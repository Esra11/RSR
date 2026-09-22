using System.Text.RegularExpressions;

namespace DesktopSteps;

internal static class PlanCompactor
{
    internal static PlanStep? FollowingEditValue(IReadOnlyList<PlanStep> steps, int index)
    {
        var click = steps[index];
        if (click.Action != "click" || click.Target is not
            { ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } } edit ||
            !string.IsNullOrWhiteSpace(edit.Name)) return null;
        for (var nextIndex = index + 1; nextIndex < Math.Min(steps.Count, index + 5); nextIndex++)
        {
            var next = steps[nextIndex];
            if (next.Target is not { ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } } target ||
                !string.Equals(target.Process, edit.Process, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(target.Window, edit.Window, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(target.AutomationId, edit.AutomationId, StringComparison.OrdinalIgnoreCase))
                return null;
            if (next.Action == "type" && next.Value is not null) return next;
            if (next.Action != "click") return null;
        }
        return null;
    }

    public static ExecutionPlan Compact(ExecutionPlan plan)
    {
        var source = NormalizeSearchEditing(plan.Steps);
        var compact = new List<PlanStep>();
        for (var i = 0; i < source.Count; i++)
        {
            var step = source[i];
            if (step.Action == "click" && step.Target is
                { Process: "explorer", Window: "Program Manager", Name: "Desktop",
                  ControlType: "ControlType.List", ClassName: "SysListView32" })
                continue;
            if (step.Action == "click" && step.ExpectedState == "unverified-click" &&
                step.Target?.ControlType == "ControlType.DataItem" &&
                string.IsNullOrWhiteSpace(step.Target.AutomationId))
            {
                var repeated = i + 1 < source.Count && source[i + 1].Action == "click" &&
                    source[i + 1].ExpectedState == "unverified-click" && source[i + 1].Target == step.Target;
                compact.Add(step with { Action = "manual-click", Explanation = repeated
                    ? "Repeat the recorded double-click in the target app; the recorder could not identify a safe control to resize."
                    : "Perform this click in the target app; the recorder could not identify a safe control." });
                if (repeated) i++;
                continue;
            }
            if (step.Action == "click" && step.ExpectedState == "window-closed" &&
                step.Target?.ControlType is ("ControlType.Window" or "ControlType.TitleBar"))
            {
                compact.Add(step with { Action = "close-window", Explanation =
                    $"Close {step.Target.Window ?? "the current window"} before continuing." });
                continue;
            }
            if (step.Intent is not null && step.Key is not null &&
                step.Key.Contains("Control", StringComparison.OrdinalIgnoreCase) &&
                step.Key.Contains("Shift", StringComparison.OrdinalIgnoreCase) &&
                step.Key.EndsWith("+I", StringComparison.OrdinalIgnoreCase) &&
                step.Intent.Contains("sort", StringComparison.OrdinalIgnoreCase))
            {
                var headerIndex = compact.FindLastIndex(x => x.Target?.ControlType == "ControlType.HeaderItem" &&
                    x.Target.Process == step.Target?.Process);
                if (headerIndex >= 0)
                {
                    var direction = step.Intent.Contains("descending", StringComparison.OrdinalIgnoreCase)
                        ? "descending" : step.Intent.Contains("ascending", StringComparison.OrdinalIgnoreCase)
                        ? "ascending" : null;
                    if (direction is not null)
                    {
                        var header = compact[headerIndex];
                        compact[headerIndex] = header with { Action = "ensure-state",
                            ExpectedState = "sort:" + direction, Intent = step.Intent,
                            Explanation = $"Ensure {header.Target?.Name} is sorted {direction}." };
                        continue; // The intent shortcut is metadata, never a replay action.
                    }
                }
            }
            if (step.Action == "click" && step.Target?.ControlType is "ControlType.Tab" or "ControlType.Window")
            {
                compact.Add(step with { Action = "manual-click", Explanation =
                    "The recorder saved this click but could not identify the control. Review it in the recording before replay." });
                continue;
            }
            if (step.Action == "click" && step.Target?.ControlType == "ControlType.HeaderItem" &&
                i + 1 < source.Count && source[i + 1].Action == "ensure-state" &&
                source[i + 1].Target == step.Target) continue;

            // Excel's inline tab editor exposes its keystrokes as a generic pane. A preceding
            // sheet tab identifies the actual object being renamed; collapse corrections into
            // the final name while preserving the selected sheet step before it.
            if (step.Action == "click" && step.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ControlType == "ControlType.TabItem" && i + 1 < source.Count &&
                (source[i + 1].Action is "type" or "key" || IsInlineSheetEditClick(source[i + 1])))
            {
                var end = i + 1;
                while (end < source.Count && (source[end].Action is "type" or "key" || IsInlineSheetEditClick(source[end])) &&
                    source[end].Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true)
                {
                    if (source[end].Key is "Enter" or "Return") break;
                    end++;
                }
                var observedName = end < source.Count &&
                    source[end].ExpectedState?.StartsWith("sheet-name:", StringComparison.Ordinal) == true
                    ? source[end].ExpectedState!["sheet-name:".Length..] : null;
                var reconstructed = TryFinalText(source.Skip(i + 1).Take(Math.Max(0, end - i - 1)).ToList(), out var typedName);
                if (end < source.Count && source[end].Key is "Enter" or "Return" &&
                    (!string.IsNullOrWhiteSpace(observedName) || reconstructed))
                {
                    var finalName = !string.IsNullOrWhiteSpace(observedName) ? observedName! : typedName;
                    // A user can add an intent immediately after committing the rename.
                    // That note may land on a second Enter from the inline sheet editor.
                    var noteIndex = end;
                    for (var next = end + 1; next < Math.Min(source.Count, end + 5); next++)
                    {
                        if (source[next].Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) != true ||
                            source[next].Target?.ControlType != "ControlType.Pane" || source[next].Action != "key") break;
                        if (source[next].Intent is not null) { noteIndex = next; break; }
                    }
                    var intent = source[noteIndex].Intent ?? source.Skip(i + 1).Take(end - i)
                        .LastOrDefault(x => x.Intent is not null)?.Intent;
                    var weekday = source[noteIndex].RelativeWeekday ?? source.Skip(i + 1).Take(end - i)
                        .LastOrDefault(x => x.RelativeWeekday is not null)?.RelativeWeekday;
                    if (weekday is not null && intent is not null)
                    {
                        var form = Regex.Match(intent, @"\b(?:form|format)\s+(?<prefix>[^()]+)\(", RegexOptions.IgnoreCase);
                        var parenthesis = finalName.IndexOf('(');
                        if (form.Success && parenthesis > 0 && form.Groups["prefix"].Value.Trim() is { Length: > 2 and < 60 } prefix)
                            finalName = prefix + " " + finalName[parenthesis..];
                    }
                    var value = DynamicText.ToTemplate(finalName, weekday);
                    compact.Add(new PlanStep(0, "rename-sheet", step.Target, value, null,
                        value, source[end].Screenshot ?? source[end - 1].Screenshot,
                        weekday is null ? $"Rename the selected sheet to {finalName}."
                            : $"Rename the selected sheet using the date of next {weekday}.",
                        intent, weekday));
                    i = noteIndex;
                    continue;
                }
            }
            compact.Add(step);
        }
        // A run of tab selections with no work on those tabs is navigation. Keep only
        // the final tab when the next action needs its context; a context action on
        // that tab already selects it, so even the final selection is unnecessary.
        var purposeful = new List<PlanStep>();
        for (var i = 0; i < compact.Count;)
        {
            if (!IsTabNavigation(compact[i])) { purposeful.Add(compact[i++]); continue; }
            var start = i;
            while (i < compact.Count && IsTabNavigation(compact[i])) i++;
            var lastTab = compact.Skip(start).Take(i - start).LastOrDefault(IsTabClick);
            if (lastTab is null) continue;
            var next = i < compact.Count ? compact[i] : null;
            if (next is not null && next.Action is "context-click" or "rename-sheet" &&
                next.Target?.Process == lastTab.Target?.Process &&
                next.Target?.ControlType == "ControlType.TabItem") continue;
            purposeful.Add(lastTab);
        }
        var final = new List<PlanStep>();
        for (var i = 0; i < purposeful.Count; i++)
        {
            var current = purposeful[i];
            var next = i + 1 < purposeful.Count ? purposeful[i + 1] : null;
            if (current.Action == "click" && current.Target is
                { ControlType: "ControlType.MenuItem", AutomationId: "Dropdown" } &&
                next?.Action == "click" && next.Target == current.Target)
                continue; // Repeated capture of the same dropdown is one menu-opening action.
            if (current.Action == "scroll" && current.Target?.ControlType == "ControlType.ToolTip" &&
                next?.Action == "scroll" && next.Target?.ControlType == "ControlType.ScrollBar" &&
                current.Target.Process == next.Target.Process && current.Key == next.Key)
                continue; // A transient tooltip intercepted the wheel; the next real scrollbar records the outcome.
            var prior = final.LastOrDefault();
            var lastHeader = final.LastOrDefault(x => x.Target?.ControlType == "ControlType.HeaderItem" &&
                x.Target.Process == current.Target?.Process && x.Target.Window == current.Target?.Window);
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.Text" &&
                lastHeader is not null && next?.Key == "Shift+Home" &&
                next.Target?.ControlType == "ControlType.ListItem" && i + 2 < purposeful.Count &&
                purposeful[i + 2].Key == "Shift+End" &&
                purposeful[i + 2].Target?.ControlType == "ControlType.ListItem" &&
                current.Target.Process == lastHeader.Target?.Process &&
                current.Target.Window == lastHeader.Target?.Window)
            {
                final.Add(current with { Action = "select-all-items", Target = lastHeader.Target,
                    Key = null, ExpectedState = null,
                    Screenshot = purposeful[i + 2].Screenshot ?? current.Screenshot,
                    Explanation = "Select every current row, regardless of row content or count." });
                i += 2;
                continue;
            }
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.Text" &&
                next?.Key == "Shift+End" && next.Target?.ControlType == "ControlType.ListItem")
            {
                final.Add(next with { Action = "select-all-items", Target = lastHeader?.Target ?? next.Target,
                    Key = null, ExpectedState = null,
                    Screenshot = next.Screenshot ?? current.Screenshot,
                    Explanation = "Select every current row, regardless of row content or count." });
                i++;
                continue;
            }
            if (current.Action == "select-all-items" && current.Target?.ControlType == "ControlType.ListItem" &&
                lastHeader is not null)
            {
                final.Add(current with { Target = lastHeader.Target, Key = null, ExpectedState = null,
                    Explanation = "Select every current row, regardless of row content or count." });
                continue;
            }
            if (current.Action == "context-selection" && prior?.Action == "select-all-items" &&
                prior.Target?.Process == current.Target?.Process && prior.Target?.Window == current.Target?.Window)
            {
                final.Add(current with { Target = prior.Target, ExpectedState = "selection-intent:all",
                    Explanation = "Open the context menu for the current selection, regardless of row text." });
                continue;
            }
            if (current.Action == "context-click" && current.Target is { } currentTarget &&
                currentTarget.ControlType is "ControlType.Text" or "ControlType.ListItem" &&
                ((prior?.Action == "select-all-items" && prior.Target is { } priorTarget &&
                  priorTarget.Process == currentTarget.Process && priorTarget.Window == currentTarget.Window) ||
                 current.ExpectedState == "selection-intent:all" ||
                 (current.ExpectedState?.StartsWith("selection-count:") == true &&
                  int.TryParse(current.ExpectedState[16..], out var selectedCount) && selectedCount > 1)))
            {
                final.Add(current with { Action = "context-selection",
                    Target = prior?.Action == "select-all-items" ? prior.Target : currentTarget,
                    ExpectedState = "selection-intent:all",
                    Explanation = "Open the context menu for all selected items, regardless of item text." });
                continue;
            }
            if (current.Action == "select-first-row" && next is not null &&
                next.Key == "Shift+End" && next.Target?.ControlType == "ControlType.ListItem")
            {
                final.Add(current with { Action = "select-all-items",
                    Screenshot = next.Screenshot ?? current.Screenshot,
                    Explanation = "Select every item in the list, regardless of its contents." });
                i++;
                continue;
            }
            if (current.Action == "context-click" && current.Target?.ControlType == "ControlType.TabItem" &&
                next?.Action == "context-click" && next.Target?.ControlType == "ControlType.TabItem" &&
                current.Target.Process == next.Target.Process && current.Target.Window == next.Target.Window)
                continue;
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.Button" &&
                current.Target.Name is "Page left" or "Scroll Left" &&
                next?.Action == "scroll" && next.Key == "Horizontal" &&
                current.Target.Process == next.Target?.Process)
                continue;
            var previous = final.LastOrDefault();
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.Text" &&
                previous?.Action == "ensure-state" && previous.Target?.ControlType == "ControlType.HeaderItem" &&
                next is not null && next.Key == "Shift+End" &&
                current.Target.Process == previous.Target.Process && current.Target.Window == previous.Target.Window)
            {
                final.Add(current with { Action = "select-all-items", Target = previous.Target,
                    Screenshot = next.Screenshot ?? current.Screenshot,
                    Explanation = "Select every item in the sorted list, regardless of its contents." });
                i++; // The following Shift+End expresses this one selection intent.
                continue;
            }
            if (current.Action == "ensure-state" && next?.Action == "ensure-state" &&
                current.Target?.ControlType == "ControlType.HeaderItem" &&
                current.Target == next.Target &&
                (current.ExpectedState?.Contains("ascending", StringComparison.OrdinalIgnoreCase) == true ||
                 current.ExpectedState?.Contains("descending", StringComparison.OrdinalIgnoreCase) == true) &&
                (next.ExpectedState?.Contains("ascending", StringComparison.OrdinalIgnoreCase) == true ||
                 next.ExpectedState?.Contains("descending", StringComparison.OrdinalIgnoreCase) == true))
                continue;
            if (current.Action == "ensure-state" && current.Target?.ControlType == "ControlType.HeaderItem" &&
                prior?.Action == "scroll" && final.Count >= 2 &&
                final[^2].Action == "ensure-state" && final[^2].Target == current.Target &&
                final[^2].ExpectedState == current.ExpectedState)
                continue;
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.ListItem" &&
                next?.Action == "click" && next.Target?.ControlType == "ControlType.ListItem" &&
                current.Target.Window == next.Target.Window && current.Target.ParentName == next.Target.ParentName)
                continue;
            final.Add(current);
        }
        for (var index = 0; index < final.Count; index++)
        {
            var button = final[index];
            if (button.Action is not ("click" or "replace-all") || button.Target?.Window != "Find and Replace" ||
                button.Target.Name != "Replace All") continue;
            string? FieldValue(string id) => final.Take(index).LastOrDefault(step => step.Target is
                { Window: "Find and Replace", AutomationId: var fieldId } && fieldId == id &&
                step.Action == "type") is { } field
                ? field.ExpectedState?.StartsWith("field-value:", StringComparison.Ordinal) == true
                    ? field.ExpectedState["field-value:".Length..] : null
                : null;
            var find = FieldValue("18");
            var replacement = FieldValue("21");
            if (!string.IsNullOrWhiteSpace(find) && replacement is not null)
                final[index] = button with { Action = "replace-all", Value = find,
                    ExpectedState = replacement,
                    Explanation = $"Click the recorded Replace All button to replace {find} with {replacement}." };
        }
        final = CollapseSupersededDialogEdits(final);
        final = ApplyFindReplaceIntent(final, source);
        for (var index = 0; index < final.Count; index++)
            if (final[index].Action is "select-all-items" or "ensure-state" && final[index].Target is
                { Process: "EXCEL", ClassName: "XLSelectAllHeader", Name: "Select All" })
                final[index] = final[index] with { Action = "click", ExpectedState = null,
                    Explanation = "Select the entire worksheet with Excel's Select All header." };
        var lastReplacement = final.FindLastIndex(step => step.Action == "replace-all");
        if (lastReplacement >= 0 && lastReplacement + 1 < final.Count &&
            final[lastReplacement + 1].Target?.Window != "Find and Replace" &&
            !final.Skip(lastReplacement + 1).Any(step => step.Action == "close-window" &&
                step.Target?.Window == "Find and Replace" ||
                step.Action == "click" && step.Target?.Window == "Find and Replace" &&
                step.Target.Name == "Close"))
        {
            final.Insert(lastReplacement + 1, new PlanStep(0, "close-window",
                new ControlRef("EXCEL", "Find and Replace", null, "Find and Replace",
                    "ControlType.Window", null, null), null, null, "window-closed", null,
                "Close Find and Replace before working in the worksheet."));
        }
        for (var index = final.Count - 1; index >= 0; index--)
            if (FollowingEditValue(final, index) is not null)
                final.RemoveAt(index);
        for (var index = final.Count - 1; index >= 0; index--)
        {
            var click = final[index];
            if (click.Action != "manual-click" || click.Target is not
                { ControlType: "ControlType.Window", Name: { Length: > 0 } windowName } windowClick)
                continue;
            var following = final.Skip(index + 1).Take(3).ToList();
            if (following.Count == 0 || !following.Any(next => next.Action == "type" &&
                    next.Target is { ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } } edit &&
                    edit.Process?.Equals(windowClick.Process, StringComparison.OrdinalIgnoreCase) == true &&
                    edit.Window?.Equals(windowName, StringComparison.OrdinalIgnoreCase) == true) ||
                following.TakeWhile(next => next.Action != "type").Any(next =>
                    next.Action != "manual-click" || next.Target?.ControlType != "ControlType.Window" ||
                    next.Target.Name != windowName || next.Target.Process != windowClick.Process))
                continue;
            final.RemoveAt(index);
        }
        return plan with { Steps = final.Select((step, index) => step with { Number = index + 1 }).ToList() };
    }

    private static List<PlanStep> ApplyFindReplaceIntent(List<PlanStep> steps, List<PlanStep> source)
    {
        // Attach conditions only to a matching, recorded Replace All action.
        // Every recorded press stays in order, including presses the note omits.
        var note = source.Select(step => step.Intent).LastOrDefault(intent =>
            intent?.Contains("replace", StringComparison.OrdinalIgnoreCase) == true);
        if (note is null) return steps;
        var conditional = Regex.Matches(note,
            @"\bif\s+(?<user>[\w\\]+)[^"".]*?\breplace\s+""(?<find>[^""]+)""\s+with\s+""(?<replacement>[^""]*)""",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        var byUser = conditional.Cast<Match>().Where(match =>
            !new[] { "the", "teh", "app" }.Contains(match.Groups["user"].Value, StringComparer.OrdinalIgnoreCase))
            .Concat(Regex.Matches(note,
                @"\bif\s+(?:the|teh)\s+app\s+is\s+running\s+by\s+(?<user>[\w\\]+)[^"".]*?\breplace\s+""(?<find>[^""]+)""\s+with\s+""(?<replacement>[^""]*)""",
                RegexOptions.IgnoreCase | RegexOptions.Singleline).Cast<Match>())
            .OrderBy(match => match.Index).ToList();
        foreach (Match match in byUser)
        {
            var user = match.Groups["user"].Value;
            var find = match.Groups["find"].Value;
            var replacement = match.Groups["replacement"].Value;
            if (find.Length == 0) continue;
            for (var index = 0; index < steps.Count; index++)
                if (steps[index].Action == "replace-all" && steps[index].Value == find &&
                    steps[index].ExpectedState == replacement && steps[index].WhenUser is null)
                    steps[index] = steps[index] with { WhenUser = user, Intent = note,
                        Explanation = $"For {user}, perform the recorded replacement of {find} with {replacement}." };
        }
        return steps;
    }

    private static List<PlanStep> CollapseSupersededDialogEdits(List<PlanStep> steps)
    {
        var keep = Enumerable.Repeat(true, steps.Count).ToArray();
        for (var start = 0; start < steps.Count;)
        {
            var end = start;
            while (end < steps.Count && steps[end].Action != "replace-all") end++;
            foreach (var fieldId in new[] { "18", "21" })
            {
                var lastValue = -1;
                for (var index = start; index < end; index++)
                    if (steps[index].Action == "type" && steps[index].Target is
                        { Window: "Find and Replace", AutomationId: var id } && id == fieldId)
                        lastValue = index;
                if (lastValue < 0) continue;
                for (var index = start; index < lastValue; index++)
                    if (steps[index].Action is "type" or "key" && steps[index].Target is
                        { Window: "Find and Replace", AutomationId: var id } && id == fieldId)
                        keep[index] = false;
            }
            start = end + 1;
        }
        return steps.Where((_, index) => keep[index]).ToList();
    }

    private static bool IsTabClick(PlanStep step) => step.Action == "click" &&
        step.Target?.ControlType == "ControlType.TabItem";

    private static bool IsInlineSheetEditClick(PlanStep step) =>
        step.Action == "click" && step.Target is
            { Process: "EXCEL", ControlType: "ControlType.Pane" } &&
        step.Key is "Left" or "Right" or "Home" or "End" or "Back" or "Backspace" or "Delete" or "Enter" or "Return";

    private static bool IsTabNavigation(PlanStep step) => IsTabClick(step) ||
        step.Action == "click" && step.Target?.ControlType == "ControlType.Button" &&
        step.Target.AutomationId == "SheetTab" && step.Target.Name?.StartsWith("Scroll ", StringComparison.Ordinal) == true;

    private static bool TryFinalText(IReadOnlyList<PlanStep> edits, out string final)
    {
        var value = "";
        var cursor = 0;
        var selectionStart = -1;
        foreach (var step in edits)
        {
            if (step.Action == "type" && step.Value is not null)
            {
                if (selectionStart >= 0)
                {
                    value = value.Remove(selectionStart, cursor - selectionStart);
                    cursor = selectionStart;
                    selectionStart = -1;
                }
                value = value.Insert(cursor, step.Value);
                cursor += step.Value.Length;
            }
            else if (step.Key?.Contains("Control", StringComparison.OrdinalIgnoreCase) == true &&
                     step.Key.Contains("Shift", StringComparison.OrdinalIgnoreCase) &&
                     step.Key.EndsWith("Left", StringComparison.OrdinalIgnoreCase))
            {
                var start = cursor;
                while (start > 0 && char.IsWhiteSpace(value[start - 1])) start--;
                while (start > 0 && !char.IsWhiteSpace(value[start - 1])) start--;
                selectionStart = start;
            }
            else if (step.Key == "Left") { cursor = Math.Max(0, cursor - 1); selectionStart = -1; }
            else if (step.Key is "Back" or "Backspace")
            {
                if (cursor > 0) { value = value.Remove(cursor - 1, 1); cursor--; }
                selectionStart = -1;
            }
            else { final = ""; return false; }
        }
        final = value;
        return final.Length > 0;
    }

    private static List<PlanStep> NormalizeSearchEditing(IReadOnlyList<PlanStep> steps)
    {
        var normalized = new List<PlanStep>(steps.Count);
        for (var index = 0; index < steps.Count;)
        {
            if (steps[index].Target is not { Process: "SearchHost", AutomationId: "SearchTextBox" } ||
                steps[index].Action is not ("type" or "key") ||
                steps[index].Key is "Enter" or "Return")
            {
                normalized.Add(steps[index++]);
                continue;
            }
            var end = index;
            while (end < steps.Count && steps[end].Target is
                { Process: "SearchHost", AutomationId: "SearchTextBox" } &&
                steps[end].Action is "type" or "key" &&
                steps[end].Key is not ("Enter" or "Return")) end++;
            if (end >= steps.Count || steps[end] is not { Action: "key", Key: "Enter" or "Return",
                Target: { Process: "SearchHost", AutomationId: "SearchTextBox" } } enter ||
                !TryFinalText(steps.Skip(index).Take(end - index).ToArray(), out var value))
            {
                normalized.Add(steps[index++]);
                continue;
            }
            var first = steps[index];
            normalized.Add(first with { Action = "type", Target = first.Target! with { Name = value },
                Value = value, Key = null, ExpectedState = "field-value:" + value,
                Screenshot = steps.Skip(index).Take(end - index)
                    .Select(item => item.Screenshot).LastOrDefault(item => !string.IsNullOrWhiteSpace(item)),
                Explanation = "Enter the final recorded Windows Search query." });
            normalized.Add(enter);
            index = end + 1;
        }
        return normalized;
    }
}
