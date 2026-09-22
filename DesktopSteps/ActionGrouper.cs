namespace DesktopSteps;

internal static class ActionGrouper
{
    public static List<RecordedEvent> Group(IReadOnlyList<RecordedEvent> source)
    {
        source = CollapseSearchEditing(source);
        var result = new List<RecordedEvent>();
        var text = new System.Text.StringBuilder();
        RecordedEvent? first = null, last = null;
        void Flush()
        {
            if (first is not null && text.Length > 0)
                result.Add(first with { Kind = "type", Value = text.ToString(), Key = null,
                    Screenshot = last?.Screenshot ?? first.Screenshot, Intent = last?.Intent ?? first.Intent,
                    AfterState = last?.AfterState ?? first.AfterState });
            text.Clear(); first = last = null;
        }
        for (var index = 0; index < source.Count; index++)
        {
            var item = source[index];
            if (item.Kind == "click" && string.IsNullOrWhiteSpace(item.Target?.Name) &&
                string.IsNullOrWhiteSpace(item.Target?.AutomationId))
            {
                Flush();
                var preceding = result.LastOrDefault();
                var following = source.Skip(index + 1).FirstOrDefault(next =>
                    next.Target?.Process?.Equals(item.Target?.Process,
                        StringComparison.OrdinalIgnoreCase) == true &&
                    !string.IsNullOrWhiteSpace(next.Target.Window));
                if (item.Target?.Process?.Equals("msedge", StringComparison.OrdinalIgnoreCase) == true &&
                    preceding is { Kind: "click", Target.Name: { } menuName } &&
                    menuName.Contains("Settings and more", StringComparison.OrdinalIgnoreCase) &&
                    following?.Target?.Window?.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase) == true &&
                    !string.Equals(item.Target.Window, following.Target.Window,
                        StringComparison.OrdinalIgnoreCase) &&
                    following.At - item.At < TimeSpan.FromSeconds(8))
                    result.Add(item with { Kind = "open-private-window", Target = preceding.Target,
                        AfterState = "opened-window:" + following.Target.Window });
                continue;
            }
            if (item.Kind == "key" && IsStandaloneModifier(item.Key)) continue;
            var character = item.Kind == "key" ? Character(item) : null;
            if (character is not null)
            {
                if (first is not null && (!SameTarget(first.Target, item.Target) ||
                    (item.At - last!.At > TimeSpan.FromSeconds(3) &&
                     first.Target?.Window != "Find and Replace"))) Flush();
                first ??= item;
                last = item;
                text.Append(character);
                continue;
            }
            Flush(); result.Add(item);
        }
        Flush();
        return MergeGridTyping(CollapseCompletedEdits(result));
    }

    private static IReadOnlyList<RecordedEvent> CollapseSearchEditing(IReadOnlyList<RecordedEvent> source)
    {
        var normalized = new List<RecordedEvent>(source.Count);
        for (var index = 0; index < source.Count;)
        {
            if (source[index] is not { Kind: "key", Target:
                { Process: "SearchHost", AutomationId: "SearchTextBox" } })
            {
                normalized.Add(source[index++]);
                continue;
            }
            var end = index;
            while (end < source.Count && source[end] is { Kind: "key", Target:
                { Process: "SearchHost", AutomationId: "SearchTextBox" } } &&
                source[end].Key is not ("Enter" or "Return")) end++;
            if (end >= source.Count || source[end].Target is not
                { Process: "SearchHost", AutomationId: "SearchTextBox" } ||
                source[end].Key is not ("Enter" or "Return") ||
                !TryFinalSearchValue(source.Skip(index).Take(end - index).ToArray(),
                    source[end], out var finalValue))
            {
                normalized.Add(source[index++]);
                continue;
            }
            var first = source[index];
            normalized.Add(first with { Kind = "type", Value = finalValue, Key = null,
                Target = first.Target! with { Name = finalValue },
                AfterState = "field-value:" + finalValue,
                Screenshot = source.Skip(index).Take(end - index).Select(item => item.Screenshot)
                    .LastOrDefault(value => !string.IsNullOrWhiteSpace(value)) });
            normalized.Add(source[end]);
            index = end + 1;
        }
        return normalized;
    }

    private static bool TryFinalSearchValue(IReadOnlyList<RecordedEvent> edits, RecordedEvent enter,
        out string value)
    {
        if (enter.BeforeState?.StartsWith("field-value:", StringComparison.Ordinal) == true &&
            enter.BeforeState.Length > "field-value:".Length)
        {
            value = enter.BeforeState["field-value:".Length..];
            return true;
        }
        var text = new System.Text.StringBuilder();
        foreach (var edit in edits)
        {
            if (IsStandaloneModifier(edit.Key)) continue;
            if (edit.Key is "Back" or "Backspace")
            {
                if (text.Length > 0) text.Length--;
                continue;
            }
            if (edit.Key is "Control+A" or "Ctrl+A") { text.Clear(); continue; }
            var character = Character(edit);
            if (character is null) { value = ""; return false; }
            text.Append(character);
        }
        value = text.ToString();
        return value.Length > 0;
    }

    private static List<RecordedEvent> MergeGridTyping(List<RecordedEvent> input)
    {
        var output = new List<RecordedEvent>();
        for (var index = 0; index < input.Count; index++)
        {
            var first = input[index];
            if (first.Kind == "type" && first.Target?.ControlType == "ControlType.DataItem")
            {
                while (index + 1 < input.Count && input[index + 1] is { Kind: "type", Target: { ControlType: "ControlType.Pane" } } next &&
                       string.IsNullOrWhiteSpace(next.Target.AutomationId) &&
                       next.Target.Process == first.Target.Process && next.Target.Window == first.Target.Window &&
                       next.At - first.At < TimeSpan.FromSeconds(2))
                {
                    first = first with { Value = (first.Value ?? "") + (next.Value ?? ""),
                        Intent = next.Intent ?? first.Intent, Screenshot = next.Screenshot ?? first.Screenshot };
                    index++;
                }
            }
            output.Add(first);
        }
        return output;
    }

    private static List<RecordedEvent> CollapseCompletedEdits(List<RecordedEvent> input)
    {
        var output = new List<RecordedEvent>();
        for (var index = 0; index < input.Count;)
        {
            var first = input[index];
            if (!IsEditable(first.Target)) { output.Add(first); index++; continue; }
            if (first.Kind == "key" && first.Key is "Enter" or "Return")
            {
                output.Add(first with { AfterState = null });
                index++;
                continue;
            }
            var end = index + 1;
            while (end < input.Count && IsEditAction(input[end]) &&
                   SameEditable(first.Target!, input[end].Target!) &&
                   input[end].At - input[end - 1].At < TimeSpan.FromSeconds(8))
            {
                if (input[end].Kind == "key" && input[end].Key is "Enter" or "Return")
                {
                    end++;
                    break;
                }
                end++;
            }
            var range = input.Skip(index).Take(end - index).ToList();
            if (range.Count > 0 && range[^1].Kind == "key" && range[^1].Key is "Enter" or "Return")
            {
                // Enter submits an editable field. Its post-key snapshot may be a
                // navigated URL, so it cannot replace the text typed before Enter.
                var beforeEnter = range.Take(range.Count - 1).ToList();
                var reconstructed = ReconstructEditableText(beforeEnter);
                var submitted = range[^1].BeforeState?.StartsWith("field-value:",
                    StringComparison.Ordinal) == true
                    ? range[^1].BeforeState!["field-value:".Length..] : null;
                var completedBeforeEnter = beforeEnter.LastOrDefault(item =>
                    item.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) == true);
                submitted ??= completedBeforeEnter?.AfterState?["field-value:".Length..];
                if (!string.IsNullOrWhiteSpace(reconstructed) &&
                    (string.IsNullOrWhiteSpace(submitted) ||
                     Uri.TryCreate(submitted, UriKind.Absolute, out var resultingAddress) &&
                     resultingAddress.Query.Length > 0 &&
                     !submitted.Equals(reconstructed, StringComparison.OrdinalIgnoreCase)))
                    submitted = reconstructed;
                if (!string.IsNullOrWhiteSpace(submitted))
                {
                    output.Add(first with { Kind = "type", Value = submitted, Key = null,
                        Target = first.Target! with { Name = submitted },
                        AfterState = "field-value:" + submitted,
                        Screenshot = beforeEnter.Select(item => item.Screenshot)
                            .LastOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? first.Screenshot });
                }
                else output.AddRange(beforeEnter);
                output.Add(range[^1] with { AfterState = null });
                index = end;
                continue;
            }
            if (first.Target is { Process: "SearchHost", AutomationId: "SearchTextBox" } &&
                input.Skip(index).Take(end - index).Any(item => item.Key is "Enter" or "Return"))
            {
                // Enter launches the selected result. The field is normally empty after
                // Search closes, so its post-key snapshot must not replace the query.
                foreach (var item in input.Skip(index).Take(end - index))
                {
                    if (item.Key is "Enter" or "Return")
                        output.Add(item with { AfterState = item.AfterState is "field-value:" ? null : item.AfterState });
                    else if (item.Kind == "type" && !string.IsNullOrWhiteSpace(item.Value) &&
                             item.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) != true)
                        output.Add(item with { AfterState = "field-value:" + item.Value });
                    else output.Add(item);
                }
                index = end;
                continue;
            }
            var completed = input.Skip(index).Take(end - index)
                .LastOrDefault(item => item.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) == true);
            if (completed is null)
            {
                output.AddRange(input.Skip(index).Take(end - index));
            }
            else
            {
                var value = completed.AfterState!["field-value:".Length..];
                output.Add(first with { Kind = "type", Value = value, Key = null,
                    Target = completed.Target, AfterState = completed.AfterState,
                    Screenshot = completed.Screenshot ?? first.Screenshot,
                    Intent = input.Skip(index).Take(end - index).LastOrDefault(x => x.Intent is not null)?.Intent });
            }
            index = end;
        }
        return output;
    }

    private static bool IsEditable(ControlRef? target) => target is not null &&
        (target.ControlType == "ControlType.Edit" || target.ClassName == "EDTBX" ||
         target.ControlType == "ControlType.Pane" && !string.IsNullOrEmpty(target.AutomationId));

    private static string? ReconstructEditableText(IReadOnlyList<RecordedEvent> edits)
    {
        var text = new System.Text.StringBuilder();
        foreach (var item in edits)
        {
            if (item.Kind == "click" || IsStandaloneModifier(item.Key)) continue;
            if (item.Key is "Back" or "Backspace")
            {
                if (text.Length > 0) text.Length--;
                continue;
            }
            if (item.Key is "Control+A" or "Ctrl+A") { text.Clear(); continue; }
            var value = item.Kind == "type" ? item.Value : Character(item);
            if (value is null) return null;
            text.Append(value);
        }
        return text.Length == 0 ? null : text.ToString();
    }
    private static bool IsEditAction(RecordedEvent item) => IsEditable(item.Target) &&
        item.Kind is "click" or "key" or "type";
    private static bool SameEditable(ControlRef a, ControlRef b) =>
        a.Process == b.Process &&
        (a.Window == b.Window || a.ProcessId is > 0 && a.ProcessId == b.ProcessId &&
            !string.IsNullOrWhiteSpace(a.AutomationId)) &&
        a.AutomationId == b.AutomationId &&
        (!string.IsNullOrEmpty(a.AutomationId) || a.Name == b.Name) &&
        a.ClassName == b.ClassName && a.ControlType == b.ControlType;

    private static bool SameTarget(ControlRef? a, ControlRef? b) => a is not null && b is not null &&
        a.Process == b.Process && a.AutomationId == b.AutomationId && a.Name == b.Name &&
        a.ControlType == b.ControlType && a.ClassName == b.ClassName;

    internal static bool IsStandaloneModifier(string? key)
    {
        var name = key?.Split('+')[^1];
        return name is "LControlKey" or "RControlKey" or "ControlKey" or
            "LShiftKey" or "RShiftKey" or "ShiftKey" or "LMenu" or "RMenu" or "Menu";
    }

    private static string? Character(RecordedEvent item)
    {
        if (item.Key?.Contains("Control", StringComparison.OrdinalIgnoreCase) == true ||
            item.Key?.Contains("Alt", StringComparison.OrdinalIgnoreCase) == true)
            return null;
        if (!string.IsNullOrEmpty(item.Value)) return item.Value;
        var key = item.Key;
        if (key is null || key.Contains('+')) return null;
        if (key == "Space") return " ";
        if (key.Length == 1 && char.IsLetter(key[0])) return key.ToLowerInvariant();
        if (key.Length == 2 && key[0] == 'D' && char.IsDigit(key[1])) return key[1].ToString();
        if (key.StartsWith("NumPad", StringComparison.Ordinal) && key.Length == 7 && char.IsDigit(key[^1])) return key[^1].ToString();
        return null;
    }
}
