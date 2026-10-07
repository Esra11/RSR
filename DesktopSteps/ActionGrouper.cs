namespace DesktopSteps;

internal static class ActionGrouper
{
    internal static List<RecordedEvent> RecoverMisclassifiedLegendDrags(IReadOnlyList<RecordedEvent> source) =>
        source.Select((item, index) =>
        {
            if (item is not { Kind: "unresolved-input",
                Diagnostic: "No supported semantic resize outcome was captured for this drag.",
                Target: { Process: "EXCEL" } target })
                return item;
            if (target.ClassName == "NetUIRepeatButton" &&
                target.ControlType == "ControlType.Button")
                return item with { Kind = "click",
                    Diagnostic = "Recovered recorded repeat-button click misclassified by legend-drag detection." };
            var preceding = index > 0 ? source[index - 1] : null;
            var following = index + 1 < source.Count ? source[index + 1] : null;
            if (target.ControlType == "ControlType.Image" && !string.IsNullOrWhiteSpace(target.Name) &&
                (ConfirmsSameChart(preceding, item, target) || ConfirmsSameChart(following, item, target)))
                return item with { Kind = "click",
                    Diagnostic = "Recovered chart selection misclassified as a drag; an adjacent input independently identified the same chart." };
            if (target.ControlType != "ControlType.DataItem") return item;
            var next = following;
            var verifiedCopy = next is { Kind: "key", Key: "Control+C" or "Ctrl+C" } &&
                next.Target?.Process == target.Process && next.Target.Window == target.Window &&
                next.AfterState?.StartsWith("excel-selection-range:", StringComparison.Ordinal) == true;
            return item with { Kind = verifiedCopy ? "click" : "unresolved-click",
                Diagnostic = verifiedCopy
                    ? "Recovered selection gesture misclassified by legend-drag detection; following Ctrl+C has a verified range."
                    : "Worksheet drag target was captured, but its final selection or fill outcome was not saved. Review this gesture manually; it was not a legend resize." };
        }).Where(item => !IsConfirmedTabReveal(item, source)).ToList();

    private static bool ConfirmsSameChart(RecordedEvent? candidate, RecordedEvent item, ControlRef target) =>
        candidate is { Kind: "click" or "context-click", Target: { } confirmed } &&
        Math.Abs((candidate.At - item.At).TotalSeconds) < 3 &&
        confirmed.Process == target.Process && confirmed.ProcessId == target.ProcessId &&
        confirmed.Window == target.Window && confirmed.Name == target.Name &&
        confirmed.ControlType == target.ControlType;

    private static bool IsConfirmedTabReveal(RecordedEvent item, IReadOnlyList<RecordedEvent> source) =>
        item is { Kind: "unresolved-input", Target: null, Key: null, Value: null,
            Diagnostic: "The mouse input identity could not be frozen before delayed processing. Re-record this section; no target was guessed.",
            Intent: "User clarification: this unresolved input only revealed worksheet tabs; navigate directly to the subsequently recorded named worksheet instead." } &&
        source.Any(next => next.At > item.At && next.At - item.At < TimeSpan.FromSeconds(5) &&
            next is { Kind: "click", Target: { Process: "EXCEL", ControlType: "ControlType.TabItem", Name: { Length: > 0 } } });

    public static List<RecordedEvent> Group(IReadOnlyList<RecordedEvent> source)
    {
        source = ConsolidateChartSourceOutcomes(source);
        source = RejectIncompleteChartSourceEdits(source);
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

        IReadOnlyList<RecordedEvent> RejectIncompleteChartSourceEdits(IReadOnlyList<RecordedEvent> source)
        {
            return source.Select(item => item.Kind == "context-click" &&
                item.Target is { Process: "EXCEL", ControlType: "ControlType.Image", ParentName: "Chart Area" } chart &&
                string.IsNullOrWhiteSpace(chart.Name)
                    ? item with { Kind = "unresolved-input",
                        Diagnostic = "The chart context-click has no frozen chart identity. No later chart identity was substituted." }
                    : item.Target is { Process: "EXCEL", Window: "Select Data Source" } &&
                item.Kind is "click" or "key" or "type" or "chart-source-edit-input"
                    ? item with { Kind = "unresolved-input",
                        Diagnostic = "Chart source dialog input has no verified final source outcome. Re-record the chart edit." }
                    : item).ToList();
        }

        IReadOnlyList<RecordedEvent> ConsolidateChartSourceOutcomes(IReadOnlyList<RecordedEvent> source)
        {
            var result = new List<RecordedEvent>();
            foreach (var item in source)
            {
                if (item.Kind == "set-chart-source-range" && item.Target is { ClassName: "ExcelChartObject" } target &&
                    item.AfterState == "chart-source-verified")
                {
                    var start = result.FindLastIndex(candidate => candidate.Kind == "context-click" &&
                        candidate.Target == target);
                    if (start >= 0 && result.Count - start <= 40)
                    {
                        var end = result.FindLastIndex(candidate =>
                            candidate.Target?.Window == "Select Data Source" &&
                            candidate.Target.ProcessId == target.ProcessId);
                        var sequence = result.Skip(start + 1).Take(end - start).ToList();
                        if (end > start && sequence.Any(candidate => candidate.Target?.Window == "Select Data Source") &&
                            sequence.All(candidate => candidate.Target is { } sequenceTarget &&
                                sequenceTarget.Process == target.Process && sequenceTarget.ProcessId == target.ProcessId &&
                                (sequenceTarget.Window == target.Window || sequenceTarget.Window == "Select Data Source") &&
                                (sequenceTarget.Window == "Select Data Source" ||
                                    candidate.Kind is "click" or "scroll" &&
                                    (sequenceTarget.Name == "Select Data..." ||
                                        sequenceTarget.ControlType == "ControlType.Image" &&
                                        sequenceTarget.Name == target.Name ||
                                        sequenceTarget.ClassName == "XLSpreadsheetCell" ||
                                        sequenceTarget.ControlType == "ControlType.ScrollBar" ||
                                        sequenceTarget.ControlType == "ControlType.Image" &&
                                        string.IsNullOrWhiteSpace(sequenceTarget.Name))) &&
                                (candidate.Kind is "click" or "scroll" or "key" or "type" ||
                                    candidate.Kind == "chart-source-edit-input" &&
                                    sequenceTarget.Window == "Select Data Source" ||
                                    candidate.Kind == "unresolved-input" &&
                                    sequenceTarget.Window == "Select Data Source" &&
                                    candidate.Diagnostic == "The keyboard input identity could not be frozen before delayed processing. Re-record this section; no target was guessed.")))
                        {
                            while (start > 0 && result[start - 1] is { Kind: "click" or "context-click",
                                Intent: null, Target: { } selection } &&
                                selection.Process == target.Process && selection.ProcessId == target.ProcessId &&
                                selection.Window == target.Window && selection.Name == target.Name &&
                                selection.ControlType == "ControlType.Image" &&
                                (result[start - 1].Kind == "click" || selection == target))
                                start--;
                            result.RemoveRange(start, end - start + 1);
                            result.Insert(start, item);
                            continue;
                        }
                    }
                }
                result.Add(item);
            }
            return result;
        }
        for (var index = 0; index < source.Count; index++)
        {
            var item = source[index];
            if (item.Kind == "chart-source-edit-input")
                item = item with { Kind = "unresolved-input",
                    Diagnostic = "Chart source dialog input has no verified final source outcome. Re-record the chart edit." };
            if (item.AfterState == "dialog-result-acknowledgement") continue;
            if (item.AfterState == "unverified-click" && index + 1 < source.Count &&
                source[index + 1] is { Kind: "click", Target: { } verified } &&
                item.Target is { } uncertain &&
                verified.Process == uncertain.Process && verified.Window == uncertain.Window &&
                verified.AutomationId == uncertain.AutomationId && verified.Name == uncertain.Name &&
                verified.ControlType == uncertain.ControlType && verified.ClassName == uncertain.ClassName &&
                source[index + 1].At - item.At < TimeSpan.FromSeconds(3))
                continue;
            if (item is { Kind: "click", Target: { ControlType: "ControlType.Document",
                    Process: { Length: > 0 } privateProcess,
                    Window: { Length: > 0 } privateDocument } } &&
                IsPrivateBrowserWindow(privateDocument) && result.LastOrDefault() is
                    { Kind: "click", Target: { } opener } &&
                string.Equals(opener.Process, privateProcess, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(opener.Window, privateDocument, StringComparison.OrdinalIgnoreCase) &&
                source.Skip(index + 1).FirstOrDefault(next =>
                    next.Target?.Process?.Equals(privateProcess,
                        StringComparison.OrdinalIgnoreCase) == true &&
                    string.Equals(next.Target.Window, privateDocument,
                        StringComparison.OrdinalIgnoreCase)) is { } privateAction &&
                privateAction.At - item.At < TimeSpan.FromSeconds(8))
            {
                Flush();
                // The private-window menu item can disappear before delayed UIA capture. A new
                // private document followed by an action in that same window verifies the outcome.
                result[^1] = item with { Kind = "open-private-window", Target = opener,
                    AfterState = "opened-window:" + privateDocument };
                continue;
            }
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
                var lastAt = first.At;
                while (index + 1 < input.Count && input[index + 1] is { Target: { ControlType: "ControlType.Pane" } } next &&
                       string.IsNullOrWhiteSpace(next.Target.AutomationId) &&
                       next.Target.Process == first.Target.Process && next.Target.Window == first.Target.Window &&
                       next.At >= lastAt && next.At - lastAt < TimeSpan.FromSeconds(3))
                {
                    if (!TryContinueGridEdit(first.Target, next.Target, first.Value ?? "",
                            next.Kind, next.Value, next.Key, out var value)) break;
                    first = first with { Value = value, AfterState = null,
                        Intent = next.Intent ?? first.Intent, Screenshot = next.Screenshot ?? first.Screenshot };
                    lastAt = next.At;
                    index++;
                }
            }
            output.Add(first);
        }
        return output;
    }

    internal static bool TryContinueGridEdit(ControlRef source, ControlRef continuation, string text,
        string action, string? value, string? key, out string result)
    {
        result = text;
        if (source.Process != continuation.Process || source.Window != continuation.Window ||
            !string.IsNullOrWhiteSpace(continuation.AutomationId) ||
            continuation.ControlType != "ControlType.Pane") return false;
        if (action == "type" && value is not null)
        {
            result = text + value;
            return true;
        }
        if (source is { Process: "EXCEL", ClassName: "XLSpreadsheetCell" } &&
            continuation.ClassName == "EXCEL6" && action == "key" && key is "Back" or "Backspace" &&
            text.Length > 0)
        {
            result = text[..^1];
            return true;
        }
        return false;
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

    private static bool IsPrivateBrowserWindow(string title) =>
        title.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase) ||
        title.Contains("Incognito", StringComparison.OrdinalIgnoreCase) ||
        title.Contains("Private Browsing", StringComparison.OrdinalIgnoreCase);

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
