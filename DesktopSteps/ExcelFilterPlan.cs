using System.Text.Json;
using System.Windows.Automation;

namespace DesktopSteps;

// Normalizes Excel checklist interactions into an explicit final filter outcome.
// Replay applies that outcome once, avoiding brittle checkbox toggle sequences whose
// starting state may differ between recording and replay.
internal static class ExcelFilterPlan
{
    private sealed record FilterSelection(bool All, HashSet<string> Exceptions);

    internal static PlanStep Confirm(PlanStep step, string mode, string[] values)
    {
        // Excluding no values is the same observable outcome as selecting every
        // value. Store that canonical outcome so replay never depends on an
        // empty exception list.
        if (mode == "exclude" && values.Length == 0) mode = "all";
        if (mode is not ("only" or "exclude" or "all") ||
            mode != "all" && values.Length == 0 ||
            values.Any(value => string.IsNullOrWhiteSpace(value) || value == "(Select All)"))
            throw new InvalidDataException("A filter requires an explicit mode and named values, or an explicit clear-filter choice.");
        var names = mode == "all" ? [] : values.Distinct(StringComparer.Ordinal).ToArray();
        return step with { Action = "filter-values", Value = JsonSerializer.Serialize(names),
            ExpectedState = "filter:" + mode,
            Explanation = $"User confirmed filter '{step.Target?.ParentName}': {mode} {string.Join(", ", names)}." };
    }

    private static PlanStep? FromExplicitIntent(PlanStep opener)
    {
        if (opener.Intent is not { } intent) return null;
        var match = System.Text.RegularExpressions.Regex.Match(intent,
            "^User confirmed filter outcome: (only|exclude) (\\[.*\\])\\.$");
        return match.Success
            ? Confirm(opener, match.Groups[1].Value,
                JsonSerializer.Deserialize<string[]>(match.Groups[2].Value)
                ?? throw new InvalidDataException("Confirmed filter intent has no values."))
            : intent == "User confirmed filter outcome: all."
                ? Confirm(opener, "all", []) : null;
    }

    internal static bool IsItem(ControlRef? target) =>
        target is { Process: "EXCEL", ControlType: "ControlType.TreeItem", ParentName: "Manual Filter" };

    internal static string? CheckState(AutomationElement element)
    {
        if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var toggle))
            return ((TogglePattern)toggle).Current.ToggleState switch
            {
                ToggleState.On => "filter:checked",
                ToggleState.Off => "filter:unchecked",
                _ => "filter:mixed"
            };
        var bounds = element.Current.BoundingRectangle;
        if (!element.Current.IsOffscreen && !bounds.IsEmpty)
        {
            // Tree-item bounds include blank space beyond a short label. MSAA
            // can hit the tree background there; probe the checkbox and label.
            foreach (var offset in new[] { 8.0, 16.0, 32.0, bounds.Width / 2 })
            {
                if (offset >= bounds.Width) continue;
                if (MsaaActions.TreeItemStateAtPoint((int)(bounds.Left + offset),
                    (int)(bounds.Top + bounds.Height / 2), element.Current.Name) is not { } state) continue;
                return (state & 0x20) != 0 ? "filter:mixed" :
                    (state & 0x10) != 0 ? "filter:checked" : "filter:unchecked";
            }
        }
        return null;
    }

    internal static string? CaptureState(ControlRef target) =>
        Automation.Resolve(target) is { } item ? CheckState(item) : null;

    internal static List<PlanStep> Compact(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>();
        var priorSelections = new Dictionary<string, FilterSelection>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < steps.Count; index++)
        {
            var opener = steps[index];
            if (opener.Action == "filter-values" && FromExplicitIntent(opener) is { } confirmedExisting)
            {
                result.Add(confirmedExisting with { OriginIntent = opener.Intent });
                continue;
            }
            if (opener is not { Action: "click", Target:
                { Process: "EXCEL", AutomationId: "Dropdown", ParentName: { Length: > 0 } } })
            { result.Add(opener); continue; }
            var end = index + 1;
            if (FromExplicitIntent(opener) is { } confirmed)
            {
                while (end < steps.Count && end - index <= 12 &&
                    steps[end].Action is "click" or "ensure-state" &&
                    steps[end].Target?.Process == opener.Target.Process &&
                    steps[end].Target?.Window == opener.Target.Window &&
                    steps[end].WhenUser == opener.WhenUser)
                {
                    if (steps[end].Target is { Name: "OK", ControlType: "ControlType.Button" }) break;
                    end++;
                }
                if (end < steps.Count && end > index + 1 &&
                    end - index <= 12 && steps[end].Action == "click" &&
                    steps[end].Target?.Process == opener.Target.Process &&
                    steps[end].Target?.Window == opener.Target.Window &&
                    steps[end].WhenUser == opener.WhenUser &&
                    steps[end].Target is { Name: "OK", ControlType: "ControlType.Button" } &&
                    steps.Skip(index + 1).Take(end - index - 1).Any(item => IsItem(item.Target)) &&
                    steps.Skip(index + 1).Take(end - index - 1).All(item =>
                        IsItem(item.Target) || item.Target?.ControlType is "ControlType.DataItem" or "ControlType.Image"))
                {
                    result.Add(confirmed with { Screenshot = steps[end].Screenshot ?? opener.Screenshot,
                        OriginIntent = opener.Intent });
                    index = end;
                    continue;
                }
            }
            end = index + 1;
            while (end < steps.Count && steps[end].Action is "click" or "ensure-state" &&
                IsItem(steps[end].Target) && steps[end].Target?.Window == opener.Target.Window) end++;
            if (end == index + 1)
            { result.Add(opener); continue; }
            var explicitApply = end < steps.Count && steps[end] is
                { Action: "click", Target: { Process: "EXCEL", Name: "OK", ControlType: "ControlType.Button" } } &&
                steps[end].Target?.Window == opener.Target.Window;
            // Excel removes the filter flyout before the mouse-up is resolved. In that
            // case UIA can report the sheet tab exposed below the former OK button.
            // Accept that lost apply click only when the next captured control proves
            // the flyout closed and the checklist contains a Select All transition.
            var implicitApply = !explicitApply && end < steps.Count &&
                !IsItem(steps[end].Target) &&
                steps.Skip(index + 1).Take(end - index - 1)
                    .Any(item => item.Target?.Name == "(Select All)") &&
                !string.IsNullOrWhiteSpace(steps[end].Screenshot);
            if (!explicitApply && !implicitApply)
            { result.Add(opener); continue; }
            if (steps.Skip(index + 1).Take(end - index).Any(item => item.WhenUser != opener.WhenUser))
            { result.Add(opener); continue; }
            var items = steps.Skip(index + 1).Take(end - index - 1).ToArray();
            // Only a definitive checked/unchecked Select All state establishes a reset.
            // Excel can expose a delayed mixed parent state as another Select All capture after
            // a named value changes; treating that mixed observation as a click reverses only/exclude.
            var reset = Array.FindLastIndex(items, item => item.Target?.Name == "(Select All)" &&
                item.ExpectedState is "filter:checked" or "filter:unchecked");
            var finalItems = (reset >= 0 ? items.Skip(reset + 1) : items)
                .Where(item => item.Target?.Name != "(Select All)")
                .GroupBy(item => item.Target!.Name).Select(group => group.Last()).ToArray();
            var mode = "unspecified";
            var names = finalItems.Select(item => item.Target!.Name!).Distinct().ToArray();
            if (reset >= 0 && items[reset].ExpectedState is "filter:checked" or "filter:unchecked" &&
                finalItems.All(item => item.ExpectedState is "filter:checked" or "filter:unchecked"))
            {
                mode = items[reset].ExpectedState == "filter:checked" ? "exclude" : "only";
                names = finalItems.Where(item => item.ExpectedState != items[reset].ExpectedState)
                    .Select(item => item.Target!.Name!).ToArray();
                if (mode == "exclude" && names.Length == 0) mode = "all";
            }
            var selectionKey = string.Join("\u001f", opener.Target.Process, opener.Target.Window,
                opener.Target.ParentName, opener.WhenUser ?? "");
            var reconstructionItems = reset >= 0
                ? items.Where((item, itemIndex) => itemIndex <= reset ||
                    item.Target?.Name != "(Select All)" || item.ExpectedState != "filter:mixed").ToArray()
                : items;
            if (mode == "unspecified" && TryReconstructSelection(opener, reconstructionItems,
                    priorSelections.GetValueOrDefault(selectionKey), out var reconstructed))
            {
                mode = reconstructed.All ? reconstructed.Exceptions.Count == 0 ? "all" : "exclude" : "only";
                names = reconstructed.Exceptions.Order(StringComparer.Ordinal).ToArray();
            }
            if (mode is "all" or "only" or "exclude")
                priorSelections[selectionKey] = mode switch
                {
                    "all" => new(true, []),
                    "exclude" => new(true, names.ToHashSet(StringComparer.Ordinal)),
                    _ => new(false, names.ToHashSet(StringComparer.Ordinal))
                };
            result.Add(opener with { Action = "filter-values", Value = JsonSerializer.Serialize(names),
                ExpectedState = "filter:" + mode, Screenshot = explicitApply
                    ? steps[end].Screenshot ?? opener.Screenshot : steps[end].Screenshot,
                Intent = string.Join("\n", steps.Skip(index).Take(end - index + (explicitApply ? 1 : 0))
                    .Select(item => item.Intent).Where(note => !string.IsNullOrWhiteSpace(note)).Distinct()) is { Length: > 0 } intent ? intent : null,
                OriginIntent = steps.Skip(index).Take(end - index + (explicitApply ? 1 : 0)).Select(item => item.OriginIntent)
                    .FirstOrDefault(note => note is not null),
                Explanation = mode == "unspecified"
                    ? "Choose the final filter selection before replay; the recording did not capture checkbox states."
                    : $"Apply verified filter outcome: {mode} {string.Join(", ", names)}." });
            // The first control exposed after an implicitly applied checklist is
            // the recorder's mistaken identity for the missing apply click. It is
            // closure evidence, not a second user action, so consume it as part of
            // the filter sequence.
            index = end;
        }
        return result;
    }

    private static bool TryReconstructSelection(PlanStep opener, IReadOnlyList<PlanStep> items,
        FilterSelection? prior, out FilterSelection selection)
    {
        selection = opener.Target?.Name == "No filter applied"
            ? new(true, [])
            : prior is not null
                ? new(prior.All, new(prior.Exceptions, StringComparer.Ordinal))
                : new(false, []);
        var sawReset = false;
        foreach (var item in items)
        {
            var name = item.Target?.Name;
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (name == "(Select All)")
            {
                sawReset = true;
                selection = item.ExpectedState switch
                {
                    "filter:checked" => new(true, []),
                    "filter:unchecked" => new(false, []),
                    _ => selection.All && selection.Exceptions.Count == 0
                        ? new(false, []) : new(true, [])
                };
                continue;
            }
            if (!sawReset) return false;
            var exceptions = new HashSet<string>(selection.Exceptions, StringComparer.Ordinal);
            if (!exceptions.Add(name)) exceptions.Remove(name);
            selection = selection with { Exceptions = exceptions };
        }
        return sawReset && (selection.All || selection.Exceptions.Count > 0);
    }

    internal static string[] Values(PlanStep step) =>
        JsonSerializer.Deserialize<string[]>(step.Value ?? throw new InvalidDataException("Missing filter values."))
        ?? throw new InvalidDataException("Empty filter values.");

    internal static ExecutionPlan RecoverRecordedOutcomes(ExecutionPlan plan,
        IReadOnlyList<RecordedEvent> recordedEvents)
    {
        var evidence = recordedEvents.Select((item, index) => new PlanStep(index + 1,
            item.Kind, item.Target, item.Value, item.Key, item.AfterState, item.Screenshot, null,
            item.Intent)).ToList();
        var recovered = Compact(evidence).Where(step => step.Action == "filter-values" &&
            step.ExpectedState is "filter:only" or "filter:exclude" or "filter:all").ToList();
        if (recovered.Count == 0) return plan;

        var cursor = 0;
        var repaired = plan.Steps.ToList();
        for (var index = 0; index < repaired.Count && cursor < recovered.Count; index++)
        {
            var saved = repaired[index];
            if (saved.Action != "filter-values") continue;
            var match = recovered.Skip(cursor).FirstOrDefault(candidate =>
                candidate.Target?.Process == saved.Target?.Process &&
                candidate.Target?.Window == saved.Target?.Window &&
                candidate.Target?.ParentName == saved.Target?.ParentName);
            if (match is null) continue;
            cursor = recovered.IndexOf(match) + 1;
            // Preserve the generated target and intent attribution. Only the recorded checklist's
            // verified final mode and values replace the earlier inferred filter outcome.
            repaired[index] = saved with { Value = match.Value, ExpectedState = match.ExpectedState,
                Screenshot = match.Screenshot ?? saved.Screenshot,
                Explanation = match.Explanation };
        }
        return plan with { Steps = repaired };
    }
}
