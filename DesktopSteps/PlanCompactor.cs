using System.Text.RegularExpressions;

namespace DesktopSteps;

// Converts verbose captured actions into verified semantic operations. Compaction is
// evidence based: it removes redundant navigation and derives dynamic table behavior
// only when recorded state or explicit user intent proves the final outcome.
internal static class PlanCompactor
{
    internal static bool IsChartSelectionBeforeLayout(PlanStep step, PlanStep? next) =>
        step is { Action: "click", Value: null, Key: null, OriginIntent: null,
            Target: { Process: { Length: > 0 }, Window: { Length: > 0 } } target } &&
        (target is { ClassName: "ExcelChartObject", ControlType: "ControlType.Image", Name: { Length: > 0 } } ||
         target is { ControlType: "ControlType.Group", Name: "Chart Area" }) &&
        string.IsNullOrWhiteSpace(step.Intent) && string.IsNullOrWhiteSpace(step.TargetStrategy) &&
        (string.IsNullOrWhiteSpace(step.ExpectedState) ||
            step.ExpectedState == target.Name + ":selected" ||
            step.ExpectedState == target.Name + " selected") &&
        next is { Action: "set-chart-legend-layout", Value: { Length: > 0 },
            Target: { Name: { Length: > 0 } } layoutTarget,
            ExpectedState: { } state } &&
        state.StartsWith("chart-legend-sheet:", StringComparison.Ordinal) &&
        state.Length > "chart-legend-sheet:".Length &&
        layoutTarget.Process == target.Process && layoutTarget.Window == target.Window &&
        (target.Name == "Chart Area" || layoutTarget.Name == target.Name) &&
        (target.ProcessId is null or 0 || layoutTarget.ProcessId == target.ProcessId) &&
        step.WhenUser == next.WhenUser;

    private static bool IsTransientChartSearchAction(PlanStep step, string workbook, string chartName)
    {
        if (step.Target is not { Process: "EXCEL", Window: { } stepWorkbook } ||
            !string.Equals(stepWorkbook, workbook, StringComparison.OrdinalIgnoreCase) ||
            step.OriginIntent is not null)
            return false;
        if (step.Action == "scroll") return string.IsNullOrWhiteSpace(step.Intent);
        if (step.Action == "click" &&
            (step.Target.ControlType == "ControlType.ScrollBar" ||
             step.Target.ClassName == "NetUIRepeatButton" && step.Target.Name is "Page left" or "Page right"))
            return string.IsNullOrWhiteSpace(step.Intent);
        if (step.Action is not ("click" or "context-click") ||
            step.Value is not null || step.Key is not null ||
            step.Target.Name != chartName ||
            step.Target.ControlType is not ("ControlType.Image" or "ControlType.Group"))
            return false;
        // A bare "previous" note describes the already captured chart gesture; it does not
        // authorize a separate pointer action before the verified native chart operation.
        return string.IsNullOrWhiteSpace(step.Intent) ||
            step.Intent.Trim().Equals("previous", StringComparison.OrdinalIgnoreCase);
    }

    private static string? CopyBoundaryState(string intent)
    {
        var values = Regex.Matches(intent, "\\bvalue\\s+\"([^\"]+)\"", RegexOptions.IgnoreCase)
            .Select(match => match.Groups[1].Value).Distinct(StringComparer.Ordinal).ToArray();
        return values.Length == 0 ? null : "copy-stop-values:" + System.Text.Json.JsonSerializer.Serialize(values);
    }

    internal static bool IsStatusBackgroundBeforeNavigation(PlanStep step, PlanStep? next) =>
        step is { Action: "click", Target: { ControlType: "ControlType.StatusBar",
                Process: { Length: > 0 }, Window: { Length: > 0 } },
            Value: null, Key: null, OriginIntent: null } &&
        string.IsNullOrWhiteSpace(step.Intent) && string.IsNullOrWhiteSpace(step.ExpectedState) &&
        string.IsNullOrWhiteSpace(step.TargetStrategy) &&
        next is { Action: "click", Target: { ControlType: "ControlType.TabItem",
            Name: { Length: > 0 } } } &&
        next.Target.Process == step.Target.Process && next.Target.Window == step.Target.Window;

    internal static bool IsTitleFocusBeforeCommand(PlanStep step, PlanStep? next) =>
        step is { Action: "click", Target: { ControlType: "ControlType.TitleBar",
            Process: { Length: > 0 } }, Value: null, Key: null } &&
        string.IsNullOrWhiteSpace(step.Intent) && string.IsNullOrWhiteSpace(step.ExpectedState) &&
        next is { Action: "click" or "optional-click" or "click-if-previous-absent",
            Target: { ControlType: "ControlType.Button" or "ControlType.MenuItem",
                Window: { Length: > 0 } } } &&
        string.Equals(step.Target.Process, next.Target.Process, StringComparison.OrdinalIgnoreCase) ||
        step is { Action: "click", Target: { ControlType: "ControlType.TitleBar",
                Process: { Length: > 0 }, Window: { Length: > 0 } },
            Value: null, Key: null, OriginIntent: null } &&
        string.IsNullOrWhiteSpace(step.Intent) && string.IsNullOrWhiteSpace(step.ExpectedState) &&
        string.IsNullOrWhiteSpace(step.TargetStrategy) &&
        next is { Action: "ensure-state", Target: { ControlType: "ControlType.HeaderItem",
                Window: { Length: > 0 } }, ExpectedState: "sort:ascending" or "sort:descending" } &&
        next.Target.Window == step.Target.Window &&
        string.Equals(step.Target.Process, next.Target.Process, StringComparison.OrdinalIgnoreCase);

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

    private static List<PlanStep> NormalizeColumnCopyIntent(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>();
        for (var index = 0; index < steps.Count; index++)
        {
            var step = steps[index];
            if (step.Action == "copy-populated-columns" &&
                step.ExpectedState?.StartsWith("excel-selection-range:", StringComparison.Ordinal) == true)
                step = step with { ExpectedState = null };
            if (step is { Action: "click", Target: { Name: "Select Data...", ControlType: "ControlType.MenuItem" },
                    Intent: { } chartIntent } &&
                chartIntent.Contains("replace the chart data source with the last pasted two-column range", StringComparison.OrdinalIgnoreCase))
            {
                var chartIndex = index + 1;
                while (chartIndex < steps.Count && steps[chartIndex].Action == "scroll") chartIndex++;
                if (index > 0 &&
                    steps[index - 1] is { Action: "context-click",
                        Target: { Name: { Length: > 0 }, ClassName: "ExcelChartObject" } capturedChart } &&
                    capturedChart.Process == step.Target.Process && capturedChart.Window == step.Target.Window &&
                    steps[index - 1].WhenUser == step.WhenUser &&
                    chartIndex + 1 < steps.Count &&
                    steps[chartIndex] is { Action: "click", Target: { ClassName: "XLSpreadsheetCell" } chartSourceCell } &&
                    chartSourceCell.Process == step.Target.Process && chartSourceCell.Window == step.Target.Window &&
                    steps[chartIndex + 1] is { Action: "click", Target: { Window: "Select Data Source", Name: "OK" } ok } &&
                    ok.Process == step.Target.Process &&
                    steps[chartIndex].WhenUser == step.WhenUser && steps[chartIndex + 1].WhenUser == step.WhenUser)
                {
                    result.RemoveAt(result.Count - 1);
                    result.Add(step with { Action = "set-chart-source-from-last-paste", Target = capturedChart,
                        Value = capturedChart.Name, OriginIntent = chartIntent,
                        Explanation = "Replace the captured chart's source with every row of the last verified two-column paste." });
                    index = chartIndex + 1;
                    continue;
                }
                var confirmedChart = Regex.Match(chartIntent, "\\btarget\\s+chart\\s+\"([^\"]+)\"", RegexOptions.IgnoreCase);
                if (confirmedChart.Success && chartIndex + 1 < steps.Count &&
                    steps[chartIndex] is { Action: "manual-click", Target: { ControlType: "ControlType.DataItem" } chartSelection } &&
                    chartSelection.Process == step.Target.Process && chartSelection.Window == step.Target.Window &&
                    steps[chartIndex + 1] is { Action: "click", Target: { Window: "Select Data Source", Name: "OK" } confirmation } &&
                    confirmation.Process == step.Target.Process &&
                    index > 0 && steps[index - 1] is { Action: "context-click", Target: { ControlType: "ControlType.Image" } context } &&
                    context.Process == step.Target.Process && context.Window == step.Target.Window)
                {
                    var layoutIndex = chartIndex + 2;
                    while (layoutIndex < steps.Count && steps[layoutIndex].Action == "scroll") layoutIndex++;
                    if (layoutIndex < steps.Count &&
                        steps[layoutIndex] is { Action: "set-chart-legend-layout", Target: { } identified } &&
                        identified.Process == step.Target.Process && identified.Window == step.Target.Window &&
                        identified.Name == confirmedChart.Groups[1].Value)
                    {
                        result.RemoveAt(result.Count - 1);
                        result.Add(step with { Action = "set-chart-source-from-last-paste", Target = identified,
                            Value = identified.Name, OriginIntent = chartIntent,
                            Explanation = "User-confirmed chart identity: replace its source with every row of the last verified two-column paste." });
                        index = chartIndex + 1;
                        continue;
                    }
                }
                if (chartIndex + 1 >= steps.Count ||
                    steps[chartIndex] is not { Action: "click", Target: { Name: { Length: > 0 }, ControlType: "ControlType.Image" } chartTarget } ||
                    chartTarget.Process != step.Target.Process || chartTarget.Window != step.Target.Window ||
                    steps[chartIndex + 1] is not { Action: "click", Target: { Window: "Select Data Source", Name: "OK" } })
                    throw new InvalidDataException("Chart source intent requires a recorded chart identity and Select Data Source confirmation.");
                result.Add(step with { Action = "set-chart-source-from-last-paste", Target = chartTarget,
                    Value = chartTarget.Name, OriginIntent = chartIntent,
                    Explanation = "Replace this chart's data source with the preceding verified two-column paste, including every pasted row." });
                index = chartIndex + 1;
                continue;
            }
            if (step is { Action: "key", Key: "Control+C" or "Ctrl+C", Intent: { } columnsIntent,
                    Target: { Process: { Length: > 0 } } columnsTarget } &&
                Regex.Match(columnsIntent, @"\bcopy\s+([A-Z]{1,3})\s+and\s+([A-Z]{1,3})\s+cells\s+that\s+contain\s+a\s+value\b",
                    RegexOptions.IgnoreCase) is { Success: true } columns &&
                Regex.Match(columnsTarget.AutomationId ?? columnsTarget.Name ?? "",
                    @"^([A-Z]{1,3})([1-9]\d*)$", RegexOptions.IgnoreCase) is { Success: true } firstCell &&
                firstCell.Groups[1].Value.Equals(columns.Groups[1].Value, StringComparison.OrdinalIgnoreCase))
            {
                var row = firstCell.Groups[2].Value;
                result.Add(step with { Action = "copy-populated-columns", Key = null,
                    Value = $"{columns.Groups[1].Value.ToUpperInvariant()}{row}:{columns.Groups[2].Value.ToUpperInvariant()}{row}",
                    OriginIntent = columnsIntent, ExpectedState = null,
                    Explanation = "Copy the rectangular range starting at the recorded row; exclude the first row with an empty cell in any selected column." });
                continue;
            }
            if (step.Action == "copy-column-until-empty" && step.Intent is { } boundaryIntent &&
                CopyBoundaryState(boundaryIntent) is { } boundary)
                step = step with { ExpectedState = boundary,
                    Explanation = "Copy downward, excluding the first empty cell or explicitly named stop value." };
            if (step is { Action: "key", Key: "Control+C" or "Ctrl+C",
                Target: { Process: { Length: > 0 } }, Intent: { } copyIntent } &&
                Regex.Match(copyIntent, @"\bcopy\s+(?:([A-Z]{1,3}[1-9]\d*)\s+down\s+to\s+the\s+first\s+empty\s+cell\b|column\s+([A-Z]{1,3})\s+until\s+before\s+(?:the\s+)?cell\s+that\s+contains\s+empty\s+value\b)",
                    RegexOptions.IgnoreCase) is { Success: true } directCopy &&
                index > 0 && steps[index - 1].Target is { Process: { Length: > 0 } } sourceCell &&
                sourceCell.Process == step.Target.Process &&
                sourceCell.Window == step.Target.Window)
            {
                var copyAddress = directCopy.Groups[1].Value.ToUpperInvariant();
                if (copyAddress.Length == 0)
                {
                    var sourceAddress = Regex.Match(sourceCell.AutomationId ?? sourceCell.Name ?? "",
                        @"^[A-Z]{1,3}([1-9]\d*)$", RegexOptions.IgnoreCase);
                    if (!sourceAddress.Success)
                        throw new InvalidDataException("Column-copy intent requires a recorded starting cell.");
                    copyAddress = directCopy.Groups[2].Value.ToUpperInvariant() + sourceAddress.Groups[1].Value;
                }
                result.Add(step with { Action = "copy-column-until-empty", Value = copyAddress, Key = null,
                    Target = sourceCell with { Name = copyAddress, AutomationId = copyAddress },
                    OriginIntent = copyIntent,
                    ExpectedState = CopyBoundaryState(copyIntent),
                    Explanation = $"Copy {copyAddress} downward, excluding the first empty cell or explicitly named stop value." });
                continue;
            }
            if (step is { Action: "resize-column", Target: { Process: { Length: > 0 },
                    ControlType: "ControlType.DataItem" }, Intent: { } fillIntent } &&
                Regex.Match(fillIntent, @"\b(?:function|formula)\s+in\s+column\s+([A-Z]{1,3}).*?\bempty\s+cell\s+in(?:\s+of)?\s+column\s+([A-Z]{1,3})\b",
                    RegexOptions.IgnoreCase) is { Success: true } formula &&
                fillIntent.Contains("last populated", StringComparison.OrdinalIgnoreCase))
            {
                var column = formula.Groups[1].Value.ToUpperInvariant();
                result.Add(step with { Action = "extend-formula-to-adjacent-data-end",
                    Target = step.Target with { Name = column + "1", AutomationId = column + "1" },
                    Value = column, Key = null,
                    ExpectedState = "adjacent-column:" + formula.Groups[2].Value.ToUpperInvariant(),
                    TargetStrategy = "last-populated-formula", OriginIntent = fillIntent,
                    Explanation = "Preserve existing formulas and extend from the last populated formula cell to before the first empty adjacent cell." });
                continue;
            }
            result.Add(step);
            if (step is not { Action: "click", Target: { Process: "EXCEL",
                ControlType: "ControlType.TabItem" }, Intent: { } intent }) continue;
            var match = Regex.Match(intent,
                @"\bcopy\s+column\s+([A-Z]{1,3})\s+until\s+before\s+(?:the\s+)?first\s+empty\s+cell\b",
                RegexOptions.IgnoreCase);
            if (!match.Success) continue;
            var cursor = index + 1;
            while (cursor < steps.Count && steps[cursor] is { Action: "scroll", Intent: null,
                OriginIntent: null, Target: { Process: "EXCEL" } scroll } &&
                scroll.Window == step.Target.Window) cursor++;
            if (cursor + 1 >= steps.Count) continue;
            var selection = steps[cursor];
            var copy = steps[cursor + 1];
            if (selection is not { Action: "click", Intent: null,
                Target: { Process: "EXCEL", ControlType: "ControlType.DataItem" } cell } ||
                cell.Window != step.Target.Window ||
                copy is not { Action: "key", Key: "Control+C" or "Ctrl+C",
                    Target: { Process: "EXCEL" } copyTarget } ||
                copyTarget.Window != step.Target.Window) continue;
            var address = Regex.Match(cell.AutomationId ?? "", @"^[A-Z]{1,3}([1-9]\d*)$", RegexOptions.IgnoreCase);
            if (!address.Success)
                address = Regex.Match(cell.Name ?? "", @"^[A-Z]{1,3}([1-9]\d*)$", RegexOptions.IgnoreCase);
            if (!address.Success) continue;
            var start = match.Groups[1].Value.ToUpperInvariant() + address.Groups[1].Value;
            result.Add(copy with
            {
                Action = "copy-column-until-empty",
                Target = cell with { Name = start, AutomationId = start, ClassName = "XLSpreadsheetCell" },
                Value = start, Key = null, Intent = intent, OriginIntent = intent,
                ExpectedState = CopyBoundaryState(intent),
                Explanation = $"Copy {start} downward, stopping before the first empty cell. The column comes from explicit intent and the starting row from the recorded selection."
            });
            index = cursor + 1;
        }
        return result;
    }

    private static List<PlanStep> NormalizeSubmittedKeyTargets(IReadOnlyList<PlanStep> steps)
    {
        var result = steps.ToList();
        for (var index = 1; index < result.Count; index++)
        {
            var submitted = result[index];
            var typed = result[index - 1];
            if (submitted is not { Action: "key", Key: "Enter" or "Return", Target: { } submittedTarget } ||
                typed is not { Action: "type", Value: not null, Target: { } typedTarget } ||
                typedTarget.ControlType is not ("ControlType.Edit" or "ControlType.ComboBox") ||
                !string.Equals(typedTarget.Process, submittedTarget.Process, StringComparison.OrdinalIgnoreCase) ||
                typed.WhenUser != submitted.WhenUser)
                continue;

            // UI Automation can finish describing Enter after navigation has already replaced the
            // focused editor. The immediately preceding typed field is the only control that could
            // have received this submit key, so retain that frozen identity for deterministic replay.
            result[index] = submitted with { Target = typedTarget };
        }
        return result;
    }

    private static List<PlanStep> NormalizePrivateWindowTransitions(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>(steps.Count);
        for (var index = 0; index < steps.Count; index++)
        {
            var opener = steps[index];
            if (index + 2 < steps.Count && opener is { Action: "click", Target: { } openerTarget } &&
                steps[index + 1] is { Action: "click", Target:
                    { ControlType: "ControlType.Document", Window: { Length: > 0 } privateWindow } } arrival &&
                IsPrivateBrowserWindow(privateWindow) &&
                string.Equals(arrival.Target!.Process, openerTarget.Process, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(arrival.Target.Window, openerTarget.Window, StringComparison.OrdinalIgnoreCase) &&
                steps[index + 2].Target is { } following &&
                string.Equals(following.Process, openerTarget.Process, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(following.Window, privateWindow, StringComparison.OrdinalIgnoreCase) &&
                opener.WhenUser == arrival.WhenUser && arrival.WhenUser == steps[index + 2].WhenUser)
            {
                // A delayed capture often sees the new private document instead of the transient
                // menu command. The following action in that window verifies the transition.
                result.Add(opener with { Action = "open-private-window",
                    ExpectedState = "opened-window:" + privateWindow,
                    Explanation = "Open and verify the recorded private browser window." });
                index++;
                continue;
            }
            result.Add(opener);
        }
        return result;
    }

    private static List<PlanStep> NormalizePostNavigationDraftTyping(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>(steps.Count);
        for (var index = 0; index < steps.Count; index++)
        {
            var draft = steps[index];
            if (index > 0 && index + 3 < steps.Count &&
                steps[index - 1] is { Action: "key", Key: "Enter" or "Return" } &&
                draft is { Action: "type", Target: { } delayedTarget, OriginIntent: null } &&
                delayedTarget.ControlType is not ("ControlType.Edit" or "ControlType.ComboBox") &&
                steps[index + 1] is { Action: "click", Target: { } editor } &&
                editor.ControlType is "ControlType.Edit" or "ControlType.ComboBox" &&
                steps[index + 2] is { Action: "type", Target: { } typedEditor } &&
                steps[index + 3] is { Action: "key", Key: "Enter" or "Return", Target: { } submittedEditor } &&
                SameInputTarget(editor, typedEditor) && SameInputTarget(editor, submittedEditor) &&
                string.Equals(delayedTarget.Process, editor.Process, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(delayedTarget.Window, editor.Window, StringComparison.OrdinalIgnoreCase) &&
                draft.WhenUser == steps[index + 1].WhenUser &&
                draft.WhenUser == steps[index + 2].WhenUser && draft.WhenUser == steps[index + 3].WhenUser)
            {
                // A page loaded while the user began typing, so delayed UIA described the early
                // characters on a non-edit node. The later clicked editor, final text and Enter are
                // complete evidence; replaying the transient draft would concatenate both values.
                continue;
            }
            result.Add(draft);
        }
        return result;
    }

    private static bool SameInputTarget(ControlRef left, ControlRef right) =>
        string.Equals(left.Process, right.Process, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(left.Window, right.Window, StringComparison.OrdinalIgnoreCase) &&
        left.ControlType == right.ControlType && left.AutomationId == right.AutomationId &&
        (left.AutomationId?.Length > 0 || left.Name == right.Name) && left.ClassName == right.ClassName;

    private static bool IsPrivateBrowserWindow(string title) =>
        title.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase) ||
        title.Contains("Incognito", StringComparison.OrdinalIgnoreCase) ||
        title.Contains("Private Browsing", StringComparison.OrdinalIgnoreCase);

    private static List<PlanStep> NormalizeWorksheetGestureIntent(IReadOnlyList<PlanStep> steps)
    {
        var result = steps.ToList();
        for (var index = 1; index < result.Count; index++)
        {
            var copy = result[index];
            if (copy is not { Action: "key", Key: "Control+C" or "Ctrl+C",
                    Target: { Process: "EXCEL" } copyTarget, ExpectedState: { } rangeState }) continue;
            var range = Regex.Match(rangeState, @"^excel-selection-range:([A-Z]{1,3})([1-9]\d*):\1[1-9]\d*$");
            var rectangle = Regex.Match(rangeState, @"^excel-selection-range:([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})[1-9]\d*$");
            if (rectangle.Success && rectangle.Groups[1].Value != rectangle.Groups[3].Value)
            {
                for (var cursor = index - 1; cursor >= Math.Max(0, index - 6); cursor--)
                {
                    var preceding = result[cursor];
                    var dialogNote = preceding.Target is { Window: "Microsoft Excel", Name: "OK",
                        ControlType: "ControlType.Button", ClassName: "NetUIButton", ProcessId: > 0 } acknowledgment &&
                        acknowledgment.ProcessId == copyTarget.ProcessId &&
                        result.Take(cursor).TakeLast(6).LastOrDefault(candidate =>
                            candidate.Target?.Process == copyTarget.Process &&
                            (candidate.Target.ClassName?.StartsWith("NetUIRibbon", StringComparison.Ordinal) == true ||
                             candidate.Target.ClassName?.StartsWith("XL", StringComparison.Ordinal) == true))
                            ?.Target?.Window == copyTarget.Window;
                    if (preceding.Target?.Process != copyTarget.Process ||
                        preceding.Target.Window != copyTarget.Window && !dialogNote ||
                        preceding.WhenUser != copy.WhenUser ||
                        preceding.Action is not ("click" or "scroll") ||
                        preceding.Target.ControlType == "ControlType.TabItem") break;
                    if (preceding.Intent is not { Length: > 0 } intent) continue;
                    var columns = Regex.Match(intent,
                        @"\bcopy\s+([A-Z]{1,3})\s+and\s+([A-Z]{1,3})\s+cells\s+that\s+contain\s+a\s+value\b",
                        RegexOptions.IgnoreCase);
                    if (columns.Success &&
                        columns.Groups[1].Value.Equals(rectangle.Groups[1].Value, StringComparison.OrdinalIgnoreCase) &&
                        columns.Groups[2].Value.Equals(rectangle.Groups[3].Value, StringComparison.OrdinalIgnoreCase))
                        result[index] = copy with { Intent = intent };
                    break;
                }
            }
            if (!range.Success) continue;
            for (var cursor = index - 1; cursor >= Math.Max(0, index - 6); cursor--)
            {
                var preceding = result[cursor];
                if (preceding.Target?.Process != copyTarget.Process || preceding.Target.Window != copyTarget.Window ||
                    !(preceding.Action is "click" or "scroll" || preceding.Action == "key" && preceding.Key == "Delete"))
                    break;
                if (preceding.Intent is not { Length: > 0 } intent) continue;
                var column = Regex.Match(intent,
                    @"\bcopy\s+column\s+([A-Z]{1,3})\s+until\s+before\s+(?:(?:the\s+)?first\s+empty\s+cell|(?:the\s+)?cell\s+that\s+contains\s+empty\s+value)\b",
                    RegexOptions.IgnoreCase);
                if (column.Success && column.Groups[1].Value.Equals(range.Groups[1].Value, StringComparison.OrdinalIgnoreCase))
                {
                    var start = range.Groups[1].Value + range.Groups[2].Value;
                    result[index] = copy with { Action = "copy-column-until-empty", Key = null, Value = start,
                        Target = copyTarget with { Name = start, AutomationId = start, ControlType = "ControlType.DataItem",
                            ClassName = "XLSpreadsheetCell" },
                        Intent = intent, OriginIntent = intent, ExpectedState = CopyBoundaryState(intent),
                        Explanation = "Copy from the verified selection's starting row to before the first empty cell or explicit stop value; do not freeze its recorded endpoint." };
                    break;
                }
            }
        }
        for (var index = 0; index < result.Count; index++)
        {
            var pasteMenu = result[index];
            if (pasteMenu is not { Action: "click", Target: { Process: "EXCEL",
                    AutomationId: "PasteMenu_Dropdown", Window: { Length: > 0 } workbook } } ||
                index + 1 < result.Count && result[index + 1] is
                    { Action: "click", Target: { Process: "EXCEL", Name: "Values",
                        ParentName: "Paste Values" } } ||
                Executor.FindPasteDestination(result, index) is null ||
                !result.Take(index).Any(candidate => candidate.Action is
                    "copy-column-until-empty" or "copy-populated-columns" &&
                    candidate.Target?.Process == "EXCEL" && candidate.Target.Window == workbook))
                continue;
            var valuesTarget = new ControlRef("EXCEL", workbook, null, "Values",
                "ControlType.Button", "NetUIRibbonButton", "Paste Values");
            result.Insert(index + 1, new PlanStep(0, "click", valuesTarget, null, null,
                "paste-values-verified", pasteMenu.Screenshot,
                "Paste the preceding verified Excel copy as values at the selected destination.",
                pasteMenu.Intent, WhenUser: pasteMenu.WhenUser, OriginIntent: pasteMenu.OriginIntent));
            index++;
        }
        for (var index = 0; index < result.Count; index++)
        {
            var step = result[index];
            if (step is not { Action: "manual-click" or "click", Target: { Process: "EXCEL",
                    ControlType: "ControlType.DataItem" } target }) continue;
            var address = Regex.Match(target.AutomationId ?? target.Name ?? "", @"^([A-Z]{1,3})([1-9]\d*)$",
                RegexOptions.IgnoreCase);
            if (!address.Success) continue;
            var intent = step.Intent;
            var alreadyExtended = false;
            for (var cursor = index - 1; cursor >= Math.Max(0, index - 6); cursor--)
            {
                var preceding = result[cursor];
                if (preceding.Target?.Process != target.Process || preceding.Target.Window != target.Window ||
                    preceding.WhenUser != step.WhenUser) break;
                if (preceding.Action == "extend-formula-to-adjacent-data-end" &&
                    preceding.Value == address.Groups[1].Value.ToUpperInvariant() &&
                    (string.IsNullOrWhiteSpace(intent) || intent == preceding.OriginIntent))
                {
                    alreadyExtended = true;
                    break;
                }
                if (!(preceding.Action == "scroll" ||
                      preceding.Action == "click" && (preceding.Target.ControlType == "ControlType.DataItem" &&
                          preceding.Target.Name == target.Name ||
                          preceding.Target is { Name: "Values", ParentName: "Paste Values" }))) break;
                if (string.IsNullOrWhiteSpace(intent) && !string.IsNullOrWhiteSpace(preceding.Intent))
                    intent = preceding.Intent;
            }
            if (intent is null || alreadyExtended) continue;
            var fixedSource = Regex.Match(intent,
                @"\bcopy\s+the\s+(?:function|formula)\s+from\s+the\s+([A-Z]{1,3}[1-9]\d*)\s+cell\s+in\s+column\s+([A-Z]{1,3})\s+to\s+the\s+last\s+cell\s+that\s+has\s+an\s+adjacent\s+existing\s+value\s+in\s+column\s+([A-Z]{1,3})\b",
                RegexOptions.IgnoreCase);
            var extension = Regex.Match(intent,
                @"\bcopy\s+the\s+(?:function|formula)\s+in\s+column\s+([A-Z]{1,3})\s+(?:till|until)\s+just\s+before\s+the\s+empty\s+cell\s+in(?:\s+of)?\s+column\s+([A-Z]{1,3})\b",
                RegexOptions.IgnoreCase);
            var column = fixedSource.Success ? fixedSource.Groups[2].Value : extension.Groups[1].Value;
            var adjacent = fixedSource.Success ? fixedSource.Groups[3].Value : extension.Groups[2].Value;
            if ((!fixedSource.Success && !extension.Success) ||
                !address.Groups[1].Value.Equals(column, StringComparison.OrdinalIgnoreCase) ||
                column.Equals(adjacent, StringComparison.OrdinalIgnoreCase)) continue;
            if (fixedSource.Success &&
                !Regex.Match(fixedSource.Groups[1].Value, @"^[A-Z]{1,3}", RegexOptions.IgnoreCase).Value
                    .Equals(column, StringComparison.OrdinalIgnoreCase)) continue;
            var source = fixedSource.Success ? fixedSource.Groups[1].Value.ToUpperInvariant() : column.ToUpperInvariant() + "1";
            result[index] = step with { Action = "extend-formula-to-adjacent-data-end",
                Target = target with { Name = source, AutomationId = source, ClassName = "XLSpreadsheetCell" },
                Value = column.ToUpperInvariant(), Key = null,
                ExpectedState = "adjacent-column:" + adjacent.ToUpperInvariant(),
                TargetStrategy = fixedSource.Success ? "formula-source:" + source : "last-populated-formula",
                Intent = intent, OriginIntent = intent,
                Explanation = fixedSource.Success
                    ? $"Fill the formula from {source} through the contiguous adjacent data, verifying relative references."
                    : "Preserve existing formulas and extend the last populated formula through the contiguous adjacent data, verifying relative references." };
        }
        return result;
    }

    public static ExecutionPlan Compact(ExecutionPlan plan)
    {
        var deduplicated = NormalizeDuplicateFormulaExtensions(NormalizeFormulaFillGestureBeforeDataCommand(
            NormalizeTransientMenuOwners(NormalizePrivateWindowTransitions(
                NormalizePostNavigationDraftTyping(NormalizeSubmittedKeyTargets(plan.Steps))))));
        var source = NormalizeMissingChartLegendDrag(NormalizeDuplicatedRowDate(NormalizeDuplicatedRowGesture(
            NormalizeColumnCopyIntent(NormalizeWorksheetGestureIntent(
                NormalizeSpreadsheetEdits(NormalizeSearchEditing(deduplicated)))))));
        source = NormalizeDuplicateFormulaExtensions(source);
        var compact = new List<PlanStep>();
        for (var i = 0; i < source.Count; i++)
        {
            var step = source[i];
            if (step is { Action: "click", ExpectedState: "worksheet-tab-focus-only",
                    Target: { Process: "EXCEL" } })
                continue; // Verified following sheet tab supersedes this workbook-surface focus.
            if (IsChartSelectionBeforeLayout(step, i + 1 < source.Count ? source[i + 1] : null))
                continue;
            if (IsStatusBackgroundBeforeNavigation(step, i + 1 < source.Count ? source[i + 1] : null))
                continue;
            if (IsTitleFocusBeforeCommand(step, i + 1 < source.Count ? source[i + 1] : null))
                continue;
            if (i + 1 < source.Count && IsSupersededExcelCellClick(step, source[i + 1]))
                continue;
            if (step.Action == "manual-click" && step.Target is
                { Process: "explorer", ControlType: "ControlType.Pane",
                  ClassName: "Shell_SecondaryTrayWnd" } &&
                i + 1 < source.Count &&
                !source[i + 1].Target?.Process?.Equals("explorer",
                    StringComparison.OrdinalIgnoreCase) == true)
                continue; // Switching applications is implicit in resolving the next stable target.
            if (step.Action == "manual-click" && step.Target?.ControlType == "ControlType.ScrollBar" &&
                step.ExpectedState?.StartsWith("range:", StringComparison.OrdinalIgnoreCase) == true &&
                step.Key is "Vertical" or "Horizontal")
            {
                compact.Add(step with { Action = "scroll",
                    Explanation = $"Restore the recorded {step.Key.ToLowerInvariant()} scroll state." });
                continue;
            }
            if (step.Action == "manual-click" && step.Target is { } stableTarget &&
                (!string.IsNullOrWhiteSpace(stableTarget.AutomationId) ||
                 step.ExpectedState == "msaa-command") &&
                stableTarget.ControlType is "ControlType.Button" or "ControlType.CheckBox" or
                    "ControlType.RadioButton" or "ControlType.MenuItem" or "ControlType.Edit" or
                    "ControlType.Pane")
            {
                compact.Add(step with { Action = "click",
                    Explanation = "Activate the recorded control using its stable accessibility identity." });
                continue;
            }
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
            if (step.Action is "click" or "manual-click" && step.ExpectedState == "window-closed" &&
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
            if (step.Action is "click" or "double-click" &&
                step.Target?.ControlType == "ControlType.TabItem" && i + 1 < source.Count &&
                (source[i + 1].Action is "type" or "key" || IsInlineSheetEditClick(source[i + 1])))
            {
                var end = i + 1;
                while (end < source.Count && (source[end].Action is "type" or "key" || IsInlineSheetEditClick(source[end])) &&
                    source[end].Target?.Process?.Equals(step.Target.Process, StringComparison.OrdinalIgnoreCase) == true &&
                    source[end].Target?.Window == step.Target.Window)
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
                    weekday ??= RelativeWeekdayFromIntent(intent);
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
            if (current is { Action: "key", Key: "Shift+End", Intent: null, OriginIntent: null,
                    ExpectedState: null, Target: { ControlType: "ControlType.ListItem",
                        Process: { Length: > 0 }, Window: { Length: > 0 } } selectionRow } &&
                prior is { Action: "ensure-state", Target: { ControlType: "ControlType.HeaderItem" } sortHeader } &&
                prior.ExpectedState?.StartsWith("sort:", StringComparison.OrdinalIgnoreCase) == true &&
                next is { Action: "select-all-items", Target: { ControlType: "ControlType.HeaderItem" } selectionHeader } &&
                selectionRow.Process == sortHeader.Process && selectionRow.Window == sortHeader.Window &&
                selectionHeader == sortHeader)
                continue; // The verified all-row operation establishes focus and replaces this raw selection key.
            if (current is { Action: "key", Key: "Shift+End" or "End", Intent: null, OriginIntent: null,
                    ExpectedState: null,
                    Target: { ControlType: "ControlType.ListItem" } row } &&
                prior is { Action: "ensure-state", Target: { ControlType: "ControlType.HeaderItem" } header } &&
                prior.ExpectedState?.StartsWith("sort:", StringComparison.OrdinalIgnoreCase) == true &&
                row.Process == header.Process && row.Window == header.Window)
            {
                var contextIndex = i + 1;
                if (next is { Action: "click", Target: { ControlType: "ControlType.Text" } cell } &&
                    cell.Process == row.Process && cell.Window == row.Window &&
                    (cell.ParentName == row.Name || current.Key == "End") &&
                    next.OriginIntent is null && string.IsNullOrWhiteSpace(next.Intent) &&
                    string.IsNullOrWhiteSpace(next.ExpectedState))
                    contextIndex++;
                if (contextIndex < purposeful.Count &&
                    purposeful[contextIndex] is { Action: "context-selection",
                        ExpectedState: "selection-intent:all", Target: { } context } &&
                    context.Process == row.Process && context.Window == row.Window)
                {
                    final.Add(current with { Action = "select-all-items", Target = header,
                        Key = null, ExpectedState = null,
                        Explanation = "Select and verify every current row in the sorted list before opening the recorded all-row context menu." });
                    i = contextIndex - 1;
                    continue;
                }
            }
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.Text" &&
                prior is { Action: "ensure-state", Target: { ControlType: "ControlType.HeaderItem" } sortedHeader } &&
                prior.ExpectedState?.StartsWith("sort:", StringComparison.OrdinalIgnoreCase) == true &&
                next is { Action: "key", Key: "End" or "Shift+End",
                    Target: { ControlType: "ControlType.ListItem" } endRow } &&
                i + 2 < purposeful.Count && purposeful[i + 2] is
                    { Action: "context-selection", ExpectedState: "selection-intent:all", Target: { } allRows } &&
                current.Target.Process == sortedHeader.Process && current.Target.Window == sortedHeader.Window &&
                endRow.Process == sortedHeader.Process && endRow.Window == sortedHeader.Window &&
                allRows.Process == sortedHeader.Process && allRows.Window == sortedHeader.Window &&
                current.OriginIntent is null && string.IsNullOrWhiteSpace(current.Intent) &&
                string.IsNullOrWhiteSpace(current.ExpectedState))
            {
                // The clicked row text is live business data and changes between refreshes. The
                // following End plus verified all-row context proves that it was only a selection
                // anchor, so select the current sorted list through UI Automation instead.
                final.Add(next with { Action = "select-all-items", Target = sortedHeader,
                    Key = null, ExpectedState = null,
                    Explanation = "Select and verify every current row in the sorted list without relying on the recorded row text." });
                i++;
                continue;
            }
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
            if (button.OriginIntent is not null) continue;
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
        final = EnsureStandingIntentReplacements(final);
        for (var index = final.Count - 1; index > 0; index--)
        {
            var previous = index - 1;
            if (final[previous] is { Action: "close-window", Target: { Process: "EXCEL", Window: "Find and Replace" } })
                previous--;
            if (previous >= 0 && IsReplacementAcknowledgement(final[index], final[previous]))
                final.RemoveAt(index);
        }
        for (var index = final.Count - 1; index > 0; index--)
        {
            var added = final[index];
            if (added.Action != "replace-all" || added.OriginIntent is null) continue;
            var start = index;
            var end = index;
            while (start > 0 && final[start - 1].Action == "replace-all") start--;
            while (end + 1 < final.Count && final[end + 1].Action == "replace-all") end++;
            if (final.Skip(start).Take(end - start + 1).Any(recorded =>
                recorded.OriginIntent is null && added.Target?.Process == recorded.Target?.Process &&
                added.Target?.Window == recorded.Target?.Window && added.Value == recorded.Value &&
                added.ExpectedState == recorded.ExpectedState && added.WhenUser == recorded.WhenUser))
                final.RemoveAt(index);
        }
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
            if (following.Count == 0 || !following.Any(next => next.Action is "type" or "key" &&
                    next.Target is { ControlType: "ControlType.Edit" or "ControlType.Pane",
                        AutomationId: { Length: > 0 } } edit &&
                    edit.Process?.Equals(windowClick.Process, StringComparison.OrdinalIgnoreCase) == true &&
                    edit.Window?.Equals(windowName, StringComparison.OrdinalIgnoreCase) == true) ||
                following.TakeWhile(next => next.Action is not ("type" or "key")).Any(next =>
                    next.Action != "manual-click" || next.Target?.ControlType != "ControlType.Window" ||
                    next.Target.Name != windowName || next.Target.Process != windowClick.Process))
                continue;
            final.RemoveAt(index);
        }
        for (var index = 0; index < final.Count; index++)
        {
            var header = final[index];
            if (header.Action is not ("click" or "ensure-state") ||
                header.Target?.ControlType != "ControlType.HeaderItem" ||
                header.Target.ClassName == "XLSelectAllHeader" || header.Action == "ensure-state" &&
                header.ExpectedState?.Contains("sort", StringComparison.OrdinalIgnoreCase) != true) continue;
            var end = index;
            while (end + 1 < final.Count && final[end + 1].Action is "click" or "ensure-state" &&
                final[end + 1].Target == header.Target) end++;
            var last = final[end];
            var state = last.ExpectedState;
            if (state?.Contains("ascending", StringComparison.OrdinalIgnoreCase) != true &&
                state?.Contains("descending", StringComparison.OrdinalIgnoreCase) != true)
            {
                var note = string.Join(" ", final.Skip(index).Take(end - index + 1).Select(step => step.Intent));
                var direction = Regex.Match(note, @"\b(ascending|descending)\b", RegexOptions.IgnoreCase);
                state = direction.Success ? "sort:" + direction.Value.ToLowerInvariant() : "sort:unspecified";
            }
            final[index] = last with { Action = "ensure-state", ExpectedState = state };
            if (end > index) final.RemoveRange(index + 1, end - index);
        }
        RecoverExcelColumnFilterOpeners(final);
        var outcomes = ExcelFilterPlan.Compact(final);
        for (var sourceIndex = outcomes.Count - 1; sourceIndex >= 0; sourceIndex--)
        {
            if (outcomes[sourceIndex] is not { Action: "set-chart-source-range",
                    ExpectedState: "chart-source-verified",
                    Target: { Process: "EXCEL", Window: { } workbook, Name: { Length: > 0 } chartName } })
                continue;
            var navigationStart = sourceIndex;
            while (navigationStart > 0 && IsTransientChartSearchAction(
                       outcomes[navigationStart - 1], workbook, chartName))
                navigationStart--;
            if (navigationStart == sourceIndex) continue;
            // Native chart source replay resolves the chart by its verified workbook/name identity.
            // Pointer clicks and viewport movement used only to reach that chart are neither needed
            // nor reliable after row counts, zoom, window size or scroll position have changed.
            outcomes.RemoveRange(navigationStart, sourceIndex - navigationStart);
            sourceIndex = navigationStart;
        }
        for (var sourceIndex = outcomes.Count - 2; sourceIndex >= 0; sourceIndex--)
        {
            if (outcomes[sourceIndex] is not { Action: "set-chart-source-range",
                    Target: { Process: "EXCEL", Window: { } workbook, Name: { Length: > 0 } chartName } })
                continue;
            var layoutIndex = -1;
            for (var candidate = sourceIndex + 1;
                 candidate < Math.Min(outcomes.Count, sourceIndex + 9); candidate++)
            {
                if (outcomes[candidate] is { Action: "set-chart-legend-layout",
                        Target: { Process: "EXCEL", Window: { } layoutWorkbook, Name: { } layoutChart },
                        ExpectedState: { } layoutState } &&
                    layoutWorkbook == workbook && layoutChart == chartName &&
                    layoutState.StartsWith("chart-legend-sheet:", StringComparison.Ordinal))
                {
                    layoutIndex = candidate;
                    break;
                }
                var between = outcomes[candidate];
                if (between.Target is not { Process: "EXCEL", Window: { } candidateWorkbook } ||
                    candidateWorkbook != workbook || between.Intent is not null || between.OriginIntent is not null ||
                    between.Action != "scroll" &&
                    !(between.Action == "click" && between.Target.ControlType is
                        "ControlType.TabItem" or "ControlType.DataItem" or "ControlType.Image"))
                    break;
            }
            if (layoutIndex <= sourceIndex + 1)
                continue;
            var chartNavigation = outcomes.Skip(sourceIndex + 1).Take(layoutIndex - sourceIndex - 1).ToArray();
            if (!chartNavigation.Any(step => step.Target?.ControlType == "ControlType.TabItem") &&
                chartNavigation.Any(step => step.Action != "scroll"))
                continue;
            // Source selection and legend resizing are verified native chart operations. Any
            // untitled sheet/cell/scroll actions captured between them, or an all-scroll sequence,
            // are transient chart navigation and must not leave a different viewport for layout.
            outcomes.RemoveRange(sourceIndex + 1, layoutIndex - sourceIndex - 1);
        }
        for (var tabIndex = outcomes.Count - 2; tabIndex >= 1; tabIndex--)
        {
            if (outcomes[tabIndex] is not { Action: "click", Target: { Process: "EXCEL",
                    ControlType: "ControlType.TabItem", Window: { } tabWorkbook },
                    Explanation: { } explanation } ||
                !explanation.StartsWith("Open the recorded worksheet", StringComparison.Ordinal) ||
                outcomes[tabIndex + 1] is not { Action: "click" or "manual-click", Target: { Process: "EXCEL",
                    ClassName: "XLGridColumnHeader", Window: { } headerWorkbook } } ||
                !string.Equals(tabWorkbook, headerWorkbook, StringComparison.OrdinalIgnoreCase) ||
                !outcomes.Skip(tabIndex + 2).Take(4).Any(candidate =>
                    candidate is { Action: "type", TargetStrategy: "first-visible-filtered-row",
                        Target: { Window: { } candidateWorkbook } } &&
                    string.Equals(candidateWorkbook, tabWorkbook, StringComparison.OrdinalIgnoreCase)))
                continue;
            var filterIndex = tabIndex - 1;
            while (filterIndex >= Math.Max(0, tabIndex - 3) &&
                outcomes[filterIndex] is { Action: "click", Intent: { Length: > 0 } intent,
                    Target: { Process: "EXCEL", ControlType: "ControlType.Button",
                        ParentName: { Length: > 0 }, Window: null or "" } } &&
                Regex.IsMatch(intent, @"\b(?:doesn['’]?t|does\s+not)\s+always\s+appear\b",
                    RegexOptions.IgnoreCase))
                filterIndex--;
            if (filterIndex < 0 || outcomes[filterIndex] is not
                { Action: "filter-values", Target: { Process: "EXCEL", Window: { } workbook } } ||
                !string.Equals(workbook, tabWorkbook, StringComparison.OrdinalIgnoreCase))
                continue;
            // Remove only a navigation step previously synthesized by worksheet recovery. The
            // surrounding semantic filter and visible-row edit establish that Excel stayed on
            // the filtered sheet. An optional collaboration response may sit between them.
            outcomes.RemoveAt(tabIndex);
        }
        for (var index = outcomes.Count - 2; index >= 0; index--)
        {
            var opener = outcomes[index];
            var filter = outcomes[index + 1];
            if (opener is not { Action: "click", Target: { Process: "EXCEL",
                    AutomationId: "Dropdown", ControlType: "ControlType.MenuItem",
                    ParentName: { Length: > 0 } column } openerTarget } ||
                filter is not { Action: "filter-values", Target: { Process: "EXCEL",
                    ParentName: { Length: > 0 } filterColumn } filterTarget } ||
                !string.Equals(column, filterColumn, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(openerTarget.Process, filterTarget.Process, StringComparison.OrdinalIgnoreCase))
                continue;
            // filter-values opens and applies the Excel checklist itself. A preceding dropdown
            // click is redundant and can retain the title of a dialog that closed just before
            // the filter (for example Find and Replace), causing replay to search the wrong tree.
            // Matching the recorded column parent keeps unrelated ribbon menus intact.
            outcomes.RemoveAt(index);
        }
        for (var index = outcomes.Count - 2; index >= 0; index--)
        {
            if (outcomes[index] is not { Action: "click", Intent: null, OriginIntent: null,
                    Target: { Process: "EXCEL", ControlType: "ControlType.Button",
                        ClassName: "NetUIRepeatButton", Name: "Page left" or "Page right",
                        Window: { } workbook } } ||
                outcomes[index + 1] is not { Action: "type", TargetStrategy: "first-visible-filtered-row",
                    Target: { Process: "EXCEL", Window: { } targetWorkbook } } ||
                workbook != targetWorkbook)
                continue;
            outcomes.RemoveAt(index);
        }
        for (var index = 0; index + 1 < outcomes.Count; index++)
        {
            var optional = outcomes[index];
            var next = outcomes[index + 1];
            if (optional is not { Action: "click", Target: { ControlType: "ControlType.Button",
                    Window: { Length: > 0 } dialog, Process: { Length: > 0 } process }, Intent: { Length: > 0 } intent } ||
                !intent.Contains('"' + dialog + '"', StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(optional.Target.Name) ||
                !Regex.IsMatch(intent, @"(?<!\w)" + Regex.Escape(optional.Target.Name) + @"(?!\w)",
                    RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(intent, @"\bif\s+present\b", RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(intent, @"\brefresh\s+button\b", RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(intent, @"\botherwise\s*,?\s*(?:do\s+)?click\s+it\b", RegexOptions.IgnoreCase) ||
                next.Action == "click-if-previous-absent" ||
                next.Target is not { Window: { Length: > 0 } view } nextTarget ||
                nextTarget.Process != process || view == dialog)
                continue;
            outcomes.Insert(index + 1, optional with { Action = "click-if-previous-absent",
                Target = new ControlRef(process, view, null, "Refresh", "ControlType.Button", null, null),
                Intent = null, OriginIntent = intent, Screenshot = null, Value = null, Key = null,
                ExpectedState = null, TargetStrategy = "unique-refresh-command",
                Explanation = "Explicit intent: when the preceding dialog is absent, invoke the unique enabled Refresh button in the resulting view." });
            index++;
        }
        for (var index = outcomes.Count - 2; index >= 0; index--)
        {
            var navigation = outcomes[index];
            if (navigation.Target is not { Process: "EXCEL" } navTarget ||
                !string.IsNullOrWhiteSpace(navigation.Intent) ||
                navigation.OriginIntent is not null ||
                !(navigation.Action == "scroll" && navigation.Key is "Horizontal" or "Vertical" ||
                  navigation.Action == "click" && navTarget.ControlType == "ControlType.Button" &&
                  navTarget.ClassName == "NetUIRepeatButton" &&
                  navTarget.Name is "Page left" or "Page right") ||
                outcomes[index + 1] is not { Action: "filter-values", Target: { } filter } ||
                filter.Process != navTarget.Process || filter.Window != navTarget.Window)
                continue;
            outcomes.RemoveAt(index);
        }
        for (var index = 1; index < outcomes.Count; index++)
        {
            var fill = outcomes[index];
            var sourceStep = outcomes[index - 1];
            if (fill is not { Action: "click" or "manual-click", Target: { Process: "EXCEL",
                    ControlType: "ControlType.DataItem" } target, Intent: { Length: > 0 } intent } ||
                sourceStep is not { Action: "type", Value: { Length: > 0 },
                    Target: { Process: "EXCEL", ClassName: "XLSpreadsheetCell",
                        AutomationId: { Length: > 0 } } address } ||
                target.Window != address.Window ||
                (!string.IsNullOrWhiteSpace(fill.ExpectedState) &&
                 !fill.ExpectedState.StartsWith("adjacent-column:", StringComparison.Ordinal) &&
                 fill.ExpectedState != "unverified-click"))
                continue;
            var sourceColumn = Regex.Match(address.AutomationId, @"^([A-Z]{1,3})[1-9]\d*$",
                RegexOptions.IgnoreCase);
            var destination = Regex.Match(intent, @"\b(?:in\s+)?column\s+([A-Z]{1,3})\b",
                RegexOptions.IgnoreCase);
            var adjacent = Regex.Match(intent,
                @"\bempty\s+adjacent\s+cell\s+in\s+([A-Z]{1,3})\s+column\b",
                RegexOptions.IgnoreCase);
            if (!sourceColumn.Success || !destination.Success || !adjacent.Success ||
                !sourceColumn.Groups[1].Value.Equals(destination.Groups[1].Value,
                    StringComparison.OrdinalIgnoreCase) ||
                sourceColumn.Groups[1].Value.Equals(adjacent.Groups[1].Value,
                    StringComparison.OrdinalIgnoreCase) ||
                !Regex.IsMatch(intent, @"\b(?:copy|fill)\b", RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(intent, @"\b(?:below|downward)\b", RegexOptions.IgnoreCase) ||
                !Regex.IsMatch(intent, @"\buntil\s+before\b", RegexOptions.IgnoreCase))
                continue;
            var adjacentState = "adjacent-column:" + adjacent.Groups[1].Value.ToUpperInvariant();
            if (!string.IsNullOrWhiteSpace(fill.ExpectedState) &&
                fill.ExpectedState != "unverified-click" && fill.ExpectedState != adjacentState) continue;
            outcomes[index] = fill with { Action = "fill-down-to-adjacent-data-end",
                ExpectedState = adjacentState, OriginIntent = intent,
                Explanation = "Fill the preceding entered value downward until before the first empty adjacent data cell, per explicit intent." };
        }
        // A single recorded workflow can enter different values into the same destination
        // column under successive filters. Recorders sometimes attach the shared fill intent
        // only to the second drag/click. Reuse that explicit column and boundary intent for an
        // earlier unpaired entry only while that workbook is already filtered.
        for (var index = 0; index < outcomes.Count; index++)
        {
            if (outcomes[index] is not { Action: "type", Value: { Length: > 0 },
                    Target: { Process: "EXCEL", Window: { } workbook,
                        AutomationId: { Length: > 0 } entryAddress } } entry ||
                index + 1 < outcomes.Count && outcomes[index + 1].Action == "fill-down-to-adjacent-data-end")
                continue;
            var entryCell = Regex.Match(entryAddress, @"^([A-Z]{1,3})[1-9]\d*$", RegexOptions.IgnoreCase);
            if (!entryCell.Success) continue;
            var precedingFilter = outcomes.Take(index).LastOrDefault(candidate =>
                candidate.Action == "filter-values" && candidate.Target?.Window == workbook);
            if (precedingFilter?.ExpectedState is null or "filter:all") continue;
            for (var cursor = index + 1; cursor + 1 < outcomes.Count; cursor++)
            {
                if (outcomes[cursor].Action == "filter-values" && outcomes[cursor].Target?.Window == workbook &&
                    outcomes[cursor].ExpectedState == "filter:all") break;
                if (outcomes[cursor] is not { Action: "type", Target: { Process: "EXCEL", Window: { } templateWorkbook,
                        AutomationId: { Length: > 0 } templateAddress } } || templateWorkbook != workbook ||
                    outcomes[cursor + 1] is not { Action: "fill-down-to-adjacent-data-end", Target: { } fillTarget } template)
                    continue;
                var templateCell = Regex.Match(templateAddress, @"^([A-Z]{1,3})[1-9]\d*$", RegexOptions.IgnoreCase);
                if (!templateCell.Success || !templateCell.Groups[1].Value.Equals(entryCell.Groups[1].Value,
                        StringComparison.OrdinalIgnoreCase)) continue;
                outcomes[index] = entry with { TargetStrategy = "first-visible-filtered-row",
                    Explanation = "Enter the recorded value in this column's first visible filtered data row; the recorded row number is not reused." };
                outcomes.Insert(index + 1, template with { Number = 0, Target = fillTarget with { Window = workbook },
                    Intent = template.Intent, OriginIntent = template.OriginIntent,
                    Explanation = "Apply the same explicit visible-row fill rule to this preceding filtered value." });
                index++;
                break;
            }
        }
        var filteredWindows = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < outcomes.Count; index++)
        {
            var step = outcomes[index];
            if (step.Action == "filter-values" && step.Target?.Window is { } window)
            {
                if (step.ExpectedState == "filter:all") filteredWindows.Remove(window);
                else filteredWindows.Add(window);
            }
            if (step is { Action: "type", Target: { Process: "EXCEL", ClassName: "XLSpreadsheetCell",
                    AutomationId: { Length: > 0 }, Window: { } workbook } } &&
                filteredWindows.Contains(workbook) && index + 1 < outcomes.Count &&
                step.TargetStrategy is null or "first-visible-filtered-row" &&
                outcomes[index + 1] is { Action: "fill-down-to-adjacent-data-end", Target: { } fillTarget } &&
                fillTarget.Window == workbook && fillTarget.Process == step.Target.Process)
                outcomes[index] = step with { TargetStrategy = "first-visible-filtered-row",
                    Explanation = "Enter the recorded value in this column's first visible filtered data row; the recorded row number is not reused." };
            if (step.Action == "context-selection" && step.ExpectedState == "selection-intent:all" &&
                index + 1 < outcomes.Count && outcomes[index + 1] is
                    { Action: "context-selection", ExpectedState: "selection-intent:all", Target: { } followingTarget } &&
                followingTarget.Process == step.Target?.Process && followingTarget.Window == step.Target?.Window)
            {
                outcomes[index] = step with { Screenshot = outcomes[index + 1].Screenshot ?? step.Screenshot };
                outcomes.RemoveAt(index + 1);
            }
        }
        for (var index = outcomes.Count - 2; index >= 0; index--)
        {
            var navigation = outcomes[index];
            var next = outcomes[index + 1];
            if (navigation is { Action: "click", ExpectedState: null, OriginIntent: null,
                    Target: { Process: "EXCEL", ControlType: "ControlType.DataItem",
                        AutomationId: null or "", ClassName: null or "" } coarse } &&
                next is { Action: "type", TargetStrategy: "first-visible-filtered-row", Target: { } entryTarget } &&
                entryTarget.Process == coarse.Process && entryTarget.Window == coarse.Window &&
                navigation.Intent == next.Intent &&
                Regex.IsMatch(coarse.Name ?? "", @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase))
            {
                outcomes.RemoveAt(index);
                continue;
            }
            if (navigation.Action == "scroll" && navigation.Target?.Process == "EXCEL" &&
                navigation.Key is "Horizontal" or "Vertical" &&
                string.IsNullOrWhiteSpace(navigation.Intent) && navigation.OriginIntent is null &&
                (next.Action == "type" && next.TargetStrategy == "first-visible-filtered-row" ||
                 next.Action == "click" && next.Target?.ControlType == "ControlType.DataItem" &&
                 Regex.IsMatch(!string.IsNullOrWhiteSpace(next.Target.AutomationId)
                     ? next.Target.AutomationId : next.Target.Name ?? "",
                     @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase)) &&
                next.Target?.Process == navigation.Target.Process &&
                next.Target.Window == navigation.Target.Window)
                outcomes.RemoveAt(index);
        }
        RemoveDuplicateOrphanFillSequences(outcomes);
        // Excel may report a whole-column header click while the user is positioning for a
        // subsequent semantic fill. Once the next step contains a complete, intent-backed
        // fill operation, retaining that unverified header click only creates an unsafe and
        // redundant replay step. Keep ordinary column selections and selections that carry
        // their own intent or verified outcome.
        for (var index = outcomes.Count - 2; index >= 0; index--)
        {
            var selection = outcomes[index];
            var operation = outcomes[index + 1];
            if (selection is { Action: "manual-click", OriginIntent: null,
                    Target: { Process: "EXCEL", Window: { Length: > 0 } workbook,
                        ClassName: "XLGridColumnHeader", Name: { Length: > 0 } column } } &&
                Regex.IsMatch(column, @"^[A-Z]{1,3}$", RegexOptions.IgnoreCase) &&
                operation is { Action: "fill-down-to-adjacent-data-end" or "extend-formula-to-adjacent-data-end",
                    Target: { Process: "EXCEL", Window: { } operationWorkbook } } &&
                (!string.IsNullOrWhiteSpace(operation.Intent) ||
                 !string.IsNullOrWhiteSpace(operation.OriginIntent)) &&
                (string.IsNullOrWhiteSpace(selection.Intent) ||
                 selection.Intent == operation.Intent || selection.Intent == operation.OriginIntent) &&
                operationWorkbook == workbook)
                outcomes.RemoveAt(index);
        }
        for (var index = 0; index + 1 < outcomes.Count; index++)
        {
            var header = outcomes[index];
            var optionalDialog = outcomes[index + 1];
            if (header is not { Action: "manual-click", ExpectedState: "unverified-click",
                    Target: { Process: "EXCEL", ClassName: "XLGridColumnHeader",
                        Name: { Length: > 0 } column } target } ||
                !Regex.IsMatch(column, @"^[A-Z]{1,3}$", RegexOptions.IgnoreCase) ||
                optionalDialog is not { Action: "click" or "optional-click", Intent: { Length: > 0 },
                    Target: { Process: "EXCEL", ControlType: "ControlType.Button",
                        ParentName: { Length: > 0 } dialogParent } } ||
                target.Window != optionalDialog.Target.Window &&
                !string.IsNullOrWhiteSpace(optionalDialog.Target.Window)) continue;
            // The header click can be the action that makes an optional Excel collaboration
            // prompt appear. Preserve that trigger as a normal live UIA click; only its delayed
            // selection outcome was unverified. The following named dialog/button remains the
            // evidence that this was a simple trigger rather than an unknown drag or fill.
            outcomes[index] = header with { Action = "click", ExpectedState = null,
                Explanation = $"Click the recorded column header, which may trigger the optional '{dialogParent}' dialog." };
        }
        for (var index = 1; index + 1 < outcomes.Count; index++)
        {
            if (outcomes[index - 1] is not { Action: "filter-values", Target: { Process: "EXCEL",
                    Window: { } workbook } } ||
                outcomes[index] is not { Action: "click", Target: { Process: "EXCEL",
                    ClassName: "XLGridColumnHeader", Window: { } headerWorkbook } } ||
                outcomes[index + 1] is not { Action: "click", Intent: { Length: > 0 } optionalIntent,
                    Target: { Process: "EXCEL", ControlType: "ControlType.Button",
                        ParentName: { Length: > 0 } } optional } ||
                !Regex.IsMatch(optionalIntent, @"\b(?:doesn['’]?t|does\s+not)\s+always\s+appear\b",
                    RegexOptions.IgnoreCase) ||
                !string.Equals(workbook, headerWorkbook, StringComparison.OrdinalIgnoreCase) ||
                !string.IsNullOrWhiteSpace(optional.Window))
                continue;
            // Filtering can raise Excel's collaboration dialog asynchronously. Handle the
            // recorded optional response before touching the worksheet again; otherwise the
            // modal window blocks the following header click and replay asks for manual help.
            (outcomes[index], outcomes[index + 1]) = (outcomes[index + 1], outcomes[index]);
            index++;
        }
        for (var filterIndex = outcomes.Count - 1; filterIndex >= 0; filterIndex--)
        {
            if (outcomes[filterIndex] is not { Action: "filter-values",
                    Target: { Process: "EXCEL", Window: { } workbook } }) continue;
            var start = filterIndex;
            while (start > 0 && outcomes[start - 1] is
                { Action: "click", OriginIntent: null, ExpectedState: null, Value: null, Key: null,
                  Screenshot: null, Target: { Process: "EXCEL", ControlType: "ControlType.DataItem",
                      AutomationId: null or "", ClassName: null or "", Window: { } navWorkbook } cell } navigation &&
                navWorkbook == workbook && Regex.IsMatch(cell.Name ?? "", @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase) &&
                (string.IsNullOrWhiteSpace(navigation.Intent) || outcomes.Take(start - 1).Any(prior =>
                    prior.Action == "fill-down-to-adjacent-data-end" && prior.Target?.Window == workbook &&
                    prior.OriginIntent == navigation.Intent)))
                start--;
            if (start < filterIndex)
            {
                outcomes.RemoveRange(start, filterIndex - start);
                filterIndex = start;
            }
        }
        for (var index = outcomes.Count - 2; index >= 0; index--)
        {
            var navigation = outcomes[index];
            var next = outcomes[index + 1];
            if (navigation is not { Action: "click", Intent: null, OriginIntent: null,
                    ExpectedState: null, Value: null, Key: null, Screenshot: null,
                    Target: { Process: "EXCEL", ControlType: "ControlType.DataItem",
                        AutomationId: null or "", ClassName: null or "", Window: { } workbook } cell } ||
                !Regex.IsMatch(cell.Name ?? "", @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase))
                continue;
            var leadsToFilter = next is { Action: "filter-values", Target: { Process: "EXCEL", Window: { } filterWorkbook } } &&
                filterWorkbook == workbook;
            var continuesNavigation = next is { Action: "click", Intent: null, OriginIntent: null,
                ExpectedState: null, Value: null, Key: null, Screenshot: null,
                Target: { Process: "EXCEL", ControlType: "ControlType.DataItem", AutomationId: null or "",
                    ClassName: null or "", Window: { } nextWorkbook } nextCell } && nextWorkbook == workbook &&
                Regex.IsMatch(nextCell.Name ?? "", @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase);
            if (leadsToFilter || continuesNavigation) outcomes.RemoveAt(index);
        }
        return plan with { Steps = outcomes
            .Select((step, index) => step with { Number = index + 1 }).ToList() };
    }

    internal static ExecutionPlan RecoverExplicitDuplicatedRowIntent(ExecutionPlan plan,
        IEnumerable<string> recordedIntents)
    {
        foreach (var intent in recordedIntents.Where(note => !string.IsNullOrWhiteSpace(note)).Distinct())
        {
            var columns = Regex.Match(intent,
                @"\bcolumns?\s+([A-Z]{1,3})\s+to\s+([A-Z]{1,3})\b", RegexOptions.IgnoreCase);
            var weekday = Regex.Match(intent,
                @"\bnext\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b",
                RegexOptions.IgnoreCase);
            if (!columns.Success || !weekday.Success || !Regex.IsMatch(intent,
                    @"\bcopy\s+(?:the\s+)?(?:functions?|formulas?)\s+of\s+(?:the\s+)?(?:previous\s+row|last\s+row\s+containing\s+values?)\b",
                    RegexOptions.IgnoreCase)) continue;
            var firstColumn = columns.Groups[1].Value.ToUpperInvariant();
            var lastColumn = columns.Groups[2].Value.ToUpperInvariant();
            if (ColumnNumber(lastColumn) < ColumnNumber(firstColumn)) continue;

            var semantic = plan.Steps.FindIndex(step => step.Action == "duplicate-range-values-and-formulas" &&
                Regex.IsMatch(step.Value ?? "", @"^[A-Z]{1,3}[1-9]\d*:[A-Z]{1,3}[1-9]\d*$",
                    RegexOptions.IgnoreCase));
            if (semantic >= 0)
            {
                var recorded = Regex.Match(plan.Steps[semantic].Value!,
                    @"^[A-Z]{1,3}(?<row>[1-9]\d*):[A-Z]{1,3}\k<row>$", RegexOptions.IgnoreCase);
                if (recorded.Success)
                {
                    var row = int.Parse(recorded.Groups["row"].Value,
                        System.Globalization.CultureInfo.InvariantCulture);
                    var sourceRange = $"{firstColumn}{row}:{lastColumn}{row}";
                    var destinationRange = $"{firstColumn}{row + 1}:{lastColumn}{row + 1}";
                    var revised = plan.Steps.ToList();
                    revised[semantic] = revised[semantic] with
                    {
                        Value = sourceRange, ExpectedState = "destination-range:" + destinationRange,
                        Intent = intent, OriginIntent = intent,
                        TargetStrategy = "last-populated-row-in-first-column",
                        Explanation = $"Copy columns {firstColumn} through {lastColumn} from the live last populated row into the next empty row and verify the result."
                    };
                    // The explicit column range and live-row wording override incidental
                    // selection endpoints such as G11 captured while the user positioned
                    // the mouse after selecting A:F.
                    return Compact(plan with { Steps = revised });
                }
            }

            // Selection capture may include an incidental cell between the copied range and its
            // destination (for example G12 after selecting A11:F11). Recover only when both range
            // endpoints, the first cell of the immediately following row, and a nearby Ctrl+F are
            // all present in the same workbook. These constraints keep the rule recording-generic
            // without turning arbitrary cell clicks into a row copy.
            for (var start = 0; start < plan.Steps.Count; start++)
            {
                if (!TryCell(plan.Steps[start], out var sourceFirst, out var sourceRow) ||
                    sourceFirst != firstColumn) continue;
                var end = FindCell(plan.Steps, start + 1, Math.Min(plan.Steps.Count, start + 5),
                    lastColumn, sourceRow, plan.Steps[start]);
                if (end < 0) continue;
                var destination = FindCell(plan.Steps, end + 1, Math.Min(plan.Steps.Count, end + 5),
                    firstColumn, sourceRow + 1, plan.Steps[start]);
                if (destination < 0) continue;
                var find = Enumerable.Range(destination + 1,
                        Math.Min(4, plan.Steps.Count - destination - 1))
                    .FirstOrDefault(index => plan.Steps[index] is { Action: "key", Key: "Control+F" or "Ctrl+F" } &&
                        SameExcelSheet(plan.Steps[start], plan.Steps[index]), -1);
                if (find < 0) continue;
                var sourceRange = $"{firstColumn}{sourceRow}:{lastColumn}{sourceRow}";
                var destinationRange = $"{firstColumn}{sourceRow + 1}:{lastColumn}{sourceRow + 1}";
                var replacement = plan.Steps[start] with
                {
                    Action = "duplicate-range-values-and-formulas", Value = sourceRange,
                    ExpectedState = "destination-range:" + destinationRange,
                    Intent = intent, OriginIntent = intent,
                    Explanation = $"Copy the explicitly described previous-row range {sourceRange} into {destinationRange} and verify the result."
                };
                var revised = plan.Steps.ToList();
                revised.RemoveRange(start, destination - start + 1);
                revised.Insert(start, replacement);
                return Compact(plan with { Steps = revised });
            }
        }
        return plan;

        static int FindCell(IReadOnlyList<PlanStep> steps, int start, int end,
            string column, int row, PlanStep anchor)
        {
            for (var index = start; index < end; index++)
                if (TryCell(steps[index], out var foundColumn, out var foundRow) &&
                    foundColumn == column && foundRow == row && SameExcelSheet(anchor, steps[index]))
                    return index;
            return -1;
        }

        static bool TryCell(PlanStep step, out string column, out int row)
        {
            column = ""; row = 0;
            if (step is not { Action: "click", Target: { Process: "EXCEL",
                    ControlType: "ControlType.DataItem" } target }) return false;
            var match = Regex.Match(target.AutomationId ?? target.Name ?? "",
                @"^([A-Z]{1,3})([1-9]\d*)$", RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            column = match.Groups[1].Value.ToUpperInvariant();
            return int.TryParse(match.Groups[2].Value, out row);
        }

        static bool SameExcelSheet(PlanStep left, PlanStep right) =>
            left.Target?.Process == right.Target?.Process && left.Target?.Window == right.Target?.Window;

        static int ColumnNumber(string column)
        {
            var number = 0;
            foreach (var character in column) number = number * 26 + character - 'A' + 1;
            return number;
        }
    }

    private static List<PlanStep> NormalizeTransientMenuOwners(IReadOnlyList<PlanStep> steps)
    {
        // Popup menus can disappear before delayed capture finishes. When that happens UIA may
        // report a sibling window that appeared after the click (for example Sample Case Query)
        // as the menu's owner. A menu opened by a recorded button, and each child whose ParentName
        // names the preceding menu, belongs to the opener's window. Restricting this correction to
        // those two structural relationships avoids rewriting unrelated menus in newly opened windows.
        var result = steps.ToList();
        for (var index = 1; index < result.Count; index++)
        {
            var previous = result[index - 1];
            var current = result[index];
            if (current is not { Action: "click", Target: { ControlType: "ControlType.MenuItem" } menu } ||
                previous is not { Action: "click", Target: { Process: { Length: > 0 },
                    Window: { Length: > 0 } owner } opener } ||
                !string.Equals(opener.Process, menu.Process, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(owner, menu.Window, StringComparison.OrdinalIgnoreCase))
                continue;

            var openedByButton = opener.ControlType == "ControlType.Button" &&
                !string.IsNullOrWhiteSpace(opener.ParentName);
            var openedByParentMenu = opener.ControlType == "ControlType.MenuItem" &&
                string.Equals(menu.ParentName, opener.Name, StringComparison.OrdinalIgnoreCase);
            if (!openedByButton && !openedByParentMenu) continue;

            result[index] = current with
            {
                Target = menu with { Window = owner },
                Explanation = current.Explanation
            };
        }
        return result;
    }

    private static List<PlanStep> NormalizeFormulaFillGestureBeforeDataCommand(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>();
        for (var index = 0; index < steps.Count; index++)
        {
            var paste = steps[index];
            if (index + 3 >= steps.Count || paste is not
                { Action: "click", Intent: { } intent,
                  Target: { Process: "EXCEL", Name: "Values", ParentName: "Paste Values",
                      Window: { Length: > 0 } workbook } } ||
                !Regex.IsMatch(intent.Trim(), @"^(?:intent\s+is\s+to\s+)?(?:copy\s+)?(?:the\s+)?(?:function|formula)s?\.?$",
                    RegexOptions.IgnoreCase) ||
                !TryCell(steps[index + 1], out var column, out var sourceRow) ||
                !TryCell(steps[index + 2], out var repeatedColumn, out var repeatedRow) ||
                !TryCell(steps[index + 3], out var nextColumn, out var nextRow) ||
                column != repeatedColumn || column != nextColumn || sourceRow != repeatedRow ||
                nextRow != sourceRow + 1 || steps[index + 1].Target?.Window != workbook ||
                steps[index + 2].Target?.Window != workbook || steps[index + 3].Target?.Window != workbook)
            {
                result.Add(paste);
                continue;
            }
            var adjacent = PreviousColumn(column);
            var boundary = steps.Skip(index + 4).Take(8).ToList();
            if (!boundary.Any(step => step is { Action: "click", Target: { Process: "EXCEL",
                    Name: "Remove Duplicates", Window: { } window } } && window == workbook) ||
                !boundary.Any(step => step is { Action: "click", Target: { Process: "EXCEL",
                    ControlType: "ControlType.DataItem", Name: { } name, Window: { } window } } &&
                    window == workbook && name.Equals(adjacent, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(paste);
                continue;
            }
            result.Add(paste);
            // Repeated source then next-row selection is the recorder's verified fill
            // gesture. Materialize it before the following Data command so destructive
            // duplicate removal cannot overtake the formula dependency.
            result.Add(steps[index + 1] with
            {
                Action = "extend-formula-to-adjacent-data-end",
                Target = steps[index + 1].Target! with { Name = column + "1", AutomationId = column + "1",
                    ClassName = "XLSpreadsheetCell" },
                Value = column, Key = null, ExpectedState = "adjacent-column:" + adjacent,
                TargetStrategy = "last-populated-formula", Intent = intent, OriginIntent = intent,
                Explanation = $"Extend the last populated formula in column {column} through the contiguous populated rows in adjacent column {adjacent}; verify the fill before continuing to Remove Duplicates."
            });
            index += 3;
        }
        return result;

        static bool TryCell(PlanStep step, out string column, out int row)
        {
            column = ""; row = 0;
            if (step is not { Action: "click", Target: { Process: "EXCEL", ControlType: "ControlType.DataItem" } target })
                return false;
            var match = Regex.Match(target.AutomationId ?? target.Name ?? "",
                @"^(?<column>[A-Z]{1,3})(?<row>[1-9]\d*)$", RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            column = match.Groups["column"].Value.ToUpperInvariant();
            return int.TryParse(match.Groups["row"].Value,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture, out row);
        }

        static string PreviousColumn(string column)
        {
            var number = column.Aggregate(0, (value, character) => value * 26 + character - 'A' + 1);
            if (number <= 1) throw new InvalidDataException("A formula fill has no adjacent preceding column.");
            number--;
            var name = "";
            while (number > 0)
            {
                number--;
                name = (char)('A' + number % 26) + name;
                number /= 26;
            }
            return name;
        }
    }

    private static void RemoveDuplicateOrphanFillSequences(List<PlanStep> steps)
    {
        for (var index = steps.Count - 1; index >= 0; index--)
        {
            var duplicate = steps[index];
            if (duplicate is not { Action: "fill-down-to-adjacent-data-end", OriginIntent: { Length: > 0 } intent,
                    Target: { Process: "EXCEL", Window: { Length: > 0 } workbook } } ||
                index > 0 && steps[index - 1].Action == "type")
                continue;
            var original = -1;
            for (var candidate = index - 1; candidate > 0; candidate--)
            {
                if (steps[candidate] is { Action: "fill-down-to-adjacent-data-end", Target: { } target } prior &&
                    prior.OriginIntent == intent && target.Process == duplicate.Target.Process &&
                    target.Window == workbook && steps[candidate - 1].Action == "type")
                { original = candidate; break; }
            }
            if (original < 0) continue;
            var start = index;
            while (start > original + 1 && steps[start - 1] is
                { Action: "click", Intent: null, OriginIntent: null, ExpectedState: null,
                  Value: null, Key: null, Screenshot: null,
                  Target: { Process: "EXCEL", ControlType: "ControlType.DataItem", Window: { } navWorkbook } nav } &&
                navWorkbook == workbook && Regex.IsMatch(nav.Name ?? nav.AutomationId ?? "",
                    @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase))
                start--;
            steps.RemoveRange(start, index - start + 1);
            index = Math.Min(start, steps.Count);
        }
    }

    private static void RecoverExcelColumnFilterOpeners(List<PlanStep> steps)
    {
        for (var index = 0; index + 1 < steps.Count; index++)
        {
            var opener = steps[index];
            if (opener is not { Action: "click" or "manual-click", Intent: null, OriginIntent: null,
                    Target: { Process: "EXCEL", Window: { Length: > 0 } workbook,
                        ControlType: "ControlType.DataItem", Name: { Length: > 0 } column } target } ||
                !string.IsNullOrWhiteSpace(target.AutomationId) ||
                !Regex.IsMatch(column, "^[A-Z]{1,3}$", RegexOptions.IgnoreCase) ||
                !ExcelFilterPlan.IsItem(steps[index + 1].Target) ||
                steps[index + 1].Target?.Window != workbook ||
                steps[index + 1].WhenUser != opener.WhenUser)
                continue;

            var parent = column.ToUpperInvariant() + "1";
            var prior = steps.Take(index).LastOrDefault(candidate =>
                candidate.Target is { Process: "EXCEL", Window: { } priorWorkbook,
                    AutomationId: "Dropdown", ParentName: { } priorParent } &&
                priorWorkbook == workbook && priorParent.Equals(parent, StringComparison.OrdinalIgnoreCase));
            steps[index] = opener with
            {
                Action = "click",
                Target = target with
                {
                    AutomationId = "Dropdown",
                    Name = prior?.Target?.Name ?? "Filter applied",
                    ControlType = "ControlType.MenuItem",
                    ParentName = parent
                },
                ExpectedState = null,
                Explanation = $"Open the recorded filter for column {column.ToUpperInvariant()}; Excel exposed only the column label while the flyout opened."
            };
        }
    }

    private static List<PlanStep> NormalizeDuplicateFormulaExtensions(IReadOnlyList<PlanStep> steps)
    {
        var result = steps.ToList();
        var duplicateGroups = result.Select((step, index) => (step, index))
            .Where(item => item.step.Action == "extend-formula-to-adjacent-data-end" &&
                           !string.IsNullOrWhiteSpace(item.step.OriginIntent))
            .GroupBy(item => string.Join("\u001f", item.step.Target?.Process, item.step.Target?.Window,
                item.step.Value, item.step.ExpectedState, item.step.TargetStrategy, item.step.OriginIntent),
                StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .ToList();
        var remove = new HashSet<int>();
        foreach (var group in duplicateGroups)
        {
            static int AnchorScore(IReadOnlyList<PlanStep> all, int index)
            {
                if (index == 0) return 0;
                var preceding = all[index - 1].Target;
                if (preceding is { Name: "Values", ParentName: "Paste Values" }) return 3;
                if (preceding?.AutomationId == "PasteMenu_Dropdown" || preceding?.Name == "More Options") return 2;
                return 0;
            }
            var keep = group.OrderByDescending(item => AnchorScore(result, item.index))
                .ThenBy(item => item.index).First().index;
            foreach (var item in group)
                if (item.index != keep) remove.Add(item.index);
        }
        foreach (var index in remove.OrderByDescending(value => value)) result.RemoveAt(index);
        return result;
    }

    private static List<PlanStep> NormalizeMissingChartLegendDrag(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>();
        for (var index = 0; index < steps.Count; index++)
        {
            var first = steps[index];
            var second = index + 1 < steps.Count ? steps[index + 1] : null;
            if (index + 2 < steps.Count && first is
                    { Action: "click", Target: { Process: "EXCEL", Name: { Length: > 0 } firstName } firstTarget } &&
                second is { Action: "click", Target: { Process: "EXCEL", Name: { Length: > 0 } secondName } secondTarget } &&
                steps[index + 2] is { Action: "click", Target: { Process: "EXCEL", ControlType: "ControlType.TabItem" } } &&
                firstName == secondName && firstTarget.Window == secondTarget.Window &&
                (firstTarget.ClassName == "ExcelChartObject" || secondTarget.ClassName == "ExcelChartObject" ||
                 firstTarget.ControlType == "ControlType.Image" || secondTarget.ControlType == "ControlType.Image"))
            {
                result.Add(first with
                {
                    Action = "expand-chart-legend-opposite-plot",
                    Target = firstTarget with
                    {
                        Name = firstName, ClassName = "ExcelChartObject", ControlType = "ControlType.Image"
                    },
                    ExpectedState = "chart-legend-expanded-right",
                    Explanation = "Expand the selected chart legend across the available side opposite the plotted data."
                });
                index++;
                continue;
            }
            result.Add(steps[index]);
        }
        return result;
    }

    private static List<PlanStep> NormalizeDuplicatedRowDate(IReadOnlyList<PlanStep> steps)
    {
        // Mark the relative "previous row" meaning even when no date rewrite
        // follows. This keeps the append behavior attached to the copy itself.
        steps = steps.Select(step => step.Action == "duplicate-range-values-and-formulas" &&
                (step.Intent ?? step.OriginIntent)?.Contains("previous row",
                    StringComparison.OrdinalIgnoreCase) == true
            ? step with { TargetStrategy = "last-populated-row-in-first-column",
                Explanation = "Copy columns from the live last populated row into the next empty row and verify the result." }
            : step).ToList();
        var weekday = steps.FirstOrDefault(step => step.Action == "rename-sheet" &&
            step.RelativeWeekday is not null)?.RelativeWeekday;
        if (weekday is null) return steps.ToList();
        var result = new List<PlanStep>();
        for (var index = 0; index < steps.Count; index++)
        {
            var duplicate = steps[index];
            if (duplicate.Action != "duplicate-range-values-and-formulas" ||
                duplicate.ExpectedState?.StartsWith("destination-range:", StringComparison.Ordinal) != true)
            {
                result.Add(duplicate);
                continue;
            }
            var rowIntent = duplicate.Intent ?? duplicate.OriginIntent ?? steps
                .Skip(index + 1).Take(16)
                .Select(step => step.Intent ?? step.OriginIntent)
                .FirstOrDefault(intent => intent is not null && Regex.IsMatch(intent,
                    @"\bcopy\s+(?:the\s+)?(?:functions?|formulas?)\s+of\s+(?:the\s+)?(?:previous\s+row|last\s+row\s+containing\s+values?)\b",
                    RegexOptions.IgnoreCase));
            var followsLivePreviousRow = duplicate.TargetStrategy == "last-populated-row-in-first-column" ||
                rowIntent is not null;
            if (rowIntent is not null)
            {
                var columns = Regex.Match(rowIntent,
                    @"\bcolumns?\s+(?<first>[A-Z]{1,3})\s+to\s+(?<last>[A-Z]{1,3})\b",
                    RegexOptions.IgnoreCase);
                var recorded = Regex.Match(duplicate.Value ?? "",
                    @"^(?<first>[A-Z]{1,3})(?<row>[1-9]\d*):(?<last>[A-Z]{1,3})\k<row>$",
                    RegexOptions.IgnoreCase);
                if (columns.Success && recorded.Success)
                {
                    var sourceRow = recorded.Groups["row"].Value;
                    var destinationRow = (int.Parse(sourceRow,
                        System.Globalization.CultureInfo.InvariantCulture) + 1).ToString(
                            System.Globalization.CultureInfo.InvariantCulture);
                    var first = columns.Groups["first"].Value.ToUpperInvariant();
                    var last = columns.Groups["last"].Value.ToUpperInvariant();
                    duplicate = duplicate with
                    {
                        Value = $"{first}{sourceRow}:{last}{sourceRow}",
                        ExpectedState = $"destination-range:{first}{destinationRow}:{last}{destinationRow}",
                        Intent = rowIntent, OriginIntent = rowIntent,
                        TargetStrategy = "last-populated-row-in-first-column",
                        Explanation = $"Copy columns {first} through {last} from the live last populated row into the next empty row and verify the result."
                    };
                }
            }
            result.Add(duplicate);
            var destination = duplicate.ExpectedState["destination-range:".Length..];
            var firstCell = destination.Split(':')[0];
            var sourceFirstCell = duplicate.Value?.Split(':')[0]
                ?? throw new InvalidDataException("Duplicated row source range is missing.");
            var sessionStart = index + 1;
            if (sessionStart < steps.Count && steps[sessionStart].Action == "update-cell-to-relative-weekday")
            {
                result.Add(steps[sessionStart] with
                {
                    Target = duplicate.Target! with { Name = firstCell, AutomationId = firstCell, ClassName = "XLSpreadsheetCell" },
                    Value = sourceFirstCell, ExpectedState = "destination-range:" + destination,
                    RelativeWeekday = weekday, TargetStrategy = followsLivePreviousRow
                        ? "copied-live-last-row-date-and-formulas" : "copied-row-date-and-formulas",
                    Explanation = $"Update the copied row's date and formula references to the next {weekday}."
                });
                index++;
                continue;
            }
            if (sessionStart >= steps.Count || steps[sessionStart] is not { Action: "key", Key: "Control+F" or "Ctrl+F" })
                continue;
            var sessionEnd = sessionStart;
            while (sessionEnd < steps.Count &&
                !(steps[sessionEnd].Action is "click" or "close-window" &&
                  steps[sessionEnd].Target is { Window: "Find and Replace", Name: "Close" }))
                sessionEnd++;
            if (sessionEnd >= steps.Count || !steps.Skip(sessionStart).Take(sessionEnd - sessionStart + 1)
                    .Any(step => step.Action == "replace-all"))
                continue;
            result.Add(new PlanStep(0, "update-cell-to-relative-weekday",
                duplicate.Target! with { Name = firstCell, AutomationId = firstCell, ClassName = "XLSpreadsheetCell" },
                sourceFirstCell, null, "destination-range:" + destination, duplicate.Screenshot,
                $"Update the copied row's date and formula references to the next {weekday}.",
                RelativeWeekday: weekday, TargetStrategy: followsLivePreviousRow
                    ? "copied-live-last-row-date-and-formulas" : "copied-row-date-and-formulas"));
            index = sessionEnd;
        }
        return result;
    }

    private static List<PlanStep> EnsureStandingIntentReplacements(List<PlanStep> steps)
    {
        var firstReplacement = steps.FindIndex(step => step.Action == "replace-all");
        if (firstReplacement < 0) return steps;
        var note = steps.Take(firstReplacement + 1).Select(step => step.Intent)
            .LastOrDefault(intent => intent?.Contains("always replace", StringComparison.OrdinalIgnoreCase) == true);
        if (note is null) return steps;
        var match = Regex.Match(note,
            "always\\s+replace\\s+\"(?<find>[^\"]+)\"\\s+with\\s+\"(?<replacement>.*?)(?=\\s+and\\s+if|\")",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (!match.Success) return steps;
        var find = match.Groups["find"].Value.Trim();
        var replacement = match.Groups["replacement"].Value.Trim();
        if (find.Length == 0 || replacement.Length == 0 || steps.Any(step =>
                step.Action == "replace-all" && step.Value == find && step.ExpectedState == replacement))
            return steps;
        var anchor = steps[firstReplacement];
        steps.Insert(firstReplacement, anchor with
        {
            Number = 0, Value = find, ExpectedState = replacement,
            Intent = note, OriginIntent = note, WhenUser = null,
            Explanation = $"Apply the standing recorded replacement of {find} with {replacement}."
        });
        return steps;
    }

    private static List<PlanStep> NormalizeDuplicatedRowGesture(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>();
        for (var index = 0; index < steps.Count; index++)
        {
            if (index + 3 < steps.Count &&
                TryCell(steps[index], out var firstColumn, out var sourceRow) &&
                TryCell(steps[index + 1], out var lastColumn, out var lastSourceRow) &&
                TryCell(steps[index + 2], out var destinationColumn, out var destinationRow) &&
                steps[index + 3] is { Action: "key", Key: "Control+F" or "Ctrl+F" } find &&
                sourceRow == lastSourceRow && sourceRow + 1 == destinationRow &&
                firstColumn == destinationColumn && ColumnNumber(lastColumn) > ColumnNumber(firstColumn) &&
                SameExcelSheet(steps[index], steps[index + 1]) &&
                SameExcelSheet(steps[index], steps[index + 2]) &&
                SameExcelSheet(steps[index], find))
            {
                var source = $"{firstColumn}{sourceRow}:{lastColumn}{sourceRow}";
                var destination = $"{firstColumn}{destinationRow}:{lastColumn}{destinationRow}";
                result.Add(steps[index] with
                {
                    Action = "duplicate-range-values-and-formulas",
                    Value = source,
                    ExpectedState = "destination-range:" + destination,
                    Explanation = $"Copy the recorded row range {source} into {destination} and verify the result."
                });
                index += 2;
                continue;
            }
            if (index + 3 < steps.Count &&
                TryCell(steps[index], out firstColumn, out sourceRow) &&
                IsUnidentifiedExcelRangeEndpoint(steps[index + 1], steps[index]) &&
                TryCell(steps[index + 2], out destinationColumn, out destinationRow) &&
                steps[index + 3] is { Action: "key", Key: "Control+F" or "Ctrl+F" } inferredFind &&
                sourceRow + 1 == destinationRow && firstColumn == destinationColumn &&
                SameExcelSheet(steps[index], steps[index + 2]) && SameExcelSheet(steps[index], inferredFind) &&
                FindCopiedColumnRange(steps, index + 3, firstColumn) is { } inferredLastColumn &&
                ColumnNumber(inferredLastColumn) >= ColumnNumber(firstColumn))
            {
                var source = $"{firstColumn}{sourceRow}:{inferredLastColumn}{sourceRow}";
                var destination = $"{firstColumn}{destinationRow}:{inferredLastColumn}{destinationRow}";
                result.Add(steps[index] with
                {
                    Action = "duplicate-range-values-and-formulas",
                    Value = source,
                    ExpectedState = "destination-range:" + destination,
                    Explanation = $"Copy the explicitly described previous-row range {source} into {destination} and verify the result."
                });
                index += 2;
                continue;
            }
            result.Add(steps[index]);
        }
        return result;

        static bool TryCell(PlanStep step, out string column, out int row)
        {
            column = ""; row = 0;
            if (step is not { Action: "click", Target: { Process: "EXCEL", ControlType: "ControlType.DataItem" } target })
                return false;
            var match = Regex.Match(target.AutomationId ?? target.Name ?? "", @"^([A-Z]{1,3})([1-9]\d*)$",
                RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            column = match.Groups[1].Value.ToUpperInvariant();
            return int.TryParse(match.Groups[2].Value, out row);
        }

        static bool SameExcelSheet(PlanStep left, PlanStep right) =>
            left.Target?.Process == right.Target?.Process && left.Target?.Window == right.Target?.Window;

        static bool IsUnidentifiedExcelRangeEndpoint(PlanStep candidate, PlanStep source) =>
            candidate.Action is "click" or "manual-click" && candidate.Target is
                { Process: "EXCEL", ControlType: "ControlType.Custom", Name: null or "" } target &&
            target.Window == source.Target?.Window && string.IsNullOrWhiteSpace(target.AutomationId) == false;

        static string? FindCopiedColumnRange(IReadOnlyList<PlanStep> all, int start, string firstColumn)
        {
            for (var cursor = start; cursor < Math.Min(all.Count, start + 24); cursor++)
            {
                var step = all[cursor];
                foreach (var note in new[] { step.Intent, step.OriginIntent })
                {
                    if (string.IsNullOrWhiteSpace(note)) continue;
                    var columns = Regex.Match(note,
                        @"\bcolumns?\s+([A-Z]{1,3})\s+to\s+([A-Z]{1,3})\b",
                        RegexOptions.IgnoreCase);
                    if (columns.Success && columns.Groups[1].Value.Equals(firstColumn,
                            StringComparison.OrdinalIgnoreCase) &&
                        Regex.IsMatch(note,
                            @"\bcopy\s+(?:the\s+)?(?:functions?|formulas?)\s+of\s+(?:the\s+)?previous\s+row\b",
                            RegexOptions.IgnoreCase))
                        return columns.Groups[2].Value.ToUpperInvariant();
                }
                if (cursor > start && step.Target is { Process: "EXCEL", ControlType: "ControlType.TabItem" }) break;
            }
            return null;
        }

        static int ColumnNumber(string column)
        {
            var number = 0;
            foreach (var character in column) number = number * 26 + character - 'A' + 1;
            return number;
        }
    }

    private static List<PlanStep> ScopeRecordedFindReplaceSessions(List<PlanStep> steps)
    {
        // A later Find/Replace session may inherit an earlier standing intent. When
        // this session contains its own recorded Replace All with the values typed
        // into both fields, that concrete pair defines this session exclusively.
        for (var start = 0; start < steps.Count;)
        {
            while (start < steps.Count && steps[start].Target?.Window != "Find and Replace") start++;
            if (start >= steps.Count) break;
            var end = start;
            while (end + 1 < steps.Count && (steps[end + 1].Target?.Window == "Find and Replace" ||
                steps[end + 1].Action == "replace-all")) end++;
            var recorded = steps.Skip(start).Take(end - start + 1).FirstOrDefault(step =>
                step.Action == "replace-all" && step.OriginIntent is null);
            if (recorded is not null)
            {
                for (var index = end; index >= start; index--)
                    if (steps[index].Action == "replace-all" && steps[index].OriginIntent is not null &&
                        (steps[index].Value != recorded.Value || steps[index].ExpectedState != recorded.ExpectedState))
                        steps.RemoveAt(index);
                end = start;
                while (end + 1 < steps.Count && (steps[end + 1].Target?.Window == "Find and Replace" ||
                    steps[end + 1].Action == "replace-all")) end++;
            }
            start = end + 1;
        }
        return steps;
    }

    private static List<PlanStep> NormalizeSpreadsheetEdits(IReadOnlyList<PlanStep> steps)
    {
        var result = new List<PlanStep>();
        for (var index = 0; index < steps.Count; index++)
        {
            var entry = steps[index];
            if (entry is { Action: "type", Value: not null,
                    Target: { Process: "EXCEL", ClassName: "XLSpreadsheetCell",
                        AutomationId: { Length: > 0 } } cell } &&
                Regex.IsMatch(cell.AutomationId, @"^[A-Z]{1,3}[1-9]\d*$"))
            {
                while (index + 1 < steps.Count && steps[index + 1] is { Target: { ClassName: "EXCEL6" } pane } next &&
                    next.OriginIntent is null && next.WhenUser == entry.WhenUser &&
                    string.IsNullOrWhiteSpace(next.Intent) && string.IsNullOrWhiteSpace(next.ExpectedState) &&
                    ActionGrouper.TryContinueGridEdit(cell, pane, entry.Value!, next.Action, next.Value, next.Key, out var value))
                {
                    entry = entry with { Value = value, ExpectedState = null,
                        Screenshot = next.Screenshot ?? entry.Screenshot,
                        Explanation = "Enter the final recorded cell text after its corrections, not intermediate edit-pane keystrokes." };
                    index++;
                }
            }
            result.Add(entry);
        }
        return result;
    }

    private static string? RelativeWeekdayFromIntent(string? intent)
    {
        if (string.IsNullOrWhiteSpace(intent)) return null;
        var match = Regex.Match(intent,
            @"\bnext\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b",
            RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : null;
    }

    internal static bool IsReplacementAcknowledgement(PlanStep step, PlanStep previous) =>
        step.Action == "click" && step.Target is
            { Process: "EXCEL", Window: "Microsoft Excel", Name: "OK", ControlType: "ControlType.Button" } &&
        previous.Action == "replace-all" && previous.Target is { Process: "EXCEL", Window: "Find and Replace" };

    internal static bool IsSupersededExcelCellClick(PlanStep click, PlanStep next) =>
        click is { Action: "click", Intent: null, ExpectedState: null, Target:
            { Process: "EXCEL", ControlType: "ControlType.DataItem" } coarse } &&
        string.IsNullOrWhiteSpace(coarse.AutomationId) && string.IsNullOrWhiteSpace(coarse.ClassName) &&
        Regex.IsMatch(coarse.Name ?? "", @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase) &&
        next is { Action: "type", Value: not null, Target:
            { Process: "EXCEL", ControlType: "ControlType.DataItem", ClassName: "XLSpreadsheetCell",
                AutomationId: { Length: > 0 } address } precise } &&
        Regex.IsMatch(address, @"^[A-Z]{1,3}[1-9]\d*$", RegexOptions.IgnoreCase) &&
        coarse.Window == precise.Window;

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
