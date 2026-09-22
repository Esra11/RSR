using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Automation;

namespace DesktopSteps;

internal static class Executor
{
    public static async Task<bool> ReplayAsync(ExecutionPlan plan, Action<string> status,
        Func<string, bool> handleUnexpectedDialog, Func<string, bool> approveBrowserChange,
        CancellationToken token)
    {
        // An older recorder could append an unresolved row-menu click after a
        // later Select All click. The duplicate Select All makes that exact
        // ordering error identifiable without changing any other plan steps.
        for (var i = 0; i + 3 < plan.Steps.Count; i++)
        {
            var steps = plan.Steps;
            if (steps[i].Action != "context-click" || steps[i].Target is not
                { Process: "EXCEL", ClassName: "XLGridRowHeader" } ||
                steps[i + 1].Action != "click" || steps[i + 1].Target is not
                { Process: "EXCEL", ClassName: "XLSelectAllHeader", Name: "Select All" } ||
                steps[i + 2].Action != "manual-click" || steps[i + 2].Target is not
                { Process: "EXCEL", Name: "row menu choice" } ||
                steps[i + 3].Action != "click" || steps[i + 3].Target != steps[i + 1].Target)
                continue;
            var repaired = steps.ToList();
            repaired.RemoveAt(i + 1);
            plan = plan with { Steps = repaired.Select((step, index) =>
                step with { Number = index + 1 }).ToList() };
            status($"Plan warning: removed an early duplicate Select All after row {steps[i].Target!.Name}; the saved row-menu choice now runs first.");
        }
        static bool RepeatedCharacterTypo(string recorded, string actual)
        {
            if (recorded.Length != actual.Length + 1) return false;
            for (var i = 0; i + 1 < recorded.Length; i++)
                if (char.ToUpperInvariant(recorded[i]) == char.ToUpperInvariant(recorded[i + 1]) &&
                    string.Equals(recorded.Remove(i, 1), actual, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
        var nearUserMappings = plan.Steps.Where(step => step.Action == "replace-all" && step.WhenUser is not null &&
            RepeatedCharacterTypo(step.WhenUser.Split('\\').Last(), Environment.UserName)).ToList();
        foreach (var replacement in plan.Steps.Where(step => step.Action == "replace-all"))
            if (replacement.Target is not { Window: "Find and Replace", Name: "Replace All" })
                throw new InvalidOperationException($"Step {replacement.Number}: Replace All targets '{replacement.Target?.Name}' in '{replacement.Target?.Window}'. Replay stopped before changing the worksheet.");
        for (var i = 0; i + 1 < plan.Steps.Count; i++)
        {
            if (plan.Steps[i] is { Action: "type", Target: { Process: "SearchHost", AutomationId: "SearchTextBox" } } search &&
                plan.Steps[i + 1] is { Action: "click", Target:
                    { Process: "explorer", ClassName: "UIProperty", Name: "Name" } })
                throw new InvalidOperationException($"Step {search.Number}: Windows Search text was followed by a File Explorer property click, not a recorded search result or launch action. Replay stopped before opening another application. Record the Search result selection or Enter key.");
        }
        foreach (var search in plan.Steps.Where(step => step.Action == "type" &&
            step.Target is { Process: "SearchHost", AutomationId: "SearchTextBox" }))
            if (string.IsNullOrWhiteSpace(search.Value))
                throw new InvalidOperationException($"Step {search.Number}: the Windows Search query is empty. Replay stopped before claiming the search succeeded; rebuild the plan from its recorded keys.");
        if (nearUserMappings.Count > 0)
        {
            var exactMapping = plan.Steps.Any(step => step.Action == "replace-all" && step.WhenUser is not null &&
                step.WhenUser.Split('\\').Last().Equals(Environment.UserName, StringComparison.OrdinalIgnoreCase));
            var names = nearUserMappings.Select(step => step.WhenUser).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (exactMapping || names.Count != 1)
                throw new InvalidOperationException("The saved plan has ambiguous user-specific replacement conditions. Review WhenUser before replay; no worksheet changes were made.");
            var recordedUser = names[0]!;
            plan = plan with { Steps = plan.Steps.Select(step => nearUserMappings.Contains(step)
                ? step with { WhenUser = Environment.UserName } : step).ToList() };
            status($"Plan warning: corrected the user condition '{recordedUser}' to the verified replay identity '{Environment.UserName}' for this run.");
        }
        for (var i = 1; i < plan.Steps.Count; i++)
        {
            var current = plan.Steps[i];
            if (current.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                current.Target.Name == "Delete" && current.Target.ControlType == "ControlType.MenuItem" &&
                plan.Steps[i - 1].Action == "context-click" &&
                plan.Steps[i - 1].Target is { } tab && IsExcelSheetTab(tab) &&
                plan.Steps.Skip(i + 1).Take(4).Any(next => next.Target?.Window == "Move or Copy"))
                throw new InvalidOperationException($"Step {current.Number}: the recording says Delete on sheet '{tab.Name}', but the following steps use Move or Copy. Replay stopped before changing the workbook. Review or re-record this menu choice.");
        }
        for (var i = 0; i < plan.Steps.Count; i++)
        {
            var copy = plan.Steps[i];
            if (copy.Action != "click" || copy.Target is not { Window: "Move or Copy", Name: "Create a copy" }) continue;
            var completion = plan.Steps.Skip(i + 1).TakeWhile(next =>
                next.Target?.Window == "Move or Copy" || next.Action == "rename-sheet" ||
                next.Target is { Process: "EXCEL", ControlType: "ControlType.TabItem" }).ToList();
            if (!completion.Any(next => next.Target is { Window: "Move or Copy", Name: "OK" } ||
                                        next.Action == "rename-sheet"))
                throw new InvalidOperationException($"Step {copy.Number}: Create a copy has no recorded OK action or grouped sheet rename to confirm the dialog. Replay stopped before changing the workbook.");
        }
        for (var i = 0; i + 1 < plan.Steps.Count; i++)
        {
            var rowMenu = plan.Steps[i];
            var following = plan.Steps[i + 1];
            if (rowMenu.Action == "context-click" && rowMenu.Target is
                { Process: "EXCEL", ClassName: "XLGridRowHeader" } &&
                following.Target is { Process: "EXCEL", Name: "Select All",
                    ClassName: "XLSelectAllHeader" })
                throw new InvalidOperationException($"Step {rowMenu.Number}: the recording opened the row menu, then selected the worksheet without saving a menu choice. Replay stopped before editing the sheet. Review the recorded screenshot and add the missing row action to the plan.");
            if (rowMenu.Action == "context-click" && rowMenu.Target is
                { Process: "EXCEL", ClassName: "XLGridRowHeader" } &&
                following.Action == "click" && following.Target is
                { Process: "EXCEL", ClassName: "XLSpreadsheetCell" })
                throw new InvalidOperationException($"Step {rowMenu.Number}: the row menu was followed by worksheet cell '{following.Target.Name}' instead of a recorded menu command. Replay stopped before editing the sheet.");
        }
        // A recording can begin using a view that was already open. Check that
        // view before earlier steps change another application. Otherwise a
        // missing navigation is only discovered after workbook edits.
        for (var i = 0; i < plan.Steps.Count; i++)
        {
            var selection = plan.Steps[i];
            if (selection.Action is not ("select-first-row" or "select-all-items") ||
                selection.Target is not { } target || target.ControlType == "ControlType.HeaderItem" ||
                string.IsNullOrWhiteSpace(target.Process) || string.IsNullOrWhiteSpace(target.Window)) continue;
            var earlier = plan.Steps.Take(i).ToList();
            if (!earlier.Any(step => step.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                    step.Action is "rename-sheet" or "replace-all" or "type" or "key") ||
                earlier.Any(step => step.Target?.Process?.Equals(target.Process, StringComparison.OrdinalIgnoreCase) == true &&
                    step.Target.Window != target.Window && step.Action is "click" or "context-click")) continue;
            var window = FindDesktopWindow(target.Process, target.Window, exactOnly: true);
            if (window == 0)
            {
                var applicationName = target.Process.EndsWith("App", StringComparison.OrdinalIgnoreCase)
                    ? target.Process[..^3] : target.Process;
                var hasTaskbarSwitch = earlier.Any(step => IsTaskbarNavigation(step) &&
                    step.Target?.Name?.Contains(applicationName, StringComparison.OrdinalIgnoreCase) == true);
                var reason = hasTaskbarSwitch
                    ? $"The recording switches to {target.Process}, but does not contain navigation that opens '{target.Window}'. That exact window was not open at replay time."
                    : $"'{target.Window}' was not open, and the recording contains no action that opens this view.";
                throw new InvalidOperationException($"Step {selection.Number}: {reason} Replay stopped before changing the workbook.");
            }
            ShowWindow(window, 9);
            SetForegroundWindow(window);
            if (await FindFirstListRowAsync(target, token) is null)
                throw new InvalidOperationException($"Step {selection.Number}: '{target.Window}' has no list row, and the recording has no action that opens or fills this view. Replay stopped before changing the workbook.");
            status($"Preflight: '{target.Window}' is open with a list row for step {selection.Number}.");
        }
        var knownWindows = plan.Steps.Select(s => s.Target?.Window)
            .Where(name => !string.IsNullOrWhiteSpace(name)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        ControlRef? copiedSheetTarget = null;
        HashSet<string>? tabsBeforeCopy = null;
        string? previousProcess = null;
        string? previousWindow = null;
        ControlRef? selectedRowsHeader = null;
        ControlRef? pendingExcelRowCheck = null;
        string? pendingExcelRowValue = null;
        string? pendingExcelFollowingRowValue = null;
        nint tuckedAwayDialog = 0;
        var replacementDialogOpen = false;
        var resolvedGridTargets = new Dictionary<string, ControlRef>(StringComparer.OrdinalIgnoreCase);
        var approvedBrowserWindows = new HashSet<nint>();
        int? freshBrowserProcessId = null;
        int? selectedBrowserProcessId = null;
        try
        {
        for (var index = 0; index < plan.Steps.Count; index++)
        {
            var step = plan.Steps[index];
            token.ThrowIfCancellationRequested();
            if (step.Action == "open-browser-process")
            {
                freshBrowserProcessId = await LaunchFreshBrowserAsync(step.Target?.BrowserLaunch,
                    step.Target?.Process ?? "msedge", status, token);
                previousProcess = null;
                previousWindow = null;
                status($"Completed step {step.Number}: opened an Edge window with the default profile.");
                continue;
            }
            if (step.Action == "click" && step.Target is
                { Process: "EXCEL", Name: "OK", ControlType: "ControlType.Button" } &&
                index > 0 && plan.Steps[index - 1].Action == "replace-all" &&
                plan.Steps[index - 1].Target?.Process?.Equals("EXCEL",
                    StringComparison.OrdinalIgnoreCase) == true)
            {
                status($"Skipped step {step.Number}: the preceding Replace All step already verified and dismissed Excel's result dialog.");
                continue;
            }
            if (step.Action == "type" && step.Value == "a" && step.Target is
                { Process: "EXCEL", Window: "Find and Replace",
                  ControlType: "ControlType.Window" } &&
                plan.Steps.Skip(index + 1).Take(2).Any(following => following is
                    { Action: "type", Value: not null, Target:
                        { Process: "EXCEL", Window: "Find and Replace",
                          AutomationId: "18", ClassName: "EDTBX" } }))
            {
                status($"Skipped step {step.Number}: the recorded window-level 'a' was superseded by the next verified Find field value.");
                continue;
            }
            if (step.Action == "context-selection" && step.Target is { } selectedTarget &&
                index + 1 < plan.Steps.Count)
            {
                var commandIndex = index + 1;
                while (commandIndex < Math.Min(plan.Steps.Count, index + 5) &&
                    (IsTaskbarNavigation(plan.Steps[commandIndex]) ||
                     plan.Steps[commandIndex] is { Action: "scroll", Target: { } intervening } &&
                     !string.Equals(intervening.Process, selectedTarget.Process, StringComparison.OrdinalIgnoreCase)))
                    commandIndex++;
                if (commandIndex > index + 1 && commandIndex < plan.Steps.Count &&
                    plan.Steps[commandIndex] is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } command } &&
                    string.Equals(command.Process, selectedTarget.Process, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(command.Window, selectedTarget.Window, StringComparison.OrdinalIgnoreCase))
                {
                    var reordered = plan.Steps.ToList();
                    var menuStep = reordered[commandIndex];
                    reordered.RemoveAt(commandIndex);
                    reordered.Insert(index + 1, menuStep);
                    plan = plan with { Steps = reordered };
                    status($"Using recorded action order: '{command.Name}' before navigation and scrolling in the next app.");
                }
            }
            if (step.Action == "click" && step.Target?.ControlType == "ControlType.MenuItem" &&
                index + 1 < plan.Steps.Count && plan.Steps[index + 1] is { Action: "click",
                    Target: { ControlType: "ControlType.MenuItem" } parent } following &&
                step.Target.ParentName == parent.Name && step.Target.Process == parent.Process &&
                step.Target.Window == parent.Window)
            {
                plan = plan with { Steps = plan.Steps.ToList() };
                (plan.Steps[index], plan.Steps[index + 1]) = (following, step);
                step = plan.Steps[index];
                status($"Using recorded menu order: '{parent.Name}' before '{plan.Steps[index + 1].Target!.Name}'.");
            }
            if (IsTaskbarNavigation(step))
            {
                // The next recorded application's own target is the reliable way
                // to foreground it. Taskbar buttons do not consistently expose
                // an invokable UIA pattern, and taskbar-only steps have no effect
                // on the work being replayed.
                status($"Skipped step {step.Number}: taskbar navigation to '{step.Target?.Name}'; the next application step activates its window.");
                continue;
            }
            if (step.Action == "click" && step.Target is
                { Process: "explorer", Window: "Program Manager", Name: "Desktop",
                  ControlType: "ControlType.List", ClassName: "SysListView32" })
            {
                status($"Skipped step {step.Number}: Desktop background click has no replay effect.");
                continue;
            }
            if (step.WhenUser is not null &&
                !Environment.UserName.Equals(step.WhenUser.Split('\\').Last(), StringComparison.OrdinalIgnoreCase))
            {
                status($"Skipped step {step.Number}: this replacement applies to {step.WhenUser}.");
                continue;
            }
            if (step.Target is { ControlType: "ControlType.DataItem" } recordedCell)
            {
                var key = $"{recordedCell.Process}|{recordedCell.Window}|{recordedCell.AutomationId}";
                if (step.TargetStrategy?.StartsWith("visible-row:", StringComparison.Ordinal) == true)
                {
                    if (!int.TryParse(step.TargetStrategy["visible-row:".Length..], out var ordinal) || ordinal < 1)
                        throw new InvalidDataException($"Step {step.Number}: the saved visible-row position is invalid.");
                    var visibleCell = await Task.Run(() => Automation.VisibleGridCell(recordedCell, ordinal), token);
                    var resolved = Automation.Describe(visibleCell);
                    if (resolved?.ControlType != "ControlType.DataItem" || string.IsNullOrWhiteSpace(resolved.AutomationId))
                        throw new InvalidOperationException($"Step {step.Number}: could not identify visible result row {ordinal} in the recorded column.");
                    resolvedGridTargets[key] = resolved;
                    step = step with { Target = resolved };
                    status($"Step {step.Number}: selected visible result row {ordinal} in the recorded column.");
                }
                else if (resolvedGridTargets.TryGetValue(key, out var liveCell))
                    step = step with { Target = liveCell };
            }
            if (step.Action == "close-window")
            {
                if (step.Target is null) throw new InvalidOperationException($"Step {step.Number}: window to close is missing.");
                await CloseRecordedWindowAsync(step.Target, token);
                replacementDialogOpen = false;
                status($"Completed step {step.Number}: closed {step.Target.Window}.");
                continue;
            }
            if (replacementDialogOpen && step.Action != "replace-all" &&
                step.Target?.Window != "Find and Replace")
            {
                await CloseCompletedReplacementDialogAsync(token);
                replacementDialogOpen = false;
                if (step.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true)
                {
                    var workbook = FindDesktopWindow("EXCEL", step.Target.Window, exactOnly: true);
                    if (workbook != 0) SetForegroundWindow(workbook);
                }
                status("Closed Find and Replace before continuing in the worksheet.");
            }
            if (step.Action == "click" && step.Target?.ControlType == "ControlType.Tab")
            {
                status($"Skipping non-actionable container at step {step.Number}");
                continue;
            }
            if (step.Action == "key" && ActionGrouper.IsStandaloneModifier(step.Key))
            {
                status($"Skipping standalone modifier at step {step.Number}");
                continue;
            }
            if (step.Action == "key" && step.Key is "Control+V" or "Ctrl+V" &&
                step.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ClassName == "XLDESK" &&
                plan.Steps.Skip(index + 1).Take(4).Any(next =>
                    next.Action == "key" && next.Key is "Control+Z" or "Ctrl+Z" &&
                    next.Target?.ClassName == "XLDESK"))
            {
                status($"Skipped step {step.Number}: this worksheet-pane paste was undone during recording.");
                continue;
            }
            if (step.Action == "key" && step.Key is "Control+Z" or "Ctrl+Z" &&
                step.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ClassName == "XLDESK" &&
                plan.Steps.Take(index).TakeLast(4).Any(previous =>
                    previous.Action == "key" && previous.Key is "Control+V" or "Ctrl+V" &&
                    previous.Target?.ClassName == "XLDESK"))
            {
                status($"Skipped step {step.Number}: its matching worksheet-pane paste was skipped.");
                continue;
            }
            if (step.Action == "key" && step.Target?.ClassName == "EDTBX" &&
                plan.Steps.Skip(index + 1).Any(next => next.Action == "type" &&
                    next.Target?.Window == step.Target.Window &&
                    next.Target?.AutomationId == step.Target.AutomationId))
            {
                status($"Skipped step {step.Number}: the next recorded field value supersedes this keystroke.");
                continue;
            }
            if (step.Action == "key" && step.Target?.ClassName == "EDTBX" &&
                step.Key is "Left" or "Right" or "Home" or "End" or "Back" or "Delete")
            {
                status($"Skipped step {step.Number}: cursor editing is already reflected in the recorded field value.");
                continue;
            }
            if (PlanCompactor.FollowingEditValue(plan.Steps, index) is { } editValue)
            {
                status($"Skipped step {step.Number}: the click only focused Edit control " +
                    $"{step.Target!.AutomationId}; step {editValue.Number} sets and verifies its value.");
                continue;
            }
            if (step is { Action: "manual-click", ExpectedState: null,
                    Target: { ControlType: "ControlType.Window", Name: { Length: > 0 } pageName,
                        Process: { Length: > 0 } pageProcess } } &&
                index + 1 < plan.Steps.Count && plan.Steps[index + 1] is
                    { Action: "click", Target: { ControlType: "ControlType.Button",
                        Name: { Length: > 0 }, Window: { Length: > 0 } pageButtonWindow } pageButton } nextStep &&
                pageButton.Process?.Equals(pageProcess, StringComparison.OrdinalIgnoreCase) == true &&
                pageButtonWindow.Equals(pageName, StringComparison.OrdinalIgnoreCase))
            {
                var nextControl = await FindAsync(pageButton, token, 2);
                try
                {
                    if (nextControl is not null && nextControl.Current.IsEnabled &&
                        !nextControl.Current.IsOffscreen)
                    {
                        status($"Skipped step {step.Number}: the unresolved page click has no recorded " +
                            $"state change, and '{pageButton.Name}' is already enabled on '{pageName}'.");
                        continue;
                    }
                }
                catch (Exception ex) when (ex is ElementNotAvailableException or
                    InvalidOperationException or COMException) { Trace.WriteLine(ex); }
            }
            if (step.Action == "manual-click")
            {
                var selectedLegacyListItem = false;
                var dependentTarget = index + 1 < plan.Steps.Count &&
                    plan.Steps[index + 1].Target is { } nextTarget &&
                    step.Target?.Process?.Equals(nextTarget.Process, StringComparison.OrdinalIgnoreCase) == true
                    ? nextTarget : null;
                if (step.Target is { ControlType: "ControlType.DataItem", Name: { Length: > 0 },
                    Process: { Length: > 0 } } clickableItem)
                {
                    for (var clickAttempt = 0; clickAttempt < 2 && !selectedLegacyListItem; clickAttempt++)
                    {
                        var item = Automation.UniqueVisibleItemByName(clickableItem.Process, clickableItem.Name);
                        if (item is null || !TryClickLiveUiaElement(item, null, false, out var clickDetail))
                            break;
                        if (!await DependentControlReadyAsync(dependentTarget, token))
                            continue;
                        status($"Completed step {step.Number}: clicked '{clickableItem.Name}' and confirmed the next control is ready. {clickDetail}");
                        previousProcess = clickableItem.Process;
                        previousWindow = Automation.Describe(item)?.Window ?? clickableItem.Window;
                        selectedLegacyListItem = true;
                    }
                }
                if (step.Target is { ControlType: "ControlType.DataItem", Name: { Length: > 0 },
                    Process: { Length: > 0 }, Window: { Length: > 0 } } listTarget)
                {
                    var candidateWindows = new List<string> { listTarget.Window };
                    if (index + 1 < plan.Steps.Count && plan.Steps[index + 1].Target is
                        { Process: { } nextProcess, Window: { Length: > 0 } nextWindow } &&
                        nextProcess.Equals(listTarget.Process, StringComparison.OrdinalIgnoreCase) &&
                        !candidateWindows.Contains(nextWindow, StringComparer.OrdinalIgnoreCase))
                        candidateWindows.Add(nextWindow);
                    foreach (var candidateWindow in candidateWindows)
                    {
                        var listWindow = FindDesktopWindow(listTarget.Process, candidateWindow, exactOnly: true);
                        if (selectedLegacyListItem || listWindow == 0 ||
                            !MsaaActions.TrySelectUnique(listWindow, listTarget.Name) ||
                            !await DependentControlReadyAsync(dependentTarget, token))
                            continue;
                        status($"Completed step {step.Number}: selected unique legacy list item '{listTarget.Name}' and confirmed the next control is ready.");
                        previousProcess = listTarget.Process;
                        previousWindow = candidateWindow;
                        selectedLegacyListItem = true;
                        break;
                    }
                }
                if (!selectedLegacyListItem && step.Target is
                    { ControlType: "ControlType.DataItem", Name: { Length: > 0 },
                      Process: { Length: > 0 } } namedItem)
                {
                    var selectable = Automation.UniqueSelectableByName(namedItem.Process, namedItem.Name);
                    if (selectable is not null && selectable.TryGetCurrentPattern(
                        SelectionItemPattern.Pattern, out var selectionObject))
                    {
                        var selection = (SelectionItemPattern)selectionObject;
                        selection.Select();
                        if (selection.Current.IsSelected &&
                            await DependentControlReadyAsync(dependentTarget, token))
                        {
                            status($"Completed step {step.Number}: selected '{namedItem.Name}' and confirmed the next control is ready.");
                            previousProcess = namedItem.Process;
                            previousWindow = Automation.Describe(selectable)?.Window ?? namedItem.Window;
                            selectedLegacyListItem = true;
                        }
                    }
                }
                if (selectedLegacyListItem) continue;
                var manualSearchResult = step.Target is
                    { Process: "SearchHost", Name: "Windows Search result" };
                ControlRef? rowCheck = null;
                string? rowValueBefore = null;
                if (step.Target?.Name == "row menu choice" && index > 0 &&
                    plan.Steps[index - 1] is { Action: "context-click", Target: { Process: "EXCEL", Name: { } rowNumber } } &&
                    int.TryParse(rowNumber, out _))
                {
                    rowCheck = pendingExcelRowCheck;
                    rowValueBefore = pendingExcelRowValue;
                    if (rowCheck is null || rowValueBefore is null)
                        throw new InvalidOperationException($"Step {step.Number}: row {rowNumber} was not measured before its menu opened. Replay stopped before later worksheet actions.");
                }
                var problem = step.Target?.ControlType == "ControlType.ToolTip"
                    ? "a pop-up hint covered the control you clicked"
                    : "the app recorded a click but could not tell which control you clicked";
                var help = $"Replay needs help: in step {step.Number}, {problem}. " +
                    "Click Open recording PDF below, find that step's screenshot, and make the click in the target program, " +
                    "then choose I've handled it. Choose Stop replay if the intended action is unclear.";
                if (manualSearchResult)
                    help = $"Replay needs help: in step {step.Number}, click the intended result in Windows Search and wait for its application window to open. " +
                        "Then choose I've handled it. Replay will check that Windows Search closed.";
                status(help);
                if (!handleUnexpectedDialog(help)) throw new OperationCanceledException("Replay stopped by user.", token);
                if (step.Target?.ControlType == "ControlType.DataItem" &&
                    !await DependentControlReadyAsync(dependentTarget, token))
                    throw new InvalidOperationException($"Step {step.Number}: the next recorded control did not become ready after selecting '{step.Target.Name}'.");
                if (manualSearchResult)
                {
                    await VerifyRecordedSearchLaunchAsync(step, token);
                }
                if (rowCheck is not null)
                {
                    string? rowValueAfter = null;
                    for (var attempt = 0; attempt < 4 && rowValueAfter is null; attempt++)
                    {
                        var cellAfter = await FindAsync(rowCheck, token, 2);
                        try
                        {
                            if (cellAfter?.TryGetCurrentPattern(ValuePattern.Pattern, out var afterValue) == true)
                                rowValueAfter = ((ValuePattern)afterValue).Current.Value;
                        }
                        catch (ElementNotAvailableException) { }
                        if (rowValueAfter is null) await Task.Delay(100, token);
                    }
                    if (rowValueAfter is null || rowValueAfter == rowValueBefore)
                        throw new InvalidOperationException($"Step {step.Number}: row 2 did not show a changed value in column B after the manual menu choice. Replay stopped before Select All.");
                    status($"Step {step.Number}: confirmed that the row menu action changed the former row 2 value.");
                }
                if (manualSearchResult)
                    status($"Completed step {step.Number}: manual Windows Search result action passed its recorded checks.");
                else if (step.Target is { ControlType: "ControlType.DataItem", Name: { Length: > 0 },
                    Process: { Length: > 0 } } manualItem)
                {
                    var selected = Automation.UniqueSelectableByName(manualItem.Process, manualItem.Name);
                    var confirmed = selected is not null && selected.TryGetCurrentPattern(
                        SelectionItemPattern.Pattern, out var selectedPattern) &&
                        ((SelectionItemPattern)selectedPattern).Current.IsSelected;
                    status(confirmed
                        ? $"Completed step {step.Number}: confirmed manual selection of '{manualItem.Name}'."
                        : $"Completed step {step.Number}: user handled '{manualItem.Name}', but its selection state was not exposed for verification.");
                }
                else
                {
                    var observed = Automation.Describe(AutomationElement.FocusedElement);
                    status($"Completed step {step.Number}: user handled the unverified recorded click. " +
                        $"Focused control afterward: {observed?.Name ?? "unknown"} ({observed?.ControlType ?? "unknown type"}).");
                }
                previousProcess = null;
                previousWindow = null;
                continue;
            }
            if (step.Action == "delete-row")
            {
                var rowTarget = index > 0 && plan.Steps[index - 1] is
                    { Action: "context-click", Target: { Process: "EXCEL", ClassName: "XLGridRowHeader" } }
                    ? plan.Steps[index - 1].Target : null;
                if (rowTarget is null || !int.TryParse(rowTarget.Name, out var rowNumber))
                    throw new InvalidOperationException($"Step {step.Number}: Delete is not preceded by a verified Excel row-header context-click.");
                if (step.Target is not { Process: "EXCEL", ControlType: "ControlType.MenuItem", Name: "Delete" } &&
                    step.Target != rowTarget)
                    throw new InvalidOperationException($"Step {step.Number}: the recorded row command is '{step.Target?.Name}' ({step.Target?.ControlType}), not a verified Delete action.");
                if (pendingExcelRowCheck is null || pendingExcelRowValue is null)
                    throw new InvalidOperationException($"Step {step.Number}: B{rowNumber} was not measured before opening the row menu.");
                if (pendingExcelFollowingRowValue is null ||
                    pendingExcelRowValue == pendingExcelFollowingRowValue)
                    throw new InvalidOperationException($"Step {step.Number}: B{rowNumber + 1} is unavailable or matches B{rowNumber}, so the row shift cannot be verified safely.");
                var invoked = false;
                if (step.Target?.ControlType == "ControlType.MenuItem")
                {
                    var liveDelete = await FindAsync(step.Target, token, 2);
                    if (liveDelete is not null)
                    {
                        try { invoked = TryClickLiveUiaElement(liveDelete, null, rightClick: false, out _); }
                        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
                        { Trace.WriteLine(ex); }
                    }
                }
                if (!invoked)
                {
                    SendKeys.SendWait("{ESC}");
                    var rowHeader = await FindAsync(rowTarget, token, 3) ??
                        throw new InvalidOperationException($"Step {step.Number}: row header {rowNumber} could not be reacquired for Delete.");
                    var firstColumnTarget = pendingExcelRowCheck with
                    { AutomationId = "A" + rowNumber, Name = "A" + rowNumber };
                    if (!TryClickLiveUiaElement(rowHeader, firstColumnTarget, rightClick: false,
                        out var rowSelectionDetail))
                        throw new InvalidOperationException($"Step {step.Number}: row {rowNumber} could not be selected for deletion. {rowSelectionDetail}");
                    await Task.Delay(120, token);
                    SendKeys.SendWait("^-");
                }
                var shifted = false;
                for (var attempt = 0; attempt < 20; attempt++)
                {
                    await Task.Delay(100, token);
                    var cell = await FindAsync(pendingExcelRowCheck, token, 1);
                    try
                    {
                        if (cell?.TryGetCurrentPattern(ValuePattern.Pattern, out var value) == true &&
                            ((ValuePattern)value).Current.Value == pendingExcelFollowingRowValue)
                        { shifted = true; break; }
                    }
                    catch (ElementNotAvailableException) { }
                }
                if (!shifted)
                    throw new InvalidOperationException($"Step {step.Number}: row {rowNumber} did not shift to the former next-row value after Delete. Replay stopped before Select All.");
                status($"Completed step {step.Number}: deleted row {rowNumber} and verified the following row moved into its place.");
                pendingExcelRowCheck = null;
                pendingExcelRowValue = null;
                pendingExcelFollowingRowValue = null;
                continue;
            }
            if (step.Target is null) throw new InvalidOperationException($"Step {step.Number} has no UI Automation target.");
            if (step.Action == "type" && step.Target is
                { Process: "SearchHost", AutomationId: "SearchTextBox" })
            {
                var searchBox = await EnsureWindowsSearchOpenAsync(step, token);
                if (searchBox is not null)
                {
                    SetRecordedEditField(searchBox, step);
                    status($"Completed step {step.Number}: entered and verified the recorded text in Windows Search.");
                }
                else
                {
                    var desired = step.Value ?? throw new InvalidOperationException($"Step {step.Number}: Windows Search has no recorded query.");
                    SendKeys.SendWait("^a");
                    SendKeys.SendWait(string.Concat(desired.Select(EscapeText)));
                    status($"Completed step {step.Number}: entered the recorded text in foreground Windows Search; its field was unavailable for value verification.");
                }
                previousProcess = null;
                previousWindow = null;
                continue;
            }
            if (step.Target is { Process: "SearchHost", AutomationId: "SearchTextBox" } &&
                step.Action == "key" && step.Key is "Control+V" or "Ctrl+V")
            {
                var searchBox = await EnsureWindowsSearchOpenAsync(step, token);
                searchBox?.SetFocus();
                keybd_event(0x11, 0, 0, 0);
                keybd_event(0x56, 0, 0, 0);
                keybd_event(0x56, 0, 2, 0);
                keybd_event(0x11, 0, 2, 0);
                status($"Completed step {step.Number}: pasted clipboard text into Windows Search.");
                previousProcess = null;
                previousWindow = null;
                continue;
            }
            if (step.Target is { Process: "SearchHost", AutomationId: "SearchTextBox" } &&
                step.Action == "key" && step.Key is not ("Enter" or "Return") &&
                !string.IsNullOrWhiteSpace(step.Key))
            {
                var searchBox = await EnsureWindowsSearchOpenAsync(step, token);
                searchBox?.SetFocus();
                SendKeys.SendWait(ToSendKeys(step.Key));
                status($"Completed step {step.Number}: replayed {step.Key} in Windows Search.");
                previousProcess = null;
                previousWindow = null;
                continue;
            }
            if (step.Target.Process?.Equals("SearchHost", StringComparison.OrdinalIgnoreCase) == true &&
                step.Action == "key" && step.Key is "Enter" or "Return")
            {
                var searchBoxTarget = step.Target with { AutomationId = "SearchTextBox", Name = null,
                    ControlType = "ControlType.Edit", ClassName = null };
                var searchBox = await EnsureWindowsSearchOpenAsync(step with { Target = searchBoxTarget }, token);
                if (searchBox is not null && searchBox.TryGetCurrentPattern(ValuePattern.Pattern, out var searchValue) &&
                    string.IsNullOrWhiteSpace(((ValuePattern)searchValue).Current.Value))
                    throw new InvalidOperationException($"Step {step.Number}: Windows Search is empty after the recorded paste; Enter was not sent.");
                var windowsBeforeLaunch = VisibleApplicationWindows();
                await Task.Delay(900, token); // Let Search update its selected result after text entry.
                SendKeys.SendWait("{ENTER}");
                status($"Step {step.Number}: pressed Enter in Windows Search; if Windows requests elevation, approve it yourself while replay waits for the application.");
                await Task.Delay(1200, token);
                if (WindowsSearchStillActive(searchBoxTarget))
                {
                    SendKeys.SendWait("{ENTER}");
                    status($"Step {step.Number}: Windows Search still showed the results; retried Enter once.");
                }
                await VerifyRecordedSearchLaunchAsync(step, token);
                var expectedApplication = plan.Steps.Skip(index + 1).Select(candidate => candidate.Target)
                    .FirstOrDefault(target => !string.IsNullOrWhiteSpace(target?.Process) &&
                        !target.Process.Equals("SearchHost", StringComparison.OrdinalIgnoreCase));
                var launchedWindow = expectedApplication is null ? null :
                    await WaitForRecordedApplicationWindowAsync(expectedApplication.Process!,
                        expectedApplication.Window, token);
                launchedWindow ??= await WaitForNewApplicationWindowAsync(windowsBeforeLaunch, token);
                if (launchedWindow is null)
                {
                    var searchState = WindowsSearchStillActive(searchBoxTarget)
                        ? "Windows Search still shows the result, and no application window was observed. Open the recorded result"
                        : "Search closed, but no new application window was observed. If Windows requested administrator approval, complete that prompt";
                    var help = $"Replay needs help after step {step.Number}: {searchState} " +
                        "and wait for the application window, then choose I've handled it.";
                    status(help);
                    if (!handleUnexpectedDialog(help))
                        throw new OperationCanceledException("Replay stopped while waiting for the launched application.", token);
                    launchedWindow = expectedApplication is null ? null :
                        await WaitForRecordedApplicationWindowAsync(expectedApplication.Process!,
                            expectedApplication.Window, token);
                    launchedWindow ??= await WaitForNewApplicationWindowAsync(windowsBeforeLaunch, token);
                    if (launchedWindow is null)
                        throw new InvalidOperationException($"Step {step.Number}: Search closed, but no new application window was observed after manual handling.");
                }
                status($"Completed step {step.Number}: Search launched visible window '{launchedWindow.Value.Title}' ({launchedWindow.Value.Process}).");
                previousProcess = null;
                previousWindow = null;
                continue;
            }
            if (step.Target.Process?.Equals("SearchHost", StringComparison.OrdinalIgnoreCase) == true &&
                step.Action == "click")
            {
                var searchBoxTarget = step.Target with { AutomationId = "SearchTextBox", Name = null,
                    ControlType = "ControlType.Edit", ClassName = null };
                await EnsureWindowsSearchOpenAsync(step with { Target = searchBoxTarget }, token);
                var result = await FindAsync(step.Target, token, 5) ??
                    throw new InvalidOperationException($"Step {step.Number}: the recorded Windows Search result '{step.Target.Name}' was not found.");
                if (!TryClickLiveUiaElement(result, null, rightClick: false, out var detail))
                    throw new InvalidOperationException($"Step {step.Number}: the recorded Windows Search result could not be clicked. {detail}");
                status($"Completed step {step.Number}: clicked the recorded Windows Search result '{step.Target.Name}'.");
                previousProcess = null;
                previousWindow = null;
                continue;
            }
            if (step.Action == "resize-column")
            {
                if (!int.TryParse(step.Value, out var distance) || Math.Abs(distance) < 6)
                    throw new InvalidDataException($"Step {step.Number}: the recorded column resize has no usable drag distance.");
                var column = await FindAsync(step.Target, token, 4);
                if (column is null || column.Current.IsOffscreen)
                    throw new InvalidOperationException($"Step {step.Number}: the recorded column header is not visible. Scroll it into view and retry.");
                var bounds = column.Current.BoundingRectangle;
                if (bounds.IsEmpty || bounds.Width < 8 || bounds.Height < 8)
                    throw new InvalidOperationException($"Step {step.Number}: the live column header has no usable bounds.");
                var startX = step.Key == "left"
                    ? (int)Math.Round(bounds.Left + 2) : (int)Math.Round(bounds.Right - 2);
                var y = (int)Math.Round(bounds.Top + bounds.Height / 2);
                var endX = startX + distance;
                if (!Screen.AllScreens.Any(screen => screen.Bounds.Contains(startX, y) &&
                    screen.Bounds.Contains(endX, y)))
                    throw new InvalidOperationException($"Step {step.Number}: the recorded resize ends outside the visible desktop.");
                if (!SetCursorPos(startX, y)) throw new InvalidOperationException("Could not reach the live column edge.");
                mouse_event(0x0002u, 0, 0, 0, 0);
                try
                {
                    for (var part = 1; part <= 6; part++)
                    {
                        SetCursorPos(startX + distance * part / 6, y);
                        await Task.Delay(35, token);
                    }
                }
                finally { mouse_event(0x0004u, 0, 0, 0, 0); }
                await Task.Delay(180, token);
                var updated = await FindAsync(step.Target, token, 2);
                var nextBounds = updated?.Current.BoundingRectangle;
                if (nextBounds is null || nextBounds.Value.IsEmpty ||
                    Math.Abs((step.Key == "left" ? nextBounds.Value.Left : nextBounds.Value.Right) -
                        (step.Key == "left" ? bounds.Left : bounds.Right)) < 3)
                    throw new InvalidOperationException($"Step {step.Number}: the boundary drag did not visibly resize the column. Check the live column edge before retrying.");
                status($"Completed step {step.Number}: resized {step.Target.Name} using its live {step.Key ?? "right"} boundary.");
                continue;
            }
            if (step.Action == "replace-all")
            {
                var result = await ApplyRecordedReplacementAsync(step, token);
                replacementDialogOpen = true;
                status($"Completed step {step.Number}: {result}");
                continue;
            }
            if (step.Action == "type" && step.Value == "a" &&
                step.Target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ControlType == "ControlType.DataItem" &&
                index + 1 < plan.Steps.Count && plan.Steps[index + 1].Target?.Window == "Find and Replace" &&
                plan.Steps.Skip(Math.Max(0, index - 3)).Take(Math.Min(3, index)).Any(previous =>
                    previous.Key is "Control+C" or "Ctrl+C" && previous.Target == step.Target))
            {
                status($"Skipped step {step.Number}: the recorded 'a' followed a copy shortcut and did not change the cell.");
                continue;
            }
            if (step.Action == "rename-sheet" && copiedSheetTarget is not null &&
                tabsBeforeCopy is not null)
            {
                var dialogStep = plan.Steps.Take(index).LastOrDefault(candidate =>
                    candidate.Target?.Process == copiedSheetTarget.Process &&
                    candidate.Target.Window == "Move or Copy");
                var dialogTarget = dialogStep?.Target ?? new ControlRef(copiedSheetTarget.Process!,
                    "Move or Copy", null, "Move or Copy", "ControlType.Window", null, null);
                var created = await ConfirmDialogAndWaitForNewTabAsync(dialogTarget,
                    copiedSheetTarget, tabsBeforeCopy, token);
                var recordedName = step.Target?.Name;
                if (!string.Equals(created, recordedName, StringComparison.OrdinalIgnoreCase))
                {
                    plan = plan with { Steps = plan.Steps.Select((candidate, position) =>
                        position >= index && candidate.Target is not null &&
                        IsExcelSheetTab(candidate.Target) && candidate.Target.Name == recordedName
                        ? candidate with { Target = candidate.Target with { Name = created } }
                        : candidate).ToList() };
                    step = plan.Steps[index];
                }
                status($"Confirmed the copy dialog and observed new sheet '{created}' before step {step.Number}.");
                copiedSheetTarget = null;
                tabsBeforeCopy = null;
            }
            PauseForUnexpectedDialog(step.Target.Process, knownWindows, handleUnexpectedDialog, status, token);
            var preservePageFocus = step.Action is "type" or "key" &&
                step.Target.ControlType == "ControlType.Pane" &&
                string.IsNullOrWhiteSpace(step.Target.Name) &&
                string.IsNullOrWhiteSpace(step.Target.AutomationId) &&
                FollowsSubmittedNavigation(plan.Steps, index) &&
                ForegroundProcessMatches(step.Target.Process);
            if (preservePageFocus)
            {
                previousProcess = step.Target.Process;
                previousWindow = step.Target.Window;
            }
            else if (!string.IsNullOrWhiteSpace(step.Target.Process) &&
                (!step.Target.Process.Equals(previousProcess, StringComparison.OrdinalIgnoreCase) ||
                 !string.Equals(step.Target.Window, previousWindow, StringComparison.OrdinalIgnoreCase)))
            {
                if (tuckedAwayDialog != 0 &&
                    string.Equals(step.Target.Window, GetNativeWindowTitle(tuckedAwayDialog), StringComparison.OrdinalIgnoreCase))
                {
                    ShowWindow(tuckedAwayDialog, 9);
                    tuckedAwayDialog = 0;
                }
                if (tuckedAwayDialog == 0 &&
                    step.Target.Process.Equals(previousProcess, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(previousWindow))
                {
                    var prior = FindDesktopWindow(previousProcess!, previousWindow, exactOnly: true);
                    if (prior != 0 && GetWindow(prior, 4) != 0)
                    {
                        ShowWindow(prior, 0);
                        tuckedAwayDialog = prior;
                        status($"Temporarily moved {previousWindow} out of the way for the next window's control.");
                    }
                }
                await EnsureApplicationWindowAsync(step.Target.Process, step.Target.Window, status, token,
                    step.Target.BrowserLaunch, step.Target, approveBrowserChange,
                    approvedBrowserWindows, freshBrowserProcessId,
                    processId => freshBrowserProcessId = processId,
                    selectedBrowserProcessId,
                    processId => selectedBrowserProcessId = processId);
                previousProcess = step.Target.Process;
                previousWindow = step.Target.Window;
            }
            if (step.Action == "click" && step.Target is
                { Process: "msedge", Name: "Maximize",
                  ClassName: "EdgeWindowsCaptionButton" })
            {
                var browserWindow = GetForegroundWindow();
                GetWindowThreadProcessId(browserWindow, out var browserProcessId);
                if (browserWindow == 0 || browserProcessId == 0 ||
                    !Process.GetProcessById((int)browserProcessId).ProcessName.Equals("msedge",
                        StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(BrowserWindowCategory(GetNativeWindowTitle(browserWindow)),
                        BrowserWindowCategory(step.Target.Window), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Step {step.Number}: the recorded Edge window was not foreground for Maximize.");
                MaximizeWindowIfSupported(browserWindow);
                if (!IsZoomed(browserWindow))
                    throw new InvalidOperationException($"Step {step.Number}: Edge did not become maximized.");
                status($"Completed step {step.Number}: verified the recorded Edge window is maximized.");
                continue;
            }
            if (step.Action == "click" &&
                step.ExpectedState?.Contains("Consent accepted", StringComparison.OrdinalIgnoreCase) == true &&
                index + 1 < plan.Steps.Count &&
                plan.Steps[index + 1].Target is { } followingConsentTarget &&
                string.Equals(followingConsentTarget.Process, step.Target.Process,
                    StringComparison.OrdinalIgnoreCase) &&
                string.Equals(followingConsentTarget.Window, step.Target.Window,
                    StringComparison.OrdinalIgnoreCase) &&
                await FindAsync(step.Target, token, 1) is null &&
                await FindAsync(followingConsentTarget, token, 2) is not null)
            {
                status($"Skipped step {step.Number}: the consent control is absent and the next recorded page control is already available.");
                continue;
            }
            if (step.Action == "open-private-window")
            {
                const string openedWindowPrefix = "opened-window:";
                var recordedPrivateWindow = step.ExpectedState is { } state &&
                    state.StartsWith(openedWindowPrefix, StringComparison.OrdinalIgnoreCase)
                    ? state[openedWindowPrefix.Length..] : null;
                var recordedCommand = step.Target.Name?.Contains("InPrivate",
                    StringComparison.OrdinalIgnoreCase) == true ||
                    recordedPrivateWindow?.Contains("[InPrivate]",
                        StringComparison.OrdinalIgnoreCase) == true;
                var followingWindow = index + 1 < plan.Steps.Count
                    ? plan.Steps[index + 1].Target?.Window : null;
                if (step.Target.Process?.Equals("msedge", StringComparison.OrdinalIgnoreCase) != true ||
                    !recordedCommand ||
                    index + 1 >= plan.Steps.Count ||
                    plan.Steps[index + 1].Target?.Process?.Equals(step.Target.Process,
                        StringComparison.OrdinalIgnoreCase) != true ||
                    followingWindow?.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase) != true ||
                    recordedPrivateWindow is not null &&
                    !string.Equals(followingWindow, recordedPrivateWindow,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"Step {step.Number}: the recorded private-window transition is incomplete.");
                var privateWindowsBefore = VisiblePrivateBrowserWindows(step.Target.Process).ToHashSet();
                SendKeys.SendWait("^+n");
                nint opened = 0;
                for (var attempt = 0; attempt < 60 && opened == 0; attempt++)
                {
                    token.ThrowIfCancellationRequested();
                    await Task.Delay(100, token);
                    opened = VisiblePrivateBrowserWindows(step.Target.Process)
                        .FirstOrDefault(handle => !privateWindowsBefore.Contains(handle));
                }
                if (opened == 0)
                    throw new InvalidOperationException($"Step {step.Number}: the private-window command did not create a new visible window. Replay stopped before the next browser action.");
                SetForegroundWindow(opened);
                GetWindowThreadProcessId(opened, out var privateProcessId);
                selectedBrowserProcessId = (int)privateProcessId;
                status($"Completed step {step.Number}: opened and verified a new private browser window.");
                previousWindow = null;
                continue;
            }
            if (step.Action == "key" && step.Key is "Control+F" or "Ctrl+F" &&
                step.Target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true)
            {
                // A recorded shortcut can refer to a header or cell whose UIA element
                // is invalidated by the preceding selection. The shortcut belongs to
                // the workbook window, not to that transient element.
                var workbook = FindDesktopWindow(step.Target.Process, step.Target.Window);
                if (workbook == 0)
                    throw new InvalidOperationException($"Step {step.Number}: Excel workbook is unavailable.");
                SetForegroundWindow(workbook);
                var afterDelete = index > 0 && plan.Steps[index - 1].Target?.Name == "Delete";
                SendKeys.SendWait(afterDelete ? "{ESC}^{HOME}^f" : "^f");
                if (!await WaitForWindowAsync("Find and Replace", token))
                    throw new InvalidOperationException($"Step {step.Number}: Find and Replace did not open.");
                status($"Completed step {step.Number}: opened Find and Replace from the workbook");
                continue;
            }
            if (step.Action is "click" or "select-all-items" or "ensure-state" &&
                step.Target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ControlType == "ControlType.HeaderItem" &&
                step.Target.Name == "Select All")
            {
                var clicked = false;
                string selectDetail = "the live header was unavailable";
                for (var attempt = 0; attempt < 3 && !clicked; attempt++)
                {
                    var header = await FindAsync(step.Target, token, attempt == 0 ? 4 : 2);
                    if (header is null) { selectDetail = "the header was not exposed"; continue; }
                    try
                    {
                        ActivateWindow(header);
                        await Task.Delay(120, token);
                        header = await FindAsync(step.Target, token, 2) ?? header;
                        clicked = TryClickLiveUiaElement(header, null, rightClick: false, out selectDetail);
                    }
                    catch (ElementNotAvailableException)
                    { selectDetail = "the header became stale after workbook activation"; }
                }
                if (!clicked)
                {
                    var workbook = FindDesktopWindow(step.Target.Process!, step.Target.Window, exactOnly: true);
                    if (workbook == 0)
                        throw new InvalidOperationException($"Step {step.Number}: Select All header was unavailable and the workbook could not be found. {selectDetail}");
                    ShowWindow(workbook, 9);
                    SetForegroundWindow(workbook);
                    if (GetForegroundWindow() != workbook)
                        throw new InvalidOperationException($"Step {step.Number}: the workbook did not take focus for Select All. {selectDetail}");
                    status($"Step {step.Number}: Select All header was stale; using Excel's select-all shortcut in the verified workbook.");
                }
                if (step.Target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true)
                {
                    SendKeys.SendWait("^a"); // Excel may first select the current table.
                    SendKeys.SendWait("^a");
                }
                status(clicked
                    ? $"Completed step {step.Number}: clicked the live Select All header. {selectDetail}"
                    : $"Completed step {step.Number}: selected all in the verified Excel workbook.");
                continue;
            }
            if (step.Action == "click" && step.Target.ControlType == "ControlType.Table" &&
                string.IsNullOrWhiteSpace(step.ExpectedState) && index + 1 < plan.Steps.Count &&
                plan.Steps[index + 1].Action == "click" &&
                plan.Steps[index + 1].Target?.Process == step.Target.Process &&
                plan.Steps[index + 1].Target?.ControlType is "ControlType.Button" or "ControlType.MenuItem")
            {
                status($"Skipped step {step.Number}: the table click only focused the view before the next menu action.");
                continue;
            }
            // An owner-drawn dialog can report the control from the previous wizard
            // page at the click point. The following list item identifies the live
            // combo box more reliably than that stale click target.
            if (step.Action == "click" && step.Target is { Process: { Length: > 0 },
                    Window: { Length: > 0 } } comboClick &&
                index + 1 < plan.Steps.Count && plan.Steps[index + 1] is
                    { Action: "click", Target: { ControlType: "ControlType.ListItem",
                        Name: { Length: > 0 }, ParentName: { Length: > 0 } } option } nextOption &&
                option.Process?.Equals(comboClick.Process, StringComparison.OrdinalIgnoreCase) == true &&
                option.Window?.Equals(comboClick.Window, StringComparison.OrdinalIgnoreCase) == true &&
                comboClick.ControlType is "ControlType.RadioButton" or "ControlType.Group" or "ControlType.Button" &&
                TrySelectUniqueNativeComboValue(comboClick, option.Name, out var selectedDetail))
            {
                status($"Completed steps {step.Number}-{nextOption.Number}: {selectedDetail} " +
                    $"The recorded '{comboClick.Name}' click belonged to an earlier dialog state.");
                index++;
                continue;
            }
            if (step.Action == "click" && step.Target is { Process: { Length: > 0 },
                    ControlType: "ControlType.Spinner" or "ControlType.RadioButton" } focusClick)
            {
                var followingIndex = index + 1;
                while (followingIndex < plan.Steps.Count && followingIndex - index < 4 &&
                    plan.Steps[followingIndex] is { Action: "click", Target:
                        { ControlType: "ControlType.Spinner" or "ControlType.RadioButton" } repeated } &&
                    repeated.Process?.Equals(focusClick.Process,
                        StringComparison.OrdinalIgnoreCase) == true &&
                    repeated.Window?.Equals(focusClick.Window,
                        StringComparison.OrdinalIgnoreCase) == true &&
                    repeated.ControlType == focusClick.ControlType &&
                    repeated.Name == focusClick.Name &&
                    repeated.AutomationId == focusClick.AutomationId)
                    followingIndex++;
                if (followingIndex < plan.Steps.Count && plan.Steps[followingIndex] is
                    { Action: "type", Value: not null, Target: { Process: { Length: > 0 },
                        ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } } followingField } followingType &&
                    followingField.Process?.Equals(focusClick.Process,
                        StringComparison.OrdinalIgnoreCase) == true &&
                    (focusClick.ControlType == "ControlType.Spinner" ||
                        followingField.Window?.Equals(focusClick.Window,
                            StringComparison.OrdinalIgnoreCase) == true &&
                        !NativeDialogHasButton(focusClick, followingField)))
                {
                    status($"Skipped steps {step.Number}-{plan.Steps[followingIndex - 1].Number}: " +
                        $"the recorded {focusClick.ControlType.Replace("ControlType.", "")} " +
                        $"click only focused a numeric field; step {followingType.Number} sets and " +
                        $"verifies Edit {followingField.AutomationId} as '{followingType.Value}'.");
                    index = followingIndex - 1;
                    continue;
                }
            }
            if (step.Action == "click" && step.Target is
                { ControlType: "ControlType.Button", Process: { Length: > 0 },
                    Window: { Length: > 0 }, Name: { Length: > 0 } } heading &&
                IsRecordedSectionGroup(heading))
            {
                var nextSectionStep = index + 1 < plan.Steps.Count ? plan.Steps[index + 1] : null;
                if (nextSectionStep is { Action: "click", Target:
                    { ControlType: "ControlType.ListItem", Name: { Length: > 0 } desiredItem } listTarget } &&
                    listTarget.Process?.Equals(heading.Process, StringComparison.OrdinalIgnoreCase) == true)
                {
                    var comboSelected = TrySelectSectionComboThroughUia(heading, listTarget,
                        out var comboDetail);
                    if (!comboSelected &&
                        !TrySelectUniqueNativeComboValue(heading, desiredItem, out comboDetail))
                        throw new InvalidOperationException($"Step {step.Number}: section '{heading.Name}' was followed by '{desiredItem}', but its native combo value could not be selected. {comboDetail}");
                    status($"Completed steps {step.Number}-{nextSectionStep.Number}: {comboDetail}");
                    index++;
                    continue;
                }
                if (nextSectionStep is { Action: "type", Target:
                    { ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } editId } editTarget } &&
                    editTarget.Process?.Equals(heading.Process, StringComparison.OrdinalIgnoreCase) == true &&
                    editTarget.Window?.Equals(heading.Window, StringComparison.OrdinalIgnoreCase) == true)
                {
                    status($"Skipped step {step.Number}: '{heading.Name}' is a section heading; step {nextSectionStep.Number} sets Edit {editId}.");
                    continue;
                }
            }
            if (step.Target.Window == "Find and Replace" && step.Target.ClassName == "EDTBX" &&
                step.Target.AutomationId == "21")
            {
                var dialogHandle = FindDesktopWindow(step.Target.Process!, step.Target.Window, exactOnly: true);
                if (dialogHandle == 0)
                    throw new InvalidOperationException($"Step {step.Number}: Find and Replace dialog is unavailable.");
                await EnsureReplaceTabOpenAsync(AutomationElement.FromHandle(dialogHandle), step.Number, token);
            }
            if (step.Action == "scroll")
            {
                status($"Finding scroll control at step {step.Number}");
                var scrollTarget = await FindAsync(step.Target, token, 4);
                if (scrollTarget is null && step.Target.ControlType != "ControlType.ScrollBar" && step.Key is "Vertical" or "Horizontal")
                    scrollTarget = await Task.Run(() => Automation.FindScrollBar(step.Target, step.Key!), token);
                if (scrollTarget is null) throw new InvalidOperationException($"Step {step.Number}: scroll control was not found.");
                ActivateWindow(scrollTarget);
                ApplyScroll(scrollTarget, step, status);
                status($"Completed step {step.Number}: scrolled {step.Key?.ToLowerInvariant()}");
                continue;
            }
            if (step.Action == "ensure-state" && step.Target.ControlType == "ControlType.HeaderItem" &&
                DesiredSortDirection(step.ExpectedState) is { } desiredDirection)
            {
                var header = await FindAsync(step.Target, token)
                    ?? throw new InvalidOperationException($"Step {step.Number}: sort header was not found.");
                ActivateWindow(header);
                var clicked = await EnsureSortDirectionAsync(header, step, desiredDirection, token);
                status(clicked
                    ? $"Completed step {step.Number}: clicked {step.Target.Name} and verified {desiredDirection} order"
                    : $"Completed step {step.Number}: {step.Target.Name} was already {desiredDirection}");
                continue;
            }
            if (step.Action is "select-first-row" or "select-all-items")
            {
                await EnsureApplicationWindowAsync(step.Target.Process!, step.Target.Window, status, token);
                previousProcess = step.Target.Process;
                previousWindow = step.Target.Window;
                var firstRow = await FindFirstListRowAsync(step.Target, token);
                if (firstRow is null)
                    throw new InvalidOperationException($"Step {step.Number}: '{step.Target.Window}' was activated, but no list row was exposed. The recording may be missing the navigation that opens this view.");
                ActivateWindow(firstRow);
                if (!firstRow.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var firstSelection))
                    throw new InvalidOperationException($"Step {step.Number}: the first row cannot be selected through UI Automation.");
                ((SelectionItemPattern)firstSelection).Select();
                firstRow.SetFocus();
                if (step.Action == "select-all-items")
                {
                    var list = ((SelectionItemPattern)firstSelection).Current.SelectionContainer;
                    if (list is null) throw new InvalidOperationException($"Step {step.Number}: the list exposes no selection container.");
                    list.SetFocus();
                    SendKeys.SendWait("^a");
                    if (!AllListItemsSelected(list))
                    {
                        firstRow.SetFocus();
                        SendKeys.SendWait("+{END}");
                    }
                    if (!AllListItemsSelected(list))
                    {
                        status($"Step {step.Number}: keyboard selection was incomplete; selecting list items through UI Automation.");
                        SelectAllListItemsViaUia(list, token);
                    }
                    if (!AllListItemsSelected(list))
                        throw new InvalidOperationException($"Step {step.Number}: the list did not confirm that all items were selected. {ListSelectionDiagnostics(list)}");
                }
                selectedRowsHeader = step.Target;
                status(step.Action == "select-all-items"
                    ? $"Completed step {step.Number}: selected all items under {step.Target.Name}"
                    : $"Completed step {step.Number}: selected the first row in {step.Target.Name}");
                continue;
            }
            if (step.Action == "ensure-state" && step.Target.ControlType == "ControlType.HeaderItem" &&
                step.ExpectedState?.Contains("sort", StringComparison.OrdinalIgnoreCase) == true)
                throw new InvalidOperationException($"Step {step.Number}: the recording did not capture ascending or descending sort direction. Record this workflow again with the updated app before replaying later case actions.");
            if (step.Action == "context-click" && IsExcelSheetTab(step.Target) &&
                index + 1 < plan.Steps.Count && plan.Steps[index + 1].Target?.Name == "Move or Copy...")
            {
                var workbookWindow = FindDesktopWindow(step.Target.Process!, step.Target.Window, exactOnly: true);
                if (workbookWindow != 0)
                {
                    ShowWindow(workbookWindow, 9);
                    SetForegroundWindow(workbookWindow);
                    await Task.Delay(150, token);
                }
                var sheet = await FindAsync(step.Target, token)
                    ?? throw new InvalidOperationException($"Step {step.Number}: sheet tab '{step.Target.Name}' was not found. " +
                        Automation.TabLookupDiagnostics(step.Target));
                ActivateWindow(sheet);
                SelectSheet(sheet);
                copiedSheetTarget = step.Target;
                tabsBeforeCopy = Automation.SiblingTabNames(sheet);
                SendKeys.SendWait("%h"); SendKeys.SendWait("o"); SendKeys.SendWait("m");
                if (!await WaitForWindowAsync("Move or Copy", token))
                    throw new InvalidOperationException("Excel did not open the Move or Copy dialog after the keyboard command.");
                status($"Completed steps {step.Number}-{plan.Steps[index + 1].Number}: opened Move or Copy dialog");
                index++;
                continue;
            }
            if (step.Action == "context-click" && IsExcelSheetTab(step.Target))
            {
                var sheet = await FindAsync(step.Target, token)
                    ?? throw new InvalidOperationException($"Step {step.Number}: sheet tab '{step.Target.Name}' was not found. " +
                        Automation.TabLookupDiagnostics(step.Target));
                ActivateWindow(sheet);
                if (!TryClickLiveUiaElement(sheet, null, rightClick: true, out var detail))
                    throw new InvalidOperationException($"Step {step.Number}: could not verify a right-click on sheet tab '{step.Target.Name}'. {detail}");
                status($"Completed step {step.Number}: opened the context menu on sheet tab '{step.Target.Name}'.");
                continue;
            }
            if (step.Action == "context-click" && step.Target.Process == "EXCEL" &&
                step.Target.ClassName == "XLGridRowHeader" &&
                int.TryParse(step.Target.Name, out var excelRowNumber))
            {
                var rowCellTarget = plan.Steps.Take(index).LastOrDefault(candidate =>
                    candidate.Target is { Process: "EXCEL" } target &&
                    target.AutomationId == "B" + excelRowNumber)?.Target;
                var rowCell = rowCellTarget is null ? null : await FindAsync(rowCellTarget, token, 3);
                if (rowCell?.TryGetCurrentPattern(ValuePattern.Pattern, out var rowValue) != true)
                    throw new InvalidOperationException($"Step {step.Number}: could not read row {excelRowNumber} before opening its context menu. Replay stopped before later worksheet actions.");
                pendingExcelRowCheck = rowCellTarget;
                pendingExcelRowValue = ((ValuePattern)rowValue).Current.Value;
                var followingTarget = rowCellTarget! with
                {
                    AutomationId = "B" + (excelRowNumber + 1), Name = "B" + (excelRowNumber + 1)
                };
                var followingCell = await FindAsync(followingTarget, token, 2);
                pendingExcelFollowingRowValue = followingCell?.TryGetCurrentPattern(ValuePattern.Pattern,
                    out var followingValue) == true ? ((ValuePattern)followingValue).Current.Value : null;
                var rowMenuTarget = new ControlRef("EXCEL", step.Target.Window, null,
                    "Row Height...", "ControlType.MenuItem", null, null);
                var firstColumnTarget = rowCellTarget! with
                {
                    AutomationId = "A" + excelRowNumber, Name = "A" + excelRowNumber
                };
                var firstColumnCell = await FindAsync(firstColumnTarget, token, 3) ??
                    throw new InvalidOperationException($"Step {step.Number}: could not locate the first cell in row {excelRowNumber} to find its live row header.");
                ActivateWindow(firstColumnCell);
                firstColumnCell = await FindAsync(firstColumnTarget, token, 2) ?? firstColumnCell;
                var cellBounds = firstColumnCell.Current.BoundingRectangle;
                if (cellBounds.IsEmpty || cellBounds.Width < 8 || cellBounds.Height < 8)
                    throw new InvalidOperationException($"Step {step.Number}: row {excelRowNumber}'s first cell has no usable live bounds.");
                var headerX = (int)Math.Round(cellBounds.Left - 12);
                var headerY = (int)Math.Round(cellBounds.Top + cellBounds.Height / 2);
                bool IsLiveRowHeader()
                {
                    if (!Screen.AllScreens.Any(screen => screen.Bounds.Contains(headerX, headerY))) return false;
                    try
                    {
                        var hit = AutomationElement.FromPoint(new System.Windows.Point(headerX, headerY));
                        var observed = Automation.Describe(hit);
                        return observed is { Process: "EXCEL", ClassName: "XLGridRowHeader" } &&
                            observed.Name == excelRowNumber.ToString();
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); return false; }
                }
                if (!IsLiveRowHeader())
                    throw new InvalidOperationException($"Step {step.Number}: the live point left of A{excelRowNumber} was not identified as row header {excelRowNumber}. Replay stopped before Delete.");
                if (!SetCursorPos(headerX, headerY) || !GetCursorPos(out var rowCursor) ||
                    rowCursor.X != headerX || rowCursor.Y != headerY)
                    throw new InvalidOperationException($"Step {step.Number}: Windows could not reach the verified row header.");
                mouse_event(0x0002u, 0, 0, 0, 0);
                Thread.Sleep(45);
                mouse_event(0x0004u, 0, 0, 0, 0);
                await Task.Delay(150, token);
                if (!IsLiveRowHeader())
                    throw new InvalidOperationException($"Step {step.Number}: row header {excelRowNumber} moved after selection. Replay stopped before Delete.");
                mouse_event(0x0008u, 0, 0, 0, 0);
                Thread.Sleep(45);
                mouse_event(0x0010u, 0, 0, 0, 0);
                var workbookHandle = FindDesktopWindow("EXCEL", step.Target.Window, exactOnly: true);
                var menuVisible = workbookHandle != 0 && await RowMenuPopupVisibleAsync(workbookHandle, token);
                if (!menuVisible && await FindAsync(rowMenuTarget, token, 1) is null)
                {
                    SendKeys.SendWait("{ESC}");
                    throw new InvalidOperationException($"Step {step.Number}: Excel did not expose a row context menu after right-clicking the verified row header {excelRowNumber}. Replay stopped before Delete.");
                }
                status($"Completed step {step.Number}: selected row {excelRowNumber} and verified its row context menu.");
                continue;
            }
            if (step.Action == "rename-sheet")
            {
                if (!IsExcelSheetTab(step.Target) || string.IsNullOrWhiteSpace(step.Value))
                    throw new InvalidOperationException($"Step {step.Number}: unsupported sheet rename target.");
                var followingIntent = index + 1 < plan.Steps.Count ? plan.Steps[index + 1] : null;
                var attachedWeekday = followingIntent?.Action == "key" &&
                    followingIntent.Key is "Enter" or "Return" &&
                    followingIntent.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                    followingIntent.Target.ControlType == "ControlType.Pane" &&
                    !string.IsNullOrWhiteSpace(followingIntent.Intent)
                    ? followingIntent.RelativeWeekday : null;
                var renameTemplate = DynamicText.ToTemplate(step.Value, attachedWeekday ?? step.RelativeWeekday);
                var repairedTemplate = Regex.Replace(renameTemplate,
                    @"\(\)\s*(\{\{next:[^}]+\}\})", "($1)", RegexOptions.IgnoreCase);
                if (repairedTemplate != renameTemplate)
                {
                    status($"Step {step.Number}: moved the recorded date inside the empty sheet-name parentheses.");
                    renameTemplate = repairedTemplate;
                }
                var intendedName = DynamicText.Resolve(renameTemplate, attachedWeekday ?? step.RelativeWeekday);
                var sheet = await FindAsync(step.Target, token)
                    ?? throw new InvalidOperationException($"Step {step.Number}: sheet to rename was not found.");
                ActivateWindow(sheet);
                SelectSheet(sheet);
                var sheetTabs = TreeWalker.ControlViewWalker.GetParent(sheet);
                var malformedName = MalformedDatedSheetName(renameTemplate);
                if (!malformedName)
                {
                    SendKeys.SendWait("%h"); SendKeys.SendWait("o"); SendKeys.SendWait("r");
                    SendKeys.SendWait(string.Concat(intendedName.Select(EscapeText)));
                    SendKeys.SendWait("{ENTER}");
                    PauseForUnexpectedDialog(step.Target.Process, knownWindows, handleUnexpectedDialog, status, token);
                }
                else status($"Step {step.Number}: the recorded sheet name is malformed; waiting for your name in Excel.");
                var renamed = step.Target with { Name = intendedName };
                var nameBeforeManualHelp = SelectedSheetName(sheetTabs);
                var manualHelpAccepted = false;
                while (true)
                {
                    token.ThrowIfCancellationRequested();
                    var oldNameStillPresent = await Task.Run(() => Automation.Resolve(step.Target) is not null, token);
                    var finalNamePresent = await Task.Run(() => Automation.Resolve(renamed) is not null, token);
                    if (!oldNameStillPresent && finalNamePresent) break;
                    var selectedName = SelectedSheetName(sheetTabs);
                    if (manualHelpAccepted && !oldNameStillPresent &&
                        !string.IsNullOrWhiteSpace(selectedName) && selectedName != nameBeforeManualHelp)
                    {
                        intendedName = selectedName;
                        break;
                    }
                    status($"Step {step.Number}: sheet rename needs your attention in Excel.");
                    var prompt = malformedName
                        ? "The recording did not capture a valid sheet name. Rename the copied sheet in Excel, then continue."
                        : $"Excel could not finish renaming the copied sheet to {intendedName}. Resolve any Excel message, then complete that rename.";
                    if (!handleUnexpectedDialog(prompt))
                        throw new InvalidOperationException($"Step {step.Number}: replay stopped before the sheet rename was resolved.");
                    manualHelpAccepted = true;
                }
                status($"Completed step {step.Number}: renamed sheet to {intendedName}");
                if (attachedWeekday is not null) index++; // The following Enter carries the rename intent.
                continue;
            }
            status($"Finding step {step.Number}: {step.Target.Name ?? step.Target.AutomationId ?? step.Target.ClassName}");
            if (step.Action == "click" && step.Target.Window == "Find and Replace" &&
                step.Target.Name == "Replace All")
            {
                var dialogHandle = FindDesktopWindow(step.Target.Process ?? "EXCEL", "Find and Replace", exactOnly: true);
                if (dialogHandle == 0)
                    throw new InvalidOperationException($"Step {step.Number}: Find and Replace is unavailable.");
                await EnsureReplaceTabOpenAsync(AutomationElement.FromHandle(dialogHandle), step.Number, token);
            }
            var menuIndex = index + 1;
            while (menuIndex < plan.Steps.Count && IsTaskbarNavigation(plan.Steps[menuIndex])) menuIndex++;
            if ((step.Action == "context-selection" || step.Action == "context-click" && selectedRowsHeader is not null) &&
                (selectedRowsHeader is null || step.Target.Process == selectedRowsHeader.Process) &&
                menuIndex < plan.Steps.Count && plan.Steps[menuIndex].Target?.ControlType == "ControlType.MenuItem" &&
                plan.Steps[menuIndex].Target?.Process == step.Target.Process)
            {
                var menuTarget = plan.Steps[menuIndex].Target!;
                AutomationElement? row = null;
                if (selectedRowsHeader is not null)
                {
                    row = selectedRowsHeader.ControlType == "ControlType.HeaderItem"
                        ? await FindAsync(selectedRowsHeader, token, 3) is { } header
                            ? await Task.Run(() => Automation.FirstListItemUnderHeader(header), token) : null
                        : await Task.Run(() => Automation.FirstListItemInWindow(selectedRowsHeader), token);
                }
                row ??= await Task.Run(() => Automation.FindSelectedListRow(step.Target), token);
                if (row is null) throw new InvalidOperationException($"Step {step.Number}: the selected list items are no longer available.");
                var selectedList = row.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectedPattern)
                    ? ((SelectionItemPattern)selectedPattern).Current.SelectionContainer : null;
                if (step.Action == "context-selection" && selectedList is null)
                    throw new InvalidOperationException($"Step {step.Number}: the list selection container is unavailable.");
                if (step.Action == "context-selection" && selectedList is not null && !AllListItemsSelected(selectedList))
                    throw new InvalidOperationException($"Step {step.Number}: all items are not selected. {ListSelectionDiagnostics(selectedList)}");
                ActivateWindow(row);
                if (selectedList is not null && !AllListItemsSelected(selectedList))
                    SelectAllListItemsViaUia(selectedList, token);
                if (step.Action == "context-selection" && selectedList is not null && !AllListItemsSelected(selectedList))
                    throw new InvalidOperationException($"Step {step.Number}: selection was lost before the context menu. {ListSelectionDiagnostics(selectedList)}");
                var exposedCommand = await FindAsync(menuTarget, token, 1);
                if (exposedCommand is not null &&
                    Automation.Describe(exposedCommand)?.Window == menuTarget.Window &&
                    exposedCommand.Current.IsEnabled &&
                    exposedCommand.TryGetCurrentPattern(InvokePattern.Pattern, out var exposedInvoke))
                {
                    ((InvokePattern)exposedInvoke).Invoke();
                    status($"Completed steps {step.Number}-{plan.Steps[menuIndex].Number}: invoked {menuTarget.Name} through UI Automation");
                    index = menuIndex;
                    continue;
                }
                AutomationElement? menu = null;
                var clickDetail = "selection container unavailable";
                if (selectedList is not null && TryRightClickSelectedItemAtLiveUiaBounds(selectedList, out clickDetail))
                {
                    status($"Step {step.Number}: right-clicked a selected row located through live UI Automation. {clickDetail}");
                    menu = await FindAsync(menuTarget, token, 4);
                    if (step.Action == "context-selection" && !AllListItemsSelected(selectedList))
                    {
                        if (menu is not null) SendKeys.SendWait("{ESC}");
                        throw new InvalidOperationException($"Step {step.Number}: the right-click changed the selected rows. {ListSelectionDiagnostics(selectedList)}");
                    }
                }
                else status($"Step {step.Number}: live UI Automation right-click unavailable: {clickDetail}");
                if (menu is not null)
                {
                    status($"Completed step {step.Number}: opened context menu on selected rows");
                    continue;
                }
                // Selection through UIA does not necessarily leave keyboard focus on the list.
                // Focus the container, then verify that its selection survived before sending keys.
                if (selectedList is not null)
                {
                    selectedList.SetFocus();
                    if (step.Action == "context-selection" && !AllListItemsSelected(selectedList))
                        throw new InvalidOperationException($"Step {step.Number}: focusing the list changed its selection. {ListSelectionDiagnostics(selectedList)}");
                    if (selectedList.Current.ClassName.Contains("SysListView32", StringComparison.OrdinalIgnoreCase) &&
                        selectedList.Current.NativeWindowHandle is var listHandle and not 0 &&
                        (int)SendMessage((nint)listHandle, 0x100C, -1, 1) < 0) // LVM_GETNEXTITEM, LVNI_FOCUSED
                    {
                        SendKeys.SendWait("^{HOME}"); // Give the list an item anchor without selecting by row text.
                        if (step.Action == "context-selection" && !AllListItemsSelected(selectedList))
                            SelectAllListItemsViaUia(selectedList, token);
                        if (step.Action == "context-selection" && !AllListItemsSelected(selectedList))
                            throw new InvalidOperationException($"Step {step.Number}: the list lost its selection while setting a keyboard anchor.");
                    }
                }
                status($"Step {step.Number}: opening selected rows' menu with keyboard focus on the list. {ContextFocusDiagnostics(selectedList)}");
                OpenContextMenu(selectedList ?? row, preserveFocus: true);
                menu = await FindAsync(menuTarget, token, 3);
                if (menu is null)
                {
                    SendContextMenuKey(selectedList ?? row, preserveFocus: true);
                    menu = await FindAsync(menuTarget, token, 3);
                }
                if (menu is null)
                {
                    SendWindowContextMenu(selectedList ?? row); // Keyboard sentinel; no pointer coordinates.
                    menu = await FindAsync(menuTarget, token, 3);
                }
                if (menu is null && selectedList is not null)
                {
                    menu = await FindAsync(menuTarget, token, 2);
                }
                if (menu is null && SendFocusedContextMenu(step.Target.Process))
                    menu = await FindAsync(menuTarget, token, 3);
                if (menu is null)
                {
                    throw new InvalidOperationException($"Step {step.Number}: the app did not expose a keyboard copy action or context menu for the selected rows. " +
                        (selectedList is null ? "Selection container unavailable." : ListSelectionDiagnostics(selectedList)));
                }
                status($"Completed step {step.Number}: opened context menu on selected rows");
                continue;
            }
            if (step.Action is "type" or "key" && step.Target.ControlType == "ControlType.Edit" &&
                FollowsSubmittedEdit(plan.Steps, index, out var submittedField))
            {
                AutomationElement? liveEditor = null;
                ControlRef? liveEditorTarget = null;
                for (var focusAttempt = 0; focusAttempt < 100; focusAttempt++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        var focused = AutomationElement.FocusedElement;
                        var described = focused is null ? null : Automation.Describe(focused);
                        if (focused is not null && described is not null &&
                            described.Process?.Equals(step.Target.Process,
                                StringComparison.OrdinalIgnoreCase) == true &&
                            !focused.Current.IsPassword && focused.Current.IsEnabled &&
                            !focused.Current.IsOffscreen &&
                            (focused.Current.ControlType == ControlType.Edit ||
                             focused.Current.ControlType == ControlType.ComboBox ||
                             focused.TryGetCurrentPattern(ValuePattern.Pattern, out _)) &&
                            !(submittedField.ClassName?.Contains("Omnibox",
                                StringComparison.OrdinalIgnoreCase) == true &&
                              described.AutomationId == submittedField.AutomationId &&
                              described.ClassName == submittedField.ClassName))
                        {
                            liveEditor = focused;
                            liveEditorTarget = described;
                            break;
                        }
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or
                        InvalidOperationException or COMException) { Trace.WriteLine(ex); }
                    await Task.Delay(100, token);
                }
                if (liveEditor is null || liveEditorTarget is null)
                    throw new InvalidOperationException($"Step {step.Number}: Enter changed the editable " +
                        "page, but replay could not verify the newly focused field. Input was not sent.");
                var focusedStep = step with { Target = liveEditorTarget };
                if (step.Action == "type")
                {
                    if (!TryClickLiveUiaElement(liveEditor, null, rightClick: false,
                        out var focusDetail))
                        throw new InvalidOperationException($"Step {step.Number}: the newly focused " +
                            $"field could not be confirmed at its live position. {focusDetail}");
                    SetRecordedEditField(liveEditor, focusedStep);
                    status($"Completed step {step.Number}: entered and verified the recorded text " +
                        "in the field focused after navigation.");
                }
                else
                {
                    FocusAndVerifyKeyboardTarget(liveEditor, focusedStep);
                    SendKeys.SendWait(ToSendKeys(step.Key ??
                        throw new InvalidDataException($"Step {step.Number}: recorded key is missing.")));
                    status($"Completed step {step.Number}: pressed {step.Key} in the field " +
                        "focused after navigation.");
                }
                continue;
            }
            if (step.Action is "type" or "key" &&
                step.Target.ControlType == "ControlType.Pane" &&
                string.IsNullOrWhiteSpace(step.Target.Name) &&
                string.IsNullOrWhiteSpace(step.Target.AutomationId) &&
                !string.IsNullOrWhiteSpace(step.Target.Process) &&
                FollowsSubmittedNavigation(plan.Steps, index))
            {
                AutomationElement? focused = null;
                ControlRef? focusedTarget = null;
                for (var focusAttempt = 0; focusAttempt < 100; focusAttempt++)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        var candidate = AutomationElement.FocusedElement;
                        var described = candidate is null ? null : Automation.Describe(candidate);
                        var liveTitle = GetNativeWindowTitle(GetForegroundWindow());
                        if (candidate is not null && described is not null &&
                            described.Process?.Equals(step.Target.Process,
                                StringComparison.OrdinalIgnoreCase) == true &&
                            described.Window?.Equals(liveTitle,
                                StringComparison.OrdinalIgnoreCase) == true &&
                            BrowserPageTitleMatches(step.Target.Window, liveTitle) &&
                            described.ClassName?.Contains("Omnibox",
                                StringComparison.OrdinalIgnoreCase) != true &&
                            !candidate.Current.IsPassword && candidate.Current.IsEnabled &&
                            !candidate.Current.IsOffscreen)
                        {
                            focused = candidate;
                            focusedTarget = described;
                            break;
                        }
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or
                        InvalidOperationException or COMException) { Trace.WriteLine(ex); }
                    await Task.Delay(100, token);
                }
                if (focused is null || focusedTarget is null)
                    throw new InvalidOperationException($"Step {step.Number}: the recorded keyboard " +
                        "pane has no stable identity and a focused page target was not confirmed " +
                        $"after navigation. Foreground window: '{GetNativeWindowTitle(GetForegroundWindow())}'. " +
                        "Input was not sent.");
                Act(focused, step, preserveFocus: true);
                status($"Completed step {step.Number}: replayed the recorded keyboard action " +
                    "in the focused application window.");
                continue;
            }
            AutomationElement? element = null;
            if ((step.Action == "key" && step.Target.ControlType == "ControlType.ListItem") ||
                (step.Action == "context-click" && step.Target.ControlType == "ControlType.Text" &&
                 index > 0 && plan.Steps[index - 1].Key?.Contains("Shift+End", StringComparison.OrdinalIgnoreCase) == true))
            {
                var focused = AutomationElement.FocusedElement;
                var focusedRef = Automation.Describe(focused);
                if (focused is not null && focusedRef is not null && focusedRef.Process == step.Target.Process &&
                    focusedRef.Window == step.Target.Window)
                    element = Automation.ListItemAncestor(focused);
                if (element is null)
                    throw new InvalidOperationException($"Step {step.Number}: the selected list row lost focus; the action cannot be sent safely.");
            }
            else element = await FindAsync(step.Target, token,
                step.Target.ControlType == "ControlType.MenuItem" ? 1 :
                FollowsHorizontalScroll(plan, index) ? 2 : 20);
            if (element is null && FollowsHorizontalScroll(plan, index))
            {
                var reveal = await RevealTargetWithUiaScrollAsync(step, status, handleUnexpectedDialog, token);
                if (reveal.HandledByUser)
                {
                    var observed = Automation.Describe(AutomationElement.FocusedElement);
                    status($"Completed step {step.Number}: user handled the target after UI Automation scrolling failed. " +
                        $"Focused control afterward: {observed?.Name ?? "unknown"} ({observed?.ControlType ?? "unknown type"}).");
                    continue;
                }
                element = reveal.Element;
            }
            if (element is null && step.Target.ControlType != "ControlType.MenuItem" &&
                !string.IsNullOrWhiteSpace(step.Target.Process))
            {
                status($"Step {step.Number}: checking that the target program is enlarged before searching again.");
                await EnsureApplicationWindowAsync(step.Target.Process, step.Target.Window, status, token);
                element = await FindAsync(step.Target, token, 5);
            }
            if (element is null && step.Target.ControlType == "ControlType.MenuItem" &&
                !string.IsNullOrWhiteSpace(step.Target.ParentName))
            {
                var parentName = step.Target.ParentName;
                var parentTarget = step.Target with
                {
                    Name = parentName, ParentName = null, AutomationId = null, ClassName = null
                };
                var parentMenuItem = await FindAsync(parentTarget, token, 3);
                if (parentMenuItem is not null)
                {
                    status($"Opening recorded parent menu '{parentName}' for step {step.Number}.");
                    if (parentMenuItem.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var expand))
                    {
                        var pattern = (ExpandCollapsePattern)expand;
                        if (pattern.Current.ExpandCollapseState == ExpandCollapseState.Collapsed)
                            pattern.Expand();
                        else if (pattern.Current.ExpandCollapseState == ExpandCollapseState.LeafNode &&
                                 !TryClickLiveUiaElement(parentMenuItem, null, rightClick: false, out var leafDetail))
                            throw new InvalidOperationException($"Step {step.Number}: parent menu '{parentName}' could not be opened. {leafDetail}");
                    }
                    else if (!TryClickLiveUiaElement(parentMenuItem, null, rightClick: false, out var parentDetail))
                        throw new InvalidOperationException($"Step {step.Number}: parent menu '{parentName}' could not be opened. {parentDetail}");
                    element = await FindAsync(step.Target, token, 5);
                }
                if (element is null)
                    throw new InvalidOperationException($"Step {step.Number}: menu item '{step.Target.Name}' was not found under recorded parent '{parentName}'.");
            }
            if (element is null && step.Target.ControlType == "ControlType.MenuItem" && index > 0)
            {
                var previous = plan.Steps[index - 1].Target;
                if (previous is not null)
                {
                    status("Opening context menu from the current selection");
                    var focused = AutomationElement.FocusedElement;
                    var anchor = focused is not null && Automation.Describe(focused)?.Process == previous.Process
                        ? Automation.ListItemAncestor(focused) : null;
                    anchor ??= await FindAsync(previous, token, 2);
                    if (anchor is not null)
                    {
                        ActivateWindow(anchor);
                        try { OpenContextMenu(anchor); }
                        catch (ElementNotAvailableException)
                        {
                            anchor = CurrentSelectedRow(previous.Process);
                            if (anchor is not null) OpenContextMenu(anchor);
                        }
                        element = await FindAsync(step.Target, token, 2);
                        if (element is null)
                        {
                            anchor = CurrentSelectedRow(previous.Process) ?? anchor;
                            try { if (anchor is not null) SendWindowContextMenu(anchor); }
                            catch (ElementNotAvailableException)
                            {
                                anchor = CurrentSelectedRow(previous.Process);
                                if (anchor is not null) SendWindowContextMenu(anchor);
                            }
                            element = await FindAsync(step.Target, token, 2);
                        }
                    }
                }
            }
            if (element is null && step.Action == "click" &&
                step.Target.ControlType is "ControlType.Button" or "ControlType.RadioButton" or "ControlType.CheckBox" &&
                !string.IsNullOrWhiteSpace(step.Target.Process) &&
                !string.IsNullOrWhiteSpace(step.Target.Window) &&
                !string.IsNullOrWhiteSpace(step.Target.Name))
            {
                var nativeWindows = FindNativeWindows(step.Target.Process, step.Target.Window);
                if (nativeWindows.Count == 1)
                {
                    var legacyWindow = nativeWindows[0].Dialog;
                    ShowWindow(nativeWindows[0].Root, 9);
                    SetForegroundWindow(nativeWindows[0].Root);
                    var roles = step.Target.ControlType switch
                    {
                        "ControlType.RadioButton" => new[] { 0x2D },
                        "ControlType.CheckBox" => new[] { 0x2C },
                        _ => new[] { 0x2B, 0x2D, 0x2C }
                    };
                    if (roles.Any(role => MsaaActions.TryInvokeUnique(legacyWindow, step.Target.Name, role)) ||
                        TryClickUniqueNativeButton(legacyWindow, step.Target.Name, step.Target.ControlType))
                    {
                        if (NormalizeNativeCaption(step.Target.Name).Equals("Save Close",
                            StringComparison.OrdinalIgnoreCase) && index + 1 < plan.Steps.Count &&
                            plan.Steps[index + 1].Target?.Window != step.Target.Window)
                        {
                            for (var wait = 0; wait < 20 && IsWindowVisible(legacyWindow); wait++)
                                await Task.Delay(100, token);
                            if (IsWindowVisible(legacyWindow))
                                throw new InvalidOperationException($"Step {step.Number}: '{step.Target.Name}' was invoked, but '{step.Target.Window}' remained open.");
                        }
                        else await Task.Delay(150, token);
                        status($"Completed step {step.Number}: invoked unique legacy control '{step.Target.Name}'.");
                        continue;
                    }
                }
            }
            if (element is null && step.Action == "type" &&
                step.Target is { ControlType: "ControlType.Edit", ClassName: "Edit",
                    AutomationId: { Length: > 0 } } nativeEdit && step.Value is not null)
            {
                if (TrySetUniqueNativeEdit(nativeEdit, step.Value, out var nativeDetail))
                {
                    status($"Completed step {step.Number}: entered and verified the recorded value in native Edit {nativeEdit.AutomationId}.");
                    continue;
                }
                throw new InvalidOperationException($"Step {step.Number}: native Edit {nativeEdit.AutomationId} could not be set: {nativeDetail}");
            }
            if (element is null)
                throw new InvalidOperationException($"Step {step.Number}: control '{step.Target.Name}' was not found through UI Automation or as a unique legacy control in '{step.Target.Window}'.");
            if (step.Action == "click" && step.Target is
                { ControlType: "ControlType.MenuItem", AutomationId: "Dropdown", ParentName: { Length: > 0 } })
            {
                if (!TryClickLiveUiaElement(element, null, rightClick: false, out var filterDetail))
                    throw new InvalidOperationException($"Step {step.Number}: the recorded column filter was not hit-testable. {filterDetail}");
                status($"Completed step {step.Number}: opened filter under {step.Target.ParentName}. {filterDetail}");
                continue;
            }
            // Excel's filter list is a transient popup. Activating its workbook window
            // dismisses the popup and leaves its UIA tree items disabled.
            var filterListItem = step.Action == "click" &&
                step.Target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ControlType == "ControlType.TreeItem";
            var filterOk = step.Action == "click" &&
                step.Target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
                step.Target.ControlType == "ControlType.Button" && step.Target.Name == "OK" &&
                index > 0 && plan.Steps[index - 1].Target?.ControlType == "ControlType.TreeItem";
            if (filterListItem || filterOk)
            {
                var detail = await ClickTransientFilterControlAsync(plan, index, element,
                    status, handleUnexpectedDialog, token);
                status($"Completed step {step.Number}: clicked filter control {step.Target.Name}. {detail}");
                continue;
            }
            if (step.Target.ControlType != "ControlType.MenuItem" &&
                !(step.Target.Window == "Find and Replace" && step.Target.ClassName == "EDTBX" &&
                  step.Action is "click" or "type")) ActivateWindow(element);
            if (step.Action == "click" && step.Target.ControlType == "ControlType.TitleBar" &&
                step.Target.Window == "Find and Replace")
            {
                var dialog = WindowAncestor(element) ??
                    throw new InvalidOperationException($"Step {step.Number}: Find and Replace dialog is unavailable.");
                await EnsureReplaceTabOpenAsync(dialog, step.Number, token);
                status($"Completed step {step.Number}: selected the Replace tab through UI Automation");
                continue;
            }
            if (step.Target.Window == "Find and Replace" && step.Target.ClassName == "EDTBX" &&
                step.Target.AutomationId == "21")
            {
                var dialog = WindowAncestor(element) ??
                    throw new InvalidOperationException($"Step {step.Number}: Find and Replace dialog is unavailable.");
                await EnsureReplaceTabOpenAsync(dialog, step.Number, token);
            }
            if (step.Action == "type" && step.Target.Window == "Find and Replace" &&
                step.Target.ClassName == "EDTBX" && !string.IsNullOrWhiteSpace(step.Target.Name))
            {
                SetRecordedEditField(element, step);
                status($"Completed step {step.Number}: entered the recorded text in " +
                    (step.Target.Label ?? (step.Target.AutomationId == "21" ? "Replace with" : "Find what")));
                continue;
            }
            if (step.Action == "type" && step.Target.ControlType == "ControlType.Edit" &&
                !string.IsNullOrWhiteSpace(step.Target.AutomationId))
            {
                SetRecordedEditField(element, step);
                status($"Completed step {step.Number}: entered and verified the recorded value in " +
                    (step.Target.Label ?? step.Target.AutomationId));
                continue;
            }
            if (step.Action == "click" && step.Target.Window == "Find and Replace" &&
                step.Target.ClassName == "EDTBX")
            {
                if (!TryClickLiveUiaElement(element, null, rightClick: false, out var editClickDetail))
                    throw new InvalidOperationException($"Step {step.Number}: the edit field could not be clicked through live UI Automation. {editClickDetail}");
                status($"Completed step {step.Number}: clicked the live Find and Replace field");
                continue;
            }
            if (step.Action == "click" && step.Target.ControlType == "ControlType.DataItem")
            {
                var anchor = PreviousDataItemAnchor(plan, index, step.Target);
                var clickDetail = await ClickDataItemWithRetriesAsync(step, element, anchor,
                    status, handleUnexpectedDialog, token);
                status(clickDetail.StartsWith("Handled by the user", StringComparison.Ordinal)
                    ? $"Completed step {step.Number}: {clickDetail}"
                    : $"Completed step {step.Number}: left-clicked its live UIA target. {clickDetail}");
                continue;
            }
            if (step.Action == "context-click" && step.Target.ControlType == "ControlType.DataItem" &&
                index + 1 < plan.Steps.Count && plan.Steps[index + 1].Target?.ControlType == "ControlType.MenuItem")
            {
                var precedingCell = PreviousDataItemAnchor(plan, index, step.Target);
                if (!TryClickLiveUiaElement(element, precedingCell, rightClick: true, out var contextDetail))
                {
                    // Window managers and application navigation can change the
                    // viewport after UIA found the element. Maximize the live app
                    // window, reacquire both elements, and retry once.
                    status($"Step {step.Number}: target was outside its visible UIA position; maximizing the app and locating it again.");
                    MaximizeTargetWindow(element);
                    await Task.Delay(200, token);
                    element = await FindAsync(step.Target, token, 5) ??
                        throw new InvalidOperationException($"Step {step.Number}: the context target disappeared after maximizing its window.");
                    if (!TryClickLiveUiaElement(element, precedingCell, rightClick: true, out contextDetail))
                        throw new InvalidOperationException($"Step {step.Number}: the context target could not be verified for a right-click after maximizing its window. {contextDetail}");
                }
                var expectedMenu = await FindAsync(plan.Steps[index + 1].Target!, token, 4);
                if (expectedMenu is null)
                {
                    // Some controls use the first right-click to give the row
                    // focus/selection. Reacquire the live target and try once
                    // more, then require the recorded menu item to be visible.
                    status($"Step {step.Number}: row became selected but its menu did not open; retrying the live UIA right-click.");
                    element = await FindAsync(step.Target, token, 4) ??
                        throw new InvalidOperationException($"Step {step.Number}: the context target disappeared after the first right-click.");
                    if (!TryClickLiveUiaElement(element, precedingCell, rightClick: true, out contextDetail))
                        throw new InvalidOperationException($"Step {step.Number}: the context target was no longer hit-testable after the first right-click. {contextDetail}");
                    expectedMenu = await FindAsync(plan.Steps[index + 1].Target!, token, 6);
                    if (expectedMenu is null)
                    {
                        // The pointer event may select the row without opening
                        // its menu. Ask the selected control for its keyboard
                        // context menu before concluding that it cannot act.
                        status($"Step {step.Number}: selected row is visible; trying its keyboard context menu.");
                        SendKeys.SendWait("+{F10}");
                        expectedMenu = await FindAsync(plan.Steps[index + 1].Target!, token, 6);
                        if (expectedMenu is null)
                            throw new InvalidOperationException($"Step {step.Number}: the expected context menu did not open after live UIA right-clicks or the keyboard context-menu command.");
                    }
                }
                status($"Completed step {step.Number}: opened context menu from its live UIA target. {contextDetail}");
                continue;
            }
            var before = Automation.State(element);
            if (step.Action == "click" && step.Target.Name == "Delete" &&
                step.Target.ControlType == "ControlType.MenuItem" && index > 0 &&
                plan.Steps[index - 1].Action == "context-click" &&
                plan.Steps[index - 1].Target?.ClassName == "XLGridRowHeader")
            {
                var rowNumber = plan.Steps[index - 1].Target!.Name;
                var anchor = plan.Steps.Take(index).LastOrDefault(candidate =>
                    candidate.Target?.Process == "EXCEL" && candidate.Target.AutomationId == "B" + rowNumber)?.Target;
                var beforeCell = anchor is null ? null : Automation.Resolve(anchor);
                if (beforeCell is null || !beforeCell.TryGetCurrentPattern(ValuePattern.Pattern, out var beforePattern))
                    throw new InvalidOperationException($"Step {step.Number}: could not read row {rowNumber} before deletion. The menu was left open.");
                var oldValue = ((ValuePattern)beforePattern).Current.Value;
                if (element.TryGetCurrentPattern(InvokePattern.Pattern, out var deleteInvoke))
                    ((InvokePattern)deleteInvoke).Invoke();
                else if (!TryClickLiveUiaElement(element, null, rightClick: false, out var detail))
                    throw new InvalidOperationException($"Step {step.Number}: the row Delete command could not be verified. {detail}");
                var changed = false;
                for (var attempt = 0; attempt < 20; attempt++)
                {
                    await Task.Delay(100, token);
                    var currentCell = Automation.Resolve(anchor!);
                    if (currentCell?.TryGetCurrentPattern(ValuePattern.Pattern, out var currentPattern) == true &&
                        ((ValuePattern)currentPattern).Current.Value != oldValue)
                    { changed = true; break; }
                }
                if (!changed)
                    throw new InvalidOperationException($"Step {step.Number}: Delete was invoked, but row {rowNumber} did not show a changed value in column B. Replay stopped before Select All.");
                status($"Completed step {step.Number}: deleted row {rowNumber} and verified its former column B value changed.");
                continue;
            }
            if (step.Action == "ensure-state" && StateMatches(element, step.ExpectedState)) continue;
            if (step.Action == "click" && step.Target.ControlType == "ControlType.ComboBox")
            {
                if (!OpenRecordedComboBox(element, step))
                    throw new InvalidOperationException($"Step {step.Number}: combo box '{step.Target.Name}' could not be opened through UI Automation or its live native handle.");
                status($"Completed step {step.Number}: opened combo box '{step.Target.Name}'.");
                continue;
            }
            var refresh = step.Action == "click" && IsRefreshControl(step.Target);
            if (refresh)
            {
                MovePointerAwayFrom(element);
                await Task.Delay(150, token); // Let the button's hover rendering settle.
            }
            var refreshWatch = refresh ? WatchRefreshAsync(element, step.Target, token) : null;
            if (refresh) await Task.Delay(50, token); // Let the watcher start before invoking the control.
            var tablePaste = step.Action == "key" && step.Key is "Control+V" or "Ctrl+V" &&
                step.Target.ControlType == "ControlType.DataItem" ? TableClipboardLength() : 0;
            var pasteSource = tablePaste > 0 ? Clipboard.GetText(TextDataFormat.UnicodeText).TrimStart() : null;
            var preserveFocus = step.Action is "key" or "context-click" && index > 0 &&
                plan.Steps[index - 1].Action == "click" && plan.Steps[index - 1].Target == step.Target;
            if (step.Action is "key" or "type" or "context-click") Act(element, step, preserveFocus);
            else
            {
                try { await Task.Run(() => Act(element, step), token); }
                catch (COMException ex) when (step.Action == "click")
                {
                    Trace.WriteLine(ex);
                    var live = await FindAsync(step.Target, token, 3);
                    if (live is null)
                        throw new InvalidOperationException($"Step {step.Number}: '{step.Target.Name}' " +
                            "became unavailable after its UI Automation action failed.", ex);
                    ActivateWindow(live);
                    live = await FindAsync(step.Target, token, 2) ?? live;
                    if (!TryClickLiveUiaElement(live, null, rightClick: false, out var detail))
                        throw new InvalidOperationException($"Step {step.Number}: UI Automation could not " +
                            $"invoke '{step.Target.Name}', and its live position could not be clicked. " +
                            detail, ex);
                    status($"Completed step {step.Number}: UI Automation invocation failed; " +
                        $"clicked the reacquired live target. {detail}");
                    await Task.Delay(150, token);
                }
            }
            if (tablePaste > 0)
            {
                await Task.Delay(150, token);
                var cell = Automation.Resolve(step.Target);
                if (cell is not null && cell.TryGetCurrentPattern(ValuePattern.Pattern, out var valuePattern))
                {
                    var pastedValue = ((ValuePattern)valuePattern).Current.Value;
                    if (pastedValue.Length > tablePaste * 0.6)
                        throw new InvalidOperationException($"Step {step.Number}: the table paste appears to have gone into one cell. Undo that paste in the target app before replaying again.");
                    if (pastedValue.Length is > 0 and <= 2 &&
                        pasteSource?.StartsWith(pastedValue, StringComparison.OrdinalIgnoreCase) == false)
                        throw new InvalidOperationException($"Step {step.Number}: the paste did not match the structured clipboard. Replay stopped before later edits.");
                }
            }
            if (step.Action == "click" && step.Target.ControlType == "ControlType.Text" &&
                Automation.ListItemAncestor(element) is { } selectedRow)
                selectedRow.SetFocus();
            if (refresh)
            {
                SendKeys.SendWait("{TAB}"); // Move focus away so refresh state can be observed.
                MovePointerAwayFrom(element);
                status($"Waiting for {step.Target.Name} to return to its prior state");
                var refreshResult = await refreshWatch!;
                if (refreshResult is null)
                {
                    status($"Could not verify that {step.Target.Name} finished refreshing.");
                    if (!handleUnexpectedDialog($"Refresh completion could not be verified for {step.Target.Name}. Wait until the app is ready, then continue."))
                        throw new InvalidOperationException($"Step {step.Number}: refresh was not confirmed.");
                }
            }
            else if (step.Action == "click" && step.Target.Window == "Find and Replace" &&
                step.Target.Name == "Replace All")
            {
                var nextRecordsAcknowledgement = index + 1 < plan.Steps.Count &&
                    plan.Steps[index + 1].Target?.Window == "Microsoft Excel" &&
                    plan.Steps[index + 1].Target?.Name == "OK";
                var result = await DismissReplaceAllResultAsync(token, !nextRecordsAcknowledgement);
                status($"Completed step {step.Number}: Replace All returned: {result}");
            }
            else if (step.Action == "click") await WaitForStableAsync(step.Target, before, token);
            PauseForUnexpectedDialog(step.Target.Process, knownWindows, handleUnexpectedDialog, status, token);
            if (copiedSheetTarget is not null && tabsBeforeCopy is not null &&
                step.Target.Window == "Move or Copy" && step.Target.Name == "OK")
            {
                var created = await WaitForNewTabAsync(copiedSheetTarget, tabsBeforeCopy, token);
                var recorded = plan.Steps.Skip(index + 1)
                    .FirstOrDefault(s => s.Target is not null && IsExcelSheetTab(s.Target))?.Target?.Name;
                if (recorded is not null && !recorded.Equals(created, StringComparison.OrdinalIgnoreCase))
                {
                    plan = plan with { Steps = plan.Steps.Select((s, i) => i > index && s.Target is not null &&
                        IsExcelSheetTab(s.Target) && s.Target.Name == recorded
                        ? s with { Target = s.Target with { Name = created } } : s).ToList() };
                    status($"New sheet is named {created}; following it instead of recorded name {recorded}");
                }
                copiedSheetTarget = null;
                tabsBeforeCopy = null;
            }
            if (step.Action == "ensure-state")
            {
                var observed = Automation.Resolve(step.Target);
                if (!StateMatches(observed, step.ExpectedState))
                    throw new InvalidOperationException($"Step {step.Number}: expected state '{step.ExpectedState}' was not observed. UIA reports: {Automation.State(observed) ?? "control unavailable"}.");
            }
            status($"Completed step {step.Number}");
        }
        if (replacementDialogOpen)
        {
            await CloseCompletedReplacementDialogAsync(token);
            replacementDialogOpen = false;
            status("Closed Find and Replace after the final replacement.");
        }
        var incompleteReplacements = CompletePendingReplaceAll(plan, status, token);
        return incompleteReplacements;
        }
        finally { if (tuckedAwayDialog != 0) ShowWindow(tuckedAwayDialog, 9); }
    }

    private static string GetNativeWindowTitle(nint handle)
    {
        var title = new StringBuilder(512);
        GetWindowText(handle, title, title.Capacity);
        return title.ToString();
    }

    private static bool CompletePendingReplaceAll(ExecutionPlan plan, Action<string> status, CancellationToken token)
    {
        var last = plan.Steps.LastOrDefault();
        if (last?.Target?.ClassName != "EDTBX" || last.Target.Window != "Find and Replace" ||
            plan.Steps.Any(step => step.Action == "replace-all" ||
                step.Action == "click" && step.Target?.Name == "Replace All")) return false;
        token.ThrowIfCancellationRequested();
        var fields = plan.Steps.Where(step => step.Action == "type" &&
                step.Target?.Window == last.Target.Window && step.Target.ClassName == "EDTBX" &&
                !string.IsNullOrWhiteSpace(step.Target.AutomationId))
            .GroupBy(step => step.Target!.AutomationId!).Select(group => group.Last().Target!).ToArray();
        if (fields.Length != 2 || fields.Any(field => string.IsNullOrWhiteSpace(field.Name)) ||
            !fields.Any(field => field.AutomationId == "18") ||
            !fields.Any(field => field.AutomationId == "21")) return false;
        var replacementValues = plan.Steps.Where(step => step.Action == "type" &&
                step.Target?.Window == "Find and Replace" && step.Target.AutomationId == "21")
            .Select(step => step.Target!.Name).Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        if (replacementValues.Length > 1)
            throw new InvalidOperationException("This plan contains multiple replacements but no recorded Replace All actions. No replacement was applied; record the button clicks with the updated recorder.");
        var find = fields.First(field => field.AutomationId == "18").Name;
        var replacement = fields.First(field => field.AutomationId == "21").Name;
        if (string.Equals(find, replacement, StringComparison.Ordinal))
            throw new InvalidOperationException("The recording gives the same final Find and Replace with values. Replace All was not clicked; record the corrected field values and button action.");
        var handle = FindDesktopWindow(last.Target.Process!, last.Target.Window, exactOnly: true);
        if (handle == 0) throw new InvalidOperationException("The recorded replacement dialog is no longer available.");
        ShowWindow(handle, 9);
        SetForegroundWindow(handle);
        var dialog = AutomationElement.FromHandle(handle);
        foreach (var field in fields)
        {
            var live = dialog.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, field.AutomationId));
            if (live is null || !RecordedEditFieldMatches(live, field.Name!))
                throw new InvalidOperationException($"Replace All was not clicked because the {field.AutomationId} field does not match its recorded value.");
        }
        var button = dialog.FindFirst(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
            new PropertyCondition(AutomationElement.NameProperty, "Replace All")));
        if (button is null || !button.Current.IsEnabled || button.Current.IsOffscreen)
        {
            if (!MsaaActions.TryInvokeUnique(handle, "Replace All", 0x2B))
                throw new InvalidOperationException("Replace All is not available through UI Automation or MSAA in the recorded dialog.");
        }
        else if (button.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
            ((InvokePattern)invoke).Invoke();
        else if (!TryClickLiveUiaElement(button, null, rightClick: false, out var detail))
            throw new InvalidOperationException("Replace All could not be clicked through live UI Automation. " + detail);
        status($"Clicked Replace All for '{find}' → '{replacement}'.");
        var mappings = plan.Steps.Where(step => step.Action == "type" &&
                step.Target?.Window == last.Target.Window && step.Target.ClassName == "EDTBX" &&
                step.Target.AutomationId is "18" or "21")
            .Select(step => (step.Target!.AutomationId, step.Target.Name)).ToArray();
        var earlierReplacementValues = mappings.Where(pair => pair.AutomationId == "21")
            .Select(pair => pair.Name).Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase).Count();
        if (earlierReplacementValues <= 1) return false;
        status("Warning: this plan contains earlier replacement values without recorded Replace All clicks. Only the final mapping was applied.");
        return true;
    }

    private static async Task<string> ApplyRecordedReplacementAsync(PlanStep step, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(step.Value) || step.ExpectedState is null)
            throw new InvalidOperationException($"Step {step.Number}: the replacement mapping is incomplete.");
        var handle = FindDesktopWindow(step.Target?.Process ?? "EXCEL", "Find and Replace", exactOnly: true);
        if (handle == 0)
            throw new InvalidOperationException($"Step {step.Number}: Find and Replace is unavailable.");
        ShowWindow(handle, 9);
        SetForegroundWindow(handle);
        var dialog = AutomationElement.FromHandle(handle);
        await EnsureReplaceTabOpenAsync(dialog, step.Number, token);
        foreach (var (id, value) in new[] { ("18", step.Value), ("21", step.ExpectedState) })
        {
            var field = dialog.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            if (field is null)
                throw new InvalidOperationException($"Step {step.Number}: replacement field {id} is unavailable.");
            var fieldTarget = new ControlRef(step.Target?.Process, "Find and Replace", id, value,
                field.Current.ControlType.ProgrammaticName, field.Current.ClassName, "Find and Replace");
            SetRecordedEditField(field, step with { Action = "type", Target = fieldTarget,
                Value = value, ExpectedState = "field-value:" + value });
            if (!RecordedEditFieldMatches(field, value))
                throw new InvalidOperationException($"Step {step.Number}: replacement field {id} did not retain its recorded value.");
        }
        var button = dialog.FindFirst(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
            new PropertyCondition(AutomationElement.NameProperty, "Replace All")));
        if (button is null || !button.Current.IsEnabled)
        {
            if (!MsaaActions.TryInvokeUnique(handle, "Replace All", 0x2B))
                throw new InvalidOperationException($"Step {step.Number}: Replace All is unavailable through UI Automation or MSAA.");
        }
        else if (button.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
            ((InvokePattern)invoke).Invoke();
        else if (!TryClickLiveUiaElement(button, null, rightClick: false, out var detail))
            throw new InvalidOperationException($"Step {step.Number}: Replace All could not be clicked. {detail}");
        var result = await DismissReplaceAllResultAsync(token, dismiss: true);
        return $"replaced '{step.Value}' with '{step.ExpectedState}'. {result}";
    }

    private static Task CloseCompletedReplacementDialogAsync(CancellationToken token) =>
        CloseRecordedWindowAsync(new ControlRef("EXCEL", "Find and Replace", null,
            "Find and Replace", "ControlType.Window", null, null), token);

    private static async Task CloseRecordedWindowAsync(ControlRef target, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(target.Process) || string.IsNullOrWhiteSpace(target.Window))
            throw new InvalidOperationException("The recorded window identity is incomplete.");
        var handle = target.Window == "Find and Replace"
            ? FindVisibleNativeWindow(target.Process, target.Window)
            : FindDesktopWindow(target.Process, target.Window, exactOnly: true);
        if (handle == 0 && target.Window == "Find and Replace")
            handle = FindDesktopWindow(target.Process, target.Window, exactOnly: true);
        if (handle == 0 || !IsWindowVisible(handle)) return;
        var dialog = AutomationElement.FromHandle(handle);
        var close = dialog.FindFirst(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
            new PropertyCondition(AutomationElement.NameProperty, "Close")));
        if (close is not null && close.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
            ((InvokePattern)invoke).Invoke();
        else
        {
            SetForegroundWindow(handle);
            SendKeys.SendWait("{ESC}");
        }
        for (var attempt = 0; attempt < 20 && IsWindowVisible(handle); attempt++)
            await Task.Delay(100, token);
        if (IsWindowVisible(handle))
            throw new InvalidOperationException($"{target.Window} stayed open after its recorded close action.");
        if (target.Window == "Find and Replace")
        {
            var remaining = FindDesktopWindow(target.Process, target.Window, exactOnly: true);
            if (FindVisibleNativeWindow(target.Process, target.Window) != 0 ||
                remaining != 0 && IsWindowVisible(remaining))
                throw new InvalidOperationException("Find and Replace remained visible after its close action.");
        }
    }

    private static nint FindVisibleNativeWindow(string processName, string windowName)
    {
        nint match = 0;
        EnumWindows((handle, _) =>
        {
            if (!IsWindowVisible(handle) || !GetNativeWindowTitle(handle).Equals(windowName,
                StringComparison.OrdinalIgnoreCase)) return true;
            GetWindowThreadProcessId(handle, out var processId);
            if (processId == 0) return true;
            try
            {
                if (Process.GetProcessById((int)processId).ProcessName.Equals(processName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    match = handle;
                    return false;
                }
            }
            catch (ArgumentException) { }
            return true;
        }, 0);
        return match;
    }

    private static async Task<string> DismissReplaceAllResultAsync(CancellationToken token, bool dismiss)
    {
        for (var attempt = 0; attempt < 50; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition)
                .Cast<AutomationElement>().ToList();
            var foreground = GetForegroundWindow();
            if (foreground != 0) windows.Insert(0, AutomationElement.FromHandle(foreground));
            foreach (AutomationElement popup in windows)
            {
                try
                {
                    var process = Process.GetProcessById(popup.Current.ProcessId).ProcessName;
                    if (!process.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)) continue;
                    var labels = popup.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                    var message = string.Join(" ", labels.Cast<AutomationElement>()
                        .Select(label => label.Current.Name).Where(value => !string.IsNullOrWhiteSpace(value)));
                    if (!(message.Contains("All done.", StringComparison.OrdinalIgnoreCase) &&
                          message.Contains("replacement", StringComparison.OrdinalIgnoreCase) ||
                          message.Contains("couldn't find anything to replace", StringComparison.OrdinalIgnoreCase))) continue;
                    if (dismiss)
                    {
                        var ok = popup.FindFirst(TreeScope.Descendants, new AndCondition(
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                            new PropertyCondition(AutomationElement.NameProperty, "OK")));
                        if (ok is null) throw new InvalidOperationException("Excel displayed a replacement result without an OK button.");
                        if (ok.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
                            ((InvokePattern)invoke).Invoke();
                        else if (!TryClickLiveUiaElement(ok, null, rightClick: false, out var detail))
                            throw new InvalidOperationException("Could not dismiss Excel's replacement result. " + detail);
                    }
                    if (message.Contains("couldn't find", StringComparison.OrdinalIgnoreCase) ||
                        Regex.IsMatch(message, @"\b0\s+replacements?\b", RegexOptions.IgnoreCase))
                        throw new InvalidOperationException("Replace All found no matching cells: " + message);
                    return message;
                }
                catch (Exception ex) when (ex is ElementNotAvailableException or System.Runtime.InteropServices.COMException or
                    ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception)
                { System.Diagnostics.Trace.WriteLine(ex); }
            }
            await Task.Delay(100, token);
        }
        throw new InvalidOperationException("Replace All was clicked, but Excel did not show a verifiable result.");
    }

    private static bool IsRefreshControl(ControlRef target) =>
        target.ControlType == "ControlType.Button" &&
        Regex.IsMatch(target.Name ?? "", @"\b(refresh|reload|sync)\b", RegexOptions.IgnoreCase);

    private static int TableClipboardLength()
    {
        try
        {
            if (!Clipboard.ContainsText(TextDataFormat.UnicodeText)) return 0;
            var text = Clipboard.GetText(TextDataFormat.UnicodeText);
            return text.Length >= 40 && (text.Contains('\t') || text.Contains('\n')) ? text.Length : 0;
        }
        catch (System.Runtime.InteropServices.ExternalException) { return 0; }
    }

    private static string? RefreshSnapshot(AutomationElement? element)
    {
        if (element is null) return null;
        try
        {
            var current = element.Current;
            var bounds = current.BoundingRectangle;
            return $"{Automation.State(element)};item={current.ItemStatus};offscreen={current.IsOffscreen};" +
                $"bounds={Math.Round(bounds.X)},{Math.Round(bounds.Y)},{Math.Round(bounds.Width)},{Math.Round(bounds.Height)};" +
                $"visual={RefreshVisualSignature(bounds)}";
        }
        catch (ElementNotAvailableException) { return null; }
    }

    private static void MovePointerAwayFrom(AutomationElement element)
    {
        var bounds = element.Current.BoundingRectangle;
        if (bounds.IsEmpty) return;
        var center = new System.Drawing.Point((int)Math.Round(bounds.Left + bounds.Width / 2),
            (int)Math.Round(bounds.Top + bounds.Height / 2));
        var area = Screen.FromPoint(center).WorkingArea;
        var corners = new[]
        {
            new System.Drawing.Point(area.Left + 12, area.Top + 12),
            new System.Drawing.Point(area.Right - 12, area.Top + 12),
            new System.Drawing.Point(area.Left + 12, area.Bottom - 12),
            new System.Drawing.Point(area.Right - 12, area.Bottom - 12)
        };
        var destination = corners.OrderByDescending(point =>
            Math.Pow(point.X - center.X, 2) + Math.Pow(point.Y - center.Y, 2)).First();
        if (!SetCursorPos(destination.X, destination.Y) || !GetCursorPos(out var cursor) ||
            cursor.X != destination.X || cursor.Y != destination.Y)
            throw new InvalidOperationException("Windows did not move the pointer away from the refresh control.");
    }

    private static string RefreshVisualSignature(System.Windows.Rect bounds)
    {
        var region = System.Drawing.Rectangle.FromLTRB((int)bounds.Left, (int)bounds.Top,
            (int)bounds.Right, (int)bounds.Bottom);
        if (region.Width < 3 || region.Height < 3 || region.Width > 250 || region.Height > 250)
            return "unavailable";
        try
        {
            using var bitmap = new Bitmap(region.Width, region.Height);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(region.Location, Point.Empty, region.Size);
            ulong hash = 14695981039346656037;
            for (var y = 1; y < bitmap.Height - 1; y += Math.Max(1, bitmap.Height / 12))
                for (var x = 1; x < bitmap.Width - 1; x += Math.Max(1, bitmap.Width / 12))
                {
                    hash ^= (uint)bitmap.GetPixel(x, y).ToArgb();
                    hash *= 1099511628211;
                }
            return hash.ToString("X16");
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or ArgumentException or OutOfMemoryException)
        {
            return "unavailable";
        }
    }

    private static Task<string?> WatchRefreshAsync(AutomationElement original, ControlRef target,
        CancellationToken token)
    {
        var before = RefreshSnapshot(original);
        return Task.Run(async () =>
        {
            var deadline = DateTime.UtcNow.AddSeconds(20);
            var changedSamples = 0;
            var returnedSamples = 0;
            var sawSustainedChange = false;
            while (DateTime.UtcNow < deadline)
            {
                token.ThrowIfCancellationRequested();
                var state = RefreshSnapshot(original);
                if (state is null) state = RefreshSnapshot(Automation.Resolve(target));
                if (state is null || state != before)
                {
                    if (++changedSamples >= 3) sawSustainedChange = true;
                    returnedSamples = 0;
                }
                else
                {
                    changedSamples = 0;
                    if (sawSustainedChange && ++returnedSamples >= 3) return "state-restored";
                }
                await Task.Delay(150, token);
            }
            return null;
        }, token);
    }

    private static void PauseForUnexpectedDialog(string? processName, HashSet<string?> knownWindows,
        Func<string, bool> handle, Action<string> status, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(processName)) return;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            var popup = UnexpectedModal(processName, knownWindows);
            if (popup is null) return;
            status($"Unexpected popup in {processName}: {popup}. Waiting for you to handle it.");
            if (!handle(popup)) throw new InvalidOperationException("Replay stopped while handling an unexpected popup.");
        }
    }

    private static string? UnexpectedModal(string processName, HashSet<string?> knownWindows)
    {
        string? ownedPopup = null;
        EnumWindows((handle, _) =>
        {
            GetWindowThreadProcessId(handle, out var processId);
            if (processId == 0 || !IsWindowVisible(handle) || GetWindow(handle, 4) == 0) return true;
            try
            {
                if (!Process.GetProcessById((int)processId).ProcessName
                    .Equals(processName, StringComparison.OrdinalIgnoreCase)) return true;
                var title = new StringBuilder(256);
                GetWindowText(handle, title, title.Capacity);
                var name = title.ToString();
                if (!string.IsNullOrWhiteSpace(name) && !knownWindows.Contains(name))
                { ownedPopup = name; return false; }
            }
            catch (ArgumentException) { }
            return true;
        }, 0);
        if (ownedPopup is not null) return ownedPopup;
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName
                    .Equals(processName, StringComparison.OrdinalIgnoreCase)) continue;
                var name = window.Current.Name;
                if (knownWindows.Contains(name)) continue;
                var modal = window.TryGetCurrentPattern(WindowPattern.Pattern, out var pattern) &&
                    ((WindowPattern)pattern).Current.IsModal;
                if (modal || window.Current.ClassName == "#32770")
                    return string.IsNullOrWhiteSpace(name) ? "Untitled dialog" : name;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or InvalidOperationException) { }
        }
        return null;
    }

    private static async Task<AutomationElement?> FindFirstListRowAsync(ControlRef target, CancellationToken token)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                var row = target.ControlType == "ControlType.HeaderItem"
                    ? await FindAsync(target, token, 2) is { } header
                        ? await Task.Run(() => Automation.FirstListItemUnderHeader(header), token) : null
                    : await Task.Run(() => Automation.FirstListItemInWindow(target), token);
                if (row is not null) return row;
            }
            catch (ElementNotAvailableException) { }
            if (attempt < 4) await Task.Delay(150, token);
        }
        return null;
    }

    private static async Task EnsureApplicationWindowAsync(string processName, string? windowName,
        Action<string> status, CancellationToken token, BrowserLaunchIdentity? browserLaunch = null,
        ControlRef? recordedTarget = null, Func<string, bool>? approveBrowserChange = null,
        HashSet<nint>? approvedBrowserWindows = null, int? freshBrowserProcessId = null,
        Action<int>? rememberFreshProcess = null, int? selectedBrowserProcessId = null,
        Action<int>? rememberSelectedProcess = null)
    {
        if (processName.Equals("SearchHost", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Windows Search must be opened through its dedicated replay path, not by starting SearchHost as a desktop application.");
        if (browserLaunch is not null)
        {
            nint checkedWindow;
            if (browserLaunch.AskForFreshProcess)
            {
                if (freshBrowserProcessId is null)
                    throw new InvalidOperationException("The plan requires a new Edge process, but its launch step did not run.");
                checkedWindow = FindMatchingBrowserWindow(processName, windowName, browserLaunch,
                    recordedTarget, processId: freshBrowserProcessId, requireCommandLine: false,
                    requireCategory: false);
            }
            else
            {
                checkedWindow = FindMatchingBrowserWindow(processName, windowName, browserLaunch,
                    recordedTarget);
                if (checkedWindow == 0 && selectedBrowserProcessId is int selectedId &&
                    windowName?.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase) == true)
                    checkedWindow = FindMatchingBrowserWindow(processName, windowName, browserLaunch,
                        null, processId: selectedId, requireCommandLine: false);
                if (checkedWindow == 0 && freshBrowserProcessId is int priorFreshProcess)
                    checkedWindow = FindMatchingBrowserWindow(processName, windowName, browserLaunch,
                        recordedTarget, processId: priorFreshProcess, requireCommandLine: false,
                        requireCategory: false);
                if (checkedWindow == 0)
                {
                    var matchingProcessExists = HasMatchingBrowserWindow(processName, browserLaunch);
                    var reason = matchingProcessExists
                        ? $"Edge has the recorded profile, but the recorded window '{windowName}' or its control is no longer available."
                        : browserLaunch.ProfilePath is not null
                            ? $"No running Edge window matches the recorded profile path '{browserLaunch.ProfilePath}'. The profile or Edge executable path may have changed."
                            : "No running Edge window matches the saved browser identity.";
                    if (approveBrowserChange is null || !approveBrowserChange(reason +
                        "\n\nMay replay automatically open a new Edge window with the default profile?"))
                        throw new OperationCanceledException("Replay stopped because the recorded Edge process is unavailable.");
                    var launchedId = await LaunchFreshBrowserAsync(browserLaunch, processName, status, token);
                    freshBrowserProcessId = launchedId;
                    rememberFreshProcess?.Invoke(launchedId);
                    checkedWindow = FindMatchingBrowserWindow(processName, windowName, browserLaunch,
                        recordedTarget, processId: launchedId, requireCommandLine: false,
                        requireCategory: false);
                    if (checkedWindow != 0) approvedBrowserWindows?.Add(checkedWindow);
                }
            }
            if (checkedWindow == 0 && freshBrowserProcessId is int newProcessId)
                for (var attempt = 0; attempt < 50 && checkedWindow == 0; attempt++)
                {
                    token.ThrowIfCancellationRequested();
                    await Task.Delay(100, token);
                    checkedWindow = FindMatchingBrowserWindow(processName, windowName,
                        browserLaunch, recordedTarget, processId: newProcessId,
                        requireCommandLine: false, requireCategory: false);
                }
            if (checkedWindow == 0)
                throw new InvalidOperationException($"The recorded browser window '{windowName}' " +
                    "could not be identified uniquely after launching Edge.");
            ShowWindow(checkedWindow, 9);
            GetWindowThreadProcessId(checkedWindow, out var selectedPid);
            rememberSelectedProcess?.Invoke((int)selectedPid);
            MaximizeWindowIfSupported(checkedWindow);
            SetForegroundWindow(checkedWindow);
            return;
        }
        var instances = Process.GetProcessesByName(processName);
        var handle = FindDesktopWindow(processName, windowName, exactOnly: !string.IsNullOrWhiteSpace(windowName));
        if (handle == 0 && !string.IsNullOrWhiteSpace(windowName) && instances.Length > 0)
        {
            // A command in the same process may just have opened a new window.
            // Wait for that window so it is maximized before its first action.
            for (var attempt = 0; attempt < 50 && handle == 0; attempt++)
            {
                token.ThrowIfCancellationRequested();
                await Task.Delay(100, token);
                handle = FindDesktopWindow(processName, windowName, exactOnly: true);
            }
        }
        if (handle == 0) handle = FindDesktopWindow(processName, windowName);
        if (handle != 0)
        {
            ShowWindow(handle, 9);
            MaximizeWindowIfSupported(handle);
            SetForegroundWindow(handle);
            return;
        }
        var executable = instances.Select(p =>
        {
            try { return p.MainModule?.FileName; }
            catch { return null; }
        }).FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));
        if (executable is null)
            throw new InvalidOperationException($"{processName} has no open desktop window and its executable could not be located.");
        status($"Opening {processName} to continue replay");
        try
        {
            if (!StartPackagedApplication(executable))
                Process.Start(new ProcessStartInfo(executable) { UseShellExecute = true });
        }
        catch (Exception ex) { throw new InvalidOperationException($"{processName} has no open desktop window and could not be opened: {ex.Message}", ex); }
        for (var attempt = 0; attempt < 30; attempt++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(200, token);
            handle = FindDesktopWindow(processName, windowName);
            if (handle == 0) continue;
            ShowWindow(handle, 9);
            MaximizeWindowIfSupported(handle);
            SetForegroundWindow(handle);
            return;
        }
        throw new InvalidOperationException($"{processName} was started but no desktop window appeared.");
    }

    private static async Task<int> LaunchFreshBrowserAsync(BrowserLaunchIdentity? recorded,
        string processName, Action<string> status, CancellationToken token)
    {
        if (!processName.Equals("msedge", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("The browser launch step does not target Edge.");
        var executable = recorded?.ExecutablePath;
        if (string.IsNullOrWhiteSpace(executable) || !File.Exists(executable))
            executable = Process.GetProcessesByName(processName).Select(process =>
            {
                try { return process.MainModule?.FileName; }
                catch { return null; }
            }).FirstOrDefault(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path));
        if (string.IsNullOrWhiteSpace(executable))
            throw new InvalidOperationException("Edge executable could not be located for the recorded automatic launch.");
        var start = new ProcessStartInfo(executable) { UseShellExecute = false };
        start.ArgumentList.Add("--profile-directory=Default");
        start.ArgumentList.Add("--new-window");
        var before = AutomationElement.RootElement.FindAll(TreeScope.Children,
            Condition.TrueCondition).Cast<AutomationElement>().Select(window =>
        {
            try { return (nint)window.Current.NativeWindowHandle; }
            catch (ElementNotAvailableException) { return 0; }
        }).Where(handle => handle != 0).ToHashSet();
        using var launched = Process.Start(start) ??
            throw new InvalidOperationException("Edge did not start.");
        status("Opened a new Edge window using the default profile.");
        for (var attempt = 0; attempt < 100; attempt++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(100, token);
            var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement window in windows)
            {
                try
                {
                    var pid = window.Current.ProcessId;
                    if (!Process.GetProcessById(pid).ProcessName.Equals(processName,
                            StringComparison.OrdinalIgnoreCase)) continue;
                    var handle = (nint)window.Current.NativeWindowHandle;
                    if (handle == 0 || before.Contains(handle) || !IsWindowVisible(handle)) continue;
                    ShowWindow(handle, 9);
                    SetForegroundWindow(handle);
                    return pid;
                }
                catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                    InvalidOperationException or COMException) { Trace.WriteLine(ex); }
            }
        }
        throw new InvalidOperationException("Edge started, but its new desktop window did not appear.");
    }

    private static nint FindMatchingBrowserWindow(string processName, string? windowName,
        BrowserLaunchIdentity recorded, ControlRef? recordedTarget,
        int? processId = null, bool requireCommandLine = true, bool requireCategory = true)
    {
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        var category = BrowserWindowCategory(windowName);
        var candidates = new List<(nint Handle, int ProcessId, bool TitleMatches, bool HasTarget)>();
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(
                        processName, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (processId is not null && window.Current.ProcessId != processId) continue;
                var handle = (nint)window.Current.NativeWindowHandle;
                if (handle == 0 || !IsWindowVisible(handle)) continue;
                if (requireCategory && category is not null && !string.Equals(category,
                        BrowserWindowCategory(window.Current.Name),
                        StringComparison.OrdinalIgnoreCase)) continue;
                var hasTarget = false;
                try { hasTarget = BrowserWindowHasRecordedControl(window, recordedTarget); }
                catch (Exception ex) when (ex is ElementNotAvailableException or
                    InvalidOperationException or COMException) { Trace.WriteLine(ex); }
                candidates.Add((handle, window.Current.ProcessId,
                    !string.IsNullOrWhiteSpace(windowName) &&
                    window.Current.Name.Equals(windowName, StringComparison.OrdinalIgnoreCase),
                    hasTarget));
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or COMException) { Trace.WriteLine(ex); }
        }
        var matches = new List<(nint Handle, bool TitleMatches, bool HasTarget)>();
        foreach (var candidate in candidates.OrderByDescending(item => item.TitleMatches))
        {
            var live = requireCommandLine ? BrowserIdentity.Read(candidate.ProcessId,
                GetNativeWindowTitle(candidate.Handle)) : null;
            if (!requireCommandLine || live is not null && BrowserIdentity.Matches(recorded, live, out _))
                matches.Add((candidate.Handle, candidate.TitleMatches, candidate.HasTarget));
        }
        if (matches.Count == 0) return 0;
        var exact = matches.Where(item => item.TitleMatches).ToList();
        if (exact.Count == 1) return exact[0].Handle;
        var withTarget = matches.Where(item => item.HasTarget).ToList();
        if (withTarget.Count == 1) return withTarget[0].Handle;
        var foreground = GetForegroundWindow();
        if (matches.Any(item => item.Handle == foreground)) return foreground;
        if (recordedTarget is { Name: "Maximize", ClassName: "EdgeWindowsCaptionButton" })
            return matches[0].Handle;
        return matches.Count == 1 ? matches[0].Handle : 0;
    }

    private static bool HasMatchingBrowserWindow(string processName,
        BrowserLaunchIdentity recorded)
    {
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children,
            Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                var handle = (nint)window.Current.NativeWindowHandle;
                if (handle == 0 || !IsWindowVisible(handle) ||
                    !Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(
                        processName, StringComparison.OrdinalIgnoreCase)) continue;
                var live = BrowserIdentity.Read(window.Current.ProcessId, window.Current.Name);
                if (live is not null && BrowserIdentity.Matches(recorded, live, out _))
                    return true;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or COMException) { Trace.WriteLine(ex); }
        }
        return false;
    }

    private static string? BrowserWindowCategory(string? title)
    {
        if (string.IsNullOrWhiteSpace(title)) return null;
        if (title.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase)) return "[InPrivate]";
        var browserSuffix = title.LastIndexOf(" - Microsoft", StringComparison.OrdinalIgnoreCase);
        if (browserSuffix < 0) return null;
        var profileStart = title.LastIndexOf(" - ", browserSuffix - 1, StringComparison.Ordinal);
        return profileStart < 0 ? null : title[(profileStart + 3)..browserSuffix];
    }

    private static bool BrowserWindowHasRecordedControl(AutomationElement window,
        ControlRef? recordedTarget)
    {
        if (string.IsNullOrWhiteSpace(recordedTarget?.AutomationId)) return true;
        var conditions = new List<Condition>
        {
            new PropertyCondition(AutomationElement.AutomationIdProperty,
                recordedTarget.AutomationId)
        };
        if (!string.IsNullOrWhiteSpace(recordedTarget.ClassName))
            conditions.Add(new PropertyCondition(AutomationElement.ClassNameProperty,
                recordedTarget.ClassName));
        var condition = conditions.Count == 1 ? conditions[0] : new AndCondition([.. conditions]);
        var control = window.FindFirst(TreeScope.Descendants, condition);
        return control is not null && control.Current.IsEnabled && !control.Current.IsOffscreen;
    }

    private static List<nint> VisiblePrivateBrowserWindows(string processName)
    {
        var matches = new List<nint>();
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!window.Current.Name.Contains("[InPrivate]", StringComparison.OrdinalIgnoreCase) ||
                    !Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(processName,
                        StringComparison.OrdinalIgnoreCase)) continue;
                var handle = (nint)window.Current.NativeWindowHandle;
                if (handle != 0 && IsWindowVisible(handle)) matches.Add(handle);
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or COMException) { Trace.WriteLine(ex); }
        }
        return matches;
    }

    private static async Task VerifyRecordedSearchLaunchAsync(PlanStep step, CancellationToken token)
    {
        for (var attempt = 0; attempt < 50 &&
            FindVisibleNativeWindow("SearchHost", "Search") != 0; attempt++)
            await Task.Delay(100, token);
        if (FindVisibleNativeWindow("SearchHost", "Search") != 0)
            throw new InvalidOperationException($"Step {step.Number}: Windows Search remained open after the result action.");
        if (step.ExpectedState?.StartsWith("launched-window:", StringComparison.Ordinal) != true)
            return;
        var identity = step.ExpectedState["launched-window:".Length..];
        var separator = identity.IndexOf('|');
        if (separator <= 0 || separator == identity.Length - 1)
            throw new InvalidDataException($"Step {step.Number}: the recorded Search launch identity is invalid.");
        var process = identity[..separator];
        var window = identity[(separator + 1)..];
        for (var attempt = 0; attempt < 30; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var handle = FindDesktopWindow(process, window, exactOnly: true);
            if (handle != 0 && IsWindowVisible(handle)) return;
            await Task.Delay(100, token);
        }
        throw new InvalidOperationException($"Step {step.Number}: Windows Search closed, but recorded window '{window}' ({process}) did not appear.");
    }

    private static bool WindowsSearchStillActive(ControlRef searchBoxTarget)
    {
        var foreground = GetForegroundWindow();
        if (foreground == 0) return false;
        GetWindowThreadProcessId(foreground, out var processId);
        if (processId == 0) return false;
        try
        {
            if (!Process.GetProcessById((int)processId).ProcessName.Equals("SearchHost",
                StringComparison.OrdinalIgnoreCase)) return false;
            var searchBox = Automation.Resolve(searchBoxTarget);
            return searchBox is not null && !searchBox.Current.IsOffscreen;
        }
        catch (Exception ex) when (ex is ArgumentException or ElementNotAvailableException or
            InvalidOperationException or COMException)
        {
            Trace.WriteLine(ex);
            return false;
        }
    }

    private static Dictionary<nint, (string Process, string Title)> VisibleApplicationWindows()
    {
        var windows = new Dictionary<nint, (string Process, string Title)>();
        var ownProcess = Process.GetCurrentProcess().ProcessName;
        EnumWindows((handle, _) =>
        {
            if (!IsWindowVisible(handle)) return true;
            var title = GetNativeWindowTitle(handle);
            if (string.IsNullOrWhiteSpace(title)) return true;
            GetWindowThreadProcessId(handle, out var processId);
            if (processId == 0) return true;
            try
            {
                var process = Process.GetProcessById((int)processId).ProcessName;
                if (process.Equals("SearchHost", StringComparison.OrdinalIgnoreCase) ||
                    process.Equals("explorer", StringComparison.OrdinalIgnoreCase) ||
                    process.Equals("consent", StringComparison.OrdinalIgnoreCase) ||
                    process.Equals(ownProcess, StringComparison.OrdinalIgnoreCase)) return true;
                windows[handle] = (process, title);
            }
            catch (ArgumentException) { }
            return true;
        }, 0);
        return windows;
    }

    private static async Task<(string Process, string Title)?> WaitForNewApplicationWindowAsync(
        IReadOnlyDictionary<nint, (string Process, string Title)> before, CancellationToken token)
    {
        for (var attempt = 0; attempt < 80; attempt++)
        {
            token.ThrowIfCancellationRequested();
            foreach (var (handle, identity) in VisibleApplicationWindows())
                if (!before.ContainsKey(handle)) return identity;
            await Task.Delay(100, token);
        }
        return null;
    }

    private static async Task<(string Process, string Title)?> WaitForRecordedApplicationWindowAsync(
        string processName, string? preferredWindow, CancellationToken token)
    {
        for (var attempt = 0; attempt < 80; attempt++)
        {
            token.ThrowIfCancellationRequested();
            if (!string.IsNullOrWhiteSpace(preferredWindow))
            {
                var exact = FindDesktopWindow(processName, preferredWindow, exactOnly: true);
                if (exact != 0 && IsWindowVisible(exact))
                    return (processName, GetNativeWindowTitle(exact));
            }
            var visible = VisibleApplicationWindows().Values.FirstOrDefault(identity =>
                identity.Process.Equals(processName, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(visible.Process)) return visible;
            await Task.Delay(100, token);
        }
        return null;
    }

    private static async Task<AutomationElement?> EnsureWindowsSearchOpenAsync(PlanStep step,
        CancellationToken token)
    {
        static AutomationElement? FindSearchBox()
        {
            try
            {
                var focused = AutomationElement.FocusedElement;
                if (focused is not null && !focused.Current.IsOffscreen &&
                    focused.Current.AutomationId == "SearchTextBox" &&
                    focused.Current.ControlType == ControlType.Edit)
                    return focused;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
            { Trace.WriteLine(ex); }
            try
            {
                var condition = new AndCondition(
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "SearchTextBox"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                foreach (AutomationElement candidate in AutomationElement.RootElement.FindAll(
                    TreeScope.Descendants, condition))
                {
                    try
                    {
                        if (!candidate.Current.IsOffscreen &&
                            candidate.Current.ClassName == "RichEditBox" &&
                            !candidate.Current.BoundingRectangle.IsEmpty)
                            return candidate;
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or
                        InvalidOperationException or COMException) { Trace.WriteLine(ex); }
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
            { Trace.WriteLine(ex); }
            return null;
        }

        var existing = await Task.Run(FindSearchBox, token);
        if (existing is not null)
        {
            try { existing.SetFocus(); return existing; }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
            { Trace.WriteLine(ex); }
        }
        keybd_event(0x5B, 0, 0, 0); // Win down.
        keybd_event(0x53, 0, 0, 0); // S down.
        keybd_event(0x53, 0, 2, 0);
        keybd_event(0x5B, 0, 2, 0);
        for (var attempt = 0; attempt < 50; attempt++)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(100, token);
            var searchBox = await Task.Run(FindSearchBox, token);
            if (searchBox is null) continue;
            try
            {
                if (searchBox.Current.IsOffscreen) continue;
                searchBox.SetFocus();
                return searchBox;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
            { Trace.WriteLine(ex); }
        }
        var foreground = GetForegroundWindow();
        if (foreground != 0)
        {
            GetWindowThreadProcessId(foreground, out var processId);
            try
            {
                if (processId != 0 && Process.GetProcessById((int)processId).ProcessName
                    .Equals("SearchHost", StringComparison.OrdinalIgnoreCase))
                    return null;
            }
            catch (ArgumentException) { }
        }
        throw new InvalidOperationException($"Step {step.Number}: Windows Search did not expose its search field, and SearchHost was not foreground. Replay stopped before typing.");
    }

    private static bool StartPackagedApplication(string executable)
    {
        var folder = new DirectoryInfo(Path.GetDirectoryName(executable)!);
        while (folder is not null)
        {
            var manifestPath = Path.Combine(folder.FullName, "AppxManifest.xml");
            if (File.Exists(manifestPath))
            {
                var manifest = XDocument.Load(manifestPath);
                var app = manifest.Descendants().FirstOrDefault(x => x.Name.LocalName == "Application" &&
                    string.Equals(Path.GetFileName((string?)x.Attribute("Executable")),
                        Path.GetFileName(executable), StringComparison.OrdinalIgnoreCase));
                var parts = folder.Name.Split('_');
                if (app is null || parts.Length < 2) return false;
                var appId = (string?)app.Attribute("Id");
                if (string.IsNullOrWhiteSpace(appId)) return false;
                var family = parts[0] + "_" + parts[^1];
                Process.Start(new ProcessStartInfo("explorer.exe")
                { ArgumentList = { $"shell:AppsFolder\\{family}!{appId}" }, UseShellExecute = true });
                return true;
            }
            folder = folder.Parent;
        }
        return false;
    }

    private static nint FindDesktopWindow(string processName, string? windowName = null, bool exactOnly = false)
    {
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        nint fallback = 0;
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (Process.GetProcessById(window.Current.ProcessId).ProcessName
                    .Equals(processName, StringComparison.OrdinalIgnoreCase))
                {
                    var handle = (nint)window.Current.NativeWindowHandle;
                    if (handle == 0) continue;
                    if (!string.IsNullOrWhiteSpace(windowName) &&
                        window.Current.Name.Equals(windowName, StringComparison.OrdinalIgnoreCase))
                        return handle;
                    if (fallback == 0) fallback = handle;
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or InvalidOperationException) { }
        }
        return exactOnly ? 0 : fallback;
    }

    private static List<(nint Dialog, nint Root)> FindNativeWindows(string processName, string windowName)
    {
        var matches = new List<(nint Dialog, nint Root)>();
        EnumWindows((root, _) =>
        {
            if (!IsWindowVisible(root)) return true;
            GetWindowThreadProcessId(root, out var processId);
            if (processId == 0) return true;
            try
            {
                if (!Process.GetProcessById((int)processId).ProcessName.Equals(processName,
                    StringComparison.OrdinalIgnoreCase)) return true;
            }
            catch (ArgumentException) { return true; }
            if (GetNativeWindowTitle(root).Equals(windowName, StringComparison.OrdinalIgnoreCase))
                matches.Add((root, root));
            EnumChildWindows(root, (child, _) =>
            {
                if (IsWindowVisible(child) && GetNativeWindowTitle(child).Equals(windowName,
                    StringComparison.OrdinalIgnoreCase)) matches.Add((child, root));
                return true;
            }, 0);
            return true;
        }, 0);
        return matches.DistinctBy(match => match.Dialog).ToList();
    }

    private static bool IsNativeDescendantOf(nint candidate, nint ancestor)
    {
        if (candidate == 0 || ancestor == 0) return false;
        for (var depth = 0; candidate != 0 && depth < 32; depth++)
        {
            if (candidate == ancestor) return true;
            var parent = GetParent(candidate);
            if (parent == candidate) break;
            candidate = parent;
        }
        return false;
    }

    private static string NormalizeNativeCaption(string? value) =>
        Regex.Replace((value ?? "").Replace("&", "").Trim(), @"\s+", " ");

    private static bool IsRecordedSectionGroup(ControlRef target)
    {
        try
        {
            var group = Automation.Resolve(target with
            {
                ControlType = "ControlType.Group", ClassName = null, AutomationId = null
            });
            if (group is not null && !group.Current.IsOffscreen)
                return true;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
        { Trace.WriteLine(ex); }
        var dialogs = FindNativeWindows(target.Process!, target.Window!);
        if (dialogs.Count != 1) return false;
        var matches = 0;
        EnumChildWindows(dialogs[0].Dialog, (child, _) =>
        {
            var className = new StringBuilder(64);
            GetClassName(child, className, className.Capacity);
            if (!className.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase) ||
                (GetWindowLong(child, -16) & 0x000F) != 0x0007) return true; // BS_GROUPBOX
            var caption = new StringBuilder(256);
            GetWindowText(child, caption, caption.Capacity);
            if (NormalizeNativeCaption(caption.ToString()).Equals(NormalizeNativeCaption(target.Name),
                StringComparison.OrdinalIgnoreCase)) matches++;
            return matches <= 1;
        }, 0);
        return matches == 1;
    }

    private static bool NativeDialogHasButton(ControlRef target, ControlRef followingEdit)
    {
        if (string.IsNullOrWhiteSpace(target.Process) || string.IsNullOrWhiteSpace(target.Window) ||
            string.IsNullOrWhiteSpace(target.Name) ||
            !int.TryParse(followingEdit.AutomationId, out var editId)) return true;
        var dialogs = FindNativeWindows(target.Process, target.Window);
        var edits = new HashSet<nint>();
        foreach (var (dialog, _) in dialogs)
        {
            EnumChildWindows(dialog, (child, _) =>
            {
                if (GetDlgCtrlID(child) != editId || !IsWindowVisible(child) ||
                    !IsWindowEnabled(child)) return true;
                var className = new StringBuilder(64);
                GetClassName(child, className, className.Capacity);
                if (className.ToString().Equals("Edit", StringComparison.OrdinalIgnoreCase))
                    edits.Add(child);
                return true;
            }, 0);
        }
        // A title can identify both a parent and its child dialog. Use the
        // nearest dialog ancestor of the following Edit, which is the control
        // the next step will actually change.
        if (edits.Count != 1) return true;
        var candidateHandles = dialogs.Select(candidate => candidate.Dialog).ToHashSet();
        nint activeDialog = 0;
        for (var current = GetParent(edits.Single()); current != 0; current = GetParent(current))
        {
            if (!candidateHandles.Contains(current)) continue;
            activeDialog = current;
            break;
        }
        if (activeDialog == 0) return true;
        var found = false;
        EnumChildWindows(activeDialog, (child, _) =>
        {
            if (!IsWindowVisible(child) || !IsWindowEnabled(child)) return true;
            var className = new StringBuilder(64);
            GetClassName(child, className, className.Capacity);
            if (!className.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase) ||
                (GetWindowLong(child, -16) & 0x000F) is not (0x0004 or 0x0009))
                return true;
            var caption = new StringBuilder(256);
            GetWindowText(child, caption, caption.Capacity);
            found = NormalizeNativeCaption(caption.ToString()).Equals(
                NormalizeNativeCaption(target.Name), StringComparison.OrdinalIgnoreCase);
            return !found;
        }, 0);
        return found;
    }

    private static bool TrySelectSectionComboThroughUia(ControlRef heading, ControlRef option,
        out string detail)
    {
        detail = "the recorded section did not expose its named combo box";
        if (string.IsNullOrWhiteSpace(option.ParentName) || string.IsNullOrWhiteSpace(option.Name))
            return false;
        try
        {
            var group = Automation.Resolve(heading with
            {
                ControlType = "ControlType.Group", ClassName = null, AutomationId = null
            });
            if (group is null) return false;
            var combos = group.FindAll(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                new PropertyCondition(AutomationElement.NameProperty, option.ParentName)));
            if (combos.Count != 1 || !combos[0].Current.IsEnabled || combos[0].Current.IsOffscreen)
                return false;
            var combo = combos[0];
            var comboTarget = heading with { Name = option.ParentName,
                ControlType = "ControlType.ComboBox", ClassName = "ComboBox" };
            ActivateWindow(combo);
            if (!OpenRecordedComboBox(combo, new PlanStep(0, "click", comboTarget, null,
                null, null, null, null)))
            {
                detail = "the named combo box could not be opened";
                return false;
            }
            var item = Automation.Resolve(option with { Window = heading.Window,
                ParentName = option.ParentName });
            if (item is null || !TryClickLiveUiaElement(item, null, false, out detail))
                return false;
            if (!NativeComboHasSelectedValue((nint)combo.Current.NativeWindowHandle, option.Name))
            {
                detail = $"the named combo box did not retain '{option.Name}' after the click";
                return false;
            }
            detail = $"selected and verified '{option.Name}' in combo box '{option.ParentName}'";
            return true;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
        {
            Trace.WriteLine(ex);
            detail = $"UI Automation could not complete the combo selection: {ex.GetType().Name}";
            return false;
        }
    }

    private static bool NativeComboHasSelectedValue(nint combo, string desired)
    {
        if (combo == 0 || SendMessageTimeout(combo, 0x0147, 0, 0, 0x0002, 500,
            out var index) == 0 || (int)index < 0) return false; // CB_GETCURSEL
        var label = new StringBuilder(Math.Max(256, desired.Length + 1));
        return SendMessageTimeout(combo, 0x0148, index, label, 0x0002, 500,
            out _) != 0 && label.ToString().Equals(desired, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TrySelectUniqueNativeComboValue(ControlRef dialogTarget, string desired,
        out string detail)
    {
        detail = "the native combo value was not found";
        var dialogs = FindNativeWindows(dialogTarget.Process!, dialogTarget.Window!);
        if (dialogs.Count == 0)
        {
            detail = "the recorded dialog was not visible";
            return false;
        }
        var candidates = new List<(nint Combo, int Index, nint Root)>();
        foreach (var (dialog, root) in dialogs)
        {
            EnumChildWindows(dialog, (child, _) =>
            {
                if (!IsWindowVisible(child) || !IsWindowEnabled(child)) return true;
                var className = new StringBuilder(64);
                GetClassName(child, className, className.Capacity);
                if (!className.ToString().Equals("ComboBox", StringComparison.OrdinalIgnoreCase)) return true;
                if (SendMessageTimeout(child, 0x0146, 0, 0, 0x0002, 500, out var countResult) == 0)
                    return true; // CB_GETCOUNT
                var count = (int)countResult;
                for (var itemIndex = 0; itemIndex < count && itemIndex < 200; itemIndex++)
                {
                    if (SendMessageTimeout(child, 0x0149, itemIndex, 0, 0x0002, 500,
                        out var lengthResult) == 0 || (int)lengthResult < 0) continue; // CB_GETLBTEXTLEN
                    var label = new StringBuilder(Math.Max(256, (int)lengthResult + 1));
                    if (SendMessageTimeout(child, 0x0148, itemIndex, label, 0x0002, 500,
                        out _) == 0) continue; // CB_GETLBTEXT
                    if (label.ToString().Equals(desired, StringComparison.OrdinalIgnoreCase))
                        candidates.Add((child, itemIndex, root));
                }
                return true;
            }, 0);
        }
        candidates = candidates.DistinctBy(candidate => (candidate.Combo, candidate.Index)).ToList();
        if (candidates.Count > 1)
        {
            var foreground = GetForegroundWindow();
            var foregroundRoot = foreground;
            for (var depth = 0; foregroundRoot != 0 && depth < 20; depth++)
            {
                var ancestor = GetParent(foregroundRoot);
                if (ancestor == 0) break;
                foregroundRoot = ancestor;
            }
            var activeCandidates = candidates.Where(candidate => candidate.Root == foreground ||
                candidate.Root == foregroundRoot).ToList();
            if (activeCandidates.Count == 1) candidates = activeCandidates;
        }
        if (candidates.Count != 1)
        {
            detail = $"found {candidates.Count} native combo items named '{desired}'";
            return false;
        }
        var (combo, index, rootWindow) = candidates[0];
        ShowWindow(rootWindow, 9);
        SetForegroundWindow(rootWindow);
        if (SendMessageTimeout(combo, 0x014E, index, 0, 0x0002, 500,
            out var selectedResult) == 0 || (int)selectedResult < 0) // CB_SETCURSEL
        {
            detail = $"native combo rejected '{desired}'";
            return false;
        }
        var parent = GetParent(combo);
        var id = GetDlgCtrlID(combo);
        if (parent != 0 && id > 0)
            SendMessageTimeout(parent, 0x0111, (id & 0xFFFF) | (1 << 16), combo,
                0x0002, 500, out _); // WM_COMMAND / CBN_SELCHANGE
        if (SendMessageTimeout(combo, 0x0147, 0, 0, 0x0002, 500,
            out var actual) == 0 || (int)actual != index) // CB_GETCURSEL
        {
            detail = $"native combo did not retain '{desired}'";
            return false;
        }
        detail = $"selected and verified native combo value '{desired}'";
        return true;
    }

    private static bool TryClickUniqueNativeButton(nint dialog, string name, string recordedType)
    {
        var expectedName = NormalizeNativeCaption(name);
        var matches = new List<nint>();
        EnumChildWindows(dialog, (child, _) =>
        {
            var className = new StringBuilder(64);
            GetClassName(child, className, className.Capacity);
            if (!className.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase) ||
                !IsWindowVisible(child) || !IsWindowEnabled(child)) return true;
            var label = new StringBuilder(256);
            GetWindowText(child, label, label.Capacity);
            if (!NormalizeNativeCaption(label.ToString()).Equals(expectedName, StringComparison.OrdinalIgnoreCase))
                return true;
            var style = GetWindowLong(child, -16) & 0x000F;
            if (style == 0x0007) return true; // A group box labels a section; it is not a command.
            var nativeType = style switch
            {
                0x0002 or 0x0003 or 0x0005 or 0x0006 => "ControlType.CheckBox",
                0x0004 or 0x0009 => "ControlType.RadioButton",
                _ => "ControlType.Button"
            };
            if (recordedType == "ControlType.Button" || recordedType == nativeType)
                matches.Add(child);
            return matches.Count <= 1;
        }, 0);
        if (matches.Count != 1) return false;
        return SendMessageTimeout(matches[0], 0x00F5, 0, 0, 0x0002, 500, out _) != 0;
    }

    private static bool TrySetUniqueNativeEdit(ControlRef target, string desired, out string detail)
    {
        detail = "the recorded control identity is incomplete";
        if (!int.TryParse(target.AutomationId, out var controlId) || controlId <= 0 ||
            string.IsNullOrWhiteSpace(target.Process) || string.IsNullOrWhiteSpace(target.Window))
            return false;
        var dialogs = FindNativeWindows(target.Process, target.Window);
        if (dialogs.Count == 0)
        {
            detail = $"found no visible native dialog named '{target.Window}' in {target.Process}";
            return false;
        }
        var matches = new List<(nint Edit, nint Dialog, nint Root)>();
        foreach (var (candidateDialog, candidateRoot) in dialogs)
        {
            EnumChildWindows(candidateDialog, (child, _) =>
            {
                if (GetDlgCtrlID(child) != controlId || !IsWindowVisible(child) ||
                    !IsWindowEnabled(child)) return true;
                var className = new StringBuilder(64);
                GetClassName(child, className, className.Capacity);
                if (className.ToString().Equals("Edit", StringComparison.OrdinalIgnoreCase))
                    matches.Add((child, candidateDialog, candidateRoot));
                return true;
            }, 0);
        }
        matches = matches.DistinctBy(match => match.Edit).ToList();
        if (matches.Count > 1)
        {
            var foreground = GetForegroundWindow();
            var active = matches.Where(match => match.Dialog == foreground ||
                match.Root == foreground ||
                IsNativeDescendantOf(foreground, match.Dialog) ||
                IsNativeDescendantOf(match.Dialog, foreground)).ToList();
            if (active.Count == 1) matches = active;
        }
        if (matches.Count != 1)
        {
            detail = $"found {matches.Count} distinct visible enabled Edit controls with ID " +
                $"{controlId} across {dialogs.Count} native handles named '{target.Window}'";
            return false;
        }
        var (edit, _, rootWindow) = matches[0];
        var current = new StringBuilder(Math.Max(256, desired.Length + 1));
        if (SendMessageTimeout(edit, 0x000D, (nint)current.Capacity, current,
            0x0002, 500, out _) == 0)
        {
            detail = "the native Edit did not return its current text";
            return false;
        }
        if (current.ToString() == desired) return true;
        var replacement = new StringBuilder(desired);
        if (SendMessageTimeout(edit, 0x000C, 0, replacement, 0x0002, 500, out _) != 0)
        {
            Thread.Sleep(100);
            var afterSet = new StringBuilder(Math.Max(256, desired.Length + 1));
            if (SendMessageTimeout(edit, 0x000D, (nint)afterSet.Capacity, afterSet,
                0x0002, 500, out _) != 0 && afterSet.ToString() == desired)
                return true;
        }
        if (!GetWindowRect(edit, out var bounds) || bounds.Right - bounds.Left < 8 ||
            bounds.Bottom - bounds.Top < 8)
        {
            detail = "the native Edit has no usable live rectangle";
            return false;
        }
        var x = bounds.Left + (bounds.Right - bounds.Left) / 3;
        var y = bounds.Top + (bounds.Bottom - bounds.Top) / 2;
        if (!Screen.AllScreens.Any(screen => screen.Bounds.Contains(x, y)))
        {
            detail = "the native Edit lies outside the visible desktop";
            return false;
        }
        ShowWindow(rootWindow, 9);
        SetForegroundWindow(rootWindow);
        GetWindowThreadProcessId(GetForegroundWindow(), out var foregroundProcess);
        GetWindowThreadProcessId(rootWindow, out var targetProcess);
        if (foregroundProcess != targetProcess || !SetCursorPos(x, y))
        {
            detail = "the verified target process could not receive a live click";
            return false;
        }
        mouse_event(0x0002u, 0, 0, 0, 0);
        mouse_event(0x0004u, 0, 0, 0, 0);
        if (SendMessageTimeout(edit, 0x00B1, 0, -1, 0x0002, 500, out _) == 0)
        {
            detail = "the native Edit did not accept full-text selection";
            return false;
        }
        SendKeys.SendWait(string.Concat(desired.Select(EscapeText)));
        Thread.Sleep(100);
        var observed = new StringBuilder(Math.Max(256, desired.Length + 1));
        if (SendMessageTimeout(edit, 0x000D, (nint)observed.Capacity, observed,
            0x0002, 500, out _) != 0 && observed.ToString() == desired) return true;
        detail = $"the native Edit still showed '{observed}' after input";
        return false;
    }

    private static async Task<string> ConfirmDialogAndWaitForNewTabAsync(ControlRef dialogTarget,
        ControlRef originalTab, HashSet<string> before, CancellationToken token)
    {
        var dialogHandle = FindDesktopWindow(dialogTarget.Process!, dialogTarget.Window, exactOnly: true);
        if (dialogHandle == 0)
            throw new InvalidOperationException($"The '{dialogTarget.Window}' dialog is unavailable and the recording contains no confirmation action. The copied sheet was not verified.");
        ShowWindow(dialogHandle, 9);
        SetForegroundWindow(dialogHandle);
        var dialog = AutomationElement.FromHandle(dialogHandle);
        var copyOptions = dialog.FindAll(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox),
            new PropertyCondition(AutomationElement.NameProperty, "Create a copy")));
        if (copyOptions.Count != 1 || !copyOptions[0].TryGetCurrentPattern(TogglePattern.Pattern, out var toggle))
            throw new InvalidOperationException("Move or Copy did not expose one verifiable 'Create a copy' checkbox. The dialog was left open; no sheet was moved.");
        if (((TogglePattern)toggle).Current.ToggleState != ToggleState.On)
            ((TogglePattern)toggle).Toggle();
        if (((TogglePattern)toggle).Current.ToggleState != ToggleState.On)
            throw new InvalidOperationException("The 'Create a copy' checkbox did not become checked. The dialog was left open; no sheet was moved.");
        var okButtons = dialog.FindAll(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
            new PropertyCondition(AutomationElement.NameProperty, "OK")));
        if (okButtons.Count > 1)
            throw new InvalidOperationException($"The '{dialogTarget.Window}' dialog exposed multiple OK buttons. The copy was not committed.");
        if (okButtons.Count == 0)
        {
            if (!MsaaActions.TryInvokeUnique(dialogHandle, "OK", 0x2B))
                throw new InvalidOperationException($"The '{dialogTarget.Window}' dialog exposed no unique OK action through UI Automation or MSAA. The copy was not committed.");
        }
        else
        {
            var ok = okButtons[0];
            if (ok.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
                ((InvokePattern)invoke).Invoke();
            else if (!TryClickLiveUiaElement(ok, null, rightClick: false, out var detail))
                throw new InvalidOperationException($"Could not confirm '{dialogTarget.Window}'. {detail}");
        }
        return await WaitForNewTabAsync(originalTab, before, token);
    }

    private static async Task<string> WaitForNewTabAsync(ControlRef originalTab,
        HashSet<string> before, CancellationToken token)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var added = await Task.Run(() =>
            {
                var tab = Automation.Resolve(originalTab);
                return tab is null ? [] : Automation.SiblingTabNames(tab).Except(before).ToArray();
            }, token);
            if (added.Length == 1) return added[0];
            if (added.Length > 1) throw new InvalidOperationException("More than one new sheet tab appeared; replay cannot identify the copy safely.");
            await Task.Delay(100, token);
        }
        throw new InvalidOperationException("A new sheet tab was not observed after the copy dialog closed.");
    }

    private static bool IsExcelSheetTab(ControlRef target) =>
        target.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true &&
        target.ControlType == "ControlType.TabItem";

    private static bool IsTaskbarNavigation(PlanStep step) =>
        step.Action == "click" && step.Target is { Process: "explorer", ControlType: "ControlType.Button" } target &&
        Regex.IsMatch(target.Name ?? "", @" - \d+ running windows?$", RegexOptions.IgnoreCase);

    private static AutomationElement? WindowAncestor(AutomationElement element)
    {
        var current = element;
        for (var depth = 0; depth < 12; depth++)
        {
            if (current.Current.ControlType == ControlType.Window) return current;
            var parent = TreeWalker.RawViewWalker.GetParent(current);
            if (parent is null || parent == AutomationElement.RootElement) break;
            current = parent;
        }
        return null;
    }

    private static bool ReplaceTabIsOpen(AutomationElement dialog)
    {
        try
        {
            var replaceAll = dialog.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                new PropertyCondition(AutomationElement.NameProperty, "Replace All")));
            if (replaceAll is not null && !replaceAll.Current.IsOffscreen) return true;
            // Some legacy dialogs expose their editable fields but not their
            // command buttons through UIA. The second field is present only on
            // the Replace tab and is the same field used by the recorded plan.
            var replacementField = dialog.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "21"));
            return replacementField is not null && !replacementField.Current.IsOffscreen;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        { System.Diagnostics.Trace.WriteLine(ex); return false; }
    }

    private static async Task EnsureReplaceTabOpenAsync(AutomationElement dialog, int stepNumber,
        CancellationToken token)
    {
        if (ReplaceTabIsOpen(dialog)) return;
        var replaceTab = dialog.FindFirst(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
            new PropertyCondition(AutomationElement.NameProperty, "Replace")));
        if (replaceTab is null)
        {
            if (!MsaaActions.TryInvokeUnique(new nint(dialog.Current.NativeWindowHandle), "Replace", 0x25))
                throw new InvalidOperationException($"Step {stepNumber}: Replace tab was not exposed by UI Automation or MSAA.");
            await Task.Delay(150, token);
            if (!ReplaceTabIsOpen(dialog))
                throw new InvalidOperationException($"Step {stepNumber}: Replace tab did not open.");
            return;
        }
        try
        {
            if (replaceTab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection))
                ((SelectionItemPattern)selection).Select();
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.Runtime.InteropServices.COMException)
        { System.Diagnostics.Trace.WriteLine(ex); }
        if (!ReplaceTabIsOpen(dialog))
        {
            TryClickLiveUiaElement(replaceTab, null, rightClick: false, out _);
            await Task.Delay(150, token);
        }
        if (!ReplaceTabIsOpen(dialog))
        {
            SendKeys.SendWait("%p");
            await Task.Delay(150, token);
        }
        if (!ReplaceTabIsOpen(dialog))
            throw new InvalidOperationException($"Step {stepNumber}: Replace tab did not open.");
    }

    private static bool MalformedDatedSheetName(string value)
    {
        var token = value.IndexOf("{{next:", StringComparison.OrdinalIgnoreCase);
        if (token < 0) return false;
        if (Regex.IsMatch(value[..token], @"\(\)\s*$")) return true;
        return Regex.IsMatch(value[..token],
            @"\d{1,4}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}\b",
            RegexOptions.IgnoreCase);
    }

    private static string? SelectedSheetName(AutomationElement? sheetTabs)
    {
        if (sheetTabs is null) return null;
        try
        {
            var child = TreeWalker.ControlViewWalker.GetFirstChild(sheetTabs);
            while (child is not null)
            {
                if (child.Current.AutomationId == "SheetTab" &&
                    child.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection) &&
                    ((SelectionItemPattern)selection).Current.IsSelected)
                    return child.Current.Name;
                child = TreeWalker.ControlViewWalker.GetNextSibling(child);
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        return null;
    }

    private static void SelectSheet(AutomationElement sheet)
    {
        if (sheet.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection))
            ((SelectionItemPattern)selection).Select();
        else if (sheet.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
            ((InvokePattern)invoke).Invoke();
        else throw new InvalidOperationException("The sheet tab cannot be selected through UI Automation.");
    }

    private static async Task<bool> WaitForWindowAsync(string name, CancellationToken token)
    {
        for (var attempt = 0; attempt < 15; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var found = await Task.Run(() => AutomationElement.RootElement.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.NameProperty, name)) is not null, token);
            if (found) return true;
            await Task.Delay(100, token);
        }
        return false;
    }

    private static async Task<bool> DependentControlReadyAsync(ControlRef? target, CancellationToken token)
    {
        if (target is null) return true;
        // The next edit's recorded Name can be its future value. Locate it by its
        // stable automation ID before that value has been entered.
        var probe = target.ControlType == "ControlType.Edit" &&
            !string.IsNullOrWhiteSpace(target.AutomationId)
            ? target with { Name = null } : target;
        for (var attempt = 0; attempt < 12; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var element = await FindAsync(probe, token, 1);
            try
            {
                if (element is not null && element.Current.IsEnabled && !element.Current.IsOffscreen)
                    return true;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
            { Trace.WriteLine(ex); }
            await Task.Delay(150, token);
        }
        return false;
    }

    private static async Task<AutomationElement?> FindAsync(ControlRef target, CancellationToken token, int attempts = 20)
    {
        for (var i = 0; i < attempts; i++)
        {
            token.ThrowIfCancellationRequested();
            var element = await Task.Run(() => target.ControlType == "ControlType.MenuItem"
                ? Automation.ResolveMenuItem(target) : Automation.Resolve(target), token);
            if (element is not null) return element;
            await Task.Delay(100, token);
        }
        return null;
    }

    private static bool FollowsSubmittedNavigation(IReadOnlyList<PlanStep> steps, int index)
    {
        var current = steps[index].Target;
        if (current is not { Process: { Length: > 0 } }) return false;
        for (var previousIndex = index - 1; previousIndex >= Math.Max(0, index - 16);
            previousIndex--)
        {
            var previous = steps[previousIndex];
            if (previous.Target?.Process?.Equals(current.Process,
                    StringComparison.OrdinalIgnoreCase) != true)
                return false;
            if (previous.Action == "key" && previous.Key is "Enter" or "Return" &&
                previous.Target.ControlType == "ControlType.Edit")
                return true;
            if (previous.Action is not ("type" or "key") ||
                previous.Target.ControlType != "ControlType.Pane" ||
                !string.IsNullOrWhiteSpace(previous.Target.Name) ||
                !string.IsNullOrWhiteSpace(previous.Target.AutomationId) ||
                previous.Target.Window != current.Window)
                return false;
        }
        return false;
    }

    private static bool ForegroundProcessMatches(string? processName)
    {
        if (string.IsNullOrWhiteSpace(processName)) return false;
        var foreground = GetForegroundWindow();
        if (foreground == 0) return false;
        GetWindowThreadProcessId(foreground, out var processId);
        try
        {
            return processId != 0 && Process.GetProcessById((int)processId).ProcessName.Equals(
                processName, StringComparison.OrdinalIgnoreCase);
        }
        catch (ArgumentException) { return false; }
    }

    private static bool BrowserPageTitleMatches(string? recorded, string live)
    {
        if (string.IsNullOrWhiteSpace(recorded) || string.IsNullOrWhiteSpace(live)) return false;
        static string Page(string title) => Regex.Replace(title.Split(" - ")[0],
            @" and \d+ more pages?", "", RegexOptions.IgnoreCase).Trim();
        return Page(recorded).Equals(Page(live), StringComparison.OrdinalIgnoreCase);
    }

    private static bool FollowsSubmittedEdit(IReadOnlyList<PlanStep> steps, int index,
        out ControlRef submittedField)
    {
        submittedField = null!;
        var current = steps[index].Target;
        if (current is not { Process: { Length: > 0 }, AutomationId: { Length: > 0 } })
            return false;
        for (var previousIndex = index - 1; previousIndex >= Math.Max(0, index - 8);
            previousIndex--)
        {
            var previous = steps[previousIndex];
            if (previous.Target?.Process?.Equals(current.Process,
                    StringComparison.OrdinalIgnoreCase) != true ||
                previous.Target.AutomationId != current.AutomationId ||
                previous.Target.ControlType != current.ControlType)
                return false;
            if (previous.Action == "key" && previous.Key is "Enter" or "Return")
            {
                submittedField = previous.Target;
                return true;
            }
            if (previous.Action is not ("type" or "key")) return false;
        }
        return false;
    }

    private static async Task<bool> RowMenuPopupVisibleAsync(nint workbook, CancellationToken token)
    {
        var thread = GetWindowThreadProcessId(workbook, out _);
        if (thread == 0) return false;
        for (var attempt = 0; attempt < 8; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var info = new GuiThreadInfo { Size = Marshal.SizeOf<GuiThreadInfo>() };
            if (GetGUIThreadInfo(thread, ref info) && info.MenuOwner != 0 &&
                IsWindowVisible(info.MenuOwner)) return true;
            await Task.Delay(75, token);
        }
        return false;
    }

    private static async Task<AutomationElement?> WaitForMenuAsync(ControlRef target, Action<string> status,
        CancellationToken token)
    {
        var watch = Stopwatch.StartNew();
        var nextUpdate = 5;
        while (watch.Elapsed < TimeSpan.FromSeconds(30))
        {
            token.ThrowIfCancellationRequested();
            var item = await Task.Run(() => Automation.ResolveMenuItem(target), token);
            if (item is not null) return item;
            if (watch.Elapsed.TotalSeconds >= nextUpdate)
            {
                status($"Still waiting for the selected items' context menu ({(int)watch.Elapsed.TotalSeconds}s).");
                nextUpdate += 5;
            }
            await Task.Delay(250, token);
        }
        return null;
    }

    private static bool FollowsHorizontalScroll(ExecutionPlan plan, int index) =>
        index > 0 && plan.Steps[index - 1].Action == "scroll" &&
        plan.Steps[index - 1].Key == "Horizontal";

    private static async Task<(AutomationElement? Element, bool HandledByUser)> RevealTargetWithUiaScrollAsync(PlanStep step,
        Action<string> status, Func<string, bool> handleUnexpectedDialog, CancellationToken token)
    {
        var target = step.Target!;
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            token.ThrowIfCancellationRequested();
            status($"Step {step.Number}: {target.Name} is not visible after horizontal scrolling; UI Automation retry {attempt} of 3.");
            var bar = await Task.Run(() => Automation.FindScrollBar(target, "Horizontal"), token);
            if (bar is null) break;
            for (var movement = 0; movement < 8; movement++)
            {
                token.ThrowIfCancellationRequested();
                if (!TryScrollUiAutomation(bar, movement < 6 ? ScrollAmount.SmallIncrement : ScrollAmount.LargeIncrement))
                    break;
                await Task.Delay(120, token);
                if (await FindAsync(target, token, 1) is { } revealed) return (revealed, false);
            }
        }
        var help = $"Replay needs help: step {step.Number} could not find '{target.Name}' after 3 UI Automation scroll retries. " +
            "In the target program, scroll until that control is visible, then choose I've handled it. " +
            "Choose Stop replay to leave the target program as it is.";
        status(help);
        if (!handleUnexpectedDialog(help)) throw new OperationCanceledException("Replay stopped by user.", token);
        return (null, true);
    }

    private static bool TryScrollUiAutomation(AutomationElement bar, ScrollAmount amount)
    {
        try
        {
            // Prefer the scrollbar's exposed increment button. Some providers
            // update RangeValue without moving the content viewport.
            var buttons = bar.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            var rightmost = buttons.Cast<AutomationElement>()
                .Where(button => button.Current.IsEnabled && !button.Current.IsOffscreen)
                .OrderByDescending(button => button.Current.BoundingRectangle.Right).FirstOrDefault();
            if (rightmost is not null)
            {
                if (rightmost.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
                {
                    ((InvokePattern)invoke).Invoke();
                    return true;
                }
                if (TryClickLiveUiaElement(rightmost, null, false, out _)) return true;
            }
            var container = Automation.ScrollableAncestor(bar);
            if (container.TryGetCurrentPattern(ScrollPattern.Pattern, out var scrollObject))
            {
                ((ScrollPattern)scrollObject).ScrollHorizontal(amount);
                return true;
            }
            if (bar.TryGetCurrentPattern(RangeValuePattern.Pattern, out var rangeObject))
            {
                var range = (RangeValuePattern)rangeObject;
                if (range.Current.IsReadOnly || range.Current.Value >= range.Current.Maximum) return false;
                var change = amount == ScrollAmount.LargeIncrement ? range.Current.LargeChange : range.Current.SmallChange;
                range.SetValue(Math.Min(range.Current.Maximum, range.Current.Value + Math.Max(change, 1)));
                return true;
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or ElementNotEnabledException or
            InvalidOperationException or COMException) { Trace.WriteLine(ex); }
        return false;
    }

    private static async Task<string> ClickDataItemWithRetriesAsync(PlanStep step,
        AutomationElement firstElement, ControlRef? anchor, Action<string> status,
        Func<string, bool> handleUnexpectedDialog, CancellationToken token)
    {
        var detail = "the control was unavailable";
        for (var attempt = 0; attempt <= 3; attempt++)
        {
            token.ThrowIfCancellationRequested();
            if (attempt > 0)
            {
                status($"Step {step.Number}: reacquiring {step.Target!.Name}, retry {attempt} of 3.");
                await Task.Delay(150, token);
            }
            var current = attempt == 0 ? firstElement : await FindAsync(step.Target!, token, 2);
            if (current is null) { detail = "the control is no longer exposed by UI Automation"; continue; }
            try
            {
                if (current.Current.IsOffscreen &&
                    current.TryGetCurrentPattern(ScrollItemPattern.Pattern, out var scrollItem))
                {
                    ((ScrollItemPattern)scrollItem).ScrollIntoView();
                    current = await FindAsync(step.Target!, token, 2) ?? current;
                }
                if (TryClickLiveUiaElement(current, anchor, false, out detail)) return detail;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ElementNotEnabledException or
                InvalidOperationException or COMException)
            { detail = $"UI Automation reported {ex.GetType().Name}: {ex.Message}"; }
        }
        var help = $"Replay needs help: step {step.Number} could not click '{step.Target!.Name}' after 3 retries. " +
            $"{detail} Bring that control into view in the target program and click it, then choose I've handled it. " +
            "Choose Stop replay to leave the target program as it is.";
        status(help);
        if (!handleUnexpectedDialog(help)) throw new OperationCanceledException("Replay stopped by user.", token);
        return "Handled by the user after retries.";
    }

    private static async Task<string> ClickTransientFilterControlAsync(ExecutionPlan plan, int index,
        AutomationElement firstElement, Action<string> status, Func<string, bool> handleUnexpectedDialog,
        CancellationToken token)
    {
        var step = plan.Steps[index];
        var target = step.Target!;
        var detail = "the filter control was unavailable";
        for (var attempt = 0; attempt <= 3; attempt++)
        {
            token.ThrowIfCancellationRequested();
            if (attempt > 0)
            {
                status($"Step {step.Number}: retry {attempt} of 3 for {target.Name}.");
                await Task.Delay(250, token);
            }
            try
            {
                var current = attempt == 0 ? firstElement : await FindAsync(target, token, 2);
                if (current is not null && TryClickLiveUiaElement(current, null, false, out detail))
                    return detail;
                if (current is null) detail = "the filter control disappeared from UI Automation";
            }
            catch (Exception ex) when (ex is ElementNotEnabledException or ElementNotAvailableException or
                InvalidOperationException or COMException)
            {
                detail = $"UI Automation reported {ex.GetType().Name}: {ex.Message}";
            }
            // A failed attempt can leave the transient menu closed. Reopen only
            // when the recorded item is no longer exposed, using its recorded column.
            if (attempt < 3 && await FindAsync(target, token, 1) is null)
            {
                var dropdown = plan.Steps.Take(index).LastOrDefault(previous =>
                    previous.Target is { ControlType: "ControlType.MenuItem", AutomationId: "Dropdown" })?.Target;
                if (dropdown is not null && await FindAsync(dropdown, token, 2) is { } button)
                    TryClickLiveUiaElement(button, null, false, out _);
            }
        }
        var help = $"Replay needs help: step {step.Number} could not click '{target.Name}' after 3 retries. " +
            $"{detail}\nIn the open filter, perform this one click yourself, then choose I've handled it. " +
            "Choose Stop replay to leave the worksheet as it is.";
        status(help.Replace('\n', ' '));
        if (!handleUnexpectedDialog(help)) throw new OperationCanceledException("Replay stopped by user.", token);
        return "Handled by the user after retries.";
    }

    private static void Act(AutomationElement element, PlanStep step, bool preserveFocus = false)
    {
        if (step.Action is "key" or "type")
        {
            if (!preserveFocus) FocusAndVerifyKeyboardTarget(element, step);
            if (step.Action == "type")
            {
                if (step.Value is null) throw new InvalidOperationException("Missing text.");
                SendKeys.SendWait(string.Concat(DynamicText.Resolve(step.Value, step.RelativeWeekday).Select(EscapeText)));
                return;
            }
            if (step.Key is null) throw new InvalidOperationException("Missing key.");
            if (step.Key is "Control+V" or "Ctrl+V")
            {
                if (step.Target?.ControlType == "ControlType.DataItem")
                {
                    SendKeys.SendWait("{ESC}"); // Leave cell edit mode before a multi-cell paste.
                    if (!preserveFocus && element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var cellSelection))
                        ((SelectionItemPattern)cellSelection).Select();
                }
                keybd_event(0x11, 0, 0, 0); // Ctrl down.
                keybd_event(0x56, 0, 0, 0); // V down.
                keybd_event(0x56, 0, 2, 0); // V up.
                keybd_event(0x11, 0, 2, 0); // Ctrl up.
                return;
            }
            SendKeys.SendWait(ToSendKeys(step.Key));
            return;
        }
        if (step.Action == "context-click")
        {
            OpenContextMenu(element, preserveFocus);
            return;
        }
        if (step.Action == "click" && step.Target?.ControlType == "ControlType.DataItem")
        {
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var item))
                ((SelectionItemPattern)item).Select();
            else element.SetFocus();
            return;
        }
        var actionable = Automation.ActionableAncestor(element) ?? element;
        if (actionable.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
            ((InvokePattern)invoke).Invoke();
        else if (actionable.TryGetCurrentPattern(TogglePattern.Pattern, out var toggle))
            ((TogglePattern)toggle).Toggle();
        else if (actionable.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection))
            ((SelectionItemPattern)selection).Select();
        else if (actionable.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var expandObject))
        {
            var expand = (ExpandCollapsePattern)expandObject;
            try
            {
                if (expand.Current.ExpandCollapseState == ExpandCollapseState.Collapsed)
                {
                    expand.Expand();
                    return;
                }
                if (expand.Current.ExpandCollapseState == ExpandCollapseState.Expanded)
                    return;
            }
            catch (Exception ex) when (ex is InvalidOperationException or COMException)
            { Trace.WriteLine(ex); }
            var detail = "the recorded action is not a click";
            if (step.Action != "click" ||
                !TryClickLiveUiaElement(actionable, null, rightClick: false, out detail))
                throw new InvalidOperationException($"Step {step.Number}: '{step.Target?.Name}' " +
                    $"exposed an unusable ExpandCollapse state and could not be clicked. {detail}");
        }
        else throw new InvalidOperationException($"Step {step.Number}: target offers no safe UI Automation action.");
    }

    private static bool OpenRecordedComboBox(AutomationElement element, PlanStep step)
    {
        try
        {
            if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var pattern))
            {
                var expand = (ExpandCollapsePattern)pattern;
                if (expand.Current.ExpandCollapseState == ExpandCollapseState.Expanded) return true;
                if (expand.Current.ExpandCollapseState == ExpandCollapseState.Collapsed)
                {
                    expand.Expand();
                    return true;
                }
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or COMException)
        { Trace.WriteLine(ex); }
        if (step.Target?.ClassName != "ComboBox" || element.Current.ClassName != "ComboBox")
            return false;
        var handle = (nint)element.Current.NativeWindowHandle;
        if (handle == 0 || SendMessageTimeout(handle, 0x014F, 1, 0,
            0x0002, 500, out _) == 0) return false; // CB_SHOWDROPDOWN
        return SendMessageTimeout(handle, 0x0157, 0, 0,
            0x0002, 500, out var dropped) != 0 && dropped != 0; // CB_GETDROPPEDSTATE
    }

    private static void FocusAndVerifyKeyboardTarget(AutomationElement element, PlanStep step)
    {
        var target = step.Target ?? throw new InvalidOperationException($"Step {step.Number}: keyboard target is missing.");
        if (target.ControlType == "ControlType.DataItem")
        {
            if (element.Current.IsOffscreen &&
                element.TryGetCurrentPattern(ScrollItemPattern.Pattern, out var scrollItem))
            {
                ((ScrollItemPattern)scrollItem).ScrollIntoView();
                element = Automation.Resolve(target) ??
                    throw new InvalidOperationException($"Step {step.Number}: keyboard input was not sent because the target disappeared after scrolling into view.");
            }
            if (!TryClickLiveUiaElement(element, null, false, out var detail))
                throw new InvalidOperationException($"Step {step.Number}: keyboard input was not sent because the recorded target could not be clicked. {detail}");
            Thread.Sleep(80);
            var live = Automation.Resolve(target) ??
                throw new InvalidOperationException($"Step {step.Number}: keyboard input was not sent because the target disappeared after selection.");
            if (live.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection) &&
                !((SelectionItemPattern)selection).Current.IsSelected)
                throw new InvalidOperationException($"Step {step.Number}: keyboard input was not sent because '{target.Name}' did not become selected.");
            var focused = Automation.Describe(AutomationElement.FocusedElement);
            if (focused?.ControlType == "ControlType.DataItem" && focused.Name != target.Name)
                throw new InvalidOperationException($"Step {step.Number}: keyboard input was not sent because focus remained on '{focused.Name}' instead of '{target.Name}'.");
            return;
        }
        element.SetFocus();
        var current = AutomationElement.FocusedElement;
        for (var depth = 0; current is not null && depth < 10; depth++)
        {
            if (current.GetRuntimeId().SequenceEqual(element.GetRuntimeId())) return;
            current = TreeWalker.RawViewWalker.GetParent(current);
        }
        throw new InvalidOperationException($"Step {step.Number}: keyboard input was not sent because UI Automation did not confirm focus on '{target.Name}'.");
    }

    private static void SetRecordedEditField(AutomationElement element, PlanStep step)
    {
        var desired = step.Value ?? step.Target?.Name ??
            throw new InvalidOperationException($"Step {step.Number}: the recorded edit field has no value.");
        if (RecordedEditFieldMatches(element, desired)) return;
        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var valueObject))
        {
            var value = (ValuePattern)valueObject;
            if (!value.Current.IsReadOnly)
            {
                try
                {
                    value.SetValue(desired);
                    if (RecordedEditFieldMatches(element, desired)) return;
                }
                catch (Exception ex) when (ex is InvalidOperationException or System.Runtime.InteropServices.COMException)
                { System.Diagnostics.Trace.WriteLine(ex); }
            }
        }
        var handle = (nint)element.Current.NativeWindowHandle;
        if (handle != 0 && step.Target?.ClassName == "Edit")
        {
            var text = new StringBuilder(desired);
            if (SendMessageTimeout(handle, 0x000C, 0, text, 0x0002, 500, out _) != 0 &&
                RecordedEditFieldMatches(element, desired)) return;
        }
        if (!TryClickLiveUiaElement(element, null, rightClick: false, out var detail))
            throw new InvalidOperationException($"Step {step.Number}: edit field could not be focused through live UI Automation. {detail}");
        SendKeys.SendWait("^a");
        SendKeys.SendWait(string.Concat(desired.Select(EscapeText)));
        if (!RecordedEditFieldMatches(element, desired))
            throw new InvalidOperationException($"Step {step.Number}: the edit field did not show the recorded value after input.");
    }

    private static bool RecordedEditFieldMatches(AutomationElement element, string desired)
    {
        Thread.Sleep(80);
        if (element.Current.Name == desired) return true;
        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out var value) &&
            ((ValuePattern)value).Current.Value == desired) return true;
        var handle = (nint)element.Current.NativeWindowHandle;
        if (handle == 0) return false;
        var text = new StringBuilder(Math.Max(256, desired.Length + 1));
        return SendMessageTimeout(handle, 0x000D, (nint)text.Capacity, text,
            0x0002, 500, out _) != 0 && text.ToString() == desired;
    }

    private static string? DesiredSortDirection(string? expected) =>
        expected?.Contains("descending", StringComparison.OrdinalIgnoreCase) == true ? "descending" :
        expected?.Contains("ascending", StringComparison.OrdinalIgnoreCase) == true ? "ascending" : null;

    private static async Task<bool> EnsureSortDirectionAsync(AutomationElement header, PlanStep step,
        string desired, CancellationToken token)
    {
        var clicked = false;
        for (var attempt = 0; attempt < 3; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var liveHeader = await Task.Run(() => Automation.Resolve(step.Target!) ?? header, token);
            var observed = await Task.Run(() => ObservedSortDirection(liveHeader), token);
            if (observed == desired) return clicked;
            if (attempt == 2)
                throw new InvalidOperationException($"Step {step.Number}: could not verify {desired} order from visible column values.");
            Act(liveHeader, step with { Action = "click" });
            clicked = true;
            await Task.Delay(150, token);
        }
        return clicked;
    }

    private static void SendWindowContextMenu(AutomationElement element)
    {
        var current = element;
        for (var depth = 0; depth < 8 && current is not null; depth++)
        {
            var handle = (nint)current.Current.NativeWindowHandle;
            if (handle != 0) { SendMessage(handle, 0x007B, handle, -1); return; }
            current = TreeWalker.ControlViewWalker.GetParent(current);
        }
    }

    private static bool SendFocusedContextMenu(string? expectedProcess)
    {
        var foreground = GetForegroundWindow();
        var thread = GetWindowThreadProcessId(foreground, out var processId);
        if (thread == 0 || processId == 0 || !Process.GetProcessById((int)processId).ProcessName
            .Equals(expectedProcess, StringComparison.OrdinalIgnoreCase)) return false;
        var info = new GuiThreadInfo { Size = Marshal.SizeOf<GuiThreadInfo>() };
        if (!GetGUIThreadInfo(thread, ref info) || info.Focus == 0) return false;
        SendMessage(info.Focus, 0x007B, info.Focus, -1); // WM_CONTEXTMENU, keyboard invocation.
        return true;
    }

    private static string ContextFocusDiagnostics(AutomationElement? list)
    {
        try
        {
            var foreground = GetForegroundWindow();
            var thread = GetWindowThreadProcessId(foreground, out _);
            var info = new GuiThreadInfo { Size = Marshal.SizeOf<GuiThreadInfo>() };
            var focus = thread != 0 && GetGUIThreadInfo(thread, ref info) ? info.Focus : 0;
            var listHandle = list?.Current.NativeWindowHandle ?? 0;
            var focusedItem = listHandle != 0 && list?.Current.ClassName.Contains("SysListView32", StringComparison.OrdinalIgnoreCase) == true
                ? (int)SendMessage((nint)listHandle, 0x100C, -1, 1) : -1;
            return $"list HWND={listHandle}; focused HWND={focus}; focused list item index={focusedItem}.";
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException)
        { return "Focus details unavailable."; }
    }

    private static bool TryRightClickSelectedItemAtLiveUiaBounds(AutomationElement list, out string detail)
    {
        detail = "no visible selected row passed hit testing";
        try
        {
            var listInfo = list.Current;
            if (listInfo.NativeWindowHandle == 0 || listInfo.IsOffscreen) return false;
            var rows = list.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
            foreach (AutomationElement row in rows)
            {
                if (row.Current.IsOffscreen ||
                    !row.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectedObject) ||
                    !((SelectionItemPattern)selectedObject).Current.IsSelected) continue;
                var visible = System.Windows.Rect.Intersect(row.Current.BoundingRectangle, listInfo.BoundingRectangle);
                if (visible.IsEmpty || visible.Width < 8 || visible.Height < 8) continue;
                var x = (int)Math.Round(visible.Left + visible.Width / 2);
                var y = (int)Math.Round(visible.Top + visible.Height / 2);
                if (!Screen.AllScreens.Any(screen => screen.Bounds.Contains(x, y))) continue;
                var hit = AutomationElement.FromPoint(new System.Windows.Point(x, y));
                var hitRow = hit is null ? null : Automation.ListItemAncestor(hit);
                if (hitRow?.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var hitObject) != true ||
                    !((SelectionItemPattern)hitObject).Current.IsSelected ||
                    ((SelectionItemPattern)hitObject).Current.SelectionContainer?.Current.NativeWindowHandle != listInfo.NativeWindowHandle)
                    continue;
                if (!SetCursorPos(x, y) || !GetCursorPos(out var cursor) || cursor.X != x || cursor.Y != y)
                {
                    detail = "Windows did not move the pointer to the UIA-verified row";
                    return false;
                }
                mouse_event(0x0008, 0, 0, 0, 0); // RIGHTDOWN at the live UIA position.
                mouse_event(0x0010, 0, 0, 0, 0); // RIGHTUP.
                detail = "the row was selected and hit-tested inside its current list";
                return true;
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        {
            detail = $"UI Automation could not validate the selected row ({ex.GetType().Name})";
        }
        return false;
    }

    private static ControlRef? PreviousDataItemAnchor(ExecutionPlan plan, int index, ControlRef target)
    {
        for (var previousIndex = index - 1; previousIndex >= Math.Max(0, index - 5); previousIndex--)
        {
            var candidate = plan.Steps[previousIndex].Target;
            if (candidate?.ControlType == "ControlType.DataItem" && candidate != target &&
                candidate.Process == target.Process && candidate.Window == target.Window)
                return candidate;
        }
        return null;
    }

    private static bool TryClickLiveUiaElement(AutomationElement element, ControlRef? precedingCell,
        bool rightClick, out string detail)
    {
        detail = "the target was not hit-testable";
        try
        {
            var bounds = element.Current.BoundingRectangle;
            var targetId = element.GetRuntimeId();
            var candidates = new List<(int X, int Y)>();
            // The provider's clickable point is tied to this live UIA element and
            // remains useful when Excel reports an empty row-header rectangle.
            if (element.TryGetClickablePoint(out var clickablePoint))
                candidates.Add(((int)Math.Round(clickablePoint.X), (int)Math.Round(clickablePoint.Y)));
            if (!element.Current.IsOffscreen && !bounds.IsEmpty && bounds.Width >= 4 && bounds.Height >= 4)
                candidates.Add(((int)Math.Round(bounds.Left + bounds.Width / 2),
                    (int)Math.Round(bounds.Top + bounds.Height / 2)));
            // Some grid providers expose a row header with an empty rectangle. Use the
            // preceding cell's live row position, then hit-test leftward for this header.
            if (precedingCell?.ControlType == "ControlType.DataItem" &&
                (bounds.IsEmpty || bounds.Width < 4 || bounds.Height < 4) &&
                Automation.Resolve(precedingCell) is { } cell)
            {
                var cellBounds = cell.Current.BoundingRectangle;
                if (!cellBounds.IsEmpty && cellBounds.Width >= 4 && cellBounds.Height >= 4)
                {
                    var y = (int)Math.Round(cellBounds.Top + cellBounds.Height / 2);
                    var screen = Screen.FromPoint(new System.Drawing.Point((int)cellBounds.Left, y));
                    for (var x = (int)cellBounds.Left - 8; x >= screen.Bounds.Left + 4; x -= 12)
                        candidates.Add((x, y));
                }
            }
            string? lastHit = null;
            foreach (var (x, y) in candidates.Distinct())
            {
                if (!Screen.AllScreens.Any(screen => screen.Bounds.Contains(x, y))) continue;
                var hit = AutomationElement.FromPoint(new System.Windows.Point(x, y));
                var matched = false;
                for (var depth = 0; hit is not null && depth < 12; depth++)
                {
                    if (depth == 0)
                        lastHit = $"hit type={hit.Current.ControlType.ProgrammaticName}, class={hit.Current.ClassName}, name={hit.Current.Name}";
                    if (hit.GetRuntimeId().SequenceEqual(targetId) ||
                        (hit.Current.ProcessId == element.Current.ProcessId &&
                         hit.Current.ControlType == element.Current.ControlType &&
                         hit.Current.ClassName == element.Current.ClassName &&
                         hit.Current.Name == element.Current.Name))
                    { matched = true; break; }
                    hit = TreeWalker.RawViewWalker.GetParent(hit);
                }
                // UIA's own clickable point is a semantic guarantee from the
                // provider even when its element-from-point tree differs.
                if (!matched && candidates[0] == (x, y) &&
                    element.TryGetClickablePoint(out var currentClickablePoint) &&
                    (int)Math.Round(currentClickablePoint.X) == x &&
                    (int)Math.Round(currentClickablePoint.Y) == y)
                    matched = true;
                if (!matched) continue;
                if (!SetCursorPos(x, y) || !GetCursorPos(out var cursor) || cursor.X != x || cursor.Y != y)
                {
                    detail = "Windows did not move the pointer to the UIA-verified target";
                    return false;
                }
                mouse_event(rightClick ? 0x0008u : 0x0002u, 0, 0, 0, 0);
                Thread.Sleep(45); // Let controls process the pressed state before release.
                mouse_event(rightClick ? 0x0010u : 0x0004u, 0, 0, 0, 0);
                detail = "the target was verified by live UIA hit testing";
                return true;
            }
            detail = candidates.Count == 0 ? "the target and preceding cell have no usable live bounds"
                : $"no point on the row header matched its UIA identity; target bounds={bounds}; {lastHit ?? "no UIA hit"}";
            return false;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        {
            detail = $"UI Automation could not validate the target ({ex.GetType().Name})";
            return false;
        }
    }

    private static AutomationElement? CurrentSelectedRow(string? process)
    {
        try
        {
            var focused = AutomationElement.FocusedElement;
            return focused is not null && Automation.Describe(focused)?.Process == process
                ? Automation.ListItemAncestor(focused) : null;
        }
        catch (ElementNotAvailableException) { return null; }
    }

    private static bool AllListItemsSelected(AutomationElement list)
    {
        try
        {
            var info = list.Current;
            if (info.ClassName.Contains("SysListView32", StringComparison.OrdinalIgnoreCase) &&
                info.NativeWindowHandle != 0)
            {
                var handle = (nint)info.NativeWindowHandle;
                var total = (int)SendMessage(handle, 0x1004, 0, 0); // LVM_GETITEMCOUNT
                var selected = (int)SendMessage(handle, 0x1032, 0, 0); // LVM_GETSELECTEDCOUNT
                if (total > 0) return selected == total;
            }
            if (!list.TryGetCurrentPattern(SelectionPattern.Pattern, out var pattern)) return false;
            var selectedItems = ((SelectionPattern)pattern).Current.GetSelection();
            var items = list.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
            return items.Count > 0 && selectedItems.Length == items.Count;
        }
        catch (Exception e) when (e is ElementNotAvailableException or InvalidOperationException) { return false; }
    }

    private static void SelectAllListItemsViaUia(AutomationElement list, CancellationToken token)
    {
        var items = list.FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
        foreach (AutomationElement item in items)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                if (item.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern) &&
                    !((SelectionItemPattern)pattern).Current.IsSelected)
                    ((SelectionItemPattern)pattern).AddToSelection();
            }
            catch (Exception e) when (e is ElementNotAvailableException or InvalidOperationException) { }
        }
    }

    private static string ListSelectionDiagnostics(AutomationElement list)
    {
        try
        {
            var count = list.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem)).Count;
            var selected = list.TryGetCurrentPattern(SelectionPattern.Pattern, out var pattern)
                ? ((SelectionPattern)pattern).Current.GetSelection().Length : -1;
            return $"Visible rows: {count}; selected rows reported by UIA: {selected}; class: {list.Current.ClassName}.";
        }
        catch (Exception e) when (e is ElementNotAvailableException or InvalidOperationException) { return "The list became unavailable."; }
    }

    internal static string? ObservedSortDirection(AutomationElement header)
    {
        var match = Regex.Match(header.Current.AutomationId, @"\d+$");
        if (!match.Success) return null;
        var cellId = "ListViewSubItem-" + match.Value;
        var ancestor = TreeWalker.ControlViewWalker.GetParent(header);
        for (var level = 0; level < 5 && ancestor is not null; level++)
        {
            var rows = ancestor.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
            var values = new List<string>();
            foreach (AutomationElement row in rows.Cast<AutomationElement>().Take(16))
            {
                var cell = row.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.AutomationIdProperty, cellId));
                var value = cell?.Current.Name;
                if (!string.IsNullOrWhiteSpace(value)) values.Add(value);
            }
            if (values.Count >= 2)
            {
                var direction = 0;
                for (var i = 1; i < values.Count; i++)
                {
                    var comparison = CompareSortValues(values[i - 1], values[i]);
                    if (comparison == 0) continue;
                    var nextDirection = comparison < 0 ? 1 : -1;
                    if (direction != 0 && direction != nextDirection) return null;
                    direction = nextDirection;
                }
                if (direction != 0) return direction > 0 ? "ascending" : "descending";
            }
            ancestor = TreeWalker.ControlViewWalker.GetParent(ancestor);
        }
        return null;
    }

    private static int CompareSortValues(string left, string right)
    {
        static double? Duration(string value)
        {
            var matches = Regex.Matches(value, @"(\d+)\s*(day|hour|minute|second)s?", RegexOptions.IgnoreCase);
            if (matches.Count == 0) return null;
            double total = 0;
            foreach (Match match in matches)
            {
                var number = double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                total += number * (match.Groups[2].Value.ToLowerInvariant() switch
                { "day" => 86400, "hour" => 3600, "minute" => 60, _ => 1 });
            }
            return total;
        }
        var first = Duration(left); var second = Duration(right);
        if (first.HasValue && second.HasValue) return first.Value.CompareTo(second.Value);
        if (double.TryParse(left, out var a) && double.TryParse(right, out var b)) return a.CompareTo(b);
        return string.Compare(left, right, StringComparison.CurrentCultureIgnoreCase);
    }

    private static void ApplyScroll(AutomationElement element, PlanStep step, Action<string> status)
    {
        if (step.ExpectedState?.StartsWith("range:", StringComparison.Ordinal) == true &&
            element.TryGetCurrentPattern(RangeValuePattern.Pattern, out var rangeObject))
        {
            var range = (RangeValuePattern)rangeObject;
            var value = double.Parse(step.ExpectedState[6..], System.Globalization.CultureInfo.InvariantCulture);
            if (!range.Current.IsReadOnly)
            {
                // A filter, zoom, or window-size change can alter a scrollbar's
                // range since recording. Keep the recorded direction and use a
                // value valid for the live control instead of failing outright.
                var liveValue = Math.Clamp(value, range.Current.Minimum, range.Current.Maximum);
                if (liveValue != value)
                    status($"Step {step.Number}: recorded scroll value {value} is outside the live range; using {liveValue}.");
                range.SetValue(liveValue);
                if (liveValue == range.Current.Minimum &&
                    step.Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true)
                    EnsureExcelViewportAtOrigin(step.Key, status);
                return;
            }
            status($"Step {step.Number}: scrollbar range is read only; trying its UI Automation scroll action.");
        }
        if (element.TryGetCurrentPattern(RangeValuePattern.Pattern, out var wheelRangeObject) &&
            int.TryParse(step.Value, out var wheelDelta))
        {
            var range = (RangeValuePattern)wheelRangeObject;
            var ticks = wheelDelta / 120.0;
            var increment = range.Current.SmallChange > 0 ? range.Current.SmallChange : 1;
            var desired = Math.Clamp(range.Current.Value - ticks * increment,
                range.Current.Minimum, range.Current.Maximum);
            if (range.Current.IsReadOnly) throw new InvalidOperationException($"Step {step.Number}: scrollbar is read only.");
            range.SetValue(desired);
            return;
        }
        var scrollElement = Automation.ScrollableAncestor(element);
        if (scrollElement.TryGetCurrentPattern(ScrollPattern.Pattern, out var scrollObject))
        {
            var scroll = (ScrollPattern)scrollObject;
            if (step.ExpectedState?.StartsWith("scroll:", StringComparison.Ordinal) == true)
            {
                var values = step.ExpectedState[7..].Split(',');
                if (values.Length == 2 &&
                    double.TryParse(values[0], System.Globalization.CultureInfo.InvariantCulture, out var horizontal) &&
                    double.TryParse(values[1], System.Globalization.CultureInfo.InvariantCulture, out var vertical))
                {
                    scroll.SetScrollPercent(horizontal, vertical);
                    return;
                }
            }
            var delta = int.TryParse(step.Value, out var parsed) ? parsed : 0;
            var amount = delta > 0 ? ScrollAmount.SmallDecrement : ScrollAmount.SmallIncrement;
            if (step.Key == "Horizontal") scroll.ScrollHorizontal(amount);
            else scroll.ScrollVertical(amount);
            return;
        }
        throw new InvalidOperationException($"Step {step.Number}: target exposes no UI Automation scroll action.");
    }

    private static void EnsureExcelViewportAtOrigin(string? axis, Action<string> status)
    {
        // Excel's UIA range can change without moving the sheet. With the workbook
        // already foreground, this keyboard command changes its actual viewport.
        if (axis is not ("Vertical" or "Horizontal")) return;
        SendKeys.SendWait("{ESC}^{HOME}");
        status("Excel viewport fallback: moved to the first row and column with Ctrl+Home");
    }

    private static void OpenContextMenu(AutomationElement element, bool preserveFocus = false)
    {
        if (!preserveFocus)
        {
            try { element.SetFocus(); }
            catch (InvalidOperationException)
            {
                var target = Automation.Describe(element);
                var focused = Automation.Describe(AutomationElement.FocusedElement);
                if (target is null || focused?.Process != target.Process || focused?.Window != target.Window)
                    throw;
            }
        }
        SendKeys.SendWait("+{F10}");
    }

    private static void SendContextMenuKey(AutomationElement element, bool preserveFocus = false)
    {
        if (!preserveFocus) element.SetFocus();
        keybd_event(0x5D, 0, 0, 0); // VK_APPS: keyboard context menu key.
        keybd_event(0x5D, 0, 2, 0);
    }

    private static void ActivateWindow(AutomationElement element)
    {
        var current = element;
        while (true)
        {
            var parent = TreeWalker.ControlViewWalker.GetParent(current);
            if (parent is null || parent == AutomationElement.RootElement) break;
            current = parent;
        }
        var handle = (nint)current.Current.NativeWindowHandle;
        if (handle == 0) throw new InvalidOperationException("The target window has no native handle to activate.");
        if (GetForegroundWindow() != handle || GetWindow(handle, 4) != 0)
            ShowWindow(handle, 9); // Restore a minimized target.
        MaximizeWindowIfSupported(current, handle);
        SetForegroundWindow(handle);
        for (var attempt = 0; attempt < 6; attempt++)
        {
            if (GetForegroundWindow() == handle) return;
            Thread.Sleep(50);
        }
        try { element.SetFocus(); } catch (InvalidOperationException) { }
        if (GetForegroundWindow() != handle)
            throw new InvalidOperationException("Windows did not bring the target window to the foreground. Leave the desktop idle and retry.");
    }

    private static void MaximizeTargetWindow(AutomationElement element)
    {
        var current = element;
        while (true)
        {
            var parent = TreeWalker.ControlViewWalker.GetParent(current);
            if (parent is null || parent == AutomationElement.RootElement) break;
            current = parent;
        }
        var handle = (nint)current.Current.NativeWindowHandle;
        if (handle == 0) throw new InvalidOperationException("The target window has no native handle to activate.");
        ShowWindow(handle, 9); // Restore a minimized window first.
        SetForegroundWindow(handle);
        MaximizeWindowIfSupported(current, handle);
        ActivateWindow(element);
    }

    private static void MaximizeWindowIfSupported(AutomationElement window, nint handle)
    {
        if (GetWindow(handle, 4) != 0) return; // Keep owned dialogs movable so their parent remains usable.
        if (window.TryGetCurrentPattern(WindowPattern.Pattern, out var pattern) &&
            ((WindowPattern)pattern).Current.CanMaximize)
            ((WindowPattern)pattern).SetWindowVisualState(WindowVisualState.Maximized);
        MaximizeWindowIfSupported(handle);
        // Dialogs and fixed-size windows are simply brought forward.
    }

    private static void MaximizeWindowIfSupported(nint handle)
    {
        if (GetWindow(handle, 4) != 0) return;
        if ((GetWindowLong(handle, -16) & 0x00010000) != 0) // WS_MAXIMIZEBOX
            ShowWindow(handle, 3); // SW_MAXIMIZE
    }

    private static string EscapeText(char character) => character switch
    {
        '{' or '}' or '+' or '^' or '%' or '~' or '(' or ')' or '[' or ']' => "{" + character + "}",
        '\n' => "{ENTER}", '\t' => "{TAB}", _ => character.ToString()
    };

    private static bool StateMatches(AutomationElement? element, string? expected)
    {
        if (element is null || string.IsNullOrWhiteSpace(expected)) return false;
        if (element.TryGetCurrentPattern(TogglePattern.Pattern, out var toggle))
        {
            var checkedState = ((TogglePattern)toggle).Current.ToggleState == ToggleState.On;
            if (Regex.IsMatch(expected, @"\b(unchecked|off)\b", RegexOptions.IgnoreCase)) return !checkedState;
            if (Regex.IsMatch(expected, @"\b(checked|on)\b", RegexOptions.IgnoreCase)) return checkedState;
        }
        var state = Automation.State(element);
        return state?.Contains(expected, StringComparison.OrdinalIgnoreCase) == true;
    }

    [DllImport("user32.dll")] private static extern nint GetForegroundWindow();
    [DllImport("user32.dll", SetLastError = true)] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] private static extern bool GetCursorPos(out CursorPoint point);
    [DllImport("user32.dll")] private static extern void mouse_event(uint flags, uint x, uint y, uint data, nuint extraInfo);
    [StructLayout(LayoutKind.Sequential)] private struct CursorPoint { public int X, Y; }
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(nint window);
    [DllImport("user32.dll")] private static extern bool ShowWindow(nint window, int command);
    [DllImport("user32.dll")] private static extern bool IsZoomed(nint window);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowLong(nint window, int index);
    private delegate bool EnumWindowsCallback(nint handle, nint parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsCallback callback, nint parameter);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(nint parent, EnumWindowsCallback callback, nint parameter);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(nint window);
    [DllImport("user32.dll")] private static extern nint GetParent(nint window);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(nint window, out NativeRect rectangle);
    [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(nint handle, StringBuilder text, int capacity);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(nint handle);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")]
    private static extern nint SendMessageTimeout(nint handle, uint message, nint wParam, nint lParam,
        uint flags, uint timeout, out nint result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")]
    private static extern nint SendMessageTimeout(nint handle, uint message, nint wParam, StringBuilder text,
        uint flags, uint timeout, out nint result);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(nint handle, out uint processId);
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo info);
    [StructLayout(LayoutKind.Sequential)]
    private struct GuiThreadInfo
    {
        public int Size;
        public int Flags;
        public nint Active, Focus, Capture, MenuOwner, MoveSize, Caret;
        public int CaretLeft, CaretTop, CaretRight, CaretBottom;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(nint handle, StringBuilder text, int capacity);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(nint handle);
    [DllImport("user32.dll")] private static extern nint GetWindow(nint handle, uint command);
    [DllImport("user32.dll")] private static extern nint SendMessage(nint handle, uint message, nint wParam, nint lParam);
    [DllImport("user32.dll")] private static extern void keybd_event(byte key, byte scan, uint flags, nuint extraInfo);

    private static async Task WaitForStableAsync(ControlRef target, string? before, CancellationToken token)
    {
        var deadline = DateTime.UtcNow.AddSeconds(15);
        var disabledSeen = false;
        var stable = 0;
        var observationEnd = DateTime.UtcNow.AddMilliseconds(250);
        while (DateTime.UtcNow < deadline)
        {
            token.ThrowIfCancellationRequested();
            await Task.Delay(75, token);
            var snapshot = await Task.Run(() =>
            {
                var windowExists = Automation.WindowExists(target);
                return (windowExists, state: windowExists ? Automation.State(Automation.Resolve(target)) : null);
            }, token);
            // Closing a dialog is a successful transition; its OK button cannot return.
            if (!snapshot.windowExists) return;
            if (snapshot.state?.StartsWith("enabled=False;", StringComparison.OrdinalIgnoreCase) == true)
                disabledSeen = true;
            if (!disabledSeen && DateTime.UtcNow >= observationEnd) return;
            stable = disabledSeen && snapshot.state is not null && snapshot.state == before ? stable + 1 : 0;
            if (stable >= 2) return;
        }
        throw new TimeoutException("The control stayed disabled after the action.");
    }

    private static string ToSendKeys(string key)
    {
        var parts = key.Replace(",", "+", StringComparison.Ordinal)
            .Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        var name = parts[^1];
        var prefix = string.Concat(parts[..^1].Select(x => x switch
        { "Control" or "Ctrl" => "^", "Shift" => "+", "Alt" => "%", _ => "" }));
        return prefix + (name switch
        { "Enter" or "Return" => "{ENTER}", "Escape" => "{ESC}", "Back" => "{BACKSPACE}",
          "Delete" => "{DELETE}", "Tab" => "{TAB}", "Up" => "{UP}", "Down" => "{DOWN}",
          "Left" => "{LEFT}", "Right" => "{RIGHT}", "Home" => "{HOME}", "End" => "{END}",
          "PageUp" => "{PGUP}", "PageDown" => "{PGDN}", "Space" => " ",
          _ when name.Length == 1 && char.IsLetterOrDigit(name[0]) => name.ToLowerInvariant(),
          _ => throw new NotSupportedException($"Key {name} cannot be replayed safely.") });
    }
}
