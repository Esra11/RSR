using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using DesktopSteps;

internal static class ExcelDialogCaptureCheck
{
    public static Task RunNativeVisibleFillAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        try
        {
            workbook = app.Workbooks.Add();
            var visible = new[] { 2, 7, 15, 20 };
            for (var row = 2; row <= 20; row++)
            {
                app.ActiveSheet.Cells[row, 11].Value2 = "Data " + row;
                app.ActiveSheet.Cells[row, 12].Value2 = "Original " + row;
                app.ActiveSheet.Rows[row].Hidden = !visible.Contains(row);
            }
            var native = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            var rows = (List<int>)native.GetMethod("VisibleDataRows", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "K", 2, CancellationToken.None])!;
            if (!rows.SequenceEqual(visible)) throw new InvalidOperationException("Sparse visible rows were not identified exactly.");
            native.GetMethod("SetValues", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "L", rows, "BL"]);
            for (var row = 2; row <= 20; row++)
            {
                var actual = (string)app.ActiveSheet.Cells[row, 12].Value2;
                var expected = visible.Contains(row) ? "BL" : "Original " + row;
                if (actual != expected) throw new InvalidOperationException($"Sparse fill changed L{row} to '{actual}', expected '{expected}'.");
            }
            var vertical = (string)native.GetMethod("ScrollViewport", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Vertical", 1, 10])!;
            var horizontal = (string)native.GetMethod("ScrollViewport", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Horizontal", 1, 5])!;
            if ((int)app.ActiveWindow.ScrollRow <= 1 || (int)app.ActiveWindow.ScrollColumn <= 1)
                throw new InvalidOperationException($"Native viewport movement was not applied: {vertical}; {horizontal}.");
            native.GetMethod("ScrollViewport", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Vertical", -1, 50]);
            native.GetMethod("ScrollViewport", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Vertical", -1, 50]);
            if ((int)app.ActiveWindow.ScrollRow != 1)
                throw new InvalidOperationException("Repeated native scrolling did not remain at Excel's top boundary.");
            Console.WriteLine("PASS: exact sparse visible rows were populated, hidden rows remained unchanged, and native viewport movement was verified.");
            return Task.CompletedTask;
        }
        finally
        {
            try { if (workbook is not null) workbook.Close(false); } catch { }
            try { app.Quit(); } catch { }
            if (workbook is not null && Marshal.IsComObject(workbook)) Marshal.FinalReleaseComObject(workbook);
            if (Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app);
        }
    }

    public static async Task RunFilteredFillAsync(bool rowCopyOnly = false)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        Exception? failure = null;
        nint handle = 0;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            if (rowCopyOnly) RetryBusy(() => app.ActiveSheet.Name = "Backlog (25 Sep 26)");
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Value2 = "Owner");
            RetryBusy(() => app.ActiveSheet.Cells[1, 11].Value2 = "Adjacent data");
            RetryBusy(() => app.ActiveSheet.Cells[1, 12].Value2 = "Role");
            var matchingRows = new HashSet<int> { 9, 37, 43, 49, 52 };
            for (var row = 2; row <= 83; row++)
            {
                var r = row;
                RetryBusy(() => app.ActiveSheet.Cells[r, 1].Value2 = matchingRows.Contains(r) ? "Alpha" : "Beta");
                RetryBusy(() => app.ActiveSheet.Cells[r, 11].Value2 = "Data " + r);
                RetryBusy(() => app.ActiveSheet.Cells[r, 12].Value2 = "Original " + r);
            }
            RetryBusy(() => app.ActiveSheet.ListObjects.Add(1, app.ActiveSheet.Range["A1:L83"],
                Type.Missing, 1));
            RetryBusy(() => app.ActiveSheet.Columns["A:L"].ColumnWidth = 25);
            RetryBusy(() => app.ActiveSheet.Range["A1"].Select());
            RetryBusy(() => app.WindowState = -4137);
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            await Task.Delay(500);
            var window = AutomationElement.FromHandle(handle);
            var target = new ControlRef("EXCEL", window.Current.Name, "Dropdown", "No filter applied",
                "ControlType.MenuItem", "", "A1", ProcessId: window.Current.ProcessId);
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            var compactor = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(300));
            var tab = target with
            {
                Name = (string)app.ActiveSheet.Name,
                AutomationId = "SheetTab",
                ControlType = "ControlType.TabItem",
                ClassName = "",
                ParentName = ""
            };
            if (!rowCopyOnly)
            {
                foreach (var (mode, value) in new[] { ("only", "FL"), ("exclude", "BL") })
                {
                    RetryBusy(() => app.ActiveWindow.ScrollColumn = 10);
                    RetryBusy(() => app.ActiveWindow.ScrollRow = 1);
                    await Task.Delay(150, timeout.Token);
                    var cell = target with
                    {
                        Name = "L9",
                        AutomationId = "L9",
                        ClassName = "XLSpreadsheetCell",
                        ControlType = "ControlType.DataItem",
                        ParentName = "Table"
                    };
                    var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Dynamic filtered fill",
                    [
                        new(1, "filter-values", target, "[\"Alpha\"]", null, "filter:" + mode, null, null),
                    new(2, "click", target with { Name = "Page left", AutomationId = "",
                        ClassName = "NetUIRepeatButton", ControlType = "ControlType.Button", ParentName = "" },
                        null, null, null, null, null),
                    new(2, "type", cell, value, null, null, null, null),
                    new(3, "fill-down-to-adjacent-data-end", cell, null, null, "adjacent-column:K", null,
                        "Fill visible rows until adjacent data is empty.")
                    ]);
                    const string fillIntent = "use the first visible filtered row rather than a fixed row number. intent is to copy the cell value to below cells in column L until before the empty adjacent cell in K column";
                    plan.Steps.RemoveAt(plan.Steps.Count - 1);
                    if (mode == "only")
                    {
                        plan.Steps[^1] = plan.Steps[^1] with { Value = "F:" };
                        var pane = cell with { Name = "", AutomationId = "", ControlType = "ControlType.Pane",
                            ClassName = "EXCEL6" };
                        plan.Steps.Add(new(4, "key", pane, null, "Back", null, null, null));
                        plan.Steps.Add(new(5, "type", pane, "L", null, null, null, null));
                    }
                    plan.Steps.Add(new(6, "click", cell with { Name = "G3", AutomationId = "G3" },
                        null, null, mode == "exclude" ? "unverified-click" : null, null, null, Intent: fillIntent));
                    plan = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [plan])!;
                    var messages = new List<string>();
                    await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                        [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                        (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                    for (var row = 2; row <= 83; row++)
                    {
                        var r = row;
                        string actual = "";
                        RetryBusy(() => actual = (string)app.ActiveSheet.Cells[r, 12].Value2);
                        var expected = matchingRows.Contains(r) ? "FL" : mode == "exclude" ? "BL" : "Original " + r;
                        if (actual != expected)
                            throw new InvalidOperationException($"Dynamic {mode} fill: row {r} value='{actual}', expected='{expected}'.");
                    }
                    object? below = null;
                    RetryBusy(() => below = app.ActiveSheet.Cells[84, 12].Value2);
                    if (below is not null) throw new InvalidOperationException("Fill continued past the empty adjacent cell.");
                    var expectedSource = mode == "only" ? "L9" : "L2";
                    if (!messages.Any(message => message.Contains("first visible filtered cell " + expectedSource)))
                        throw new InvalidOperationException("Dynamic seed did not select the live first visible row.");
                    Console.WriteLine($"PASS: after Page left, {mode} filter seeds live {expectedSource}; verifies all 82 rows, preserves hidden rows and stops before blank adjacent data.");
                }
                RetryBusy(() => app.ActiveWindow.ScrollColumn = 10);
                var clear = new ExecutionPlan(1, DateTimeOffset.Now, "Offscreen filter clear",
                    [new(1, "filter-values", target, "[]", null, "filter:all", null, null)]);
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [clear, (Action<string>)(_ => { }), (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                for (var row = 2; row <= 83; row++)
                {
                    var r = row;
                    bool hidden = false;
                    RetryBusy(() => hidden = (bool)app.ActiveSheet.Rows[r].Hidden);
                    if (hidden) throw new InvalidOperationException("Offscreen clear left a hidden row.");
                }
                Console.WriteLine("PASS: filter clear brings its offscreen header into view and exposes all 82 rows.");
                var bulkRows = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelNativeSheet")!
                    .GetMethod("VisibleDataRows", BindingFlags.NonPublic | BindingFlags.Static)!;
                RetryBusy(() => app.ActiveSheet.Cells[10, 11].Value2 = null);
                var beforeGap = (List<int>)bulkRows.Invoke(null,
                    [(object)app.ActiveSheet, "K", 2, timeout.Token])!;
                if (!beforeGap.SequenceEqual(Enumerable.Range(2, 8)))
                    throw new InvalidOperationException("Bulk adjacent read skipped an internal visible blank.");
                RetryBusy(() => app.ActiveSheet.Rows[10].Hidden = true);
                var hiddenGap = (List<int>)bulkRows.Invoke(null,
                    [(object)app.ActiveSheet, "K", 2, timeout.Token])!;
                if (hiddenGap.Count != 81 || hiddenGap.Contains(10) || hiddenGap[^1] != 83)
                    throw new InvalidOperationException("Bulk adjacent read treated a hidden blank as the visible stopping point.");
                RetryBusy(() => app.ActiveSheet.Rows[10].Hidden = false);
                RetryBusy(() => app.ActiveSheet.Cells[10, 11].Value2 = "Data 10");
                Console.WriteLine("PASS: bulk adjacent read stops at internal visible blanks and ignores hidden blanks.");
                var sourceSheetName = (string)app.ActiveSheet.Name;
                RetryBusy(() => app.ActiveSheet.Cells[2, 11].Formula = "=\"Data \"&ROW()");
                RetryBusy(() => app.ActiveSheet.Cells[10, 11].Value2 = null);
                RetryBusy(() => app.ActiveWorkbook.Worksheets.Add());
                var destinationSheetName = (string)app.ActiveSheet.Name;
                var copyTab = target with
                {
                    Name = sourceSheetName,
                    AutomationId = "SheetTab",
                    ControlType = "ControlType.TabItem",
                    ClassName = "",
                    ParentName = ""
                };
                var copyCell = target with
                {
                    Name = "E2",
                    AutomationId = "",
                    ClassName = "",
                    ControlType = "ControlType.DataItem"
                };
                var copyPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Column copy and Paste Values",
                    [new(1, "click", copyTab, null, null, null, null, null,
                    Intent: "intent is to copy column K until before first empty cell"),
                 new(2, "scroll", target with { ControlType = "ControlType.ScrollBar" },
                    null, "Horizontal", "range:1138", null, null),
                 new(3, "click", copyCell, null, null, null, null, null),
                 new(4, "key", target with { Name = "", AutomationId = "",
                    ControlType = "ControlType.Pane", ClassName = "XLDESK" },
                    null, "Control+C", null, null, null),
                 new(5, "click", copyTab with { Name = destinationSheetName }, null, null, null, null, null),
                 new(6, "click", copyCell with { Name = "A1", AutomationId = "A1",
                    ClassName = "XLSpreadsheetCell" }, null, null, null, null, null),
                 new(7, "click", target with { Name = "More Options", AutomationId = "PasteMenu_Dropdown",
                    ControlType = "ControlType.MenuItem", ClassName = "NetUIRibbonButton", ParentName = null },
                    null, null, null, null, null),
                 new(8, "click", target with { Name = "Values", AutomationId = "",
                    ControlType = "ControlType.ListItem", ClassName = "NetUIGalleryButton",
                    ParentName = "Paste Values" }, null, null, null, null, null)]);
                copyPlan = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [copyPlan])!;
                if (copyPlan.Steps.Count != 6 || copyPlan.Steps[1].Action != "copy-column-until-empty" ||
                    copyPlan.Steps[1].Value != "K2")
                    throw new InvalidOperationException("Explicit column-copy intent was not normalized correctly.");
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [copyPlan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                for (var row = 1; row <= 8; row++)
                    if ((string)app.ActiveSheet.Cells[row, 1].Value2 != "Data " + (row + 1))
                        throw new InvalidOperationException("Paste Values did not receive the copied column data.");
                if ((bool)app.ActiveSheet.Cells[1, 1].HasFormula)
                    throw new InvalidOperationException("Paste Values copied a formula instead of its result.");
                if (app.ActiveSheet.Cells[9, 1].Value2 is not null)
                    throw new InvalidOperationException("Column copy continued beyond the internal blank.");
                RetryBusy(() => app.Worksheets[sourceSheetName].Activate());
                RetryBusy(() => app.ActiveSheet.Cells[10, 11].Value2 = "Data 10");
                Console.WriteLine("PASS: explicit K2 copy stops before internal blank, switches sheet and pastes the copied values into A1.");
                RetryBusy(() => app.ActiveSheet.ListObjects[1].Unlist());
                var columnHeader = target with
                {
                    Name = "A",
                    AutomationId = "",
                    ClassName = "XLGridColumnHeader",
                    ControlType = "ControlType.DataItem",
                    ParentName = "Grid"
                };
                var deleteColumn = new ExecutionPlan(1, DateTimeOffset.Now, "Clear selected column",
                    [new(1, "click", columnHeader, null, null, null, null, null),
                 new(2, "key", columnHeader with { Name = "A1", AutomationId = "A1",
                     ClassName = "XLSpreadsheetCell" }, null, "Delete", "field-value:", null, null)]);
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [deleteColumn, (Action<string>)(_ => { }), (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                double remaining = -1;
                RetryBusy(() => remaining = (double)app.WorksheetFunction.CountA(app.ActiveSheet.Columns["A"]));
                if (remaining != 0 || (string)app.ActiveSheet.Cells[1, 11].Value2 != "Adjacent data")
                    throw new InvalidOperationException("Column clear left content or changed the neighboring columns.");
                Console.WriteLine("PASS: whole-column selection plus Delete clears all A contents and leaves K unchanged.");
                var originalName = (string)app.ActiveSheet.Name;
                RetryBusy(() => app.ActiveWorkbook.Worksheets.Add());
                var otherName = (string)app.ActiveSheet.Name;
                tab = tab with { Name = originalName };
                var datedPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Dynamic dated tab reference",
                    [new(1, "rename-sheet", tab, "Backlog ({{next:Tuesday:dd MMM yy}})", null,
                    null, null, null),
                 new(2, "click", tab with { Name = otherName }, null, null, null, null, null),
                 new(3, "click", tab with { Name = "Backlog (25 Sep 26)" }, null, null, null, null, null)]);
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [datedPlan, (Action<string>)(_ => { }), (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                var resolvedName = (string)app.ActiveSheet.Name;
                if (!resolvedName.StartsWith("Backlog (") || resolvedName == otherName)
                    throw new InvalidOperationException("Missing dated reference did not return to the sheet renamed in this replay.");
                Console.WriteLine("PASS: missing old dated tab returns to the sheet explicitly renamed during replay after switching to another sheet.");
                RetryBusy(() => app.ActiveWorkbook.Worksheets.Add());
                RetryBusy(() => app.ActiveSheet.Name = "Backlog (25 Sep 26)");
                var existingDatePlan = new ExecutionPlan(1, DateTimeOffset.Now, "Existing dated tab reference",
                    [new(1, "rename-sheet", tab with { Name = resolvedName },
                    "Backlog ({{next:Wednesday:dd MMM yy}})", null, null, null, null),
                 new(2, "click", tab with { Name = "Backlog (25 Sep 26)" },
                    null, null, null, null, null)]);
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [existingDatePlan, (Action<string>)(_ => { }), (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                if ((string)app.ActiveSheet.Name != "Backlog (25 Sep 26)")
                    throw new InvalidOperationException("An existing dated tab was incorrectly redirected.");
                Console.WriteLine("PASS: existing old dated tabs are preserved rather than redirected to a generated sheet.");
            }
            var existingTab = tab with { Name = "Backlog (25 Sep 26)" };
            var dialogTarget = target with
            {
                Window = "Move or Copy",
                Name = "Create a copy",
                AutomationId = "",
                ControlType = "ControlType.CheckBox",
                ClassName = "NetUICheckbox",
                ParentName = "Move or Copy"
            };
            var sheetCopyPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Copy dialog regression",
                [new(1, "context-click", existingTab, null, null, null, null, null),
                 new(2, "click", target with { Name = "Move or Copy...",
                    ControlType = "ControlType.MenuItem" }, null, null, null, null, null),
                 new(3, "click", dialogTarget, null, null, null, null, null),
                 new(4, "click", dialogTarget with { Name = "Backlog (25 Sep 26)",
                    ControlType = "ControlType.ListItem", ClassName = "NetUIListViewItem",
                    ParentName = "Before sheet" }, null, null, null, null, null),
                 new(5, "click", dialogTarget with { Name = "OK",
                    ControlType = "ControlType.Button", ClassName = "NetUIButton" },
                    null, null, null, null, null)]);
            var copyDuration = System.Diagnostics.Stopwatch.StartNew();
            var countBeforeCopy = (int)app.Worksheets.Count;
            await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [sheetCopyPlan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            if ((int)app.Worksheets.Count != countBeforeCopy + 1 ||
                copyDuration.Elapsed > TimeSpan.FromSeconds(20))
                throw new InvalidOperationException("Copy dialog failed to create a sheet within 20 seconds.");
            Console.WriteLine($"PASS: copy dialog created one new sheet in {copyDuration.Elapsed.TotalSeconds:F1} seconds.");
            RetryBusy(() => app.ActiveSheet.Cells[2, 1].Value2 = "Row two");
            RetryBusy(() => app.ActiveSheet.Cells[2, 2].Value2 = "Before");
            RetryBusy(() => app.ActiveSheet.Cells[3, 1].Value2 = "Row three");
            RetryBusy(() => app.ActiveSheet.Cells[3, 2].Value2 = "After");
            RetryBusy(() => app.ActiveWindow.ScrollColumn = 2);
            var b2 = target with
            {
                Name = "B2",
                AutomationId = "B2",
                ClassName = "XLSpreadsheetCell",
                ControlType = "ControlType.DataItem",
                ParentName = "Grid"
            };
            var rowHeader = target with
            {
                Name = "2",
                AutomationId = "",
                ClassName = "XLGridRowHeader",
                ControlType = "ControlType.DataItem",
                ParentName = "Grid"
            };
            var rowDelete = new ExecutionPlan(1, DateTimeOffset.Now, "Offscreen A2 row deletion",
                [new(1, "click", b2, null, null, null, null, null),
                 new(2, "click", rowHeader, null, null, null, null, null),
                 new(3, "context-click", rowHeader, null, null, null, null, null),
                 new(4, "click", target with { Name = "Delete", AutomationId = "",
                    ControlType = "ControlType.MenuItem", ClassName = "NetUITWBtnMenuItem",
                    ParentName = "Context Menu" }, null, null, null, null, null)]);
            await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [rowDelete, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            if ((string)app.ActiveSheet.Cells[2, 2].Value2 != "After")
                throw new InvalidOperationException("Row 2 did not delete when A2 initially lay outside the viewport.");
            Console.WriteLine("PASS: row 2 deletion reveals initially offscreen A2 and verifies row 3 moved up.");
        }
        catch (Exception ex) { failure = ex; throw; }
        finally
        {
            try
            {
                SetForegroundWindow(handle);
                AutomationElement.FromHandle(handle).SetFocus();
                await Task.Delay(150);
                SendKeys.SendWait("{ESC}{ESC}");
                RetryBusy(() => app.DisplayAlerts = false);
                if (workbook is not null)
                {
                    RetryBusy(() => workbook.Close(false));
                    Marshal.FinalReleaseComObject(workbook);
                }
                RetryBusy(() => app.Quit());
                Marshal.FinalReleaseComObject(app);
            }
            catch (Exception ex) when (failure is not null)
            { Console.Error.WriteLine("Scratch filtered fill cleanup failed: " + ex.Message); }
        }
    }

    public static async Task RunCellNavigationAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        Exception? failure = null;
        nint handle = 0;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Value2 = "Owner");
            RetryBusy(() => app.ActiveSheet.Cells[2, 1].Value2 = "Alpha");
            RetryBusy(() => app.ActiveSheet.Cells[3, 1].Value2 = "Beta");
            RetryBusy(() => app.ActiveSheet.Cells[2, 12].Value2 = "Before");
            RetryBusy(() => app.ActiveSheet.Cells[3, 12].Value2 = "Untouched");
            RetryBusy(() => app.ActiveSheet.Cells[1, 7].Value2 = "Owner");
            RetryBusy(() => app.ActiveSheet.Range["A1:L3"].AutoFilter(1, "Alpha"));
            RetryBusy(() => app.ActiveSheet.Range["A1"].Select());
            RetryBusy(() => app.WindowState = -4137);
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            await Task.Delay(500);
            var window = AutomationElement.FromHandle(handle);
            var target = new ControlRef("EXCEL", window.Current.Name, "", "F9", "ControlType.DataItem",
                "", "Table", ProcessId: window.Current.ProcessId);
            var input = new PlanStep(2, "type", target with
            {
                Name = "L2",
                AutomationId = "L2",
                ClassName = "XLSpreadsheetCell"
            }, "FL", null, null, null, null);
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Filtered offscreen input",
                [new(1, "click", target, null, null, null, null, null), input]);
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            var messages = new List<string>();
            await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            SendKeys.SendWait("^g");
            await Task.Delay(100);
            SendKeys.SendWait("G1{ENTER}");
            await Task.Delay(100);
            string value = "", other = "";
            string navigationTarget = "";
            RetryBusy(() => value = (string)app.ActiveSheet.Cells[2, 12].Value2);
            RetryBusy(() => other = (string)app.ActiveSheet.Cells[3, 12].Value2);
            RetryBusy(() => navigationTarget = (string)app.ActiveSheet.Cells[1, 7].Value2);
            if (value != "FL" || other != "Untouched" || navigationTarget != "Owner")
                throw new InvalidOperationException($"Fixed-address input was not committed before the next Go To command: value='{value}', other='{other}', G1='{navigationTarget}'; {string.Join("; ", messages)}");
            Console.WriteLine("PASS: fixed-address Excel input commits and verifies its exact value before a following Go To command; the navigation address is not appended to the cell text.");
            var hidden = input with
            {
                Number = 1,
                Target = input.Target! with { Name = "L3", AutomationId = "L3" },
                Value = "Must not enter"
            };
            try
            {
                await ((Task)executor.GetMethod("TypeIntoExcelCellByAddressAsync", BindingFlags.NonPublic | BindingFlags.Static)!
                    .Invoke(null, [hidden, "L3", timeout.Token])!).WaitAsync(timeout.Token);
                throw new InvalidOperationException("Hidden filtered row was accepted for text entry.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("No text was entered") ||
                ex.Message.Contains("instead of L3"))
            { Console.WriteLine("PASS: hidden filtered row stops before text input rather than selecting a substitute row."); }
            RetryBusy(() => other = (string)app.ActiveSheet.Cells[3, 12].Value2);
            if (other != "Untouched") throw new InvalidOperationException("Hidden row was modified.");
            var nativeSheet = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            RetryBusy(() => nativeSheet.GetMethod("SelectEntireColumn", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [app.ActiveSheet, "A"]));
            RetryBusy(() => nativeSheet.GetMethod("VerifyColumnSelection", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [app.ActiveSheet, "A"]));
            RetryBusy(() => app.ActiveSheet.Range["A1"].Select());
            Console.WriteLine("PASS: native Excel selection selects and verifies the entire requested column.");
        }
        catch (Exception ex) { failure = ex; throw; }
        finally
        {
            try
            {
                SetForegroundWindow(handle);
                AutomationElement.FromHandle(handle).SetFocus();
                await Task.Delay(200);
                SendKeys.SendWait("{ESC}{ESC}");
                await Task.Delay(200);
                RetryBusy(() => app.DisplayAlerts = false);
                if (workbook is not null)
                {
                    RetryBusy(() => workbook.Close(false));
                    Marshal.FinalReleaseComObject(workbook);
                }
                RetryBusy(() => app.Quit());
                Marshal.FinalReleaseComObject(app);
            }
            catch (Exception ex) when (failure is not null)
            { Console.Error.WriteLine("Scratch cell navigation cleanup failed: " + ex.Message); }
        }
    }

    public static async Task RunColumnSelectionAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Value2 = "Keep until verified delete");
            var nativeSheet = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            RetryBusy(() => nativeSheet.GetMethod("SelectEntireColumn", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [app.ActiveSheet, "A"]));
            RetryBusy(() => nativeSheet.GetMethod("VerifyColumnSelection", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [app.ActiveSheet, "A"]));
            if ((string)app.ActiveSheet.Cells[1, 1].Value2 != "Keep until verified delete")
                throw new InvalidOperationException("Selecting the entire column changed its contents.");
            Console.WriteLine("PASS: native Excel selection selects and verifies the entire requested column without changing cells.");
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Interior.Color = 65535);
            RetryBusy(() => nativeSheet.GetMethod("ClearColumnContents", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [app.ActiveSheet, "A"]));
            object? cleared = null;
            int color = 0;
            RetryBusy(() => cleared = app.ActiveSheet.Cells[1, 1].Value2);
            RetryBusy(() => color = (int)app.ActiveSheet.Cells[1, 1].Interior.Color);
            if (cleared is not null || color != 65535)
                throw new InvalidOperationException("Native column clear did not preserve formatting or remove all contents.");
            Console.WriteLine("PASS: native Excel column clear removes contents and preserves formatting without UI selection.");
            await Task.CompletedTask;
        }
        finally
        {
            RetryBusy(() => app.DisplayAlerts = false);
            if (workbook is not null)
            {
                RetryBusy(() => workbook.Close(false));
                Marshal.FinalReleaseComObject(workbook);
            }
            RetryBusy(() => app.Quit());
            Marshal.FinalReleaseComObject(app);
        }
    }

    public static async Task RunFormulaExtensionAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            for (var row = 1; row <= 20; row++)
            {
                var current = row;
                RetryBusy(() => app.ActiveSheet.Cells[current, 3].Value2 = "Category " + current);
            }
            RetryBusy(() => app.ActiveSheet.Cells[23, 3].Value2 = "After gap");
            RetryBusy(() => app.ActiveSheet.Cells[21, 3].Formula = "=\"invalid\"+0");
            RetryBusy(() => app.ActiveSheet.Cells[22, 3].Value2 = "#VALUE!");
            for (var row = 1; row <= 10; row++)
            {
                var current = row;
                RetryBusy(() => app.ActiveSheet.Cells[current, 6].Formula = $"=COUNTIF($C$1:$C$20,E{current})");
            }
            RetryBusy(() => app.ActiveSheet.Cells[20, 6].Value2 = "stale constant");
            RetryBusy(() => app.ActiveSheet.Cells[1, 6].Formula = "=99");
            RetryBusy(() => app.ActiveSheet.Range["A1"].Select());
            RetryBusy(() => app.WindowState = -4137);
            nint handle = 0;
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            var window = AutomationElement.FromHandle(handle);
            var cell = new ControlRef("EXCEL", window.Current.Name, "C1", "C1",
                "ControlType.DataItem", "XLSpreadsheetCell", "Grid", ProcessId: window.Current.ProcessId);
            SetForegroundWindow(handle);
            await Task.Delay(500);
            var copyIntent = "intent is to copy column C until before cell that contains empty value or the value \"#VALUE!\". User clarification: copy C1 down to the first empty cell and paste values starting at E1.";
            var formulaIntent = "intent is to copy the function in column F till just before the empty cell in of column E. User clarification: extend from the last populated F formula cell; preserve existing formulas.";
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Clarified range and formula fill",
                [new(1, "click", cell, null, null, null, null, null),
                         new(2, "key", cell with { Name = "", AutomationId = "", ClassName = "XLDESK",
                            ControlType = "ControlType.Pane" }, null, "Control+C", null, null, null, Intent: copyIntent),
                         new(3, "click", cell with { Name = "E1", AutomationId = "E1" }, null, null, null, null, null),
                         new(4, "click", cell with { Name = "More Options", AutomationId = "PasteMenu_Dropdown",
                            ControlType = "ControlType.MenuItem" }, null, null, null, null, null),
                         new(5, "click", cell with { Name = "Values", AutomationId = "",
                            ControlType = "ControlType.ListItem", ParentName = "Paste Values" }, null, null, null, null, null,
                            Intent: formulaIntent),
                         new(6, "click", cell with { Name = "F10", AutomationId = "F10" },
                            null, null, null, null, null),
                         new(7, "click", cell with { Name = "F10", AutomationId = "F10" },
                            null, null, null, null, null)]);
            var assembly = typeof(ControlRef).Assembly;
            plan = (ExecutionPlan)assembly.GetType("DesktopSteps.PlanCompactor")!
                .GetMethod("Compact")!.Invoke(null, [plan])!;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(180));
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [plan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                            (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            for (var row = 1; row <= 20; row++)
            {
                if ((string)app.ActiveSheet.Cells[row, 5].Value2 != "Category " + row)
                    throw new InvalidOperationException($"Range copy did not populate E{row}.");
                var expectedFormula = row == 1 ? "=99" : $"=COUNTIF($C$1:$C$20,E{row})";
                if ((string)app.ActiveSheet.Cells[row, 6].Formula != expectedFormula)
                    throw new InvalidOperationException($"Formula extension changed existing formulas or relative references at F{row}.");
            }
            if (app.ActiveSheet.Cells[21, 5].Value2 is not null ||
                app.ActiveSheet.Cells[21, 6].Value2 is not null ||
                app.ActiveSheet.Cells[23, 5].Value2 is not null)
                throw new InvalidOperationException("Copy or formula extension continued past the first empty row.");
            Console.WriteLine("PASS: C1:C20 values paste into E1:E20; F10 extends through F20 with relative references; F1:F9 and rows after gap remain unchanged.");
            string formulaSourceSheet = app.ActiveSheet.Name;
            dynamic explicitSheet = app.Worksheets.Add();
            for (var row = 1; row <= 5; row++)
            {
                var current = row;
                RetryBusy(() => app.ActiveSheet.Cells[current, 1].Value2 = current);
            }
            RetryBusy(() => app.ActiveSheet.Cells[1, 3].Formula = "=A1*2");
            RetryBusy(() => app.ActiveSheet.Cells[2, 3].Formula = "=99");
            RetryBusy(() => app.ActiveSheet.Cells[7, 3].Value2 = "After gap");
            var explicitFormula = new ExecutionPlan(1, DateTimeOffset.Now, "Explicit formula source",
                [new(1, "click", cell,
                    null, null, null, null, null,
                    Intent: "intent is to copy the function from the C1 cell in column C to the last cell that has an adjacent existing value in column A"),
                 new(2, "click", cell,
                    null, null, null, null, null,
                    Intent: "intent is to copy the function from the C1 cell in column C to the last cell that has an adjacent existing value in column A")]);
            explicitFormula = (ExecutionPlan)assembly.GetType("DesktopSteps.PlanCompactor")!
                .GetMethod("Compact")!.Invoke(null, [explicitFormula])!;
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [explicitFormula, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                            (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            for (var row = 1; row <= 5; row++)
                if ((string)app.ActiveSheet.Cells[row, 3].Formula != $"=A{row}*2")
                    throw new InvalidOperationException("Explicit formula fill did not use its named source or adjust references.");
            if (app.ActiveSheet.Cells[6, 3].Value2 is not null ||
                (string)app.ActiveSheet.Cells[7, 3].Value2 != "After gap")
                throw new InvalidOperationException("Explicit formula fill exceeded its adjacent data boundary.");
            Marshal.FinalReleaseComObject(explicitSheet);
            RetryBusy(() => app.Worksheets[formulaSourceSheet].Activate());
            Console.WriteLine("PASS: ordinary C1 clicks fill C1:C5 from C1 through A's live data boundary; relative references verified.");
            var ribbonPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Recorded ribbon command",
                [new(1, "click", cell with { Name = "E", AutomationId = "", ClassName = "XLGridColumnHeader",
                        ControlType = "ControlType.DataItem" }, null, null, null, null, null),
                 new(2, "click", cell with { Name = "Data", AutomationId = "TabData", ClassName = "NetUIRibbonTab",
                        ControlType = "ControlType.TabItem", ParentName = "Ribbon Tabs" }, null, null, null, null, null),
                 new(3, "click", cell with { Name = "Remove Duplicates", AutomationId = "RemoveDuplicates",
                        ClassName = "NetUIRibbonButton", ControlType = "ControlType.Button", ParentName = "Data Tools" },
                        null, null, null, null, null),
                 new(4, "click", new ControlRef("EXCEL", "Remove Duplicates Warning", "", "Remove Duplicates...",
                        "ControlType.Button", "", "Remove Duplicates Warning"), null, null, null, null, null),
                 new(5, "click", new ControlRef("EXCEL", "Remove Duplicates", "", "OK",
                        "ControlType.Button", "NetUIButton", "Remove Duplicates"), null, null, null, null, null),
                 new(6, "click", new ControlRef("EXCEL", "Microsoft Excel", "2", "OK",
                        "ControlType.Button", "NetUIButton", "Microsoft Excel"), null, null, null, null, null)]);
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [ribbonPlan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                            (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            await Task.Delay(300, timeout.Token);
            if (!window.Current.IsEnabled ||
                window.FindAll(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, "OK")))
                    .Cast<AutomationElement>().Any(button => button.Current.IsEnabled && !button.Current.IsOffscreen))
                throw new InvalidOperationException("Remove Duplicates result remained open after replay acknowledged OK.");
            Console.WriteLine("PASS: recorded Remove Duplicates warning, settings OK and result OK complete without activating the disabled workbook owner.");
            string sourceSheet = app.ActiveSheet.Name;
            dynamic destinationSheet = app.Worksheets.Add();
            destinationSheet.Name = "Renamed destination";
            app.Worksheets[sourceSheet].Activate();
            var columnsIntent = "intent is to copy E and F cells that contain a value";
            dynamic chartObject = destinationSheet.ChartObjects().Add(100, 100, 400, 250);
            chartObject.Name = "Fixture chart";
            chartObject.Chart.ChartType = 51;
            Marshal.FinalReleaseComObject(chartObject);
            var chartIntent = "User clarification: replace the chart data source with the last pasted two-column range.";
            var returnPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Two-column return paste",
                [new(1, "click", cell with { Name = "E1", AutomationId = "E1" }, null, null, null, null, null, Intent: columnsIntent),
                 new(2, "key", cell with { Name = "E1", AutomationId = "E1" }, null, "Control+C",
                    "excel-selection-range:E1:F10", null, null),
                 new(3, "click", cell with { Name = "Renamed destination", AutomationId = "SheetTab",
                    ControlType = "ControlType.TabItem", ClassName = "" }, null, null, null, null, null),
                 new(4, "click", cell with { Name = "O7", AutomationId = "O7" }, null, null, null, null, null),
                 new(5, "click", cell with { Name = "File Tab", AutomationId = "FileTabButton",
                    ControlType = "ControlType.Button", ClassName = "NetUIRibbonTab", ParentName = "Ribbon" },
                    null, null, null, null, null),
                 new(6, "click", cell with { Name = "Back", AutomationId = "FileTabButton",
                    ControlType = "ControlType.ListItem", ClassName = "NetUISimpleButton", ParentName = "File" },
                    null, null, null, null, null),
                 new(5, "click", cell with { Name = "Home", AutomationId = "TabHome",
                    ControlType = "ControlType.TabItem", ClassName = "NetUIRibbonTab" }, null, null, null, null, null),
                 new(6, "click", cell with { Name = "More Options", AutomationId = "PasteMenu_Dropdown",
                    ControlType = "ControlType.MenuItem" }, null, null, null, null, null),
                 new(7, "click", cell with { Name = "Values", AutomationId = "", ParentName = "Paste Values",
                    ControlType = "ControlType.ListItem" }, null, null, null, null, null),
                 new(8, "context-click", cell with { Name = "Fixture chart", AutomationId = "",
                    ControlType = "ControlType.Image", ClassName = "ExcelChartObject" },
                    null, null, null, null, null),
                 new(9, "click", cell with { Name = "Select Data...", AutomationId = "",
                    ControlType = "ControlType.MenuItem" }, null, null, null, null, null, Intent: chartIntent),
                 new(10, "click", cell with { Name = "O7", AutomationId = "O7" },
                    null, null, null, null, null),
                 new(11, "click", cell with { Name = "OK", Window = "Select Data Source",
                    ControlType = "ControlType.Button" }, null, null, null, null, null)]);
            returnPlan = (ExecutionPlan)assembly.GetType("DesktopSteps.PlanCompactor")!
                .GetMethod("Compact")!.Invoke(null, [returnPlan])!;
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [returnPlan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                            (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            for (var row = 1; row <= 20; row++)
            {
                if ((string)destinationSheet.Cells[row + 6, 15].Value2 != "Category " + row ||
                    Convert.ToDouble(destinationSheet.Cells[row + 6, 16].Value2) != (row == 1 ? 99 : 1))
                    throw new InvalidOperationException("Two-column paste lost a value at row " + row);
            }
            if (destinationSheet.Cells[27, 15].Value2 is not null || destinationSheet.Cells[27, 16].Value2 is not null ||
                (bool)destinationSheet.Range["O7:P26"].HasFormula)
                throw new InvalidOperationException("Two-column paste exceeded the boundary or pasted formulas.");
            await CheckLegendDragAsync(app, destinationSheet, window, assembly, timeout.Token);
            Marshal.FinalReleaseComObject(destinationSheet);
            app.Worksheets[sourceSheet].Activate();
            Console.WriteLine("PASS: E1:F20 copied dynamically beyond recorded F10 and pasted as 40 values at O7:P26 after sheet switching and File/Back/Home; no formulas or extra rows.");
            await CheckWorksheetDragAsync(app, window, assembly, timeout.Token);
            RetryBusy(() => app.ActiveSheet.Range["C1:C20"].Select());
            var native = assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            var literalEnd = (int)native.GetMethod("ContiguousDataEnd", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "C", 22, timeout.Token, new[] { "#VALUE!" }])!;
            if (literalEnd != 21) throw new InvalidOperationException("Literal #VALUE! boundary was included.");
            var captured = (string)native.GetMethod("SelectionAddress", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet])!;
            if (captured != "C1:C20") throw new InvalidOperationException("Copy selection capture returned only the active cell.");
            var capturedPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Captured rectangle replay",
                [new(1, "click", cell, null, null, null, null, null),
                         new(2, "key", cell, null, "Control+C", "excel-selection-range:" + captured, null, null)]);
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [capturedPlan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                            (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            if ((string)app.Selection.Address != "$C$1:$C$20")
                throw new InvalidOperationException("Recorded range was collapsed during replay.");
            Console.WriteLine("PASS: captured Ctrl+C range is restored and verified instead of copying only C1.");
            var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
            using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
            var captureDirectory = Path.Combine(Path.GetTempPath(), "RSR-range-capture-" + Guid.NewGuid().ToString("N"));
            recorderType.GetMethod("Start")!.Invoke(recorder, [captureDirectory]);
            try
            {
                SendKeys.SendWait("^c");
                await Task.Delay(500, timeout.Token);
                recorderType.GetMethod("Stop")!.Invoke(recorder, null);
                var recorded = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
                if (!recorded.Any(item => item.Kind == "key" && item.Key == "Control+C" &&
                    item.AfterState == "excel-selection-range:C1:C20"))
                    throw new InvalidOperationException("Recorder did not preserve the copied rectangle at Ctrl+C: " +
                        string.Join("; ", recorded.Select(item => $"{item.Kind}, {item.Key}, {item.AfterState}, {item.Diagnostic}")));
                Console.WriteLine("PASS: live recorder Ctrl+C captures C1:C20 instead of just the active cell.");
            }

            finally
            {
                recorderType.GetMethod("Stop")!.Invoke(recorder, null);
                if (Directory.Exists(captureDirectory))
                {
                    foreach (var file in Directory.GetFiles(captureDirectory)) File.Delete(file);
                    Directory.Delete(captureDirectory);
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Formula/duplicate replay failure before cleanup: " + ex);
            throw;
        }
        finally
        {
            RetryBusy(() => app.DisplayAlerts = false);
            if (workbook is not null)
            {
                RetryBusy(() => workbook.Close(false));
                Marshal.FinalReleaseComObject(workbook);
            }
            RetryBusy(() => app.Quit());
            Marshal.FinalReleaseComObject(app);
        }
    }

    private static async Task CheckWorksheetDragAsync(dynamic app, AutomationElement window, Assembly assembly,
        CancellationToken token)
    {
        app.ActiveSheet.Range["C1"].Select();
        await Task.Delay(200, token);
        var cell = window.FindFirst(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.NameProperty, "C1"),
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem)))
            ?? throw new InvalidOperationException("Scratch worksheet cell unavailable.");
        var bounds = cell.Current.BoundingRectangle;
        var x = (int)(bounds.Left + bounds.Width / 2);
        var y = (int)(bounds.Top + bounds.Height / 2);
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var folder = Path.Combine(Path.GetTempPath(), "RSR-worksheet-drag-" + Guid.NewGuid().ToString("N"));
        recorderType.GetMethod("Start")!.Invoke(recorder, [folder]);
        try
        {
            Cursor.Position = new System.Drawing.Point(x, y);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(150, token);
            Cursor.Position = new System.Drawing.Point(x, y + (int)(bounds.Height * 3));
            await Task.Delay(150, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(300, token);
            SendKeys.SendWait("^c");
            await Task.Delay(300, token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            if (events.Any(item => item.Kind is "unresolved-input" or "set-chart-legend-layout") ||
                !events.Any(item => item.Kind == "key" && item.Key == "Control+C" &&
                    item.AfterState == "excel-selection-range:C1:C4"))
                throw new InvalidOperationException("Worksheet drag failed selection capture or was classified as legend resize: " +
                    string.Join("; ", events.Select(item => $"{item.Kind}: {item.AfterState}: {item.Diagnostic}")));
            Console.WriteLine("PASS: live C1:C4 selection drag remains non-chart input and following Ctrl+C captures the full range.");
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            if (Directory.Exists(folder))
            {
                foreach (var file in Directory.GetFiles(folder)) File.Delete(file);
                Directory.Delete(folder);
            }
        }
    }

    public static async Task RunChartCaptureAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        nint scratchHandle = 0;
        uint scratchProcess = 0;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            for (var row = 1; row <= 20; row++)
            {
                var current = row;
                RetryBusy(() => app.ActiveSheet.Cells[current, 1].Value2 = "Category " + current);
                RetryBusy(() => app.ActiveSheet.Cells[current, 2].Value2 = current);
            }
            dynamic chartObject = app.ActiveSheet.ChartObjects().Add(300, 100, 400, 250);
            chartObject.Name = "Fixture chart";
            chartObject.Chart.ChartType = 51;
            chartObject.Chart.SetSourceData(app.ActiveSheet.Range["A1:B20"]);
            for (var index = 0; index < 3; index++)
            {
                object otherChart = app.ActiveSheet.ChartObjects().Add(100, 700 + index * 300, 200, 200);
                try { ((dynamic)otherChart).Name = "Other chart " + index; }
                finally { Marshal.ReleaseComObject(otherChart); }
            }
            var native = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            var sourceType = native.GetNestedType("ChartSource", BindingFlags.NonPublic)!;
            var capturedSource = native.GetMethod("ReadChartSource", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Fixture chart"]);
            Console.WriteLine("Captured native chart source: " + System.Text.Json.JsonSerializer.Serialize(capturedSource));
            object originalSheet = app.ActiveSheet;
            object otherSheet = workbook.Worksheets.Add();
            try
            {
                ((dynamic)otherSheet).Activate();
                var inactiveSource = native.GetMethod("ReadChartSourceFromSheet", BindingFlags.NonPublic | BindingFlags.Static)!
                    .Invoke(null, [(object)app.ActiveSheet, (string)((dynamic)originalSheet).Name, "Fixture chart"]);
                if (System.Text.Json.JsonSerializer.Serialize(inactiveSource) !=
                    System.Text.Json.JsonSerializer.Serialize(capturedSource))
                    throw new InvalidOperationException("Chart source changed when read while another worksheet was active.");
                Console.WriteLine("PASS: chart source remains capturable while Excel's range selector activates another worksheet.");
            }
            finally
            {
                ((dynamic)originalSheet).Activate();
                Marshal.ReleaseComObject(otherSheet);
                Marshal.ReleaseComObject(originalSheet);
            }
            app.ActiveSheet.Cells[1, 15].Value2 = "First";
            app.ActiveSheet.Cells[1, 16].Value2 = 7;
            app.ActiveSheet.Cells[2, 15].Value2 = "Second";
            app.ActiveSheet.Cells[2, 16].Value2 = 9;
            var source = Activator.CreateInstance(sourceType, [(string)app.ActiveSheet.Name, "A1:B2"]);
            native.GetMethod("ApplyChartSource", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Fixture chart", source, null]);
            Marshal.FinalReleaseComObject(chartObject);
            nint handle = (nint)(int)app.Hwnd;
            scratchHandle = handle;
            GetWindowThreadProcessId(handle, out scratchProcess);
            RetryBusy(() => app.WindowState = -4137);
            SetForegroundWindow(handle);
            await Task.Delay(500);
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            await CheckChartSourceRecordingAsync(app, AutomationElement.FromHandle(handle),
                typeof(ControlRef).Assembly, timeout.Token);
            await CheckLegendDragAsync(app, app.ActiveSheet, AutomationElement.FromHandle(handle),
                typeof(ControlRef).Assembly, timeout.Token);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Chart capture check failed before cleanup: " + ex);
            throw;
        }
        finally
        {
            try
            {
                if (scratchHandle != 0 && GetForegroundWindow() == scratchHandle)
                    SendKeys.SendWait("{ESC}{ESC}");
                RetryBusy(() => app.DisplayAlerts = false);
                if (workbook is not null) RetryBusy(() => workbook.Close(false));
                RetryBusy(() => app.Quit());
            }
            catch (Exception cleanup) when (cleanup is COMException or System.ComponentModel.Win32Exception)
            {
                Console.Error.WriteLine("Scratch Excel cleanup: " + cleanup.Message);
                if (scratchProcess != 0)
                    System.Diagnostics.Process.GetProcessById((int)scratchProcess).Kill();
            }
            finally
            {
                if (workbook is not null) Marshal.FinalReleaseComObject(workbook);
                Marshal.FinalReleaseComObject(app);
            }
        }
    }

    public static Task InspectChartAsync(int processId, int x, int y)
    {
        using var process = System.Diagnostics.Process.GetProcessById(processId);
        var assembly = typeof(ControlRef).Assembly;
        var native = assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
        var read = native.GetMethod("Read", BindingFlags.NonPublic | BindingFlags.Static)!
            .MakeGenericMethod(typeof(bool));
        read.Invoke(null, [process.MainWindowHandle, (Func<object, bool>)(sheetObject =>
        {
            dynamic sheet = sheetObject;
            Console.WriteLine("Native active sheet: " + (string)sheet.Name);
            object charts = sheet.ChartObjects();
            try
            {
                for (var index = 1; index <= (int)((dynamic)charts).Count; index++)
                {
                    object chart = ((dynamic)charts).Item(index);
                    try { Console.WriteLine("Native chart: " + (string)((dynamic)chart).Name); }
                    finally { Marshal.ReleaseComObject(chart); }
                }
            }
            finally { Marshal.ReleaseComObject(charts); }
            var element = AutomationElement.FromPoint(new(x, y));
            var names = new List<string>();
            for (var depth = 0; element is not null && depth < 12; depth++)
            {
                Console.WriteLine("UIA ancestor: " + element.Current.ControlType.ProgrammaticName + "/" + element.Current.Name);
                if (element.Current.ControlType == ControlType.Image && !string.IsNullOrWhiteSpace(element.Current.Name))
                    names.Add(element.Current.Name);
                element = TreeWalker.RawViewWalker.GetParent(element);
            }
            var capture = native.GetMethod("CaptureChartAtPoint", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [sheetObject, x, y, true, names]);
            Console.WriteLine("Capture: " + capture);
            return true;
        })]);
        return Task.CompletedTask;
    }

    private static async Task CheckChartSourceRecordingAsync(dynamic app, AutomationElement window,
        Assembly assembly, CancellationToken token)
    {
        app.ActiveWindow.ScrollRow = 1;
        app.ActiveWindow.ScrollColumn = 1;
        dynamic chartObject = app.ActiveSheet.ChartObjects().Item("Fixture chart");
        chartObject.Activate();
        await Task.Delay(400, token);
        assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
            BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
        chartObject.Activate();
        await Task.Delay(300, token);
        AutomationElement? image = null;
        for (var attempt = 0; image is null && attempt < 30; attempt++)
        {
            image = window.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "Fixture chart"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Image)));
            if (image is null) await Task.Delay(100, token);
        }
        if (image is null) throw new InvalidOperationException("Chart is unavailable for source capture.");
        var bounds = image.Current.BoundingRectangle;
        Cursor.Position = new((int)(bounds.Left + 15), (int)(bounds.Top + 15));
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var folder = Path.Combine(Path.GetTempPath(), "RSR-chart-source-" + Guid.NewGuid().ToString("N"));
        recorderType.GetMethod("Start")!.Invoke(recorder, [folder]);
        recorderType.GetMethod("StopChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(recorder, null);
        try
        {
            assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
            System.Drawing.Point? unnamedHit = null;
            foreach (var horizontal in new[] { .3, .5, .7 })
            {
                foreach (var vertical in new[] { .2, .4, .6 })
                {
                    var point = new System.Drawing.Point((int)(bounds.Left + bounds.Width * horizontal),
                        (int)(bounds.Top + bounds.Height * vertical));
                    var hit = AutomationElement.FromPoint(new(point.X, point.Y));
                    if (hit.Current.ControlType == ControlType.Image &&
                        string.IsNullOrWhiteSpace(hit.Current.Name) &&
                        TreeWalker.ControlViewWalker.GetParent(hit)?.Current.Name == "Chart Area")
                    {
                        unnamedHit = point;
                        break;
                    }
                }
                if (unnamedHit is not null) break;
            }
            Cursor.Position = unnamedHit ??
                throw new InvalidOperationException("Cold source check could not reproduce an unnamed inner Chart Area hit; no source gesture was sent.");
            await Task.Delay(100, token);
            if (recorderType.GetField("hoveredChart", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .GetValue(recorder) is not null)
                throw new InvalidOperationException("Cold chart-source check unexpectedly had a hover snapshot.");
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(200, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0008, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0010, 0, 0, 0, 0);
            await Task.Delay(400, token);
            AutomationElement? menu = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Select Data..."));
            var menuBounds = menu?.Current.BoundingRectangle ?? System.Windows.Rect.Empty;
            if (menuBounds.IsEmpty)
            {
                var probe = assembly.GetType("DesktopSteps.MsaaActions")!.GetMethod("Probe")!;
                var windows = AutomationElement.RootElement.FindAll(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ProcessIdProperty, window.Current.ProcessId));
                foreach (AutomationElement popup in windows)
                {
                    var observations = (System.Collections.IEnumerable)probe.Invoke(null,
                        [(nint)popup.Current.NativeWindowHandle])!;
                    foreach (var observation in observations)
                    {
                        var type = observation.GetType();
                        if (((string?)type.GetProperty("Name")!.GetValue(observation))?.Replace("&", "") != "Select Data...") continue;
                        menuBounds = new((int)type.GetProperty("X")!.GetValue(observation)!,
                            (int)type.GetProperty("Y")!.GetValue(observation)!,
                            (int)type.GetProperty("Width")!.GetValue(observation)!,
                            (int)type.GetProperty("Height")!.GetValue(observation)!);
                    }
                }
            }
            if (menuBounds.IsEmpty) throw new InvalidOperationException("Chart context menu did not expose Select Data through UIA or MSAA.");
            Cursor.Position = new((int)(menuBounds.Left + menuBounds.Width / 2), (int)(menuBounds.Top + menuBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(500, token);
            var dialog = AutomationElement.RootElement.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.NameProperty, "Select Data Source"))
                ?? throw new InvalidOperationException("Select Data Source did not open.");
            var edit = dialog.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit))
                .Cast<AutomationElement>().FirstOrDefault(element => element.Current.IsEnabled &&
                    element.TryGetCurrentPattern(ValuePattern.Pattern, out _))
                ?? throw new InvalidOperationException("Chart data range edit is unavailable.");
            var editBounds = edit.Current.BoundingRectangle;
            Cursor.Position = new((int)(editBounds.Left + editBounds.Width / 2), (int)(editBounds.Top + editBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(200, token);
            Cursor.Position = new((int)(editBounds.Right - 40), (int)(editBounds.Top + editBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            Cursor.Position = new((int)(editBounds.Left + 2), (int)(editBounds.Top + editBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(300, token);
            Cursor.Position = new((int)(editBounds.Right - 10), (int)(editBounds.Top + editBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(300, token);
            var firstCell = window.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "A1"),
                new PropertyCondition(AutomationElement.ClassNameProperty, "XLSpreadsheetCell")))
                ?? throw new InvalidOperationException("A1 is not exposed for chart range selection.");
            var lastCell = window.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "B20"),
                new PropertyCondition(AutomationElement.ClassNameProperty, "XLSpreadsheetCell")))
                ?? throw new InvalidOperationException("B20 is not exposed for chart range selection.");
            var firstBounds = firstCell.Current.BoundingRectangle;
            var lastBounds = lastCell.Current.BoundingRectangle;
            Cursor.Position = new((int)(firstBounds.Left + firstBounds.Width / 2), (int)(firstBounds.Top + firstBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            Cursor.Position = new((int)(lastBounds.Left + lastBounds.Width / 2), (int)(lastBounds.Top + lastBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(300, token);
            var collapsed = AutomationElement.RootElement.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.NameProperty, "Select Data Source"))!;
            var collapsedEdit = collapsed.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ClassNameProperty, "NetUIRefedit"))!;
            var collapsedBounds = collapsedEdit.Current.BoundingRectangle;
            var chosen = ((ValuePattern)collapsedEdit.GetCurrentPattern(ValuePattern.Pattern)).Current.Value;
            if (chosen != "=Sheet1!$A$1:$B$20")
                throw new InvalidOperationException("The live range picker did not replace the source: " + chosen);
            Cursor.Position = new((int)(collapsedBounds.Right - 10),
                (int)(collapsedBounds.Top + collapsedBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(300, token);
            dialog = AutomationElement.RootElement.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.NameProperty, "Select Data Source"))
                ?? throw new InvalidOperationException("Chart source dialog did not expand after range selection.");
            var ok = dialog.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "OK"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)))!;
            var okBounds = ok.Current.BoundingRectangle;
            Cursor.Position = new((int)(okBounds.Left + okBounds.Width / 2), (int)(okBounds.Top + okBounds.Height / 2));
            await Task.Delay(150, token);
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(100, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(2000, token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            AssertCapturedInputOrder(events);
            Console.WriteLine("Source after actual dialog: " + assembly.GetType("DesktopSteps.ExcelNativeSheet")!
                .GetMethod("ReadChartSource", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Fixture chart"]));
            if (events.SingleOrDefault(item => item.Kind == "context-click")?.Target is not
                { Name: "Fixture chart", ClassName: "ExcelChartObject", ParentName: "Sheet Sheet1" })
                throw new InvalidOperationException("The cold context-click did not freeze its named chart and worksheet identity.");
            var outcome = events.SingleOrDefault(item => item.Kind == "set-chart-source-range")
                ?? throw new InvalidOperationException("Actual chart source edit was not captured: " +
                    string.Join("; ", events.Select(item => $"{item.Kind} {item.Target?.Name} {item.Diagnostic}")));
            Console.WriteLine("Recorded source outcome: " + outcome.Value);
            var grouped = (List<RecordedEvent>)assembly.GetType("DesktopSteps.ActionGrouper")!
                .GetMethod("Group")!.Invoke(null, [events])!;
            Console.WriteLine("Grouped chart source: " + string.Join("; ", grouped.Select(item =>
                $"{item.Kind} {item.Target?.Name} {item.Target?.Window} {item.Diagnostic}")));
            if (grouped.Count != 1 || grouped[0].Kind != "set-chart-source-range")
                throw new InvalidOperationException("Chart edit did not consolidate into one replayable source outcome: " +
                    string.Join("; ", events.Select(item => $"{item.At:HH:mm:ss.fff} {item.Kind} {item.Target} {item.Diagnostic}")));
            chartObject.Chart.SetSourceData(app.ActiveSheet.Range["O1:P2"]);
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Captured chart source",
                grouped.Select((item, index) => new PlanStep(index + 1, item.Kind, item.Target,
                    item.Value, item.Key, item.AfterState, item.Screenshot, null)).ToList());
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [plan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), token])!).WaitAsync(token);
            var source = assembly.GetType("DesktopSteps.ExcelNativeSheet")!
                .GetMethod("ReadChartSource", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Fixture chart"])!;
            if ((string)source.GetType().GetProperty("Range")!.GetValue(source)! != "A1:B20")
                throw new InvalidOperationException("Replay did not restore the actually recorded chart source.");
            var extendedValues = new object?[25, 2];
            for (var row = 0; row < 25; row++)
            {
                extendedValues[row, 0] = "Category " + (row + 1);
                extendedValues[row, 1] = (double)(row + 1);
                app.ActiveSheet.Cells[row + 1, 1].Value2 = extendedValues[row, 0];
                app.ActiveSheet.Cells[row + 1, 2].Value2 = extendedValues[row, 1];
            }
            var native = assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            var pasteType = assembly.GetType("DesktopSteps.TablePaste")!;
            var paste = Activator.CreateInstance(pasteType, ["Sheet1", "A1", extendedValues]);
            native.GetMethod("ApplyChartSource", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Fixture chart", source, paste]);
            var extendedSource = native.GetMethod("ReadChartSource", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(object)app.ActiveSheet, "Fixture chart"])!;
            if ((string)extendedSource.GetType().GetProperty("Range")!.GetValue(extendedSource)! != "A1:B25")
                throw new InvalidOperationException("Captured source froze its endpoint instead of following the verified paste.");
            Console.WriteLine("PASS: an unnamed Chart Area hit with NO hover cache freezes its containing chart; real Select Data recording/replay restores all 20 rows.");
            Console.WriteLine("PASS: captured chart source follows all 25 rows of a matching verified paste.");
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            Marshal.FinalReleaseComObject(chartObject);
            if (Directory.Exists(folder))
            {
                foreach (var file in Directory.GetFiles(folder)) File.Delete(file);
                Directory.Delete(folder);
            }
        }
    }

    private static async Task CheckLegendDragAsync(dynamic app, dynamic sheet, AutomationElement window,
        Assembly assembly, CancellationToken token)
    {
        sheet.Activate();
        app.ActiveWindow.Zoom = 60;
        app.ActiveWindow.ScrollRow = 1;
        app.ActiveWindow.ScrollColumn = 1;
        dynamic chartObject = sheet.ChartObjects().Item("Fixture chart");
        chartObject.Left = 100;
        object anchor = sheet.Cells[2000, 1];
        try { chartObject.Top = (double)((dynamic)anchor).Top; }
        finally { Marshal.ReleaseComObject(anchor); }
        app.ActiveWindow.ScrollRow = 1995;
        dynamic chart = chartObject.Chart;
        chart.HasLegend = true;
        chartObject.Width = 700;
        chartObject.Height = 350;
        dynamic legend = chart.Legend;
        legend.Width = 40;
        chartObject.Width = 770;
        chartObject.Height = 385;
        legend.Height = 180;
        legend.Left = 450;
        legend.Top = 60;
        chartObject.Activate();
        legend.Select();
        assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
            BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
        chartObject.Activate();
        legend.Select();
        await Task.Delay(400, token);
        AutomationElement? image = null;
        for (var attempt = 0; image is null && attempt < 30; attempt++)
        {
            image = window.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "Fixture chart"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Image)));
            if (image is null) await Task.Delay(100, token);
        }
        if (image is null) throw new InvalidOperationException("Scratch chart was not exposed through UIA.");
        var bounds = image.Current.BoundingRectangle;
        var legendElements = image.FindAll(TreeScope.Descendants, Condition.TrueCondition)
            .Cast<AutomationElement>().Where(item =>
                item.Current.Name.Contains("Legend", StringComparison.OrdinalIgnoreCase)).ToList();
        Console.WriteLine("Live legend UIA names: " + string.Join("; ", legendElements.Select(item =>
            item.Current.Name)));
        var legendElement = legendElements.FirstOrDefault(item => item.Current.Name == "Legend")
            ?? throw new InvalidOperationException("No live Legend element was available; no drag was sent.");
        var native = assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
        var layout = native.GetMethod("ReadLegendLayout", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [(object)sheet, "Fixture chart"])!;
        double Dimension(string name) => (double)layout.GetType().GetProperty(name)!.GetValue(layout)!;
        var legendBounds = new System.Windows.Rect(bounds.Left + Dimension("Left") * bounds.Width,
            bounds.Top + Dimension("Top") * bounds.Height,
            Dimension("Width") * bounds.Width, Dimension("Height") * bounds.Height);
        var hitMethod = native.GetMethod("ChartNameAtPoint", BindingFlags.NonPublic | BindingFlags.Static)!;
        var hitName = (string?)hitMethod.Invoke(null,
            [(object)sheet, (int)(legendBounds.Left + legendBounds.Width / 2),
                (int)(legendBounds.Top + legendBounds.Height / 2)]);
        if (hitName != "Fixture chart")
            throw new InvalidOperationException("Native chart hit did not identify the legend's containing chart: " + hitName);
        var startX = (int)Math.Round(legendBounds.Right);
        var startY = (int)Math.Round(legendBounds.Top + legendBounds.Height / 2);
        var handleFound = false;
        foreach (var offset in Enumerable.Range(0, 21).Concat(Enumerable.Range(1, 20).Select(value => -value)))
        {
            Cursor.Position = new System.Drawing.Point(startX + offset, startY);
            await Task.Delay(60, token);
            var cursor = new CursorInfo { Size = Marshal.SizeOf<CursorInfo>() };
            if (GetCursorInfo(ref cursor) && cursor.Handle == LoadCursor(0, 32644))
            { startX += offset; handleFound = true; break; }
        }
        if (!handleFound) throw new InvalidOperationException(
            "No horizontal resize cursor on the live legend edge; no drag was sent.");
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var folder = Path.Combine(Path.GetTempPath(), "RSR-legend-" + Guid.NewGuid().ToString("N"));
        recorderType.GetMethod("Start")!.Invoke(recorder, [folder]);
        try
        {
            Cursor.Position = new System.Drawing.Point(startX, startY);
            await Task.Delay(700, token);
            assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
            chartObject.Activate();
            legend.Select();
            Cursor.Position = new System.Drawing.Point(startX, startY);
            await Task.Delay(250, token);
            if (Cursor.Position != new System.Drawing.Point(startX, startY))
                throw new InvalidOperationException("The desktop pointer moved before the legend gesture; no drag was sent.");
            var liveCursor = new CursorInfo { Size = Marshal.SizeOf<CursorInfo>() };
            if (!GetCursorInfo(ref liveCursor) || liveCursor.Handle != LoadCursor(0, 32644))
                throw new InvalidOperationException("The legend resize handle was no longer active after workbook activation; no drag was sent.");
            var chartSnapshot = recorderType.GetField("hoveredChart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(recorder);
            if (chartSnapshot is null)
                throw new InvalidOperationException($"No fresh chart identity was captured at the live resize edge {startX},{startY}; bounds {bounds}. No drag was sent.");
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(250, token);
            for (var offset = 10; offset <= 160; offset += 10)
            {
                Cursor.Position = new System.Drawing.Point(startX + offset, startY);
                await Task.Delay(30, token);
            }
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(500, token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            AssertCapturedInputOrder(events);
            var resized = events.SingleOrDefault(item => item.Kind == "set-chart-legend-layout")
                ?? throw new InvalidOperationException("Live legend drag did not capture a semantic resize: " +
                    string.Join("; ", events.Select(item => $"{item.Kind}: {item.Target?.Name}: {item.Diagnostic}")));
            if (resized.Target?.ClassName != "ExcelChartObject")
                throw new InvalidOperationException("Legend drag did not use native chart identity frozen at mouse-down.");
            if ((double)legend.Width <= 60) throw new InvalidOperationException("Live drag did not enlarge legend. " +
                $"Before: {resized.BeforeState}; after: {resized.Value}; bounds: {bounds}; start: {startX},{startY}");
            var beforeLayout = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, double>>(resized.BeforeState!)!;
            var afterLayout = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, double>>(resized.Value!)!;
            if (Math.Abs(beforeLayout["Left"] - afterLayout["Left"]) > .005 ||
                Math.Abs(beforeLayout["Top"] - afterLayout["Top"]) > .005)
                throw new InvalidOperationException("Resize test moved the legend instead of holding its position.");
            legend.Width = 40;
            var replay = new ExecutionPlan(1, DateTimeOffset.Now, "Recorded legend resize",
                [new(1, "click", resized.Target, null, null, resized.Target!.Name + ":selected", null, null),
                 new(2, resized.Kind, resized.Target, resized.Value, null, resized.AfterState, null, null)]);
            await ((Task<bool>)assembly.GetType("DesktopSteps.Executor")!.GetMethod("ReplayAsync")!
                .Invoke(null, [replay, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                            (Func<string, bool>)(_ => false), token])!).WaitAsync(token);
            if ((double)legend.Width <= 60) throw new InvalidOperationException("Replay did not restore captured legend width.");
            Console.WriteLine("PASS: real legend handle drag captures chart-relative layout, survives Stop, and replay restores widened legend from a narrow baseline.");
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            if (Directory.Exists(folder))
            {
                foreach (var file in Directory.GetFiles(folder)) File.Delete(file);
                Directory.Delete(folder);
            }
            Marshal.FinalReleaseComObject(legend);
            Marshal.FinalReleaseComObject(chart);
            Marshal.FinalReleaseComObject(chartObject);
        }
    }
    private static void AssertCapturedInputOrder(IReadOnlyList<RecordedEvent> events)
    {
        var inputs = events.Where(item => item.Kind is "click" or "context-click" or "key" or
            "chart-source-edit-input" or "unresolved-input").ToList();
        for (var index = 1; index < inputs.Count; index++)
            if (inputs[index].At < inputs[index - 1].At - TimeSpan.FromMilliseconds(2))
                throw new InvalidOperationException("Live capture reordered input: " +
                    $"{inputs[index - 1].Kind} at {inputs[index - 1].At:O} before " +
                    $"{inputs[index].Kind} at {inputs[index].At:O}.");
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct CursorInfo
    {
        public int Size;
        public int Flags;
        public nint Handle;
        public System.Drawing.Point Position;
    }
    [DllImport("user32.dll")]
    private static extern bool GetCursorInfo(ref CursorInfo info);
    [DllImport("user32.dll")]
    private static extern nint LoadCursor(nint instance, int resource);

    public static async Task RunFiltersAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        Exception? failure = null;
        nint scratchHandle = 0;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Value2 = "Owner");
            for (var row = 2; row <= 5; row++)
            {
                var r = row;
                RetryBusy(() => app.ActiveSheet.Cells[r, 1].Value2 = r % 2 == 0 ? "Alpha" : "Beta");
            }
            RetryBusy(() => app.ActiveSheet.Range["A1:A5"].AutoFilter());
            RetryBusy(() => app.WindowState = -4137);
            nint handle = 0;
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            scratchHandle = handle;
            await Task.Delay(500);
            var window = AutomationElement.FromHandle(handle);
            var target = new ControlRef("EXCEL", window.Current.Name, "Dropdown", "No filter applied",
                "ControlType.MenuItem", "", "A1", ProcessId: window.Current.ProcessId);
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            foreach (var mode in new[] { "only", "only", "exclude", "all" })
            {
                var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Scratch filter",
                    [new(1, "filter-values", target, "[\"Alpha\"]", null, "filter:" + mode, null, null)]);
                var messages = new List<string>();
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                        (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
                for (var row = 2; row <= 5; row++)
                {
                    var r = row;
                    bool hidden = false;
                    RetryBusy(() => hidden = (bool)app.ActiveSheet.Rows[r].Hidden);
                    var expectedHidden = mode == "only" ? r % 2 != 0 : mode == "exclude" && r % 2 == 0;
                    if (hidden != expectedHidden)
                        throw new InvalidOperationException($"Filter {mode}: row {r} had Hidden={hidden}.");
                }
                Console.WriteLine($"PASS: actual Excel filter {mode} verifies checkbox outcome, popup closure and hidden rows.");
            }
            var automation = typeof(ControlRef).Assembly.GetType("DesktopSteps.Automation")!;
            var dropdown = (AutomationElement)automation.GetMethod("ResolveMenuItem")!.Invoke(null, [target])!;
            object?[] clickArgs = [dropdown, null, false, null];
            if (!(bool)executor.GetMethod("TryClickLiveUiaElement", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, clickArgs)!)
                throw new InvalidOperationException("Scratch filter could not reopen for capture verification.");
            await Task.Delay(300, timeout.Token);
            var filterPlan = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelFilterPlan")!;
            var capture = filterPlan.GetMethod("CaptureState", BindingFlags.NonPublic | BindingFlags.Static)!;
            var alpha = target with
            {
                Name = "Alpha",
                AutomationId = "",
                ControlType = "ControlType.TreeItem",
                ParentName = "Manual Filter"
            };
            if ((string?)capture.Invoke(null, [alpha]) != "filter:checked")
                throw new InvalidOperationException("Recorder could not capture the filter's checked state.");
            var alphaElement = (AutomationElement)automation.GetMethod("Resolve")!.Invoke(null, [alpha, 2500])!;
            await ((Task)executor.GetMethod("SetExcelFilterCheckAsync", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [alphaElement, false, timeout.Token])!).WaitAsync(timeout.Token);
            if ((string?)capture.Invoke(null, [alpha]) != "filter:unchecked")
                throw new InvalidOperationException("Recorder could not capture the filter's unchecked state.");
            SendKeys.SendWait("{ESC}");
            await Task.Delay(200, timeout.Token);
            Console.WriteLine("PASS: recorder's post-click capture reads actual checked and unchecked Excel filter states.");
            await CheckRecordedFilterStateAsync(target, timeout.Token);
            var empty = new PlanStep(1, "filter-values", target, "[\"Alpha\",\"Beta\"]", null, "filter:exclude", null, null);
            try
            {
                await ((Task)executor.GetMethod("ApplyExcelFilterAsync", BindingFlags.NonPublic | BindingFlags.Static)!
                    .Invoke(null, [empty, timeout.Token])!).WaitAsync(timeout.Token);
                throw new InvalidOperationException("Empty filter was accepted.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("no selected values"))
            { Console.WriteLine("PASS: empty selection stops before disabled OK and worksheet input."); }
            SendKeys.SendWait("{ESC}");
        }
        catch (Exception ex) { failure = ex; throw; }
        finally
        {
            try
            {
                // Close only this fixture's popup and unsaved workbook.
                SetForegroundWindow(scratchHandle);
                SendKeys.SendWait("{ESC}");
                await Task.Delay(200);
                var window = AutomationElement.FromHandle(scratchHandle);
                var cancel = window.FindFirst(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, "Cancel"),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "NetUIButton")));
                if (cancel is not null && cancel.Current.IsEnabled &&
                    cancel.TryGetCurrentPattern(InvokePattern.Pattern, out var invoke))
                    ((InvokePattern)invoke).Invoke();
                RetryBusy(() => app.DisplayAlerts = false);
                if (workbook is not null)
                {
                    RetryBusy(() => workbook.Close(false));
                    Marshal.FinalReleaseComObject(workbook);
                }
                RetryBusy(() => app.Quit());
                Marshal.FinalReleaseComObject(app);
            }
            catch (Exception ex) when (failure is not null)
            { Console.Error.WriteLine("Scratch filter cleanup failed: " + ex.Message); }
        }
    }

    public static async Task RunReplacementsAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        AutomationElement? dialog = null;
        Exception? failure = null;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Value2 = "First source");
            RetryBusy(() => app.ActiveSheet.Cells[2, 1].Value2 = "Second source");
            RetryBusy(() => app.ActiveSheet.Range["A1:A2"].Select());
            nint handle = 0;
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            var id = AutomationElement.FromHandle(handle).Current.ProcessId;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            var command = Task.Run(() => RetryBusy(() => app.CommandBars.ExecuteMso("ReplaceDialog")));
            dialog = await FindWindowAsync(id, "Find and Replace", timeout.Token);
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            var apply = executor.GetMethod("ApplyRecordedReplacementAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
            var target = new ControlRef("EXCEL", "Find and Replace", null, "Replace All",
                "ControlType.Button", null, null, ProcessId: id);
            foreach (var (find, replacement) in new[] { ("First source", "First result"),
                ("Second source", "Second result") })
            {
                var step = new PlanStep(1, "replace-all", target, find, null, replacement, null, null);
                var result = await ((Task<string>)apply.Invoke(null, [step, timeout.Token, null])!).WaitAsync(timeout.Token);
                Console.WriteLine("PASS: actual Excel consecutive replacement result: " + result);
            }
            var messages = new List<string>();
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Replacement with legacy result acknowledgement",
            [
                new(1, "replace-all", target, "Missing source", null, "Unused result", null, null),
                new(2, "close-window", target, null, null, "window-closed", null, null),
                new(3, "click", target with { Window = "Microsoft Excel", Name = "OK" },
                    null, null, null, null, null)
            ]);
            await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            if (!messages.Any(message => message.Contains("find anything to replace", StringComparison.OrdinalIgnoreCase)) ||
                !messages.Any(message => message.StartsWith("Skipped step 3:")) ||
                messages.Any(message => message.StartsWith("Finding step 3:")))
                throw new InvalidOperationException("Legacy replay did not report no matches and skip the consumed result OK.");
            Console.WriteLine("PASS: actual Excel no-match warning, close, and legacy redundant OK replay.");
            dialog = null;
            await command.WaitAsync(timeout.Token);
            string first = "", second = "";
            RetryBusy(() => first = (string)app.ActiveSheet.Cells[1, 1].Value2);
            RetryBusy(() => second = (string)app.ActiveSheet.Cells[2, 1].Value2);
            if (first != "First result" || second != "Second result")
                throw new InvalidOperationException("Scratch replacements did not produce the requested values.");
            Console.WriteLine("PASS: actual Excel replacement dialog closed and both scratch cell values verified.");
        }
        catch (Exception ex) { failure = ex; throw; }
        finally
        {
            try
            {
                if (dialog is not null) await InvokeAsync(dialog, "Close");
                RetryBusy(() => app.DisplayAlerts = false);
                if (workbook is not null)
                {
                    RetryBusy(() => workbook.Close(false));
                    Marshal.FinalReleaseComObject(workbook);
                }
                RetryBusy(() => app.Quit());
                Marshal.FinalReleaseComObject(app);
            }
            catch (Exception ex) when (failure is not null)
            { Console.Error.WriteLine("Scratch Excel cleanup failed: " + ex.Message); }
        }
    }

    public static async Task RunCopyCaptureAsync(bool fastRelease = false)
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        var assembly = typeof(ControlRef).Assembly;
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var folder = Path.Combine(Path.GetTempPath(), "RSR-copy-capture-" + Guid.NewGuid().ToString("N"));
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Name = "Copy capture fixture");
            RetryBusy(() => app.WindowState = -4137);
            nint handle = 0;
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            SetForegroundWindow(handle);
            await Task.Delay(500);
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(45));
            var window = AutomationElement.FromHandle(handle);
            AutomationElement? tab = null;
            for (var attempt = 0; tab is null && attempt < 30; attempt++)
            {
                tab = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, "Copy capture fixture"));
                if (tab is null) await Task.Delay(100, timeout.Token);
            }
            if (tab is null)
            {
                foreach (AutomationElement candidate in window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)))
                    Console.WriteLine("Tab: " + candidate.Current.Name);
                throw new InvalidOperationException("Scratch worksheet tab was not found.");
            }
            var bounds = tab.Current.BoundingRectangle;
            Console.WriteLine($"Worksheet hit: {tab.Current.ControlType.ProgrammaticName} / {tab.Current.Name} / {bounds}");
            recorderType.GetMethod("Start")!.Invoke(recorder, [folder]);
            assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
            await Task.Delay(400, timeout.Token);
            Cursor.Position = new((int)(bounds.Left + bounds.Width / 2), (int)(bounds.Top + bounds.Height / 2));
            mouse_event(0x0008, 0, 0, 0, 0);
            await Task.Delay(150, timeout.Token);
            mouse_event(0x0010, 0, 0, 0, 0);
            AutomationElement? menuItem = null;
            for (var attempt = 0; menuItem is null && attempt < 60; attempt++)
            {
                var candidates = AutomationElement.RootElement.FindAll(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ProcessIdProperty, window.Current.ProcessId),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)));
                foreach (AutomationElement candidate in candidates)
                {
                    if (candidate.Current.Name.StartsWith("Move or Copy", StringComparison.Ordinal))
                        menuItem = candidate;
                }
                if (menuItem is null) await Task.Delay(100, timeout.Token);
            }
            if (menuItem is null) throw new InvalidOperationException("Move or Copy menu item was not found.");
            await Task.Delay(250, timeout.Token);
            bounds = menuItem.Current.BoundingRectangle;
            var x = (int)(bounds.Left + bounds.Width / 2);
            var y = (int)(bounds.Top + bounds.Height / 2);
            var savedMenus = (Array?)recorderType.GetField("pointerMenus",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder);
            if (savedMenus is null || savedMenus.Length == 0)
                throw new InvalidOperationException("Context menu was not frozen before selection.");
            var menuType = savedMenus.GetValue(0)!.GetType();
            var selected = savedMenus.Cast<object>().Single(item =>
            {
                var snapshot = menuType.GetProperty("Snapshot")!.GetValue(item)!;
                return ((ControlRef)snapshot.GetType().GetProperty("Target")!.GetValue(snapshot)!).Name == "Move or Copy...";
            });
            var at = (DateTimeOffset)menuType.GetProperty("At")!.GetValue(selected)!;
            var generation = (long)menuType.GetProperty("Generation")!.GetValue(selected)!;
            var foreground = (nint)menuType.GetProperty("Foreground")!.GetValue(selected)!;
            var freshness = recorderType.GetMethod("IsPointerMenuFresh", BindingFlags.NonPublic | BindingFlags.Static)!;
            bool Fresh(DateTimeOffset time, long inputGeneration, nint owner, int hitX, int hitY) =>
                (bool)freshness.Invoke(null, [selected, time, inputGeneration, owner, hitX, hitY])!;
            if (!Fresh(at, generation, foreground, x, y) ||
                Fresh(at.AddMilliseconds(1500), generation, foreground, x, y) ||
                Fresh(at.AddMilliseconds(-1), generation, foreground, x, y) ||
                Fresh(at, generation + 1, foreground, x, y) ||
                Fresh(at, generation, foreground + 1, x, y) ||
                Fresh(at, generation, foreground, -1, -1))
                throw new InvalidOperationException("Frozen menu freshness, input, window or point guards failed.");
            Cursor.Position = new(x, y);
            // The application can replace the popup before queued recorder work runs.
            var click = Task.Run(() =>
            {
                mouse_event(0x0002, 0, 0, 0, 0);
                if (!fastRelease) Thread.Sleep(150);
                mouse_event(0x0004, 0, 0, 0, 0);
            });
            Thread.Sleep(3000);
            await click.WaitAsync(timeout.Token);
            var dialog = await FindWindowAsync(window.Current.ProcessId, "Move or Copy", timeout.Token);
            await Task.Delay(700, timeout.Token);
            await InvokeAsync(dialog, "Cancel");
            await Task.Delay(500, timeout.Token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            foreach (var item in events)
                Console.WriteLine($"{item.Kind}: {item.Target?.Name} / {item.Target?.ControlType} / {item.Diagnostic}");
            foreach (var line in (IEnumerable<string>)recorderType.GetField("captureDiagnostics",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!)
                Console.WriteLine(line);
            if (!((IEnumerable<string>)recorderType.GetField("captureDiagnostics",
                    BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!)
                .Any(line => line.Contains("menu identity frozen before click")))
                throw new InvalidOperationException("Fast-click test did not exercise the pre-click menu snapshot.");
            if (!events.Any(item => item.Kind == "click" &&
                    item.Target is { Name: "Move or Copy...", ControlType: "ControlType.MenuItem" }) ||
                events.Any(item => item.Kind is "unresolved-click" or "unresolved-input" ||
                    item.Kind == "click" && item.Target is
                        { Window: "Move or Copy", ControlType: "ControlType.Window" }))
                throw new InvalidOperationException("Move or Copy lost its menu identity during the dialog transition.");
            Console.WriteLine($"PASS: {(fastRelease ? "fast press/release" : "150 ms press/release")} retains Move or Copy menu identity despite a three-second recorder UI stall.");
            assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
            await Task.Delay(300, timeout.Token);
            bounds = tab.Current.BoundingRectangle;
            Cursor.Position = new((int)(bounds.Left + bounds.Width / 2), (int)(bounds.Top + bounds.Height / 2));
            mouse_event(0x0008, 0, 0, 0, 0);
            await Task.Delay(150, timeout.Token);
            mouse_event(0x0010, 0, 0, 0, 0);
            await Task.Delay(400, timeout.Token);
            menuItem = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ProcessIdProperty, window.Current.ProcessId),
                new PropertyCondition(AutomationElement.NameProperty, "Move or Copy..."),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)))
                ?? throw new InvalidOperationException("Fallback menu fixture did not open.");
            bounds = menuItem.Current.BoundingRectangle;
            x = (int)(bounds.Left + bounds.Width / 2);
            y = (int)(bounds.Top + bounds.Height / 2);
            var fallback = recorderType.GetMethod("CaptureDirectMenuAtPointCore",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [x, y, false]);
            var captured = (ControlRef?)fallback?.GetType().GetProperty("Target")!.GetValue(fallback);
            if (captured is not { Name: "Move or Copy...", ControlType: "ControlType.MenuItem" })
                throw new InvalidOperationException("Missing MSAA hit did not resolve through the immediate popup-only UIA fallback.");
            await InvokeAsync(AutomationElement.RootElement, "Move or Copy...");
            dialog = await FindWindowAsync(window.Current.ProcessId, "Move or Copy", timeout.Token);
            if (recorderType.GetMethod("CaptureDirectMenuAtPointCore",
                    BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [x, y, false]) is not null)
                throw new InvalidOperationException("Popup menu fallback accepted a newly opened dialog as a menu command.");
            var pointType = recorderType.GetNestedType("Point", BindingFlags.NonPublic)!;
            var point = Activator.CreateInstance(pointType)!;
            pointType.GetField("X")!.SetValue(point, x);
            pointType.GetField("Y")!.SetValue(point, y);
            var infoType = recorderType.GetNestedType("MouseInfo", BindingFlags.NonPublic)!;
            var info = Activator.CreateInstance(infoType)!;
            infoType.GetField("Point")!.SetValue(info, point);
            using (var delayedRecorder = (IDisposable)Activator.CreateInstance(recorderType)!)
            {
                recorderType.GetMethod("ProcessMouse", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .Invoke(delayedRecorder, [0, (nint)0x201, info, null, null, captured, fallback, DateTimeOffset.Now]);
                var delayedEvents = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(delayedRecorder)!;
                if (delayedEvents.Count != 1 || delayedEvents[0].Target != captured ||
                    (bool)recorderType.GetMethod("IsUnresolvedDialogClick")!.Invoke(delayedRecorder, [0])!)
                    throw new InvalidOperationException("Frozen popup fallback was replaced by the newly opened dialog.");
            }
            await InvokeAsync(dialog, "Cancel");
            Console.WriteLine("PASS: unavailable MSAA menu hit uses immediate UIA fallback; delayed delivery preserves the command instead of the open dialog.");
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            SetForegroundWindow((nint)(int)app.Hwnd);
            SendKeys.SendWait("{ESC}{ESC}");
            await Task.Delay(200);
            RetryBusy(() => app.DisplayAlerts = false);
            if (workbook is not null)
            {
                RetryBusy(() => workbook.Close(false));
                Marshal.FinalReleaseComObject(workbook);
            }
            RetryBusy(() => app.Quit());
            Marshal.FinalReleaseComObject(app);
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }

    public static async Task RunRowMenuCaptureAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        var assembly = typeof(ControlRef).Assembly;
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var folder = Path.Combine(Path.GetTempPath(), "RSR-row-menu-capture-" + Guid.NewGuid().ToString("N"));
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Range("A2").Value2 = "row two");
            RetryBusy(() => app.WindowState = -4137);
            nint handle = 0;
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            SetForegroundWindow(handle);
            await Task.Delay(500);
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(120));
            var window = AutomationElement.FromHandle(handle);
            AutomationElement? header = null;
            for (var attempt = 0; header is null && attempt < 30; attempt++)
            {
                header = window.FindFirst(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "2"),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "XLGridRowHeader")));
                if (header is null) await Task.Delay(100, timeout.Token);
            }
            if (header is null) throw new InvalidOperationException("Row header 2 was not found.");
            var bounds = header.Current.BoundingRectangle;
            recorderType.GetMethod("Start")!.Invoke(recorder, [folder]);
            assembly.GetType("DesktopSteps.Executor")!.GetMethod("ActivateWindow",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [window]);
            await Task.Delay(400, timeout.Token);
            Cursor.Position = new((int)(bounds.Right - 6), (int)(bounds.Top + bounds.Height / 2));
            mouse_event(0x0008, 0, 0, 0, 0);
            await Task.Delay(150, timeout.Token);
            mouse_event(0x0010, 0, 0, 0, 0);
            AutomationElement? menuItem = null;
            for (var attempt = 0; menuItem is null && attempt < 30; attempt++)
            {
                await Task.Delay(200, timeout.Token);
                var popup = GetForegroundWindow();
                if (popup == handle) continue;
                menuItem = AutomationElement.FromHandle(popup).FindFirst(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Insert"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)));
            }
            // The search-box row menu becomes the foreground window; the recorder must treat it as the menu scope.
            if (menuItem is null) throw new InvalidOperationException("Row menu Insert item was not found in a foreground menu popup.");
            Console.WriteLine($"Foreground while row menu is open: {GetForegroundWindow():X} (Excel {handle:X})");
            await Task.Delay(250, timeout.Token);
            bounds = menuItem.Current.BoundingRectangle;
            Cursor.Position = new((int)(bounds.Left + bounds.Width / 2), (int)(bounds.Top + bounds.Height / 2));
            var click = Task.Run(() =>
            {
                mouse_event(0x0002, 0, 0, 0, 0);
                mouse_event(0x0004, 0, 0, 0, 0);
            });
            Thread.Sleep(3000);
            await click.WaitAsync(timeout.Token);
            await Task.Delay(700, timeout.Token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            foreach (var item in events)
                Console.WriteLine($"{item.Kind}: {item.Target?.Name} / {item.Target?.ControlType} / {item.Diagnostic}");
            foreach (var line in (IEnumerable<string>)recorderType.GetField("captureDiagnostics",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!)
                Console.WriteLine(line);
            string? moved = null, original = null;
            for (var attempt = 0; moved != "row two" && attempt < 10; attempt++)
            {
                await Task.Delay(300, timeout.Token);
                RetryBusy(() => moved = Convert.ToString(app.ActiveSheet.Range("A3").Value2));
                RetryBusy(() => original = Convert.ToString(app.ActiveSheet.Range("A2").Value2));
            }
            if (moved != "row two")
                throw new InvalidOperationException($"Insert was not executed on row 2 (A2='{original}', A3='{moved}').");
            if (!events.Any(item => item.Kind == "click" &&
                    item.Target is { Name: "Insert", ControlType: "ControlType.MenuItem" }) ||
                events.Any(item => item.Kind is "unresolved-click" or "unresolved-input"))
                throw new InvalidOperationException("Row header menu choice lost its identity.");
            Console.WriteLine("PASS: row-header context menu choice keeps its menu identity despite a three-second recorder UI stall.");
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            SetForegroundWindow((nint)(int)app.Hwnd);
            SendKeys.SendWait("{ESC}{ESC}");
            await Task.Delay(200);
            RetryBusy(() => app.DisplayAlerts = false);
            if (workbook is not null)
            {
                RetryBusy(() => workbook.Close(false));
                Marshal.FinalReleaseComObject(workbook);
            }
            RetryBusy(() => app.Quit());
            Marshal.FinalReleaseComObject(app);
            if (Directory.Exists(folder)) Directory.Delete(folder, true);
        }
    }

    public static async Task RunAsync()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("Excel is not installed.");
        dynamic app = Activator.CreateInstance(excelType)!;
        dynamic? workbook = null;
        Exception? failure = null;
        AutomationElement? openDialog = null;
        string? cleanupCommand = null;
        try
        {
            RetryBusy(() => app.Visible = true);
            RetryBusy(() => workbook = app.Workbooks.Add());
            RetryBusy(() => app.ActiveSheet.Cells[1, 1].Value2 = "Key");
            RetryBusy(() => app.ActiveSheet.Cells[1, 2].Value2 = "Value");
            RetryBusy(() => app.ActiveSheet.Cells[2, 1].Value2 = "Duplicate");
            RetryBusy(() => app.ActiveSheet.Cells[3, 1].Value2 = "Duplicate");
            RetryBusy(() => app.ActiveSheet.Cells[2, 2].Value2 = "First");
            RetryBusy(() => app.ActiveSheet.Cells[3, 2].Value2 = "Second");
            RetryBusy(() => app.ActiveSheet.Range["A1:A3"].Select());
            nint handle = 0;
            RetryBusy(() => handle = (nint)(int)app.Hwnd);
            var window = AutomationElement.FromHandle(handle);
            var processId = window.Current.ProcessId;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            var command = Task.Run(() => RetryBusy(() => app.CommandBars.ExecuteMso("RemoveDuplicates")));
            var warning = await FindWindowAsync(processId, "Remove Duplicates Warning", timeout.Token);
            openDialog = warning;
            cleanupCommand = "Cancel";
            CheckCoarseHits(warning);
            await CheckRecordedDuplicateFlowAsync(warning, processId, timeout.Token);
            openDialog = null;
            await command.WaitAsync(timeout.Token);
            Console.WriteLine("PASS: actual Excel Remove Duplicates warning, command and result clicks are recorded without unresolved targets.");

            var replaceCommand = Task.Run(() => RetryBusy(() => app.CommandBars.ExecuteMso("ReplaceDialog")));
            var replace = await FindWindowAsync(processId, "Find and Replace", timeout.Token);
            openDialog = replace;
            cleanupCommand = "Close";
            CheckCoarseHits(replace);
            await CheckRecordedReplaceFlowAsync(replace, processId, timeout.Token);
            openDialog = null;
            await replaceCommand.WaitAsync(timeout.Token);
            Console.WriteLine("PASS: actual Excel Find/Replace tab, fields, Replace All, acknowledgement and Close are recorded.");

            RetryBusy(() => app.ActiveSheet.Range["A1:B3"].AutoFilter());
            RetryBusy(() => app.WindowState = -4137);
            AutomationElementCollection? dropdowns = null;
            for (var attempt = 0; attempt < 50; attempt++)
            {
                await Task.Delay(100, timeout.Token);
                RetryBusy(() => handle = (nint)(int)app.Hwnd);
                window = AutomationElement.FromHandle(handle);
                dropdowns = window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "Dropdown"));
                if (dropdowns.Count == 2) break;
            }
            var recorder = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
            var capture = recorder.GetMethod("CaptureMouseDownTarget", BindingFlags.NonPublic | BindingFlags.Static)!;
            var count = 0;
            foreach (AutomationElement dropdown in dropdowns!)
            {
                var state = dropdown.Current;
                if (state.IsOffscreen || !state.IsEnabled || state.BoundingRectangle.IsEmpty) continue;
                var parent = TreeWalker.ControlViewWalker.GetParent(dropdown);
                var bounds = state.BoundingRectangle;
                var captured = capture.Invoke(null, [(int)(bounds.Left + bounds.Width / 2),
                    (int)(bounds.Top + bounds.Height / 2)]);
                var target = (ControlRef?)captured?.GetType().GetProperty("Target")!.GetValue(captured);
                if (target?.AutomationId != "Dropdown" || string.IsNullOrWhiteSpace(target.ParentName) ||
                    target.ParentName != parent?.Current.Name)
                    throw new InvalidOperationException("Excel dropdown capture did not retain its column parent.");
                var automation = typeof(ControlRef).Assembly.GetType("DesktopSteps.Automation")!;
                var resolved = (AutomationElement?)automation.GetMethod("ResolveMenuItem")!.Invoke(null, [target]);
                if (resolved is null || TreeWalker.ControlViewWalker.GetParent(resolved)?.Current.Name != target.ParentName)
                    throw new InvalidOperationException("Captured filter did not resolve within its recorded column.");
                count++;
            }
            if (count != 2) throw new InvalidOperationException($"Expected two identifiable scratch filters, found {count}.");
            Console.WriteLine("PASS: actual Excel filter captures preserve both distinct column parents.");
        }
        catch (Exception ex)
        {
            failure = ex;
            throw;
        }
        finally
        {
            try
            {
                if (openDialog is not null && cleanupCommand is not null)
                {
                    try { await InvokeAsync(openDialog, cleanupCommand); }
                    catch (Exception ex) when (failure is not null)
                    { Console.Error.WriteLine("Scratch dialog cleanup failed: " + ex.Message); }
                }
                RetryBusy(() => app.DisplayAlerts = false);
                if (workbook is not null)
                {
                    RetryBusy(() => workbook.Close(false));
                    Marshal.FinalReleaseComObject(workbook);
                }
                RetryBusy(() => app.Quit());
                Marshal.FinalReleaseComObject(app);
            }
            catch (Exception ex) when (failure is not null)
            { Console.Error.WriteLine("Excel fixture cleanup failed: " + ex.Message); }
        }
    }

    private static void RetryBusy(Action action)
    {
        for (var attempt = 0; ; attempt++)
        {
            try { action(); return; }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80010001) && attempt < 30)
            { Thread.Sleep(100); }
        }
    }

    private static async Task CheckRecordedFilterStateAsync(ControlRef target, CancellationToken token)
    {
        var assembly = typeof(ControlRef).Assembly;
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        var automation = assembly.GetType("DesktopSteps.Automation")!;
        var executor = assembly.GetType("DesktopSteps.Executor")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var output = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rsr-filter-capture-" + Guid.NewGuid().ToString("N"));
        var cursor = Cursor.Position;
        recorderType.GetMethod("Start")!.Invoke(recorder, [output]);
        try
        {
            var dropdown = (AutomationElement)automation.GetMethod("ResolveMenuItem")!.Invoke(null, [target])!;
            await ClickAsync(dropdown);
            var all = target with
            {
                Name = "(Select All)",
                AutomationId = "",
                ControlType = "ControlType.TreeItem",
                ClassName = "",
                ParentName = "Manual Filter"
            };
            var alpha = all with { Name = "Alpha" };
            var allElement = (AutomationElement)automation.GetMethod("Resolve")!.Invoke(null, [all, 2500])!;
            await ((Task)executor.GetMethod("SetExcelFilterCheckAsync", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [allElement, true, token])!).WaitAsync(token);
            await ClickAsync(allElement);
            await ClickAsync((AutomationElement)automation.GetMethod("Resolve")!.Invoke(null, [alpha, 2500])!);
            var ok = (AutomationElement)automation.GetMethod("Resolve")!.Invoke(null,
                [target with { Name = "OK", AutomationId = "", ControlType = "ControlType.Button",
                    ParentName = null }, 2500])!;
            var okBounds = ok.Current.BoundingRectangle;
            var okX = (int)(okBounds.Left + okBounds.Width / 2);
            var okY = (int)(okBounds.Top + okBounds.Height / 2);
            var commandSnapshot = recorderType.GetMethod("CaptureOwnedPopupCommandAtPoint",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [okX, okY]);
            var commandTarget = (ControlRef?)commandSnapshot?.GetType().GetProperty("Target")!.GetValue(commandSnapshot);
            if (commandTarget is not { Name: "OK", ControlType: "ControlType.Button" } ||
                commandTarget.Process != target.Process || commandTarget.Window != target.Window)
                throw new InvalidOperationException("Filter confirmation did not freeze through generic owned-popup MSAA capture.");
            Cursor.Position = new(okX, okY);
            var sendConfirmation = Task.Run(() =>
            {
                mouse_event(0x0002, 0, 0, 0, 0);
                Thread.Sleep(150);
                mouse_event(0x0004, 0, 0, 0, 0);
            });
            Thread.Sleep(3000);
            await sendConfirmation.WaitAsync(token);
            await Task.Delay(750, token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            if (!events.Any(item => item.Target?.Name == "(Select All)" && item.AfterState == "filter:unchecked") ||
                !events.Any(item => item.Target?.Name == "Alpha" && item.AfterState == "filter:checked") ||
                !events.Any(item => item.Target?.Name == "OK" && item.Kind == "click") ||
                events.Any(item => item.Kind == "unresolved-input"))
                throw new InvalidOperationException("Actual recorded checklist clicks lost final checkbox states: " +
                    System.Text.Json.JsonSerializer.Serialize(events));
            Console.WriteLine("PASS: actual filter checklist states and OK survive a 3-second recorder UI stall with a frozen popup command.");
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            SendKeys.SendWait("{ESC}");
            Cursor.Position = cursor;
            foreach (var file in System.IO.Directory.GetFiles(output)) System.IO.File.Delete(file);
            System.IO.Directory.Delete(output);
        }

        async Task ClickAsync(AutomationElement element)
        {
            var bounds = element.Current.BoundingRectangle;
            Cursor.Position = new System.Drawing.Point((int)(bounds.Left +
                (element.Current.ControlType == ControlType.TreeItem ? Math.Min(8, bounds.Width / 2) : bounds.Width / 2)),
                (int)(bounds.Top + bounds.Height / 2));
            if (element.Current.ControlType == ControlType.TreeItem)
            {
                var snapshot = recorderType.GetMethod("CaptureFilterItemAtPoint",
                    BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [Cursor.Position.X, Cursor.Position.Y]);
                var frozen = (ControlRef?)snapshot?.GetType().GetProperty("Target")!.GetValue(snapshot);
                if (frozen?.Name != element.Current.Name || frozen.ParentName != "Manual Filter" ||
                    frozen.ControlType != "ControlType.TreeItem")
                    throw new InvalidOperationException("Mouse-down MSAA filter identity did not capture the checkbox before queued processing. " +
                        $"Expected={element.Current.Name}; snapshot={System.Text.Json.JsonSerializer.Serialize(frozen)}; " +
                        $"nativeName={assembly.GetType("DesktopSteps.MsaaActions")!.GetMethod("FilterItemNameAtPoint", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [Cursor.Position.X, Cursor.Position.Y])}");
            }
            mouse_event(0x0002, 0, 0, 0, 0);
            await Task.Delay(150, token);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(750, token);
        }
    }

    private static async Task CheckRecordedDuplicateFlowAsync(AutomationElement warning, int processId,
        CancellationToken token)
    {
        var recorderType = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var output = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rsr-duplicate-capture-" + Guid.NewGuid().ToString("N"));
        var notices = 0;
        Action<int, string> notice = (_, _) => notices++;
        recorderType.GetEvent("UnresolvedDialogClick")!.AddEventHandler(recorder, notice);
        recorderType.GetMethod("Start")!.Invoke(recorder, [output]);
        var previousCursor = Cursor.Position;
        try
        {
            await Task.Delay(600, token);
            var radio = warning.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton))
                .Cast<AutomationElement>().Single(control =>
                    control.Current.Name.Contains("Continue", StringComparison.OrdinalIgnoreCase));
            await ClickAsync(radio);
            var continueButton = warning.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button))
                .Cast<AutomationElement>().Single(control =>
                    control.Current.Name.Replace("&", "").StartsWith("Remove Duplicates", StringComparison.Ordinal));
            await ClickAsync(continueButton);
            var dialog = await FindWindowAsync(processId, "Remove Duplicates", token);
            await ClickAsync(dialog.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "OK")));
            var result = await FindWindowAsync(processId, "Microsoft Excel", token);
            await Task.Delay(600, token);
            await ClickAsync(result.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "OK")));
            await Task.Delay(5500, token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            if (notices != 0 || events.Any(item => item.Kind is "unresolved-input" or "unresolved-click") ||
                !events.Any(item => item.Target is
                {
                    Window: "Remove Duplicates Warning",
                    ControlType: "ControlType.RadioButton"
                }) ||
                !events.Any(item => item.Target is
                {
                    Window: "Remove Duplicates Warning",
                    ControlType: "ControlType.Button"
                } && item.Target.Name?.StartsWith("Remove Duplicates") == true) ||
                !events.Any(item => item.Target is
                {
                    Window: "Remove Duplicates", Name: "OK",
                    ControlType: "ControlType.Button"
                }) ||
                !events.Any(item => item.Target is
                {
                    Window: "Microsoft Excel", Name: "OK",
                    ControlType: "ControlType.Button"
                }))
                throw new InvalidOperationException("The real Remove Duplicates flow lost a command identity: " +
                    System.Text.Json.JsonSerializer.Serialize(events));
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            Cursor.Position = previousCursor;
            foreach (var file in System.IO.Directory.GetFiles(output)) System.IO.File.Delete(file);
            System.IO.Directory.Delete(output);
        }

        static async Task ClickAsync(AutomationElement? control)
        {
            if (control is null) throw new InvalidOperationException("Scratch dialog command is missing.");
            var bounds = control.Current.BoundingRectangle;
            if (!control.Current.IsEnabled || control.Current.IsOffscreen || bounds.IsEmpty)
                throw new InvalidOperationException("Scratch dialog command is not clickable.");
            Cursor.Position = new System.Drawing.Point((int)(bounds.Left + bounds.Width * 0.75),
                (int)(bounds.Top + bounds.Height / 2));
            mouse_event(0x0002, 0, 0, 0, 0);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(600);
        }
    }

    private static async Task CheckRecordedReplaceFlowAsync(AutomationElement dialog, int processId,
        CancellationToken token)
    {
        var recorderType = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        var output = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "rsr-replace-capture-" + Guid.NewGuid().ToString("N"));
        var notices = 0;
        Action<int, string> notice = (_, _) => notices++;
        recorderType.GetEvent("UnresolvedDialogClick")!.AddEventHandler(recorder, notice);
        recorderType.GetMethod("Start")!.Invoke(recorder, [output]);
        var previousCursor = Cursor.Position;
        try
        {
            await Task.Delay(600, token);
            await ClickNamedAsync("Find", ControlType.TabItem);
            await ClickNamedAsync("Replace", ControlType.TabItem);
            var fields = dialog.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ClassNameProperty, "EDTBX"))
                .Cast<AutomationElement>().Where(element => !element.Current.IsOffscreen)
                .OrderBy(element => element.Current.BoundingRectangle.Top).ToArray();
            if (fields.Length != 2) throw new InvalidOperationException("Replace fixture fields are missing.");
            await ClickControlAsync(fields[0]);
            SendKeys.SendWait("^a");
            SendKeys.SendWait("Duplicate");
            await Task.Delay(500, token);
            await ClickControlAsync(fields[1]);
            SendKeys.SendWait("^a");
            SendKeys.SendWait("Unique");
            await Task.Delay(500, token);
            await ClickNamedAsync("Replace All", ControlType.Button);
            var result = await FindWindowAsync(processId, "Microsoft Excel", token);
            await Task.Delay(600, token);
            await ClickControlAsync(result.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "OK"))!);
            await ClickNamedAsync("Close", ControlType.Button);
            await Task.Delay(5500, token);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            if (notices != 0 || events.Any(item => item.Kind is "unresolved-input" or "unresolved-click") ||
                events.Any(item => item.Target?.Name == "Find and Replace" && item.Kind == "click") ||
                !events.Any(item => item.Target is
                {
                    Name: "Replace All", Window: "Find and Replace",
                    ControlType: "ControlType.Button"
                }) ||
                !events.Any(item => item.Target is { Name: "Close", Window: "Find and Replace" }) ||
                !events.Any(item => item.Kind == "key" && item.Target?.ClassName == "EDTBX"))
                throw new InvalidOperationException("Find and Replace recording lost a command or field: " +
                    System.Text.Json.JsonSerializer.Serialize(events));
        }
        finally
        {
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            Cursor.Position = previousCursor;
            foreach (var file in System.IO.Directory.GetFiles(output)) System.IO.File.Delete(file);
            System.IO.Directory.Delete(output);
        }

        Task ClickNamedAsync(string name, ControlType type) => ClickControlAsync(
            dialog.FindFirst(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, name),
                new PropertyCondition(AutomationElement.ControlTypeProperty, type)))
            ?? throw new InvalidOperationException($"Scratch command '{name}' is missing."));

        async Task ClickControlAsync(AutomationElement control)
        {
            var bounds = control.Current.BoundingRectangle;
            Cursor.Position = new System.Drawing.Point((int)(bounds.Left + bounds.Width * 0.6),
                (int)(bounds.Top + bounds.Height / 2));
            mouse_event(0x0002, 0, 0, 0, 0);
            mouse_event(0x0004, 0, 0, 0, 0);
            await Task.Delay(700, token);
        }
    }

    private static async Task<AutomationElement> FindWindowAsync(int processId, string name, CancellationToken token)
    {
        for (var attempt = 0; attempt < 100; attempt++)
        {
            token.ThrowIfCancellationRequested();
            var roots = AutomationElement.RootElement.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ProcessIdProperty, processId));
            foreach (AutomationElement root in roots)
            {
                if (root.Current.Name == name) return root;
                var nested = root.FindFirst(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                    new PropertyCondition(AutomationElement.NameProperty, name)));
                if (nested is not null) return nested;
            }
            await Task.Delay(100, token);
        }
        throw new InvalidOperationException($"Excel fixture dialog '{name}' did not appear.");
    }

    private static void CheckCoarseHits(AutomationElement window)
    {
        var recorder = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
        var snapshot = recorder.GetMethod("CreateMouseDownSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
        var controls = window.FindAll(TreeScope.Descendants, Condition.TrueCondition);
        var checkedControls = 0;
        foreach (AutomationElement control in controls)
        {
            var current = control.Current;
            if (current.ControlType != ControlType.Button && current.ControlType != ControlType.RadioButton)
                continue;
            if (!current.IsEnabled || current.IsOffscreen || current.BoundingRectangle.IsEmpty) continue;
            var bounds = current.BoundingRectangle;
            var captured = snapshot.Invoke(null, [window, (nint)window.Current.NativeWindowHandle,
                (int)(bounds.Left + bounds.Width / 2), (int)(bounds.Top + bounds.Height / 2)]);
            var target = (ControlRef?)captured?.GetType().GetProperty("Target")!.GetValue(captured);
            if (target?.Name != current.Name.Replace("&", "") ||
                target.ControlType != current.ControlType.ProgrammaticName || target.Window != window.Current.Name)
                throw new InvalidOperationException($"Excel control '{current.Name}' ({current.ControlType.ProgrammaticName}) " +
                    $"in '{window.Current.Name}' was captured as '{target?.Name}' ({target?.ControlType}) in '{target?.Window}'.");
            checkedControls++;
        }
        if (checkedControls == 0) throw new InvalidOperationException("Excel dialog exposed no command controls.");
        var handle = (nint)window.Current.NativeWindowHandle;
        if (handle == 0)
        {
            var state = window.Current;
            var identity = new ControlRef(System.Diagnostics.Process.GetProcessById(state.ProcessId).ProcessName,
                state.Name, null, state.Name, "ControlType.Window", null, null);
            handle = (nint)recorder.GetMethod("ResolveDialogWindow",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [(nint)0, identity, null, null])!;
        }
        if (handle == 0)
            throw new InvalidOperationException("Scratch Excel dialog has no identifiable native window.");
        var foregroundThread = GetWindowThreadProcessId(GetForegroundWindow(), out _);
        var currentThread = GetCurrentThreadId();
        var attached = foregroundThread != currentThread &&
            AttachThreadInput(currentThread, foregroundThread, true);
        try
        {
            ShowWindow(handle, 9);
            SetWindowPos(handle, (nint)(-1), 0, 0, 0, 0, 0x0043);
            SetForegroundWindow(handle);
        }
        finally
        {
            if (attached) AttachThreadInput(currentThread, foregroundThread, false);
        }
        for (var attempt = 0; attempt < 20 && GetForegroundWindow() != handle; attempt++)
        {
            SetForegroundWindow(handle);
            Application.DoEvents();
            Thread.Sleep(50);
        }
        if (GetForegroundWindow() != handle)
            throw new InvalidOperationException($"Scratch Excel dialog could not acquire foreground focus: " +
                $"handle {handle}, root {GetAncestor(handle, 2)}, foreground {GetForegroundWindow()}, " +
                $"dialog '{window.Current.Name}', attached {attached}.");
        var liveCapture = recorder.GetMethod("CaptureMouseDownTarget", BindingFlags.NonPublic | BindingFlags.Static)!;
        foreach (AutomationElement control in controls)
        {
            var current = control.Current;
            if (current.ControlType != ControlType.Button && current.ControlType != ControlType.RadioButton)
                continue;
            if (!current.IsEnabled || current.IsOffscreen || current.BoundingRectangle.IsEmpty) continue;
            // Include option-label clicks, not just the small radio glyph.
            var bounds = current.BoundingRectangle;
            var pointX = (int)(bounds.Left + bounds.Width * 0.75);
            var pointY = (int)(bounds.Top + bounds.Height / 2);
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var filterHit = recorder.GetMethod("CaptureFilterItemAtPoint",
                BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [pointX, pointY]);
            if (filterHit is not null || watch.ElapsedMilliseconds >= 250)
                throw new InvalidOperationException($"Filter probe entered dialog '{window.Current.Name}' " +
                    $"at '{current.Name}' or took {watch.ElapsedMilliseconds} ms; expected rejection below 250 ms.");
            var captured = liveCapture.Invoke(null, [pointX, pointY]);
            var target = (ControlRef?)captured?.GetType().GetProperty("Target")!.GetValue(captured);
            if (target?.Name != current.Name.Replace("&", "") ||
                target.ControlType != current.ControlType.ProgrammaticName || target.Window != window.Current.Name)
            {
                var hit = AutomationElement.FromPoint(new System.Windows.Point(pointX, pointY)).Current;
                throw new InvalidOperationException($"Actual mouse-down lookup failed for '{current.Name}': " +
                    $"captured '{target?.Name}' ({target?.ControlType}) in '{target?.Window}'; " +
                    $"live hit '{hit.Name}' ({hit.ControlType.ProgrammaticName}), class {hit.ClassName}, " +
                    $"process {hit.ProcessId}, point {pointX},{pointY}, expected bounds {bounds}.");
            }

        }
    }

    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(nint window);
    [DllImport("user32.dll")] private static extern nint GetForegroundWindow();
    [DllImport("user32.dll")] private static extern nint GetAncestor(nint window, uint flag);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(nint window, nint after, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint thread, uint other, bool attach);
    [DllImport("user32.dll")] private static extern bool ShowWindow(nint window, int command);
    [DllImport("user32.dll")] private static extern void mouse_event(uint flags, uint x, uint y, uint data, nuint extraInfo);
    [DllImport("user32.dll")] private static extern void keybd_event(byte key, byte scan, uint flags, nuint extra);

    private static async Task InvokeAsync(AutomationElement window, string name)
    {
        var control = window.FindFirst(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.NameProperty, name))
            ?? throw new InvalidOperationException($"Excel fixture command '{name}' missing.");
        await Task.Run(() => ((InvokePattern)control.GetCurrentPattern(InvokePattern.Pattern)).Invoke());
    }
}
