using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Automation;
using DesktopSteps;

internal static class Program
{
    private const string HostTitle = "RSR menu regression host";

    [STAThread]
    private static int Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        if (args.Contains("--host"))
        {
            if (args.Contains("--dialog-capture"))
            {
                using var dialog = new Form { Text = "RSR transient dialog fixture", Width = 360, Height = 230 };
                dialog.Controls.Add(new Button { Text = "OK", Left = 40, Top = 110, Width = 100,
                    DialogResult = DialogResult.OK });
                dialog.Controls.Add(new RadioButton { Text = "Continue with selection", Left = 40,
                    Top = 40, Width = 230 });
                Application.Run(dialog);
                return 0;
            }
            using var form = new Form { Text = HostTitle, Width = 600, Height = 400 };
            if (args.Contains("--refresh-watch"))
            {
                var refreshButton = new Button { Text = "Refresh View", Name = "refreshButton", Dock = DockStyle.Top };
                var otherButton = new Button { Text = "Other control", Dock = DockStyle.Bottom };
                using var refreshTimer = new System.Windows.Forms.Timer { Interval = 1200 };
                refreshTimer.Tick += (_, _) =>
                {
                    refreshTimer.Stop();
                    refreshButton.Enabled = true;
                };
                refreshButton.Click += (_, _) =>
                {
                    refreshButton.Enabled = false;
                    otherButton.Focus();
                    refreshTimer.Start();
                };
                form.Controls.Add(refreshButton);
                form.Controls.Add(otherButton);
                Application.Run(form);
                return 0;
            }
            if (args.Contains("--refresh-choice"))
            {
                var unrelated = new Button { Text = "Refresh", Name = "unrelatedRefresh",
                    Left = 30, Top = 25, Width = 130 };
                var desired = new Button { Text = "Refresh View", Name = "viewRefresh",
                    Left = 30, Top = 190, Width = 130 };
                var anchor = new Label { Text = "Case Age Tag", Name = "viewAnchor",
                    Left = 30, Top = 235, Width = 150 };
                var grid = new DataGridView { Name = "triageGrid", Left = 20, Top = 265,
                    Width = 520, Height = 110 };
                grid.Columns.Add("caseAge", "Case Age Tag");
                grid.Rows.Add("42 days");
                form.Controls.Add(unrelated);
                form.Controls.Add(desired);
                form.Controls.Add(anchor);
                form.Controls.Add(grid);
                Application.Run(form);
                return 0;
            }
            if (args.Contains("--embedded-dialog"))
            {
                var panel = new Panel { AccessibleName = "Others are also making changes",
                    Left = 80, Top = 80, Width = 420, Height = 180 };
                var button = new Button { Text = "See just mine", Name = "seeMine",
                    Left = 40, Top = 90, Width = 140 };
                button.Click += (_, _) => panel.Visible = false;
                panel.Controls.Add(button);
                form.Controls.Add(panel);
                Application.Run(form);
                return 0;
            }
            if (args.Contains("--resized"))
            {
                form.Size = new System.Drawing.Size(900, 600);
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new System.Drawing.Point(80, 120);
            }
            var strip = new StatusStrip { Name = "statusStrip" };
            var opener = new ToolStripButton("Work menu");
            using var menu = new ContextMenuStrip();
            menu.Items.Add("Browse items", null, (_, _) => form.Text = HostTitle + " - Invoked");
            if (args.Contains("--nested"))
            {
                var command = menu.Items[0];
                menu.Items.Remove(command);
                var parent = new ToolStripMenuItem("Utilities");
                parent.DropDownItems.Add(command);
                menu.Items.Add(parent);
            }
            if (args.Contains("--disabled")) menu.Items[0].Enabled = false;
            if (args.Contains("--duplicate"))
                menu.Items.Add("Browse items", null, (_, _) => form.Text = HostTitle + " - Wrong command");
            using var timer = new System.Windows.Forms.Timer { Interval = 700 };
            timer.Tick += (_, _) =>
            {
                timer.Stop();
                menu.Show(strip, new System.Drawing.Point(0, 0));
            };
            opener.Click += (_, _) => timer.Start();
            strip.Items.Add(opener);
            form.Controls.Add(strip);
            if (args.Contains("--grid-focus"))
            {
                form.Controls.Add(new DataGridView { Name = "summaryGrid", Dock = DockStyle.Fill });
                var refresh = new Button { Name = "commandButton", Text = "Run command", Dock = DockStyle.Top };
                refresh.Click += (_, _) => form.Text = HostTitle + " - Refreshed";
                form.Controls.Add(refresh);
                form.Controls.Add(new Label { Name = "statusLabel", Text = "Ready", Dock = DockStyle.Bottom });
            }
            if (args.Contains("--selection"))
            {
                var list = new ListView { Dock = DockStyle.Fill, View = View.Details,
                    MultiSelect = true, AccessibleName = "Selection regression list" };
                list.Columns.Add("Items", 250);
                for (var row = 0; row < 51; row++) list.Items.Add($"Current row {row + 1}");
                form.Controls.Add(list);
                form.Shown += (_, _) => { list.Items[0].Selected = true; list.Focus(); };
            }
            using var optionalTimer = new System.Windows.Forms.Timer { Interval = 9500 };
            using var readyTimer = new System.Windows.Forms.Timer { Interval = 3000 };
            readyTimer.Tick += (_, _) =>
            {
                readyTimer.Stop();
                var next = new Button { Name = "nextButton", Text = "Next action", Dock = DockStyle.Right };
                next.Click += (_, _) => form.Text = HostTitle + " - Ready action executed";
                form.Controls.Add(next);
            };
            if (args.Contains("--optional"))
            {
                optionalTimer.Tick += (_, _) =>
                {
                    optionalTimer.Stop();
                    using var dialog = new Form { Text = "RSR optional regression dialog", Width = 300, Height = 150 };
                    var yes = new Button { Text = "Yes", Dock = DockStyle.Fill };
                    yes.Click += (_, _) =>
                    {
                        if (args.Contains("--optional-ready")) readyTimer.Start();
                        else form.Text = HostTitle + " - Optional handled";
                        dialog.Close();
                    };
                    dialog.Controls.Add(yes);
                    dialog.ShowDialog(form);
                };
                form.Shown += (_, _) => optionalTimer.Start();
            }
            Application.Run(form);
            return 0;
        }
        SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
        var task = args is ["--excel-capture", var outputDirectory]
            ? ExcelCaptureCheck.RunAsync(outputDirectory)
            : args is ["--excel-row-menu-capture"]
                ? ExcelDialogCaptureCheck.RunRowMenuCaptureAsync()
            : args is ["--excel-copy-fast-capture"]
                ? ExcelDialogCaptureCheck.RunCopyCaptureAsync(fastRelease: true)
            : args is ["--excel-copy-capture"]
                ? ExcelDialogCaptureCheck.RunCopyCaptureAsync()
            : args is ["--excel-dialog-capture"]
                ? ExcelDialogCaptureCheck.RunAsync()
            : args is ["--excel-replacements"]
                ? ExcelDialogCaptureCheck.RunReplacementsAsync()
            : args is ["--excel-filters"]
                ? ExcelDialogCaptureCheck.RunFiltersAsync()
            : args is ["--excel-cell-navigation"]
                ? ExcelDialogCaptureCheck.RunCellNavigationAsync()
            : args is ["--excel-column-selection"]
                ? ExcelDialogCaptureCheck.RunColumnSelectionAsync()
            : args is ["--excel-filtered-fill"]
                ? ExcelDialogCaptureCheck.RunFilteredFillAsync()
            : args is ["--excel-native-visible-fill"]
                ? ExcelDialogCaptureCheck.RunNativeVisibleFillAsync()
            : args is ["--excel-row-copy"]
                ? ExcelDialogCaptureCheck.RunFilteredFillAsync(rowCopyOnly: true)
            : args is ["--excel-formula-extension"]
                ? ExcelDialogCaptureCheck.RunFormulaExtensionAsync()
            : args is ["--excel-chart-capture"]
                ? ExcelDialogCaptureCheck.RunChartCaptureAsync()
            : args is ["--excel-chart-inspect", var processId, var x, var y]
                ? ExcelDialogCaptureCheck.InspectChartAsync(int.Parse(processId), int.Parse(x), int.Parse(y))
            : args is ["--planner-tests"]
                ? FoundryChecks.RunAsync()
            : args is ["--planner-fill-tests"]
                ? FoundryChecks.RunFilteredFillRecoveryAsync()
            : args is ["--filter-confirmation-tests"]
                ? FoundryChecks.RunFilterConfirmationChecksAsync()
            : args is ["--conditional-grid-focus-tests"]
                ? CheckConditionalGridFocusAsync()
            : args is ["--duplicate-fill-tests"]
                ? FoundryChecks.RunDuplicateFillChecksAsync()
            : args is ["--intent-input-tests"]
                ? CheckIntentInputAsync()
            : args is ["--menu-window-tests"]
                ? CheckMenuWindowFallbackAsync()
            : args is ["--list-selection"]
                ? CheckAllRowSelectionAsync()
            : args is ["--readiness-tests"]
                ? CheckStableReadinessAsync()
            : args is ["--window-activation-tests"]
                ? CheckWindowActivationRuleAsync()
            : args is ["--refresh-tests"]
                ? CheckRefreshCompletionAsync()
            : args is ["--refresh-input-tests"]
                ? CheckRefreshInputAsync()
            : args is ["--refresh-choice-tests"]
                ? CheckRefreshChoiceAsync()
            : args is ["--embedded-dialog-tests"]
                ? CheckEmbeddedDialogAsync()
            : args is ["--recorder-input"]
                ? CheckRecorderInputAsync()
            : args is ["--recording-validation", var planPath]
                ? CheckRecordingValidationAsync(planPath)
            : args is ["--compact-recording", var recordingDirectory]
                ? CompactRecordingAsync(recordingDirectory)
            : args is ["--foundry-check", var configuration]
                ? CheckFoundryAsync(configuration)
            : args is ["--generate-recording", var settings, var capturedDirectory]
                ? GenerateRecordingAsync(settings, capturedDirectory)
            : args is ["--regenerate-recording", var rebuildSettings, var rebuildDirectory]
                ? GenerateRecordingAsync(rebuildSettings, rebuildDirectory, overwrite: true)
            : args.Contains("--missing-only") ? CheckAsync("missing") : RunAsync();
        while (!task.IsCompleted)
        {
            Application.DoEvents();
            Thread.Sleep(10);
        }
        try { task.GetAwaiter().GetResult(); return 0; }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static async Task CheckMenuWindowFallbackAsync()
    {
        var owner = new ControlRef("SampleDesktopApp", "Sample application main", "", "Case Work",
            "ControlType.Button", "", "statusStrip1");
        var menu = new ControlRef("SampleDesktopApp", "Sample Case Query", "", "Tools",
            "ControlType.MenuItem", "", null);
        var child = menu with { Name = "Legacy command", ParentName = "Tools" };
        var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!
            .GetMethod("Compact")!;
        var normalized = (ExecutionPlan)compact.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now,
            "Transient menu owner", [new(1, "click", owner, null, null, null, null, null),
                new(2, "click", menu, null, null, null, null, null),
                new(3, "click", child, null, null, null, null, null)])])!;
        if (normalized.Steps[1].Target?.Window != owner.Window ||
            normalized.Steps[2].Target?.Window != owner.Window)
            throw new InvalidOperationException("Transient menu ownership was not recovered from its opener.");
        await CheckAsync("legacy");
        await CheckAsync("wrong-window");
        await CheckAsync("wrong-process");
        await CheckAsync("duplicate");
    }

    private static Task CheckIntentInputAsync()
    {
        var assembly = typeof(ControlRef).Assembly;
        var normalizeFill = assembly.GetType("DesktopSteps.Foundry")!
            .GetMethod("TryNormalizeExplicitFilteredFill", BindingFlags.NonPublic | BindingFlags.Static)!;
        var staleCell = new ControlRef("EXCEL", "Book.xlsx - Excel", "I7", "I7",
            "ControlType.DataItem", "XLSpreadsheetCell", "Grid");
        object?[] fillArguments =
        [
            "use the first visible filtered row rather than a fixed row number. intent is to copy the cell value to below cells in column L until before the empty adjacent cell in K column",
            staleCell,
            null,
            null
        ];
        if (!(bool)normalizeFill.Invoke(null, fillArguments)! ||
            fillArguments[2] is not ControlRef { Name: "L7", AutomationId: "L7" } ||
            fillArguments[3] as string != "adjacent-column:K")
            throw new InvalidOperationException("Explicit filtered fill did not override a stale scrolled cell identity.");
        object?[] vagueFillArguments = ["copy cells down", staleCell, null, null];
        if ((bool)normalizeFill.Invoke(null, vagueFillArguments)!)
            throw new InvalidOperationException("A vague fill instruction was accepted without explicit columns.");
        var method = assembly.GetType("DesktopSteps.MainForm")!
            .GetMethod("IsIntentSubmitKey", BindingFlags.NonPublic | BindingFlags.Static)!;
        bool Submit(Keys keys) => (bool)method.Invoke(null, [keys])!;
        if (!Submit(Keys.Enter) || !Submit(Keys.Return) ||
            Submit(Keys.Shift | Keys.Enter) || Submit(Keys.Control | Keys.Enter))
            throw new InvalidOperationException("Intent editor Enter and Shift+Enter behavior regressed.");
        var compact = assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
        var header = new ControlRef("EXCEL", "Book.xlsx - Excel", "", "A", "ControlType.DataItem",
            "XLGridColumnHeader", "Grid");
        var cell = new ControlRef("EXCEL", "Book.xlsx - Excel", "F11", "F11", "ControlType.DataItem",
            "XLSpreadsheetCell", "Grid");
        var orphan = new PlanStep(1, "manual-click", header, null, null, "unverified-click", null, null);
        var fill = new PlanStep(2, "fill-down-to-adjacent-data-end", cell, null, null,
            "adjacent-column:E", null, null, "copy the function in column F until column E is empty");
        ExecutionPlan Compact(params PlanStep[] steps) => (ExecutionPlan)compact.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Intent input regression", steps.ToList())])!;
        var address = new ControlRef("browser", "New tab", "address", "Address",
            "ControlType.Edit", "AddressEditor", "Toolbar");
        var loadedSearch = new ControlRef("browser", "Loaded page", "search", "Search",
            "ControlType.ComboBox", "PageSearch", "Page");
        var submittedNavigation = Compact(
            new PlanStep(1, "type", address, "example.test", null, null, null, null),
            new PlanStep(2, "key", loadedSearch, null, "Enter", null, null, null));
        if (submittedNavigation.Steps[1].Target != address)
            throw new InvalidOperationException("A submit key retained a post-navigation target instead of its typed editor.");
        Console.WriteLine("PASS: Enter immediately after text keeps the pre-navigation editor identity.");
        var browserMenu = new ControlRef("browser", "Search results", "menu", "Browser menu",
            "ControlType.Button", "MenuButton", "Toolbar");
        var privateDocument = new ControlRef("browser", "New tab - [InPrivate] - Browser", "document",
            "New tab", "ControlType.Document", "", "");
        var privateAddress = address with { Window = privateDocument.Window };
        var privateTransition = Compact(
            new PlanStep(1, "click", browserMenu, null, null, null, null, null),
            new PlanStep(2, "click", privateDocument, null, null, null, null, null),
            new PlanStep(3, "click", privateAddress, null, null, null, null, null));
        if (privateTransition.Steps.Count != 2 ||
            privateTransition.Steps[0].Action != "open-private-window" ||
            privateTransition.Steps[0].ExpectedState != "opened-window:" + privateDocument.Window ||
            privateTransition.Steps[1].Target != privateAddress)
            throw new InvalidOperationException("A verified private-window transition remained a document click.");
        Console.WriteLine("PASS: a resulting private document becomes a verified private-window command.");
        var scrollRepeatCount = assembly.GetType("DesktopSteps.Executor")!.GetMethod(
            "ScrollRepeatCount", BindingFlags.NonPublic | BindingFlags.Static)!;
        if ((int)scrollRepeatCount.Invoke(null, [-840])! != 7 ||
            (int)scrollRepeatCount.Invoke(null, [120])! != 1 ||
            (int)scrollRepeatCount.Invoke(null, [0])! != 1)
            throw new InvalidOperationException("Recorded wheel distance was not preserved by page scrolling.");
        Console.WriteLine("PASS: UIA page scrolling preserves the recorded wheel distance.");
        var loadedGroup = new ControlRef("browser", "Loaded page", "", "Loading content",
            "ControlType.Group", "PageGroup", "Page");
        var pageSearch = new ControlRef("browser", "Loaded page", "search", "Search",
            "ControlType.ComboBox", "PageSearch", "Page");
        var finalSearch = Compact(
            new PlanStep(1, "type", address, "example.test", null, null, null, null),
            new PlanStep(2, "key", address, null, "Enter", null, null, null),
            new PlanStep(3, "type", loadedGroup, "partial", null, null, null, null),
            new PlanStep(4, "click", pageSearch, null, null, null, null, null),
            new PlanStep(5, "type", pageSearch, "complete query", null, null, null, null),
            new PlanStep(6, "key", pageSearch, null, "Enter", null, null, null));
        if (finalSearch.Steps.Any(step => step.Value == "partial") ||
            finalSearch.Steps.Count(step => step.Action == "type") != 2)
            throw new InvalidOperationException("Transient post-navigation draft text was replayed before the final query.");
        Console.WriteLine("PASS: delayed non-edit draft text cannot concatenate with a verified final query.");
        if (Compact(orphan, fill).Steps.Any(step => step.Action == "manual-click") ||
            !Compact(orphan, fill with { Intent = null }).Steps.Any(step => step.Action == "manual-click"))
            throw new InvalidOperationException("Orphan whole-column selection cleanup regressed.");
        var filterCompactor = assembly.GetType("DesktopSteps.ExcelFilterPlan")!
            .GetMethod("Compact", BindingFlags.NonPublic | BindingFlags.Static)!;
        var filterTarget = new ControlRef("EXCEL", "Book.xlsx - Excel", "Dropdown", "Filter applied",
            "ControlType.MenuItem", "", "G1");
        var filterItem = new ControlRef("EXCEL", "Book.xlsx - Excel", "", "(Select All)",
            "ControlType.TreeItem", "", "Manual Filter");
        var openFilter = new PlanStep(1, "click", filterTarget, null, null, null, "open.jpg", null);
        var uncheckedAll = new PlanStep(2, "click", filterItem, null, null, "filter:unchecked", "unchecked.jpg", null);
        var delayedMixed = new PlanStep(3, "click", filterItem, null, null, "filter:mixed", "mixed.jpg", null);
        var wesam = new PlanStep(4, "click", filterItem with { Name = "Wesam (walziadat)" },
            null, null, null, "wesam.jpg", null);
        var applyFilter = new PlanStep(5, "click", filterItem with
            { Name = "OK", ControlType = "ControlType.Button", ParentName = "" }, null, null, null, "ok.jpg", null);
        var reconstructedFilters = (List<PlanStep>)filterCompactor.Invoke(null,
            [new List<PlanStep> { openFilter, uncheckedAll, delayedMixed, wesam, applyFilter,
                openFilter with { Number = 6 }, delayedMixed with { Number = 7 },
                wesam with { Number = 8 }, applyFilter with { Number = 9 },
                openFilter with { Number = 10 }, uncheckedAll with
                    { Number = 11, ExpectedState = null }, applyFilter with { Number = 12 } }])!;
        if (reconstructedFilters.Count != 3 ||
            reconstructedFilters[0] is not { ExpectedState: "filter:only", Value: "[\"Wesam (walziadat)\"]" } ||
            reconstructedFilters[1] is not { ExpectedState: "filter:exclude", Value: "[\"Wesam (walziadat)\"]" } ||
            reconstructedFilters[2].ExpectedState != "filter:all")
            throw new InvalidOperationException("Delayed mixed Select All evidence reversed the only/exclude/all filter sequence.");
        var filterOpener = new PlanStep(1, "click",
            new ControlRef("EXCEL", "Find and Replace", "Dropdown", "No filter applied",
                "ControlType.MenuItem", "", "G1"), null, null, null, null, null);
        var semanticFilter = new PlanStep(2, "filter-values",
            filterOpener.Target! with { Window = "Book.xlsx - Excel" }, "[\"Wesam\"]", null,
            "filter:only", null, null);
        var filtered = Compact(filterOpener, semanticFilter);
        if (filtered.Steps.Count != 1 || filtered.Steps[0].Action != "filter-values")
            throw new InvalidOperationException("A stale filter opener survived before its semantic filter.");
        var falseTab = new PlanStep(3, "click",
            new ControlRef("EXCEL", "Book.xlsx - Excel", "SheetTab", "Chart",
                "ControlType.TabItem", "", "Book"), null, null, null, "screen-tab.jpg",
            "Open the recorded worksheet 'Chart' before continuing with its grid.");
        var columnHeader = new PlanStep(4, "click", header, null, null, "unverified-click",
            "screen-header.jpg", null);
        var optionalPrompt = new PlanStep(5, "optional-click",
            new ControlRef("EXCEL", "", "", "See just mine", "ControlType.Button",
                "NetUIButton", "Others are also making changes"), null, null, "window-closed", null, null);
        var filteredEntry = new PlanStep(6, "type", cell, "FL", null, null, null, null,
            null, null, null, "first-visible-filtered-row");
        var occasionalPrompt = optionalPrompt with { Action = "click",
            Intent = "the see everyone or just see mine dialog doesn't always appear" };
        var repairedFilter = Compact(semanticFilter, occasionalPrompt, falseTab, columnHeader, filteredEntry);
        if (repairedFilter.Steps.Any(step => step.Target?.ControlType == "ControlType.TabItem"))
            throw new InvalidOperationException("A false recovered sheet tab survived a filtered-row continuation: " +
                string.Join(" | ", repairedFilter.Steps.Select(step =>
                    $"{step.Action}:{step.Target?.Name}:{step.Target?.Window}:{step.TargetStrategy}:{step.Explanation}")));
        var orderedDialog = Compact(semanticFilter, columnHeader, occasionalPrompt, filteredEntry);
        if (orderedDialog.Steps.Count < 3 || orderedDialog.Steps[1].Target?.Name != "See just mine" ||
            orderedDialog.Steps[2].Target?.Name != "A")
            throw new InvalidOperationException("An occasional Excel dialog was not moved ahead of the worksheet click it can block.");
        var recover = assembly.GetType("DesktopSteps.MainForm")!.GetMethod(
            "RecoverOmittedWorksheetNavigation", BindingFlags.NonPublic | BindingFlags.Static)!;
        var navigationPlan = new ExecutionPlan(1, DateTimeOffset.Now, "False tab recovery",
            [semanticFilter, occasionalPrompt, columnHeader, filteredEntry]);
        var tabEvent = new RecordedEvent(DateTimeOffset.Now, "click", falseTab.Target, null, null,
            "screen-tab.jpg", "enabled=True", null, null, 500, 900, null);
        var headerEvent = new RecordedEvent(DateTimeOffset.Now, "click", header, null, null,
            "screen-header.jpg", "enabled=True", "unverified-click", null, 200, 300, null);
        object?[] recoverArgs = [navigationPlan, new List<RecordedEvent> { tabEvent, headerEvent }, null];
        if ((bool)recover.Invoke(null, recoverArgs)! ||
            navigationPlan.Steps.Any(step => step.Target?.ControlType == "ControlType.TabItem"))
            throw new InvalidOperationException("Worksheet recovery reintroduced a blocked, stale tab identity.");
        var sapTab = falseTab with { Target = falseTab.Target! with { Name = "SAP rough draft" },
            Screenshot = "screen-sap.jpg" };
        var a1 = columnHeader with { Target = cell with { AutomationId = "A1", Name = "A1" },
            Screenshot = null };
        var delete = new PlanStep(5, "key", a1.Target, null, "Delete", "field-value:",
            "screen-delete.jpg", null);
        var missingTabPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Missing worksheet tab", [a1, delete]);
        var sapTabEvent = new RecordedEvent(DateTimeOffset.Now, "click", sapTab.Target, null, null,
            sapTab.Screenshot, "enabled=True", null, null, 100, 900, null);
        var a1Event = new RecordedEvent(DateTimeOffset.Now, "click", a1.Target, null, null,
            null, "enabled=True", null, null, 100, 100, null);
        var deleteEvent = new RecordedEvent(DateTimeOffset.Now, "key", delete.Target, null, "Delete",
            delete.Screenshot, null, "field-value:");
        object?[] missingTabArgs = [missingTabPlan,
            new List<RecordedEvent> { sapTabEvent, a1Event, deleteEvent }, null];
        if (!(bool)recover.Invoke(null, missingTabArgs)! ||
            missingTabPlan.Steps.FirstOrDefault()?.Target?.Name != "SAP rough draft")
            throw new InvalidOperationException("A screenshot-free cell action did not recover its verified preceding worksheet tab.");
        var represented = assembly.GetType("DesktopSteps.MainForm")!.GetMethod("RecordedIntentIsRepresented",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        var chartSource = new PlanStep(1, "set-chart-source-range",
            cell with { Name = "Chart 7", ClassName = "ExcelChartObject", ControlType = "ControlType.Image" },
            "{\"Sheet\":\"Sheet1\",\"Range\":\"S4:T14\"}", null, "chart-source-verified", null, null);
        var chartPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Chart source", [chartSource]);
        bool Represented(string note) => (bool)represented.Invoke(null, [chartPlan, note])!;
        if (!Represented("intent is to add data to be the cells that we pasted in previous step") ||
            !Represented("previous") ||
            Represented("previous row") ||
            Represented("add data to the report") ||
            Represented("use cells from another worksheet"))
            throw new InvalidOperationException("Verified chart-source intent recognition regressed.");
        var rowPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Live row", [
            new PlanStep(1, "duplicate-range-values-and-formulas", cell, "A11:F11", null,
                "destination-range:A12:F12", null, null),
            new PlanStep(2, "update-cell-to-relative-weekday", cell, "A11", null,
                "destination-range:A12:F12", null, null, RelativeWeekday: "Thursday")]);
        if (!(bool)represented.Invoke(null, [rowPlan,
                "tent is to copy functions of the last row containing values from columns A to F, then use next Thursday"])! ||
            !(bool)represented.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now, "Optional", []),
                "the see everyone or just see mine dialog doesn't always appear"])!)
            throw new InvalidOperationException("Relative last-row or informational optional-dialog intent recognition regressed.");
        var chartTab = falseTab with { Target = falseTab.Target! with { Name = "Evolution" } };
        var chartScroll = new PlanStep(2, "scroll",
            new ControlRef("EXCEL", "Book.xlsx - Excel", "", "", "ControlType.ScrollBar",
                "NetUIScrollBar", "", "Vertical"), "-120", "Vertical", "range:50", null, null);
        var legend = new PlanStep(3, "set-chart-legend-layout", chartSource.Target,
            "{\"Left\":0.5,\"Top\":0.1,\"Width\":0.4,\"Height\":0.7}", null,
            "chart-legend-sheet:Sheet1", null, null);
        var realEvolution = chartTab with { Number = 4, Screenshot = "screen-real-evolution.jpg",
            Explanation = "Open Evolution before its row update." };
        var duplicateRow = new PlanStep(5, "duplicate-range-values-and-formulas",
            cell with { AutomationId = "A11", Name = "A11" }, "A11:F11", null,
            "destination-range:A12:F12", null, null, "copy the previous row A to F");
        var chartCell = columnHeader with { Target = cell with { AutomationId = "C19", Name = "C19" },
            ExpectedState = null, Screenshot = null };
        var chartSequence = Compact(chartSource, chartTab, chartCell, chartScroll, legend,
            realEvolution, duplicateRow);
        var legendPosition = chartSequence.Steps.FindIndex(step => step.Action == "set-chart-legend-layout");
        var evolutionPosition = chartSequence.Steps.FindIndex(step => step.Target?.Name == "Evolution");
        if (legendPosition != 1 || evolutionPosition != 2 ||
            chartSequence.Steps.Count(step => step.Target?.Name == "Evolution") != 1 ||
            chartSequence.Steps.Last().Action != "duplicate-range-values-and-formulas" ||
            chartSequence.Steps.Last().TargetStrategy != "last-populated-row-in-first-column")
            throw new InvalidOperationException("Transient chart-edit navigation was not removed while preserving the later row-update sheet transition.");
        var pageRight = new PlanStep(2, "click", chartScroll.Target! with
            { Name = "Page right", ControlType = "ControlType.Button", ClassName = "NetUIRepeatButton" },
            null, null, null, null, null);
        var chartPointer = new PlanStep(3, "click", chartSource.Target, null, null, null, null, null);
        var chartContext = chartPointer with { Number = 4, Action = "context-click" };
        var chartPrevious = chartPointer with { Number = 5, Intent = "previous" };
        var preSource = Compact(chartScroll, pageRight, chartPointer, chartContext, chartPrevious,
            chartSource with { Number = 6 });
        if (preSource.Steps.Count != 1 || preSource.Steps[0].Action != "set-chart-source-range")
            throw new InvalidOperationException("Pointer navigation before verified native chart source was retained.");
        var betweenChartOperations = Compact(chartSource, chartScroll, chartScroll, legend);
        if (betweenChartOperations.Steps.Count != 2 ||
            betweenChartOperations.Steps[0].Action != "set-chart-source-range" ||
            betweenChartOperations.Steps[1].Action != "set-chart-legend-layout")
            throw new InvalidOperationException("Scroll-only navigation between verified chart operations was retained.");
        if (!Compact(chartTab, chartPointer, chartSource).Steps.Any(step =>
                step.Target?.ControlType == "ControlType.TabItem") ||
            !Compact(chartScroll with { Intent = "Review this position" }, chartSource).Steps.Any(step =>
                step.Action == "scroll") ||
            !Compact(chartPointer with { Target = chartPointer.Target! with { Name = "Other Chart" } }, chartSource)
                .Steps.Any(step => step.Target?.Name == "Other Chart"))
            throw new InvalidOperationException("Chart cleanup removed a worksheet transition, explicit navigation intent or a different chart.");
        var formulaGesture = Compact(
            new PlanStep(1, "click", cell with { Name = "Values", AutomationId = "",
                ControlType = "ControlType.Button", ParentName = "Paste Values" }, null, null,
                null, null, null, "function"),
            new PlanStep(2, "click", cell with { Name = "F11", AutomationId = "F11" }, null, null, null, null, null),
            new PlanStep(3, "click", cell with { Name = "F11", AutomationId = "F11" }, null, null, null, null, null),
            new PlanStep(4, "click", cell with { Name = "F12", AutomationId = "F12" }, null, null, null, null, null),
            new PlanStep(5, "scroll", chartScroll.Target, "120", "Vertical", "range:0", null, null),
            new PlanStep(6, "click", cell with { Name = "E", AutomationId = "",
                ClassName = "XLGridColumnHeader" }, null, null, null, null, null),
            new PlanStep(7, "click", cell with { Name = "Data", AutomationId = "TabData",
                ControlType = "ControlType.TabItem" }, null, null, null, null, null),
            new PlanStep(8, "click", cell with { Name = "Remove Duplicates", AutomationId = "",
                ControlType = "ControlType.Button" }, null, null, null, null, null));
        var fillPosition = formulaGesture.Steps.FindIndex(step => step.Action == "extend-formula-to-adjacent-data-end");
        var removePosition = formulaGesture.Steps.FindIndex(step => step.Target?.Name == "Remove Duplicates");
        if (fillPosition < 0 || removePosition <= fillPosition ||
            formulaGesture.Steps[fillPosition].Value != "F" ||
            formulaGesture.Steps[fillPosition].ExpectedState != "adjacent-column:E" ||
            formulaGesture.Steps.Any(step => step.Target?.Name is "F11" or "F12" && step.Action == "click"))
            throw new InvalidOperationException("Formula-fill gesture did not become a verified dependency before Remove Duplicates.");
        var workbookFocus = new PlanStep(1, "click", new ControlRef("EXCEL", "Book.xlsx - Excel", "",
            "Book.xlsx", "ControlType.Custom", "", ""), null, null,
            "worksheet-tab-focus-only", null, "Focus transition.");
        var verifiedSheet = new PlanStep(2, "click", new ControlRef("EXCEL", "Book.xlsx - Excel", "SheetTab",
            "Clean Sheet", "ControlType.TabItem", "", "Book.xlsx"), null, null, null, null, "Open sheet.");
        var focusSequence = Compact(workbookFocus, verifiedSheet);
        if (focusSequence.Steps.Count != 1 || focusSequence.Steps[0].Target?.Name != "Clean Sheet")
            throw new InvalidOperationException("Excel workbook focus before a verified sheet tab remained replayable.");
        Console.WriteLine("PASS: Excel workbook focus before a verified sheet tab is discarded without losing the tab.");
        Console.WriteLine("PASS: repeated formula-source gesture is completed before the following Remove Duplicates command.");
        Console.WriteLine("PASS: Enter submits an intent while Shift+Enter remains available for a new line.");
        Console.WriteLine("PASS: an orphan whole-column selection is removed only before a complete intent-backed fill.");
        Console.WriteLine("PASS: verified chart source represents the matching pasted-cells intent without accepting broad notes.");
        Console.WriteLine("PASS: transient navigation inside verified chart editing is removed while the later row-update transition is preserved.");
        Console.WriteLine("PASS: a stale filter opener is removed before the verified semantic filter for that column.");
        Console.WriteLine("PASS: a filtered-row continuation rejects a stale sheet-tab identity from a blocked click.");
        Console.WriteLine("PASS: an occasional Excel collaboration dialog is handled before its modal window can block the next worksheet click.");
        Console.WriteLine("PASS: a screenshot-free Excel cell action recovers its worksheet only when its following action also matches.");
        return Task.CompletedTask;
    }

    private static Task CheckConditionalGridFocusAsync()
    {
        var method = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsGridFocusBeforeCommand", BindingFlags.NonPublic | BindingFlags.Static)!;
        const string intent = "If the confirmation dialog is present, click Yes; otherwise click Refresh.";
        var grid = new PlanStep(11, "click",
            new ControlRef("SampleDesktopApp", "Sample Case Query - Case Triage (any)", "dgvSummary", "",
                "ControlType.DataGrid", "WindowsForms10.Window", ""),
            null, null, null, null, "Focus the grid.", intent);
        var fallback = new PlanStep(13, "click-if-previous-absent",
            new ControlRef("SampleDesktopApp", "Sample Case Query - Case Triage (any)", "", "Refresh View",
                "ControlType.Button", "", "toolStrip1"),
            null, null, null, null, "Refresh when the dialog is absent.", intent);
        if (!(bool)method.Invoke(null, [grid, fallback])! ||
            (bool)method.Invoke(null, [grid with { Intent = "Different intent" }, fallback])!)
            throw new InvalidOperationException("Conditional dialog grid-focus suppression regressed.");
        Console.WriteLine("PASS: unnamed grid focus before a matching optional-dialog fallback is skipped safely.");
        return Task.CompletedTask;
    }

    private static Task CheckWindowActivationRuleAsync()
    {
        var rule = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsWindowActivationClick", BindingFlags.NonPublic | BindingFlags.Static)!;
        var title = new PlanStep(10, "click",
            new ControlRef("SampleDesktopApp", "Sample Case Query", "TitleBar",
                "Sample Case Query", "ControlType.TitleBar", "", "Sample Case Query"),
            null, null, null, null,
            "Focus the Sample Case Query window.",
            "If the optional question appears click Yes; otherwise continue.");
        if (!(bool)rule.Invoke(null, [title])! ||
            (bool)rule.Invoke(null, [title with { ExpectedState = "window-moved" }])! ||
            (bool)rule.Invoke(null, [title with { Target = title.Target! with
                { Name = "Different", ParentName = "Different" } }])! ||
            (bool)rule.Invoke(null, [title with { Target = title.Target! with
                { Window = "Find and Replace", Name = "Find and Replace",
                    ParentName = "Find and Replace" } }])!)
            throw new InvalidOperationException("Window activation rule accepted the wrong title-bar action.");
        Console.WriteLine("PASS: a matching title-bar focus with an attached workflow note is replayed as verified window activation; meaningful and special dialog title actions remain explicit.");
        return Task.CompletedTask;
    }

    private static async Task CheckRefreshChoiceAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --refresh-choice")
        {
            UseShellExecute = false
        }) ?? throw new InvalidOperationException("Could not start refresh choice fixture.");
        try
        {
            for (var attempt = 0; host.MainWindowHandle == 0 && attempt < 100; attempt++)
                await Task.Delay(50);
            if (host.MainWindowHandle == 0) throw new InvalidOperationException("Refresh choice fixture did not open.");
            var process = host.ProcessName;
            var target = new ControlRef(process, HostTitle, null, "Refresh", "ControlType.Button", null, null);
            var context = new ControlRef(process, HostTitle, "viewAnchor", "Case Age Tag", "ControlType.Text", null, null);
            var automation = typeof(ControlRef).Assembly.GetType("DesktopSteps.Automation")!;
            var resolve = automation.GetMethod("ResolveUniqueRefreshCommand", BindingFlags.NonPublic | BindingFlags.Static)!;
            var chosen = (AutomationElement?)resolve.Invoke(null, [target, context]);
            if (chosen?.Current.AutomationId != "viewRefresh")
                throw new InvalidOperationException("Refresh fallback did not choose the command in the recorded header's view region.");
            try
            {
                _ = resolve.Invoke(null, [target, null]);
                throw new InvalidOperationException("Ambiguous Refresh commands were accepted without recorded context.");
            }
            catch (TargetInvocationException ex) when (ex.InnerException is InvalidOperationException) { }
            Console.WriteLine("PASS: duplicate Refresh commands resolve only through the next recorded view context.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(5000)) host.Kill(true);
            }
        }
    }

    private static async Task CheckEmbeddedDialogAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --embedded-dialog")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Could not start embedded dialog fixture.");
        try
        {
            for (var attempt = 0; host.MainWindowHandle == 0 && attempt < 100; attempt++) await Task.Delay(50);
            var target = new ControlRef(host.ProcessName, "", "", "See just mine", "ControlType.Button", "",
                "Others are also making changes");
            var automation = typeof(ControlRef).Assembly.GetType("DesktopSteps.Automation")!;
            var resolve = automation.GetMethod("ResolveUniqueEmbeddedButton", BindingFlags.NonPublic | BindingFlags.Static)!;
            var button = (AutomationElement?)resolve.Invoke(null, [target]);
            if (button?.Current.AutomationId != "seeMine")
                throw new InvalidOperationException("Embedded dialog button was not resolved by its recorded parent identity.");
            ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
            await Task.Delay(200);
            if (resolve.Invoke(null, [target]) is not null)
                throw new InvalidOperationException("Handled embedded dialog remained available.");
            Console.WriteLine("PASS: embedded optional dialog resolves uniquely and disappears after invocation.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(5000)) host.Kill(true);
            }
        }
    }

    private static async Task GenerateRecordingAsync(string configuration, string directory, bool overwrite = false)
    {
        var path = Path.Combine(directory, "execution_plan.json");
        if (!overwrite && File.Exists(path)) throw new InvalidOperationException("A plan already exists; use the application's rebuild command.");
        var rawPath = Path.Combine(directory, "recorded_events.json");
        var raw = await File.ReadAllBytesAsync(rawPath);
        var events = System.Text.Json.JsonSerializer.Deserialize<List<RecordedEvent>>(raw)
            ?? throw new InvalidDataException("Recording is empty.");
        var assembly = typeof(ControlRef).Assembly;
        var foundryType = assembly.GetType("DesktopSteps.Foundry")!;
        var foundry = Activator.CreateInstance(foundryType, [configuration])!;
        using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(5));
        var plan = await (Task<ExecutionPlan>)foundryType.GetMethod("PlanAsync")!
            .Invoke(foundry, [events, timeout.Token, null])!;
        var temporaryPdf = Path.Combine(directory, "manual_steps.generating.pdf");
        var temporaryPlan = Path.Combine(directory, "execution_plan.generating.json");
        try
        {
            await (Task)assembly.GetType("DesktopSteps.ManualPdf")!.GetMethod("CreateAsync")!
                .Invoke(null, [plan, directory, temporaryPdf, timeout.Token])!;
            await File.WriteAllTextAsync(temporaryPlan, System.Text.Json.JsonSerializer.Serialize(plan,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true }), timeout.Token);
            var currentRaw = await File.ReadAllBytesAsync(rawPath, timeout.Token);
            if (!raw.SequenceEqual(currentRaw))
                throw new InvalidOperationException("Recording changed during generation; no plan was installed.");
            File.Move(temporaryPlan, path, overwrite);
            File.Move(temporaryPdf, Path.Combine(directory, "manual_steps.pdf"), overwrite);
            await File.WriteAllTextAsync(Path.Combine(directory, "summary.txt"), plan.Summary, timeout.Token);
        }
        finally
        {
            File.Delete(temporaryPlan);
            File.Delete(temporaryPdf);
        }
        Console.WriteLine($"Generated {plan.Steps.Count} plan/PDF steps through Foundry; raw recording unchanged, no desktop input.");
    }

    private static async Task CompactRecordingAsync(string directory)
    {
        var path = Path.Combine(directory, "execution_plan.json");
        var options = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
        var saved = System.Text.Json.JsonSerializer.Deserialize<ExecutionPlan>(await File.ReadAllTextAsync(path), options)
            ?? throw new InvalidDataException("Recording plan is empty.");
        var assembly = typeof(ControlRef).Assembly;
        var compacted = (ExecutionPlan)assembly.GetType("DesktopSteps.PlanCompactor")!
            .GetMethod("Compact")!.Invoke(null, [saved])!;
        var rawPath = Path.Combine(directory, "recorded_events.json");
        if (File.Exists(rawPath))
        {
            var events = System.Text.Json.JsonSerializer.Deserialize<List<RecordedEvent>>(
                await File.ReadAllTextAsync(rawPath), options) ?? [];
            compacted = (ExecutionPlan)assembly.GetType("DesktopSteps.PlanCompactor")!
                .GetMethod("RecoverExplicitDuplicatedRowIntent", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [compacted, events.Select(item => item.Intent).Where(note => note is not null)
                    .Select(note => note!)])!;
            compacted = (ExecutionPlan)assembly.GetType("DesktopSteps.ExcelFilterPlan")!
                .GetMethod("RecoverRecordedOutcomes", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [compacted, events])!;
            object?[] navigationArgs = [compacted, events, null];
            assembly.GetType("DesktopSteps.MainForm")!
                .GetMethod("RecoverOmittedWorksheetNavigation", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, navigationArgs);
        }
        var unresolved = compacted.Steps.Where(step => step.Action is "manual-click" or "unresolved-input").ToArray();
        if (unresolved.Length > 0)
            throw new InvalidDataException("Compacted plan still contains unresolved operations: " +
                string.Join(", ", unresolved.Select(step =>
                {
                    var index = compacted.Steps.IndexOf(step);
                    var next = index + 1 < compacted.Steps.Count ? compacted.Steps[index + 1] : null;
                    return $"step {step.Number} {step.Action} '{step.Target?.Name}' " +
                        $"class={step.Target?.ClassName} intent='{step.Intent}' origin='{step.OriginIntent}' " +
                        $"next={next?.Action} next-origin='{next?.OriginIntent}'";
                })) +
                "; no files were regenerated.");
        var temporaryPdf = Path.Combine(directory, "manual_steps.compacting.pdf");
        var temporaryPlan = Path.Combine(directory, "execution_plan.compacting.json");
        try
        {
            await (Task)assembly.GetType("DesktopSteps.ManualPdf")!.GetMethod("CreateAsync")!
                .Invoke(null, [compacted, directory, temporaryPdf, CancellationToken.None])!;
            await File.WriteAllTextAsync(temporaryPlan, System.Text.Json.JsonSerializer.Serialize(compacted, options));
            File.Move(temporaryPdf, Path.Combine(directory, "manual_steps.pdf"), true);
            File.Move(temporaryPlan, path, true);
            await File.WriteAllTextAsync(Path.Combine(directory, "summary.txt"), compacted.Summary);
        }
        finally
        {
            File.Delete(temporaryPdf);
            File.Delete(temporaryPlan);
        }
        Console.WriteLine($"Regenerated {compacted.Steps.Count} static plan/PDF steps in {directory}; no desktop input or Foundry request.");
    }

    private static async Task CheckRecorderInputAsync()
    {
        CheckIncompleteReplacementIntent();
        await CheckWheelDuringUiStallAsync();
        await CheckDialogCaptureAsync();
    }

    private static async Task CheckStableReadinessAsync()
    {
        var method = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("WaitForStableReadinessAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
        Task<bool> Wait(Func<Task<bool>> probe, TimeSpan timeout, CancellationToken token = default) =>
            (Task<bool>)method.Invoke(null, [probe, timeout, token])!;

        var watch = Stopwatch.StartNew();
        var probes = 0;
        if (!await Wait(() =>
            {
                probes++;
                return Task.FromResult(watch.Elapsed >= TimeSpan.FromSeconds(2.5));
            }, TimeSpan.FromSeconds(5)) || probes < 12)
            throw new InvalidOperationException("Readiness stopped at the old short lookup budget.");

        var sequence = new Queue<bool>([true, true, false, true, true, true]);
        if (!await Wait(() => Task.FromResult(sequence.Dequeue()), TimeSpan.FromSeconds(3)) ||
            sequence.Count != 0)
            throw new InvalidOperationException("Readiness did not reset after a transient unavailable view.");
        if (await Wait(() => Task.FromResult(false), TimeSpan.FromMilliseconds(400)))
            throw new InvalidOperationException("A missing view was reported as ready.");
        if (await Wait(async () =>
            {
                await Task.Delay(150);
                return true;
            }, TimeSpan.FromMilliseconds(100)))
            throw new InvalidOperationException("A late probe bypassed the readiness deadline.");
        using var cancellation = new CancellationTokenSource(100);
        try
        {
            await Wait(() => Task.FromResult(false), TimeSpan.FromSeconds(3), cancellation.Token);
            throw new InvalidOperationException("Readiness ignored cancellation.");
        }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { }
        Console.WriteLine("PASS: delayed view readiness exceeds short lookup, requires three stable probes, times out and cancels safely.");
    }

    private static async Task CheckRefreshCompletionAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --refresh-watch")
            { UseShellExecute = false }) ?? throw new InvalidOperationException("Refresh host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            AutomationElement? button = null;
            for (var attempt = 0; attempt < 100 && button is null; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0)
                    button = AutomationElement.FromHandle(host.MainWindowHandle).FindFirst(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.AutomationIdProperty, "refreshButton"));
                if (button is null) await Task.Delay(50, timeout.Token);
            }
            if (button is null) throw new InvalidOperationException("Refresh fixture button unavailable.");
            var target = new ControlRef(host.ProcessName, HostTitle, "refreshButton", "Refresh View",
                "ControlType.Button", null, null);
            var method = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
                .GetMethod("WatchRefreshAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
            var watch = Stopwatch.StartNew();
            var completion = (Task<string?>)method.Invoke(null, [button, target, timeout.Token, TimeSpan.FromSeconds(20)])!;
            ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
            if (await completion.WaitAsync(timeout.Token) != "state-restored" ||
                watch.Elapsed > TimeSpan.FromSeconds(10))
                throw new InvalidOperationException("Refresh completion did not recognize busy-to-ready without identical focus visuals.");
            watch.Restart();
            var unchanged = (Task<string?>)method.Invoke(null, [button, target, timeout.Token, TimeSpan.FromSeconds(20)])!;
            if (await unchanged.WaitAsync(timeout.Token) is not null ||
                watch.Elapsed < TimeSpan.FromSeconds(19) ||
                watch.Elapsed > TimeSpan.FromSeconds(23))
                throw new InvalidOperationException("Unchanged refresh state bypassed the 20-second completion limit.");
            Console.WriteLine("PASS: refresh busy-to-ready is confirmed with changed keyboard focus and without visual restoration.");
            Console.WriteLine("PASS: unchanged ready state is not treated as a completed refresh; timeout remains 20 seconds.");
            var configuredPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Configured refresh",
                [new(1, "click", target, null, null, null, null, null)], RefreshExpectedSeconds: 10);
            watch.Restart();
            var logs = new List<string>();
            var replay = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
                .GetMethod("ReplayAsync")!;
            await (Task<bool>)replay.Invoke(null, [configuredPlan, (Action<string>)logs.Add,
                (Func<string, bool>)(_ => false), (Func<string, bool>)(_ => false), timeout.Token])!;
            if (watch.Elapsed >= TimeSpan.FromSeconds(12) ||
                !logs.Any(line => line.Contains("Refresh completion timeout: 12 seconds")))
                throw new InvalidOperationException("Refresh replay did not continue on verified completion before its configured timeout.");
            Console.WriteLine("PASS: configured expectation is a timeout, not a minimum wait; verified completion continues early.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    private static async Task CheckRefreshInputAsync()
    {
        var timing = typeof(ControlRef).Assembly.GetType("DesktopSteps.RefreshTiming")!;
        var parser = timing.GetMethod("TryExpectedSeconds", BindingFlags.NonPublic | BindingFlags.Static)!;
        foreach (var text in new[] { "1", "18", "20", "001", (int.MaxValue - 2).ToString() })
        {
            object?[] input = [text, 0];
            if (!(bool)parser.Invoke(null, input)! || (int)input[1]! != int.Parse(text))
                throw new InvalidOperationException("Natural seconds were rejected: " + text);
        }
        foreach (var text in new[] { "", "0", "000", "-1", "+2", "1.5", "1e2", " 2", "2 ",
            "2\n", "abc", "2147483646", "9999999999999999999999", "\u0662" })
        {
            object?[] input = [text, 0];
            if ((bool)parser.Invoke(null, input)!)
                throw new InvalidOperationException("Invalid seconds were accepted: " + text);
        }
        var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Refresh",
            [new(1, "click-if-previous-absent", new ControlRef("Fixture", "View", null, "Refresh",
                "ControlType.Button", null, null), null, null, null, null, null,
                TargetStrategy: "unique-refresh-command")], RefreshExpectedSeconds: 7);
        var hasRefresh = timing.GetMethod("HasRefreshStep", BindingFlags.NonPublic | BindingFlags.Static)!;
        if (!(bool)hasRefresh.Invoke(null, [plan])! ||
            (bool)hasRefresh.Invoke(null, [plan with { Steps = [] }])!)
            throw new InvalidOperationException("Refresh prompt detection did not distinguish conditional and absent steps.");
        var loaded = System.Text.Json.JsonSerializer.Deserialize<ExecutionPlan>(
            System.Text.Json.JsonSerializer.Serialize(plan))!;
        if (loaded.RefreshExpectedSeconds != 7)
            throw new InvalidOperationException("Refresh expectation was not retained in JSON.");
        var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
        foreach (var invalid in new[] { 0, -1, int.MaxValue })
        {
            try
            {
                await (Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [plan with { RefreshExpectedSeconds = invalid }, (Action<string>)(_ => { }),
                        (Func<string, bool>)(_ => false), (Func<string, bool>)(_ => false), CancellationToken.None])!;
                throw new InvalidOperationException("Invalid saved refresh expectation ran.");
            }
            catch (InvalidDataException) { }
        }
        Console.WriteLine("PASS: refresh prompt accepts positive whole seconds only, detects conditional Refresh, persists expectation and rejects invalid saved values.");
    }

    private static async Task CheckRecordingValidationAsync(string path)
    {
        CheckIncompleteReplacementIntent();
        var plan = System.Text.Json.JsonSerializer.Deserialize<ExecutionPlan>(await File.ReadAllTextAsync(path),
            new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("The execution plan is empty.");
        var mainForm = typeof(ControlRef).Assembly.GetType("DesktopSteps.MainForm")!;
        var incomplete = (List<int>)mainForm.GetMethod("IncompleteReplacements",
            BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [plan])!;
        if (incomplete.Count > 0)
            throw new InvalidOperationException("Incomplete replacement steps: " + string.Join(", ", incomplete));
        var problem = mainForm.GetMethod("ReplacementPlanProblem",
            BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [plan]);
        if (problem is not null) throw new InvalidOperationException(problem.ToString());
        Console.WriteLine("PASS: saved recording passes replacement preflight without changing its JSON or captured evidence.");
        var conditionalRule = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsConditionalStep", BindingFlags.NonPublic | BindingFlags.Static)!;
        var optional = new PlanStep(1, "click",
            new ControlRef("FixtureApp", "Question", "6", "Yes", "ControlType.Button", "Button", "Question"),
            null, null, null, null, "Confirm the optional prompt if it appears.",
            Intent: "The pop up \"Question\", if present intent is to select Yes, otherwise proceed to next steps.");
        if (!(bool)conditionalRule.Invoke(null, [optional])! ||
            (bool)conditionalRule.Invoke(null, [optional with { Intent = null }])! ||
            (bool)conditionalRule.Invoke(null, [optional with
                { Target = optional.Target! with { Window = "Different question" } }])! ||
            (bool)conditionalRule.Invoke(null, [optional with
                { Target = optional.Target! with { Name = "No" } }])!)
            throw new InvalidOperationException("Optional dialog interpretation lost exact window/button scope.");
        var recordedOptional = plan.Steps.FirstOrDefault(step =>
            step.Target?.Window == "Sample application question");
        if (recordedOptional is not null &&
            !(bool)conditionalRule.Invoke(null, [recordedOptional])!)
            throw new InvalidOperationException("The recorded explicitly optional prompt remained mandatory.");
        Console.WriteLine("PASS: explicitly named recorded optional dialog is conditional; unrelated windows/buttons stay mandatory.");
        var compactor = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!;
        var compactedCopy = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [plan])!;
        var columnCopy = compactedCopy.Steps.Single(step => step.Action == "copy-column-until-empty" && step.Value == "K2");
        if (!compactedCopy.Steps.Any(step => step.Action == "copy-populated-columns" && step.Value == "E1:F1"))
            throw new InvalidOperationException("The populated two-column copy intent remained scalar copy.");
        if (!compactedCopy.Steps.Any(step =>
                step.Action == "set-chart-source-from-last-paste" && step.Value == "Chart 7" ||
                step is { Action: "set-chart-source-range", ExpectedState: "chart-source-verified",
                    Target: { Name: "Chart 7" }, Value: { Length: > 0 } }))
            throw new InvalidOperationException("The clarified chart-source action was not normalized.");
        if (!compactedCopy.Steps.Any(step => step.Action == "copy-column-until-empty" && step.Value == "C1") ||
            !compactedCopy.Steps.Any(step => step.Action == "extend-formula-to-adjacent-data-end" &&
                step.Value == "F" && step.ExpectedState == "adjacent-column:E"))
            throw new InvalidOperationException("Clarified copy and formula intents remained scalar copy or resize actions.");
        if (compactedCopy.Steps.Single(step => step.Action == "copy-column-until-empty" && step.Value == "C1")
                .ExpectedState != "copy-stop-values:[\"#VALUE!\"]")
            throw new InvalidOperationException("The explicit copy error boundary was lost.");
        var portableCell = columnCopy.Target! with { Process = "DifferentTableApp", Name = "J1",
            AutomationId = "J1", ClassName = "GenericTableCell" };
        var portablePlan = new ExecutionPlan(1, DateTimeOffset.Now, "Generic table intent",
            [new(1, "click", portableCell, null, null, null, null, null),
             new(2, "key", portableCell, null, "Control+C", null, null, null,
                Intent: "intent is to copy column J until before cell that contains empty value or the value \"STOP\""),
             new(3, "resize-column", portableCell, "-20", "right", null, null, null,
                Intent: "copy the formula in column J till before the empty cell in column H from the last populated formula")]);
        var portableResult = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [portablePlan])!;
        if (portableResult.Steps[1].Action != "copy-column-until-empty" ||
            portableResult.Steps[1].Value != "J1" ||
            portableResult.Steps[1].ExpectedState != "copy-stop-values:[\"STOP\"]" ||
            portableResult.Steps[2].Action != "extend-formula-to-adjacent-data-end" ||
            portableResult.Steps[2].ExpectedState != "adjacent-column:H")
            throw new InvalidOperationException("Intent planning hard-coded a program or column.");
        var adapters = typeof(ControlRef).Assembly.GetType("DesktopSteps.TableReplayAdapters")!;
        var capability = typeof(ControlRef).Assembly.GetType("DesktopSteps.TableCapability")!;
        try
        {
            adapters.GetMethod("Require", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [(nint)0, Enum.Parse(capability, "FormulaExtension")]);
            throw new InvalidOperationException("Unsupported table capability was accepted.");
        }
        catch (TargetInvocationException ex) when (ex.InnerException is NotSupportedException) { }
        Console.WriteLine("PASS: table intent planning supports arbitrary program/column identities; unavailable adapter stops explicitly.");
        if (columnCopy.Value != "K2" || columnCopy.OriginIntent != "intent is to copy column K until before first empty cell" ||
            compactedCopy.Steps.Any(step => step.Action == "click" && step.Target?.Name == "E2") ||
            !((ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [compactedCopy])!).Steps
                .SequenceEqual(compactedCopy.Steps))
            throw new InvalidOperationException("Saved column-copy repair lost explicit intent, start row or idempotence.");
        var manual = typeof(ControlRef).Assembly.GetType("DesktopSteps.ManualPdf")!;
        var pdfDirectory = Path.Combine(Path.GetTempPath(), "RSR-pdf-numbering-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(pdfDirectory);
        var pdfPath = Path.Combine(pdfDirectory, "manual.pdf");
        try
        {
            var withoutImages = compactedCopy with { Steps = compactedCopy.Steps.Select(step =>
                step with { Screenshot = null }).ToList() };
            await (Task)manual.GetMethod("CreateAsync")!.Invoke(null,
                [withoutImages, pdfDirectory, pdfPath, CancellationToken.None])!;
            var pdf = await File.ReadAllTextAsync(pdfPath);
            if (!pdf.Contains($"/Count {withoutImages.Steps.Count} ") ||
                withoutImages.Steps.Any(step => !pdf.Contains($"Step {step.Number}:")))
                throw new InvalidOperationException("PDF step numbers differ from the replay plan.");
        }
        finally { File.Delete(pdfPath); Directory.Delete(pdfDirectory); }
        Console.WriteLine("PASS: saved explicit K2 copy is attributed and idempotent; PDF preserves every replay step number.");
        var titleRule = compactor.GetMethod("IsTitleFocusBeforeCommand",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        var title = new PlanStep(1, "click",
            new ControlRef("FixtureApp", "Query", "TitleBar", "Query",
                "ControlType.TitleBar", "", "Query"), null, null, null, null, null);
        var yes = title with { Number = 2, Target = title.Target! with
            { Window = "Question", Name = "Yes", ControlType = "ControlType.Button" } };
        if (!(bool)titleRule.Invoke(null, [title, yes])! ||
            (bool)titleRule.Invoke(null, [title with { Intent = "Move the window." }, yes])! ||
            (bool)titleRule.Invoke(null, [title with { ExpectedState = "window-closed" }, yes])! ||
            (bool)titleRule.Invoke(null, [title, yes with { Action = "key" }])! ||
            (bool)titleRule.Invoke(null, [title, yes with
                { Target = yes.Target! with { Process = "OtherApp" } }])!)
            throw new InvalidOperationException("Title focus rule ignored meaningful actions or crossed applications.");
        var activationRule = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsWindowActivationClick", BindingFlags.NonPublic | BindingFlags.Static)!;
        var recordedQueryTitle = title with { Intent =
            "If the optional question appears click Yes; otherwise continue." };
        if (!(bool)activationRule.Invoke(null, [recordedQueryTitle])! ||
            (bool)activationRule.Invoke(null, [recordedQueryTitle with
                { ExpectedState = "window-moved" }])! ||
            (bool)activationRule.Invoke(null, [recordedQueryTitle with { Target =
                recordedQueryTitle.Target! with { Name = "Other title", ParentName = "Other parent" } }])! ||
            (bool)activationRule.Invoke(null, [recordedQueryTitle with { Target =
                recordedQueryTitle.Target! with { Window = "Find and Replace", Name = "Find and Replace",
                    ParentName = "Find and Replace" } }])!)
            throw new InvalidOperationException("Recorded title-bar focus was not scoped to verified window activation.");
        Console.WriteLine("PASS: a recorded title-bar focus activates its matching window even when a workflow note is attached; meaningful and special dialog title actions remain explicit.");
        var compact = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [plan])!;
        var datedRule = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsDatedSheetReference", BindingFlags.NonPublic | BindingFlags.Static)!;
        const string datedTemplate = "Backlog ({{next:Tuesday:dd MMM yy}})";
        if (!(bool)datedRule.Invoke(null, [datedTemplate, "Backlog (25 Sep 26)"])! ||
            (bool)datedRule.Invoke(null, [datedTemplate, "Other (25 Sep 26)"])! ||
            (bool)datedRule.Invoke(null, [datedTemplate, "Backlog (not a date)"])! ||
            (bool)datedRule.Invoke(null, ["Backlog", "Backlog (25 Sep 26)"])!)
            throw new InvalidOperationException("Dated sheet reference repair crossed explicit rename template scope.");
        var statusRule = compactor.GetMethod("IsStatusBackgroundBeforeNavigation",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        var statusBackground = title with { Target = title.Target! with
            { ControlType = "ControlType.StatusBar", Name = "Status Bar" } };
        var sheetTab = title with { Target = title.Target! with
            { ControlType = "ControlType.TabItem", Name = "Sheet" } };
        if (!(bool)statusRule.Invoke(null, [statusBackground, sheetTab])! ||
            (bool)statusRule.Invoke(null, [statusBackground with { Intent = "Change status settings" }, sheetTab])! ||
            (bool)statusRule.Invoke(null, [statusBackground with { ExpectedState = "changed" }, sheetTab])! ||
            (bool)statusRule.Invoke(null, [statusBackground, sheetTab with
                { Target = sheetTab.Target! with { Window = "Other window" } }])! ||
            (bool)statusRule.Invoke(null, [statusBackground, null])!)
            throw new InvalidOperationException("Status-bar navigation skipped a meaningful or unrelated action.");
        if (compact.Steps.Any(step => step.Target?.ControlType == "ControlType.StatusBar" &&
            step.Target.Name == "Status Bar"))
            throw new InvalidOperationException("Saved inert status-bar navigation remained.");
        Console.WriteLine("PASS: status-bar background before same-window tab navigation is omitted; intent, state and unrelated actions remain.");
        var sorted = new PlanStep(1, "ensure-state",
            new ControlRef("FixtureApp", "Query", "header-3", "Age", "ControlType.HeaderItem", "", "Header"),
            null, null, "sort:descending", null, null);
        var range = sorted with { Number = 2, Action = "key", Key = "Shift+End", ExpectedState = null,
            Target = sorted.Target! with { ControlType = "ControlType.ListItem",
                AutomationId = "ListViewItem-0", Name = "Old row", ParentName = "Rows" } };
        var refocus = range with { Number = 3, Action = "click", Key = null,
            Target = range.Target! with { ControlType = "ControlType.Text",
                Name = "Old value", ParentName = "Old row" } };
        var contextAll = refocus with { Number = 4, Action = "context-selection",
            ExpectedState = "selection-intent:all" };
        ExecutionPlan CompactSelection(params PlanStep[] steps) => (ExecutionPlan)compactor
            .GetMethod("Compact")!.Invoke(null,
                [new ExecutionPlan(1, DateTimeOffset.Now, "Sorted selection", steps.ToList())])!;
        var selected = CompactSelection(sorted, range, refocus, contextAll);
        if (selected.Steps.Count != 3 || selected.Steps[1].Action != "select-all-items" ||
            selected.Steps[1].Target != sorted.Target ||
            selected.Steps[2].Target != sorted.Target ||
            selected.Steps[2].ExpectedState != "selection-intent:all" ||
            CompactSelection(sorted, range, refocus, contextAll with { ExpectedState = null })
                .Steps.Any(step => step.Action == "select-all-items") ||
            CompactSelection(sorted, range, refocus, contextAll with
                { Target = contextAll.Target! with { Window = "Other query" } })
                .Steps.Any(step => step.Action == "select-all-items") ||
            CompactSelection(sorted, range).Steps.Any(step => step.Action == "select-all-items"))
            throw new InvalidOperationException("Sorted all-row selection was not scoped to explicit same-view context intent.");
        var endAfterAnchor = CompactSelection(sorted, refocus with { Number = 2 },
            range with { Number = 3, Key = "End" }, contextAll);
        if (endAfterAnchor.Steps.Count != 3 ||
            endAfterAnchor.Steps[1].Action != "select-all-items" ||
            endAfterAnchor.Steps[1].Target != sorted.Target ||
            endAfterAnchor.Steps[2].Action != "context-selection" ||
            CompactSelection(sorted, refocus with { Number = 2 }, range with { Number = 3, Key = "End" })
                .Steps.Any(step => step.Action == "select-all-items"))
            throw new InvalidOperationException("A stale clicked value before End was not safely replaced by verified all-row selection.");
        if (compact.Steps.Any(step => step.Action == "key" && step.Key == "Shift+End" &&
            step.Target?.ControlType == "ControlType.ListItem" &&
            step.Target.Window == "Sample Case Query - Case Triage (any)"))
            throw new InvalidOperationException("Saved sorted-list selection still relies on incidental row focus.");
        var twice = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [selected])!;
        if (!twice.Steps.SequenceEqual(selected.Steps))
            throw new InvalidOperationException("Sorted selection compaction is not idempotent.");
        Console.WriteLine("PASS: sorted range with explicit all-row context uses verified dynamic selection; unrelated and partial selections are preserved.");
        var recoveredFill = compact.Steps.FirstOrDefault(step =>
            step.Action == "fill-down-to-adjacent-data-end" && step.Target?.Name == "F35");
        var seed = new PlanStep(1, "type",
            new ControlRef("EXCEL", "Workbook", "L9", "L9", "ControlType.DataItem", "XLSpreadsheetCell", "Table"),
            "FL", null, null, null, null);
        var nav = seed with { Action = "scroll", Value = null, Key = "Horizontal",
            Target = seed.Target! with { ControlType = "ControlType.ScrollBar" } };
        var directSeed = seed with { TargetStrategy = "first-visible-filtered-row" };
        if (CompactSelection(nav, nav, directSeed).Steps.Count != 1 ||
            CompactSelection(nav with { Intent = "Inspect the worksheet" }, directSeed).Steps.Count != 2 ||
            CompactSelection(nav, directSeed with { Target = seed.Target! with { Window = "Other workbook" } })
                .Steps.Count != 2 ||
            CompactSelection(nav, seed).Steps.Count != 2)
            throw new InvalidOperationException("Direct filtered navigation removed meaningful or unrelated scrolling.");
        var endpoint = seed with { Number = 2, Action = "click", Value = null,
            Target = seed.Target! with { Name = "F35", AutomationId = "", ClassName = "" },
            Intent = "use the first visible filtered row rather than a fixed row number. intent is to copy the cell value to below cells in column L until before the empty adjacent cell in K column" };
        if (CompactSelection(seed, endpoint).Steps[1].Action != "fill-down-to-adjacent-data-end" ||
            CompactSelection(seed, endpoint with { Intent = null }).Steps[1].Action != "click" ||
            CompactSelection(seed, endpoint with { Intent = endpoint.Intent!.Replace("column L", "column M") })
                .Steps[1].Action != "click" ||
            CompactSelection(seed, endpoint with { Target = endpoint.Target! with { Window = "Other workbook" } })
                .Steps[1].Action != "click" ||
            CompactSelection(endpoint).Steps[0].Action != "click")
            throw new InvalidOperationException("Explicit fill interpretation guessed a source or crossed workbook/column scope.");
        if (plan.Steps.Any(step => step.Target?.Name == "F35" &&
                step.Intent?.Contains("empty adjacent cell in K column") == true) &&
            (recoveredFill is null || recoveredFill.ExpectedState != "adjacent-column:K" ||
             compact.Steps[compact.Steps.IndexOf(recoveredFill) - 1].TargetStrategy != "first-visible-filtered-row"))
            throw new InvalidOperationException("The first visible FL fill instruction remained an endpoint click.");
        var gridRule = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsGridFocusBeforeCommand", BindingFlags.NonPublic | BindingFlags.Static)!;
        var pane = new PlanStep(1, "click",
            new ControlRef("EXCEL", "Workbook", "Workbook", "Workbook", "ControlType.Pane", "ExcelGrid", ""),
            null, null, null, null, null);
        var page = pane with { Number = 2, Target = pane.Target! with
            { ControlType = "ControlType.Button", ClassName = "NetUIRepeatButton", Name = "Page left" } };
        if (!(bool)gridRule.Invoke(null, [pane, page])! ||
            !(bool)gridRule.Invoke(null, [pane, pane with { Action = "filter-values",
                Target = pane.Target! with { ControlType = "ControlType.MenuItem", AutomationId = "Dropdown" } }])! ||
            (bool)gridRule.Invoke(null, [pane with { Intent = "Select worksheet contents" }, page])! ||
            (bool)gridRule.Invoke(null, [pane with { Target = pane.Target! with { ClassName = "OtherPane" } }, page])! ||
            (bool)gridRule.Invoke(null, [pane, page with { Target = page.Target! with { Window = "Other workbook" } }])!)
            throw new InvalidOperationException("Excel grid focus rule crossed intent or window boundaries.");
        Console.WriteLine("PASS: saved FL intent becomes visible fill; inert Excel grid focus is skipped only before same-workbook commands.");
        var filterStep = pane with { Action = "filter-values", Value = "[]",
            ExpectedState = "filter:all", Target = pane.Target! with
                { ControlType = "ControlType.MenuItem", AutomationId = "Dropdown", ParentName = "G1" } };
        var pageRight = page with { Target = page.Target! with { Name = "Page right" } };
        if (CompactSelection(pageRight, filterStep).Steps.Count != 1 ||
            CompactSelection(pageRight with { Intent = "Inspect this view" }, filterStep).Steps.Count != 2)
            throw new InvalidOperationException("Filter navigation compaction discarded intent or retained redundant paging.");
        var fallbackIntent = optional with { Intent =
            "The pop up \"Question\", if present click Yes, otherwise proceed. Automatic refresh occurs if present so no need to click refresh button, otherwise, do click it." };
        var branch = CompactSelection(fallbackIntent, sorted);
        if (branch.Steps.Count != 3 || branch.Steps[1].Action != "click-if-previous-absent" ||
            branch.Steps[1].TargetStrategy != "unique-refresh-command" ||
            branch.Steps[1].OriginIntent != fallbackIntent.Intent ||
            !((ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [branch])!).Steps.SequenceEqual(branch.Steps) ||
            CompactSelection(optional, sorted).Steps.Count != 2 ||
            CompactSelection(fallbackIntent, sorted with { Target = sorted.Target! with { Process = "OtherApp" } })
                .Steps.Count != 2)
            throw new InvalidOperationException("Explicit refresh fallback lost provenance, scope or idempotence.");
        if (recordedOptional is not null && !compact.Steps.Any(step =>
            step.TargetStrategy == "unique-refresh-command" && step.Action == "click-if-previous-absent"))
            throw new InvalidOperationException("The saved explicit absent-dialog Refresh instruction was omitted.");
        Console.WriteLine("PASS: explicit Refresh fallback is persisted, scoped and idempotent; grid focus before filtering is redundant.");
        if (plan.Steps.Any(step => step.Target?.ControlType == "ControlType.TitleBar" &&
                step.Target.Window == "Sample Case Query") &&
            compact.Steps.Any(step => step.Target?.ControlType == "ControlType.TitleBar" &&
                step.Target.Window == "Sample Case Query"))
            throw new InvalidOperationException("The recorded redundant query-title navigation remained in the compact plan.");
        Console.WriteLine("PASS: title focus before same-application modal command is skipped; intent, state and unrelated actions remain.");
        await CheckDialogCaptureAsync();
    }

    private static async Task RunAsync()
    {
        CheckKeyboardMappings();
        CheckGridFocusBeforeCommand();
        await CheckGridFocusReplayAsync();
        await CheckOptionalBranchReplayAsync(present: true);
        await CheckOptionalBranchReplayAsync(present: false);
        await CheckFrozenKeyboardAsync();
        await CheckWheelDuringUiStallAsync();
        await FoundryChecks.RunAsync();
        await CheckDialogCaptureAsync();
        CheckIncompleteReplacementIntent();
        await CheckUnscopedFilterPreflightAsync();
        await CheckAllRowSelectionAsync();
        await CheckOptionalDialogAsync();
        await CheckAsync("replay");
        await CheckAsync("resized");
        await CheckAsync("nested");
        await CheckAsync("capture");
        await CheckAsync("legacy");
        await CheckAsync("duplicate");
        await CheckAsync("disabled");
        await CheckAsync("wrong-window");
        await CheckAsync("wrong-process");
        await CheckAsync("missing");
    }

    private static async Task CheckFoundryAsync(string configuration)
    {
        var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.Foundry")!;
        var foundry = Activator.CreateInstance(type, [configuration])!;
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
        var task = (Task<string>)type.GetMethod("AskAsync")!.Invoke(foundry,
            ["Return ONLY JSON: {\"summary\":\"Connectivity check\",\"steps\":[]}.", timeout.Token])!;
        Console.WriteLine(await task);
        var target = new ControlRef("SyntheticApp", "Synthetic window", null, "Sample button",
            "ControlType.Button", null, null);
        var events = new List<RecordedEvent>
        {
            new(DateTimeOffset.Now, "click", target, null, null, null, null, null)
        };
        var plan = await (Task<ExecutionPlan>)type.GetMethod("PlanAsync")!.Invoke(foundry,
            [events, timeout.Token, null])!;
        if (plan.Steps.Count != 1 || plan.Steps[0].Target != target ||
            plan.Summary.Contains("fallback", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Synthetic Foundry planning did not preserve the recorded action.");
        Console.WriteLine("PASS: Foundry generated a complete synthetic plan without fallback.");
    }

    private static void CheckKeyboardMappings()
    {
        var method = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("ToSendKeys", BindingFlags.NonPublic | BindingFlags.Static)!;
        foreach (var (key, expected) in new[]
        {
            ("Insert", "{INSERT}"), ("Control, Insert", "^{INSERT}"),
            ("Ctrl+Insert", "^{INSERT}"), ("Shift, Insert", "+{INSERT}"),
            ("Control, Shift, Insert", "^+{INSERT}"), ("End", "{END}"),
            ("Shift, End", "+{END}"), ("Control, A", "^a")
        })
            if ((string)method.Invoke(null, [key])! != expected)
                throw new InvalidOperationException($"Incorrect keyboard replay mapping: {key}.");
        Console.WriteLine("PASS: Insert and modifier keyboard mappings.");
    }

    private static void CheckGridFocusBeforeCommand()
    {
        var method = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsGridFocusBeforeCommand", BindingFlags.NonPublic | BindingFlags.Static)!;
        var grid = new ControlRef("GridHost", "Grid window", "summaryGrid", "",
            "ControlType.DataGrid", null, null);
        var click = new PlanStep(1, "click", grid, null, null, null, null, null);
        var command = click with { Number = 2, Target = grid with
            { Name = "Refresh", ControlType = "ControlType.Button" } };
        foreach (var type in new[] { "ControlType.DataGrid", "ControlType.Table" })
            if (!(bool)method.Invoke(null, [click with { Target = grid with { ControlType = type } }, command])!)
                throw new InvalidOperationException("Grid focus before a command was not recognized.");
        foreach (var candidate in new[]
        {
            click with { Action = "context-click" },
            click with { ExpectedState = "selected rows" },
            click with { Intent = "Select a particular row" },
            click with { Value = "selection-count:3" },
            click with { Target = grid with { Name = "Specific data" } },
            click with { Target = grid with { ControlType = "ControlType.DataItem" } }
        })
            if ((bool)method.Invoke(null, [candidate, command])!)
                throw new InvalidOperationException("Meaningful grid action was incorrectly skipped.");
        if ((bool)method.Invoke(null, [click, command with { Action = "key", Key = "Delete" }])! ||
            (bool)method.Invoke(null, [click, command with { Target = command.Target! with { Process = "OtherHost" } }])! ||
            (bool)method.Invoke(null, [click, null])!)
            throw new InvalidOperationException("Grid focus skip ignored the dependent command.");
        if ((bool)method.Invoke(null, [click, command with { Target = command.Target! with { Window = "Other window" } }])!)
            throw new InvalidOperationException("Grid focus skip crossed windows.");
        Console.WriteLine("PASS: DataGrid/Table focus before command; meaningful selection and key actions preserved.");
    }

    private static async Task CheckGridFocusReplayAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --grid-focus")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Grid test host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) break;
                await Task.Delay(50, timeout.Token);
            }
            var grid = new ControlRef(host.ProcessName, HostTitle, "summaryGrid", "",
                "ControlType.DataGrid", null, null);
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Grid focus then refresh",
            [
                new(1, "click", grid, null, null, null, null, "Focus the grid."),
                new(2, "click", grid with { AutomationId = "commandButton", Name = "Run command",
                    ControlType = "ControlType.Button" }, null, null, null, null, null)
            ]);
            var messages = new List<string>();
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            await (Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!;
            host.Refresh();
            if (host.MainWindowTitle != HostTitle + " - Refreshed" ||
                !messages.Any(message => message.StartsWith("Skipped step 1:")))
                throw new InvalidOperationException("Grid-background replay did not invoke its following command.");
            var label = grid with { Window = host.MainWindowTitle, AutomationId = "statusLabel",
                Name = "Ready", ControlType = "ControlType.Text", ClassName = "WindowsForms10.Static.app.test" };
            var labelPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Status label then command",
            [
                new(1, "click", label, null, null, null, null, null),
                new(2, "click", label with { AutomationId = "commandButton", Name = "Run command",
                    ControlType = "ControlType.Button", ClassName = null }, null, null, null, null, null)
            ]);
            messages.Clear();
            await (Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [labelPlan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!;
            if (!messages.Any(message => message.Contains("static status label")) ||
                !messages.Any(message => message.StartsWith("Completed step 2")))
                throw new InvalidOperationException("Static label navigation did not preserve the following command.");
            Console.WriteLine("PASS: live grid-background replay invokes the independently targeted command button.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    private static async Task CheckOptionalBranchReplayAsync(bool present)
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!,
            "--host --grid-focus" + (present ? " --optional --optional-ready" : ""))
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Optional branch host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) break;
                await Task.Delay(50, timeout.Token);
            }
            var target = new ControlRef(host.ProcessName, HostTitle, "commandButton", "Run command",
                "ControlType.Button", null, null);
            var note = "If present, click Yes in \"RSR optional regression dialog\"; otherwise click Run command.";
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Optional branch",
            [
                new(1, "click", target with { AutomationId = "summaryGrid", Name = "",
                    ControlType = "ControlType.DataGrid" }, null, null, null, null, null),
                new(2, "optional-click", target with { Window = "RSR optional regression dialog",
                    Name = "Yes", AutomationId = null }, null, null, null, null, null, note, OriginIntent: note),
                new(3, "click-if-previous-absent", target, null, null, null, null, null, note)
            ]);
            if (present)
                plan.Steps.Add(new(4, "click", target with { AutomationId = "nextButton", Name = "Next action" },
                    null, null, null, null, null));
            var messages = new List<string>();
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            await (Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), timeout.Token])!;
            host.Refresh();
            var expected = HostTitle + (present ? " - Ready action executed" : " - Refreshed");
            if (host.MainWindowTitle != expected ||
                !messages.Any(message => message.StartsWith("Skipped step 1:")) ||
                present && !messages.Any(message => message.StartsWith("Skipped conditional step 3:")) ||
                !present && !messages.Any(message => message.StartsWith("Completed step 3")))
                throw new InvalidOperationException("Optional branch ran the wrong command: " + string.Join("; ", messages));
            if (present && !messages.Any(message => message.StartsWith("The resulting view is ready")))
                throw new InvalidOperationException("Optional branch did not wait for its delayed resulting view.");
            var waitReady = executor.GetMethod("WaitForReadyControlAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
            try
            {
                await (Task)waitReady.Invoke(null, [target with { Window = host.MainWindowTitle, Name = "Missing control",
                    AutomationId = "missing" }, timeout.Token, TimeSpan.FromMilliseconds(300), null,
                    "The optional action's resulting view"])!;
                throw new InvalidOperationException("Missing resulting control was accepted as ready.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("did not expose an enabled, populated view")) { }
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            try
            {
                await (Task)waitReady.Invoke(null, [target, cancelled.Token, TimeSpan.FromSeconds(2), null,
                    "The optional action's resulting view"])!;
                throw new InvalidOperationException("Readiness wait ignored cancellation.");
            }
            catch (OperationCanceledException) { }
            Console.WriteLine($"PASS: optional dialog {(present ? "appears; fallback is skipped" : "absent; recorded fallback executes")}.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    private static async Task CheckFrozenKeyboardAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --selection")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Keyboard host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            var recorderType = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
            var capture = recorderType.GetMethod("CaptureKeyboardTarget", BindingFlags.NonPublic | BindingFlags.Static)!;
            object? snapshot = null;
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) snapshot = capture.Invoke(null, null);
                if (snapshot is not null &&
                    ((ControlRef?)snapshot.GetType().GetProperty("Target")!.GetValue(snapshot))?.Window == HostTitle) break;
                await Task.Delay(50, timeout.Token);
            }
            if (snapshot is null) throw new InvalidOperationException("Keyboard snapshot was not captured.");
            var snapshotType = snapshot.GetType();
            var frozen = (ControlRef?)snapshotType.GetProperty("Target")!.GetValue(snapshot);
            if (frozen?.Process != host.ProcessName || frozen.Window != HostTitle)
                throw new InvalidOperationException("Keyboard snapshot came from the wrong application.");
            using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
            using var other = new Form { Text = "Different focused window" };
            other.Show(); other.Activate();
            Application.DoEvents();
            var externalTarget = frozen with { Process = "ExternalKeyboardHost" };
            var external = Activator.CreateInstance(snapshotType,
                [snapshotType.GetProperty("Window")!.GetValue(snapshot)!,
                 snapshotType.GetProperty("Focus")!.GetValue(snapshot)!,
                 externalTarget, snapshotType.GetProperty("Element")!.GetValue(snapshot),
                 snapshotType.GetProperty("Screen")!.GetValue(snapshot)!])!;
            var infoType = recorderType.GetNestedType("KeyboardInfo", BindingFlags.NonPublic)!;
            var info = Activator.CreateInstance(infoType)!;
            infoType.GetField("VirtualKey")!.SetValue(info, (uint)Keys.A);
            var process = recorderType.GetMethod("ProcessKeyboard", BindingFlags.NonPublic | BindingFlags.Instance)!;
            var at = DateTimeOffset.Now.AddSeconds(-3);
            process.Invoke(recorder, [(nint)0x100, info, "Control", at, external, externalTarget]);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            if (events.Count != 1 || events[0].Target != externalTarget || events[0].Key != "Control+A" ||
                events[0].At != at)
                throw new InvalidOperationException("Queued keyboard input was retargeted or retimed.");
            process.Invoke(recorder, [(nint)0x100, info, "Control", at, null, null]);
            if (events.Count != 2 || events[1].Kind != "unresolved-input" || events[1].Target is not null)
                throw new InvalidOperationException("Unresolved keyboard input guessed a live target.");
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Delayed input", [
                new(1, "unresolved-input", null, null, null, null, null, null)]);
            var messages = new List<string>();
            try
            {
                await (Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                        (Func<string, bool>)(_ => false), timeout.Token])!;
                throw new InvalidOperationException("Unresolved input was replayed.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("no verified target"))
            {
                if (messages.Count != 0) throw new InvalidOperationException("Delayed-input validation happened after replay.");
            }
            Console.WriteLine("PASS: queued keyboard identity/modifiers/time survive focus changes; unresolved input blocks replay.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern void mouse_event(uint flags, uint x, uint y, uint data, nuint extra);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern void keybd_event(byte key, byte scan, uint flags, nuint extra);

    private static async Task CheckWheelDuringUiStallAsync()
    {
        var fixture = Path.Combine(AppContext.BaseDirectory, "RSRHookFixture.exe");
        File.Copy(Environment.ProcessPath!, fixture, overwrite: true);
        using var host = Process.Start(new ProcessStartInfo(fixture, "--host --selection")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Wheel host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) break;
                await Task.Delay(50, timeout.Token);
            }
            var window = AutomationElement.FromHandle(host.MainWindowHandle);
            var list = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Selection regression list"))
                ?? throw new InvalidOperationException("Wheel fixture list was not initialized.");
            list.SetFocus();
            for (var attempt = 0; attempt < 20 &&
                AutomationElement.FocusedElement?.Current.ProcessId != host.Id; attempt++)
                await Task.Delay(50, timeout.Token);
            if (AutomationElement.FocusedElement?.Current.ProcessId != host.Id)
                throw new InvalidOperationException("Wheel fixture could not establish keyboard focus before injection.");
            var bounds = window.Current.BoundingRectangle;
            var previousCursor = Cursor.Position;
            Cursor.Position = new System.Drawing.Point((int)bounds.Left + 120, (int)bounds.Top + 100);
            var recorderType = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
            using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
            var output = Path.Combine(Path.GetTempPath(), "rsr-wheel-" + Guid.NewGuid().ToString("N"));
            try
            {
                recorderType.GetMethod("Start")!.Invoke(recorder, [output]);
                var hookThread = (Thread?)recorderType.GetField("hookThread", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .GetValue(recorder);
                if (hookThread is null || hookThread.ManagedThreadId == Environment.CurrentManagedThreadId)
                    throw new InvalidOperationException("Input hooks still share the UI thread.");
                var start = DateTimeOffset.Now;
                var injected = Task.Run(() =>
                {
                    Thread.Sleep(100);
                    for (var i = 0; i < 12; i++)
                    {
                        mouse_event(0x0800, 0, 0, 120, 0);
                        Thread.Sleep(30);
                    }
                    keybd_event((byte)Keys.ControlKey, 0, 0, 0);
                    Thread.Sleep(30);
                    keybd_event((byte)Keys.A, 0, 0, 0);
                    keybd_event((byte)Keys.A, 0, 2, 0);
                    keybd_event((byte)Keys.ControlKey, 0, 2, 0);
                });
                // Deliberately stop the recorder UI pump while native hook delivery continues.
                Thread.Sleep(3200);
                await injected;
                Application.DoEvents();
                recorderType.GetMethod("Stop")!.Invoke(recorder, null);
                var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
                if (events.Any(item => item.Kind == "unresolved-input") ||
                    events.Where(item => item.Kind == "scroll").Sum(item => int.Parse(item.Value!)) != 1440 ||
                    !events.Any(item => item.Kind == "key" && item.Key == "Control+A") ||
                    events.Any(item => item.Target?.Process != host.ProcessName ||
                        item.At < start || item.At > start.AddSeconds(2)))
                    throw new InvalidOperationException("Wheel input was lost, misidentified or retimed during a UI stall: " +
                        System.Text.Json.JsonSerializer.Serialize(events));
                recorderType.GetMethod("Start")!.Invoke(recorder, [output]);
                recorderType.GetMethod("Stop")!.Invoke(recorder, null);
                Console.WriteLine("PASS: native wheel burst and Ctrl+A survive a 3.2-second recorder UI stall with exact delta/time/owner; hook restart succeeds.");
            }
            finally
            {
                Cursor.Position = previousCursor;
                foreach (var file in Directory.GetFiles(output)) File.Delete(file);
                Directory.Delete(output);
            }
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
            File.Delete(fixture);
        }
    }

    private static async Task CheckOptionalDialogAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --optional")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Optional test host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(40));
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) break;
                await Task.Delay(50, timeout.Token);
            }
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            var wait = executor.GetMethod("WaitForOptionalStepAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
            var ready = executor.GetMethod("OptionalNextControlReadyAsync", BindingFlags.NonPublic | BindingFlags.Static)!;
            var next = new ControlRef(host.ProcessName, HostTitle, null, "Work menu", "ControlType.Button", null, "statusStrip");
            var optional = next with { Window = "RSR optional regression dialog", Name = "Yes", ParentName = null };
            var step = new PlanStep(1, "click", optional, null, null, null, null, "If present, click Yes; otherwise continue.");
            var watch = Stopwatch.StartNew();
            if (!await (Task<bool>)wait.Invoke(null, [step, timeout.Token, TimeSpan.FromSeconds(20)])!)
                throw new InvalidOperationException("Dialog appearing after eight seconds was skipped.");
            if (watch.Elapsed < TimeSpan.FromSeconds(8))
                throw new InvalidOperationException("Delayed dialog fixture did not exercise the old timeout.");
            host.Refresh();
            if (host.MainWindowTitle != HostTitle + " - Optional handled")
                throw new InvalidOperationException("Optional dialog was not invoked automatically.");
            if (!await (Task<bool>)ready.Invoke(null, [next with { Window = host.MainWindowTitle }, timeout.Token])!)
                throw new InvalidOperationException("Available next control was not recognized.");
            if (await (Task<bool>)ready.Invoke(null, [next with { Name = "Missing next control" }, timeout.Token])!)
                throw new InvalidOperationException("Window presence alone was accepted as readiness.");
            if (await (Task<bool>)ready.Invoke(null, [null, timeout.Token])!)
                throw new InvalidOperationException("Missing next target was accepted as readiness.");
            if (await (Task<bool>)wait.Invoke(null, [step, timeout.Token, TimeSpan.FromMilliseconds(200)])!)
                throw new InvalidOperationException("Absent optional dialog was reported as completed.");
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            try
            {
                await (Task<bool>)wait.Invoke(null, [step, cancelled.Token, TimeSpan.FromSeconds(20)])!;
                throw new InvalidOperationException("Optional wait ignored cancellation.");
            }
            catch (OperationCanceledException) { }
            Console.WriteLine("PASS: delayed optional dialog, absent dialog, next-control readiness and cancellation.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    private static async Task CheckAllRowSelectionAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --selection")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Selection host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            AutomationElement? list = null;
            for (var attempt = 0; attempt < 100 && list is null; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0)
                    list = AutomationElement.FromHandle(host.MainWindowHandle).FindFirst(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.NameProperty, "Selection regression list"));
                if (list is null) await Task.Delay(50, timeout.Token);
            }
            if (list is null) throw new InvalidOperationException("Selection fixture list unavailable.");
            var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
            var ensure = executor.GetMethod("EnsureRecordedAllRowSelection", BindingFlags.NonPublic | BindingFlags.Static)!;
            var all = executor.GetMethod("AllListItemsSelected", BindingFlags.NonPublic | BindingFlags.Static)!;
            var window = AutomationElement.FromHandle(host.MainWindowHandle);
            var header = window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.HeaderItem))
                ?? throw new InvalidOperationException("Selection fixture header unavailable.");
            var first = list.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem))
                ?? throw new InvalidOperationException("Selection fixture row unavailable.");
            var target = new ControlRef(host.ProcessName, HostTitle, header.Current.AutomationId,
                header.Current.Name, "ControlType.HeaderItem", header.Current.ClassName, "Header");
            var selectionPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Selection independent of focus",
                [new(1, "select-all-items", target, null, null, null, null, null)]);
            var step = new PlanStep(16, "context-selection", null, null, null, null, null, null);
            try
            {
                ensure.Invoke(null, [list, step, timeout.Token]);
                throw new InvalidOperationException("Missing all-row intent was accepted.");
            }
            catch (TargetInvocationException ex) when (ex.InnerException is InvalidOperationException &&
                ex.InnerException.Message.Contains("no all-row selection intent")) { }
            if ((bool)all.Invoke(null, [list])!)
                throw new InvalidOperationException("Selection changed without recorded intent.");
            ensure.Invoke(null, [list, step with { ExpectedState = "selection-intent:all" }, timeout.Token]);
            if (!(bool)all.Invoke(null, [list])!)
                throw new InvalidOperationException("All 51 current rows were not selected.");
            ensure.Invoke(null, [list, step with { ExpectedState = "selection-intent:all" }, timeout.Token]);
            using (var focusWindow = new Form { Text = "RSR unrelated focus fixture" })
            {
                focusWindow.Show();
                focusWindow.Activate();
                Application.DoEvents();
                await ((Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                    [selectionPlan, (Action<string>)Console.WriteLine, (Func<string, bool>)(_ => false),
                     (Func<string, bool>)(_ => false), timeout.Token])!).WaitAsync(timeout.Token);
            }
            if (!(bool)all.Invoke(null, [list])!)
                throw new InvalidOperationException("Verified selection depended on a previously focused row.");
            Console.WriteLine("PASS: explicit all-row intent selects and verifies 51 live rows; missing intent is rejected.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    private static void CheckIncompleteReplacementIntent()
    {
        var mainForm = typeof(ControlRef).Assembly.GetType("DesktopSteps.MainForm")!;
        var incomplete = mainForm.GetMethod("IncompleteReplacements", BindingFlags.NonPublic | BindingFlags.Static)!;
        var target = new ControlRef("EXCEL", "Find and Replace", null, "Replace",
            "ControlType.TabItem", null, null);
        var steps = new List<PlanStep>
        {
            new(1, "click", target, null, null, null, null, null, "Always replace the entered text."),
            new(2, "click", target with { Window = "Workbook", Name = "Filter" }, null, null, null, null, null)
        };
        var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Replacement intent", steps);
        if (((List<int>)incomplete.Invoke(null, [plan])!).Single() != 1)
            throw new InvalidOperationException("Replacement intent without a captured command was not blocked.");
        steps.Insert(1, new(2, "replace-all", target with { Name = "Replace All", ControlType = "ControlType.Button" },
            "old", null, "new", null, null, "Replace old with new."));
        if (((List<int>)incomplete.Invoke(null, [plan])!).Count != 0)
            throw new InvalidOperationException("Captured replacement command was incorrectly flagged.");
        var note = steps[0].Intent;
        var restarted = new ExecutionPlan(1, DateTimeOffset.Now, "Resumed replacement", [
            steps[0],
            steps[^1] with { Number = 2 },
            steps[0] with { Number = 3 },
            steps[1] with { Number = 4, Intent = note },
            steps[^1] with { Number = 5 }]);
        if (((List<int>)incomplete.Invoke(null, [restarted])!).Count != 0)
            throw new InvalidOperationException("A preparatory tab carrying contextual intent was falsely treated as a missing replacement.");
        var typedMissing = restarted with { Steps = [.. restarted.Steps] };
        typedMissing.Steps[0] = typedMissing.Steps[0] with { Action = "type", Value = "Unsubmitted text" };
        if (((List<int>)incomplete.Invoke(null, [typedMissing])!).Single() != 1)
            throw new InvalidOperationException("An abandoned typed replacement was silently satisfied by a later session.");
        var differentIntent = restarted with { Steps = [.. restarted.Steps] };
        differentIntent.Steps[3] = differentIntent.Steps[3] with { Intent = "Replace something else." };
        if (((List<int>)incomplete.Invoke(null, [differentIntent])!).Single() != 1)
            throw new InvalidOperationException("An unrelated later replacement satisfied a missing command.");
        Console.WriteLine("PASS: uncaptured replacement intent is rejected; captured Replace All remains valid.");
    }

    private static async Task CheckUnscopedFilterPreflightAsync()
    {
        var executor = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!;
        var filter = new ControlRef("SyntheticGrid", "Synthetic workbook", "Dropdown",
            "No filter applied", "ControlType.MenuItem", null, null);
        var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Missing filter scope",
            [new(1, "click", filter, null, null, null, null, null)]);
        var messages = new List<string>();
        try
        {
            await (Task<bool>)executor.GetMethod("ReplayAsync")!.Invoke(null,
                [plan, (Action<string>)messages.Add, (Func<string, bool>)(_ => false),
                    (Func<string, bool>)(_ => false), CancellationToken.None])!;
            throw new InvalidOperationException("Unscoped filter was accepted.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("did not save the column parent"))
        {
            if (messages.Count != 0) throw new InvalidOperationException("Filter validation occurred after replay started.");
        }
        Console.WriteLine("PASS: filter without its column parent is blocked before replay starts.");
    }

    private static async Task CheckDialogCaptureAsync()
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--host --dialog-capture")
        { UseShellExecute = false }) ?? throw new InvalidOperationException("Dialog capture host could not start.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) break;
                await Task.Delay(50, timeout.Token);
            }
            var window = AutomationElement.FromHandle(host.MainWindowHandle);
            var recorderType = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
            var snapshotMethod = recorderType.GetMethod("CreateMouseDownSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
            object? saved = null;
            int x = 0, y = 0;
            foreach (var name in new[] { "Continue with selection", "OK" })
            {
                var control = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, name));
                var bounds = control.Current.BoundingRectangle;
                x = (int)(bounds.Left + bounds.Width / 2);
                y = (int)(bounds.Top + bounds.Height / 2);
                var snapshot = snapshotMethod.Invoke(null, [window, host.MainWindowHandle, x, y])!;
                var captured = (ControlRef)snapshot.GetType().GetProperty("Target")!.GetValue(snapshot)!;
                if (captured.Name != name || captured.Window != "RSR transient dialog fixture" ||
                    captured.ControlType != (name == "OK" ? "ControlType.Button" : "ControlType.RadioButton"))
                    throw new InvalidOperationException("Coarse dialog hit was not resolved to its actual child command.");
                saved = snapshot;
            }
            using (var retryRecorder = (IDisposable)Activator.CreateInstance(recorderType)!)
            {
                var retryEvents = (List<RecordedEvent>)recorderType.GetField("events",
                    BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(retryRecorder)!;
                retryEvents.Add(new(DateTimeOffset.Now, "click",
                    new ControlRef(host.ProcessName, window.Current.Name, null, window.Current.Name,
                        "ControlType.Window", null, null), null, null, null, null, null,
                    ClickX: x, ClickY: y));
                recorderType.GetMethod("QueueDialogCommandRetry",
                    BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(retryRecorder,
                        [0, host.MainWindowHandle, x, y]);
                for (var attempt = 0; attempt < 40 && retryEvents[0].Target?.ControlType != "ControlType.Button"; attempt++)
                    await Task.Delay(50);
                if (retryEvents[0].Target?.Name != "OK" || retryEvents[0].Target?.ControlType != "ControlType.Button" ||
                    (bool)recorderType.GetMethod("IsUnresolvedDialogClick")!.Invoke(retryRecorder, [0])!)
                    throw new InvalidOperationException("UIA dialog retry failed or still reports the repaired click as unresolved.");
                retryEvents.Add(new(DateTimeOffset.Now, "unresolved-click", null, null, null, null, null, null));
                if (!(bool)recorderType.GetMethod("IsUnresolvedDialogClick")!.Invoke(retryRecorder, [1])!)
                    throw new InvalidOperationException("An unresolved menu click no longer requests attention.");
                var recovery = (string)recorderType.GetMethod("DescribeRecovery")!.Invoke(retryRecorder, [1])!;
                if (!recovery.Contains("action 2") || !recovery.Contains("Last kept action 1") ||
                    !recovery.Contains("actions 2-2"))
                    throw new InvalidOperationException("Recovery instructions do not identify the affected action and retained boundary.");
            }
            var snapshotType = saved!.GetType();
            var frozen = (ControlRef)snapshotType.GetProperty("Target")!.GetValue(saved)!;
            var external = Activator.CreateInstance(snapshotType,
                [window, frozen with { Process = "RecordedDialogHost" }])!;
            snapshotType.GetProperty("WindowHandle")!.SetValue(external, host.MainWindowHandle);
            snapshotType.GetProperty("IsDialog")!.SetValue(external, true);
            host.CloseMainWindow();
            if (!host.WaitForExit(3000)) throw new InvalidOperationException("Transient fixture did not close.");
            using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
            var pointType = recorderType.GetNestedType("Point", BindingFlags.NonPublic)!;
            var point = Activator.CreateInstance(pointType)!;
            pointType.GetField("X")!.SetValue(point, x);
            pointType.GetField("Y")!.SetValue(point, y);
            var infoType = recorderType.GetNestedType("MouseInfo", BindingFlags.NonPublic)!;
            var info = Activator.CreateInstance(infoType)!;
            infoType.GetField("Point")!.SetValue(info, point);
            recorderType.GetMethod("ProcessMouse", BindingFlags.NonPublic | BindingFlags.Instance)!
                .Invoke(recorder, [0, (nint)0x201, info, null, null, null, external, DateTimeOffset.Now]);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            if (events.Count != 1 || events[0].Target?.Name != "OK" ||
                events[0].Target?.Window != frozen.Window)
                throw new InvalidOperationException("Queued capture lost the command after its dialog closed.");
            Console.WriteLine("PASS: coarse dialog hits resolve radio/button controls and retain command after closing.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(3000)) host.Kill(entireProcessTree: true);
            }
        }
    }

    private static async Task CheckAsync(string scenario)
    {
        using var host = Process.Start(new ProcessStartInfo(Environment.ProcessPath!,
            scenario is "duplicate" or "disabled" or "resized" or "nested" or "capture"
                ? $"--host --{(scenario == "capture" ? "nested" : scenario)}" : "--host")
        {
            UseShellExecute = false
        }) ?? throw new InvalidOperationException("Test host could not be started.");
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            for (var attempt = 0; attempt < 100; attempt++)
            {
                host.Refresh();
                if (host.MainWindowHandle != 0) break;
                await Task.Delay(100, timeout.Token);
            }
            var target = new ControlRef(host.ProcessName, HostTitle, null, "Work menu",
                "ControlType.Button", null, "statusStrip");
            var command = target with { Name = "Browse items", ControlType = "ControlType.MenuItem", ParentName = null };
            if (scenario == "nested") command = command with { ParentName = "Utilities" };
            var plan = new ExecutionPlan(1, DateTimeOffset.Now, "Menu regression",
                [new(1, "click", target, null, null, null, null, null),
                 new(2, "click", command, null, null, null, null, null)]);
            var executor = typeof(ExecutionPlan).Assembly.GetType("DesktopSteps.Executor", throwOnError: true)!;
            var replay = executor.GetMethod("ReplayAsync", BindingFlags.Public | BindingFlags.Static)!;
            var messages = new List<string>();
            Action<string> status = message => { messages.Add(message); Console.WriteLine(message); };
            if (scenario == "capture")
            {
                var window = AutomationElement.FromHandle(host.MainWindowHandle);
                var button = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, target.Name));
                ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                await Task.Delay(1000, timeout.Token);
                var parent = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, "Utilities"));
                ((ExpandCollapsePattern)parent.GetCurrentPattern(ExpandCollapsePattern.Pattern)).Expand();
                await Task.Delay(200, timeout.Token);
                var item = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, command.Name));
                var bounds = item.Current.BoundingRectangle;
                var x = (int)(bounds.Left + bounds.Width / 2);
                var y = (int)(bounds.Top + bounds.Height / 2);
                var assembly = typeof(ControlRef).Assembly;
                var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
                var msaa = assembly.GetType("DesktopSteps.MsaaActions")!;
                var capturedParent = (string?)msaa.GetMethod("MenuParentAtPoint")!.Invoke(null, [x, y]);
                if (capturedParent != "Utilities")
                    throw new InvalidOperationException("Mouse-down capture lost the submenu parent.");
                var popupTarget = (ControlRef?)recorderType.GetMethod("CapturePopupMenuTarget",
                    BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [x, y]);
                if (popupTarget?.Name != command.Name || popupTarget.ParentName != capturedParent)
                    throw new InvalidOperationException("UIA fallback did not capture the popup command and parent.");
                var snapshot = recorderType.GetMethod("CaptureMouseDownTarget",
                    BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [x, y]);
                var snapshotTarget = (ControlRef?)snapshot?.GetType().GetProperty("Target")!.GetValue(snapshot);
                if (snapshotTarget?.Name != command.Name || snapshotTarget.ParentName != capturedParent ||
                    snapshotTarget.Window != HostTitle)
                    throw new InvalidOperationException("Physical-point snapshot did not preserve menu owner identity.");
                var nativeRef = recorderType.GetMethod("NativeControlRef", BindingFlags.NonPublic | BindingFlags.Static)!;
                var frozen = ((ControlRef?)nativeRef.Invoke(null,
                    [(nint)window.Current.NativeWindowHandle, null, command.Name, "ControlType.MenuItem", null])
                    ?? throw new InvalidOperationException("Mouse-down target was not captured."))
                    // The fixture runs the same executable as the test runner;
                    // use an external-app identity to avoid recorder self-exclusion.
                    with { Process = "RecordedMenuHost", ParentName = capturedParent };
                // Close the menu before delivering the queued recorder event.
                SendKeys.SendWait("{ESC}{ESC}");
                using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
                var pointType = recorderType.GetNestedType("Point", BindingFlags.NonPublic)!;
                var point = Activator.CreateInstance(pointType)!;
                pointType.GetField("X")!.SetValue(point, x);
                pointType.GetField("Y")!.SetValue(point, y);
                var infoType = recorderType.GetNestedType("MouseInfo", BindingFlags.NonPublic)!;
                var info = Activator.CreateInstance(infoType)!;
                infoType.GetField("Point")!.SetValue(info, point);
                recorderType.GetMethod("ProcessMouse", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .Invoke(recorder, [0, (nint)0x201, info, null, capturedParent, frozen, null, DateTimeOffset.Now]);
                var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
                if (events.Count != 1 || events[0].Target != frozen)
                    throw new InvalidOperationException("Queued recorder delivery replaced the frozen menu identity.");
                var generated = new ExecutionPlan(1, DateTimeOffset.Now, "Captured menu",
                    [new(1, events[0].Kind, events[0].Target, null, null, null, null, null)]);
                var compactor = assembly.GetType("DesktopSteps.PlanCompactor")!;
                generated = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null, [generated])!;
                var persisted = System.Text.Json.JsonSerializer.Deserialize<ExecutionPlan>(
                    System.Text.Json.JsonSerializer.Serialize(generated))!;
                if (persisted.Steps.Single().Target != frozen)
                    throw new InvalidOperationException("Plan generation lost the captured menu identity.");
                var headerTarget = frozen with { Name = "E", ParentName = "Grid",
                    ControlType = "ControlType.DataItem", ClassName = "XLGridColumnHeader", AutomationId = "" };
                var snapshotType = recorderType.GetNestedType("MouseDownSnapshot", BindingFlags.NonPublic)!;
                var headerSnapshot = Activator.CreateInstance(snapshotType, [button, headerTarget])!;
                recorderType.GetMethod("ProcessMouse", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .Invoke(recorder, [0, (nint)0x201, info, null, null, null, headerSnapshot, DateTimeOffset.Now]);
                recorderType.GetMethod("FinishPendingClick", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .Invoke(recorder, null);
                if (events.Count != 2 || events[1].Target != headerTarget ||
                    events[1].AfterState == "unverified-click")
                    throw new InvalidOperationException("Delayed completion replaced a verified header with a later point/focus target.");
                Console.WriteLine("PASS: closed-menu queued capture preserves command and parent through plan serialization.");
                Console.WriteLine("PASS: verified header identity survives delayed click completion.");
                return;
            }
            if (scenario is "legacy" or "duplicate" or "disabled" or "wrong-window" or "wrong-process")
            {
                var window = AutomationElement.FromHandle(host.MainWindowHandle);
                var button = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, target.Name));
                ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                await Task.Delay(1000, timeout.Token);
                var invokeLegacy = executor.GetMethod("TryInvokeRecordedLegacyMenu",
                    BindingFlags.NonPublic | BindingFlags.Static)!;
                var legacyTarget = scenario switch
                {
                    "wrong-window" => command with { Window = "A different recorded window" },
                    "wrong-process" => command with { Process = "A different recorded process" },
                    _ => command
                };
                var invoked = (bool)invokeLegacy.Invoke(null, [legacyTarget])!;
                if (invoked != (scenario is "legacy" or "wrong-window"))
                    throw new InvalidOperationException($"Legacy menu uniqueness check failed: {scenario}.");
                await Task.Delay(200, timeout.Token);
                host.Refresh();
                if ((host.MainWindowTitle == HostTitle + " - Invoked") !=
                    (scenario is "legacy" or "wrong-window"))
                    throw new InvalidOperationException("Legacy menu invocation did not match its reported result.");
                Console.WriteLine($"PASS: {scenario} menu command.");
                return;
            }
            if (scenario == "missing")
                plan = plan with { Steps = [plan.Steps[0],
                    plan.Steps[1] with { Target = command with { Name = "Not recorded in this menu" } }] };
            var result = (Task<bool>)replay.Invoke(null,
                [plan, status, (Func<string, bool>)(_ => false),
                 (Func<string, bool>)(_ => false), timeout.Token])!;
            try
            {
                await result;
                if (scenario == "missing")
                    throw new InvalidOperationException("Missing menu command was incorrectly accepted.");
            }
            catch (InvalidOperationException ex) when (scenario == "missing" &&
                ex.Message.Contains("was not found through UI Automation"))
            {
                if (messages.Any(message => message.Contains("Opening context menu")))
                    throw new InvalidOperationException("An ordinary dropdown triggered context-menu recovery.");
                var visibleCommand = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ProcessIdProperty, host.Id),
                        new PropertyCondition(AutomationElement.NameProperty, command.Name),
                        new PropertyCondition(AutomationElement.IsOffscreenProperty, false)));
                if (visibleCommand is null)
                    throw new InvalidOperationException("Searching for a missing item dismissed the open menu.");
                Console.WriteLine("PASS: missing command stops safely without closing the dropdown.");
                return;
            }
            host.Refresh();
            if (host.MainWindowTitle != HostTitle + " - Invoked")
                throw new InvalidOperationException("Replay reported success without invoking the menu command.");
            Console.WriteLine("PASS: status-strip dropdown replay invoked its recorded command.");
        }
        finally
        {
            if (!host.HasExited)
            {
                host.CloseMainWindow();
                if (!host.WaitForExit(2000)) host.Kill();
            }
        }
    }
}
