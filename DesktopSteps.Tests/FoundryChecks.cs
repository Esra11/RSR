using System.Net;
using System.Net.Http;
using System.IO;
using System.Reflection;
using System.Text.Json;
using DesktopSteps;

internal static class FoundryChecks
{
    internal static Task RunFilterConfirmationChecksAsync()
    {
        CheckMissingFilterValues();
        return Task.CompletedTask;
    }

    internal static Task RunDuplicateFillChecksAsync()
    {
        const string intent = "fill the entered value down to the adjacent data end";
        var cell = new ControlRef("EXCEL", "Workbook", "L2", "L2", "ControlType.DataItem",
            "XLSpreadsheetCell", "Table");
        var steps = new List<PlanStep>
        {
            new(1, "type", cell, "BL", null, null, null, "Enter value."),
            new(2, "fill-down-to-adjacent-data-end", cell, null, null, "adjacent-column:K", null,
                "Fill value.", intent, null, null, null, intent),
            new(3, "click", cell with { AutomationId = "", Name = "D3", ClassName = "" }, null, null,
                null, null, "Navigate."),
            new(4, "click", cell with { AutomationId = "", Name = "D25", ClassName = "" }, null, null,
                null, null, "Navigate."),
            new(5, "fill-down-to-adjacent-data-end", cell with { AutomationId = "", Name = "D13", ClassName = "" },
                null, null, "adjacent-column:K", null, "Duplicate fill.", intent, null, null, null, intent),
            new(6, "filter-values", cell with { AutomationId = "Dropdown", Name = "Filter applied",
                ControlType = "ControlType.MenuItem", ClassName = "", ParentName = "G1" }, "[]", null,
                "filter:all", null, "Clear filter.")
        };
        var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!
            .GetMethod("Compact", BindingFlags.Public | BindingFlags.Static)!;
        var compacted = (ExecutionPlan)compact.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Duplicate fill", steps)])!;
        if (compacted.Steps.Count != 3 || compacted.Steps.Any(step => step.Target?.Name is "D3" or "D25" or "D13") ||
            compacted.Steps.Count(step => step.Action == "fill-down-to-adjacent-data-end") != 1)
            throw new InvalidOperationException("A duplicate orphan fill sequence was retained.");
        const string rowIntent = "copy functions of previous row from columns A to F and use next Thursday";
        var represented = typeof(ControlRef).Assembly.GetType("DesktopSteps.MainForm")!
            .GetMethod("RecordedIntentIsRepresented", BindingFlags.NonPublic | BindingFlags.Static)!;
        var semanticPlan = new ExecutionPlan(1, DateTimeOffset.Now, "Row date", [
            new PlanStep(1, "duplicate-range-values-and-formulas", cell, "A16:F16", null,
                "destination-range:A17:F17", null, "Duplicate."),
            new PlanStep(2, "update-cell-to-relative-weekday", cell, "A16", null,
                "destination-range:A17:F17", null, "Update.", RelativeWeekday: "Thursday")]);
        if (!(bool)represented.Invoke(null, [semanticPlan, rowIntent])! ||
            (bool)represented.Invoke(null, [semanticPlan, "joseor"])!)
            throw new InvalidOperationException("Semantic intent representation validation regressed.");
        Console.WriteLine("PASS: duplicate orphan fill and its meaningless cell navigation are removed without step-number rules.");
        return Task.CompletedTask;
    }
    internal static Task RunFilteredFillRecoveryAsync()
    {
        CheckWorksheetIntentRecovery();
        var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!;
        var skip = type.GetMethod("IsChartSelectionBeforeLayout", BindingFlags.NonPublic | BindingFlags.Static)!;
        var chartArea = new PlanStep(1, "click", new ControlRef("EXCEL", "Workbook", null, "Chart Area",
            "ControlType.Group", "", null), null, null, null, null, null);
        var layout = new PlanStep(2, "set-chart-legend-layout", chartArea.Target! with { Name = "Chart 7",
            ControlType = "ControlType.Image", ClassName = "ExcelChartObject" }, "{}", null,
            "chart-legend-sheet:Sheet1", null, null);
        if (!(bool)skip.Invoke(null, [chartArea, layout])!)
            throw new InvalidOperationException("A chart-area UIA click was still required before native legend layout.");
        Console.WriteLine("PASS: repeated filtered values recover the shared explicit fill intent without fixed rows.");
        return Task.CompletedTask;
    }
    private static async Task CheckMissingChartHoverElementAsync()
    {
        var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(type)!;
        var core = type.GetMethod("PrimeHoveredChartCore", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var calls = 0;
        Func<System.Windows.Point, System.Windows.Automation.AutomationElement?> missing = _ =>
        {
            Interlocked.Increment(ref calls);
            return null;
        };
        var request = (AutoResetEvent)type.GetField("chartSampleRequested",
            BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!;
        type.GetMethod("StartChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(recorder, [(Action)(() => core.Invoke(recorder, [missing]))]);
        try
        {
            for (var round = 1; round <= 3; round++)
            {
                request.Set();
                for (var attempt = 0; Volatile.Read(ref calls) < round && attempt < 100; attempt++)
                    await Task.Delay(10);
                if (Volatile.Read(ref calls) < round)
                    throw new InvalidOperationException("A missing UIA hit stopped the chart sampling worker.");
            }
            type.GetMethod("StopChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!
                .Invoke(recorder, null);
            var diagnostics = (IEnumerable<string>)type.GetField("captureDiagnostics",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!;
            if (diagnostics.Count(line => line.Contains("pointer UIA hit returned no element")) < 3 ||
                type.GetField("hoveredChart", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .GetValue(recorder) is not null)
                throw new InvalidOperationException("Missing chart hover hits were not diagnosed or produced a guessed snapshot.");
            Console.WriteLine("PASS: null UIA point hits are diagnosed, publish no chart identity and leave the STA sampler alive.");

            var deniedCalls = 0;
            Func<System.Windows.Point, System.Windows.Automation.AutomationElement?> denied = _ =>
            {
                Interlocked.Increment(ref deniedCalls);
                throw new System.ComponentModel.Win32Exception(5, "Access is denied.");
            };
            type.GetField("chartSampleStopping", BindingFlags.NonPublic | BindingFlags.Instance)!
                .SetValue(recorder, false);
            core.Invoke(recorder, [denied]);
            if (deniedCalls != 1)
                throw new InvalidOperationException("The inaccessible-window chart fixture did not execute.");
            diagnostics = (IEnumerable<string>)type.GetField("captureDiagnostics",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!;
            if (!diagnostics.Any(line => line.Contains("Win32Exception")))
                throw new InvalidOperationException("Access denial during chart sampling was not contained and diagnosed.");
            Console.WriteLine("PASS: elevated-window access denial is contained by chart sampling and cannot terminate recording.");
        }
        finally
        {
            type.GetMethod("StopChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!
                .Invoke(recorder, null);
        }
    }

    private static void CheckOrderedCaptureDelivery()
    {
        var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.OrderedCaptureQueue")!;
        var queue = Activator.CreateInstance(type)!;
        var reserve = type.GetMethod("Reserve")!;
        var complete = type.GetMethod("Complete")!;
        long Reserve() => (long)reserve.Invoke(queue, null)!;
        void Complete(long sequence, Action action) => complete.Invoke(queue, [sequence, action]);
        var order = new List<string>();
        var down1 = Reserve();
        var up1 = Reserve();
        var key = Reserve();
        var down2 = Reserve();
        var up2 = Reserve();
        var last = Reserve();
        int? down = null;
        var gestures = new List<int>();
        Complete(up2, () => { order.Add("up2"); gestures.Add(20 - down!.Value); down = null; });
        Complete(down2, () => { order.Add("down2"); down = 20; });
        Complete(key, () => order.Add("key"));
        Complete(up1, () => { order.Add("up1"); gestures.Add(10 - down!.Value); down = null; });
        if (order.Count != 0) throw new InvalidOperationException("Input overtook an incomplete earlier hook capture.");
        Complete(down1, () =>
        {
            order.Add("down1");
            Complete(last, () => order.Add("last"));
            if (order.Count != 1) throw new InvalidOperationException("Reentrant dispatch overtook the current input.");
            down = 10;
        });
        if (!order.SequenceEqual(new[] { "down1", "up1", "key", "down2", "up2", "last" }) ||
            gestures.Count != 2 || gestures.Any(distance => distance != 0))
            throw new InvalidOperationException("Reverse completion corrupted mouse/keyboard order or manufactured a drag.");
        var empty = Reserve();
        var following = Reserve();
        Complete(following, () => order.Add("following"));
        Complete(empty, () => { });
        if (order[^1] != "following") throw new InvalidOperationException("An ignored input left a delivery gap.");
        var stop1 = Reserve();
        var stop2 = Reserve();
        var enqueue = type.GetMethod("Enqueue")!;
        enqueue.Invoke(queue, [stop2, (Action)(() => order.Add("stop2"))]);
        enqueue.Invoke(queue, [stop1, (Action)(() => order.Add("stop1"))]);
        if (order[^1] != "following") throw new InvalidOperationException("Hook completion processed input off the UI dispatcher.");
        type.GetMethod("Drain")!.Invoke(queue, null);
        if (!order.TakeLast(2).SequenceEqual(new[] { "stop1", "stop2" }))
            throw new InvalidOperationException("Stop did not drain completed hooks in order before saving.");
        Console.WriteLine("PASS: reversed/reentrant hook completion preserves mouse/keyboard order and pairs each release with its own click.");
    }

    private static async Task CheckSlowChartSamplingAsync()
    {
        var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(type)!;
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var samples = 0;
        var callerThread = Environment.CurrentManagedThreadId;
        var sampleThread = callerThread;
        var sampleApartment = ApartmentState.Unknown;
        type.GetMethod("StartChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(recorder, [(Action)(() =>
            {
                Interlocked.Increment(ref samples);
                sampleThread = Environment.CurrentManagedThreadId;
                sampleApartment = Thread.CurrentThread.GetApartmentState();
                entered.Set();
                release.Wait(TimeSpan.FromSeconds(5));
            })]);
        try
        {
            for (var attempt = 0; !entered.IsSet && attempt < 100; attempt++) await Task.Delay(10);
            if (!entered.IsSet) throw new InvalidOperationException("Chart sampler did not start.");
            var request = (AutoResetEvent)type.GetField("chartSampleRequested",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(recorder)!;
            for (var index = 0; index < 20; index++) request.Set();
            var ticks = 0;
            using var heartbeat = new System.Windows.Forms.Timer { Interval = 20 };
            heartbeat.Tick += (_, _) => ticks++;
            heartbeat.Start();
            await Task.Delay(2600);
            if (ticks < 20 || samples != 1 || sampleThread == callerThread || sampleApartment != ApartmentState.STA)
                throw new InvalidOperationException("Slow chart sampling blocked the UI or spawned overlapping samples.");
            release.Set();
            type.GetMethod("StopChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(recorder, null);
            if (type.GetField("chartSampleThread", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .GetValue(recorder) is not null)
                throw new InvalidOperationException("Chart sampling thread survived Stop.");
            Console.WriteLine("PASS: a 2.6-second chart sample leaves the UI responsive, runs on one STA worker and stops cleanly.");
        }
        finally
        {
            release.Set();
            type.GetMethod("StopChartSampling", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(recorder, null);
        }
    }

    public static async Task RunAsync()
    {
        CheckOrderedCaptureDelivery();
        await CheckSlowChartSamplingAsync();
        await CheckMissingChartHoverElementAsync();
        var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.Foundry")!;
        var configuration = Path.Combine(Path.GetTempPath(), $"rsr-foundry-{Guid.NewGuid():N}.env");
        try
        {
            await File.WriteAllTextAsync(configuration,
                "PROJECT_ENDPOINT=https://synthetic.invalid/api/projects/test\nAGENT_NAME=test\nAZURE_SUBSCRIPTION_ID=synthetic-subscription\n");
            using var client = new HttpClient(new StubHandler());
            var foundry = Activator.CreateInstance(type, [configuration, client])!;
            var planRecording = type.GetMethod("PlanAsync")!;
            var incompleteChart = new List<RecordedEvent>
            {
                new(DateTimeOffset.Now, "context-click",
                    new ControlRef("EXCEL", "Recorded workbook", "", "", "ControlType.Image",
                        "", "Chart Area", ProcessId: 123),
                    null, null, null, null, null),
                new(DateTimeOffset.Now, "click",
                    new ControlRef("EXCEL", "Select Data Source", "", "OK", "ControlType.Button",
                        "", "Select Data Source", ProcessId: 123),
                    null, null, null, null, null)
            };
            try
            {
                await (Task<ExecutionPlan>)planRecording.Invoke(foundry,
                    [incompleteChart, CancellationToken.None, null])!;
                throw new InvalidOperationException("Foundry accepted chart capture with missing identity/source.");
            }
            catch (InvalidDataException ex) when (ex.Message.Contains("no frozen chart identity")) { }
            static async Task CheckRenamePlanningAnchorsAsync(Type foundry)
            {
                var path = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                    "20261005-235456", "recorded_events.json");
                var original = File.Exists(path) ? File.ReadAllBytes(path) : null;
                var tab = new ControlRef("EXCEL", "Recorded workbook", "SheetTab", "Clean Sheet",
                    "ControlType.TabItem", "", "");
                var pane = tab with { Name = "", AutomationId = "", ControlType = "ControlType.Pane" };
                var noteText = "Rename to the date of next Thursday in the form EMEA Backlog (NextTuesdayDate Month year)";
                var events = original is not null ? JsonSerializer.Deserialize<List<RecordedEvent>>(original)! :
                    new List<RecordedEvent>
                    {
                        new(DateTimeOffset.Now, "context-click", tab, null, null, null, null, null),
                        new(DateTimeOffset.Now, "click", tab with { Name = "Move or Copy...", ControlType = "ControlType.MenuItem" },
                            null, null, null, null, null),
                        new(DateTimeOffset.Now, "click", tab with { Window = "Move or Copy", Name = "Create a copy",
                            ControlType = "ControlType.CheckBox" }, null, null, null, null, null),
                        new(DateTimeOffset.Now, "click", tab with { Window = "Move or Copy", Name = "OK",
                            ControlType = "ControlType.Button" }, null, null, null, null, null),
                        new(DateTimeOffset.Now, "double-click", tab with { Name = "Clean Sheet (2)" },
                            null, null, null, null, null),
                        new(DateTimeOffset.Now, "type", pane, "EMEA Backlog (26 Oct 26)", null, null, null, null),
                        new(DateTimeOffset.Now, "key", pane, null, "Enter", null, null,
                            "sheet-name:EMEA Backlog (26 Oct 26)", noteText)
                    };
                var assembly = typeof(ControlRef).Assembly;
                var grouped = (List<RecordedEvent>)assembly.GetType("DesktopSteps.ActionGrouper")!
                    .GetMethod("Group")!.Invoke(null, [events])!;
                var note = grouped.FindIndex(item => item.Intent is not null) + 1;
                var steps = grouped.Select((item, index) => new PlanStep(index + 1, item.Kind, item.Target,
                    item.Value, item.Key, item.AfterState, item.Screenshot, null, item.Intent)).ToList();
                var plan = (ExecutionPlan)assembly.GetType("DesktopSteps.PlanCompactor")!
                    .GetMethod("Compact")!.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now, "Rename regression", steps)])!;
                var rename = plan.Steps.Single(item => item.Action == "rename-sheet");
                if (rename.Target?.Name != "Clean Sheet (2)" ||
                    rename.Value != "EMEA Backlog ({{next:Thursday:dd MMM yy}})" ||
                    !plan.Steps.Any(item => item.Action == "context-click" && item.Target?.Name == "Clean Sheet") ||
                    !plan.Steps.Any(item => item.Target?.Name == "Create a copy") ||
                    !plan.Steps.Any(item => item.Target is { Window: "Move or Copy", Name: "OK" }))
                    throw new InvalidOperationException("Saved cloning and corrected rename evidence no longer compacts to the working semantic sequence.");
                string Answer(int intentNumber, bool extra) => JsonSerializer.Serialize(new
                {
                    summary = "Recorded clone and rename",
                    steps = grouped.Select((item, index) => new { number = index + 1, action = item.Kind }),
                    intentActions = extra
                        ? new[] { new { intentNumber, targetNumber = 1, action = "click" } }
                        : []
                });
                var validate = foundry.GetMethod("RequestValidatedPlanningChunkAsync",
                    BindingFlags.NonPublic | BindingFlags.Static)!;
                var requests = 0;
                Func<string, Task<string>> corrected = prompt =>
                {
                    requests++;
                    if (requests == 2 && !prompt.Contains("Valid recorded intent numbers"))
                        throw new InvalidOperationException("Correction request omitted the valid anchor numbers.");
                    return Task.FromResult(requests == 1 ? Answer(note + 100, true) : Answer(note, false));
                };
                var answer = await (Task<string>)validate.Invoke(null, [corrected, "Synthetic prompt", grouped, 0, grouped.Count])!;
                if (requests != 2 || !answer.Contains("Recorded clone"))
                    throw new InvalidOperationException("Invalid intent anchor did not trigger exactly one corrected response.");
                Func<string, Task<string>> invalid = _ => Task.FromResult(Answer(note + 100, true));
                try
                {
                    await (Task<string>)validate.Invoke(null, [invalid, "Synthetic prompt", grouped, 0, grouped.Count])!;
                    throw new InvalidOperationException("Repeated invalid intent anchors were accepted.");
                }
                catch (InvalidDataException ex) when (ex.Message.Contains("after one correction attempt")) { }
                if (original is not null && !original.SequenceEqual(File.ReadAllBytes(path)))
                    throw new InvalidOperationException("Rename regression changed the saved raw recording.");
                Console.WriteLine("PASS: saved sheet cloning and rename preserve corrected text and next Thursday; invalid AI anchors retry once, then fail explicitly without changing evidence.");
            }
            await CheckRenamePlanningAnchorsAsync(type);
            incompleteChart.RemoveAt(0);
            try
            {
                await (Task<ExecutionPlan>)planRecording.Invoke(foundry,
                    [incompleteChart, CancellationToken.None, null])!;
                throw new InvalidOperationException("Foundry accepted a source confirmation with no verified outcome.");
            }
            catch (InvalidDataException ex) when (ex.Message.Contains("no verified final source outcome")) { }
            var process = (System.Diagnostics.ProcessStartInfo)type.GetMethod("CliTokenProcess",
                BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(foundry, null)!;
            if (!process.ArgumentList.Contains("synthetic-subscription") &&
                !process.Arguments.Contains("synthetic-subscription", StringComparison.Ordinal))
                throw new InvalidOperationException("Token request did not bind the configured subscription.");
            var ask = type.GetMethod("AskWithTokenAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
            var result = await (Task<string>)ask.Invoke(foundry, ["success", "synthetic-token", CancellationToken.None])!;
            if (result != "Synthetic response") throw new InvalidOperationException("Response text was not extracted.");
            try
            {
                await (Task<string>)ask.Invoke(foundry, ["failure", "synthetic-token", CancellationToken.None])!;
                throw new InvalidOperationException("Foundry HTTP failure was swallowed.");
            }

            catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.NotFound &&
                ex.Message.Contains("WorkspaceNotFound", StringComparison.Ordinal)) { }

            var chunk = type.GetMethod("RequestPlanningChunkAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
            var delayed = await (Task<string>)chunk.Invoke(foundry,
                ["delayed", "synthetic-token", 0, 30, CancellationToken.None, TimeSpan.FromSeconds(2)])!;
            if (delayed != "Synthetic response")
                throw new InvalidOperationException("Delayed successful chunk was not preserved.");
            try
            {
                await (Task<string>)chunk.Invoke(foundry,
                    ["timeout", "synthetic-token", 30, 60, CancellationToken.None, TimeSpan.FromMilliseconds(50)])!;
                throw new InvalidOperationException("Chunk timeout was swallowed.");
            }
            catch (TimeoutException ex) when (ex.Message.Contains("actions 31-60") &&
                ex.Message.Contains("Finish latest") && !ex.Message.Contains("Verify PROJECT_ENDPOINT")) { }
            using var cancelled = new CancellationTokenSource();
            cancelled.CancelAfter(50);
            try
            {
                await (Task<string>)chunk.Invoke(foundry,
                    ["timeout", "synthetic-token", 0, 30, cancelled.Token, TimeSpan.FromSeconds(2)])!;
                throw new InvalidOperationException("User cancellation was swallowed.");
            }
            catch (OperationCanceledException) when (cancelled.IsCancellationRequested) { }
            try
            {
                await (Task<string>)chunk.Invoke(foundry,
                    ["failure", "synthetic-token", 0, 30, CancellationToken.None, TimeSpan.FromSeconds(2)])!;
                throw new InvalidOperationException("Permanent service error was swallowed.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("WorkspaceNotFound") &&
                ex.Message.Contains("Verify PROJECT_ENDPOINT")) { }

            var parse = type.GetMethod("ParseSuggestions", BindingFlags.NonPublic | BindingFlags.Static)!;
            var parsed = ((string? Summary, JsonElement[] Steps))parse.Invoke(null,
                ["{\"summary\":\"Complete\",\"steps\":[{\"number\":1},{\"number\":2},{\"number\":3}]}", 0, 2])!;
            if (parsed.Steps.Length != 2 || parsed.Summary != "Complete")
                throw new InvalidOperationException("Chunk context was not filtered correctly.");
            foreach (var answer in new[]
            {
                "{\"steps\":[]}", "{\"steps\":[{\"number\":1}]}",
                "{\"steps\":[{\"number\":1},{\"number\":1},{\"number\":2}]}", "not JSON"
            })
            {
                try
                {
                    parse.Invoke(null, [answer, 0, 2]);
                    throw new InvalidOperationException("Incomplete or duplicated Foundry output was accepted.");
                }
                catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
            }
            CheckIntentExpansion(type);
            CheckOptionalIntentExpansion(type);
            CheckSortCompaction();
            CheckIntentEditsAndAcknowledgements();
            CheckFilterOutcomes();
            CheckSupersededCellClick();
            CheckLegendDragRecovery();
            CheckWorksheetIntentRecovery();
            CheckRedundantListSelection();
            CheckConfirmedTabReveal();
            CheckMissingFilterValues();
            CheckTitleSortAndQueuedSelection();
            CheckGridEditCorrections();
            Console.WriteLine("PASS: Foundry subscription, response, HTTP error and complete-chunk checks.");
        }
            finally { File.Delete(configuration); }
        }

        private static void CheckGridEditCorrections()
        {
            var assembly = typeof(ControlRef).Assembly;
            var cell = new ControlRef("EXCEL", "Workbook", "J9", "J9", "ControlType.DataItem", "XLSpreadsheetCell", "Grid");
            var pane = cell with { Name = "", AutomationId = "", ControlType = "ControlType.Pane", ClassName = "EXCEL6" };
            var at = DateTimeOffset.Now;
            var events = new List<RecordedEvent>
            {
                new(at, "key", cell, "F", "Shift+F", null, null, null),
                new(at.AddMilliseconds(700), "key", pane, ":", "Shift+OemSemicolon", null, null, null),
                new(at.AddMilliseconds(1300), "key", pane, null, "Back", null, null, null),
                new(at.AddMilliseconds(1700), "key", pane, "L", "Shift+L", null, null, null)
            };
            var group = assembly.GetType("DesktopSteps.ActionGrouper")!.GetMethod("Group")!;
            var grouped = (List<RecordedEvent>)group.Invoke(null, [events])!;
            if (grouped is not [{ Kind: "type", Value: "FL", Target.AutomationId: "J9" }])
                throw new InvalidOperationException("Recorded in-cell corrections became separate pane input.");
            var compact = assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
            var input = new List<PlanStep>
            {
                new(1, "type", cell, "F:", null, null, null, null),
                new(2, "key", pane, null, "Back", null, null, null),
                new(3, "type", pane, "L", null, null, null, null)
            };
            ExecutionPlan Compact(List<PlanStep> steps) => (ExecutionPlan)compact.Invoke(null,
                [new ExecutionPlan(1, at, "Corrections", steps)])!;
            if (Compact(input).Steps is not [{ Action: "type", Value: "FL", Target.AutomationId: "J9" }] ||
                Compact([input[0], input[1] with { Key = "Left" }, input[2]]).Steps.Count != 3 ||
                Compact([input[0], input[1] with { Target = pane with { Window = "Other" } }, input[2]]).Steps.Count != 3)
                throw new InvalidOperationException("Grid edit reconstruction guessed cursor movement or crossed windows.");
            var path = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "Online Excel End to End", "execution_plan.json");
            if (File.Exists(path))
            {
                var saved = JsonSerializer.Deserialize<ExecutionPlan>(File.ReadAllBytes(path))!;
                var recovered = Compact(saved.Steps);
                var fills = recovered.Steps.Select((step, index) => (step, index))
                    .Where(item => item.step.Action == "fill-down-to-adjacent-data-end").ToArray();
                if (fills.Length != 2 ||
                    recovered.Steps[fills[0].index - 1] is not { Value: "FL", TargetStrategy: "first-visible-filtered-row" } ||
                    recovered.Steps[fills[1].index - 1] is not { Value: "BL", TargetStrategy: "first-visible-filtered-row" } ||
                    recovered.Steps.Any(step => step.Target?.ClassName == "EXCEL6" && step.Key == "Back") ||
                    !Compact(recovered.Steps).Steps.SequenceEqual(recovered.Steps))
                    throw new InvalidOperationException("Saved recording still replays intermediate text or fixed-row filtered fills.");
            }
            Console.WriteLine("PASS: cell edit F:/Back/L becomes FL; cursor movement and other windows remain separate; both saved filtered fills use dynamic rows.");
        }

        private static void CheckTitleSortAndQueuedSelection()
        {
            var type = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!;
            var compact = type.GetMethod("Compact")!;
            var title = new ControlRef("FixtureApp", "Query", "TitleBar", "Query", "ControlType.TitleBar", "", "Query");
            var focus = new PlanStep(1, "click", title, null, null, null, null, null);
            var header = title with { Name = "Age", ControlType = "ControlType.HeaderItem", AutomationId = "HeaderItem 3" };
            var sort = focus with { Number = 2, Action = "ensure-state", Target = header, ExpectedState = "sort:descending" };
            var key = focus with { Number = 3, Action = "key", Key = "End", Target = title with {
                Name = "Old row", AutomationId = "ListViewItem-49", ControlType = "ControlType.ListItem" } };
            var cell = focus with { Number = 4, Target = title with { Name = "Old value", AutomationId = "subitem",
                ControlType = "ControlType.Text", ParentName = "Old row" } };
            var context = cell with { Number = 5, Action = "context-selection", ExpectedState = "selection-intent:all" };
            ExecutionPlan Compact(params PlanStep[] steps) => (ExecutionPlan)compact.Invoke(null,
                [new ExecutionPlan(1, DateTimeOffset.Now, "Sorted selection", steps.ToList())])!;
            var result = Compact(focus, sort, key, cell, context);
            if (result.Steps.Count != 3 || result.Steps[0].Action != "ensure-state" ||
                result.Steps[1].Action != "select-all-items" || result.Steps[1].Target != header ||
                result.Steps[2].Target != header ||
                !((ExecutionPlan)compact.Invoke(null, [result])!).Steps.SequenceEqual(result.Steps))
                throw new InvalidOperationException("Title navigation and delayed row selection did not normalize to verified sort/selection.");
            var rule = type.GetMethod("IsTitleFocusBeforeCommand", BindingFlags.NonPublic | BindingFlags.Static)!;
            foreach (var pair in new[]
            {
                (focus with { Intent = "Move the window" }, sort),
                (focus with { ExpectedState = "resized" }, sort),
                (focus, sort with { Target = header with { Window = "Other" } }),
                (focus, sort with { ExpectedState = "sort:unspecified" })
            })
                if ((bool)rule.Invoke(null, [pair.Item1, pair.Item2])!)
                    throw new InvalidOperationException("Meaningful or unrelated title action was removed.");
            if (!Compact(sort, key, cell).Steps.Any(step => step.Action == "key") ||
                !Compact(sort, key, cell, context with { Target = cell.Target! with { Window = "Other" } })
                    .Steps.Any(step => step.Action == "key"))
                throw new InvalidOperationException("Unverified partial selection was replaced with all rows.");
            Console.WriteLine("PASS: same-window title focus before verified sort is omitted; queued End/row-click requires explicit all-row context.");
        }

        private static void CheckMissingFilterValues()
        {
            var assembly = typeof(ControlRef).Assembly;
            var filterType = assembly.GetType("DesktopSteps.ExcelFilterPlan")!;
            var confirm = filterType.GetMethod("Confirm", BindingFlags.NonPublic | BindingFlags.Static)!;
            var compact = assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
            var column = new ControlRef("EXCEL", "Workbook", "Dropdown", "Filter applied",
                "ControlType.MenuItem", "", "G1");
            var opener = new PlanStep(1, "click", column, null, null, null, null, null,
                "User confirmed filter outcome: exclude [\"Alpha\"].");
            var all = opener with { Number = 2, Intent = null, Target = column with {
                Name = "(Select All)", AutomationId = "", ControlType = "ControlType.TreeItem", ParentName = "Manual Filter" } };
            var wrong = all with { Number = 3, Target = column with { Name = "F22",
                AutomationId = "F22", ControlType = "ControlType.DataItem", ParentName = "Grid" } };
            var ok = wrong with { Number = 4, Target = column with { Name = "OK", AutomationId = "",
                ControlType = "ControlType.Button" } };
            ExecutionPlan Compact(List<PlanStep> steps) => (ExecutionPlan)compact.Invoke(null,
                [new ExecutionPlan(1, DateTimeOffset.Now, "Missing checkbox", steps)])!;
            var fixedPlan = Compact([opener, all, wrong, ok]);
            if (fixedPlan.Steps is not [{ Action: "filter-values", ExpectedState: "filter:exclude", Value: "[\"Alpha\"]" }] ||
                Compact([opener with { Intent = null }, all, wrong, ok]).Steps.Count == 1 ||
                Compact([opener, all, wrong]).Steps.Count == 1 ||
                Compact([opener, all, wrong, ok with { Target = ok.Target! with { Window = "Other" } }]).Steps.Count == 1 ||
                Compact([opener, all, wrong with { Target = wrong.Target! with { Window = "Other" } }, ok]).Steps.Count == 1)
                throw new InvalidOperationException("Explicit filter repair discarded unconfirmed or incomplete operations.");
            var empty = opener with { Action = "filter-values", Intent = null, Value = "[]", ExpectedState = "filter:unspecified" };
            foreach (var mode in new[] { "only", "exclude" })
            {
                var selected = (PlanStep)confirm.Invoke(null, [empty, mode, new[] { "Alpha" }])!;
                if (selected.Value != "[\"Alpha\"]" || selected.ExpectedState != "filter:" + mode)
                    throw new InvalidOperationException("Missing named filter values cannot be supplied explicitly.");
                if (mode == "exclude")
                {
                    var cleared = (PlanStep)confirm.Invoke(null, [empty, mode, Array.Empty<string>()])!;
                    if (cleared.ExpectedState != "filter:all" || cleared.Value != "[]")
                        throw new InvalidOperationException("An empty exclusion was not normalized to clear-filter.");
                }
                else try
                {
                    confirm.Invoke(null, [empty, mode, Array.Empty<string>()]);
                    throw new InvalidOperationException("An empty keep-only filter was accepted even though Excel cannot apply it.");
                }
                catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
            }
            var mainForm = assembly.GetType("DesktopSteps.MainForm")!;
            var canConfirm = mainForm.GetMethod("CanConfirmFilterSelection", BindingFlags.NonPublic | BindingFlags.Static)!;
            var effectiveValues = mainForm.GetMethod("EffectiveFilterValues", BindingFlags.NonPublic | BindingFlags.Static)!;
            bool Enabled(int index, params string[] names) => (bool)canConfirm.Invoke(null, [index, names])!;
            if (Enabled(-1) || Enabled(0) || !Enabled(0, "Alpha") || !Enabled(1) || !Enabled(2))
                throw new InvalidOperationException("Filter confirmation button rules regressed.");
            var reused = (string[])effectiveValues.Invoke(null, [Array.Empty<string>(), new[] { "Wesam" }])!;
            var recordedWins = (string[])effectiveValues.Invoke(null, [new[] { "Esraa" }, new[] { "Wesam" }])!;
            if (!reused.SequenceEqual(["Wesam"]) || !recordedWins.SequenceEqual(["Esraa"]))
                throw new InvalidOperationException("A missing filter value did not reuse the prior same-column value safely.");
            Console.WriteLine("PASS: filter confirmation permits clear/no-exception choices without text, while keep-only requires a value.");
        }

        private static void CheckConfirmedTabReveal()
        {
            var method = typeof(ControlRef).Assembly.GetType("DesktopSteps.ActionGrouper")!
                .GetMethod("RecoverMisclassifiedLegendDrags", BindingFlags.NonPublic | BindingFlags.Static)!;
            var at = DateTimeOffset.Now;
            var input = new RecordedEvent(at, "unresolved-input", null, null, null, null, null, null,
                Intent: "User clarification: this unresolved input only revealed worksheet tabs; navigate directly to the subsequently recorded named worksheet instead.",
                Diagnostic: "The mouse input identity could not be frozen before delayed processing. Re-record this section; no target was guessed.");
            var tab = input with { At = at.AddSeconds(1), Kind = "click",
                Target = new ControlRef("EXCEL", "Workbook", "SheetTab", "Named worksheet", "ControlType.TabItem", "", ""),
                Intent = null, Diagnostic = "Control identity verified at mouse-down." };
            List<RecordedEvent> Recover(params RecordedEvent[] events) =>
                (List<RecordedEvent>)method.Invoke(null, [events])!;
            if (!Recover(input, tab).SequenceEqual(new[] { tab }) ||
                Recover(input with { Intent = null }, tab).Count != 2 ||
                Recover(input).Count != 1 ||
                Recover(input, tab with { At = at.AddMinutes(1) }).Count != 2 ||
                Recover(input with { Diagnostic = "Different unresolved input" }, tab).Count != 2)
                throw new InvalidOperationException("Tab navigation recovery omitted unconfirmed or unbounded input.");
            var path = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "Online Excel End to End", "recorded_events.json");
            if (File.Exists(path))
            {
                var before = File.ReadAllBytes(path);
                var events = JsonSerializer.Deserialize<List<RecordedEvent>>(before)!;
                var recovered = Recover(events.ToArray());
                if (events.Count != 198 || recovered.Count != 197 ||
                    recovered.Any(item => item.Kind == "unresolved-input") ||
                    !before.SequenceEqual(File.ReadAllBytes(path)))
                    throw new InvalidOperationException("New full recording recovery lost evidence or retained its confirmed navigation blocker.");
            }
            Console.WriteLine("PASS: user-confirmed tab reveal recovers only beside captured named navigation; raw unresolved evidence remains preserved.");
        }

        private static void CheckRedundantListSelection()
        {
            var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
            var header = new ControlRef("FixtureApp", "Current list", "HeaderItem 3", "Age",
                "ControlType.HeaderItem", "", "Header");
            var sort = new PlanStep(1, "ensure-state", header, null, null, "sort:descending", null, null);
            var key = new PlanStep(2, "key", header with { Name = "Old row", AutomationId = "ListViewItem-49",
                ControlType = "ControlType.ListItem" }, null, "Shift+End", null, "row.jpg", null);
            var select = new PlanStep(3, "select-all-items", header, null, null, null, "all.jpg", null);
            ExecutionPlan Compact(params PlanStep[] steps) => (ExecutionPlan)compact.Invoke(null,
                [new ExecutionPlan(1, DateTimeOffset.Now, "Selection", steps.ToList())])!;
            var result = Compact(sort, key, select);
            if (result.Steps.Count != 2 || result.Steps[1].Action != "select-all-items" ||
                !((ExecutionPlan)compact.Invoke(null, [result])!).Steps.SequenceEqual(result.Steps))
                throw new InvalidOperationException("Redundant row key survived verified selection or compaction is not idempotent.");
            foreach (var unsafePlan in new[]
            {
                Compact(sort, key),
                Compact(sort, key with { Intent = "Select only a subset" }, select),
                Compact(sort, key, select with { Target = header with { Window = "Other list" } }),
                Compact(sort, key with { Key = "Shift+Home" }, select)
            })
                if (!unsafePlan.Steps.Any(step => step.Action == "key"))
                    throw new InvalidOperationException("Selection-key recovery removed an unrelated or partial selection.");
            Console.WriteLine("PASS: redundant post-sort Shift+End is replaced only by the same-view verified all-row operation.");
        }

        private static void CheckWorksheetIntentRecovery()
        {
            var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
            ExecutionPlan Compact(ExecutionPlan plan) => (ExecutionPlan)compact.Invoke(null, [plan])!;
            var cell = new ControlRef("EXCEL", "Synthetic workbook", "J3", "J3",
                "ControlType.DataItem", "XLSpreadsheetCell", "Grid");
            const string fillIntent = "use the first visible filtered row rather than a fixed row number. intent is to copy the cell value to below cells in column J until before the empty adjacent cell in H column";
            var filter = new PlanStep(1, "filter-values", cell with { Name = "Filter", AutomationId = "Dropdown",
                ControlType = "ControlType.MenuItem", ParentName = "B1" }, "[\"Alpha\"]", null, "filter:only", null, null);
            var entry = new PlanStep(2, "type", cell, "tag", null, null, null, null, fillIntent);
            var gesture = new PlanStep(3, "manual-click", cell with { Name = "D3", AutomationId = "D3" },
                null, null, "adjacent-column:H", null, null, fillIntent);
            var recovered = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Fill", [filter, entry, gesture]));
            if (recovered.Steps[1].TargetStrategy != "first-visible-filtered-row" ||
                recovered.Steps[2].Action != "fill-down-to-adjacent-data-end")
                throw new InvalidOperationException("Explicit filtered fill remained an unresolved gesture.");
            var conflict = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Conflict",
                [filter, entry, gesture with { ExpectedState = "adjacent-column:X" }]));
            if (conflict.Steps.Last().Action != "manual-click")
                throw new InvalidOperationException("Conflicting adjacent-column evidence was ignored.");
            var excluded = filter with { ExpectedState = "filter:exclude" };
            var secondFilter = filter with { Number = 3, ExpectedState = "filter:only" };
            var repeated = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Repeated filtered fill",
                [excluded, entry with { Number = 2, Value = "BL", Intent = null }, secondFilter,
                 entry with { Number = 4, Value = "FL" }, gesture with { Number = 5 }]));
            var repeatedFills = repeated.Steps.Select((step, index) => (step, index))
                .Where(item => item.step.Action == "fill-down-to-adjacent-data-end").ToArray();
            if (repeatedFills.Length != 2 ||
                repeated.Steps[repeatedFills[0].index - 1] is not { Value: "BL", TargetStrategy: "first-visible-filtered-row" } ||
                repeated.Steps[repeatedFills[1].index - 1] is not { Value: "FL", TargetStrategy: "first-visible-filtered-row" } ||
                !Compact(repeated).Steps.SequenceEqual(repeated.Steps))
                throw new InvalidOperationException("A repeated filtered fill lost the first value's shared explicit fill rule.");
            var formula = gesture with { Target = cell, ExpectedState = null,
                Intent = "intent is to copy the function in column J till just before the empty cell in of column H" };
            var extended = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Formula", [formula]));
            if (extended.Steps.Single().TargetStrategy != "last-populated-formula" ||
                extended.Steps.Single().ExpectedState != "adjacent-column:H")
                throw new InvalidOperationException("Dynamic formula intent reused the recorded drag row.");
            var note = entry with { Action = "scroll", Intent =
                "intent is to copy the function from the J3 cell in column J to the last cell that has an adjacent existing value in column H" };
            var fixedFormula = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Explicit source",
                [note, formula with { Intent = null }]));
            if (fixedFormula.Steps.Last().TargetStrategy != "formula-source:J3")
                throw new InvalidOperationException("Explicit formula source was replaced by a guessed source row.");
            var unrelated = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Different workbook",
                [note with { Target = cell with { Window = "Other workbook" } }, formula with { Intent = null }]));
            if (unrelated.Steps.Last().Action != "manual-click")
                throw new InvalidOperationException("Formula intent crossed workbook boundaries.");
            var ordinary = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Ordinary formula clicks",
                [note, formula with { Action = "click", Intent = null },
                 formula with { Action = "click", Intent = note.Intent }]));
            if (ordinary.Steps.Count(step => step.Action == "extend-formula-to-adjacent-data-end") != 1 ||
                ordinary.Steps[1].TargetStrategy != "formula-source:J3" ||
                !Compact(ordinary).Steps.SequenceEqual(ordinary.Steps))
                throw new InvalidOperationException("Ordinary formula clicks lost explicit source or emitted duplicate fills.");
            var pasteNote = formula with { Action = "click", Target = cell with { Name = "Values",
                ParentName = "Paste Values", ControlType = "ControlType.ListItem" } };
            var fromPaste = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Following formula",
                [pasteNote, formula with { Action = "click", Intent = null },
                 formula with { Action = "click", Intent = null }]));
            if (fromPaste.Steps.Count(step => step.Action == "extend-formula-to-adjacent-data-end") != 1)
                throw new InvalidOperationException("Following formula intent was lost at Paste Values.");
            foreach (var barrier in new[]
            {
                pasteNote with { Target = cell with { ControlType = "ControlType.TabItem" } },
                pasteNote with { WhenUser = "Different user" },
                pasteNote with { Target = pasteNote.Target! with { Window = "Other workbook" } }
            })
                if (Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Scope barrier",
                    [barrier, formula with { Action = "click", Intent = null }])).Steps
                    .Any(step => step.Action == "extend-formula-to-adjacent-data-end"))
                    throw new InvalidOperationException("Formula intent crossed a scope barrier.");

            var rowCell = cell with { Window = "Evolution", Name = "A16", AutomationId = "A16" };
            var rowEnd = rowCell with { Name = "G16", AutomationId = "G16" };
            var nextRow = rowCell with { Name = "A17", AutomationId = "A17" };
            var find = new PlanStep(4, "key", nextRow, null, "Control+F", "field-value:Oct 06th", null, null);
            var replaceAll = new PlanStep(5, "replace-all", rowCell with { Window = "Find and Replace",
                Name = "Replace All", ControlType = "ControlType.Button" }, "06", null, "08", null, null);
            var closeFind = new PlanStep(6, "click", rowCell with { Window = "Find and Replace",
                Name = "Close", ControlType = "ControlType.Button" }, null, null, null, null, null);
            var renamed = new PlanStep(0, "rename-sheet", rowCell with { Name = "Clean Sheet (2)" },
                "EMEA Backlog ({date:dd MMM yy})", null, null, null, null, RelativeWeekday: "Thursday");
            var duplicated = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Duplicate previous row",
                [renamed, formula with { Number = 1, Action = "click", Target = rowCell, Intent = null },
                 formula with { Number = 2, Action = "click", Target = rowEnd, Intent = null },
                 formula with { Number = 3, Action = "click", Target = nextRow, Intent = null },
                 find, replaceAll, closeFind]));
            if (!duplicated.Steps.Any(step => step is { Action: "duplicate-range-values-and-formulas",
                    Value: "A16:G16", ExpectedState: "destination-range:A17:G17" }) ||
                !duplicated.Steps.Any(step => step is { Action: "update-cell-to-relative-weekday",
                    Value: "A16", ExpectedState: "destination-range:A17:G17", RelativeWeekday: "Thursday" }) ||
                duplicated.Steps.Any(step => step.Action is "key" or "replace-all" &&
                    (step.Key is "Control+F" or "Ctrl+F" || step.Value == "06")) ||
                !Compact(duplicated).Steps.SequenceEqual(duplicated.Steps))
                throw new InvalidOperationException("Recorded source-row boundary and next-row destination did not compact into a verified dynamic duplicate.");
            const string copiedRowIntent = "intent is to copy functions of previous row from columns A to F, and then replace the date in the new pasted cells to next Thursday";
            var unidentifiedEndpoint = formula with { Number = 2, Action = "click", Intent = null,
                Target = rowCell with { Name = "", AutomationId = "automation id",
                    ControlType = "ControlType.Custom" } };
            var inferredDuplicate = Compact(new ExecutionPlan(1, DateTimeOffset.Now,
                "Duplicate previous row with unidentified range endpoint",
                [renamed, formula with { Number = 1, Action = "click", Target = rowCell, Intent = null },
                 unidentifiedEndpoint,
                 formula with { Number = 3, Action = "click", Target = nextRow, Intent = null },
                 find,
                 replaceAll with { Intent = copiedRowIntent, OriginIntent = copiedRowIntent },
                 closeFind]));
            if (!inferredDuplicate.Steps.Any(step => step is { Action: "duplicate-range-values-and-formulas",
                    Value: "A16:F16", ExpectedState: "destination-range:A17:F17" }) ||
                !inferredDuplicate.Steps.Any(step => step is { Action: "update-cell-to-relative-weekday",
                    Value: "A16", ExpectedState: "destination-range:A17:F17" }) ||
                inferredDuplicate.Steps.Any(step => step.Target?.ControlType == "ControlType.Custom") ||
                !Compact(inferredDuplicate).Steps.SequenceEqual(inferredDuplicate.Steps))
                throw new InvalidOperationException("Explicit copied-column intent did not recover an unidentified Excel range endpoint.");
            var endToEndPath = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "End-to-End-3", "execution_plan.json");
            if (File.Exists(endToEndPath))
            {
                var endToEnd = Compact(JsonSerializer.Deserialize<ExecutionPlan>(File.ReadAllBytes(endToEndPath))!);
                var duplicateIndex = endToEnd.Steps.FindIndex(step => step.Action == "duplicate-range-values-and-formulas");
                if (duplicateIndex < 0 || duplicateIndex + 1 >= endToEnd.Steps.Count ||
                    endToEnd.Steps[duplicateIndex] is not { Value: "A16:G16", ExpectedState: "destination-range:A17:G17" } ||
                    endToEnd.Steps[duplicateIndex + 1] is not { Action: "update-cell-to-relative-weekday",
                        Value: "A16", ExpectedState: "destination-range:A17:G17", RelativeWeekday: "Thursday" } ||
                    endToEnd.Steps.Skip(duplicateIndex + 2).Any(step =>
                        step.Target?.Name == "This cell is inconsistent with the column formula.") ||
                    !Compact(endToEnd).Steps.SequenceEqual(endToEnd.Steps))
                    throw new InvalidOperationException("End-to-End-3 did not recover its recorded row duplication before date replacement.");
            }
            var newRecordingPath = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "20261006-201253", "execution_plan.json");
            if (File.Exists(newRecordingPath))
            {
                var newRecording = Compact(JsonSerializer.Deserialize<ExecutionPlan>(File.ReadAllBytes(newRecordingPath))!);
                if (newRecording.Steps.Any(step => step is { Action: "manual-click",
                        Target: { Process: "EXCEL", Name: "G" } }) ||
                    !newRecording.Steps.Any(step => step is { Action: "filter-values",
                        Target.ParentName: "G1", Value: "[\"Wesam (walziadat)\"]" }) ||
                    !Compact(newRecording).Steps.SequenceEqual(newRecording.Steps))
                    throw new InvalidOperationException("The new recording did not recover its bare column-G filter opener generically.");
            }
            var unidentifiedRangeRecordingPath = Path.Combine(Environment.CurrentDirectory, "bin", "RSR",
                "Recordings", "20261007-085438", "execution_plan.json");
            if (File.Exists(unidentifiedRangeRecordingPath))
            {
                var unidentifiedRangeRecording = Compact(JsonSerializer.Deserialize<ExecutionPlan>(
                    File.ReadAllBytes(unidentifiedRangeRecordingPath))!);
                if (!unidentifiedRangeRecording.Steps.Any(step => step is
                        { Action: "duplicate-range-values-and-formulas", Value: "A16:F16",
                          ExpectedState: "destination-range:A17:F17" }) ||
                    !unidentifiedRangeRecording.Steps.Any(step => step is
                        { Action: "update-cell-to-relative-weekday", Value: "A16",
                          ExpectedState: "destination-range:A17:F17", RelativeWeekday: "Thursday" }) ||
                    unidentifiedRangeRecording.Steps.Any(step => step.Target?.ControlType == "ControlType.Custom") ||
                    !Compact(unidentifiedRangeRecording).Steps.SequenceEqual(unidentifiedRangeRecording.Steps))
                    throw new InvalidOperationException("The preserved recording did not compact its unidentified range endpoint into the explicit A:F copy intent.");
            }
            var bareFilter = formula with { Action = "manual-click", Intent = null,
                ExpectedState = "unverified-click", Target = cell with { Name = "C", AutomationId = "",
                    ControlType = "ControlType.DataItem", ParentName = "Grid" } };
            var filterItem = formula with { Action = "click", Intent = null, Target = cell with {
                Name = "(Select All)", AutomationId = "", ControlType = "ControlType.TreeItem",
                ParentName = "Manual Filter" }, ExpectedState = "filter:unchecked" };
            var genericFilter = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Generic bare-column filter",
                [bareFilter, filterItem,
                 filterItem with { Target = filterItem.Target! with { Name = "Alpha" },
                     ExpectedState = "filter:checked" },
                 formula with { Action = "click", Intent = null, Target = cell with { Name = "OK",
                     AutomationId = "", ControlType = "ControlType.Button" }, ExpectedState = null }]));
            if (genericFilter.Steps is not [{ Action: "filter-values", Target.ParentName: "C1",
                    ExpectedState: "filter:only", Value: "[\"Alpha\"]" }] ||
                !Compact(genericFilter).Steps.SequenceEqual(genericFilter.Steps))
                throw new InvalidOperationException("Bare filter recovery was tied to a specific recorded column or value.");

            var destinationMethod = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
                .GetMethod("FindPasteDestination", BindingFlags.NonPublic | BindingFlags.Static)!;
            var destination = formula with { Action = "click", Intent = null };
            var file = destination with { Target = cell with { Name = "File Tab",
                AutomationId = "FileTabButton", ClassName = "NetUIRibbonTab",
                ControlType = "ControlType.Button" } };
            var back = file with { Target = file.Target! with { Name = "Back",
                ClassName = "NetUISimpleButton", ControlType = "ControlType.ListItem", ParentName = "File" } };
            var home = file with { Target = file.Target! with { Name = "Home", AutomationId = "TabHome",
                ControlType = "ControlType.TabItem" } };
            ControlRef? FindDestination(List<PlanStep> steps) =>
                (ControlRef?)destinationMethod.Invoke(null, [steps, steps.Count - 1]);
            if (FindDestination([destination, file, back, home, pasteNote]) != cell)
                throw new InvalidOperationException("Paired File/Back navigation lost the paste destination.");
            foreach (var invalid in new List<PlanStep>[]
            {
                [destination, file, home, pasteNote],
                [destination, back, home, pasteNote],
                [destination, home with { Target = home.Target! with { Window = "Other workbook" } }, pasteNote],
                [destination, file with { Target = cell with { ControlType = "ControlType.TabItem" } }, pasteNote],
                [destination, home with { Action = "key", Key = "Delete" }, pasteNote]
            })
                if (FindDestination(invalid) is not null)
                    throw new InvalidOperationException("Paste destination lookup crossed an unsafe action.");
            var sourceCell = cell with { Name = "J1", AutomationId = "J1", ProcessId = 1234 };
            var command = destination with { Target = sourceCell with { Name = "Data command",
                ControlType = "ControlType.Button", ClassName = "NetUIRibbonButton" } };
            var resultNote = command with { Target = sourceCell with { Window = "Microsoft Excel",
                Name = "OK", ControlType = "ControlType.Button", ClassName = "NetUIButton" },
                Intent = "intent is to copy J and K cells that contain a value" };
            var rectangleCopy = destination with { Action = "key", Key = "Control+C", Target = sourceCell,
                ExpectedState = "excel-selection-range:J1:K10" };
            ExecutionPlan RectanglePlan(PlanStep noteStep) => new(1, DateTimeOffset.Now, "Result-note copy",
                [command, noteStep, destination with { Target = sourceCell }, rectangleCopy]);
            if (Compact(RectanglePlan(resultNote)).Steps.Last() is not
                { Action: "copy-populated-columns", Value: "J1:K1", ExpectedState: null })
                throw new InvalidOperationException("Verified result note did not recover dynamic two-column copy.");
            foreach (var invalidNote in new[]
            {
                resultNote with { Target = resultNote.Target! with { ProcessId = 5678 } },
                resultNote with { Intent = "intent is to copy J and L cells that contain a value" },
                resultNote with { WhenUser = "Different user" },
                resultNote with { Target = resultNote.Target! with { Window = "Other workbook" } }
            })
                if (Compact(RectanglePlan(invalidNote)).Steps.Last().Action == "copy-populated-columns")
                    throw new InvalidOperationException("Two-column copy note crossed conflicting scope or columns.");

            var onlinePath = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "Online Excel End to End", "execution_plan.json");
            if (File.Exists(onlinePath))
            {
                var online = Compact(JsonSerializer.Deserialize<ExecutionPlan>(File.ReadAllBytes(onlinePath))!);
                if (online.Steps.Count(step => step.Action == "extend-formula-to-adjacent-data-end") != 2 ||
                    !online.Steps.Any(step => step.TargetStrategy == "formula-source:C1" &&
                        step.ExpectedState == "adjacent-column:A") ||
                    !online.Steps.Any(step => step.Action == "extend-formula-to-adjacent-data-end" &&
                        step.Value == "F" && step.ExpectedState == "adjacent-column:E") ||
                    !online.Steps.Any(step => step.Action == "copy-populated-columns" && step.Value == "E1:F1") ||
                    !Compact(online).Steps.SequenceEqual(online.Steps))
                    throw new InvalidOperationException("Online recording lost formula fills, dynamic rectangle or idempotence: " +
                        string.Join("; ", online.Steps.Where(step => step.Action is
                            "extend-formula-to-adjacent-data-end" or "copy-populated-columns").Select(step =>
                            $"{step.Number} {step.Action} {step.Value} {step.TargetStrategy} {step.ExpectedState}")) +
                        "; idempotent=" + Compact(online).Steps.SequenceEqual(online.Steps));
                var finalPaste = online.Steps.FindLastIndex(step => step.Target?.AutomationId == "PasteMenu_Dropdown");
                if (((ControlRef?)destinationMethod.Invoke(null, [online.Steps, finalPaste]))?.Name != "O7")
                    throw new InvalidOperationException("Online recording lost its final O7 paste destination.");
            }
            const string chartIntent = "replace the chart data source with the last pasted two-column range";
            var chartContext = destination with { Action = "context-click", Target = cell with {
                Name = "Captured chart", ControlType = "ControlType.Image", ClassName = "ExcelChartObject" } };
            var chartCommand = destination with { Target = cell with { Name = "Select Data...",
                ControlType = "ControlType.MenuItem" }, Intent = chartIntent };
            var chartOk = destination with { Target = cell with { Window = "Select Data Source", Name = "OK" } };
            var chartPlan = Compact(new ExecutionPlan(1, DateTimeOffset.Now, "Captured source replacement",
                [chartContext, chartCommand, destination, chartOk]));
            if (chartPlan.Steps.Single() is not { Action: "set-chart-source-from-last-paste", Value: "Captured chart" } ||
                !Compact(chartPlan).Steps.SequenceEqual(chartPlan.Steps))
                throw new InvalidOperationException("Captured chart identity did not consolidate source replacement.");
            var group = typeof(ControlRef).Assembly.GetType("DesktopSteps.ActionGrouper")!.GetMethod("Group")!;
            var capturedChart = chartContext.Target! with { ProcessId = 123 };
            var dialogTarget = capturedChart with { Window = "Select Data Source", Name = "Chart data range",
                ClassName = "NetUIRefedit", ControlType = "ControlType.Edit" };
            var at = DateTimeOffset.Now;
            var sourceEvents = new List<RecordedEvent>
            {
                new(at, "context-click", capturedChart, null, null, null, null, null),
                new(at, "click", capturedChart with { Name = "Select Data...", ControlType = "ControlType.MenuItem" },
                    null, null, null, null, null),
                new(at, "unresolved-input", dialogTarget with { Name = null, ControlType = "ControlType.Window" },
                    null, null, null, null, null, Diagnostic:
                    "The keyboard input identity could not be frozen before delayed processing. Re-record this section; no target was guessed."),
                new(at, "click", dialogTarget with { Name = "OK", ControlType = "ControlType.Button" },
                    null, null, null, null, null),
                new(at, "set-chart-source-range", capturedChart, """{"Sheet":"Sheet1","Range":"A1:B20"}""",
                    null, null, null, "chart-source-verified")
            };
            var sourceGrouped = (List<RecordedEvent>)group.Invoke(null, [sourceEvents])!;
            if (sourceGrouped.Count != 1 || sourceGrouped[0] != sourceEvents[^1] ||
                sourceEvents.Count != 5 || sourceEvents[2].Kind != "unresolved-input")
                throw new InvalidOperationException("Verified chart outcome did not safely consolidate its transient edit evidence.");
            var withSelection = sourceEvents.ToList();
            withSelection.Insert(0, sourceEvents[0] with { Kind = "click" });
            if (((List<RecordedEvent>)group.Invoke(null, [withSelection])!).Count != 1)
                throw new InvalidOperationException("The verified chart edit retained its redundant same-chart selection.");
            var withRetry = withSelection.ToList();
            withRetry.Insert(1, sourceEvents[0]);
            if (((List<RecordedEvent>)group.Invoke(null, [withRetry])!).Count != 1)
                throw new InvalidOperationException("The verified chart edit retained its same-chart context-menu retry.");
            withRetry[1] = withRetry[1] with { Target = capturedChart with { Name = "Other chart" } };
            if (((List<RecordedEvent>)group.Invoke(null, [withRetry])!)
                .All(item => item.Kind != "context-click" || item.Target?.Name != "Other chart"))
                throw new InvalidOperationException("Chart source consolidation removed another chart's context-click.");
            withSelection[0] = withSelection[0] with { Target = capturedChart with { Name = "Other chart" } };
            if (((List<RecordedEvent>)group.Invoke(null, [withSelection])!).Count != 2)
                throw new InvalidOperationException("Chart consolidation removed another chart's selection.");
            var dialogInputs = sourceEvents.ToList();
            dialogInputs[2] = dialogInputs[2] with { Kind = "chart-source-edit-input" };
            if (((List<RecordedEvent>)group.Invoke(null, [dialogInputs])!).Count != 1 ||
                !((List<RecordedEvent>)group.Invoke(null, [dialogInputs.Take(4).ToList()])!)
                    .Any(item => item.Kind == "unresolved-input"))
                throw new InvalidOperationException("Transient chart dialog input must require a verified final source outcome.");
            var lateOutcome = sourceEvents.ToList();
            lateOutcome[1] = lateOutcome[1] with { Target = capturedChart with { ClassName = "" } };
            var laterDrag = new RecordedEvent(at, "unresolved-input", capturedChart,
                null, null, null, null, null, Diagnostic: "No supported semantic resize outcome was captured for this drag.");
            lateOutcome.Insert(4, laterDrag);
            var lateGrouped = (List<RecordedEvent>)group.Invoke(null, [lateOutcome])!;
            if (lateGrouped.Count != 2 || lateGrouped[0].Kind != "set-chart-source-range" ||
                lateGrouped[1] != laterDrag)
                throw new InvalidOperationException("Late chart outcome swallowed a later drag or required an unavailable menu label.");
            var failedChartRecording = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "20261005-144829", "recorded_events.json");
            if (File.Exists(failedChartRecording))
            {
                var raw = File.ReadAllBytes(failedChartRecording);
                var captured = JsonSerializer.Deserialize<List<RecordedEvent>>(raw)!;
                var chartRecovered = (List<RecordedEvent>)group.Invoke(null, [captured])!;
                if (chartRecovered.Any(item => item.Kind == "context-click" && item.Target?.Name == "Chart 7") ||
                    chartRecovered.Any(item => item.Diagnostic?.StartsWith("Chart source dialog input has no verified") == true) ||
                    chartRecovered.Count(item => item.Kind == "unresolved-input") != 1 ||
                    !chartRecovered.Any(item => item.Kind == "scroll") ||
                    !chartRecovered.Any(item => item.Kind == "set-chart-source-range") ||
                    !raw.SequenceEqual(File.ReadAllBytes(failedChartRecording)))
                    throw new InvalidOperationException("Saved chart source recovery lost a later operation or changed raw evidence.");
            }
            foreach (var barrier in new[] {
                sourceEvents[2] with { Target = null },
                sourceEvents[2] with { Target = dialogTarget with { ProcessId = 456 } },
                sourceEvents[2] with { Target = dialogTarget with { Window = "Other workbook" } },
                sourceEvents[2] with { Kind = "key", Key = "Control+S", Target = capturedChart },
                sourceEvents[2] with { Diagnostic = "Different unresolved input" } })
            {
                var guarded = sourceEvents.ToList();
                guarded[2] = barrier;
                if (((List<RecordedEvent>)group.Invoke(null, [guarded])!).Count == 1)
                    throw new InvalidOperationException("Chart outcome swallowed unrelated or unowned input.");
            }
            var native = typeof(ControlRef).Assembly.GetType("DesktopSteps.ExcelNativeSheet")!;
            var sourceType = native.GetNestedType("ChartSource", BindingFlags.NonPublic)!;
            var validateSource = native.GetMethod("ValidateChartSource", BindingFlags.Static | BindingFlags.NonPublic)!;
            foreach (var range in new[] { "A1:C20", "A20:B1", "XFD1:XFE2", "A1:B1048577", "A0:B2" })
            {
                try
                {
                    validateSource.Invoke(null, [Activator.CreateInstance(sourceType, ["Sheet1", range])]);
                    throw new InvalidOperationException("Invalid chart range was accepted: " + range);
                }
                catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
            }
            Console.WriteLine("PASS: chart outcomes preserve raw evidence, reject scope barriers and validate ranges before mutation.");
            var missingSourcePath = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "20261005-162233", "recorded_events.json");
            if (File.Exists(missingSourcePath))
            {
                var raw = File.ReadAllBytes(missingSourcePath);
                var recording = JsonSerializer.Deserialize<List<RecordedEvent>>(raw)!;
                var incomplete = (List<RecordedEvent>)group.Invoke(null, [recording])!;
                if (!incomplete.Any(item => item.Kind == "unresolved-input" &&
                        item.Diagnostic?.Contains("no frozen chart identity") == true) ||
                    !incomplete.Any(item => item.Kind == "unresolved-input" &&
                        item.Diagnostic?.Contains("no verified final source outcome") == true) ||
                    incomplete.Single(item => item.Kind == "set-chart-legend-layout") !=
                        recording.Single(item => item.Kind == "set-chart-legend-layout") ||
                    !raw.SequenceEqual(File.ReadAllBytes(missingSourcePath)))
                    throw new InvalidOperationException("Incomplete chart-source capture was accepted, altered raw evidence or lost its valid legend outcome.");
            }
            var latestChartPlan = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings",
                "20261005-174524", "execution_plan.json");
            if (File.Exists(latestChartPlan))
            {
                var raw = File.ReadAllBytes(latestChartPlan);
                var saved = JsonSerializer.Deserialize<ExecutionPlan>(raw)!;
                var selection = saved.Steps.Single(step => step.Action == "click");
                var layout = saved.Steps.Single(step => step.Action == "set-chart-legend-layout");
                var predicate = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!
                    .GetMethod("IsChartSelectionBeforeLayout", BindingFlags.NonPublic | BindingFlags.Static)!;
                bool Redundant(PlanStep click, PlanStep next) => (bool)predicate.Invoke(null, [click, next])!;
                if (!Redundant(selection, layout) || Compact(saved).Steps.Count != 3 ||
                    Compact(saved).Steps.Any(step => step.Action == "click") ||
                    !Compact(Compact(saved)).Steps.SequenceEqual(Compact(saved).Steps) ||
                    !raw.SequenceEqual(File.ReadAllBytes(latestChartPlan)))
                    throw new InvalidOperationException("Verified layout retained redundant selection or changed saved evidence.");
                foreach (var incompatible in new[]
                {
                    layout with { Target = layout.Target! with { Name = "Other chart" } },
                    layout with { Target = layout.Target! with { Window = "Other workbook" } },
                    layout with { Target = layout.Target! with { ProcessId = 999 } },
                    layout with { WhenUser = "Different user" },
                    layout with { ExpectedState = "chart-legend-sheet:" },
                    layout with { Action = "click" }
                })
                    if (Redundant(selection, incompatible))
                        throw new InvalidOperationException("Layout selection suppression crossed an identity or action barrier.");
                if (Redundant(selection with { Intent = "Select for another purpose" }, layout) ||
                    Redundant(selection with { ExpectedState = "Other state" }, layout))
                    throw new InvalidOperationException("Layout suppression discarded explicit selection intent/state.");
            }
            var readChartBounds = typeof(ControlRef).Assembly.GetType("DesktopSteps.Recorder")!
                .GetMethod("ReadChartBounds", BindingFlags.NonPublic | BindingFlags.Static)!;
            var validBounds = new System.Windows.Rect(10, 20, 100, 50);
            if ((System.Windows.Rect)readChartBounds.Invoke(null,
                    [(Func<System.Windows.Rect>)(() => validBounds)])! != validBounds)
                throw new InvalidOperationException("Valid chart bounds were discarded.");
            foreach (var read in new Func<System.Windows.Rect>[] {
                () => throw new ArgumentException("Width and Height must be non-negative."),
                () => System.Windows.Rect.Empty,
                () => new System.Windows.Rect(0, 0, 0, 1),
                () => new System.Windows.Rect(0, 0, double.PositiveInfinity, 1) })
                if (!((System.Windows.Rect)readChartBounds.Invoke(null, [read])!).IsEmpty)
                    throw new InvalidOperationException("Invalid provider bounds were accepted.");
            Console.WriteLine("PASS: negative-size provider exceptions do not escape chart sampling; invalid bounds are never used.");

            var path = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings", "full", "execution_plan.json");
            if (File.Exists(path))
            {
                var before = File.ReadAllBytes(path);
                var saved = JsonSerializer.Deserialize<ExecutionPlan>(before)!;
                var clarified = saved with { Steps = saved.Steps.Select(step =>
                    step.Target?.Name == "Select Data..." ? step with { Intent =
                        (step.Intent ?? "") + " User clarification: replace the chart data source with the last pasted two-column range; target chart \"Chart 7\"." } : step).ToList() };
                var full = Compact(clarified);
                if (full.Steps.Any(step => step.Action == "manual-click") ||
                    full.Steps.Count(step => step.Action == "fill-down-to-adjacent-data-end") != 2 ||
                    !full.Steps.Any(step => step.Action == "extend-formula-to-adjacent-data-end" &&
                        step.TargetStrategy == "formula-source:C1") ||
                    !full.Steps.Any(step => step.Action == "extend-formula-to-adjacent-data-end" &&
                        step.Value == "F" && step.TargetStrategy == "last-populated-formula") ||
                    !full.Steps.Any(step => step.Action == "copy-column-until-empty" && step.Value == "K2") ||
                    !full.Steps.Any(step => step.Action == "copy-column-until-empty" && step.Value == "C1" &&
                        step.ExpectedState == "copy-stop-values:[\"#VALUE!\"]") ||
                    !full.Steps.Any(step => step.Action == "set-chart-source-from-last-paste" && step.Value == "Chart 7") ||
                    !full.Steps.Any(step => step.Action == "set-chart-legend-layout") ||
                    !Compact(full).Steps.SequenceEqual(full.Steps) ||
                    !File.ReadAllBytes(path).SequenceEqual(before))
                    throw new InvalidOperationException("Full recording intent recovery lost an operation, boundary, layout or idempotence.");
                var form = typeof(ControlRef).Assembly.GetType("DesktopSteps.MainForm")!;
                var warnings = (List<string>)form.GetMethod("ValidatePlanTransitions",
                    BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [full])!;
                var blocking = form.GetMethod("IsBlockingPlanWarning", BindingFlags.NonPublic | BindingFlags.Static)!;
                if (warnings.Any(warning => (bool)blocking.Invoke(null, [warning])!))
                    throw new InvalidOperationException("Recovered full recording still fails replay preflight: " + string.Join(" ", warnings));
            }
            Console.WriteLine("PASS: worksheet gesture recovery preserves explicit sources, dynamic boundaries and confirmed chart identity.");
        }

        private static void CheckLegendDragRecovery()
        {
            var assembly = typeof(ControlRef).Assembly;
            var method = assembly.GetType("DesktopSteps.ActionGrouper")!
                .GetMethod("RecoverMisclassifiedLegendDrags", BindingFlags.NonPublic | BindingFlags.Static)!;
            var target = new ControlRef("EXCEL", "Workbook", "C1", "C1",
                "ControlType.DataItem", "XLSpreadsheetCell", "Grid");
            var diagnostic = "No supported semantic resize outcome was captured for this drag.";
            var source = new List<RecordedEvent>
            {
                new(DateTimeOffset.Now, "unresolved-input", target with { Name = "Page left",
                    ControlType = "ControlType.Button", ClassName = "NetUIRepeatButton" },
                    null, null, null, null, null, Diagnostic: diagnostic),
                new(DateTimeOffset.Now, "unresolved-input", target, null, null, null, null, null, Diagnostic: diagnostic),
                new(DateTimeOffset.Now, "key", target, null, "Control+C", null, null, "excel-selection-range:C1:C20"),
                new(DateTimeOffset.Now, "unresolved-input", target, null, null, null, null, null, Diagnostic: diagnostic),
                new(DateTimeOffset.Now, "unresolved-input", target with { Name = "Chart Area",
                    ControlType = "ControlType.Group" }, null, null, null, null, null, Diagnostic: diagnostic),
                new(DateTimeOffset.Now, "unresolved-input", target, null, null, null, null, null, Diagnostic: "Delayed input")
            };
            var recovered = (List<RecordedEvent>)method.Invoke(null, [source])!;
            if (recovered[0].Kind != "click" || recovered[1].Kind != "click" ||
                recovered[3].Kind != "unresolved-click" || recovered[4].Kind != "unresolved-input" ||
                recovered[5].Kind != "unresolved-input" || source[0].Kind != "unresolved-input" ||
                recovered[2].AfterState != source[2].AfterState)
                throw new InvalidOperationException("Legend-drag recovery lost evidence or repaired unsupported input.");
            Console.WriteLine("PASS: only known non-chart legend-drag regressions recover; unverified worksheet gestures remain manual, raw evidence and genuine unresolved input remain intact.");
            var saved = Path.Combine(Environment.CurrentDirectory, "bin", "RSR", "Recordings", "20261005-075330", "recorded_events.json");
            if (File.Exists(saved))
            {
                var original = File.ReadAllBytes(saved);
                var captured = JsonSerializer.Deserialize<List<RecordedEvent>>(original)!;
                var repaired = (List<RecordedEvent>)method.Invoke(null, [captured])!;
                if (repaired.Any(item => item.Kind == "unresolved-input") ||
                    repaired.Count(item => item.Kind == "unresolved-click") != 5 ||
                    repaired.Count != captured.Count || !original.SequenceEqual(File.ReadAllBytes(saved)))
                    throw new InvalidOperationException("Saved full recording recovery lost events, changed raw evidence or left a planning blocker.");
                Console.WriteLine("PASS: full 154-event recording clears false blockers; five unverified worksheet gestures remain reviewable; raw JSON is unchanged.");
            }
        }
    private static void CheckIntentExpansion(Type foundry)
    {
        var note = "Always replace \"Jose\" with \"Jose Ortega\". If esshrouf runs the app, " +
            "replace \"- Me -\" with \"Esraa\"; if hulopesv runs it, replace \"- Me -\" with \"Hugo\".";
        var target = new ControlRef("EXCEL", "Find and Replace", null, "Replace", "ControlType.TabItem", null, null);
        var events = new List<RecordedEvent>
        {
            new(DateTimeOffset.Now, "click", target, null, null, null, null, null),
            new(DateTimeOffset.Now, "click", target with { Window = "Workbook", Name = "Next action" },
                null, null, null, null, null, note)
        };
        var steps = events.Select((item, i) => new PlanStep(i + 1, item.Kind, item.Target, null, null,
            null, null, null, item.Intent)).ToList();
        using var json = JsonDocument.Parse("""
            [
              {"intentNumber":2,"targetNumber":1,"position":"after","action":"replace-all","value":"Jose","expectedState":"Jose Ortega","whenUser":null,"explanation":"Requested common replacement"},
              {"intentNumber":2,"targetNumber":1,"position":"after","action":"replace-all","value":"- Me -","expectedState":"Esraa","whenUser":"esshrouf","explanation":"Requested conditional replacement"},
              {"intentNumber":2,"targetNumber":1,"position":"after","action":"replace-all","value":"- Me -","expectedState":"Hugo","whenUser":"hulopesv","explanation":"Requested alternate conditional replacement"}
            ]
            """);
        var additions = json.RootElement.EnumerateArray().Select(item => item.Clone()).ToArray();
        var expand = foundry.GetMethod("ExpandIntentActions", BindingFlags.NonPublic | BindingFlags.Static)!;
        var expanded = (List<PlanStep>)expand.Invoke(null, [steps, events, additions])!;
        if (expanded.Count != 5 || expanded.Skip(1).Take(3).Any(step => step.OriginIntent != note ||
            step.Intent != note || step.Target?.Name != "Replace All") ||
            expanded[2].WhenUser != "esshrouf" || expanded[3].WhenUser != "hulopesv")
            throw new InvalidOperationException("A later note did not generate multiple attributed actions in the earlier context.");
        var assembly = typeof(ControlRef).Assembly;
        var compact = assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
        var plan = (ExecutionPlan)compact.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Intent regression", expanded)])!;
        var twice = (ExecutionPlan)compact.Invoke(null, [plan])!;
        if (!plan.Steps.SequenceEqual(twice.Steps) ||
            plan.Steps.Count(step => step.OriginIntent is not null) != 3 ||
            plan.Steps.Single(step => step.Value == "Jose").ExpectedState != "Jose Ortega")
            throw new InvalidOperationException("Compaction changed or duplicated intent-generated replacements.");
        var form = assembly.GetType("DesktopSteps.MainForm")!;
        var problem = form.GetMethod("ReplacementPlanProblem", BindingFlags.NonPublic | BindingFlags.Static)!;
        var incomplete = form.GetMethod("IncompleteReplacements", BindingFlags.NonPublic | BindingFlags.Static)!;
        if (problem.Invoke(null, [plan]) is not null || ((List<int>)incomplete.Invoke(null, [plan])!).Count != 0)
            throw new InvalidOperationException("Explicit intent-generated replacements were incorrectly blocked.");
        var persisted = JsonSerializer.Deserialize<ExecutionPlan>(JsonSerializer.Serialize(plan))!;
        if (persisted.Steps.Count(step => step.OriginIntent == note) != 3)
            throw new InvalidOperationException("Intent provenance did not survive serialization.");
        var dialogSteps = new List<PlanStep>
        {
            new(1, "type", target with { AutomationId = "18", ControlType = "ControlType.Edit" },
                "- Me -", null, "field-value:- Me -", null, null),
            new(2, "type", target with { AutomationId = "21", ControlType = "ControlType.Edit" },
                "Esraa", null, "field-value:Esraa", null, null),
            new(3, "replace-all", target with { Name = "Replace All", ControlType = "ControlType.Button" },
                "- Me -", null, "Esraa", null, null, note, WhenUser: "esshrouf"),
            expanded[1] with { Number = 4 },
            expanded[3] with { Number = 5 },
            expanded[2] with { Number = 6 }
        };
        var recordedAndAdded = (ExecutionPlan)compact.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Recorded and additional replacements", dialogSteps)])!;
        var mappings = recordedAndAdded.Steps.Where(step => step.Action == "replace-all").ToArray();
        if (mappings.Length != 3 || mappings.Count(step => step.Value == "- Me -" && step.ExpectedState == "Esraa") != 1 ||
            !mappings.Any(step => step.WhenUser == "hulopesv" && step.ExpectedState == "Hugo") ||
            !mappings.Any(step => step.Value == "Jose" && step.ExpectedState == "Jose Ortega"))
            throw new InvalidOperationException("Recorded fields overwrote intent replacements or duplicate additions survived.");
        var forwardEvents = new List<RecordedEvent>
        {
            events[1],
            events[0],
            events[0]
        };
        var forwardSteps = forwardEvents.Select((item, index) => new PlanStep(index + 1,
            item.Kind, item.Target, null, null, null, null, null, item.Intent)).ToList();
        using var forwardJson = JsonDocument.Parse("""
            [
              {"intentNumber":1,"targetNumber":2,"position":"after","action":"replace-all","value":"Jose","expectedState":"Jose Ortega","explanation":"First upcoming action"},
              {"intentNumber":1,"targetNumber":3,"position":"after","action":"replace-all","value":"- Me -","expectedState":"Esraa","whenUser":"esshrouf","explanation":"Second upcoming action"}
            ]
            """);
        var forward = (List<PlanStep>)expand.Invoke(null, [forwardSteps, forwardEvents,
            forwardJson.RootElement.EnumerateArray().Select(item => item.Clone()).ToArray()])!;
        if (forward.Count != 5 || forward[2].Value != "Jose" || forward[4].Value != "- Me -")
            throw new InvalidOperationException("A note did not apply across multiple upcoming actions.");
        foreach (var invalid in new[]
        {
            """[{"intentNumber":2,"targetNumber":1,"position":"after","action":"replace-all","value":"find: Jose -> replace: Jose Ortega","expectedState":"All occurrences replaced","explanation":"Invalid prose"}]""",
            """[{"intentNumber":2,"targetNumber":1,"position":"after","action":"replace-all","value":"Jose","expectedState":"Jose Ortega","whenUser":"when user esshrouf runs it","explanation":"Invalid condition"}]""",
            """[{"intentNumber":2,"targetNumber":2,"position":"after","action":"replace-all","value":"Jose","expectedState":"Jose Ortega","explanation":"Wrong window"}]"""
        })
        {
            using var bad = JsonDocument.Parse(invalid);
            try
            {
                expand.Invoke(null, [steps, events,
                    bad.RootElement.EnumerateArray().Select(item => item.Clone()).ToArray()]);
                throw new InvalidOperationException("Unsafe intent-generated action was accepted.");
            }
            catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
        }
        events[1] = events[1] with { Intent = null };
        try
        {
            expand.Invoke(null, [steps, events, additions]);
            throw new InvalidOperationException("Unattributed generated actions were accepted.");
        }
        catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
        Console.WriteLine("PASS: multi-step backward intent, conditional added actions, provenance and idempotence.");
    }

    private static void CheckOptionalIntentExpansion(Type foundry)
    {
        var note = "If present, click Yes in \"RSR optional regression dialog\"; otherwise click the refresh button.";
        var target = new ControlRef("SyntheticHost", "Main view", "refresh", "Refresh View",
            "ControlType.Button", null, "Commands");
        var events = new List<RecordedEvent>
        {
            new(DateTimeOffset.Now, "click", target, null, null, null, null, null, note)
        };
        var steps = new List<PlanStep> { new(1, "click", target, null, null, null, null, null, note) };
        using var json = JsonDocument.Parse("""
            [{"intentNumber":1,"targetNumber":1,"position":"before","action":"optional-click",
              "window":"RSR optional regression dialog","name":"Yes","explanation":"Requested optional dialog"}]
            """);
        var action = json.RootElement[0];
        var expand = foundry.GetMethod("ExpandIntentActions", BindingFlags.NonPublic | BindingFlags.Static)!;
        var expanded = (List<PlanStep>)expand.Invoke(null, [steps, events, new[] { action }])!;
        if (expanded.Count != 2 || expanded[0].Action != "optional-click" ||
            expanded[0].Target?.Name != "Yes" || expanded[0].Target?.AutomationId is not null ||
            expanded[0].Target?.Window != "RSR optional regression dialog" ||
            expanded[0].OriginIntent != note || expanded[1].Action != "click-if-previous-absent" ||
            expanded[1].Target != target)
            throw new InvalidOperationException("Optional intent did not preserve its exact dialog and recorded fallback.");
        var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
        var plan = (ExecutionPlan)compact.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now, "Branch", expanded)])!;
        if (!plan.Steps.SequenceEqual(expanded))
            throw new InvalidOperationException("Compaction changed the optional branch.");
        var timing = typeof(ControlRef).Assembly.GetType("DesktopSteps.RefreshTiming")!.GetMethod("AddObservedWait")!;
        var timed = (ExecutionPlan)timing.Invoke(null, [plan, new List<RecordedEvent>
        {
            events[0],
            events[0] with { At = events[0].At.AddSeconds(18), Target = target with { Name = "Next control" } }
        }])!;
        if (timed.Steps[1].Value != "refresh-observed-ms:18000")
            throw new InvalidOperationException("Optional fallback lost its recorded refresh wait.");
        var conditional = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsConditionalStep", BindingFlags.NonPublic | BindingFlags.Static)!;
        if ((bool)conditional.Invoke(null, [steps[0]])! ||
            !(bool)conditional.Invoke(null, [expanded[0]])!)
            throw new InvalidOperationException("Optional dialog explanation made the Refresh command optional.");
        var actualNote = "the pop up \"RSR optional regression dialog\", if present intent is to click Yes button, " +
            "otherwise proceed to next steps. automatic refresh will occur if box appears so no need to click refresh button in this case, otherwise, do click it.";
        var launch = target with { Name = "(Legacy) SD Case Triage", ControlType = "ControlType.MenuItem" };
        var workflowEvents = new List<RecordedEvent>
        {
            events[0] with { Target = launch, Intent = actualNote },
            events[0] with { Intent = null }
        };
        var workflowSteps = new List<PlanStep>
        {
            steps[0] with { Target = launch, Intent = actualNote },
            steps[0] with { Number = 2, Intent = actualNote }
        };
        using var misplaced = JsonDocument.Parse("""
            [{"intentNumber":1,"targetNumber":1,"position":"before","action":"optional-click",
              "window":"RSR optional regression dialog","name":"Yes","explanation":"Wrong anchor"}]
            """);
        try
        {
            expand.Invoke(null, [workflowSteps, workflowEvents, new[] { misplaced.RootElement[0] }]);
            throw new InvalidOperationException("Optional wait before its launch command was accepted.");
        }
        catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
        using var placed = JsonDocument.Parse(misplaced.RootElement.GetRawText().Replace("\"targetNumber\":1", "\"targetNumber\":2"));
        var ordered = (List<PlanStep>)expand.Invoke(null,
            [workflowSteps, workflowEvents, new[] { placed.RootElement[0] }])!;
        if (ordered[0].Action != "click" || ordered[0].Target != launch ||
            ordered[1].Action != "optional-click" || ordered[2].Action != "click-if-previous-absent" ||
            (bool)conditional.Invoke(null, [workflowSteps[1]])!)
            throw new InvalidOperationException("Workflow intent changed launch/wait/fallback dependency order.");
        foreach (var invalid in new[]
        {
            action.GetRawText().Replace("RSR optional regression dialog", "Unspecified dialog"),
            action.GetRawText().Replace("\"Yes\"", "\"Delete\""),
            action.GetRawText().Replace("\"before\"", "\"after\"")
        })
        {
            using var bad = JsonDocument.Parse(invalid);
            try
            {
                expand.Invoke(null, [steps, events, new[] { bad.RootElement }]);
                throw new InvalidOperationException("An unspecified optional target was accepted.");
            }
            catch (TargetInvocationException ex) when (ex.InnerException is InvalidDataException) { }
        }
        Console.WriteLine("PASS: explicit optional-dialog generation, recorded fallback, scope and invalid-target rejection.");
    }

    private static void CheckSortCompaction()
    {
        var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
        var target = new ControlRef("SyntheticHost", "Cases", "age", "Age", "ControlType.HeaderItem", null, null);
        foreach (var desired in new[] { "unspecified", "ascending", "descending" })
        {
            var steps = new List<PlanStep>
            {
                new(1, "click", target, null, null, null, null, null),
                new(2, "click", target, null, null, desired == "unspecified" ? null : "sort:" + desired, null, null)
            };
            var plan = (ExecutionPlan)compact.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now, "Sort", steps)])!;
            if (plan.Steps.Count != 1 || plan.Steps[0].Action != "ensure-state" ||
                plan.Steps[0].ExpectedState != "sort:" + desired)
                throw new InvalidOperationException("Header clicks did not compact into their final sort outcome.");
            var twice = (ExecutionPlan)compact.Invoke(null, [plan])!;
            if (!twice.Steps.SequenceEqual(plan.Steps))
                throw new InvalidOperationException("Sort compaction is not idempotent.");
        }
        var explicitIntent = new ExecutionPlan(1, DateTimeOffset.Now, "Sort intent",
            [new(1, "click", target, null, null, null, null, null, "Sort Age descending.")]);
        var inferred = (ExecutionPlan)compact.Invoke(null, [explicitIntent])!;
        if (inferred.Steps[0].ExpectedState != "sort:descending")
            throw new InvalidOperationException("Explicit sort intent was ignored.");
        Console.WriteLine("PASS: sort toggles collapse into verified final state, explicit intent or unspecified replay choice.");
    }

    private static void CheckSupersededCellClick()
    {
        var assembly = typeof(ControlRef).Assembly;
        var compactor = assembly.GetType("DesktopSteps.PlanCompactor")!;
        var target = new ControlRef("EXCEL", "Synthetic workbook", "", "F9", "ControlType.DataItem", "", "Table");
        var click = new PlanStep(1, "click", target, null, null, null, null, null);
        var type = new PlanStep(2, "type", target with { Name = "L9", AutomationId = "L9",
            ClassName = "XLSpreadsheetCell" }, "FL", null, null, null, null);
        var matches = compactor.GetMethod("IsSupersededExcelCellClick", BindingFlags.NonPublic | BindingFlags.Static)!;
        bool Match(PlanStep a, PlanStep b) => (bool)matches.Invoke(null, [a, b])!;
        if (!Match(click, type) || Match(click with { Intent = "Select this range" }, type) ||
            Match(click with { ExpectedState = "selected" }, type) ||
            Match(click with { Target = target with { ClassName = "XLSpreadsheetCell", AutomationId = "F9" } }, type) ||
            Match(click, type with { Target = type.Target! with { Window = "Another workbook" } }) ||
            Match(click, type with { Action = "key", Key = "Control+V" }) ||
            Match(click, type with { Target = type.Target! with { AutomationId = "L0" } }))
            throw new InvalidOperationException("Coarse grid-click compaction removed meaningful selections or missed a superseded click.");
        var plan = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Precise typing", [click, type])])!;
        if (plan.Steps is not [{ Action: "type", Target.AutomationId: "L9", Value: "FL" }])
            throw new InvalidOperationException("Precise Excel typing target was changed during compaction.");
        var filter = new PlanStep(1, "filter-values", target with { Name = "Filter applied", ParentName = "G1",
            AutomationId = "Dropdown", ControlType = "ControlType.MenuItem" }, "[\"Alpha\"]", null, "filter:only", null, null);
        var fill = new PlanStep(4, "fill-down-to-adjacent-data-end", target, null, null, "adjacent-column:K", null, null);
        var filtered = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Filtered fill", [filter, click, type, fill])])!;
        if (filtered.Steps[1].TargetStrategy != "first-visible-filtered-row" ||
            filtered.Steps[1].Target?.AutomationId != "L9" || filtered.Steps[2].Action != fill.Action)
            throw new InvalidOperationException("Filtered fill did not retain column evidence with a dynamic visible-row strategy.");
        var fixedEntry = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Fixed entry", [filter, type])])!;
        if (fixedEntry.Steps[1].TargetStrategy is not null)
            throw new InvalidOperationException("An ordinary fixed-address entry was made relative without fill intent.");
        var header = target with { ControlType = "ControlType.HeaderItem", Name = "Synthetic header" };
        var context = new PlanStep(1, "context-selection", header, null, null, "selection-intent:all", null, null);
        var rowContext = context with { Number = 2, Target = target with { Name = "Recorded person",
            ControlType = "ControlType.Text" } };
        var copy = new PlanStep(3, "click", target with { Name = "Copy selected rows",
            ControlType = "ControlType.MenuItem" }, null, null, null, null, null);
        var selected = (ExecutionPlan)compactor.GetMethod("Compact")!.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Selected rows", [context, rowContext, copy])])!;
        if (selected.Steps.Count != 2 || selected.Steps[0].Target != header)
            throw new InvalidOperationException("All-row context selection retained a recorded person's locator.");
        Console.WriteLine("PASS: coarse pre-typing grid click is superseded only by independently verified same-workbook cell input.");
    }

    private static void CheckFilterOutcomes()
    {
        var compact = typeof(ControlRef).Assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
        var column = new ControlRef("EXCEL", "Synthetic workbook", "Dropdown", "No filter applied",
            "ControlType.MenuItem", "", "C1");
        var item = column with { AutomationId = "", Name = "(Select All)", ControlType = "ControlType.TreeItem",
            ParentName = "Manual Filter" };
        var steps = new List<PlanStep>
        {
            new(1, "click", column, null, null, null, null, null),
            new(2, "click", item, null, null, null, null, null),
            new(3, "click", item with { Name = "Alpha" }, null, null, null, null, null),
            new(4, "click", item with { Name = "Alpha" }, null, null, null, null, null),
            new(5, "click", column with { AutomationId = "", Name = "OK", ControlType = "ControlType.Button",
                ClassName = "NetUIButton", ParentName = "" }, null, null, null, null, null)
        };
        ExecutionPlan Run(List<PlanStep> source) => (ExecutionPlan)compact.Invoke(null,
            [new ExecutionPlan(1, DateTimeOffset.Now, "Filter", source)])!;
        var unknown = Run(steps);
        if (unknown.Steps is not [{ Action: "filter-values", ExpectedState: "filter:unspecified", Value: "[\"Alpha\"]" }] ||
            Run(unknown.Steps).Steps.Count != 1)
            throw new InvalidOperationException("Legacy filter clicks guessed a state or did not compact idempotently.");
        steps[1] = steps[1] with { ExpectedState = "filter:unchecked" };
        steps[2] = steps[2] with { ExpectedState = "filter:unchecked" };
        steps[3] = steps[3] with { ExpectedState = "filter:checked" };
        if (Run(steps).Steps[0] is not { ExpectedState: "filter:only", Value: "[\"Alpha\"]" })
            throw new InvalidOperationException("Recorded final checked state was not honored.");
        steps[1] = steps[1] with { ExpectedState = "filter:checked" };
        steps[3] = steps[3] with { ExpectedState = "filter:unchecked" };
        if (Run(steps).Steps[0] is not { ExpectedState: "filter:exclude", Value: "[\"Alpha\"]" })
            throw new InvalidOperationException("Recorded exclusion was not honored.");
        var boundary = new PlanStep(10, "click", column with { AutomationId = "SheetTab", Name = "Results",
            ControlType = "ControlType.TabItem", ParentName = "Synthetic workbook" }, null, null, null,
            "filter-closed.jpg", null);
        var implicitSequences = new List<PlanStep>
        {
            new(1, "click", column, null, null, null, null, null),
            new(2, "click", item, null, null, "filter:mixed", null, null),
            new(3, "click", item with { Name = "Alpha" }, null, null, null, null, null),
            boundary,
            new(5, "click", column with { Name = "Filter applied" }, null, null, null, null, null),
            new(6, "click", item, null, null, "filter:mixed", null, null),
            new(7, "click", item with { Name = "Alpha" }, null, null, null, null, null),
            boundary with { Number = 8 },
            new(9, "click", column with { Name = "Filter applied" }, null, null, null, null, null),
            new(10, "click", item, null, null, "filter:checked", null, null),
            boundary with { Number = 11 }
        };
        var recoveredPlan = Run(implicitSequences);
        var recovered = recoveredPlan.Steps.Where(step => step.Action == "filter-values").ToArray();
        if (recovered is not
            [
                { ExpectedState: "filter:only", Value: "[\"Alpha\"]" },
                { ExpectedState: "filter:exclude", Value: "[\"Alpha\"]" },
                { ExpectedState: "filter:all", Value: "[]" }
            ])
            throw new InvalidOperationException("Closed filter flyouts with a lost apply click were not reconstructed from their transitions.");
        if (recoveredPlan.Steps.Any(step => step.Target?.Name == "Results"))
            throw new InvalidOperationException("The control exposed below a closed filter flyout was replayed as a second click.");
        var alreadyCompacted = Run([
            recovered[0],
            boundary with { Screenshot = recovered[0].Screenshot }
        ]);
        if (alreadyCompacted.Steps.Count != 1 || alreadyCompacted.Steps[0].Action != "filter-values")
            throw new InvalidOperationException("A saved plan retained the sheet tab misidentified beneath a closed filter flyout.");
        var bareColumn = new PlanStep(20, "manual-click", column with { AutomationId = "", Name = "C",
            ControlType = "ControlType.DataItem", ParentName = "Grid" }, null, null, "unverified-click", null, null);
        var priorFilter = new PlanStep(19, "filter-values", column, "[\"Alpha\"]", null,
            "filter:exclude", null, null);
        var recoveredBareColumn = Run([
            priorFilter,
            bareColumn,
            new PlanStep(21, "click", item, null, null, "filter:mixed", null, null),
            new PlanStep(22, "click", item with { Name = "Alpha" }, null, null, "filter:unchecked", null, null),
            boundary with { Number = 23 }
        ]);
        if (recoveredBareColumn.Steps.Any(step => step.Action == "manual-click") ||
            recoveredBareColumn.Steps.Last() is not
                { Action: "filter-values", ExpectedState: "filter:exclude", Value: "[\"Alpha\"]" } ||
            !Run(recoveredBareColumn.Steps).Steps.SequenceEqual(recoveredBareColumn.Steps))
            throw new InvalidOperationException("A bare Excel column filter opener was not recovered from its completed checklist.");
        var conditional = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsConditionalStep", BindingFlags.NonPublic | BindingFlags.Static)!;
        var occasionalDialog = new PlanStep(1, "click", column with { Process = "EXCEL", Window = "",
            AutomationId = "", Name = "See just mine", ControlType = "ControlType.Button",
            ParentName = "Others are also making changes" }, null, null, null, null, null,
            "the see everyone or just see mine dialog doesn't always appear");
        if (!(bool)conditional.Invoke(null, [occasionalDialog])!)
            throw new InvalidOperationException("A recorded occasional dialog was treated as mandatory.");
        var occasional = typeof(ControlRef).Assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsRecordedOccasionalDialog", BindingFlags.NonPublic | BindingFlags.Static)!;
        if (!(bool)occasional.Invoke(null, [occasionalDialog])! ||
            (bool)occasional.Invoke(null, [occasionalDialog with { Intent = "Click this button." }])!)
            throw new InvalidOperationException("Occasional dialog absence was not distinguished from a mandatory recorded click.");
        Console.WriteLine("PASS: duplicate filter clicks become confirmed final outcomes; unknown states require a choice.");
    }

    private static void CheckIntentEditsAndAcknowledgements()
    {
        var assembly = typeof(ControlRef).Assembly;
        var noMatch = assembly.GetType("DesktopSteps.Executor")!
            .GetMethod("IsNoMatchReplacementResult", BindingFlags.NonPublic | BindingFlags.Static)!;
        foreach (var message in new[] { "We couldn't find anything to replace.", "Cannot find anything to replace.",
            "Can't find anything to replace.", "Could not find anything to replace.", "We made 0 replacements." })
            if (!(bool)noMatch.Invoke(null, [message])!)
                throw new InvalidOperationException("Replacement warning was misclassified: " + message);
        if ((bool)noMatch.Invoke(null, ["All done. We made 9 replacements."])!)
            throw new InvalidOperationException("Replacement success was misclassified as a warning.");
        var target = new ControlRef("Synthetic", "Window", "field", "Field", "ControlType.Edit", null, null);
        var source = new List<RecordedEvent>
        {
            new(DateTimeOffset.Now, "key", target, "value", "A", "image.jpg", "before", "after", "Original"),
            new(DateTimeOffset.Now, "click", target, null, null, null, null, null, "Original"),
            new(DateTimeOffset.Now, "click", target, null, null, null, null, null, "Remove me")
        };
        var edits = new Dictionary<string, string> { ["Original"] = "Updated", ["Remove me"] = "" };
        var updated = (List<RecordedEvent>)assembly.GetType("DesktopSteps.MainForm")!
            .GetMethod("ApplyIntentEdits", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [source, edits])!;
        if (updated[0] != source[0] with { Intent = "Updated" } ||
            updated[1] != source[1] with { Intent = "Updated" } ||
            updated[2] != source[2] with { Intent = null } || source[0].Intent != "Original")
            throw new InvalidOperationException("Editing intents altered recorded evidence or did not update shared notes.");
        var compact = assembly.GetType("DesktopSteps.PlanCompactor")!.GetMethod("Compact")!;
        var replacement = new PlanStep(1, "replace-all", target with { Process = "EXCEL", Window = "Find and Replace",
            Name = "Replace All", ControlType = "ControlType.Button" }, "old", null, "new", null, null,
            "Replace old with new", OriginIntent: "Replace old with new");
        var ok = new PlanStep(2, "click", replacement.Target! with { Window = "Microsoft Excel", Name = "OK" },
            null, null, null, null, null);
        var close = replacement with { Number = 2, Action = "close-window", Value = null,
            ExpectedState = "window-closed", OriginIntent = null };
        foreach (var steps in new[] { new List<PlanStep> { replacement, ok },
            new List<PlanStep> { replacement, close, ok with { Number = 3 } } })
        {
            var plan = (ExecutionPlan)compact.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now, "Acknowledgement", steps)])!;
            if (plan.Steps.Any(step => step.Action == "click" && step.Target?.Name == "OK"))
                throw new InvalidOperationException("Already-consumed replacement acknowledgement remained in the plan.");
        }
        var unrelated = (ExecutionPlan)compact.Invoke(null, [new ExecutionPlan(1, DateTimeOffset.Now, "Unrelated OK",
            [replacement with { Action = "click", Target = replacement.Target! with { Name = "Other" } }, ok])])!;
        if (!unrelated.Steps.Any(step => step.Target?.Name == "OK"))
            throw new InvalidOperationException("An unrelated OK button was incorrectly removed.");
        Console.WriteLine("PASS: intent edits preserve all recorded evidence; consumed result OK is removed and unrelated OK retained.");
    }

    private sealed class StubHandler : HttpMessageHandler
    {
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            if (request.Headers.Authorization?.Parameter != "synthetic-token" ||
                request.RequestUri?.AbsolutePath != "/api/projects/test/agents/test/endpoint/protocols/openai/responses")
                throw new InvalidOperationException("Unexpected Foundry request.");
            var body = await request.Content!.ReadAsStringAsync(token);
            if (body.Contains("timeout", StringComparison.Ordinal))
                await Task.Delay(Timeout.InfiniteTimeSpan, token);
            if (body.Contains("delayed", StringComparison.Ordinal))
                await Task.Delay(100, token);
            var failure = body.Contains("failure", StringComparison.Ordinal);
            return new HttpResponseMessage(failure ? HttpStatusCode.NotFound : HttpStatusCode.OK)
            {
                Content = new StringContent(failure
                    ? "{\"error\":{\"code\":\"WorkspaceNotFound\"}}"
                    : "{\"output\":[{\"content\":[{\"text\":\"Synthetic response\"}]}]}")
            };
        }
    }
}
