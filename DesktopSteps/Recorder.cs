using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Channels;
using System.Windows.Automation;

namespace DesktopSteps;

// Captures low-level input promptly and enriches it with accessibility evidence off
// the hook thread. Ordered delivery prevents slow UIA providers from reordering the
// user's actions, while unresolved identities remain explicit instead of being guessed.
internal sealed class Recorder : IDisposable
{
    private const int WH_KEYBOARD_LL = 13, WH_MOUSE_LL = 14;
    private const int WM_KEYDOWN = 0x100, WM_SYSKEYDOWN = 0x104, WM_LBUTTONDOWN = 0x201,
        WM_LBUTTONUP = 0x202,
        WM_RBUTTONDOWN = 0x204, WM_RBUTTONUP = 0x205, WM_MOUSEWHEEL = 0x20A, WM_MOUSEHWHEEL = 0x20E;
    private readonly List<RecordedEvent> events = [];
    private readonly HashSet<string> ignoredProcesses = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<Task> pendingScrollCaptures = [];
    private readonly List<Task> pendingFieldCaptures = [];
    private readonly List<Task> pendingWindowCaptures = [];
    private readonly List<Task> pendingProbeCaptures = [];
    private sealed record RowAnchorSnapshot(int Row, string Current, string Next, string Following);
    private readonly Dictionary<string, (DateTimeOffset At, Task<RowAnchorSnapshot?> Capture)> rowAnchorCaptures =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<nint, (DateTime At, NativeRect WindowBounds,
        List<(string Name, string Type, System.Windows.Rect Bounds, ControlRef Target)> Controls)> dialogControlCache = [];
    private readonly HashSet<nint> queuedDialogSnapshots = [];
    private readonly Dictionary<nint, int> dialogCacheVersions = [];
    private readonly object dialogCacheLock = new();
    private readonly Dictionary<nint, (AutomationElement Root, AutomationEventHandler Invoked,
        AutomationEventHandler Selected)> dialogActionHandlers = [];
    private readonly Channel<Action> enrichmentQueue = Channel.CreateUnbounded<Action>(
        new UnboundedChannelOptions { SingleReader = true, SingleWriter = false });
    private readonly Task enrichmentWorker;
    private readonly Control hookDispatcher = new();
    private readonly OrderedCaptureQueue inputDeliveries = new();
    private Thread? hookThread;
    private Control? hookLoop;
    private readonly HookProc keyboardProc, mouseProc;
    private readonly System.Windows.Forms.Timer typingCapture = new() { Interval = 450 };
    private readonly System.Windows.Forms.Timer clickCapture = new() { Interval = 400 };
    private readonly System.Windows.Forms.Timer scrollCapture = new() { Interval = 500 };
    private readonly System.Windows.Forms.Timer chartHoverCapture = new() { Interval = 150 };
    private readonly AutoResetEvent chartSampleRequested = new(false);
    private Thread? chartSampleThread;
    private volatile bool chartSampleStopping;
    private volatile bool chartSamplePaused;
    private long chartInputGeneration;
    private readonly System.Windows.Forms.Timer legendOutcomeCapture = new() { Interval = 300 };
    private Action? pendingLegendOutcome;
    private readonly System.Collections.Concurrent.ConcurrentQueue<string> captureDiagnostics = new();
    private bool captureDiagnosticsSaved;

    private void CaptureDiagnostic(string message)
    {
        if (captureDiagnostics.Count < 2000)
            captureDiagnostics.Enqueue($"{DateTimeOffset.Now:O} {message}");
    }

    private T MeasureCapture<T>(string stage, Func<T> capture)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        try { return capture(); }
        finally
        {
            if (watch.ElapsedMilliseconds >= 100)
                CaptureDiagnostic($"{stage}: {watch.ElapsedMilliseconds} ms");
        }
    }
    private nint keyboardHook, mouseHook;
    private string? directory;
    private int imageNumber;
    private Screen? typingScreen, scrollScreen;
    private AutomationElement? scrollElement;
    private int scrollEventIndex = -1;
    private int pendingClickIndex = -1;
    private int dragClickIndex = -1;
    private System.Drawing.Point dragClickPoint;
    private string? dragChartName;
    private nint dragChartWindow;
    private string? dragChartSheet;
    private ExcelNativeSheet.LegendLayout? dragLegendBefore;
    private sealed record HoveredChart(DateTimeOffset At, System.Windows.Rect Bounds, MouseDownSnapshot Snapshot);
    private volatile HoveredChart? hoveredChart;
    private sealed record PointerMenu(DateTimeOffset At, long Generation, nint Foreground,
        MouseDownSnapshot Snapshot);
    private volatile PointerMenu[]? pointerMenus;
    private readonly List<Task> pendingMenuCaptures = [];
    private volatile bool chartPointerDown;
    private sealed record ChartSourceCapture(nint Window, ControlRef Target, string? Sheet,
        ExcelNativeSheet.ChartSource? Before);
    private volatile ChartSourceCapture? chartSourceCapture;
    private long lastChartInputTick;
    private Screen? pendingClickScreen;
    private System.Drawing.Point? pendingClickPoint;
    private nint pendingClickWindow;
    private string? pendingClickTabBefore;
    private ControlRef? pendingClickFocusBefore;
    private int resizeStartIndex = -1;
    private System.Drawing.Point resizeStartPoint;
    private string resizeEdge = "right";
    private bool scrollDragPending;
    private string? scrollAxis;
    private volatile bool capturePausedForIntent;
    private volatile bool capturePausedByUser;
    private int finishingTypingCapture;
    private nint intentDialogWindow;
    private nint cachedKeyboardFocus, cachedKeyboardWindow;
    private nint lastForegroundWindow;
    private ControlRef? cachedKeyboardTarget;
    private Screen? cachedKeyboardScreen;
    private ControlRef? cachedEditableTarget;
    private AutomationElement? cachedEditableElement;
    private Screen? cachedEditableScreen;
    public bool IsRecording => keyboardHook != 0 || mouseHook != 0;
    public bool IsPaused => IsRecording && capturePausedByUser;
    public IReadOnlyList<RecordedEvent> Events => events;
    public event Action<RecordedEvent>? Captured;
    public event Action<nint>? IntentRequested;
    public event Action<ControlRef>? TargetCorrected;
    public event Action<int, string>? UnresolvedDialogClick;

    public bool IsUnresolvedDialogClick(int index)
    {
        lock (events)
            return index >= 0 && index < events.Count &&
                (events[index].Kind == "unresolved-click" || IsUnresolvedDialogEvent(events[index]));
    }

    public string DescribeRecovery(int index)
    {
        lock (events)
        {
            if (index < 0 || index >= events.Count)
                return "The affected action is no longer present in this recording.";
            static string Describe(RecordedEvent item) =>
                $"{item.At:HH:mm:ss.fff}: {item.Kind} " +
                $"'{item.Target?.Name ?? item.Key ?? "(unidentified control)"}' " +
                $"in '{item.Target?.Window ?? "(unknown window)"}'";
            var item = events[index];
            return $"Repeat from recorded action {index + 1} ({Describe(item)}).\n" +
                (index > 0 ? $"Last kept action {index}: {Describe(events[index - 1])}.\n" : "") +
                $"Affected range: actions {index + 1}-{events.Count}, through {events[^1].At:HH:mm:ss.fff}.\n" +
                string.Join("\n", events.Skip(index + 1).Take(5)
                    .Select((following, offset) => $"Also repeat action {index + offset + 2}: {Describe(following)}.")) +
                (events.Count - index > 6 ? "\nRepeat the remaining later actions as well." : "");
        }
    }

    private static bool IsUnresolvedDialogEvent(RecordedEvent item) =>
        item.Kind == "click" && item.AfterState is not
            ("dialog-command" or "dialog-command-result" or "window-closed" or
             "dialog-result-acknowledgement" or "worksheet-tab-focus-only") &&
        item.Target is { Name.Length: > 0, ControlType: "ControlType.Window" or
            "ControlType.Custom" or "ControlType.Tab" };

    public void SetIgnoredProcesses(IEnumerable<string> processes)
    {
        ignoredProcesses.Clear();
        foreach (var process in processes)
            ignoredProcesses.Add(Path.GetFileNameWithoutExtension(process.Trim()));
    }

    private bool IsIgnored(ControlRef? target)
    {
        if (target?.Process is not { } process) return false;
        if (ignoredProcesses.Contains(process)) return true;
        if (!process.Equals("explorer", StringComparison.OrdinalIgnoreCase) ||
            target.ControlType != "ControlType.Button") return false;
        // Windows owns taskbar buttons in explorer, even when the button
        // represents an application excluded by a recording rule.
        var app = System.Text.RegularExpressions.Regex.Match(target.Name ?? "",
            @"^(?<app>.+?) - \d+ running windows?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return app.Success && ignoredProcesses.Contains(app.Groups["app"].Value);
    }

    private bool IsIgnoredForeground()
    {
        var window = GetForegroundWindow();
        if (window == 0) return false;
        GetWindowThreadProcessId(window, out var processId);
        if (processId == 0) return false;
        try { return ignoredProcesses.Contains(System.Diagnostics.Process.GetProcessById((int)processId).ProcessName); }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return false; }
    }

    public Recorder()
    {
        hookDispatcher.CreateControl();
        enrichmentWorker = Task.Run(async () =>
        {
            await foreach (var action in enrichmentQueue.Reader.ReadAllAsync())
                try { action(); }
                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
        });
        keyboardProc = OnKeyboard; mouseProc = OnMouse;
        typingCapture.Tick += (_, _) => FinishTypingCapture();
        clickCapture.Tick += (_, _) => FinishPendingClick();
        scrollCapture.Tick += (_, _) => CompleteScroll();
        chartHoverCapture.Tick += (_, _) =>
        {
            CompleteChartSourceCapture();
            chartSamplePaused = capturePausedByUser || capturePausedForIntent || pendingLegendOutcome is not null;
            chartSampleRequested.Set();
        };
        legendOutcomeCapture.Tick += (_, _) => FinishLegendOutcome();
    }

    public void Start(string outputDirectory)
    {
        if (keyboardHook != 0 || mouseHook != 0) throw new InvalidOperationException("Already recording.");
        if (pendingMenuCaptures.Any(task => !task.IsCompleted))
            throw new InvalidOperationException("The previous context-menu capture has not stopped.");
        pendingMenuCaptures.Clear();
        pointerMenus = null;
        Interlocked.Increment(ref chartInputGeneration);
        directory = outputDirectory;
        captureDiagnostics.Clear();
        captureDiagnosticsSaved = false;
        Directory.CreateDirectory(outputDirectory);
        events.Clear(); imageNumber = 0;
        capturePausedByUser = false;
        cachedKeyboardFocus = cachedKeyboardWindow = 0;
        lastForegroundWindow = 0;
        cachedKeyboardTarget = null; cachedKeyboardScreen = null;
        cachedEditableTarget = null; cachedEditableScreen = null;
        cachedEditableElement = null;
        pendingClickIndex = -1; pendingClickScreen = null; pendingClickPoint = null;
        dragClickIndex = -1; dragChartName = null; dragLegendBefore = null;
        hoveredChart = null; chartPointerDown = false;
        chartSourceCapture = null;
        pendingClickWindow = 0; pendingClickTabBefore = null;
        pendingClickFocusBefore = null;
        resizeStartIndex = -1;
        pendingScrollCaptures.Clear();
        pendingFieldCaptures.Clear();
        pendingWindowCaptures.Clear();
        pendingProbeCaptures.Clear();
        rowAnchorCaptures.Clear();
        lock (dialogCacheLock)
        {
            dialogControlCache.Clear();
            queuedDialogSnapshots.Clear();
            dialogCacheVersions.Clear();
        }
        RemoveDialogActionHandlers();
        CompleteChartSourceCapture();
        StartChartSampling(PrimeHoveredChart);
        chartHoverCapture.Start();
        using var ready = new ManualResetEventSlim();
        Exception? hookError = null;
        hookThread = new Thread(() =>
        {
            using var loop = new Control();
            loop.CreateControl();
            hookLoop = loop;
            using var dialogs = new System.Windows.Forms.Timer { Interval = 150 };
            dialogs.Tick += (_, _) => PrimeVisibleDialogCaches();
            keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardProc, 0, 0);
            mouseHook = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, 0, 0);
            if (keyboardHook == 0 || mouseHook == 0)
                hookError = new System.ComponentModel.Win32Exception();
            ready.Set();
            try
            {
                if (hookError is null)
                {
                    dialogs.Start();
                    Application.Run();
                }
            }
            finally
            {
                if (keyboardHook != 0) UnhookWindowsHookEx(keyboardHook);
                if (mouseHook != 0) UnhookWindowsHookEx(mouseHook);
                keyboardHook = mouseHook = 0;
                hookLoop = null;
            }
        }) { IsBackground = true, Name = "RSR input hooks" };
        hookThread.SetApartmentState(ApartmentState.STA);
        hookThread.Start();
        ready.Wait();
        if (hookError is not null)
        {
            hookThread.Join();
            hookThread = null;
            chartHoverCapture.Stop();
            StopChartSampling();
            throw hookError;
        }
    }

    public void Stop()
    {
        Interlocked.Increment(ref chartInputGeneration);
        pointerMenus = null;
        chartHoverCapture.Stop();
        chartSampleStopping = true;
        chartSampleRequested.Set();
        capturePausedByUser = false;
        capturePausedForIntent = false;
        intentDialogWindow = 0;
        if (keyboardHook != 0) UnhookWindowsHookEx(keyboardHook);
        if (mouseHook != 0) UnhookWindowsHookEx(mouseHook);
        keyboardHook = mouseHook = 0;
        if (hookLoop is { } loop && !loop.IsDisposed)
            loop.BeginInvoke(new Action(Application.ExitThread));
        if (hookThread is not null)
        {
            if (!hookThread.Join(TimeSpan.FromSeconds(10)))
                throw new TimeoutException("Recording hooks did not stop within ten seconds.");
            hookThread = null;
        }
        inputDeliveries.Drain();
        StopChartSampling();
        FinishLegendOutcome();
        FinishTypingCapture();
        FinishPendingClick();
        CompleteChartSourceCapture();
        CompleteScroll();
        RemoveDialogActionHandlers();
        try { Task.WaitAll(pendingScrollCaptures.ToArray()); }
        catch (AggregateException ex) { System.Diagnostics.Trace.WriteLine(ex); }
        pendingScrollCaptures.Clear();
        try { Task.WaitAll(pendingFieldCaptures.ToArray()); }
        catch (AggregateException ex) { System.Diagnostics.Trace.WriteLine(ex); }
        pendingFieldCaptures.Clear();
        try { Task.WaitAll(pendingWindowCaptures.ToArray()); }
        catch (AggregateException ex) { System.Diagnostics.Trace.WriteLine(ex); }
        pendingWindowCaptures.Clear();
        try { Task.WaitAll(pendingProbeCaptures.ToArray(), TimeSpan.FromSeconds(15)); }
        catch (AggregateException ex) { System.Diagnostics.Trace.WriteLine(ex); }
        pendingProbeCaptures.Clear();
        if (!Task.WaitAll(pendingMenuCaptures.ToArray(), TimeSpan.FromSeconds(10)))
            throw new TimeoutException("Context-menu accessibility capture did not stop within ten seconds.");
        pendingMenuCaptures.Clear();
        var drained = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        if (enrichmentQueue.Writer.TryWrite(() => drained.TrySetResult()))
            drained.Task.Wait(TimeSpan.FromSeconds(10));
        RemoveDialogActionHandlers();
        if (directory is not null && !captureDiagnosticsSaved)
        {
            File.WriteAllLines(Path.Combine(directory, "capture_diagnostics.log"), captureDiagnostics.ToArray());
            captureDiagnosticsSaved = true;
        }
    }

    private nint OnMouse(int code, nint message, nint data)
    {
        if (code >= 0 && message == WM_RBUTTONUP && !capturePausedByUser && !capturePausedForIntent)
        {
            var generation = Interlocked.Read(ref chartInputGeneration);
            var foreground = GetForegroundWindow();
            hookDispatcher.BeginInvoke(new Action(() =>
            {
                if (!IsRecording || generation != Interlocked.Read(ref chartInputGeneration)) return;
                pendingMenuCaptures.RemoveAll(task => task.IsCompleted);
                pendingMenuCaptures.Add(CaptureReleasedContextMenuAsync(foreground, generation));
            }));
            return CallNextHookEx(0, code, message, data);
        }
        if (capturePausedByUser || capturePausedForIntent || code < 0 ||
            message is not (WM_LBUTTONDOWN or WM_LBUTTONUP or WM_RBUTTONDOWN or WM_MOUSEWHEEL or WM_MOUSEHWHEEL))
            return CallNextHookEx(0, code, message, data);
        var sequence = inputDeliveries.Reserve();
        var actions = new List<Action>();
        try { return OnMouseCore(code, message, data, actions.Add); }
        catch (ArgumentException ex)
        {
            System.Diagnostics.Trace.WriteLine("Mouse accessibility capture: " + ex);
            actions.Add(() => Add(new(DateTimeOffset.Now,
                "unresolved-input", null, null, null, null, null, null,
                Diagnostic: "The accessibility provider returned invalid control bounds during mouse capture: " +
                    ex.Message + " No target was guessed.")));
            return CallNextHookEx(0, code, message, data);
        }
        finally
        {
            inputDeliveries.Enqueue(sequence, () =>
            {
                foreach (var action in actions) action();
            });
            hookDispatcher.BeginInvoke(new Action(inputDeliveries.Drain));
        }
    }

    private nint OnMouseCore(int code, nint message, nint data, Action<Action> dispatch)
    {
        if (capturePausedByUser || capturePausedForIntent || code < 0)
            return CallNextHookEx(0, code, message, data);
        if (message is not (WM_LBUTTONDOWN or WM_LBUTTONUP or WM_RBUTTONDOWN or
            WM_MOUSEWHEEL or WM_MOUSEHWHEEL))
            return CallNextHookEx(0, code, message, data);
        var info = Marshal.PtrToStructure<MouseInfo>(data);
        var generationBeforeInput = Interlocked.Read(ref chartInputGeneration);
        var menuBeforeInput = pointerMenus;
        if (message is WM_LBUTTONDOWN or WM_RBUTTONDOWN or WM_MOUSEWHEEL or WM_MOUSEHWHEEL)
        {
            pointerMenus = null;
            Interlocked.Increment(ref chartInputGeneration);
        }
        var at = InputTimestamp(unchecked((uint)info.Time));
        if (chartSourceCapture is { Target.ProcessId: > 0 } pendingSource)
        {
            GetWindowThreadProcessId(GetForegroundWindow(), out var foregroundProcessId);
            if (foregroundProcessId == pendingSource.Target.ProcessId)
                Interlocked.Exchange(ref lastChartInputTick, Environment.TickCount64);
        }
        var menusAtPoint = menuBeforeInput?.Where(candidate =>
            candidate.Snapshot.InputBounds.Contains(info.Point.X, info.Point.Y)).Take(2).ToArray();
        var menu = menusAtPoint is { Length: 1 } ? menusAtPoint[0] : null;
        if (message == WM_LBUTTONDOWN && menu is not null &&
            IsPointerMenuFresh(menu, at, generationBeforeInput, GetForegroundWindow(), info.Point.X, info.Point.Y) &&
            IsWindowVisible(menu.Snapshot.WindowHandle) &&
            GetAncestor(WindowFromPoint(new(info.Point.X, info.Point.Y)), 2) == menu.Snapshot.WindowHandle)
        {
            CaptureDiagnostic("mouse-hook: menu identity frozen before click");
            dispatch(() => ProcessMouse(code, message, info, null, null,
                menu.Snapshot.Target, menu.Snapshot, at));
            return CallNextHookEx(0, code, message, data);
        }
        if (chartSourceCapture is { } chartEdit && message is WM_LBUTTONDOWN or WM_RBUTTONDOWN &&
            GetForegroundWindow() is var dialogWindow &&
            NativeControlRef(dialogWindow, null, "", "ControlType.Window", null) is
                { Process: "EXCEL", Window: "Select Data Source" } dialogOwner &&
            dialogOwner.ProcessId == chartEdit.Target.ProcessId)
        {
            Interlocked.Exchange(ref lastChartInputTick, Environment.TickCount64);
            dispatch(() =>
            {
                dragClickIndex = -1;
                dragChartName = null;
                dragLegendBefore = null;
                Add(new(at, "chart-source-edit-input",
                    dialogOwner, null, null, null, null, null,
                    Diagnostic: "Chart source dialog input; replay requires its verified final source outcome."));
            });
            return CallNextHookEx(0, code, message, data);
        }
        if (message is WM_MOUSEWHEEL or WM_MOUSEHWHEEL)
        {
            var axis = message == WM_MOUSEHWHEEL ? "Horizontal" : "Vertical";
            var target = CaptureWheelTarget(info.Point.X, info.Point.Y, axis);
            var delta = unchecked((short)((uint)info.MouseData >> 16)).ToString();
            dispatch(() =>
            {
                if (target is null)
                {
                    Add(new RecordedEvent(at, "unresolved-input", null, delta, axis, null, null, null,
                        Diagnostic: "The wheel window could not be identified at input time. " +
                            "Re-record this section; no target was guessed."));
                    return;
                }
                ProcessWheel(at, target, delta, axis, info.Point.X, info.Point.Y);
            });
            return CallNextHookEx(0, code, message, data);
        }
        var cachedChart = hoveredChart;
        if (message is WM_LBUTTONDOWN or WM_RBUTTONDOWN) chartPointerDown = true;
        if (message == WM_LBUTTONUP) chartPointerDown = false;
        if (message is WM_LBUTTONDOWN or WM_RBUTTONDOWN && cachedChart is not null &&
            at >= cachedChart.At && at - cachedChart.At < TimeSpan.FromMilliseconds(500) &&
            cachedChart.Bounds.Contains(info.Point.X, info.Point.Y) &&
            (WindowFromPoint(new(info.Point.X, info.Point.Y)) == cachedChart.Snapshot.HitWindow ||
                GetAncestor(WindowFromPoint(new(info.Point.X, info.Point.Y)), 2) == cachedChart.Snapshot.WindowHandle))
        {
            dispatch(() =>
                ProcessMouse(code, message, info, null, null, null, cachedChart.Snapshot, at));
            return CallNextHookEx(0, code, message, data);
        }
        // Resolve Excel's transient filter checklist before the general UIA
        // probes. Its provider can be slow, and running this specialized MSAA
        // capture last allowed a valid checklist click to disappear first.
        var filterTarget = message != WM_LBUTTONUP
            ? MeasureCapture("mouse-hook filter hit", () => CaptureFilterItemAtPoint(info.Point.X, info.Point.Y))
            : null;
        var mouseDownTarget = message != WM_LBUTTONUP
            ? filterTarget ??
                MeasureCapture("mouse-hook direct menu hit", () => CaptureDirectMenuAtPoint(info.Point.X, info.Point.Y)) ??
                MeasureCapture("mouse-hook owned popup command", () => CaptureOwnedPopupCommandAtPoint(info.Point.X, info.Point.Y)) ??
                MeasureCapture("mouse-hook cached dialog hit", () => CachedDialogCommandAtPoint(info.Point.X, info.Point.Y)) ??
                MeasureCapture("mouse-hook native dialog hit", () => CaptureNativeDialogCommandAtPoint(info.Point.X, info.Point.Y)) ??
                MeasureCapture("mouse-hook UIA hit", () => CaptureMouseDownTarget(info.Point.X, info.Point.Y)) : null;
        if (mouseDownTarget is { IsDialog: true, Target.ControlType: "ControlType.Button" or
                "ControlType.RadioButton" or "ControlType.CheckBox" or "ControlType.TabItem" or
                "ControlType.Edit" } && message == WM_LBUTTONDOWN)
        {
            if (mouseDownTarget.Target.ControlType == "ControlType.TabItem")
                InvalidateDialogCache(mouseDownTarget.WindowHandle);
            dispatch(() =>
                ProcessMouse(code, message, info, null, null, null, mouseDownTarget, at));
            return CallNextHookEx(0, code, message, data);
        }
        // Popup menus can disappear or switch branches before the UI-thread recorder
        // processes the queued mouse event. Freeze the MSAA identity at mouse-down,
        // with a popup-only UIA fallback; screenshots and enrichment remain queued.
        var transientMenu = message == WM_LBUTTONDOWN
            && mouseDownTarget?.Target.ControlType != "ControlType.MenuItem"
            ? MsaaCommandAtPoint(info.Point.X, info.Point.Y) : null;
        var transientMenuParent = transientMenu is { Type: "ControlType.MenuItem" }
            ? MsaaActions.MenuParentAtPoint(info.Point.X, info.Point.Y) : null;
        var transientMenuTarget = mouseDownTarget?.Target.ControlType == "ControlType.MenuItem"
            ? mouseDownTarget.Target
            : transientMenu is { Type: "ControlType.MenuItem" } command
            ? NativeControlRef(WindowFromPoint(new System.Drawing.Point(info.Point.X, info.Point.Y)),
                null, command.Name, command.Type, null) is { } menuTarget
                ? menuTarget with { ParentName = transientMenuParent } : null
            : null;
        transientMenuTarget ??= message == WM_LBUTTONDOWN
            ? CapturePopupMenuTarget(info.Point.X, info.Point.Y) : null;
        dispatch(() =>
            ProcessMouse(code, message, info, transientMenu, transientMenuParent, transientMenuTarget,
                mouseDownTarget, at));
        return CallNextHookEx(0, code, message, data);
    }

    private static ControlRef? CaptureWheelTarget(int x, int y, string axis)
    {
        var point = new System.Drawing.Point(x, y);
        // A transient native scroll/tooltip window can disappear between queries.
        // Retry the live hit test, never reuse the previous event's application.
        for (var attempt = 0; attempt < 3; attempt++)
        {
            var pointer = WindowFromPoint(point);
            if (pointer == 0) continue;
            var root = GetAncestor(pointer, 2);
            if (root == 0) root = pointer;
            var target = NativeControlRef(root, null, "", "ControlType.ScrollBar", null) ??
                NativeControlRef(pointer, null, "", "ControlType.ScrollBar", null);
            if (target is not null) return target with { Orientation = axis };
        }
        return null;
    }

    private void ProcessWheel(DateTimeOffset at, ControlRef target, string delta, string axis, int x, int y)
    {
        if (IsOwnWindow(target) || IsIgnored(target)) return;
        FinishPendingClick();
        FinishTypingCapture();
        if (scrollEventIndex >= 0 && !scrollDragPending && events[scrollEventIndex].Key == axis &&
            events[scrollEventIndex].Target?.Process == target.Process &&
            events[scrollEventIndex].Target?.Window == target.Window &&
            int.TryParse(events[scrollEventIndex].Value, out var previous))
            events[scrollEventIndex] = events[scrollEventIndex] with
                { Value = (previous + int.Parse(delta)).ToString() };
        else
        {
            CompleteScroll();
            Add(new(at, "scroll", target, delta, axis, null, null, null,
                Diagnostic: "Wheel application and axis frozen through native window identity."));
            scrollEventIndex = events.Count - 1;
            scrollElement = null;
            scrollScreen = Screen.FromPoint(new System.Drawing.Point(x, y));
            scrollAxis = axis;
        }
        scrollCapture.Stop();
        scrollCapture.Start();
    }

    private void ProcessMouse(int code, nint message, MouseInfo info,
        (string Name, string Type, int X, int Y, int Width, int Height)? transientMenu,
        string? transientMenuParent, ControlRef? transientMenuTarget,
        MouseDownSnapshot? mouseDownTarget, DateTimeOffset at)
    {
        if (message == WM_LBUTTONDOWN)
        {
            dragClickIndex = -1; dragChartName = null; dragLegendBefore = null;
        }
        // Provider latency does not invalidate a click. This callback completes
        // before Windows delivers the click to the target application, so the
        // live resolution below still observes the clicked surface. The former
        // age check converted slow but valid Excel clicks into unresolved input.
        if (code >= 0 && message == WM_LBUTTONUP && resizeStartIndex >= 0)
        {
            var release = info.Point;
            var index = resizeStartIndex;
            resizeStartIndex = -1;
            var dx = release.X - resizeStartPoint.X;
            if (Math.Abs(dx) >= 6 && index < events.Count)
            {
                clickCapture.Stop();
                if (pendingClickIndex == index) pendingClickIndex = -1;
                RecordedEvent resized;
                lock (events) events[index] = resized = events[index] with
                {
                    Kind = "resize-column", Value = dx.ToString(), Key = resizeEdge,
                    AfterState = "column-boundary-drag", Screenshot = Capture(pendingClickScreen)
                };
                Captured?.Invoke(resized);
            }
        }
        if (code >= 0 && message == WM_LBUTTONUP && scrollDragPending)
        {
            scrollDragPending = false;
            scrollCapture.Stop(); scrollCapture.Start();
        }
        if (code >= 0 && message == WM_LBUTTONUP && dragClickIndex >= 0)
        {
            var index = dragClickIndex;
            dragClickIndex = -1;
            if (index < events.Count && events[index].Kind == "click" &&
                (Math.Abs(info.Point.X - dragClickPoint.X) >= 6 || Math.Abs(info.Point.Y - dragClickPoint.Y) >= 6))
            {
                var recorded = events[index];
                clickCapture.Stop();
                if (pendingClickIndex == index) pendingClickIndex = -1;
                var chartName = dragChartName;
                var chartSheet = dragChartSheet;
                var chartWindow = dragChartWindow;
                var legendBefore = dragLegendBefore;
                var outcomeScreen = pendingClickScreen;
                FinishLegendOutcome();
                chartSamplePaused = true;
                pendingLegendOutcome = () =>
                {
                try
                {
                    if (chartName is null || chartSheet is null || legendBefore is null || recorded.Target is null)
                        throw new InvalidOperationException("No supported semantic resize outcome was captured for this drag.");
                    var after = ExcelNativeSheet.Read<ExcelNativeSheet.LegendLayout>(chartWindow, sheet =>
                    {
                        if ((string)sheet.Name != chartSheet)
                            throw new InvalidOperationException("The chart sheet changed during drag capture.");
                        return ExcelNativeSheet.ReadLegendLayout(sheet, chartName);
                    });
                    if (Math.Abs(after.Width - legendBefore.Width) < .005 &&
                        Math.Abs(after.Height - legendBefore.Height) < .005 &&
                        Math.Abs(after.Left - legendBefore.Left) < .005 &&
                        Math.Abs(after.Top - legendBefore.Top) < .005)
                        throw new InvalidOperationException("The drag's legend layout change could not be verified. " +
                            $"Before: {legendBefore}; after: {after}.");
                    recorded = recorded with { Kind = "set-chart-legend-layout",
                        Target = recorded.Target with { Name = chartName, ControlType = "ControlType.Image", AutomationId = "" },
                        Value = System.Text.Json.JsonSerializer.Serialize(after),
                        BeforeState = System.Text.Json.JsonSerializer.Serialize(legendBefore),
                        AfterState = "chart-legend-sheet:" + chartSheet, Screenshot = Capture(outcomeScreen),
                        Diagnostic = "Legend layout captured before and after drag; replay uses chart-relative dimensions, not pointer coordinates." };
                }
                catch (Exception ex) when (ex is InvalidOperationException or COMException)
                {
                    recorded = recorded with { Kind = "unresolved-input", Diagnostic = ex.Message,
                        AfterState = null, Screenshot = Capture(outcomeScreen) };
                }
                lock (events) events[index] = recorded;
                Captured?.Invoke(recorded);
                };
                legendOutcomeCapture.Start();
            }
            dragChartName = null; dragLegendBefore = null;
        }
        if (code >= 0 && (message == WM_LBUTTONDOWN || message == WM_RBUTTONDOWN ||
            message == WM_MOUSEWHEEL || message == WM_MOUSEHWHEEL))
        {
            var precedingRowMenu = message == WM_LBUTTONDOWN && events.Count > 0 &&
                events[^1].Kind == "context-click" &&
                events[^1].Target?.ClassName == "XLGridRowHeader" &&
                at - events[^1].At < TimeSpan.FromSeconds(8);
            if (precedingRowMenu)
            {
                // A popup menu can vanish on mouse-up. Finishing the previous
                // click here takes a screenshot and calls UIA while the next
                // mouse-down is still waiting in this low-level hook.
                clickCapture.Stop();
                pendingClickIndex = -1;
                pendingClickScreen = null;
                pendingClickPoint = null;
                pendingClickWindow = 0;
                pendingClickTabBefore = null;
                pendingClickFocusBefore = null;
            }
            else FinishPendingClick();
            FinishTypingCapture();
            cachedEditableTarget = null; cachedEditableElement = null;
            cachedKeyboardTarget = null; cachedKeyboardScreen = null; // Pointer may change focus.
            try
            {
                if (message == WM_LBUTTONDOWN && transientMenuTarget is not null)
                {
                    // Use the identity frozen at mouse-down. Re-hit-testing here
                    // can see another item after the popup closes or changes branch.
                    if (!IsOwnWindow(transientMenuTarget) && !IsIgnored(transientMenuTarget) &&
                        !IsIgnoredForeground())
                        Add(new RecordedEvent(at, "click", transientMenuTarget,
                            null, null, Capture(Screen.FromPoint(
                                new System.Drawing.Point(info.Point.X, info.Point.Y))),
                            null, "msaa-menu-item", ClickX: info.Point.X, ClickY: info.Point.Y,
                            Diagnostic: "Menu identity captured synchronously at mouse-down."));
                    return;
                }
                if (message == WM_LBUTTONDOWN && mouseDownTarget is
                    { IsDialog: true, Target.ControlType: "ControlType.Button" or
                        "ControlType.RadioButton" or "ControlType.CheckBox" or "ControlType.TabItem" or
                        "ControlType.Edit" } dialogClick)
                {
                    var frozen = dialogClick.Target;
                    if (IsOwnWindow(frozen) || IsIgnored(frozen) || IsIgnoredForeground()) return;
                    Add(new RecordedEvent(at, "click", frozen, null, null,
                        Capture(Screen.FromPoint(new System.Drawing.Point(info.Point.X, info.Point.Y))),
                        null, null, ClickX: info.Point.X, ClickY: info.Point.Y,
                        Diagnostic: "Control identity verified at mouse-down."));
                    var capturedIndex = events.Count - 1;
                    if (frozen.Window == "Find and Replace")
                        QueueDialogFieldSnapshot(capturedIndex, dialogClick.WindowHandle);
                    if (frozen.ControlType == "ControlType.TabItem")
                        RefreshDialogCacheAfterTab(dialogClick.WindowHandle);
                    if (frozen.ControlType == "ControlType.Edit")
                    {
                        cachedEditableTarget = frozen;
                        cachedEditableElement = dialogClick.Element;
                        cachedEditableScreen = Screen.FromHandle(dialogClick.WindowHandle);
                        cachedKeyboardWindow = dialogClick.WindowHandle;
                        cachedKeyboardFocus = 0;
                    }
                    if (frozen.ControlType == "ControlType.Button")
                        QueueNativeCloseConfirmation(capturedIndex, dialogClick.WindowHandle);
                    return;
                }
                if (precedingRowMenu)
                {
                    var rowTarget = events.LastOrDefault(item => item.Kind == "context-click" &&
                        item.Target?.ClassName == "XLGridRowHeader")?.Target;
                    if (rowTarget is not null && !IsIgnored(rowTarget))
                    {
                        if (RowMenuCommandAtPoint(info.Point.X, info.Point.Y) is
                            { Type: "ControlType.MenuItem" } menu)
                        {
                            var menuTarget = rowTarget with { Name = menu.Name, ControlType = menu.Type,
                                AutomationId = null, ClassName = null, ParentName = null };
                            Add(new RecordedEvent(at, "click", menuTarget, null, null,
                                null, null, "msaa-menu-item", ClickX: info.Point.X, ClickY: info.Point.Y));
                        }
                        else
                        {
                            var unresolved = new RecordedEvent(at, "unresolved-click",
                                rowTarget with { Name = "row menu choice", ControlType = "ControlType.Window",
                                    AutomationId = null, ClassName = null, ParentName = null },
                                null, null, null, null, "menu choice was not identified",
                                ClickX: info.Point.X, ClickY: info.Point.Y,
                                Diagnostic: "row-menu-hit:pointer=" + DescribeNativeWindow(
                                    WindowFromPoint(new System.Drawing.Point(info.Point.X, info.Point.Y))) +
                                    ";foreground=" + DescribeNativeWindow(GetForegroundWindow()) +
                                    ";menu-owner=" + DescribeNativeWindow(CurrentMenuOwner()));
                            Add(unresolved);
                            int unresolvedIndex;
                            lock (events) unresolvedIndex = events.FindIndex(item => ReferenceEquals(item, unresolved));
                            UnresolvedDialogClick?.Invoke(unresolvedIndex,
                                rowTarget.Window ?? rowTarget.Name ?? "context menu");
                            QueueRowMenuDiagnostic(unresolvedIndex, info.Point.X, info.Point.Y, unresolved.At);
                            var key = rowTarget.Window + "/" + rowTarget.Name;
                            if (rowAnchorCaptures.TryGetValue(key, out var cached) &&
                                unresolved.At - cached.At < TimeSpan.FromSeconds(10) &&
                                cached.Capture.IsCompletedSuccessfully &&
                                cached.Capture.Result is { } anchor)
                            {
                                lock (events)
                                {
                                    var contextIndex = events.FindLastIndex(item => item.Kind == "context-click" &&
                                        item.Target?.ClassName == "XLGridRowHeader");
                                    if (contextIndex >= 0)
                                        events[contextIndex] = events[contextIndex] with
                                        { BeforeState = "row-anchor:" + JsonSerializer.Serialize(anchor) };
                                }
                                QueueRowDeletionOutcome(unresolvedIndex, unresolved.At, rowTarget, anchor);
                            }
                        }
                    }
                    return;
                }
                if (!precedingRowMenu) PrimeDialogCache(CompactWindowAtPoint(info.Point.X, info.Point.Y));
                if (mouseDownTarget is null && message == WM_LBUTTONDOWN &&
                    TryRecordNativeDialogClick(info.Point.X, info.Point.Y))
                    return;
                var native = message == WM_LBUTTONDOWN
                    ? FindNativeActionableAtPoint(info.Point.X, info.Point.Y) : null;
                var compactWindow = CompactWindowAtPoint(info.Point.X, info.Point.Y);
                // Hook points and UIA FromPoint both use physical screen coordinates.
                var automationPoint = new System.Drawing.Point(info.Point.X, info.Point.Y);
                var hit = mouseDownTarget?.Element ?? (native is null
                    ? AutomationElement.FromPoint(new System.Windows.Point(automationPoint.X, automationPoint.Y))
                    : AutomationElement.FromHandle(native.Value.Handle));
                if (hit is null) return;
                var compactClick = message == WM_LBUTTONDOWN &&
                    IsCompactForegroundWindow(info.Point.X, info.Point.Y);
                var element = mouseDownTarget?.Element ?? (compactClick
                    ? FindActionableControlAtPoint(hit, automationPoint.X, automationPoint.Y) ?? hit : hit);
                var target = mouseDownTarget?.Target ?? Automation.Describe(element);
                if (message == WM_LBUTTONDOWN && target is not null &&
                    transientMenu is { Type: "ControlType.MenuItem" } liveMenu)
                    target = target with
                    {
                        Name = liveMenu.Name,
                        ControlType = liveMenu.Type,
                        AutomationId = null,
                        ClassName = null,
                        ParentName = transientMenuParent ?? Automation.MenuParentName(element)
                    };
                if (mouseDownTarget is null && compactClick &&
                    target?.ControlType is "ControlType.DataItem" or "ControlType.ListItem")
                    target = CorrectCompactListWindow(target, info.Point.X, info.Point.Y);
                string? desktopHitDiagnostic = null;
                if (IsDesktopBackground(target))
                {
                    var pointerWindow = WindowFromPoint(new System.Drawing.Point(info.Point.X, info.Point.Y));
                    var pointerRoot = pointerWindow == 0 ? 0 : GetAncestor(pointerWindow, 2);
                    var foreground = GetForegroundWindow();
                    desktopHitDiagnostic = $"desktop-hit:point={info.Point.X},{info.Point.Y};" +
                        $"window={DescribeNativeWindow(pointerWindow)};" +
                        $"root={DescribeNativeWindow(pointerRoot)};" +
                        $"foreground={DescribeNativeWindow(foreground)}";
                    // UIA can briefly return Program Manager during an application
                    // switch. Trust a native window at the pointer only when it
                    // belongs to another process; never assign the foreground
                    // window's controls to a real desktop click.
                    AutomationElement? nativeElement = null;
                    var owner = DescribeNativeWindowOwner(pointerWindow);
                    if (owner is not null &&
                        !owner.Equals("explorer", StringComparison.OrdinalIgnoreCase) &&
                        owner.Equals(DescribeNativeWindowOwner(foreground), StringComparison.OrdinalIgnoreCase))
                        try { nativeElement = AutomationElement.FromHandle(pointerWindow); }
                        catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                    if (nativeElement is not null)
                    {
                        var nativeTarget = Automation.Describe(nativeElement);
                        if (nativeTarget is not null && !IsWeakSwitchTarget(nativeTarget) &&
                            !string.IsNullOrWhiteSpace(nativeTarget.Name) &&
                            string.Equals(nativeTarget.Process, owner, StringComparison.OrdinalIgnoreCase) &&
                            !IsOwnWindow(nativeTarget) && !IsIgnored(nativeTarget))
                        {
                            element = nativeElement;
                            target = nativeTarget;
                        }
                    }
                }
                var foregroundProcess = DescribeNativeWindowOwner(GetForegroundWindow());
                var mismatchedWeakTarget = IsWeakSwitchTarget(target) &&
                    !string.IsNullOrWhiteSpace(foregroundProcess) &&
                    !string.Equals(target?.Process, foregroundProcess, StringComparison.OrdinalIgnoreCase);
                var searchForegroundHandle = GetForegroundWindow();
                var searchForegroundMismatch = message == WM_LBUTTONDOWN &&
                    string.Equals(foregroundProcess, "SearchHost", StringComparison.OrdinalIgnoreCase) &&
                    target is not null &&
                    !string.Equals(target.Process, "SearchHost", StringComparison.OrdinalIgnoreCase) &&
                    GetWindowRect(searchForegroundHandle, out var searchBounds) &&
                    info.Point.X >= searchBounds.Left && info.Point.X < searchBounds.Right &&
                    info.Point.Y >= searchBounds.Top && info.Point.Y < searchBounds.Bottom;
                if (mismatchedWeakTarget)
                    desktopHitDiagnostic = $"{desktopHitDiagnostic};process-mismatch:" +
                        $"target={target?.Process};foreground={foregroundProcess};" +
                        $"point={info.Point.X},{info.Point.Y};" +
                        $"window={DescribeNativeWindow(WindowFromPoint(new System.Drawing.Point(info.Point.X, info.Point.Y)))}";
                if (searchForegroundMismatch)
                {
                    desktopHitDiagnostic = (desktopHitDiagnostic ?? "") + ";windows-search-underlay:" +
                        $"target={target!.Process}/{target.Name};foreground={DescribeNativeWindow(searchForegroundHandle)}";
                    target = new ControlRef("SearchHost", "Search", null, "Windows Search result",
                        "ControlType.Window", null, "Search");
                }
                if (target is { Process: "EXCEL", ClassName: "XLSpreadsheetCell", AutomationId: { } cellId } &&
                    cellId.Length > 1 && cellId[0] == 'A' && int.TryParse(cellId[1..], out var rowNumber))
                {
                    try
                    {
                        var cellBounds = element.Current.BoundingRectangle;
                        if (!cellBounds.IsEmpty && info.Point.X < cellBounds.Left &&
                            info.Point.X >= cellBounds.Left - 60)
                        {
                            var header = AutomationElement.FromPoint(new System.Windows.Point(
                                cellBounds.Left - 6, info.Point.Y));
                            var headerTarget = header is null ? null : Automation.Describe(header);
                            if (headerTarget is { ClassName: "XLGridRowHeader" } &&
                                headerTarget.Name == rowNumber.ToString() &&
                                headerTarget.Process == target.Process)
                            {
                                element = header!;
                                target = headerTarget;
                            }
                        }
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                }
                var cachedAction = compactClick && native is null &&
                    target?.ControlType is "ControlType.Window" or "ControlType.Pane" or "ControlType.Custom"
                    ? CachedDialogControlAtPoint(info.Point.X, info.Point.Y) : null;
                if (mouseDownTarget is null && native is { Name: { Length: > 0 } } button && target is not null &&
                    (target.ControlType is not ("ControlType.Button" or "ControlType.TabItem") ||
                     !string.Equals(target.Name, button.Name, StringComparison.OrdinalIgnoreCase)))
                    target = target with { Name = button.Name,
                        ControlType = button.ClassName.Equals("Button", StringComparison.OrdinalIgnoreCase)
                            ? NativeButtonControlType(button.Handle) : "ControlType.TabItem" };
                else if (mouseDownTarget is null && cachedAction is { } action && target is not null &&
                    target.ControlType is "ControlType.Window" or "ControlType.Pane" or "ControlType.Custom")
                    target = target with { Name = action.Name, ControlType = action.Type,
                        AutomationId = null, ClassName = null,
                        ParentName = action.Type == "ControlType.MenuItem"
                            ? transientMenuParent ?? Automation.MenuParentName(element) : null };
                nint dialogRoot = 0;
                if (target is not null && compactClick)
                {
                    try
                    {
                        var targetWindow = WindowAncestor(element);
                        var candidate = targetWindow is null ? 0 : (nint)targetWindow.Current.NativeWindowHandle;
                        dialogRoot = ResolveDialogWindow(candidate, target, null, null);
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                    if (dialogRoot == 0)
                        dialogRoot = ResolveDialogWindow(CompactWindowAtPoint(info.Point.X, info.Point.Y),
                            target, null, null);
                }
                if (message == WM_LBUTTONDOWN && target is not null && dialogRoot != 0 &&
                    target.ControlType is "ControlType.Window" or "ControlType.Pane" or
                        "ControlType.Custom" or "ControlType.Tab")
                {
                    // Identify transient native and MSAA commands before the click changes
                    // or closes the dialog. The point is recording evidence only; replay
                    // receives the resulting accessibility identity and never these coordinates.
                    var rawPoint = new System.Drawing.Point(info.Point.X, info.Point.Y);
                    var uiaAction = FindDialogActionableAtPoint(dialogRoot, rawPoint.X, rawPoint.Y);
                    var nativeAction = uiaAction is null ? FindNativeActionableAtPoint(dialogRoot,
                        rawPoint.X, rawPoint.Y, 0x20, 80) : null;
                    var accessibleAction = uiaAction is null && nativeAction is null
                        ? MsaaActions.CommandAtPoint(dialogRoot, rawPoint.X, rawPoint.Y)
                        : null;
                    if (uiaAction is not null || nativeAction is not null || accessibleAction is not null)
                    {
                        var actionType = nativeAction is not null
                            ? nativeAction.Value.ClassName.Contains("Tab",
                                StringComparison.OrdinalIgnoreCase)
                                ? "ControlType.TabItem" : "ControlType.Button"
                            : accessibleAction?.Type;
                        var actionName = nativeAction?.Name ?? accessibleAction?.Name;
                        target = uiaAction ??
                            (actionName is not null && actionType is not null
                                ? NativeControlRef(dialogRoot, null, actionName, actionType, null)
                                : null) ?? target;
                    }
                }
                if (message == WM_LBUTTONDOWN && target is not null &&
                    NativeCloseAtPoint(CompactWindowAtPoint(info.Point.X, info.Point.Y), info.Point.X, info.Point.Y))
                    target = target with { Name = "Close", ControlType = "ControlType.Button" };
                if (target is not null && !IsOwnWindow(target) && !IsIgnored(target) && !IsIgnoredForeground())
                {
                    if (message == WM_RBUTTONDOWN && target.Process == "EXCEL" &&
                        target.ControlType == "ControlType.Image" && !string.IsNullOrWhiteSpace(target.Name) &&
                        (target.ClassName == "ExcelChartObject" ||
                            target.ParentName?.StartsWith("Sheet ", StringComparison.Ordinal) == true))
                    {
                        var chartSheet = mouseDownTarget?.ChartSheet ??
                            (target.ParentName?.StartsWith("Sheet ", StringComparison.Ordinal) == true
                                ? target.ParentName["Sheet ".Length..] : null);
                        target = target with { ClassName = "ExcelChartObject" };
                        chartSourceCapture = new(mouseDownTarget?.WindowHandle ??
                            GetAncestor(WindowFromPoint(new(info.Point.X, info.Point.Y)), 2),
                            target, chartSheet, mouseDownTarget?.ChartSource);
                    }
                    var screen = Screen.FromPoint(new System.Drawing.Point(info.Point.X, info.Point.Y));
                    if (!searchForegroundMismatch && IsEditableTarget(target, element))
                    {
                        cachedEditableTarget = target;
                        cachedEditableElement = element;
                        cachedEditableScreen = screen;
                        cachedKeyboardWindow = GetForegroundWindow();
                        cachedKeyboardFocus = 0;
                    }
                    if (message is WM_MOUSEWHEEL or WM_MOUSEHWHEEL ||
                        target.ControlType is "ControlType.ScrollBar" or "ControlType.Thumb")
                    {
                        var axis = message == WM_MOUSEHWHEEL ? "Horizontal" :
                            message == WM_MOUSEWHEEL ? "Vertical" : Automation.ScrollBarAxis(element);
                        var delta = message is WM_MOUSEWHEEL or WM_MOUSEHWHEEL
                            ? unchecked((short)((uint)info.MouseData >> 16)).ToString() : null;
                        if (delta is not null && scrollEventIndex >= 0 && !scrollDragPending &&
                            events[scrollEventIndex].Key == axis && events[scrollEventIndex].Target?.Process == target.Process &&
                            events[scrollEventIndex].Target?.Window == target.Window &&
                            int.TryParse(events[scrollEventIndex].Value, out var previous))
                        {
                            events[scrollEventIndex] = events[scrollEventIndex] with
                            { Value = (previous + int.Parse(delta)).ToString() };
                        }
                        else
                        {
                            CompleteScroll();
                            Add(new(at, "scroll", target, delta, axis, null,
                                Automation.ScrollState(element), null));
                            scrollEventIndex = events.Count - 1;
                            scrollElement = element;
                            scrollScreen = screen;
                            scrollAxis = axis;
                        }
                        scrollCapture.Stop();
                        scrollDragPending = delta is null && message == WM_LBUTTONDOWN;
                        if (!scrollDragPending) scrollCapture.Start();
                    }
                    else
                    {
                        var before = searchForegroundMismatch || target.Window == "Find and Replace"
                            ? null : Automation.State(element);
                        Add(new(at, searchForegroundMismatch ? "unresolved-click" :
                            message == WM_RBUTTONDOWN ? "context-click" : "click",
                            target, null, null, null, before,
                            message == WM_RBUTTONDOWN ? SelectionCount(element) : null,
                            ClickX: info.Point.X, ClickY: info.Point.Y,
                            Diagnostic: mouseDownTarget is not null &&
                                !IsWeakClickTarget(target)
                                ? "Control identity verified at mouse-down." : desktopHitDiagnostic));
                        if (message == WM_LBUTTONDOWN)
                        {
                            dragClickPoint = new(info.Point.X, info.Point.Y);
                            dragChartName = null; dragLegendBefore = null;
                            if (target.Process == "EXCEL")
                            {
                                try
                                {
                                    if (target.ClassName == "ExcelChartObject")
                                        dragChartName = target.Name;
                                    var ancestor = element;
                                    for (var depth = 0; dragChartName is null && ancestor is not null && depth < 10; depth++)
                                    {
                                        if (ancestor.Current.ControlType == System.Windows.Automation.ControlType.Image &&
                                            !string.IsNullOrWhiteSpace(ancestor.Current.Name))
                                        {
                                            dragChartName = ancestor.Current.Name;
                                            break;
                                        }
                                        ancestor = TreeWalker.RawViewWalker.GetParent(ancestor);
                                    }
                                    if (dragChartName is not null)
                                    {
                                        dragClickIndex = events.Count - 1;
                                        var root = GetAncestor(WindowFromPoint(dragClickPoint), 2);
                                        dragChartWindow = root;
                                        dragLegendBefore = mouseDownTarget?.LegendBefore;
                                        dragChartSheet = mouseDownTarget?.ChartSheet;
                                        if (dragLegendBefore is null)
                                        {
                                            // The known-working recorder sampled this state directly at
                                            // mouse-down. Keep the hover cache as the fast path, but never
                                            // discard a valid drag merely because that cache was late.
                                            dragLegendBefore = ExcelNativeSheet.Read<ExcelNativeSheet.LegendLayout>(root,
                                                sheet =>
                                                {
                                                    dragChartSheet = (string)sheet.Name;
                                                    return ExcelNativeSheet.ReadLegendLayout(sheet, dragChartName);
                                                });
                                        }
                                    }
                                }
                                catch (Exception ex) when (ex is InvalidOperationException or COMException)
                                { System.Diagnostics.Trace.WriteLine(ex); dragLegendBefore = null; }
                            }
                        }
                        if (searchForegroundMismatch)
                        {
                            QueueSearchResultRetry(events.Count - 1, info.Point.X, info.Point.Y);
                            QueueSearchLaunchOutcome(events.Count - 1, searchForegroundHandle);
                        }
                        if (message == WM_LBUTTONDOWN && target is
                            { Process: "EXCEL", ClassName: "XLGridRowHeader" } &&
                            int.TryParse(target.Name, out var selectedRow))
                        {
                            var capture = Task.Run(() => TryReadRowAnchor(target, selectedRow));
                            rowAnchorCaptures[target.Window + "/" + target.Name] = (DateTimeOffset.Now, capture);
                            pendingProbeCaptures.RemoveAll(task => task.IsCompleted);
                            pendingProbeCaptures.Add(capture);
                        }
                        if (IsDesktopBackground(target) || mismatchedWeakTarget)
                            QueueApplicationSwitchRetry(events.Count - 1, info.Point.X, info.Point.Y,
                                foregroundProcess);
                        if (message == WM_LBUTTONDOWN && compactClick &&
                            target.Window == "Find and Replace")
                            QueueDialogFieldSnapshot(events.Count - 1, dialogRoot);
                        if (message == WM_LBUTTONDOWN && compactClick &&
                            target.ControlType is "ControlType.Window" or "ControlType.Pane" or
                                "ControlType.Custom" or "ControlType.Tab" &&
                            target.ClassName != "EDTBX")
                        {
                            QueueDialogCommandRetry(events.Count - 1,
                                dialogRoot,
                                info.Point.X, info.Point.Y);
                            QueueDialogClickProbe(events.Count - 1, dialogRoot,
                                info.Point.X, info.Point.Y);
                        }
                        if (message == WM_LBUTTONDOWN && compactClick &&
                            target.ControlType == "ControlType.TabItem")
                            RefreshDialogCacheAfterTab(dialogRoot);
                        if (IsEditableTarget(target, element))
                        {
                            var fieldIndex = events.Count - 1;
                            var fieldElement = element;
                            pendingFieldCaptures.RemoveAll(task => task.IsCompleted);
                            pendingFieldCaptures.Add(Task.Run(() =>
                            {
                                try
                                {
                                    var label = Automation.FieldLabel(fieldElement);
                                    if (string.IsNullOrWhiteSpace(label)) return;
                                    lock (events)
                                        events[fieldIndex] = events[fieldIndex] with
                                        { Target = events[fieldIndex].Target! with { Label = label } };
                                }
                                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
                            }));
                        }
                        if (message == WM_LBUTTONDOWN && target.Name is { Length: > 0 } &&
                            target.ClassName == "XLGridColumnHeader")
                        {
                            var bounds = element.Current.BoundingRectangle;
                            if (!bounds.IsEmpty && bounds.Height is >= 8 and <= 60 &&
                                (Math.Abs(info.Point.X - bounds.Right) <= 9 ||
                                 Math.Abs(info.Point.X - bounds.Left) <= 9))
                            {
                                resizeStartIndex = events.Count - 1;
                                resizeStartPoint = new(info.Point.X, info.Point.Y);
                                resizeEdge = Math.Abs(info.Point.X - bounds.Right) <=
                                    Math.Abs(info.Point.X - bounds.Left) ? "right" : "left";
                            }
                        }
                        if (message == WM_LBUTTONDOWN && target.ControlType == "ControlType.MenuItem")
                        {
                            // The menu closes on mouse-up; capture the actual chosen item while it is visible.
                            var index = events.Count - 1;
                            if (target.Window != "Find and Replace")
                                events[index] = events[index] with { Screenshot = Capture(screen) };
                        }
                        else
                        {
                            pendingClickIndex = events.Count - 1;
                            pendingClickScreen = screen;
                            pendingClickPoint = new System.Drawing.Point(info.Point.X, info.Point.Y);
                            var clickedWindow = target.ControlType is "ControlType.Window" or "ControlType.TitleBar" ||
                                target.Name == "Close" || target.Window == "Find and Replace"
                                ? WindowAncestor(element) : null;
                            pendingClickWindow = clickedWindow is null ? 0 :
                                (nint)clickedWindow.Current.NativeWindowHandle;
                            pendingClickTabBefore = clickedWindow is null ||
                                target.ControlType != "ControlType.TabItem"
                                ? null : SelectedTab(clickedWindow,
                                    sheetOnly: target.AutomationId == "SheetTab")?.Current.Name;
                            pendingClickFocusBefore = target.ControlType == "ControlType.DataItem"
                                ? Automation.Describe(AutomationElement.FocusedElement) : null;
                            if (IsEditableTarget(target, element))
                            {
                                pendingClickIndex = -1;
                                pendingClickWindow = 0;
                                pendingClickTabBefore = null;
                            }
                            else { clickCapture.Stop(); clickCapture.Start(); }
                        }
                        if (message == WM_LBUTTONDOWN && target.ControlType == "ControlType.HeaderItem")
                        {
                            var index = events.Count - 1;
                            var priorDirection = Executor.ObservedSortDirection(element);
                            pendingScrollCaptures.Add(Task.Run(async () =>
                            {
                                try
                                {
                                    var deadline = DateTime.UtcNow.AddSeconds(4);
                                    string? last = null;
                                    var stable = 0;
                                    while (DateTime.UtcNow < deadline)
                                    {
                                        await Task.Delay(150);
                                        var liveHeader = Automation.Resolve(target);
                                        var direction = liveHeader is null ? null : Executor.ObservedSortDirection(liveHeader);
                                        stable = direction is not null && direction == last ? stable + 1 : 0;
                                        last = direction;
                                        if (stable >= 3 && direction != priorDirection) break;
                                    }
                                    if (last is not null && stable >= 3)
                                        lock (events) events[index] = events[index] with { AfterState = "sort:" + last };
                                }
                                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
                            }));
                        }
                    }
                }
            }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
            finally { if (message == WM_RBUTTONDOWN) chartPointerDown = false; }
        }
        return;
    }

    private nint OnKeyboard(int code, nint message, nint data)
    {
        if (capturePausedByUser || capturePausedForIntent || code < 0 ||
            message is not (WM_KEYDOWN or WM_SYSKEYDOWN))
            return CallNextHookEx(0, code, message, data);
        var sequence = inputDeliveries.Reserve();
        var actions = new List<Action>();
        try { return OnKeyboardCore(code, message, data, actions.Add); }
        finally
        {
            inputDeliveries.Enqueue(sequence, () =>
            {
                foreach (var action in actions) action();
            });
            hookDispatcher.BeginInvoke(new Action(inputDeliveries.Drain));
        }
    }

    private nint OnKeyboardCore(int code, nint message, nint data, Action<Action> dispatch)
    {
        if (capturePausedByUser || capturePausedForIntent || code < 0)
            return CallNextHookEx(0, code, message, data);
        if (message != WM_KEYDOWN && message != WM_SYSKEYDOWN)
            return CallNextHookEx(0, code, message, data);
        var info = Marshal.PtrToStructure<KeyboardInfo>(data);
        if ((message == WM_KEYDOWN || message == WM_SYSKEYDOWN) &&
            info.VirtualKey == (uint)Keys.I && KeyDown(Keys.ControlKey) &&
            KeyDown(Keys.Menu) && KeyDown(Keys.ShiftKey))
        {
            capturePausedForIntent = true;
            var previousWindow = GetForegroundWindow();
            dispatch(() => IntentRequested?.Invoke(previousWindow));
            return 1;
        }
        if (info.VirtualKey is (uint)Keys.ShiftKey or (uint)Keys.LShiftKey or (uint)Keys.RShiftKey or
            (uint)Keys.ControlKey or (uint)Keys.LControlKey or (uint)Keys.RControlKey or
            (uint)Keys.Menu or (uint)Keys.LMenu or (uint)Keys.RMenu)
            return CallNextHookEx(0, code, message, data);
        var owner = DescribeNativeWindowOwner(GetForegroundWindow());
        if (owner is not null && (ignoredProcesses.Contains(owner) ||
            owner.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), StringComparison.OrdinalIgnoreCase)))
            return CallNextHookEx(0, code, message, data);
        Interlocked.Increment(ref chartInputGeneration);
        var modifiers = CurrentModifiers();
        var at = InputTimestamp(info.Time);
        if (chartSourceCapture is { Target.ProcessId: > 0 } pendingSource)
        {
            GetWindowThreadProcessId(GetForegroundWindow(), out var foregroundProcessId);
            if (foregroundProcessId == pendingSource.Target.ProcessId)
                Interlocked.Exchange(ref lastChartInputTick, Environment.TickCount64);
        }
        // Low-level hook timestamps can already be old when a provider held the
        // hook thread. Windows has not delivered this key to the application yet,
        // so age alone must not discard its current foreground/focus identity.
        var inputOwner = NativeControlRef(GetForegroundWindow(), null, "", "ControlType.Window", null);
        if (chartSourceCapture is not { } pendingChart ||
            inputOwner is not { Process: "EXCEL", Window: "Select Data Source" } ||
            inputOwner.ProcessId != pendingChart.Target.ProcessId)
            inputOwner = null;
        if (inputOwner is not null)
        {
            var dialogOwner = inputOwner;
            Interlocked.Exchange(ref lastChartInputTick, Environment.TickCount64);
            dispatch(() => Add(new(at, "chart-source-edit-input",
                dialogOwner, null, null, null, null, null,
                Diagnostic: "Chart source dialog input; replay requires its verified final source outcome.")));
            return CallNextHookEx(0, code, message, data);
        }
        var snapshot = inputOwner is null ? CaptureKeyboardTarget() : null;
        dispatch(() => ProcessKeyboard(message, info, modifiers, at, snapshot, inputOwner));
        return CallNextHookEx(0, code, message, data);
    }

    private static DateTimeOffset InputTimestamp(uint time) =>
        DateTimeOffset.Now - TimeSpan.FromMilliseconds(unchecked((uint)Environment.TickCount64 - time));

    private void RecordDelayedInput(DateTimeOffset at, string input, ControlRef? owner = null) =>
        Add(new RecordedEvent(at, "unresolved-input", owner, null, null, null, null, null,
            Diagnostic: $"The {input} input identity could not be frozen before delayed processing. Re-record this section; no target was guessed."));

    private sealed record KeyboardSnapshot(nint Window, nint Focus, ControlRef? Target,
        AutomationElement? Element, Screen Screen);

    private static KeyboardSnapshot? CaptureKeyboardTarget()
    {
        var window = GetForegroundWindow();
        var focus = ForegroundFocusHandle();
        try
        {
            if (TryDescribeNativeEdit(window, focus, out var native) && native is not null)
                return new(window, focus, (GetWindowLong(focus, -16) & 0x20) != 0 ? null : native,
                    null, Screen.FromHandle(window));
            var element = AutomationElement.FocusedElement;
            if (element is null) return null;
            if (element.Current.IsPassword)
                return new(window, focus, null, null, Screen.FromHandle(window));
            var current = element.Current;
            GetWindowThreadProcessId(window, out var owner);
            if (current.ProcessId != owner && !IsHostedElementInWindow(element, window)) return null;
            var target = NativeControlRef(window, current.AutomationId, current.Name,
                current.ControlType.ProgrammaticName, current.ClassName);
            if (target is null) return null;
            target = target with { Process = System.Diagnostics.Process.GetProcessById(current.ProcessId).ProcessName,
                ParentName = TreeWalker.ControlViewWalker.GetParent(element)?.Current.Name,
                ProcessId = current.ProcessId };
            return new(window, focus, target, element, Screen.FromHandle(window));
        }

        catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    private static bool IsHostedElementInWindow(AutomationElement element, nint window)
    {
        for (var depth = 0; depth < 32; depth++)
        {
            var handle = (nint)element.Current.NativeWindowHandle;
            if (handle != 0 && (handle == window || GetAncestor(handle, 2) == window)) return true;
            var parent = TreeWalker.RawViewWalker.GetParent(element);
            if (parent is null) return false;
            element = parent;
        }
        return false;
    }

    private void ProcessKeyboard(nint message, KeyboardInfo info, string modifiers,
        DateTimeOffset at, KeyboardSnapshot? snapshot, ControlRef? inputOwner)
    {
        if (message == WM_KEYDOWN || message == WM_SYSKEYDOWN)
        {
            try
            {
                var keyCode = (int)info.VirtualKey;
                if (snapshot is null)
                {
                    RecordDelayedInput(at, "keyboard", inputOwner);
                    return;
                }
                if (snapshot.Target is null) return;
                if (keyCode is (int)Keys.Tab or (int)Keys.F6)
                    FinishTypingCapture();
                var foreground = snapshot.Window;
                var focus = snapshot.Focus;
                // Resolve the focused editor once after a click or focus change.
                // The control under the pointer can be the dialog pane instead of
                // the field receiving the keystrokes.
                if (cachedEditableTarget is not null && foreground == cachedKeyboardWindow &&
                    focus != 0 && focus != cachedKeyboardFocus)
                {
                    cachedEditableTarget = null;
                    cachedEditableElement = null;
                    cachedKeyboardTarget = null;
                }
                if (keyCode is (int)Keys.Tab or (int)Keys.F6)
                {
                    cachedEditableTarget = null;
                    cachedEditableElement = null; // Keyboard navigation can change fields without a click.
                    cachedKeyboardTarget = null;
                }
                var target = snapshot.Target;
                var targetScreen = snapshot.Screen;
                cachedKeyboardFocus = focus;
                cachedKeyboardWindow = foreground;
                cachedKeyboardTarget = target;
                cachedKeyboardScreen = targetScreen;
                cachedEditableTarget = IsEditableTarget(target, snapshot.Element) ? target : null;
                cachedEditableElement = cachedEditableTarget is null ? null : snapshot.Element;
                cachedEditableScreen = cachedEditableTarget is null ? null : targetScreen;
                if (!IsOwnWindow(target) && !IsIgnored(target))
                {
                    var key = ((Keys)keyCode).ToString();
                    var character = TextCharacter(info, modifiers);
                    var submittingEdit = keyCode == (int)Keys.Enter &&
                        IsEditableTarget(target, cachedEditableElement);
                    if (submittingEdit && foreground == GetForegroundWindow() &&
                        focus == ForegroundFocusHandle() && DateTimeOffset.Now - at < TimeSpan.FromSeconds(2))
                    {
                        try
                        {
                            var liveFocus = AutomationElement.FocusedElement;
                            var liveTarget = liveFocus is null ? null : Automation.Describe(liveFocus);
                            if (liveFocus is not null && liveTarget is not null &&
                                liveTarget.Process?.Equals(target.Process,
                                    StringComparison.OrdinalIgnoreCase) == true &&
                                IsEditableTarget(liveTarget, liveFocus))
                            {
                                target = liveTarget;
                                targetScreen = ScreenFor(liveFocus);
                                cachedEditableElement = liveFocus;
                            }
                        }
                        catch (Exception ex) when (ex is ElementNotAvailableException or
                            InvalidOperationException or System.Runtime.InteropServices.COMException)
                        { System.Diagnostics.Trace.WriteLine(ex); }
                    }
                    var searchEnter = target is { Process: "SearchHost", AutomationId: "SearchTextBox" } &&
                        keyCode == (int)Keys.Enter;
                    string? submittedValue = null;
                    if (submittingEdit && cachedEditableElement is { } submittedElement)
                        try
                        {
                            if (submittedElement.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern))
                                submittedValue = ((ValuePattern)pattern).Current.Value;
                        }
                        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                    string? selectionState = null;
                    if (target.Process == "EXCEL" && keyCode == (int)Keys.C &&
                        modifiers == "Control")
                    {
                        try
                        {
                            if (foreground != GetForegroundWindow() ||
                                DateTimeOffset.Now - at > TimeSpan.FromSeconds(2))
                                throw new InvalidOperationException("Excel copy selection changed before capture.");
                            selectionState = "excel-selection-range:" + ExcelNativeSheet.Read<string>(foreground,
                                sheet => ExcelNativeSheet.SelectionAddress(sheet));
                        }
                        catch (Exception ex) when (ex is InvalidOperationException or COMException)
                        {
                            Add(new(at, "unresolved-input", target, null, "Control+C", null, null, null,
                                Diagnostic: "Excel copy range could not be captured: " + ex.Message));
                            return;
                        }
                    }
                    Add(new(at, "key", target, character,
                        modifiers == "None" ? key : $"{modifiers}+{key}", null,
                        string.IsNullOrWhiteSpace(submittedValue) ? null : "field-value:" + submittedValue,
                        selectionState));
                    if (searchEnter)
                        QueueSearchLaunchOutcome(events.Count - 1, foreground);
                    if (submittingEdit)
                    {
                        // Browser pages can replace the focused editor without changing
                        // their top-level HWND or native keyboard-focus handle.
                        cachedEditableTarget = null;
                        cachedEditableElement = null;
                        cachedKeyboardTarget = null;
                        cachedKeyboardFocus = 0;
                    }
                    typingScreen = targetScreen;
                    typingCapture.Stop();
                    typingCapture.Start();
                }
            }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
        }
    }

    private static bool IsOwnWindow(ControlRef target) =>
        target.Process?.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), StringComparison.OrdinalIgnoreCase) == true;

    private static bool IsEditableTarget(ControlRef target, AutomationElement? element = null)
    {
        if (target.ControlType == "ControlType.Edit" || target.ClassName == "EDTBX") return true;
        if (target.ControlType == "ControlType.Pane" && !string.IsNullOrEmpty(target.AutomationId)) return true;
        try { return element?.TryGetCurrentPattern(ValuePattern.Pattern, out _) == true; }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); return false; }
    }

    public bool AddIntentNote(string note)
    {
        lock (events)
        {
            for (var index = events.Count - 1; index >= 0; index--)
            {
                if (events[index].Kind == "key" && ActionGrouper.IsStandaloneModifier(events[index].Key)) continue;
                events[index] = events[index] with { Intent = note.Trim() };
                return true;
            }
        }
        return false;
    }

    // Called after the low-level keyboard hook has returned. UI Automation, COM,
    // screenshot capture, and WinForms timer work must not run on the hook stack.
    public void PrepareIntentDialog() => FinishTypingCapture();

    public void IntentDialogOpened(nint window) => intentDialogWindow = window;
    public void IntentDialogClosed()
    {
        capturePausedForIntent = false;
        intentDialogWindow = 0;
        cachedEditableTarget = null;
    }

    public int Pause(int? discardFromEvent = null)
    {
        if (!IsRecording || capturePausedByUser) return 0;
        FinishTypingCapture();
        FinishPendingClick();
        CompleteScroll();
        var discarded = 0;
        if (discardFromEvent is int start)
        {
            lock (events)
            {
                if (start >= 0 && start < events.Count)
                {
                    discarded = events.Count - start;
                    events.RemoveRange(start, discarded);
                }
            }
        }
        capturePausedByUser = true;
        chartSamplePaused = true;
        Interlocked.Increment(ref chartInputGeneration);
        hoveredChart = null;
        cachedEditableTarget = null;
        cachedEditableElement = null;
        cachedKeyboardTarget = null;
        cachedKeyboardScreen = null;
        return discarded;
    }

    public void Resume()
    {
        if (!IsRecording || !capturePausedByUser) return;
        capturePausedByUser = false;
        chartSamplePaused = capturePausedForIntent || pendingLegendOutcome is not null;
        chartPointerDown = false;
        chartSampleRequested.Set();
        cachedEditableTarget = null;
        cachedEditableElement = null;
        cachedKeyboardTarget = null;
        cachedKeyboardScreen = null;
        cachedKeyboardFocus = 0;
        cachedKeyboardWindow = 0;
    }


    private void FinishPendingClick()
    {
        var index = pendingClickIndex;
        try { FinishPendingClickCore(); }
        catch (ArgumentException ex)
        {
            System.Diagnostics.Trace.WriteLine("Click outcome accessibility capture: " + ex);
            if (index < 0 || index >= events.Count) return;
            RecordedEvent unresolved;
            lock (events)
                events[index] = unresolved = events[index] with { Kind = "unresolved-input",
                    Diagnostic = "The accessibility provider returned invalid bounds while verifying the click: " +
                        ex.Message + " Its outcome was not guessed." };
            Captured?.Invoke(unresolved);
        }
    }

    private void FinishPendingClickCore()
    {
        clickCapture.Stop();
        if (pendingClickIndex < 0 || pendingClickIndex >= events.Count) return;
        var index = pendingClickIndex;
        var screen = pendingClickScreen;
        var point = pendingClickPoint;
        var window = pendingClickWindow;
        var tabBefore = pendingClickTabBefore;
        var focusBefore = pendingClickFocusBefore;
        pendingClickIndex = -1; pendingClickScreen = null; pendingClickPoint = null;
        pendingClickWindow = 0; pendingClickTabBefore = null;
        pendingClickFocusBefore = null;
        string? screenshot;
        ControlRef? changedTab = null;
        if (window != 0 && tabBefore is not null)
        {
            try
            {
                var selected = SelectedTab(AutomationElement.FromHandle(window),
                    events[index].Target?.Window != "Find and Replace");
                if (selected is not null && selected.Current.Name != tabBefore)
                    changedTab = Automation.Describe(selected);
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        }
        var target = changedTab ?? events[index].Target;
        if (changedTab is not null && window != 0)
        {
            lock (dialogCacheLock) dialogControlCache.Remove(window);
            QueueDialogControlSnapshot(window);
        }
        var afterState = events[index].AfterState;
        if (ExcelFilterPlan.IsItem(target))
        {
            try
            {
                afterState = ExcelFilterPlan.CaptureState(target!);
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        }
        screenshot = events[index].Target?.Window == "Find and Replace" &&
            events[index].Target?.ControlType != "ControlType.Window" ? null : Capture(screen);

        // A result dialog can hide the command that opened it. After its
        // acknowledgement closes, focus commonly returns to that exact command.
        // Use that live accessibility identity to repair the preceding weak
        // dialog click and mark this click as the result acknowledgement.
        if (target?.ControlType is "ControlType.Window" or "ControlType.Pane" or
                "ControlType.Custom" or "ControlType.Tab")
        {
            try
            {
                var focusedElement = AutomationElement.FocusedElement;
                var returnedCommand = Automation.Describe(focusedElement);
                if (returnedCommand is { ControlType: "ControlType.Button" or "ControlType.MenuItem" or
                        "ControlType.TabItem" or "ControlType.CheckBox" or "ControlType.RadioButton",
                        Name.Length: > 0 } &&
                    returnedCommand.Process == target.Process && returnedCommand.Window == target.Window)
                {
                    lock (events)
                    {
                        var priorIndex = events.FindLastIndex(index - 1, prior =>
                            prior.Kind == "click" && prior.Target is { } priorTarget &&
                            priorTarget.Process == target.Process && priorTarget.Window == target.Window &&
                            priorTarget.ControlType is "ControlType.Window" or "ControlType.Pane" or
                                "ControlType.Custom" or "ControlType.Tab" &&
                            events[index].At - prior.At < TimeSpan.FromSeconds(15));
                        if (priorIndex >= 0)
                        {
                            events[priorIndex] = events[priorIndex] with
                            { Target = returnedCommand, AfterState = "dialog-command-result" };
                            TargetCorrected?.Invoke(returnedCommand);
                            afterState = "dialog-result-acknowledgement";
                        }
                    }
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        }
        if (changedTab is null && events[index].Diagnostic != "Control identity verified at mouse-down." &&
            point is { } clickPoint && target is not null &&
            target.ClassName != "XLGridRowHeader" &&
            (target.ControlType == "ControlType.DataItem" || target.ControlType == "ControlType.ToolTip") &&
            events[index].Kind is "click" or "context-click")
        {
            try
            {
                var live = AutomationElement.FromPoint(new System.Windows.Point(clickPoint.X, clickPoint.Y));
                var liveBounds = live.Current.BoundingRectangle;
                var liveTarget = Automation.Describe(live);
                if (liveTarget is not null && !IsWeakClickTarget(liveTarget) &&
                    (target.ControlType != "ControlType.DataItem" || liveTarget.ControlType == "ControlType.DataItem") &&
                    liveTarget.Process == target.Process &&
                    liveBounds.Contains(clickPoint.X, clickPoint.Y))
                    target = liveTarget;
                else if (target.ControlType == "ControlType.DataItem" &&
                    AutomationElement.FocusedElement is { } focusedElement &&
                    Automation.Describe(focusedElement) is { ControlType: "ControlType.DataItem" } focused &&
                    focused.Process == target.Process && focused.Window == target.Window &&
                    !string.IsNullOrWhiteSpace(focused.AutomationId) &&
                    (focused != focusBefore || focused.Name == target.Name ||
                     focusedElement.Current.BoundingRectangle.Contains(clickPoint.X, clickPoint.Y)))
                    target = focused;
                else
                {
                    afterState = "unverified-click";
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException)
            {
                afterState = "unverified-click";
                System.Diagnostics.Trace.WriteLine(ex);
            }
        }
        if (changedTab is null && events[index].Diagnostic != "Control identity verified at mouse-down." &&
            point is { } selectedPoint && target?.ControlType is
            "ControlType.DataItem" or "ControlType.ListItem")
        {
            target = CorrectCompactListWindow(target, selectedPoint.X, selectedPoint.Y);
            if (afterState == "unverified-click")
            {
                var selected = MsaaActions.IsSelectedListItemAtPoint(
                    selectedPoint.X, selectedPoint.Y, target.Name);
                if (!selected)
                    try
                    {
                        var live = AutomationElement.FromPoint(new System.Windows.Point(
                            selectedPoint.X, selectedPoint.Y));
                        selected = live is not null && live.Current.Name == target.Name &&
                            live.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selection) &&
                            ((SelectionItemPattern)selection).Current.IsSelected;
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                if (selected) afterState = "native-selection-verified";
            }
        }
        if (window != 0 && (target?.ControlType is ("ControlType.Window" or "ControlType.TitleBar") ||
            target?.Name == "Close") &&
            !IsWindowVisible(window))
            afterState = "window-closed";
        lock (events)
        {
            var current = events[index];
            if (target?.ControlType is "ControlType.Window" or "ControlType.Pane" or
                    "ControlType.Custom" or "ControlType.Tab" &&
                current.Target?.ControlType is "ControlType.Button" or "ControlType.TabItem" or
                    "ControlType.CheckBox" or "ControlType.RadioButton")
                target = current.Target;
            if (current.AfterState == "dialog-command" ||
                IsWeakSwitchTarget(target) && current.Target is not null &&
                !IsWeakSwitchTarget(current.Target)) target = current.Target;
            if ((current.AfterState?.StartsWith("dialog-probe:", StringComparison.Ordinal) == true &&
                 afterState is null) || current.AfterState == "dialog-command")
                afterState = current.AfterState;
            events[index] = current with
            { Screenshot = screenshot, Target = target, AfterState = afterState };
        }
        if (index > 0)
        {
            RecordedEvent? unresolved = null;
            lock (events)
            {
                var prior = events[index - 1];
                if (IsUnresolvedDialogEvent(prior) &&
                    prior.Target is { Process: "EXCEL", Window: { Length: > 0 } priorWindow } &&
                    target is { Process: "EXCEL", ControlType: "ControlType.TabItem",
                        Window: { Length: > 0 } tabWindow, Name: { Length: > 0 } } &&
                    string.Equals(priorWindow, tabWindow, StringComparison.OrdinalIgnoreCase))
                {
                    // Excel can expose the workbook surface itself while focus moves between
                    // sheet tabs. The following verified tab proves that this weak click had
                    // no independent command outcome, so retain it only as discardable focus
                    // evidence and do not interrupt a valid recording.
                    events[index - 1] = prior with { AfterState = "worksheet-tab-focus-only" };
                }
                else if (IsUnresolvedDialogEvent(prior)) unresolved = prior;
            }
            if (unresolved is not null)
                QueueUnresolvedDialogNotice(index - 1, unresolved);
        }
        if (afterState != "unverified-click" && target?.ControlType == "ControlType.DataItem" &&
            !string.IsNullOrWhiteSpace(target.AutomationId) && IsFilteredGridContext(index))
        {
            var capturedTarget = target;
            pendingFieldCaptures.RemoveAll(task => task.IsCompleted);
            pendingFieldCaptures.Add(Task.Run(() =>
            {
                try
                {
                    var ordinal = Automation.VisibleGridOrdinal(capturedTarget);
                    if (ordinal is not > 0) return;
                    lock (events)
                    {
                        if (events[index].Target == capturedTarget)
                            events[index] = events[index] with { AfterState = "visible-row:" + ordinal };
                    }
                }
                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
            }));
        }
        if (window != 0 && target?.ControlType is ("ControlType.Window" or "ControlType.TitleBar") &&
            afterState != "window-closed")
        {
            pendingWindowCaptures.RemoveAll(task => task.IsCompleted);
            pendingWindowCaptures.Add(Task.Run(async () =>
            {
                for (var attempt = 0; attempt < 12; attempt++)
                {
                    await Task.Delay(150);
                    if (IsWindowVisible(window)) continue;
                    lock (events) events[index] = events[index] with { AfterState = "window-closed" };
                    return;
                }
            }));
        }
        if (target is not null && (IsEditableTarget(target) ||
            target.ControlType == "ControlType.Pane" && !string.IsNullOrEmpty(target.AutomationId)))
            CaptureCompletedField(index, target, null);
    }

    private void QueueUnresolvedDialogNotice(int index, RecordedEvent recorded)
    {
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            // Command retries and the diagnostic probe may still be identifying
            // an owner-drawn control after the next click has been processed.
            await Task.Delay(5000);
            hookDispatcher.BeginInvoke(new Action(() =>
            {
                lock (events)
                    if (index >= events.Count || events[index].At != recorded.At ||
                        !IsUnresolvedDialogEvent(events[index])) return;
                if (!IsRecording || IsPaused) return;
                UnresolvedDialogClick?.Invoke(index,
                    recorded.Target?.Window ?? recorded.Target?.Name ?? "dialog");
            }));
        }));
    }

    private static bool IsWeakClickTarget(ControlRef target) =>
        target.ControlType is "ControlType.ToolTip" or "ControlType.Window" or
            "ControlType.TitleBar" or "ControlType.Pane" or "ControlType.Custom" ||
        string.IsNullOrWhiteSpace(target.Name);

    private static ControlRef CorrectCompactListWindow(ControlRef target, int x, int y)
    {
        var window = CompactWindowAtPoint(x, y);
        var dialog = window == 0 ? null : NativeControlRef(window, null, "", "ControlType.Window", null);
        return dialog is not null && dialog.Process == target.Process &&
            !string.IsNullOrWhiteSpace(dialog.Window)
            ? target with { Window = dialog.Window } : target;
    }

    private bool IsFilteredGridContext(int index)
    {
        var recent = events.Skip(Math.Max(0, index - 18)).Take(Math.Min(18, index)).ToList();
        var lastFilterOk = recent.FindLastIndex(item => item.Target?.Name == "OK" &&
            item.Target.ControlType == "ControlType.Button");
        if (lastFilterOk < 0) return false;
        return recent.Take(lastFilterOk).TakeLast(8).Any(item => item.Target?.ControlType == "ControlType.TreeItem") &&
            recent.Take(lastFilterOk).TakeLast(12).Any(item => item.Target?.AutomationId == "Dropdown");
    }

    private static nint CompactWindowAtPoint(int x, int y)
    {
        var child = WindowFromPoint(new System.Drawing.Point(x, y));
        var screen = Screen.FromPoint(new System.Drawing.Point(x, y)).Bounds;
        // An owned popup can be a small top-level window whose root resolves
        // to the maximized application. Use that popup before walking to root.
        // Child edits and buttons still resolve through their dialog root.
        if (child != 0 && GetAncestor(child, 1) == 0 &&
            IsCompactAtPoint(child, x, y, screen)) return child;
        var window = child == 0 ? 0 : GetAncestor(child, 2);
        if (window == 0) window = child;
        if (IsCompactAtPoint(window, x, y, screen)) return window;
        // Owner-drawn dialogs can give WindowFromPoint the application's frame.
        // UIA may still report the dialog itself at the same pointer location.
        try
        {
            var hit = AutomationElement.FromPoint(new System.Windows.Point(x, y));
            var dialog = hit is null ? null : WindowAncestor(hit);
            var handle = dialog is null ? 0 : (nint)dialog.Current.NativeWindowHandle;
            return IsCompactAtPoint(handle, x, y, screen) ? handle : 0;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        { return 0; }
    }

    private static bool IsCompactAtPoint(nint window, int x, int y,
        System.Drawing.Rectangle screen)
    {
        if (window == 0 || !GetWindowRect(window, out var bounds) ||
            x < bounds.Left || x >= bounds.Right || y < bounds.Top || y >= bounds.Bottom)
            return false;
        return (long)(bounds.Right - bounds.Left) * (bounds.Bottom - bounds.Top) <
            (long)screen.Width * screen.Height * 8 / 10;
    }

    private static bool IsCompactForegroundWindow(int x, int y) => CompactWindowAtPoint(x, y) != 0;

    private static bool IsNativeEditClass(string name) =>
        name.Equals("Edit", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("EDTBX", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("RichEdit", StringComparison.OrdinalIgnoreCase);

    private static bool TryDescribeNativeEdit(nint root, nint child, out ControlRef? target)
    {
        target = null;
        if (root == 0 || child == 0 || GetAncestor(child, 2) != root) return false;
        var classBuffer = new System.Text.StringBuilder(128);
        GetClassName(child, classBuffer, classBuffer.Capacity);
        var className = classBuffer.ToString();
        if (!IsNativeEditClass(className)) return false;
        var id = GetDlgCtrlID(child);
        if (id <= 0 || id >= 65536) return false;
        target = NativeControlRef(root, id.ToString(), "", className == "Edit"
            ? "ControlType.Edit" : "ControlType.Pane", className);
        return target is not null;
    }

    private static ControlRef? NativeControlRef(nint root, string? id, string name,
        string type, string? className)
    {
        var title = new System.Text.StringBuilder(512);
        GetWindowText(root, title, title.Capacity);
        GetWindowThreadProcessId(root, out var processId);
        if (processId == 0) return null;
        if (title.Length == 0)
        {
            // Popup menus can be untitled windows owned by the workbook.
            // Keep the recorded control tied to its application's window.
            var owners = new List<nint> { GetAncestor(root, 3) };
            for (var owner = GetWindow(root, 4); owner != 0 && owners.Count < 16;
                owner = GetWindow(owner, 4))
            {
                if (owners.Contains(owner)) break;
                owners.Add(owner);
            }
            owners.Add(GetForegroundWindow());
            foreach (var candidate in owners)
            {
                if (candidate == 0) continue;
                GetWindowThreadProcessId(candidate, out var candidateProcess);
                if (candidateProcess != processId) continue;
                GetWindowText(candidate, title, title.Capacity);
                if (title.Length > 0) break;
            }
        }
        if (title.Length == 0) return null;
        string process;
        try { process = System.Diagnostics.Process.GetProcessById((int)processId).ProcessName; }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException)
        { System.Diagnostics.Trace.WriteLine(ex); return null; }
        return new ControlRef(process, title.ToString(), id, name, type, className, null,
            ProcessId: (int)processId);
    }

    private static ControlRef? CapturePopupMenuTarget(int x, int y) =>
        CapturePopupMenuSnapshot(x, y)?.Target;

    private static MouseDownSnapshot? CapturePopupMenuSnapshot(int x, int y)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        var popup = pointer == 0 ? 0 : GetAncestor(pointer, 2);
        if (popup == 0 || !IsWindowVisible(popup)) return null;
        var className = new System.Text.StringBuilder(128);
        GetClassName(popup, className, className.Capacity);
        // Limit synchronous UIA access to a small, transient native popup, not
        // a potentially large application tree.
        if (className.ToString() != "#32768" &&
            ((GetWindowLong(popup, -16) & unchecked((int)0x80000000)) == 0 ||
             GetWindow(popup, 4) == 0)) return null;
        try
        {
            var element = AutomationElement.FromPoint(new System.Windows.Point(x, y));
            for (var depth = 0; element is not null && depth < 4; depth++)
            {
                if (element.Current.ControlType == ControlType.MenuItem &&
                    element.Current.IsEnabled && !element.Current.IsOffscreen)
                {
                    var bounds = element.Current.BoundingRectangle;
                    if (bounds.IsEmpty || !bounds.Contains(x, y)) return null;
                    var described = Automation.Describe(element);
                    if (described is null || string.IsNullOrWhiteSpace(described.Name)) return null;
                    var owner = NativeControlRef(popup, described.AutomationId,
                        described.Name, described.ControlType!, described.ClassName);
                    if (owner is null || owner.Process != described.Process ||
                        owner.ProcessId != described.ProcessId) return null;
                    return new(null, owner with { ParentName = Automation.RecordedMenuParentName(element) })
                    { WindowHandle = popup, InputBounds = bounds };
                }
                element = TreeWalker.ControlViewWalker.GetParent(element);
            }
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        return null;
    }

    private void CompleteChartSourceCapture()
    {
        if (chartSourceCapture is not { } pending) return;
        if (mouseHook != 0 && Environment.TickCount64 - Interlocked.Read(ref lastChartInputTick) < 1500)
            return;
        var contextIndex = events.FindLastIndex(item => item.Kind == "context-click" &&
            item.Target == pending.Target);
        if (contextIndex < 0) return;
        var sequence = events.Skip(contextIndex + 1).ToList();
        if (sequence.Any(item => item.Target is { } target &&
            (target.Process != pending.Target.Process || target.ProcessId != pending.Target.ProcessId ||
                target.Window != pending.Target.Window && target.Window != "Select Data Source")))
        {
            chartSourceCapture = null;
            return;
        }
        if (sequence.Any(item => item.Target is { Window: "Select Data Source", Name: "Cancel" }))
        {
            chartSourceCapture = null;
            return;
        }
        if (!sequence.Any(item => item.Target?.Window == "Select Data Source")) return;
        var foreground = GetForegroundWindow();
        if (foreground != pending.Window || !IsWindowEnabled(pending.Window)) return;
        try
        {
            if (pending.Sheet is null)
                throw new InvalidOperationException("The chart worksheet identity was not captured.");
            var after = ExcelNativeSheet.Read<ExcelNativeSheet.ChartSource>(pending.Window,
                sheet => ExcelNativeSheet.ReadChartSourceFromSheet(sheet, pending.Sheet, pending.Target.Name!));
            if (after != pending.Before)
            {
                Add(new(DateTimeOffset.Now, "set-chart-source-range", pending.Target,
                    System.Text.Json.JsonSerializer.Serialize(after), null, Capture(Screen.FromHandle(pending.Window)),
                    System.Text.Json.JsonSerializer.Serialize(pending.Before), "chart-source-verified",
                    Diagnostic: "Captured the chart's actual category/value source after Select Data confirmation."));
                chartSourceCapture = null;
            }
        }
        catch (Exception ex) when (ex is InvalidOperationException or COMException or NotSupportedException)
        {
            // Range selection can temporarily expose a partial/unsupported source
            // while Excel's collapsed selector is still open. Keep observing until
            // recording stops; only then is the absence of a valid outcome final.
            if (IsRecording) return;
            chartSourceCapture = null;
            Add(new(DateTimeOffset.Now, "unresolved-input", pending.Target, null, null, null, null, null,
                Diagnostic: "The confirmed chart source could not be captured: " + ex.Message +
                    " Re-record this chart edit with a supported adjacent category/value range."));
        }
    }

    private void PrimeHoveredChart() =>
        PrimeHoveredChartCore(point => AutomationElement.FromPoint(point));

    private void PrimeHoveredChartCore(Func<System.Windows.Point, AutomationElement?> hitTest)
    {
        if (chartSampleStopping || chartSamplePaused || chartPointerDown || chartSourceCapture is not null) return;
        var inputGeneration = Interlocked.Read(ref chartInputGeneration);
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var stage = "pointer UIA hit";
        var attemptedChart = false;
        var previous = hoveredChart;
        try
        {
            var point = Cursor.Position;
            var root = GetAncestor(WindowFromPoint(point), 2);
            var element = hitTest(new(point.X, point.Y));
            if (element is null)
            {
                CaptureDiagnostic("chart hover: pointer UIA hit returned no element; no snapshot published");
                return;
            }
            var tooltip = element.Current.ControlType == ControlType.ToolTip;
            if (tooltip && root != GetForegroundWindow())
            {
                GetWindowThreadProcessId(root, out var tooltipOwner);
                GetWindowThreadProcessId(GetForegroundWindow(), out var foregroundOwner);
                if (tooltipOwner != foregroundOwner) return;
                root = GetForegroundWindow();
            }
            if (root == 0 || root != GetForegroundWindow() || !IsWindowEnabled(root) ||
                NativeControlRef(root, "", "", "ControlType.Image", "") is not { Process: "EXCEL" } target)
                return;
            var names = new List<string>();
            var bounds = ReadChartBounds(() => element.Current.BoundingRectangle);
            var ancestor = element;
            var chartAncestor = false;
            for (var depth = 0; ancestor is not null && depth < 10; depth++)
            {
                if (ancestor.Current.ControlType == ControlType.Image)
                {
                    chartAncestor = true;
                    if (!string.IsNullOrWhiteSpace(ancestor.Current.Name)) names.Add(ancestor.Current.Name);
                    var rectangle = ReadChartBounds(() => ancestor.Current.BoundingRectangle);
                    if (!rectangle.IsEmpty && rectangle.Contains(point.X, point.Y) &&
                        (bounds.IsEmpty || rectangle.Width * rectangle.Height > bounds.Width * bounds.Height))
                        bounds = rectangle;
                }
                ancestor = TreeWalker.RawViewWalker.GetParent(ancestor);
            }
            if (!chartAncestor && !tooltip) return;
            attemptedChart = true;
            stage = "native chart identity and legend";
            var chart = ExcelNativeSheet.Read<ExcelNativeSheet.ChartCapture?>(root,
                sheet => ExcelNativeSheet.CaptureChartAtPoint(sheet, point.X, point.Y, true, names));
            if (chart is null) { CaptureDiagnostic("chart hover: native identity did not match"); return; }
            if (tooltip || bounds.IsEmpty)
            {
                if (previous?.Snapshot is not { } captured || captured.WindowHandle != root ||
                    captured.Target.Name != chart.Name || captured.ChartSheet != chart.Sheet ||
                    captured.Element is null) { CaptureDiagnostic("chart hover: tooltip/invalid bounds has no corroborated prior element"); return; }
                element = captured.Element;
                bounds = ReadChartBounds(() => element.Current.BoundingRectangle);
            }
            if (bounds.IsEmpty || !bounds.Contains(point.X, point.Y))
            { CaptureDiagnostic("chart hover: bounds unavailable or do not contain pointer"); return; }
            var livePoint = Cursor.Position;
            if (chartSampleStopping || chartSamplePaused || capturePausedByUser || capturePausedForIntent ||
                inputGeneration != Interlocked.Read(ref chartInputGeneration) ||
                chartPointerDown || !bounds.Contains(livePoint.X, livePoint.Y) ||
                WindowFromPoint(livePoint) != WindowFromPoint(point) ||
                GetForegroundWindow() != root)
            { CaptureDiagnostic("chart hover: input, pointer or foreground changed during sampling"); return; }
            hoveredChart = new(DateTimeOffset.Now, bounds, new(element, target with {
                Name = chart.Name, ClassName = "ExcelChartObject" }) {
                WindowHandle = root, HitWindow = WindowFromPoint(point), ChartSheet = chart.Sheet,
                LegendBefore = chart.Legend });
            CaptureDiagnostic($"chart hover: snapshot ready; legend={chart.Legend is not null}; {watch.ElapsedMilliseconds} ms");
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
            InvalidOperationException or COMException or System.ComponentModel.Win32Exception)
        {
            // A non-elevated recorder cannot inspect an elevated window. Chart hover sampling is
            // advisory, so access denial must discard this sample instead of terminating RSR.
            CaptureDiagnostic($"chart hover failed at {stage}: {ex.GetType().Name}; {watch.ElapsedMilliseconds} ms");
            System.Diagnostics.Trace.WriteLine("Chart hover capture could not sample a valid accessibility rectangle: " + ex.Message);
        }
        finally
        {
            if (watch.ElapsedMilliseconds >= 100 && (attemptedChart || watch.ElapsedMilliseconds >= 500))
                CaptureDiagnostic($"chart hover stage {stage}: {watch.ElapsedMilliseconds} ms");
        }
    }

    private void StartChartSampling(Action sample)
    {
        if (chartSampleThread is not null)
            throw new InvalidOperationException("The previous chart sampler has not stopped.");
        chartSampleStopping = false;
        chartSamplePaused = false;
        chartSampleThread = new Thread(() =>
        {
            while (!chartSampleStopping)
            {
                chartSampleRequested.WaitOne();
                if (chartSampleStopping) continue;
                try { sample(); }
                catch (Exception ex)
                {
                    // This worker only improves chart recognition. Keep recorder input hooks alive
                    // when any inaccessible or short-lived application defeats a sampling provider.
                    CaptureDiagnostic($"chart sampler ignored {ex.GetType().Name}: {ex.Message}");
                    System.Diagnostics.Trace.WriteLine("Chart sampling failed without stopping recording: " + ex);
                }
            }
        }) { IsBackground = true, Name = "RSR accessibility chart sampler" };
        chartSampleThread.SetApartmentState(ApartmentState.STA);
        chartSampleThread.Start();
        chartSampleRequested.Set();
    }

    private void StopChartSampling()
    {
        chartSampleStopping = true;
        chartSampleRequested.Set();
        if (chartSampleThread is null) return;
        if (!chartSampleThread.Join(TimeSpan.FromSeconds(10)))
            throw new TimeoutException("The accessibility chart sampler did not stop within ten seconds.");
        chartSampleThread = null;
        hoveredChart = null;
    }

    private void FinishLegendOutcome()
    {
        legendOutcomeCapture.Stop();
        var outcome = pendingLegendOutcome;
        pendingLegendOutcome = null;
        outcome?.Invoke();
    }

    internal static System.Windows.Rect ReadChartBounds(Func<System.Windows.Rect> read)
    {
        try
        {
            var rectangle = read();
            if (!rectangle.IsEmpty && double.IsFinite(rectangle.X) && double.IsFinite(rectangle.Y) &&
                double.IsFinite(rectangle.Width) && double.IsFinite(rectangle.Height) &&
                rectangle.Width > 0 && rectangle.Height > 0)
                return rectangle;
            System.Diagnostics.Trace.WriteLine("Chart accessibility bounds are empty or invalid; no pointer identity was inferred.");
        }
        catch (ArgumentException ex)
        {
            System.Diagnostics.Trace.WriteLine("Chart accessibility provider returned invalid bounds: " + ex.Message);
        }
        return System.Windows.Rect.Empty;
    }

    private sealed record MouseDownSnapshot(AutomationElement? Element, ControlRef Target)
    {
        public System.Windows.Rect InputBounds { get; init; } = System.Windows.Rect.Empty;
        public nint WindowHandle { get; init; }
        public nint HitWindow { get; init; }
        public bool IsDialog { get; init; }
        public string? ChartSheet { get; init; }
        public ExcelNativeSheet.LegendLayout? LegendBefore { get; init; }
        public ExcelNativeSheet.ChartSource? ChartSource { get; init; }
    }

    private static MouseDownSnapshot? CaptureFilterItemAtPoint(int x, int y)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        var root = pointer == 0 ? 0 : GetAncestor(pointer, 2);
        if (root == 0 || GetWindow(root, 4) == 0 ||
            (GetWindowLong(root, -16) & 0x00C00000) != 0 ||
            (GetWindowLong(root, -16) & unchecked((int)0x80000000)) == 0 ||
            !IsCompactAtPoint(root, x, y, Screen.FromPoint(new(x, y)).Bounds) ||
            !string.Equals(DescribeNativeWindowOwner(root), "EXCEL", StringComparison.OrdinalIgnoreCase) ||
            MsaaActions.FilterItemNameAtPoint(x, y) is not { } name) return null;
        var target = NativeControlRef(root, "", name, "ControlType.TreeItem", "");
        return target is null ? null : new(null, target with { ParentName = "Manual Filter" })
        { WindowHandle = root };
    }

    private static MouseDownSnapshot? CaptureDirectMenuAtPoint(int x, int y) =>
        CaptureDirectMenuAtPointCore(x, y, true);

    // The released-menu loop refreshes continuously while the popup stays open; a
    // slow UIA read (search-box menus take ~0.5 s) must not age out the newest snapshot.
    internal static readonly TimeSpan PointerMenuLifetime = TimeSpan.FromMilliseconds(1500);

    private static bool IsPointerMenuFresh(PointerMenu menu, DateTimeOffset at, long generation,
        nint foreground, int x, int y) =>
        menu.Generation == generation && IsMenuScopeForeground(menu.Foreground, foreground) &&
        at >= menu.At && at - menu.At < PointerMenuLifetime &&
        menu.Snapshot.InputBounds.Contains(x, y);

    // Context menus with an embedded search box take keyboard focus. Treat that
    // same-process popup as part of the original menu scope, never another window.
    private static bool IsMenuScopeForeground(nint original, nint current)
    {
        if (current == original) return current != 0;
        if (current == 0 || original == 0 || !IsWindowVisible(current)) return false;
        GetWindowThreadProcessId(current, out var currentProcess);
        GetWindowThreadProcessId(original, out var originalProcess);
        if (currentProcess == 0 || currentProcess != originalProcess) return false;
        var root = GetAncestor(current, 2);
        if (root == 0 || root == GetAncestor(original, 2)) return false;
        if (GetAncestor(root, 3) == GetAncestor(original, 3)) return true;
        var style = unchecked((uint)GetWindowLong(root, -16));
        return (style & 0x80000000) != 0 && (style & 0x00C00000) == 0;
    }

    private async Task CaptureReleasedContextMenuAsync(nint foreground, long generation)
    {
        try
        {
            await Task.Run(async () =>
            {
                var attempt = 0;
                for (; attempt < 40 && IsRecording && !capturePausedByUser &&
                    !capturePausedForIntent && generation == Interlocked.Read(ref chartInputGeneration) &&
                    IsMenuScopeForeground(foreground, GetForegroundWindow()); attempt++)
                {
                    var started = DateTimeOffset.Now;
                    GetWindowThreadProcessId(foreground, out var processId);
                    var popups = new List<nint>();
                    EnumWindows((popup, _) =>
                    {
                        GetWindowThreadProcessId(popup, out var owner);
                        if (owner != processId || !IsWindowVisible(popup) ||
                            !GetWindowRect(popup, out var rect)) return true;
                        var className = new System.Text.StringBuilder(128);
                        GetClassName(popup, className, className.Capacity);
                        var center = new System.Drawing.Point((rect.Left + rect.Right) / 2,
                            (rect.Top + rect.Bottom) / 2);
                        if (className.ToString() == "#32768" ||
                            GetWindow(popup, 4) != 0 && (GetWindowLong(popup, -16) & 0x00C00000) == 0 &&
                            IsCompactAtPoint(popup, center.X, center.Y, Screen.FromPoint(center).Bounds))
                            popups.Add(popup);
                        return popups.Count < 8;
                    }, 0);
                    var captured = new List<PointerMenu>();
                    foreach (var popup in popups)
                    {
                        var items = AutomationElement.FromHandle(popup).FindAll(TreeScope.Descendants,
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));
                        if (items.Count > 128) continue;
                        foreach (AutomationElement item in items)
                        {
                            var state = item.Current;
                            var bounds = state.BoundingRectangle;
                            if (state.ProcessId != processId || !state.IsEnabled || state.IsOffscreen ||
                                string.IsNullOrWhiteSpace(state.Name) || bounds.IsEmpty ||
                                bounds.Width <= 0 || bounds.Height <= 0) continue;
                            var center = new System.Drawing.Point((int)(bounds.Left + bounds.Width / 2),
                                (int)(bounds.Top + bounds.Height / 2));
                            if (GetAncestor(WindowFromPoint(center), 2) != popup) continue;
                            var target = NativeControlRef(popup, state.AutomationId, state.Name,
                                "ControlType.MenuItem", state.ClassName);
                            if (target is null) continue;
                            captured.Add(new(started, generation, foreground,
                                new(null, target with { ParentName = Automation.RecordedMenuParentName(item) })
                                { WindowHandle = popup, InputBounds = bounds }));
                        }
                    }
                    if (generation != Interlocked.Read(ref chartInputGeneration) ||
                        !IsRecording || capturePausedByUser || capturePausedForIntent ||
                        !IsMenuScopeForeground(foreground, GetForegroundWindow())) break;
                    var completed = DateTimeOffset.Now;
                    if (completed - started < PointerMenuLifetime)
                        pointerMenus = captured.Select(item => item with { At = completed }).ToArray();
                    if (captured.Count > 0)
                        CaptureDiagnostic($"released menu snapshot: {captured.Count} items; {(DateTimeOffset.Now - started).TotalMilliseconds:F0} ms");
                    await Task.Delay(150);
                }
                if (attempt < 40 && IsRecording)
                    CaptureDiagnostic($"released menu capture stopped after {attempt} attempt(s): " +
                        (generation != Interlocked.Read(ref chartInputGeneration) ? "new input" :
                        capturePausedByUser || capturePausedForIntent ? "paused" :
                        "foreground left menu scope: " + DescribeNativeWindow(GetForegroundWindow())));
            }).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is ArgumentException or ElementNotAvailableException or
            InvalidOperationException or COMException)
        { CaptureDiagnostic("released menu capture rejected: " + ex.GetType().Name + ": " + ex.Message); }
    }

    private static MouseDownSnapshot? CaptureDirectMenuAtPointCore(int x, int y, bool useMsaa)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        if (pointer == 0) return null;
        var root = GetAncestor(pointer, 2);
        if (root == 0 || !IsWindowVisible(root)) return null;
        var className = new System.Text.StringBuilder(128);
        GetClassName(root, className, className.Capacity);
        if (CurrentMenuOwner() == 0 && className.ToString() != "#32768" &&
            (GetWindow(root, 4) == 0 ||
                !IsCompactAtPoint(root, x, y, Screen.FromPoint(new(x, y)).Bounds)))
            return null;
        var command = useMsaa ? MsaaCommandAtPoint(x, y) : null;
        if (command is not { Type: "ControlType.MenuItem" } item)
        {
            return CapturePopupMenuSnapshot(x, y);
        }
        var target = NativeControlRef(root, null, item.Name, item.Type, null);
        if (target is null) return null;
        return new(null, target with { ParentName = MsaaActions.MenuParentAtPoint(x, y) })
        { WindowHandle = root, InputBounds = new(item.X, item.Y, item.Width, item.Height) };
    }

    private static MouseDownSnapshot? CaptureOwnedPopupCommandAtPoint(int x, int y)
    {
        var hit = WindowFromPoint(new System.Drawing.Point(x, y));
        var popup = hit == 0 ? 0 : GetAncestor(hit, 2);
        if (popup == 0 || !IsWindowVisible(popup) ||
            GetWindow(popup, 4) == 0 ||
            !IsCompactAtPoint(popup, x, y, Screen.FromPoint(new(x, y)).Bounds))
            return null;
        GetWindowThreadProcessId(popup, out var processId);
        GetWindowThreadProcessId(GetWindow(popup, 4), out var ownerId);
        if (processId == 0 || processId != ownerId) return null;
        var command = MsaaCommandAtPoint(x, y);
        if (command is not { Type: "ControlType.Button" or "ControlType.CheckBox" or
                "ControlType.RadioButton" } action)
            return null;
        var target = NativeControlRef(popup, null, action.Name, action.Type, null);
        return target is null ? null : new(null, target) { WindowHandle = popup, IsDialog = true };
    }

    private static MouseDownSnapshot? CaptureNativeDialogCommandAtPoint(int x, int y)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        var screen = Screen.FromPoint(new System.Drawing.Point(x, y)).Bounds;
        var window = pointer;
        var direct = false;
        (string Name, string Type, int X, int Y, int Width, int Height)? command = null;
        for (var depth = 0; window != 0 && depth < 12;
            window = GetAncestor(window, 1), depth++)
        {
            if (!IsCompactAtPoint(window, x, y, screen)) continue;
            if (!IsNativeDialogSurface(window)) continue;
            var caption = new System.Text.StringBuilder(512);
            GetWindowText(window, caption, caption.Capacity);
            if (caption.Length == 0) continue;
            var className = new System.Text.StringBuilder(128);
            GetClassName(window, className, className.Capacity);
            if (className.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase) ||
                IsNativeEditClass(className.ToString()) ||
                className.ToString().Contains("Tab", StringComparison.OrdinalIgnoreCase)) continue;
            if (!direct)
            {
                command = MsaaCommandAtPoint(x, y);
                direct = true;
            }
            if (command is not { Type: "ControlType.Button" or "ControlType.RadioButton" } action)
                continue;
            var target = NativeControlRef(window, null, action.Name, action.Type, null);
            if (target is null || target.Window == action.Name) continue;
            return new(null, target) { WindowHandle = window, IsDialog = true };
        }
        return null;
    }

    private static MouseDownSnapshot? CaptureMouseDownTarget(int x, int y)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        var root = pointer == 0 ? 0 : GetAncestor(pointer, 2);
        if (root == 0) return null;
        try
        {
            var element = AutomationElement.FromPoint(new System.Windows.Point(x, y));
            var dialog = element.Current.ControlType == ControlType.TabItem ? null : WindowAncestor(element);
            if (dialog is not null && dialog.Current.NativeWindowHandle != 0 &&
                dialog.Current.ProcessId == element.Current.ProcessId &&
                dialog.Current.BoundingRectangle.Contains(x, y))
                root = (nint)dialog.Current.NativeWindowHandle;
            return CreateMouseDownSnapshot(element, root, x, y);
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    private static MouseDownSnapshot? CreateMouseDownSnapshot(AutomationElement element, nint root, int x, int y)
    {
        try
        {
            var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
            MouseDownSnapshot Snapshot(AutomationElement captured, ControlRef target)
            {
                var tab = captured.Current.ControlType == ControlType.TabItem;
                var dialog = tab ? null : WindowAncestor(captured);
                var screen = Screen.FromPoint(new System.Drawing.Point(x, y)).Bounds;
                var isDialog = false;
                var handle = root;
                if (tab)
                    isDialog = IsCompactAtPoint(root, x, y, screen) && IsNativeDialogSurface(root);
                if (dialog is not null)
                {
                    var state = dialog.Current;
                    isDialog = state.ProcessId == captured.Current.ProcessId &&
                        !string.IsNullOrWhiteSpace(state.Name) &&
                        state.BoundingRectangle.Contains(x, y) &&
                        state.BoundingRectangle.Width * state.BoundingRectangle.Height <
                            (long)screen.Width * screen.Height * 8 / 10;
                    if (isDialog)
                    {
                        target = target with { Window = state.Name };
                        if (state.NativeWindowHandle != 0) handle = (nint)state.NativeWindowHandle;
                    }
                }
                return new(captured, target)
                {
                    WindowHandle = handle,
                    IsDialog = isDialog
                };
            }
            if (NativeCloseAtPoint(root, x, y))
            {
                var close = NativeControlRef(root, null, "Close", "ControlType.Button", null);
                if (close is not null) return Snapshot(element, close);
            }
            var type = element.Current.ControlType;
            if (type == ControlType.Window || type == ControlType.Pane ||
                type == ControlType.Custom || type == ControlType.Tab || type == ControlType.Text ||
                type == ControlType.Group)
            {
                // A window-level hit is not a verified command. Resolve native
                // children while the clicked dialog still exists, not after delivery.
                var button = FindNativeActionableAtPoint(root, x, y, 0x0002, 30);
                if (button is { } action)
                {
                    var command = NativeControlRef(root, null, action.Name,
                        action.ClassName.Contains("Tab", StringComparison.OrdinalIgnoreCase)
                            ? "ControlType.TabItem" : NativeButtonControlType(action.Handle), action.ClassName);
                    if (command is not null)
                        return Snapshot(AutomationElement.FromHandle(action.Handle), command);
                }
                var msaa = pointer != 0 && (GetAncestor(pointer, 2) == root ||
                    GetAncestor(root, 2) == GetAncestor(pointer, 2))
                    ? MsaaCommandAtPoint(x, y) : null;
                if (msaa is { } accessible &&
                    NativeControlRef(root, null, accessible.Name, accessible.Type, null) is { } legacy)
                    return Snapshot(element, legacy);
                element = FindActionableControlAtPoint(element, x, y) ?? element;
            }
            if (element.Current.ControlType == ControlType.Image &&
                string.IsNullOrWhiteSpace(element.Current.Name) &&
                string.IsNullOrWhiteSpace(element.Current.AutomationId))
                element = NamedContainingImage(element, root, x, y) ?? element;
            var current = element.Current;
            if (current.IsOffscreen || !current.IsEnabled ||
                current.BoundingRectangle.IsEmpty || !current.BoundingRectangle.Contains(x, y) ||
                string.IsNullOrWhiteSpace(current.Name) && string.IsNullOrWhiteSpace(current.AutomationId))
                return null;
            var native = NativeControlRef(root, current.AutomationId, current.Name,
                current.ControlType.ProgrammaticName, current.ClassName);
            GetWindowThreadProcessId(root, out var processId);
            if (native is null || current.ProcessId != processId) return null;
            var parent = current.ControlType == ControlType.MenuItem
                ? Automation.RecordedMenuParentName(element)
                : TreeWalker.ControlViewWalker.GetParent(element)?.Current.Name;
            return Snapshot(element, native with { ParentName = parent, ProcessId = (int)processId });
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    private static AutomationElement? NamedContainingImage(AutomationElement hit, nint root, int x, int y)
    {
        GetWindowThreadProcessId(root, out var processId);
        AutomationElement? named = null;
        AutomationElement? element = hit;
        for (var depth = 0; element is not null && depth < 12; depth++)
        {
            var current = element.Current;
            if (current.ProcessId != processId) break;
            if (current.ControlType == ControlType.Image && !string.IsNullOrWhiteSpace(current.Name) &&
                current.IsEnabled && !current.IsOffscreen)
            {
                var bounds = ReadChartBounds(() => element.Current.BoundingRectangle);
                if (!bounds.IsEmpty && bounds.Contains(x, y)) named = element;
            }
            if (current.NativeWindowHandle == root) break;
            element = TreeWalker.RawViewWalker.GetParent(element);
        }
        return named;
    }

    private bool TryRecordNativeDialogClick(int x, int y)
    {
        var root = CompactWindowAtPoint(x, y);
        if (root == 0) return false;
        if (NativeCloseAtPoint(root, x, y))
        {
            var close = NativeControlRef(root, null, "Close", "ControlType.Button", null);
            if (close is not null && !IsOwnWindow(close) && !IsIgnored(close))
            {
                Add(new RecordedEvent(DateTimeOffset.Now, "click", close, null, null, null,
                    null, null, ClickX: x, ClickY: y));
                QueueNativeCloseConfirmation(events.Count - 1, root);
                return true;
            }
        }
        var child = WindowFromPoint(new System.Drawing.Point(x, y));
        if (TryDescribeNativeEdit(root, child, out var edit) && edit is not null)
        {
            var screen = Screen.FromPoint(new System.Drawing.Point(x, y));
            cachedEditableTarget = edit;
            cachedEditableElement = null;
            cachedEditableScreen = screen;
            cachedKeyboardWindow = root;
            cachedKeyboardFocus = child;
            cachedKeyboardTarget = edit;
            cachedKeyboardScreen = screen;
            Add(new RecordedEvent(DateTimeOffset.Now, "click", edit, null, null, null,
                null, null, ClickX: x, ClickY: y));
            return true;
        }
        var msaa = MsaaCommandAtPoint(x, y) ?? MsaaActions.CommandAtPoint(root, x, y);
        if (msaa is not null)
        {
            var command = NativeControlRef(root, null, msaa.Value.Name,
                msaa.Value.Type, null);
            if (command is not null && !IsOwnWindow(command) && !IsIgnored(command))
            {
                Add(new RecordedEvent(DateTimeOffset.Now, "click", command, null, null, null,
                    null, "msaa-command", ClickX: x, ClickY: y));
                if (command.Window == "Find and Replace")
                    QueueDialogFieldSnapshot(events.Count - 1, root);
                if (command.ControlType == "ControlType.TabItem") RefreshDialogCacheAfterTab(root);
                if (command.Name == "Close") QueueNativeCloseConfirmation(events.Count - 1, root);
                return true;
            }
        }
        var button = FindNativeActionableAtPoint(x, y);
        if (button is null) return false;
        var target = NativeControlRef(root, null, button.Value.Name,
            button.Value.ClassName.Contains("Tab", StringComparison.OrdinalIgnoreCase)
                ? "ControlType.TabItem" : NativeButtonControlType(button.Value.Handle),
            button.Value.ClassName);
        if (target is null || IsOwnWindow(target) || IsIgnored(target)) return false;
        Add(new RecordedEvent(DateTimeOffset.Now, "click", target, null, null, null,
            null, null, ClickX: x, ClickY: y));
        if (target.Window == "Find and Replace")
            QueueDialogFieldSnapshot(events.Count - 1, root);
        if (target.ControlType == "ControlType.TabItem") RefreshDialogCacheAfterTab(root);
        if (target.Name == "Close") QueueNativeCloseConfirmation(events.Count - 1, root);
        return true;
    }

    private void RefreshDialogCacheAfterTab(nint window)
    {
        InvalidateDialogCache(window);
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            await Task.Delay(180);
            InvalidateDialogCache(window);
            QueueDialogControlSnapshot(window);
        }));
    }

    private void InvalidateDialogCache(nint window)
    {
        lock (dialogCacheLock)
        {
            dialogControlCache.Remove(window);
            dialogCacheVersions[window] = dialogCacheVersions.GetValueOrDefault(window) + 1;
        }
    }

    private void QueueNativeCloseConfirmation(int index, nint root)
    {
        RecordedEvent recorded;
        lock (events) recorded = events[index];
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            for (var attempt = 0; attempt < 12; attempt++)
            {
                await Task.Delay(100);
                if (IsWindowVisible(root)) continue;
                lock (events)
                    if (index < events.Count && events[index].At == recorded.At &&
                        events[index].Target == recorded.Target)
                        events[index] = events[index] with { AfterState = "window-closed" };
                return;
            }
        }));
    }

    private static bool NativeCloseAtPoint(nint window, int x, int y)
    {
        if (window == 0 || CompactWindowAtPoint(x, y) != window) return false;
        var packedPoint = unchecked((nint)((y << 16) | (x & 0xffff)));
        return SendMessageTimeout(window, 0x0084, 0, packedPoint, 0x0002, 30,
            out var hit) != 0 && hit == 20; // WM_NCHITTEST / HTCLOSE
    }

    private (string Name, string Type)? CachedDialogControlAtPoint(int x, int y)
    {
        var window = CompactWindowAtPoint(x, y);
        lock (dialogCacheLock)
        {
            if (!dialogControlCache.TryGetValue(window, out var entry) ||
                !GetWindowRect(window, out var currentBounds) ||
                !currentBounds.Equals(entry.WindowBounds)) return null;
            var found = entry.Controls.Where(control => control.Bounds.Contains(x, y))
                .OrderBy(control => control.Bounds.Width * control.Bounds.Height).FirstOrDefault();
            return string.IsNullOrWhiteSpace(found.Name) ? null : (found.Name, found.Type);
        }
    }

    private MouseDownSnapshot? CachedDialogCommandAtPoint(int x, int y)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        var frame = pointer == 0 ? 0 : GetAncestor(pointer, 2);
        if (frame == 0) return null;
        lock (dialogCacheLock)
        {
            foreach (var (window, entry) in dialogControlCache)
            {
                if (!IsWindowVisible(window) || !IsWindowEnabled(window) ||
                    GetAncestor(window, 2) != frame ||
                    !GetWindowRect(window, out var bounds) || !bounds.Equals(entry.WindowBounds))
                    continue;
                var matches = entry.Controls.Where(control => control.Bounds.Contains(x, y) &&
                    control.Type is "ControlType.Button" or "ControlType.RadioButton" or
                        "ControlType.CheckBox" or "ControlType.TabItem" or "ControlType.Edit")
                    .OrderBy(control => control.Bounds.Width * control.Bounds.Height).ToArray();
                if (matches.Length == 0) continue;
                GetWindowThreadProcessId(window, out var processId);
                if (matches[0].Target.ProcessId != processId) continue;
                var title = new System.Text.StringBuilder(512);
                GetWindowText(window, title, title.Capacity);
                if (title.Length > 0 && title.ToString() != matches[0].Target.Window) continue;
                return new(null, matches[0].Target) { WindowHandle = window, IsDialog = true };
            }
        }
        return null;
    }

    private void PrimeVisibleDialogCaches()
    {
        if (capturePausedByUser || capturePausedForIntent) return;
        var foreground = GetForegroundWindow();
        if (foreground == 0) return;
        GetWindowThreadProcessId(foreground, out var processId);
        void Consider(nint window)
        {
            if (!IsWindowVisible(window) || !GetWindowRect(window, out var bounds)) return;
            if (!IsNativeDialogSurface(window)) return;
            var screen = Screen.FromPoint(new System.Drawing.Point(bounds.Left, bounds.Top)).Bounds;
            if ((long)(bounds.Right - bounds.Left) * (bounds.Bottom - bounds.Top) >=
                (long)screen.Width * screen.Height * 8 / 10) return;
            var caption = new System.Text.StringBuilder(512);
            GetWindowText(window, caption, caption.Capacity);
            if (caption.Length == 0 && GetWindow(window, 4) == 0) return;
            var className = new System.Text.StringBuilder(128);
            GetClassName(window, className, className.Capacity);
            if (className.ToString().Contains("Button", StringComparison.OrdinalIgnoreCase) ||
                IsNativeEditClass(className.ToString()) ||
                className.ToString().Contains("Tab", StringComparison.OrdinalIgnoreCase)) return;
            lock (dialogCacheLock)
                if (dialogControlCache.TryGetValue(window, out var entry) &&
                    entry.WindowBounds.Equals(bounds) &&
                    DateTime.UtcNow - entry.At < TimeSpan.FromMilliseconds(500)) return;
            QueueDialogControlSnapshot(window);
        }
        Consider(foreground);
        EnumChildWindows(foreground, (child, _) => { Consider(child); return true; }, 0);
        EnumWindows((window, _) =>
        {
            GetWindowThreadProcessId(window, out var owner);
            if (owner == processId) Consider(window);
            return true;
        }, 0);
    }

    private static bool IsNativeDialogSurface(nint window)
    {
        var style = GetWindowLong(window, -16);
        var owner = GetWindow(window, 4);
        var childCaption = (style & 0x40000000) != 0 &&
            (style & 0x00C00000) == 0x00C00000;
        var ownedCaption = owner != 0 && (style & 0x00C00000) == 0x00C00000 &&
            (IsWindowVisible(owner) || (style & 0x00020000) == 0);
        return childCaption || ownedCaption ||
            owner != 0 && IsWindowVisible(owner) && !IsWindowEnabled(owner);
    }

    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(nint window);

    private void PrimeDialogCache(nint window)
    {
        if (window == 0) return;
        if (window != lastForegroundWindow)
            enrichmentQueue.Writer.TryWrite(ReleaseClosedDialogHandlers);
        lock (dialogCacheLock)
            if (window == lastForegroundWindow && dialogControlCache.ContainsKey(window)) return;
        lastForegroundWindow = window;
        QueueDialogControlSnapshot(window);
    }

    private void QueueDialogControlSnapshot(nint window)
    {
        if (window == 0 || !IsNativeDialogSurface(window) ||
            !GetWindowRect(window, out var bounds)) return;
        var screen = Screen.FromPoint(new System.Drawing.Point(bounds.Left, bounds.Top)).Bounds;
        if ((long)(bounds.Right - bounds.Left) * (bounds.Bottom - bounds.Top) >
            (long)screen.Width * screen.Height * 8 / 10) return;
        int version;
        lock (dialogCacheLock)
        {
            if (!queuedDialogSnapshots.Add(window)) return;
            version = dialogCacheVersions.GetValueOrDefault(window);
        }
        enrichmentQueue.Writer.TryWrite(() =>
        {
            try
            {
                if (!IsWindowVisible(window)) return;
                var dialog = AutomationElement.FromHandle(window);
                RegisterDialogActionHandlers(window, dialog);
                var controls = new List<(string Name, string Type, System.Windows.Rect Bounds, ControlRef Target)>();
                foreach (var child in RawDialogControls(dialog))
                {
                    var state = child.Current;
                    if (state.IsOffscreen || (!IsActionableDialogElement(child) &&
                        state.ControlType != ControlType.Edit && state.ClassName != "EDTBX") ||
                        (string.IsNullOrWhiteSpace(state.Name) && string.IsNullOrWhiteSpace(state.AutomationId)) ||
                        state.BoundingRectangle.IsEmpty) continue;
                    var target = Automation.Describe(child);
                    if (target is null) continue;
                    controls.Add((state.Name, state.ClassName == "EDTBX" ? "ControlType.Edit" :
                        state.ControlType.ProgrammaticName, state.BoundingRectangle,
                        target with { Window = dialog.Current.Name,
                            ControlType = state.ClassName == "EDTBX" ? "ControlType.Edit" : target.ControlType }));
                }
                if (!GetWindowRect(window, out var finalBounds) || !finalBounds.Equals(bounds)) return;
                lock (dialogCacheLock)
                    if (dialogCacheVersions.GetValueOrDefault(window) == version)
                        dialogControlCache[window] = (DateTime.UtcNow, bounds, controls);
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
            finally { lock (dialogCacheLock) queuedDialogSnapshots.Remove(window); }
        });
    }

    private void ReleaseClosedDialogHandlers()
    {
        List<(AutomationElement Root, AutomationEventHandler Invoked, AutomationEventHandler Selected)> closed;
        lock (dialogCacheLock)
        {
            var handles = dialogActionHandlers.Keys.Where(handle => !IsWindowVisible(handle)).ToArray();
            closed = handles.Select(handle => dialogActionHandlers[handle]).ToList();
            foreach (var handle in handles)
            {
                dialogActionHandlers.Remove(handle);
                dialogControlCache.Remove(handle);
            }
        }
        foreach (var (root, invoked, selected) in closed)
        {
            try
            {
                System.Windows.Automation.Automation.RemoveAutomationEventHandler(InvokePattern.InvokedEvent, root, invoked);
                System.Windows.Automation.Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, root, selected);
            }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
        }
    }

    private void RegisterDialogActionHandlers(nint window, AutomationElement dialog)
    {
        lock (dialogCacheLock)
            if (dialogActionHandlers.ContainsKey(window)) return;
        AutomationEventHandler invoked = (sender, _) => RecordDialogAction(window, sender, "ControlType.Button");
        AutomationEventHandler selected = (sender, _) => RecordDialogAction(window, sender, "ControlType.TabItem");
        try
        {
            System.Windows.Automation.Automation.AddAutomationEventHandler(InvokePattern.InvokedEvent, dialog,
                TreeScope.Descendants, invoked);
            System.Windows.Automation.Automation.AddAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, dialog,
                TreeScope.Descendants, selected);
            lock (dialogCacheLock) dialogActionHandlers[window] = (dialog, invoked, selected);
        }
        catch (Exception ex)
        {
            try { System.Windows.Automation.Automation.RemoveAutomationEventHandler(InvokePattern.InvokedEvent, dialog, invoked); }
            catch { }
            System.Diagnostics.Trace.WriteLine(ex);
        }
    }

    private void RecordDialogAction(nint window, object sender, string expectedType)
    {
        if (!IsRecording || GetForegroundWindow() != window || !IsWindowVisible(window) ||
            sender is not AutomationElement element) return;
        ControlRef? action;
        try
        {
            var root = WindowAncestor(element);
            if (root is null || root.Current.NativeWindowHandle != window) return;
            action = Automation.Describe(element);
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return; }
        if (action is null || string.IsNullOrWhiteSpace(action.Name) ||
            IsOwnWindow(action) || IsIgnored(action)) return;
        var now = DateTimeOffset.Now;
        lock (events)
        {
            for (var index = events.Count - 1; index >= 0; index--)
            {
                var recorded = events[index];
                if (now - recorded.At > TimeSpan.FromSeconds(1)) break;
                if (recorded.Kind != "click" || recorded.Target?.Window != action.Window) continue;
                if (recorded.Target?.ControlType == action.ControlType && recorded.Target?.Name == action.Name)
                    return; // The pointer path already recorded this command.
                if (recorded.Target?.ControlType is not ("ControlType.Window" or "ControlType.Pane" or
                    "ControlType.Custom" or "ControlType.Tab")) break;
                events[index] = recorded with { Target = action, AfterState = "dialog-command" };
                if (expectedType == "ControlType.TabItem") RefreshDialogCacheAfterTab(window);
                return;
            }
        }
        // An accessibility event alone is not evidence of user input. Providers
        // also raise these while an application updates its own controls.
    }

    private void RemoveDialogActionHandlers()
    {
        List<(AutomationElement Root, AutomationEventHandler Invoked, AutomationEventHandler Selected)> handlers;
        lock (dialogCacheLock)
        {
            handlers = dialogActionHandlers.Values.ToList();
            dialogActionHandlers.Clear();
        }
        foreach (var (root, invoked, selected) in handlers)
        {
            try
            {
                System.Windows.Automation.Automation.RemoveAutomationEventHandler(InvokePattern.InvokedEvent, root, invoked);
                System.Windows.Automation.Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, root, selected);
            }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
        }
    }

    private void QueueDialogFieldSnapshot(int eventIndex, nint window)
    {
        if (window == 0) return;
        ControlRef? expected;
        lock (events) expected = eventIndex < events.Count ? events[eventIndex].Target : null;
        pendingFieldCaptures.RemoveAll(task => task.IsCompleted);
        pendingFieldCaptures.Add(Task.Run(() =>
        {
            for (var attempt = 0; attempt < 15; attempt++)
            {
                var live = ResolveDialogWindow(window, expected, null, null);
                if (live == 0) { Thread.Sleep(120); continue; }
                var fields = NativeDialogFields(live, 0x20, 80);
                if (fields.Count > 0) SaveDialogFieldSnapshot(eventIndex, fields);
                if (fields.Count < 2) CaptureDialogFieldsWithUia(eventIndex, live, fields);
                lock (events)
                    if (eventIndex < events.Count &&
                        events[eventIndex].BeforeState?.Contains("\"AutomationId\":\"21\"") == true)
                        return;
                Thread.Sleep(120);
            }
        }));
    }

    private void QueueDialogCommandRetry(int eventIndex, nint window, int x, int y)
    {
        if (window == 0) return;
        RecordedEvent? recorded;
        lock (events) recorded = eventIndex < events.Count ? events[eventIndex] : null;
        if (recorded is null) return;
        var expected = recorded.Target;
        pendingWindowCaptures.RemoveAll(task => task.IsCompleted);
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            for (var attempt = 0; attempt < 30; attempt++)
            {
                await Task.Delay(attempt == 0 ? 75 : 150);
                var live = ResolveDialogWindow(window, expected, x, y);
                if (live == 0) continue;
                var rawPoint = new System.Drawing.Point(x, y);
                var uia = FindDialogActionableAtPoint(live, rawPoint.X, rawPoint.Y);
                var action = uia is null ? FindNativeActionableAtPoint(live, x, y, 0x20, 80) : null;
                var msaa = uia is null && action is null ? MsaaActions.CommandAtPoint(live, x, y) : null;
                if (uia is null && action is null && msaa is null) continue;
                var corrected = uia;
                if (corrected is null && action is { } native)
                    corrected = NativeControlRef(live, null, native.Name,
                        native.ClassName.Contains("Tab", StringComparison.OrdinalIgnoreCase)
                            ? "ControlType.TabItem" : NativeButtonControlType(native.Handle),
                        native.ClassName);
                if (corrected is null && msaa is { } accessible)
                    corrected = NativeControlRef(live, null, accessible.Name, accessible.Type, null);
                if (corrected is null) return;
                lock (events)
                {
                    if (eventIndex >= events.Count || events[eventIndex].At != recorded.At ||
                        events[eventIndex].ClickX != x ||
                        events[eventIndex].ClickY != y ||
                        events[eventIndex].Target?.ControlType is not
                            ("ControlType.Window" or "ControlType.Pane" or "ControlType.Custom" or "ControlType.Tab"))
                        return;
                    events[eventIndex] = events[eventIndex] with { Target = corrected };
                }
                TargetCorrected?.Invoke(corrected);
                return;
            }
        }));
    }

    private static ControlRef? FindDialogActionableAtPoint(nint window, int x, int y)
    {
        try
        {
            var root = AutomationElement.FromHandle(window);
            var matches = root.FindAll(TreeScope.Descendants, Condition.TrueCondition)
                .Cast<AutomationElement>()
                .Select(element => (Element: element, State: element.Current))
                .Where(item => !item.State.IsOffscreen && item.State.IsEnabled &&
                    !string.IsNullOrWhiteSpace(item.State.Name) &&
                    !item.State.BoundingRectangle.IsEmpty &&
                    item.State.BoundingRectangle.Contains(x, y) &&
                    IsActionableDialogElement(item.Element))
                .OrderBy(item => item.State.BoundingRectangle.Width *
                    item.State.BoundingRectangle.Height)
                .FirstOrDefault();
            return matches.Element is null ? null : Automation.Describe(matches.Element);
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        {
            System.Diagnostics.Trace.WriteLine(ex);
            return null;
        }
    }

    private static nint ResolveDialogWindow(nint captured, ControlRef? expected,
        int? x, int? y)
    {
        bool Matches(nint candidate)
        {
            if (candidate == 0 || !IsWindow(candidate) ||
                !GetWindowRect(candidate, out var bounds) ||
                x is int px && (px < bounds.Left || px >= bounds.Right) ||
                y is int py && (py < bounds.Top || py >= bounds.Bottom)) return false;
            var identity = NativeControlRef(candidate, null, "", "ControlType.Window", null);
            return identity is not null &&
                string.Equals(identity.Window, expected?.Window, StringComparison.Ordinal) &&
                string.Equals(identity.Process, expected?.Process, StringComparison.OrdinalIgnoreCase);
        }
        if (x is int pointX && y is int pointY)
        {
            var atPoint = CompactWindowAtPoint(pointX, pointY);
            if (Matches(atPoint)) return atPoint;
        }
        if (Matches(captured)) return captured;
        var matches = new List<nint>();
        EnumWindows((candidate, _) =>
        {
            if (Matches(candidate)) matches.Add(candidate);
            return matches.Count < 2;
        }, 0);
        return matches.Count == 1 ? matches[0] : 0;
    }

    private void QueueDialogClickProbe(int eventIndex, nint window, int x, int y)
    {
        if (window == 0) return;
        RecordedEvent? recorded;
        lock (events) recorded = eventIndex < events.Count ? events[eventIndex] : null;
        if (recorded is null) return;
        var expected = recorded.Target;
        pendingProbeCaptures.RemoveAll(task => task.IsCompleted);
        pendingProbeCaptures.Add(Task.Run(async () =>
        {
            try
            {
                // Give command identification priority over this expensive diagnostic walk.
                await Task.Delay(2000);
                var atPoint = CompactWindowAtPoint(x, y);
                var live = ResolveDialogWindow(window, expected, x, y);
                var visible = live != 0 && IsWindowVisible(live);
                var native = WindowFromPoint(new System.Drawing.Point(x, y));
                var nativeClass = new System.Text.StringBuilder(128);
                if (native != 0) GetClassName(native, nativeClass, nativeClass.Capacity);
                var msaaPoint = live != 0 ? MsaaPointProbe(x, y) : null;
                var msaaChildren = live != 0 ? MsaaActions.Probe(live) : [];
                var observed = new List<object>();
                try
                {
                    if (live != 0)
                    {
                        var dialog = AutomationElement.FromHandle(live);
                        var controls = dialog.FindAll(TreeScope.Descendants, Condition.TrueCondition);
                        foreach (AutomationElement control in controls)
                        {
                            if (observed.Count >= 100) break;
                            try
                            {
                                var state = control.Current;
                                var bounds = state.BoundingRectangle;
                                observed.Add(new { state.Name, state.AutomationId, state.ClassName,
                                    Type = state.ControlType.ProgrammaticName, state.IsOffscreen,
                                    Actionable = IsActionableDialogElement(control),
                                    X = bounds.X, Y = bounds.Y, Width = bounds.Width, Height = bounds.Height });
                            }
                            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                                System.Runtime.InteropServices.COMException)
                            { System.Diagnostics.Trace.WriteLine(ex); }
                        }
                    }
                }
                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
                var probe = "dialog-probe:" + System.Text.Json.JsonSerializer.Serialize(new
                { X = x, Y = y, CapturedWindow = DescribeWindow(window, x, y),
                    PointWindow = DescribeWindow(atPoint, x, y),
                    LiveWindow = DescribeWindow(live, x, y), WindowVisible = visible,
                    NativeClass = nativeClass.ToString(), NativeChildren = NativeChildren(live),
                    Msaa = msaaPoint, MsaaChildren = msaaChildren,
                    Controls = observed });
                ControlRef? corrected = null;
                if (live != 0 && MsaaActions.CommandFromObservations(msaaChildren, x, y) is { } command)
                    corrected = NativeControlRef(live, null, command.Name, command.Type, null);
                var applied = false;
                lock (events)
                    if (eventIndex < events.Count && events[eventIndex].At == recorded.At)
                    {
                        var current = events[eventIndex];
                        if (corrected is not null && current.ClickX == x && current.ClickY == y &&
                            current.Target?.ControlType is "ControlType.Window" or "ControlType.Pane" or
                                "ControlType.Custom" or "ControlType.Tab" &&
                            string.Equals(current.Target.Window, corrected.Window, StringComparison.Ordinal) &&
                            string.Equals(current.Target.Process, corrected.Process, StringComparison.OrdinalIgnoreCase))
                        {
                            current = current with { Target = corrected };
                            applied = true;
                        }
                        events[eventIndex] = current with { Diagnostic = probe };
                    }
                if (applied) TargetCorrected?.Invoke(corrected!);
            }
            catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
        }));
    }

    private static object DescribeWindow(nint window, int x, int y)
    {
        var title = new System.Text.StringBuilder(512);
        if (window != 0 && IsWindow(window)) GetWindowText(window, title, title.Capacity);
        NativeRect bounds = default;
        var hasRect = window != 0 && GetWindowRect(window, out bounds);
        return new { Handle = window.ToInt64(), Exists = window != 0 && IsWindow(window),
            Visible = window != 0 && IsWindowVisible(window), Title = title.ToString(),
            Rect = hasRect ? new { bounds.Left, bounds.Top, bounds.Right, bounds.Bottom } : null,
            PointInRect = hasRect && x >= bounds.Left && x < bounds.Right &&
                y >= bounds.Top && y < bounds.Bottom };
    }

    private static List<object> NativeChildren(nint window)
    {
        var children = new List<object>();
        if (window == 0 || !IsWindow(window)) return children;
        EnumChildWindows(window, (child, _) =>
        {
            if (children.Count >= 100) return false;
            var classBuffer = new System.Text.StringBuilder(128);
            GetClassName(child, classBuffer, classBuffer.Capacity);
            var caption = new System.Text.StringBuilder(256);
            if (classBuffer.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase) ||
                classBuffer.ToString().Contains("Tab", StringComparison.OrdinalIgnoreCase))
                SendMessageTimeout(child, 0x000D, caption.Capacity, caption, 0x20, 80, out _);
            var hasRect = GetWindowRect(child, out var bounds);
            children.Add(new { Handle = child.ToInt64(), ClassName = classBuffer.ToString(),
                Name = caption.ToString(), Id = GetDlgCtrlID(child), Visible = IsWindowVisible(child),
                Rect = hasRect ? new { bounds.Left, bounds.Top, bounds.Right, bounds.Bottom } : null });
            return true;
        }, 0);
        return children;
    }

    private void CaptureDialogFieldsWithUia(int eventIndex, nint window,
        List<DialogFieldSnapshot> fields)
    {
        try
        {
                var dialog = AutomationElement.FromHandle(window);
                var candidates = dialog.FindAll(TreeScope.Descendants, new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                    new PropertyCondition(AutomationElement.ClassNameProperty, "EDTBX")));
                foreach (AutomationElement field in candidates)
                {
                    var state = field.Current;
                    if (state.IsPassword || string.IsNullOrEmpty(state.AutomationId)) continue;
                    var value = field.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern)
                        ? ((ValuePattern)pattern).Current.Value : state.Name;
                    if (value is null || value == dialog.Current.Name) continue;
                    if (Automation.Describe(field) is { } target &&
                        !fields.Any(saved => saved.Target.AutomationId == target.AutomationId))
                        fields.Add(new DialogFieldSnapshot(target, value));
                }
                if (fields.Count > 0) SaveDialogFieldSnapshot(eventIndex, fields);
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
    }

    private void SaveDialogFieldSnapshot(int eventIndex, List<DialogFieldSnapshot> fields)
    {
        var snapshot = "dialog-fields:" + System.Text.Json.JsonSerializer.Serialize(fields);
        lock (events)
            if (eventIndex < events.Count)
            {
                var previous = events[eventIndex].BeforeState;
                if (previous?.StartsWith("dialog-fields:", StringComparison.Ordinal) == true)
                {
                    try
                    {
                        var existing = System.Text.Json.JsonSerializer.Deserialize<List<DialogFieldSnapshot>>(
                            previous["dialog-fields:".Length..]);
                        if (existing is not null && existing.Count >= fields.Count) return;
                    }
                    catch (System.Text.Json.JsonException) { }
                }
                events[eventIndex] = events[eventIndex] with { BeforeState = snapshot };
            }
    }

    private static List<DialogFieldSnapshot> NativeDialogFields(nint window,
        uint flags = 0x0002, uint timeout = 40)
    {
        var fields = new List<DialogFieldSnapshot>();
        EnumChildWindows(window, (child, _) =>
        {
            var classBuffer = new System.Text.StringBuilder(128);
            GetClassName(child, classBuffer, classBuffer.Capacity);
            var className = classBuffer.ToString();
            if (!IsNativeEditClass(className)) return true;
            var id = GetDlgCtrlID(child);
            if (id <= 0 || id >= 65536 ||
                SendMessageTimeout(child, 0x000E, 0, 0, flags, timeout, out var length) == 0 ||
                length is < 0 or > 32767) return true;
            var value = new System.Text.StringBuilder((int)length + 1);
            if (SendMessageTimeout(child, 0x000D, value.Capacity, value, flags, timeout, out _) == 0)
                return true;
            var target = NativeControlRef(window, id.ToString(), value.ToString(),
                className == "Edit" ? "ControlType.Edit" : "ControlType.Pane", className);
            if (target is not null) fields.Add(new DialogFieldSnapshot(target, value.ToString()));
            return true;
        }, 0);
        return fields.DistinctBy(field => field.Target.AutomationId).ToList();
    }

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

    private static AutomationElement? FindActionableControlAtPoint(AutomationElement hit, int x, int y)
    {
        try
        {
            if (hit.Current.ClassName == "EDTBX" || hit.Current.ControlType == ControlType.Edit ||
                hit.Current.ControlType != ControlType.Window &&
                    hit.TryGetCurrentPattern(ValuePattern.Pattern, out _)) return null;
            var dialog = WindowAncestor(hit);
            if (dialog is null || IsActionableDialogElement(hit))
                return null;
            var windowBounds = dialog.Current.BoundingRectangle;
            var screenBounds = Screen.FromPoint(new System.Drawing.Point(x, y)).Bounds;
            if (windowBounds.Width * windowBounds.Height > screenBounds.Width * screenBounds.Height * 0.8)
                return null;
            var ancestor = TreeWalker.RawViewWalker.GetParent(hit);
            for (var depth = 0; ancestor is not null && depth < 6 && ancestor != dialog; depth++)
            {
                if (IsActionableDialogElement(ancestor) &&
                    ancestor.Current.ProcessId == dialog.Current.ProcessId &&
                    ancestor.Current.BoundingRectangle.Contains(x, y))
                    return ancestor;
                ancestor = TreeWalker.RawViewWalker.GetParent(ancestor);
            }
            // Some dialogs report their window or pane at a child control's point.
            // Search only this small dialog and verify the control contains the click.
            var nativeChild = WindowFromPoint(new System.Drawing.Point(x, y));
            if (nativeChild != 0)
            {
                var nativeElement = AutomationElement.FromHandle(nativeChild);
                if (IsActionableDialogElement(nativeElement) &&
                    nativeElement.Current.ProcessId == dialog.Current.ProcessId &&
                    nativeElement.Current.BoundingRectangle.Contains(x, y))
                    return nativeElement;
            }
            // Raw View includes owner-drawn option controls hidden from Control View.
            var controls = new List<AutomationElement>();
            var queue = new Queue<AutomationElement>();
            queue.Enqueue(dialog);
            while (queue.Count > 0 && controls.Count < 300)
            {
                var current = queue.Dequeue();
                controls.Add(current);
                var child = TreeWalker.RawViewWalker.GetFirstChild(current);
                while (child is not null && queue.Count + controls.Count < 300)
                {
                    queue.Enqueue(child);
                    child = TreeWalker.RawViewWalker.GetNextSibling(child);
                }
            }
            return controls
                .Where(control => IsActionableDialogElement(control) &&
                    control.Current.BoundingRectangle.Contains(x, y))
                .OrderBy(control => control.Current.BoundingRectangle.Width * control.Current.BoundingRectangle.Height)
                .FirstOrDefault();
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException)
        { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    private static bool IsActionableDialogElement(AutomationElement element)
    {
        try
        {
            var state = element.Current;
            if (state.ControlType == ControlType.Window || state.ControlType == ControlType.Group)
                return false;
            return state.IsEnabled && !state.IsOffscreen &&
                !string.IsNullOrWhiteSpace(state.Name) &&
                (state.ControlType == ControlType.Button || state.ControlType == ControlType.TabItem ||
                 state.ControlType == ControlType.RadioButton || state.ControlType == ControlType.CheckBox ||
                 element.TryGetCurrentPattern(InvokePattern.Pattern, out _) ||
                 element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out _) ||
                 element.TryGetCurrentPattern(TogglePattern.Pattern, out _));
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { return false; }
    }

    private static IEnumerable<AutomationElement> RawDialogControls(AutomationElement dialog)
    {
        var queue = new Queue<AutomationElement>();
        queue.Enqueue(dialog);
        var visited = 0;
        while (queue.Count > 0 && visited < 300)
        {
            var current = queue.Dequeue();
            visited++;
            yield return current;
            var child = TreeWalker.RawViewWalker.GetFirstChild(current);
            while (child is not null && visited + queue.Count < 300)
            {
                queue.Enqueue(child);
                child = TreeWalker.RawViewWalker.GetNextSibling(child);
            }
        }
    }

    private static (nint Handle, string Name, string ClassName)? FindNativeActionableAtPoint(int x, int y)
    {
        var root = CompactWindowAtPoint(x, y);
        if (root == 0) return null;
        return FindNativeActionableAtPoint(root, x, y, 0x0002, 30);
    }

    private static string NativeButtonControlType(nint handle)
    {
        var style = GetWindowLong(handle, -16) & 0x000F;
        return style switch
        {
            0x0002 or 0x0003 or 0x0005 or 0x0006 => "ControlType.CheckBox",
            0x0004 or 0x0009 => "ControlType.RadioButton",
            _ => "ControlType.Button"
        };
    }

    private static (nint Handle, string Name, string ClassName)? FindNativeActionableAtPoint(
        nint root, int x, int y, uint flags, uint timeout)
    {
        if (!IsWindow(root) || !GetWindowRect(root, out var rootBounds) ||
            x < rootBounds.Left || x >= rootBounds.Right ||
            y < rootBounds.Top || y >= rootBounds.Bottom) return null;
        var candidates = new List<(nint Handle, string Name, string ClassName, long Area)>();
        void Consider(nint child)
        {
            if (child == 0 || child == root || GetAncestor(child, 2) != root ||
                !GetWindowRect(child, out var bounds) ||
                x < bounds.Left || x >= bounds.Right || y < bounds.Top || y >= bounds.Bottom)
                return;
            var classBuffer = new System.Text.StringBuilder(128);
            GetClassName(child, classBuffer, classBuffer.Capacity);
            var className = classBuffer.ToString();
            if (!className.Equals("Button", StringComparison.OrdinalIgnoreCase) &&
                !className.Contains("Tab", StringComparison.OrdinalIgnoreCase)) return;
            if (className.Equals("Button", StringComparison.OrdinalIgnoreCase) &&
                (GetWindowLong(child, -16) & 0x000F) == 0x0007) return; // BS_GROUPBOX is a label, not a command.
            var caption = new System.Text.StringBuilder(256);
            if (SendMessageTimeout(child, 0x000D, caption.Capacity, caption, flags, timeout, out _) == 0)
                return;
            var name = caption.ToString().Replace("&", "").Trim();
            if (string.IsNullOrWhiteSpace(name)) return;
            candidates.Add((child, name, className,
                (long)(bounds.Right - bounds.Left) * (bounds.Bottom - bounds.Top)));
        }
        Consider(WindowFromPoint(new System.Drawing.Point(x, y)));
        if (candidates.Count == 0)
            EnumChildWindows(root, (child, _) => { Consider(child); return true; }, 0);
        var best = candidates.OrderBy(candidate => candidate.Area).FirstOrDefault();
        return best.Handle == 0 ? null : (best.Handle, best.Name, best.ClassName);
    }

    private static (string Name, string Type, int X, int Y, int Width, int Height)? MsaaCommandAtPoint(int x, int y)
    {
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                out var child) != 0 || accessible is null) return null;
            var role = accessible.get_accRole(child);
            var type = role is int value ? value switch
            {
                0x0C => "ControlType.MenuItem", // ROLE_SYSTEM_MENUITEM
                0x2B => "ControlType.Button", // ROLE_SYSTEM_PUSHBUTTON
                0x25 => "ControlType.TabItem", // ROLE_SYSTEM_PAGETAB
                0x2C => "ControlType.CheckBox",
                0x2D => "ControlType.RadioButton",
                _ => null
            } : null;
            if (type is null) return null;
            var name = accessible.get_accName(child)?.Replace("&", "").Trim();
            if (string.IsNullOrWhiteSpace(name)) return null;
            var state = accessible.get_accState(child);
            if (state is int flags && (flags & (0x1 | 0x8000 | 0x10000)) != 0)
                return null; // Unavailable, invisible, or offscreen.
            accessible.accLocation(out var left, out var top, out var width, out var height, child);
            if (width <= 0 || height <= 0 || x < left || x >= left + width ||
                y < top || y >= top + height) return null;
            return (name, type, left, top, width, height);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine(ex);
            return null;
        }
    }

    private static (string Name, string Type, int X, int Y, int Width, int Height)? RowMenuCommandAtPoint(int x, int y)
    {
        var direct = MsaaCommandAtPoint(x, y);
        if (direct is { Type: "ControlType.MenuItem" }) return direct;
        var menuOwner = CurrentMenuOwner();
        var pointerWindow = WindowFromPoint(new System.Drawing.Point(x, y));
        var root = pointerWindow == 0 ? 0 : GetAncestor(pointerWindow, 2);
        foreach (var window in new[] { menuOwner, pointerWindow, root, GetForegroundWindow() }
            .Where(handle => handle != 0).Distinct())
        {
            if (!string.Equals(DescribeNativeWindowOwner(window), "EXCEL",
                StringComparison.OrdinalIgnoreCase)) continue;
            var command = MsaaActions.MenuCommandAtPoint(window, x, y);
            if (command is { Type: "ControlType.MenuItem" }) return command;
        }
        return null;
    }

    private static nint CurrentMenuOwner()
    {
        var current = new GuiThreadInfo { Size = (uint)Marshal.SizeOf<GuiThreadInfo>() };
        if (GetGUIThreadInfo(0, ref current) && current.MenuOwner != 0)
            return current.MenuOwner;
        var foreground = GetForegroundWindow();
        if (foreground == 0) return 0;
        var thread = GetWindowThreadProcessId(foreground, out _);
        if (thread == 0) return 0;
        var info = new GuiThreadInfo { Size = (uint)Marshal.SizeOf<GuiThreadInfo>() };
        return GetGUIThreadInfo(thread, ref info) ? info.MenuOwner : 0;
    }

    private static string MenuProbeDiagnostic(int x, int y)
    {
        var pointer = WindowFromPoint(new System.Drawing.Point(x, y));
        var root = pointer == 0 ? 0 : GetAncestor(pointer, 2);
        var windows = new[] { CurrentMenuOwner(), pointer, root, GetForegroundWindow() }
            .Where(handle => handle != 0).Distinct();
        return string.Join("|", windows.Select(window =>
        {
            var items = MsaaActions.ProbeMenu(window).Where(item => item.Role == 0x0C)
                .Take(12).Select(item => $"{item.Name}@{item.X},{item.Y},{item.Width},{item.Height}");
            return DescribeNativeWindow(window) + ":[" + string.Join(",", items) + "]";
        }));
    }

    private void QueueRowMenuDiagnostic(int index, int x, int y, DateTimeOffset at)
    {
        pendingProbeCaptures.RemoveAll(task => task.IsCompleted);
        pendingProbeCaptures.Add(Task.Run(() =>
        {
            string evidence;
            try { evidence = MenuProbeDiagnostic(x, y); }
            catch (Exception ex)
            {
                evidence = "probe-error:" + ex.GetType().Name;
                System.Diagnostics.Trace.WriteLine(ex);
            }
            lock (events)
            {
                if (index < 0 || index >= events.Count) return;
                var current = events[index];
                if (current.Kind != "unresolved-click" || current.At != at ||
                    current.ClickX != x || current.ClickY != y ||
                    current.Target?.Name != "row menu choice") return;
                events[index] = current with { Diagnostic =
                    (current.Diagnostic ?? "") + ";menu-items=" + evidence };
            }
        }));
    }

    private static RowAnchorSnapshot? TryReadRowAnchor(ControlRef rowTarget, int row)
    {
        string? CellValue(int number)
        {
            try
            {
                var id = "B" + number;
                var target = new ControlRef("EXCEL", rowTarget.Window, id, id,
                    "ControlType.DataItem", "XLSpreadsheetCell", null);
                var cell = Automation.Resolve(target);
                return cell?.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern) == true
                    ? ((ValuePattern)pattern).Current.Value : null;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException)
            { System.Diagnostics.Trace.WriteLine(ex); return null; }
        }
        var current = CellValue(row);
        var next = CellValue(row + 1);
        var following = CellValue(row + 2);
        return current is not null && next is not null && following is not null &&
            current != next && next != following
            ? new RowAnchorSnapshot(row, current, next, following) : null;
    }

    private void QueueRowDeletionOutcome(int index, DateTimeOffset at, ControlRef rowTarget,
        RowAnchorSnapshot before)
    {
        pendingProbeCaptures.RemoveAll(task => task.IsCompleted);
        pendingProbeCaptures.Add(Task.Run(async () =>
        {
            for (var attempt = 0; attempt < 12; attempt++)
            {
                await Task.Delay(150);
                var after = TryReadRowAnchor(rowTarget, before.Row);
                if (after?.Current != before.Next || after.Next != before.Following) continue;
                lock (events)
                {
                    if (index < 0 || index >= events.Count) return;
                    var current = events[index];
                    if (current.Kind != "unresolved-click" || current.At != at ||
                        current.Target?.Name != "row menu choice") return;
                    events[index] = current with { Kind = "delete-row", Target = rowTarget,
                        BeforeState = "row-anchor-verified", AfterState = "row-anchor-shifted",
                        Diagnostic = (current.Diagnostic ?? "") + ";verified-row-delete" };
                }
                TargetCorrected?.Invoke(rowTarget with { Name = "Delete row " + before.Row });
                return;
            }
        }));
    }

    private static object? MsaaPointProbe(int x, int y)
    {
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                out var child) != 0 || accessible is null) return null;
            var name = accessible.get_accName(child);
            var role = accessible.get_accRole(child);
            var state = accessible.get_accState(child);
            accessible.accLocation(out var left, out var top, out var width, out var height, child);
            return new { Name = name, Role = role, State = state, X = left, Y = top,
                Width = width, Height = height };
        }
        catch (Exception ex) { return new { Error = ex.GetType().Name }; }
    }

    private static AutomationElement? SelectedTab(AutomationElement root, bool sheetOnly = true)
        => Automation.SelectedTab(root, sheetOnly);

    private void FinishTypingCapture()
    {
        if (Interlocked.Exchange(ref finishingTypingCapture, 1) != 0) return;
        try
        {
            typingCapture.Stop();
            for (var index = events.Count - 1; index >= 0 && events[index].Kind == "key"; index--)
            {
                if (ActionGrouper.IsStandaloneModifier(events[index].Key)) continue;
                if (events[index].Target is { } target && IsEditableTarget(target, cachedEditableElement) &&
                    events[index].AfterState?.StartsWith("excel-selection-range:", StringComparison.Ordinal) != true &&
                    events[index].Key is not ("Enter" or "Return"))
                    CaptureCompletedField(index, target, cachedEditableElement);
                if (events[index].Key is "Enter" or "Return" &&
                    events[index].Target?.Process?.Equals("EXCEL", StringComparison.OrdinalIgnoreCase) == true)
                {
                    try
                    {
                        var window = GetForegroundWindow();
                        var selected = window == 0 ? null : SelectedTab(AutomationElement.FromHandle(window));
                        if (selected is not null)
                            events[index] = events[index] with { AfterState = "sheet-name:" + selected.Current.Name };
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                }
                if (events[index].Screenshot is null && events[index].Target?.Window != "Find and Replace")
                    events[index] = events[index] with { Screenshot = Capture(typingScreen) };
                break;
            }
        }
        finally { Volatile.Write(ref finishingTypingCapture, 0); }
    }

    private void CaptureCompletedField(int index, ControlRef target, AutomationElement? knownField)
    {
        // Read the native edit while it still has focus. A queued UIA read can run
        // after the next click and capture a different field or a partial value.
        if (cachedKeyboardFocus != 0 && target.Window == "Find and Replace" &&
            cachedKeyboardTarget?.AutomationId == target.AutomationId &&
            GetAncestor(cachedKeyboardFocus, 2) == cachedKeyboardWindow &&
            SendMessageTimeout(cachedKeyboardFocus, 0x000E, 0, 0, 0x0002, 40,
                out var lengthResult) != 0 && lengthResult is > 0 and < 32768)
        {
            var text = new System.Text.StringBuilder((int)lengthResult + 1);
            if (SendMessageTimeout(cachedKeyboardFocus, 0x000D, text.Capacity, text,
                0x0002, 40, out _) != 0)
            {
                lock (events) events[index] = events[index] with { AfterState = "field-value:" + text };
                return;
            }
        }
        pendingFieldCaptures.RemoveAll(task => task.IsCompleted);
        pendingFieldCaptures.Add(Task.Run(() =>
        {
            try
            {
                var field = knownField;
                if (field is null || field.Current.AutomationId != target.AutomationId)
                    field = Automation.Resolve(target);
                if (field is null || field.Current.IsPassword) return;
                if (field.Current.AutomationId != target.AutomationId) return;
                var value = field.TryGetCurrentPattern(ValuePattern.Pattern, out var pattern)
                    ? ((ValuePattern)pattern).Current.Value : field.Current.Name;
                if (value is null || value == target.Window) return;
                lock (events) events[index] = events[index] with { AfterState = "field-value:" + value };
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        }));
    }

    private static string? SelectionCount(AutomationElement element)
    {
        try
        {
            var row = Automation.ListItemAncestor(element);
            if (row?.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var itemPattern) != true) return null;
            var list = ((SelectionItemPattern)itemPattern).Current.SelectionContainer;
            return list?.TryGetCurrentPattern(SelectionPattern.Pattern, out var selectionPattern) == true
                ? "selection-count:" + ((SelectionPattern)selectionPattern).Current.GetSelection().Length : null;
        }
        catch (Exception e) when (e is ElementNotAvailableException or InvalidOperationException) { return null; }
    }

    private void CompleteScroll()
    {
        scrollCapture.Stop();
        if (scrollEventIndex >= 0 && scrollEventIndex < events.Count)
        {
            var index = scrollEventIndex;
            var recorded = events[index];
            var element = scrollElement;
            var screen = scrollScreen;
            var axis = scrollAxis;
            pendingScrollCaptures.Add(Task.Run(() =>
            {
                try
                {
                    var bar = (recorded.Value is not null || recorded.Target?.ControlType == "ControlType.Thumb") &&
                        recorded.Target is not null && axis is not null
                        ? Automation.FindScrollBar(recorded.Target, axis) : null;
                    var updated = recorded with
                    { Target = Automation.Describe(bar) ?? recorded.Target,
                      AfterState = Automation.ScrollState(bar ?? element), Screenshot = Capture(screen) };
                    lock (events) events[index] = updated;
                }
                catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
            }));
        }
        scrollEventIndex = -1;
        scrollElement = null;
        scrollScreen = null;
        scrollAxis = null;
    }

    private static Screen ScreenFor(AutomationElement element)
    {
        try
        {
            var bounds = element.Current.BoundingRectangle;
            if (bounds.Width > 0 && bounds.Height > 0)
                return Screen.FromRectangle(System.Drawing.Rectangle.Round(new System.Drawing.RectangleF(
                    (float)bounds.Left, (float)bounds.Top, (float)bounds.Width, (float)bounds.Height)));
        }
        catch (ElementNotAvailableException) { }
        return Screen.PrimaryScreen ?? Screen.AllScreens[0];
    }

    private static bool IsDesktopBackground(ControlRef? target) => target is
        { Process: "explorer", Window: "Program Manager", Name: "Desktop",
          ControlType: "ControlType.List", ClassName: "SysListView32" };

    private static bool IsWeakSwitchTarget(ControlRef? target) => IsDesktopBackground(target) ||
        target is { ControlType: "ControlType.Window" or "ControlType.Pane" or
            "ControlType.Custom" or "ControlType.TitleBar" };

    private void QueueSearchResultRetry(int index, int x, int y)
    {
        DateTimeOffset at;
        lock (events) at = index < events.Count ? events[index].At : default;
        pendingWindowCaptures.RemoveAll(task => task.IsCompleted);
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            for (var attempt = 0; attempt < 12; attempt++)
            {
                await Task.Delay(100);
                if (!string.Equals(DescribeNativeWindowOwner(GetForegroundWindow()), "SearchHost",
                    StringComparison.OrdinalIgnoreCase)) return;
                try
                {
                    var element = AutomationElement.FromPoint(new System.Windows.Point(x, y));
                    var candidate = Automation.Describe(element);
                    if (candidate is not { Process: "SearchHost", Name: { Length: > 0 } } ||
                        IsWeakSwitchTarget(candidate) || IsOwnWindow(candidate) || IsIgnored(candidate))
                        continue;
                    var bounds = element.Current.BoundingRectangle;
                    if (bounds.IsEmpty || !bounds.Contains(x, y)) continue;
                    lock (events)
                    {
                        if (index < 0 || index >= events.Count) return;
                        var current = events[index];
                        if (current.At != at || current.Kind != "unresolved-click" ||
                            current.Target?.Name != "Windows Search result") return;
                        events[index] = current with { Kind = "click", Target = candidate,
                            Diagnostic = (current.Diagnostic ?? "") + ";reidentified-in-windows-search" };
                    }
                    TargetCorrected?.Invoke(candidate);
                    return;
                }
                catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                    InvalidOperationException or System.Runtime.InteropServices.COMException)
                { System.Diagnostics.Trace.WriteLine(ex); }
            }
            lock (events)
            {
                if (index >= 0 && index < events.Count && events[index].At == at &&
                    events[index].Target?.Name == "Windows Search result")
                    events[index] = events[index] with { Diagnostic =
                        (events[index].Diagnostic ?? "") + ";search-result-retry-exhausted" };
            }
        }));
    }

    private void QueueSearchLaunchOutcome(int index, nint searchWindow)
    {
        DateTimeOffset at;
        lock (events) at = index < events.Count ? events[index].At : default;
        pendingWindowCaptures.RemoveAll(task => task.IsCompleted);
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            nint stableForeground = 0;
            var stableCount = 0;
            for (var attempt = 0; attempt < 60; attempt++)
            {
                await Task.Delay(100);
                if (searchWindow != 0 && IsWindowVisible(searchWindow)) continue;
                var foreground = GetForegroundWindow();
                if (foreground == 0 || foreground == searchWindow || !IsWindowVisible(foreground)) continue;
                var process = DescribeNativeWindowOwner(foreground);
                if (string.IsNullOrWhiteSpace(process) ||
                    process.Equals("SearchHost", StringComparison.OrdinalIgnoreCase) ||
                    process.Equals("explorer", StringComparison.OrdinalIgnoreCase) ||
                    process.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath),
                        StringComparison.OrdinalIgnoreCase)) continue;
                var title = new System.Text.StringBuilder(512);
                GetWindowText(foreground, title, title.Capacity);
                if (string.IsNullOrWhiteSpace(title.ToString())) continue;
                stableCount = foreground == stableForeground ? stableCount + 1 : 1;
                stableForeground = foreground;
                if (stableCount < 3) continue;
                lock (events)
                {
                    if (index < 0 || index >= events.Count || events[index].At != at) return;
                    var current = events[index];
                    if (current.Target?.Process != "SearchHost" ||
                        current.Kind is not ("key" or "unresolved-click" or "click")) return;
                    events[index] = current with
                    {
                        AfterState = "launched-window:" + process + "|" + title
                    };
                }
                return;
            }
        }));
    }

    private void QueueApplicationSwitchRetry(int index, int x, int y, string? expectedProcess)
    {
        if (string.IsNullOrWhiteSpace(expectedProcess)) return;
        pendingWindowCaptures.RemoveAll(task => task.IsCompleted);
        pendingWindowCaptures.Add(Task.Run(async () =>
        {
            var lastReason = "no candidate observed";
            for (var attempt = 0; attempt < 10; attempt++)
            {
                await Task.Delay(100);
                try
                {
                    var foregroundOwner = DescribeNativeWindowOwner(GetForegroundWindow());
                    if (!string.Equals(foregroundOwner, expectedProcess,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        lastReason = $"foreground={foregroundOwner ?? "none"}";
                        continue;
                    }
                    var element = AutomationElement.FromPoint(new System.Windows.Point(x, y));
                    var candidate = Automation.Describe(element);
                    if (candidate is null || IsDesktopBackground(candidate) ||
                        !string.Equals(candidate.Process, expectedProcess, StringComparison.OrdinalIgnoreCase) ||
                        IsWeakSwitchTarget(candidate) || string.IsNullOrWhiteSpace(candidate.Name) ||
                        IsOwnWindow(candidate) || IsIgnored(candidate))
                    {
                        lastReason = $"candidate={candidate?.Process ?? "none"}/" +
                            $"{candidate?.ControlType ?? "none"}/{candidate?.Name ?? ""}";
                        continue;
                    }
                    var bounds = element.Current.BoundingRectangle;
                    if (bounds.IsEmpty || !bounds.Contains(x, y))
                    {
                        lastReason = "candidate outside recorded point";
                        continue;
                    }
                    lock (events)
                    {
                        if (index >= events.Count || !IsWeakSwitchTarget(events[index].Target)) return;
                        events[index] = events[index] with
                        { Target = candidate, Diagnostic = events[index].Diagnostic + ";reidentified-after-app-switch" };
                    }
                    TargetCorrected?.Invoke(candidate);
                    return;
                }
                catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                    InvalidOperationException or System.Runtime.InteropServices.COMException)
                { lastReason = ex.GetType().Name; System.Diagnostics.Trace.WriteLine(ex); }
            }
            lock (events)
            {
                if (index < events.Count && IsWeakSwitchTarget(events[index].Target))
                    events[index] = events[index] with
                    {
                        Diagnostic = (events[index].Diagnostic ?? "") +
                            $";app-switch-retry-exhausted:expected={expectedProcess};" +
                            $"foreground={DescribeNativeWindowOwner(GetForegroundWindow()) ?? "none"};" +
                            $"pointer={DescribeNativeWindow(WindowFromPoint(new System.Drawing.Point(x, y)))};" +
                            $"reason={lastReason}"
                    };
            }
        }));
    }

    private static string? DescribeNativeWindowOwner(nint window)
    {
        if (window == 0) return null;
        GetWindowThreadProcessId(window, out var processId);
        if (processId == 0) return null;
        try { return System.Diagnostics.Process.GetProcessById((int)processId).ProcessName; }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    private static string DescribeNativeWindow(nint window)
    {
        if (window == 0) return "none";
        var className = new System.Text.StringBuilder(128);
        var title = new System.Text.StringBuilder(256);
        GetClassName(window, className, className.Capacity);
        GetWindowText(window, title, title.Capacity);
        return $"{DescribeNativeWindowOwner(window) ?? "unknown"}/{className}/{title}";
    }

    private string? Capture(Screen? screen)
    {
        if (directory is null) return null;
        try
        {
            var bounds = (screen ?? Screen.PrimaryScreen ?? Screen.AllScreens[0]).Bounds;
            using var bitmap = new Bitmap(bounds.Width, bounds.Height);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size);
            var name = $"screen-{Interlocked.Increment(ref imageNumber):0000}.jpg";
            bitmap.Save(Path.Combine(directory, name), ImageFormat.Jpeg);
            return name;
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    private void Add(RecordedEvent item) { lock (events) events.Add(item); Captured?.Invoke(item); }
    public void Dispose()
    {
        Stop();
        enrichmentQueue.Writer.TryComplete();
        try { enrichmentWorker.Wait(TimeSpan.FromSeconds(2)); }
        catch (AggregateException ex) { System.Diagnostics.Trace.WriteLine(ex); }
        typingCapture.Dispose(); clickCapture.Dispose(); scrollCapture.Dispose(); chartHoverCapture.Dispose(); legendOutcomeCapture.Dispose();
        chartSampleRequested.Dispose();
        hookDispatcher.Dispose();
    }

    private static string? TextCharacter(KeyboardInfo info, string modifiers)
    {
        if (modifiers.Contains("Control", StringComparison.OrdinalIgnoreCase) ||
            modifiers.Contains("Alt", StringComparison.OrdinalIgnoreCase)) return null;
        if (modifiers.Contains("Control") || modifiers.Contains("Alt")) return null;
        var state = new byte[256];
        if (!GetKeyboardState(state)) return null;
        if (modifiers.Contains("Shift", StringComparison.OrdinalIgnoreCase))
        {
            state[(int)Keys.ShiftKey] |= 0x80;
            state[(int)Keys.LShiftKey] |= 0x80;
        }
        state[info.VirtualKey] |= 0x80;
        var buffer = new System.Text.StringBuilder(8);
        var count = ToUnicodeEx(info.VirtualKey, info.ScanCode, state, buffer, buffer.Capacity, 0, GetKeyboardLayout(0));
        return count == 1 && !char.IsControl(buffer[0]) ? buffer[0].ToString() : null;
    }

    private static nint ForegroundFocusHandle()
    {
        var info = new GuiThreadInfo { Size = (uint)Marshal.SizeOf<GuiThreadInfo>() };
        return GetGUIThreadInfo(0, ref info) ? info.Focus : 0;
    }

    [StructLayout(LayoutKind.Sequential)] private struct GuiThreadInfo
    {
        public uint Size, Flags;
        public nint Active, Focus, Capture, MenuOwner, MoveSize, Caret;
        public int CaretLeft, CaretTop, CaretRight, CaretBottom;
    }
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo info);
    [DllImport("oleacc.dll", PreserveSig = true)] private static extern int AccessibleObjectFromPoint(
        Point point, [MarshalAs(UnmanagedType.Interface)] out Accessibility.IAccessible accessible,
        [MarshalAs(UnmanagedType.Struct)] out object child);

    private static bool KeyDown(Keys key) => (GetAsyncKeyState((int)key) & 0x8000) != 0;

    private static string CurrentModifiers()
    {
        var keys = new List<string>();
        if (KeyDown(Keys.ControlKey) || KeyDown(Keys.LControlKey) || KeyDown(Keys.RControlKey)) keys.Add("Control");
        if (KeyDown(Keys.ShiftKey) || KeyDown(Keys.LShiftKey) || KeyDown(Keys.RShiftKey)) keys.Add("Shift");
        if (KeyDown(Keys.Menu) || KeyDown(Keys.LMenu) || KeyDown(Keys.RMenu)) keys.Add("Alt");
        return keys.Count == 0 ? "None" : string.Join(", ", keys);
    }

    private delegate nint HookProc(int code, nint message, nint data);
    [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
    private delegate bool EnumWindowProc(nint window, nint parameter);
    [StructLayout(LayoutKind.Sequential)] private struct MouseInfo
    { public Point Point; public int MouseData, Flags, Time; public nint ExtraInfo; }
    [StructLayout(LayoutKind.Sequential)] private struct KeyboardInfo
    { public uint VirtualKey, ScanCode, Flags, Time; public nint ExtraInfo; }
    [DllImport("user32.dll", SetLastError = true)] private static extern nint SetWindowsHookEx(int id, HookProc callback, nint module, uint thread);
    [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(nint hook);
    [DllImport("user32.dll")] private static extern nint CallNextHookEx(nint hook, int code, nint message, nint data);
    [DllImport("user32.dll")] private static extern bool GetKeyboardState(byte[] state);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")] private static extern nint GetForegroundWindow();
    [DllImport("user32.dll")] private static extern nint GetWindow(nint window, uint command);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(nint window, out uint processId);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(nint window);
    [DllImport("user32.dll")] private static extern bool IsWindow(nint window);
    [DllImport("user32.dll")] private static extern nint WindowFromPoint(System.Drawing.Point point);
    [DllImport("user32.dll")] private static extern nint GetAncestor(nint window, uint flags);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(nint window);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(nint parent, EnumWindowProc callback, nint parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowProc callback, nint parameter);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(nint window, out NativeRect rectangle);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(nint window, System.Text.StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")]
    private static extern nint SendMessageTimeout(nint window, uint message, nint wParam,
        nint lParam, uint flags, uint timeout, out nint result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")]
    private static extern nint SendMessageTimeout(nint window, uint message, nint wParam,
        System.Text.StringBuilder text, uint flags, uint timeout, out nint result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(nint window, System.Text.StringBuilder text, int capacity);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowLong(nint window, int index);
    [DllImport("user32.dll")] private static extern nint GetKeyboardLayout(uint thread);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int ToUnicodeEx(uint key, uint scan,
        byte[] state, System.Text.StringBuilder buffer, int length, uint flags, nint layout);
}
