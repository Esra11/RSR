using System.Diagnostics;
using System.Reflection;
using System.Windows.Automation;
using DesktopSteps;

internal static class ExcelCaptureCheck
{
    public static async Task RunAsync(string outputDirectory)
    {
        var process = Process.GetProcessesByName("EXCEL").Single();
        var window = AutomationElement.FromHandle(process.MainWindowHandle);
        var tabs = window.FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.AutomationIdProperty, "SheetTab"));
        for (var attempt = 0; tabs.Count == 0 && attempt < 20; attempt++)
        {
            await Task.Delay(100);
            tabs = window.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "SheetTab"));
        }
        var selected = tabs.Cast<AutomationElement>().FirstOrDefault(tab =>
            tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var value) &&
            ((SelectionItemPattern)value).Current.IsSelected);
        var tab = tabs.Cast<AutomationElement>().FirstOrDefault(tab => tab.Current.Name == "Clean Sheet")
            ?? selected ?? throw new InvalidOperationException("Excel fixture exposes no selected worksheet tab.");
        var assembly = typeof(ControlRef).Assembly;
        var executor = assembly.GetType("DesktopSteps.Executor")!;
        var recorderType = assembly.GetType("DesktopSteps.Recorder")!;
        using var recorder = (IDisposable)Activator.CreateInstance(recorderType)!;
        try
        {
            executor.GetMethod("ActivateWindow", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [tab]);
            ((SelectionItemPattern)tab.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            await Task.Delay(200);
            recorderType.GetMethod("Start")!.Invoke(recorder, [outputDirectory]);
            await ClickAsync(executor, tab, true);
            await Task.Delay(500);
            var command = Find(process.Id, ControlType.MenuItem, "Move or Copy...");
            var snapshot = recorderType.GetMethod("CaptureMouseDownTarget", BindingFlags.NonPublic | BindingFlags.Static)!;
            var bounds = command.Current.BoundingRectangle;
            var captured = snapshot.Invoke(null,
                [(int)(bounds.Left + bounds.Width / 2), (int)(bounds.Top + bounds.Height / 2)]);
            var capturedTarget = (ControlRef?)captured?.GetType().GetProperty("Target")!.GetValue(captured);
            if (capturedTarget?.Name != "Move or Copy..." || capturedTarget.Window != window.Current.Name)
                throw new InvalidOperationException("Menu snapshot did not retain the originating workbook.");
            await ClickAsync(executor, command, false);
            await Task.Delay(500);
            var cancel = Find(process.Id, ControlType.Button, "Cancel");
            await ClickAsync(executor, cancel, false);
            await Task.Delay(500);
            recorderType.GetMethod("Stop")!.Invoke(recorder, null);
            var events = (IReadOnlyList<RecordedEvent>)recorderType.GetProperty("Events")!.GetValue(recorder)!;
            foreach (var item in events)
                Console.WriteLine($"Captured {item.Kind}: {item.Target?.Name} in {item.Target?.Window}; {item.Diagnostic}");
            if (!events.Any(item => item.Target?.Name == "Move or Copy..." &&
                    item.Target.Window == window.Current.Name) ||
                !events.Any(item => item.Target?.Name == "Cancel" && item.Target.Window == "Move or Copy" &&
                    item.Diagnostic == "Control identity verified at mouse-down."))
                throw new InvalidOperationException("Real recorder did not preserve menu and dialog command identities.");
            Console.WriteLine("PASS: real menu and dialog identities preserved; no sheet copied.");
        }
        finally
        {
            SendKeys.SendWait("{ESC}");
            if (selected is not null)
                ((SelectionItemPattern)selected.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
        }
    }

    private static AutomationElement Find(int processId, ControlType type, string name)
    {
        var roots = AutomationElement.RootElement.FindAll(TreeScope.Children,
            new PropertyCondition(AutomationElement.ProcessIdProperty, processId));
        foreach (AutomationElement root in roots)
        {
            var items = root.FindAll(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, type),
                new PropertyCondition(AutomationElement.NameProperty, name)));
            foreach (AutomationElement item in items)
                if (item.Current.IsEnabled && !item.Current.IsOffscreen) return item;
        }
        throw new InvalidOperationException($"Live fixture control '{name}' not found.");
    }

    private static async Task ClickAsync(Type executor, AutomationElement element, bool rightClick)
    {
        var parameters = new object?[] { element, null, rightClick, null };
        var method = executor.GetMethod("TryClickLiveUiaElement", BindingFlags.NonPublic | BindingFlags.Static)!;
        var clicked = await Task.Run(() => (bool)method.Invoke(null, parameters)!);
        if (!clicked) throw new InvalidOperationException($"Live fixture click failed: {parameters[3]}");
    }
}
