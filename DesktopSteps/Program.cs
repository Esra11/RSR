namespace DesktopSteps;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        if (args is ["--render-manual-pdf", var sessionDirectory, var outputPath])
        {
            var planPath = Path.Combine(sessionDirectory, "execution_plan.json");
            var plan = JsonFile.LoadAsync<ExecutionPlan>(planPath).GetAwaiter().GetResult()
                ?? throw new InvalidDataException("The selected recording has no execution plan.");
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
            ManualPdf.CreateAsync(plan, sessionDirectory, outputPath, CancellationToken.None)
                .GetAwaiter().GetResult();
            return;
        }
        if (args is ["--rebuild-recording", var recordingDirectory])
        {
            try
            {
                Task.Run(() => RebuildRecordingAsync(recordingDirectory)).GetAwaiter().GetResult();
                File.Delete(Path.Combine(recordingDirectory, "rebuild_error.txt"));
            }
            catch (Exception ex)
            {
                File.WriteAllText(Path.Combine(recordingDirectory, "rebuild_error.txt"), ex.ToString());
                Environment.ExitCode = 1;
            }
            return;
        }
        if (args is ["--repair-recording", var repairDirectory])
        {
            var planPath = Path.Combine(repairDirectory, "execution_plan.json");
            var eventsPath = Path.Combine(repairDirectory, "recorded_events.json");
            var plan = PlanCompactor.Compact(JsonFile.LoadAsync<ExecutionPlan>(planPath)
                .GetAwaiter().GetResult() ?? throw new InvalidDataException("The recording has no plan."));
            var events = JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath).GetAwaiter().GetResult()
                ?? throw new InvalidDataException("The recording has no captured events.");
            MainForm.RecoverOmittedWorksheetNavigation(plan, events, out _);
            JsonFile.SaveAsync(planPath, plan).GetAwaiter().GetResult();
            return;
        }
        Application.Run(new MainForm());
    }

    private static async Task RebuildRecordingAsync(string directory)
    {
        var eventsPath = Path.Combine(directory, "recorded_events.json");
        var events = await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath)
            ?? throw new InvalidDataException("The selected recording has no captured events.");
        events.RemoveAll(item => item.Kind == "click" && item.AfterState == "dialog-command" &&
            item.ClickX is null && item.ClickY is null);
        events = MainForm.ExpandDialogFieldSnapshots(
            MainForm.RecoverRecordedDialogCommands(events));
        await JsonFile.SaveAsync(eventsPath, events);
        var rulesPath = Path.Combine(directory, "recording_rules.json");
        var rules = File.Exists(rulesPath)
            ? System.Text.Json.JsonSerializer.Deserialize<string[]>(await File.ReadAllTextAsync(rulesPath)) ?? []
            : [];
        var plan = RefreshTiming.AddObservedWait(
            await new Foundry(MainForm.EnvPath()).PlanAsync(events, CancellationToken.None, rules), events);
        await JsonFile.SaveAsync(Path.Combine(directory, "execution_plan.json"), plan);
        await File.WriteAllTextAsync(Path.Combine(directory, "summary.txt"), plan.Summary);
        await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"),
            CancellationToken.None);
        File.Delete(Path.Combine(directory, "intents_need_rebuild.txt"));
    }
}
