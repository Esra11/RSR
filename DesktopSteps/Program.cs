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
        Application.Run(new MainForm());
    }
}
