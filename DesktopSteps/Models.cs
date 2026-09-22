using System.Text.Json;

namespace DesktopSteps;

public sealed record ControlRef(string? Process, string? Window, string? AutomationId,
    string? Name, string? ControlType, string? ClassName, string? ParentName,
    string? Orientation = null, string? Label = null, int? ProcessId = null,
    BrowserLaunchIdentity? BrowserLaunch = null);

public sealed record BrowserLaunchIdentity(string ExecutablePath, string? UserDataDirectory,
    string? ProfileDirectory, string? CommandLineSha256 = null,
    bool AskForFreshProcess = false, string? CommandLine = null,
    string? ProfilePath = null);

public sealed record RecordedEvent(DateTimeOffset At, string Kind, ControlRef? Target,
    string? Value, string? Key, string? Screenshot, string? BeforeState, string? AfterState,
    string? Intent = null, int? ClickX = null, int? ClickY = null,
    string? Diagnostic = null);

public sealed record DialogFieldSnapshot(ControlRef Target, string Value);

public sealed record PlanStep(int Number, string Action, ControlRef? Target, string? Value,
    string? Key, string? ExpectedState, string? Screenshot, string? Explanation,
    string? Intent = null, string? RelativeWeekday = null, string? WhenUser = null,
    string? TargetStrategy = null);

public sealed record ExecutionPlan(int Version, DateTimeOffset Created, string Summary,
    List<PlanStep> Steps);

internal static class JsonFile
{
    public static readonly JsonSerializerOptions Options = new() { WriteIndented = true,
        PropertyNameCaseInsensitive = true };
    public static async Task SaveAsync<T>(string path, T value, CancellationToken token = default)
    {
        await using var stream = File.Create(path);
        await JsonSerializer.SerializeAsync(stream, value, Options, token);
    }
    public static async Task<T?> LoadAsync<T>(string path, CancellationToken token = default)
    {
        await using var stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<T>(stream, Options, token);
    }
}
