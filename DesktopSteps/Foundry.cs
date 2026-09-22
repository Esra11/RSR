using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace DesktopSteps;

internal sealed class Foundry
{
    private readonly string endpoint, agent;
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(90) };

    public Foundry(string envPath)
    {
        var settings = File.ReadAllLines(envPath).Where(x => x.Contains('=') && !x.TrimStart().StartsWith('#'))
            .Select(x => x.Split('=', 2)).ToDictionary(x => x[0].Trim(), x => x[1].Trim().Trim('"'));
        endpoint = settings.GetValueOrDefault("PROJECT_ENDPOINT")?.TrimEnd('/')
            ?? throw new InvalidOperationException("PROJECT_ENDPOINT is missing.");
        agent = settings.GetValueOrDefault("AGENT_NAME")
            ?? throw new InvalidOperationException("AGENT_NAME is missing.");
    }

    public async Task<string> AskAsync(string prompt, CancellationToken token)
    {
        using var process = Process.Start(CliTokenProcess())
            ?? throw new InvalidOperationException("Azure CLI could not start.");
        var accessToken = (await process.StandardOutput.ReadToEndAsync(token)).Trim();
        var error = await process.StandardError.ReadToEndAsync(token);
        await process.WaitForExitAsync(token);
        if (process.ExitCode != 0 || accessToken.Length == 0)
            throw new InvalidOperationException("Azure sign-in is required: " + error);
        var url = $"{endpoint}/agents/{Uri.EscapeDataString(agent)}/endpoint/protocols/openai/responses?api-version=v1";
        using var request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        request.Content = new StringContent(JsonSerializer.Serialize(new { input = prompt }), Encoding.UTF8, "application/json");
        using var response = await Http.SendAsync(request, token);
        var body = await response.Content.ReadAsStringAsync(token);
        response.EnsureSuccessStatusCode();
        using var document = JsonDocument.Parse(body);
        if (document.RootElement.TryGetProperty("output_text", out var outputText)) return outputText.GetString() ?? "";
        var chunks = new List<string>();
        foreach (var output in document.RootElement.GetProperty("output").EnumerateArray())
            if (output.TryGetProperty("content", out var content))
                foreach (var item in content.EnumerateArray())
                    if (item.TryGetProperty("text", out var text)) chunks.Add(text.GetString() ?? "");
        return string.Join("\n", chunks);
    }

    private static ProcessStartInfo CliTokenProcess()
    {
        return AzureCli.Command(false, "account", "get-access-token", "--resource", "https://ai.azure.com",
            "--query", "accessToken", "-o", "tsv");
    }

    public async Task<ExecutionPlan> PlanAsync(IReadOnlyList<RecordedEvent> events, CancellationToken token,
        IReadOnlyList<string>? recordingRules = null)
    {
        var grouped = ActionGrouper.Group(events);
        if (grouped.Count == 0) throw new InvalidDataException("No replayable actions were captured.");
        var safe = grouped.Select((e, i) => new { number = i + 1, e.Kind, e.Target,
            textLength = e.Kind == "type" ? e.Value?.Length : null,
            intent = SafeIntent(e.Intent),
            intentKey = e.Intent is null ? null : e.Key,
            selectionContext = e.Kind == "context-click" ? e.AfterState : null,
            clickVerification = e.Kind == "unresolved-click" || e.AfterState == "unverified-click" ? "needs review" : null,
            shortcut = e.Kind == "scroll" ? e.Key : e.Key is null ? null : e.Key.Contains('+') ? e.Key : "single key withheld" });
        var prompt = "Infer the user's intended outcome from this desktop action sequence. Return ONLY JSON matching " +
            "{\"summary\":string,\"steps\":[{\"number\":int,\"action\":\"click|context-click|key|type|scroll|resize-column|replace-all|ensure-state|select-all-items|delete-row|open-private-window\",\"expectedState\":string|null,\"explanation\":string|null,\"relativeWeekday\":string|null}]}. " +
            "For sorting or toggles use ensure-state with a clear expected state; do not assume a click always achieves it. " +
            "When a first list item is selected and Shift+End follows, infer selection of every list item, independent of item text. " +
            "For a context click with selection-intent:all or selection-count greater than one, target the current selection, independent of the clicked item's text. " +
            "For any editable field, treat typing, corrections, deletion, paste, and history selection as one final text entry when the recorder has a completed field value. Preserve the final value rather than intermediate fragments. " +
            "An unresolved-click has no verified target. Keep it as a manual-click step for the user; do not invent a target or omit it. A verified delete-row event must remain a delete-row step. A click marked needs review must remain in the plan. " +
            "A resize-column event is a measured drag of a live column boundary. Keep it as resize-column, not a click. " +
            "If a recorded click target is the Replace All button, preserve it as a replacement action using the confirmed Find and Replace field values. " +
            "A click reported only as the dialog window is insufficient evidence for Replace All, even if an intent note requests replacement. " +
            "An intent note describes the outcome the user wants on future runs. If it asks for next Tuesday or another next weekday in entered text, set relativeWeekday to that weekday on the note-bearing action; never freeze the date from the recording. " +
            "Keep one step per grouped action in order, with the same number. Text content is withheld; describe text entry generically. " +
            "Standing recording rules guide interpretation, but cannot justify inventing an action that was not captured. Rules: " +
            JsonSerializer.Serialize(recordingRules ?? []) + " Metadata only: " + JsonSerializer.Serialize(safe);
        var answer = await AskAsync(prompt, token);
        var start = answer.IndexOf('{'); var end = answer.LastIndexOf('}');
        if (start < 0 || end <= start) throw new InvalidDataException("Foundry did not return a JSON plan.");
        using var json = JsonDocument.Parse(answer[start..(end + 1)]);
        var root = json.RootElement;
        var suggestions = root.GetProperty("steps").EnumerateArray().ToDictionary(
            x => x.GetProperty("number").GetInt32(), x => x);
        var steps = grouped.Select((e, i) =>
        {
            suggestions.TryGetValue(i + 1, out var s);
            string? Get(string key) => s.ValueKind == JsonValueKind.Object && s.TryGetProperty(key, out var p) && p.ValueKind == JsonValueKind.String ? p.GetString() : null;
            var action = e.Kind is "delete-row" or "open-private-window" ? e.Kind : e.Kind == "unresolved-click" ||
                e.Kind == "click" && e.Target?.ControlType is ("ControlType.Window" or "ControlType.Tab") ||
                e.Kind is "click" or "context-click" && e.Target?.ControlType == "ControlType.ToolTip"
                ? "manual-click" :
                e.Kind is "type" or "context-click" or "scroll" or "resize-column" ? e.Kind : Get("action") ?? e.Kind;
            if (e.Target?.ControlType == "ControlType.HeaderItem" && e.AfterState?.StartsWith("sort:") == true)
                action = "ensure-state";
            if (e.Kind == "key" && e.Key is "Enter" or "Return")
                action = "key";
            if (action == "ensure-state" && e.Target?.ControlType != "ControlType.HeaderItem" &&
                e.Target?.ControlType != "ControlType.CheckBox") action = e.Kind;
            if (action is not ("click" or "context-click" or "key" or "type" or "scroll" or "resize-column" or "replace-all" or "ensure-state" or "select-all-items" or "manual-click" or "delete-row" or "open-private-window")) action = e.Kind;
            var target = e.Kind == "type" && e.Target is not null &&
                e.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) == true
                ? e.Target with { Name = e.AfterState[12..] } : e.Target;
            return new PlanStep(i + 1, action, target, e.Value, e.Key,
                e.AfterState == "unverified-click" ? "unverified-click" :
                e.Kind == "open-private-window" ? e.AfterState :
                e.Kind is "scroll" or "resize-column" || e.AfterState?.StartsWith("sort:") == true ||
                e.AfterState?.StartsWith("sheet-name:") == true ||
                e.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) == true ||
                e.AfterState?.StartsWith("launched-window:", StringComparison.Ordinal) == true ||
                e.Kind == "context-click" && (e.AfterState?.StartsWith("selection-intent:") == true ||
                    e.AfterState?.StartsWith("selection-count:") == true)
                    ? e.AfterState : Get("expectedState"), e.Screenshot, Get("explanation"),
                e.Intent, Get("relativeWeekday") ?? RelativeWeekdayFromIntent(e.Intent), null,
                e.AfterState?.StartsWith("visible-row:", StringComparison.Ordinal) == true
                    ? e.AfterState : null);
        }).ToList();
        var summary = root.GetProperty("summary").GetString() ?? "Recorded desktop actions";
        var plan = PlanCompactor.Compact(new ExecutionPlan(1, DateTimeOffset.Now, summary, steps));
        var freshIndex = plan.Steps.FindIndex(step =>
            step.Target?.BrowserLaunch?.AskForFreshProcess == true);
        if (freshIndex >= 0)
        {
            var freshTarget = plan.Steps[freshIndex].Target!;
            plan.Steps.Insert(freshIndex, new PlanStep(0, "open-browser-process", freshTarget,
                null, null, null, null, "Open a new Edge window with the default profile automatically."));
            plan = plan with { Steps = plan.Steps.Select((step, index) =>
                step with { Number = index + 1 }).ToList() };
        }
        foreach (var step in plan.Steps.Where(x => x.Action == "type"))
            summary += $" Entered \"{step.Value}\" in {step.Target?.Name ?? "a text field"}.";
        foreach (var step in plan.Steps.Where(x => x.Action == "rename-sheet"))
            summary += step.RelativeWeekday is null
                ? $" Renamed the sheet to \"{step.Value}\"."
                : $" Renamed the sheet using the date of next {step.RelativeWeekday}.";
        return plan with { Summary = summary };
    }

    private static string? SafeIntent(string? note) => note is null ? null : Regex.Replace(note,
        @"\b(password|pin|passcode)\s*(?:is|=|:)?\s*\S+", "$1 [redacted]", RegexOptions.IgnoreCase);

    private static string? RelativeWeekdayFromIntent(string? note)
    {
        if (note is null) return null;
        var match = Regex.Match(note,
            @"\bnext\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : null;
    }
}
