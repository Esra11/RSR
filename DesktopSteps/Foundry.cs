using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace DesktopSteps;

internal sealed class Foundry
{
    private readonly string endpoint, agent;
    private readonly string? subscription;
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(150) };
    private readonly HttpClient client;

    public Foundry(string envPath) : this(envPath, Http) { }

    public Foundry(string envPath, HttpClient client)
    {
        this.client = client;
        var settings = File.ReadAllLines(envPath).Where(x => x.Contains('=') && !x.TrimStart().StartsWith('#'))
            .Select(x => x.Split('=', 2)).ToDictionary(x => x[0].Trim(), x => x[1].Trim().Trim('"'));
        endpoint = settings.GetValueOrDefault("PROJECT_ENDPOINT")?.TrimEnd('/')
            ?? throw new InvalidOperationException("PROJECT_ENDPOINT is missing.");
        agent = settings.GetValueOrDefault("AGENT_NAME")
            ?? throw new InvalidOperationException("AGENT_NAME is missing.");
        subscription = settings.GetValueOrDefault("AZURE_SUBSCRIPTION_ID");
    }

    public async Task<string> AskAsync(string prompt, CancellationToken token)
    {
        var accessToken = await GetAccessTokenAsync(token);
        return await AskWithTokenAsync(prompt, accessToken, token);
    }

    private async Task<string> GetAccessTokenAsync(CancellationToken token)
    {
        using var process = Process.Start(CliTokenProcess())
            ?? throw new InvalidOperationException("Azure CLI could not start.");
        var accessToken = (await process.StandardOutput.ReadToEndAsync(token)).Trim();
        var error = await process.StandardError.ReadToEndAsync(token);
        await process.WaitForExitAsync(token);
        if (process.ExitCode != 0 || accessToken.Length == 0)
            throw new InvalidOperationException("Azure sign-in is required: " + error);
        return accessToken;
    }

    private async Task<string> AskWithTokenAsync(string prompt, string accessToken, CancellationToken token)
    {
        var url = $"{endpoint}/agents/{Uri.EscapeDataString(agent)}/endpoint/protocols/openai/responses?api-version=v1";
        using var request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        request.Content = new StringContent(JsonSerializer.Serialize(new { input = prompt }), Encoding.UTF8, "application/json");
        using var response = await client.SendAsync(request, token);
        var body = await response.Content.ReadAsStringAsync(token);
        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException($"Foundry returned {(int)response.StatusCode} ({response.ReasonPhrase}). {body}",
                null, response.StatusCode);
        if (string.IsNullOrWhiteSpace(body)) throw new InvalidDataException("Foundry returned an empty response.");
        using var document = JsonDocument.Parse(body);
        if (document.RootElement.TryGetProperty("output_text", out var outputText)) return outputText.GetString() ?? "";
        var chunks = new List<string>();
        foreach (var output in document.RootElement.GetProperty("output").EnumerateArray())
            if (output.TryGetProperty("content", out var content))
                foreach (var item in content.EnumerateArray())
                    if (item.TryGetProperty("text", out var text)) chunks.Add(text.GetString() ?? "");
        return string.Join("\n", chunks);
    }

    private ProcessStartInfo CliTokenProcess()
    {
        var args = new List<string> { "account", "get-access-token", "--resource", "https://ai.azure.com",
            "--query", "accessToken", "-o", "tsv" };
        if (!string.IsNullOrWhiteSpace(subscription))
            args.AddRange(["--subscription", subscription]);
        return AzureCli.Command(false, args.ToArray());
    }

    public async Task<ExecutionPlan> PlanAsync(IReadOnlyList<RecordedEvent> events, CancellationToken token,
        IReadOnlyList<string>? recordingRules = null)
    {
        events = ActionGrouper.RecoverMisclassifiedLegendDrags(events);
        var grouped = ActionGrouper.Group(events);
        if (grouped.FirstOrDefault(item => item.Kind == "unresolved-input") is { } unresolved)
            throw new InvalidDataException("Recording is incomplete: " + unresolved.Diagnostic +
                " Captured events were preserved; planning cannot invent the missing outcome or identity.");
        if (grouped.Count == 0) throw new InvalidDataException("No replayable actions were captured.");
        var safe = grouped.Select((e, i) => new { number = i + 1, e.Kind, e.Target,
            textLength = e.Kind == "type" ? e.Value?.Length : null,
            intent = SafeIntent(e.Intent),
            intentKey = e.Intent is null ? null : e.Key,
            selectionContext = e.Kind == "context-click" ? e.AfterState : null,
            clickVerification = e.Kind == "unresolved-click" || e.AfterState == "unverified-click" ? "needs review" : null,
            filterState = ExcelFilterPlan.IsItem(e.Target) ? e.AfterState : null,
            shortcut = e.Kind == "scroll" ? e.Key : e.Key is null ? null : e.Key.Contains('+') ? e.Key : "single key withheld" }).ToArray();
        var intents = grouped.Select((e, i) => new { number = i + 1, note = SafeIntent(e.Intent) })
            .Where(item => !string.IsNullOrWhiteSpace(item.note)).ToArray();
        var promptPrefix = "Infer the user's intended outcome from this numbered portion of a desktop action sequence. Return ONLY JSON matching " +
            "{\"summary\":string,\"steps\":[{\"number\":int,\"action\":\"click|context-click|key|type|scroll|resize-column|replace-all|ensure-state|select-all-items|delete-row|open-private-window|fill-down-to-adjacent-data-end\",\"expectedState\":string|null,\"explanation\":string|null,\"relativeWeekday\":string|null,\"intentNumber\":int|null}],\"intentActions\":[{\"intentNumber\":int,\"targetNumber\":int,\"position\":\"before|after\",\"action\":\"replace-all|select-all-items|type|key|click|optional-click|ensure-state|close-window\",\"window\":string|null,\"name\":string|null,\"value\":string|null,\"key\":string|null,\"expectedState\":string|null,\"whenUser\":string|null,\"explanation\":string}]}. " +
            "Intent notes are workflow instructions, not captions of their anchor action. They may describe several earlier or later actions, or request extra actions not performed during recording. " +
            "Infer scope from meaning and surrounding controls. Associate each affected recorded step with its source intentNumber, including steps before the note. " +
            "Keep every recorded step. Put any additional explicitly requested actions in intentActions, never fabricate recorded evidence. " +
            "For extra actions, targetNumber must reference an existing action in the requested core range with a suitable control; position specifies placement relative to it. " +
            "intentNumber MUST be one of the Workflow intents numbers below, not a target number or an index into the intents list. If there are no Workflow intents, intentActions MUST be empty. " +
            "A recorded sheet rename is already handled locally from its tab and typing sequence. For weekday-based renaming, associate the existing rename input with the note and set relativeWeekday; do not add click, key or type actions to repeat that rename. " +
            "Each chunk handles only its own controls. If an intent concerns a different part of the workflow and no suitable control is in this chunk, do not generate that action here; another chunk will handle it. " +
            "Do not duplicate actions already satisfied by the recorded steps. Use whenUser only for explicit user conditions. " +
            "A replace-all intent action may reference a recorded control in Find and Replace (including its Replace tab) and supply the exact requested find and replacement strings; the executor handles that dialog operation. " +
            "For replace-all, targetNumber MUST be an action whose Target.Window is exactly Find and Replace. Never anchor a replacement to a worksheet cell, filter or later unrelated action. " +
            "For replace-all, value is ONLY the literal find text, expectedState is ONLY the literal replacement text (not a sentence or desired-state description). " +
            "Example: replacing \"Old\" with \"New\" means value:\"Old\", expectedState:\"New\". Both strings must appear verbatim in the user's note. " +
            "whenUser must be ONLY the exact login username (for example esshrouf or hulopesv), or null; never write a sentence in whenUser. " +
            "A close-window intent action must be explicitly requested. Click actions may reference only a recorded named control explicitly named in the note. Other operations require a suitable recorded control and explicit instructions. " +
            "EXCEPTION: an optional dialog explicitly named in an intent may use optional-click, with window the exact quoted dialog title and name the exact button named in the note. Use the existing application's recorded command as targetNumber, position before. " +
            "If the note says to click that recorded command only when the dialog is absent, generate ONE optional-click before that command; the planner will turn the recorded command into its absent-dialog fallback. Do not generate another click for the fallback or invent a dialog title/button. " +
            "targetNumber for optional-click MUST be the absent-dialog fallback, NEVER the command that launches the workflow or causes the dialog. Preserve the launching command unconditionally BEFORE the optional wait. Example: launch a query, then Yes if present, otherwise Refresh View: anchor optional-click BEFORE Refresh View, NOT before launch query. " +
            "Do not add fill-down actions to intentActions. Fill-down is supported only as interpretation of a recorded action directly following a completed type event with a known source value and address. Otherwise preserve the recorded sequence and explain the note without guessing a source row or formula. " +
            "For sorting or toggles use ensure-state with a clear expected state; do not assume a click always achieves it. " +
            "Two clicks on a sort header do not imply a fixed direction. Infer ascending/descending only from an explicit intent or recorded sort state. Otherwise use expectedState sort:unspecified so replay asks for the desired outcome. " +
            "Preserve filter checkbox clicks and their recorded filterState; do not infer inclusion/exclusion or checked state from click counts or control names. The local planner will consolidate the final verified filter outcome or ask before replay. " +
            "When a first list item is selected and Shift+End follows, infer selection of every list item, independent of item text. " +
            "For a context click with selection-intent:all or selection-count greater than one, target the current selection, independent of the clicked item's text. " +
            "For any editable field, treat typing, corrections, deletion, paste, and history selection as one final text entry when the recorder has a completed field value. Preserve the final value rather than intermediate fragments. " +
            "An unresolved-click has no verified target. Keep it as a manual-click step for the user; do not invent a target or omit it. A verified delete-row event must remain a delete-row step. A click marked needs review must remain in the plan. " +
            "A resize-column event is a measured drag of a live column boundary. Keep it as resize-column, not a click. " +
            "When an intent-bearing spreadsheet action explicitly says to copy or fill the preceding entered cell downward until the adjacent column becomes empty, use fill-down-to-adjacent-data-end and set expectedState to adjacent-column:<recorded column letters>. Do not preserve the drag endpoint as an ordinary click. " +
            "If a recorded click target is the Replace All button, preserve it as a replacement action using the confirmed Find and Replace field values. " +
            "A click reported only as the dialog window is insufficient evidence for a recorded Replace All; use a separately attributed intent action when explicitly requested. " +
            "An intent note describes the outcome the user wants on future runs. If it asks for next Tuesday or another next weekday in entered text, set relativeWeekday to that weekday on the note-bearing action; never freeze the date from the recording. " +
            "Keep one step per requested action in order, with the same number. Do not renumber steps. Text content is withheld; describe text entry generically. " +
            "Standing recording rules guide interpretation. Rules: " +
            JsonSerializer.Serialize(recordingRules ?? []) + ". Workflow intents (numbers are timing anchors, NOT fixed scope): " +
            JsonSerializer.Serialize(intents) + ". ";
        const int chunkSize = 30;
        const int overlap = 2;
        var accessToken = await GetAccessTokenAsync(token);
        using var slots = new SemaphoreSlim(2);
        var requests = Enumerable.Range(0, (safe.Length + chunkSize - 1) / chunkSize).Select(async chunkIndex =>
        {
            var coreStart = chunkIndex * chunkSize;
            var coreEnd = Math.Min(safe.Length, coreStart + chunkSize);
            var contextStart = Math.Max(0, coreStart - overlap);
            var contextEnd = Math.Min(safe.Length, coreEnd + overlap);
            var payload = safe[contextStart..contextEnd];
            var prompt = promptPrefix + $"Return steps only for action numbers {coreStart + 1} through {coreEnd}. " +
                "The adjacent actions are context only. Metadata only: " + JsonSerializer.Serialize(payload);
            await slots.WaitAsync(token);
            try
            {
                var answer = await RequestValidatedPlanningChunkAsync(
                    correctedPrompt => RequestPlanningChunkAsync(correctedPrompt, accessToken, coreStart, coreEnd,
                        token, TimeSpan.FromSeconds(120)), prompt, grouped, coreStart, coreEnd);
                var parsed = ParseSuggestions(answer, coreStart, coreEnd);
                return (Index: chunkIndex, parsed.Summary, parsed.Steps,
                    IntentActions: ParseIntentActions(answer, coreStart, coreEnd));
            }
            finally { slots.Release(); }
        }).ToArray();
        var results = await Task.WhenAll(requests);
        var suggestions = results.SelectMany(result => result.Steps)
            .ToDictionary(step => step.GetProperty("number").GetInt32(), step => step);
        var steps = grouped.Select((e, i) =>
        {
            suggestions.TryGetValue(i + 1, out var s);
            string? Get(string key) => s.ValueKind == JsonValueKind.Object && s.TryGetProperty(key, out var p) && p.ValueKind == JsonValueKind.String ? p.GetString() : null;
            var relatedIntent = s.ValueKind == JsonValueKind.Object &&
                s.TryGetProperty("intentNumber", out var intentNumber) && intentNumber.ValueKind == JsonValueKind.Number &&
                intentNumber.TryGetInt32(out var noteIndex) &&
                noteIndex > 0 && noteIndex <= grouped.Count ? grouped[noteIndex - 1].Intent : null;
            var stepIntent = e.Intent?.StartsWith("User confirmed filter outcome: ", StringComparison.Ordinal) == true
                ? e.Intent : relatedIntent ?? e.Intent;
            var action = e.Kind is "delete-row" or "open-private-window" or "set-chart-legend-layout" or "set-chart-source-range" ? e.Kind : e.Kind == "unresolved-click" ||
                e.Kind == "click" && e.Target?.ControlType is ("ControlType.Window" or "ControlType.Tab") ||
                e.Kind is "click" or "context-click" && e.Target?.ControlType == "ControlType.ToolTip"
                ? "manual-click" :
                e.Kind is "type" or "context-click" or "scroll" or "resize-column" ? e.Kind : Get("action") ?? e.Kind;
            if (e.Target?.ControlType == "ControlType.HeaderItem" && e.Target.ClassName != "XLSelectAllHeader" &&
                e.Kind == "click")
                action = "ensure-state";
            if (e.Kind == "key" && e.Key is "Enter" or "Return")
                action = "key";
            if (action == "replace-all" && e.Target is not { Window: "Find and Replace", Name: "Replace All" })
                action = e.Kind;
            if (action == "ensure-state" && e.Target?.ControlType != "ControlType.HeaderItem" &&
                e.Target?.ControlType != "ControlType.CheckBox") action = e.Kind;
            if (action is not ("click" or "context-click" or "key" or "type" or "scroll" or "resize-column" or "replace-all" or "ensure-state" or "select-all-items" or "manual-click" or "delete-row" or "open-private-window" or "fill-down-to-adjacent-data-end")) action = e.Kind;
            var target = e.Kind == "type" && e.Target is not null &&
                e.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) == true
                ? e.Target with { Name = e.AfterState[12..] } : e.Target;
            return new PlanStep(i + 1, action, target, e.Value, e.Key,
                e.AfterState == "unverified-click" ? "unverified-click" :
                e.Kind == "open-private-window" ? e.AfterState :
                e.Kind is "scroll" or "resize-column" or "set-chart-legend-layout" or "set-chart-source-range" || e.AfterState?.StartsWith("sort:") == true ||
                e.AfterState?.StartsWith("filter:") == true ||
                e.AfterState?.StartsWith("sheet-name:") == true ||
                e.AfterState?.StartsWith("field-value:", StringComparison.Ordinal) == true ||
                e.AfterState?.StartsWith("excel-selection-range:", StringComparison.Ordinal) == true ||
                e.AfterState?.StartsWith("launched-window:", StringComparison.Ordinal) == true ||
                e.Kind == "context-click" && (e.AfterState?.StartsWith("selection-intent:") == true ||
                    e.AfterState?.StartsWith("selection-count:") == true)
                    ? e.AfterState : e.Kind == "click" && e.Target?.ControlType == "ControlType.HeaderItem" &&
                        e.Target.ClassName != "XLSelectAllHeader"
                        ? ExplicitSortState(relatedIntent ?? e.Intent) : Get("expectedState"), e.Screenshot,
                e.Kind is "set-chart-legend-layout" or "set-chart-source-range"
                    ? e.Diagnostic : Get("explanation"),
                stepIntent, NormalizeRelativeWeekday(Get("relativeWeekday")) ?? RelativeWeekdayFromIntent(stepIntent), null,
                e.AfterState?.StartsWith("visible-row:", StringComparison.Ordinal) == true
                    ? e.AfterState : null);
        }).ToList();
        steps = ExpandIntentActions(steps, grouped, results.SelectMany(result => result.IntentActions).ToArray());
        var summary = string.Join(" ", results.OrderBy(result => result.Index)
            .Select(result => result.Summary).Where(text => !string.IsNullOrWhiteSpace(text)));
        if (summary.Length == 0) summary = "Recorded desktop actions";
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

    private static JsonElement[] ParseIntentActions(string answer, int coreStart, int coreEnd)
    {
        using var json = JsonDocument.Parse(answer[answer.IndexOf('{')..(answer.LastIndexOf('}') + 1)]);
        if (!json.RootElement.TryGetProperty("intentActions", out var actions)) return [];
        return actions.EnumerateArray().Where(action =>
            action.GetProperty("targetNumber").GetInt32() > coreStart &&
            action.GetProperty("targetNumber").GetInt32() <= coreEnd).Select(action => action.Clone()).ToArray();
    }

    private static async Task<string> RequestValidatedPlanningChunkAsync(Func<string, Task<string>> request,
        string prompt, IReadOnlyList<RecordedEvent> events, int coreStart, int coreEnd)
    {
        var validIntents = events.Select((item, index) => (item, number: index + 1))
            .Where(entry => !string.IsNullOrWhiteSpace(entry.item.Intent))
            .Select(entry => entry.number).ToArray();
        for (var attempt = 0; attempt < 2; attempt++)
        {
            var answer = await request(prompt);
            try
            {
                ParseSuggestions(answer, coreStart, coreEnd);
                foreach (var action in ParseIntentActions(answer, coreStart, coreEnd))
                {
                    var hasNote = action.TryGetProperty("intentNumber", out var note);
                    if (!hasNote || note.ValueKind != JsonValueKind.Number || !note.TryGetInt32(out var number) ||
                        !validIntents.Contains(number))
                        throw new InvalidDataException($"An added action references invalid intentNumber {(hasNote ? note.ToString() : "(missing)")}. " +
                            $"Valid recorded intent numbers: [{string.Join(", ", validIntents)}].");
                    var target = action.GetProperty("targetNumber").GetInt32();
                    if (target < 1 || target > events.Count || events[target - 1].Target is null)
                        throw new InvalidDataException($"An added action references targetNumber {target} without a recorded control.");
                }
                return answer;
            }
            catch (InvalidDataException ex)
            {
                if (attempt == 1)
                    throw new InvalidDataException($"Foundry returned invalid planning anchors for actions " +
                        $"{coreStart + 1}-{coreEnd} after one correction attempt. {ex.Message} " +
                        "Captured events were preserved; no substitute action or plan was generated.", ex);
                System.Diagnostics.Trace.WriteLine("Retrying invalid Foundry planning response: " + ex.Message);
                prompt += " Your previous response was rejected: " + ex.Message +
                    " Return the complete JSON again with the original action numbers. " +
                    "Do not invent an intent anchor or repeat an already recorded sheet rename.";
            }
        }
        throw new InvalidOperationException("Planning response validation did not complete.");
    }

    private static List<PlanStep> ExpandIntentActions(List<PlanStep> steps,
        IReadOnlyList<RecordedEvent> events, JsonElement[] actions)
    {
        var before = new Dictionary<int, List<PlanStep>>();
        var after = new Dictionary<int, List<PlanStep>>();
        var fallbacks = new HashSet<int>();
        foreach (var item in actions)
        {
            var noteNumber = item.GetProperty("intentNumber").GetInt32();
            var targetNumber = item.GetProperty("targetNumber").GetInt32();
            if (noteNumber < 1 || noteNumber > events.Count || targetNumber < 1 || targetNumber > steps.Count ||
                string.IsNullOrWhiteSpace(events[noteNumber - 1].Intent))
                throw new InvalidDataException("Foundry generated an action without a valid recorded intent or target anchor.");
            var note = events[noteNumber - 1].Intent!;
            var target = events[targetNumber - 1].Target ??
                throw new InvalidDataException("Intent action has no recorded target context.");
            string? Get(string key) => item.TryGetProperty(key, out var value) &&
                value.ValueKind == JsonValueKind.String ? value.GetString() : null;
            var action = Get("action");
            var value = Get("value");
            var state = Get("expectedState");
            var key = Get("key");
            if (action == "optional-click")
            {
                var window = Get("window");
                var name = Get("name");
                var valid = !(string.IsNullOrWhiteSpace(window) || string.IsNullOrWhiteSpace(name) ||
                    !note.Contains("\"" + window + "\"", StringComparison.Ordinal) ||
                    !Regex.IsMatch(note, @"(?<!\w)" + Regex.Escape(name) + @"(?!\w)", RegexOptions.IgnoreCase) ||
                    !Regex.IsMatch(note, @"\bif\s+(present|.{0,100}\bappears)\b", RegexOptions.IgnoreCase) ||
                    !Regex.IsMatch(note, @"\botherwise\b", RegexOptions.IgnoreCase) ||
                    !Regex.IsMatch(note, @"\botherwise\b.{0,160}\bclick\b", RegexOptions.IgnoreCase | RegexOptions.Singleline) ||
                    Get("position") != "before" || Get("whenUser") is not null ||
                    steps[targetNumber - 1].Action != "click" ||
                    target.ControlType is not ("ControlType.Button" or "ControlType.MenuItem") ||
                    !IntentNamesFallback(note, target.Name) ||
                    string.IsNullOrWhiteSpace(target.Process) ||
                    !fallbacks.Add(targetNumber));
                if (!valid && RecordedOptionalCanBeCompacted(events, note))
                    continue;
                // A vague "doesn't always appear" note establishes optionality
                // but does not authorize inventing a dialog title or control.
                // Keep the recorded action and ignore only the unsupported AI
                // addition; deterministic compaction may still recover it when
                // surrounding recorded evidence is sufficient.
                if (!valid && Regex.IsMatch(note,
                    @"\b(?:doesn['’]?t|does\s+not)\s+always\s+appear\b", RegexOptions.IgnoreCase))
                    continue;
                if (!valid)
                    throw new InvalidDataException("Optional dialog intent needs an exact quoted title, named button, and a single recorded fallback command.");
                target = target with { Window = window, Name = name, ControlType = "ControlType.Button",
                    AutomationId = null, ClassName = null, ParentName = null, ProcessId = null };
                value = key = state = null;
            }
            else if (action == "replace-all" && target.Window == "Find and Replace" &&
                Regex.IsMatch(note, @"\breplace\b", RegexOptions.IgnoreCase) &&
                !string.IsNullOrWhiteSpace(value) && state is not null &&
                note.Contains(value, StringComparison.Ordinal) && note.Contains(state, StringComparison.Ordinal))
                target = target with { Name = "Replace All", ControlType = "ControlType.Button",
                    AutomationId = null, ClassName = null, ParentName = null };
            else if (action == "select-all-items" &&
                target.ControlType is "ControlType.HeaderItem" or "ControlType.ListItem") { }
            else if (action == "type" && value is not null &&
                target.ControlType is "ControlType.Edit" or "ControlType.Pane" or "ControlType.DataItem") { }
            else if (action == "key" && !string.IsNullOrWhiteSpace(key)) { }
            else if (action == "click" && !string.IsNullOrWhiteSpace(target.Name) &&
                note.Contains(target.Name, StringComparison.OrdinalIgnoreCase)) { }
            else if (action == "ensure-state" && !string.IsNullOrWhiteSpace(state) &&
                target.ControlType is "ControlType.CheckBox" or "ControlType.HeaderItem") { }
            else if (action == "ensure-state" && target is
                    { Process: "EXCEL", ControlType: "ControlType.DataItem" } &&
                Regex.IsMatch(note, @"\bcopy\s+(?:the\s+)?(?:functions?|formulas?)\s+of\s+(?:the\s+)?previous\s+row\b",
                    RegexOptions.IgnoreCase) &&
                Regex.IsMatch(note, @"\bcolumns?\s+[A-Z]{1,3}\s+to\s+[A-Z]{1,3}\b",
                    RegexOptions.IgnoreCase))
                continue; // Deterministic worksheet compaction owns this range operation.
            else if (action == "fill-down-to-adjacent-data-end" && target is
                    { Process: "EXCEL", ControlType: "ControlType.DataItem" } &&
                TryNormalizeExplicitFilteredFill(note, target, out target, out state))
            {
                // Excel can expose a stale cell identity after horizontal scrolling. The intent's
                // explicit columns are authoritative; replay later resolves the first visible row.
            }
            else if (action == "close-window" && !string.IsNullOrWhiteSpace(target.Window) &&
                Regex.IsMatch(note, @"\b(close|dismiss)\b", RegexOptions.IgnoreCase)) { }
            else throw new InvalidDataException($"Foundry requested unsupported or insufficiently specified intent action '{action}' " +
                $"at action {targetNumber} ('{target.Name}' in '{target.Window}', {target.ControlType}). " +
                "Clarify the intent; no partial plan was generated.");
            var position = Get("position");
            if (position is not ("before" or "after"))
                throw new InvalidDataException("Intent action placement must be before or after its recorded anchor.");
            var whenUser = Get("whenUser");
            if (whenUser is not null && (!Regex.IsMatch(whenUser, @"^(?:[\w.-]+\\)?[\w.-]+$") ||
                !Regex.IsMatch(note, @"(?<![\w.-])" + Regex.Escape(whenUser.Split('\\').Last()) +
                    @"(?![\w.-])", RegexOptions.IgnoreCase)))
                throw new InvalidDataException("Intent action user condition must be an exact login named in the recorded note.");
            var destination = position == "before" ? before : after;
            if (!destination.TryGetValue(targetNumber, out var additions))
                destination[targetNumber] = additions = [];
            additions.Add(new PlanStep(0, action!, target, value, key, state,
                events[targetNumber - 1].Screenshot, "Intent-generated: " + Get("explanation"),
                note, null, whenUser, OriginIntent: note));
        }
        var expanded = new List<PlanStep>();
        foreach (var step in steps)
        {
            if (before.TryGetValue(step.Number, out var preceding)) expanded.AddRange(preceding);
            expanded.Add(fallbacks.Contains(step.Number)
                ? step with { Action = "click-if-previous-absent" } : step);
            if (after.TryGetValue(step.Number, out var following)) expanded.AddRange(following);
        }
        return expanded.Select((step, index) => step with { Number = index + 1 }).ToList();
    }

    private static bool RecordedOptionalCanBeCompacted(IReadOnlyList<RecordedEvent> events, string note) =>
        Regex.IsMatch(note, @"\bif\s+present\b", RegexOptions.IgnoreCase) &&
        Regex.IsMatch(note, @"\brefresh\s+button\b", RegexOptions.IgnoreCase) &&
        Regex.IsMatch(note, @"\botherwise\s*,?\s*(?:do\s+)?click\s+it\b", RegexOptions.IgnoreCase) &&
        events.Any(recorded => recorded is { Kind: "click", Intent: { } recordedNote,
                Target: { ControlType: "ControlType.Button", Window: { Length: > 0 } window,
                    Name: { Length: > 0 } name } } &&
            recordedNote == note && note.Contains('"' + window + '"', StringComparison.OrdinalIgnoreCase) &&
            Regex.IsMatch(note, @"(?<!\w)" + Regex.Escape(name) + @"(?!\w)", RegexOptions.IgnoreCase));

    private static bool IntentNamesFallback(string note, string? name)
    {
        var words = Regex.Matches(name ?? "", @"[\p{L}\p{N}]+")
            .Select(match => match.Value)
            .Where(word => !new[] { "view", "button", "menu", "command" }
                .Contains(word, StringComparer.OrdinalIgnoreCase)).ToArray();
        return words.Length > 0 && words.All(word =>
            Regex.IsMatch(note, @"(?<!\w)" + Regex.Escape(word) + @"(?!\w)", RegexOptions.IgnoreCase));
    }

    private static string ExplicitSortState(string? intent)
    {
        var directions = Regex.Matches(intent ?? "", @"\b(ascending|descending)\b", RegexOptions.IgnoreCase)
            .Select(match => match.Value.ToLowerInvariant()).Distinct().ToArray();
        return directions.Length == 1 ? "sort:" + directions[0] : "sort:unspecified";
    }

    private async Task<string> RequestPlanningChunkAsync(string prompt, string accessToken,
        int coreStart, int coreEnd, CancellationToken token, TimeSpan timeout)
    {
        using var chunkTimeout = CancellationTokenSource.CreateLinkedTokenSource(token);
        chunkTimeout.CancelAfter(timeout);
        try { return await AskWithTokenAsync(prompt, accessToken, chunkTimeout.Token); }
        catch (OperationCanceledException ex) when (!token.IsCancellationRequested)
        {
            throw new TimeoutException($"Foundry planning timed out after {timeout.TotalSeconds:0.##} seconds " +
                $"for actions {coreStart + 1}-{coreEnd}. Captured events were preserved; " +
                "use Finish latest to retry generation. No substitute plan was generated.", ex);
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException(
                $"Foundry planning failed for actions {coreStart + 1}-{coreEnd}. " +
                "No substitute plan was generated. " +
                (ex.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden or HttpStatusCode.NotFound
                    ? "Verify PROJECT_ENDPOINT, AGENT_NAME and AZURE_SUBSCRIPTION_ID in .env. "
                    : "Captured events were preserved; use Finish latest to retry generation. ") +
                ex.Message, ex);
        }
    }

    private static (string? Summary, JsonElement[] Steps) ParseSuggestions(string answer, int coreStart, int coreEnd)
    {
        var start = answer.IndexOf('{'); var end = answer.LastIndexOf('}');
        if (start < 0 || end <= start) throw new InvalidDataException("Foundry did not return a JSON plan.");
        using var json = JsonDocument.Parse(answer[start..(end + 1)]);
        var root = json.RootElement;
        var summary = root.TryGetProperty("summary", out var summaryNode) ? summaryNode.GetString() : null;
        var steps = root.GetProperty("steps").EnumerateArray()
            .Where(step => step.GetProperty("number").GetInt32() >= coreStart + 1 &&
                step.GetProperty("number").GetInt32() <= coreEnd)
            .Select(step => step.Clone()).ToArray();
        var expected = Enumerable.Range(coreStart + 1, coreEnd - coreStart);
        if (!steps.Select(step => step.GetProperty("number").GetInt32()).Order().SequenceEqual(expected))
            throw new InvalidDataException($"Foundry omitted or duplicated actions in group {coreStart + 1}-{coreEnd}. " +
                "The recording was preserved; no partial plan was generated.");
        return (summary, steps);
    }

    private static string? SafeIntent(string? note) => note is null ? null : Regex.Replace(note,
        @"\b(password|pin|passcode)\s*(?:is|=|:)?\s*\S+", "$1 [redacted]", RegexOptions.IgnoreCase);

    private static bool TryNormalizeExplicitFilteredFill(string note, ControlRef recordedTarget,
        out ControlRef normalizedTarget, out string? expectedState)
    {
        normalizedTarget = recordedTarget;
        expectedState = null;
        var destination = Regex.Match(note, @"\b(?:in\s+)?column\s+([A-Z]{1,3})\b",
            RegexOptions.IgnoreCase);
        var adjacent = Regex.Match(note,
            @"\bempty\s+adjacent\s+cell\s+in\s+([A-Z]{1,3})\s+column\b",
            RegexOptions.IgnoreCase);
        if (!destination.Success || !adjacent.Success ||
            destination.Groups[1].Value.Equals(adjacent.Groups[1].Value,
                StringComparison.OrdinalIgnoreCase) ||
            !Regex.IsMatch(note, @"\bfirst\s+visible\s+filtered\s+row\b", RegexOptions.IgnoreCase) ||
            !Regex.IsMatch(note, @"\b(?:copy|fill)\b", RegexOptions.IgnoreCase) ||
            !Regex.IsMatch(note, @"\buntil\s+before\b", RegexOptions.IgnoreCase))
            return false;

        var row = Regex.Match(recordedTarget.AutomationId ?? recordedTarget.Name ?? "", @"[1-9]\d*$").Value;
        if (row.Length == 0) return false;
        var address = destination.Groups[1].Value.ToUpperInvariant() + row;
        normalizedTarget = recordedTarget with { Name = address, AutomationId = address };
        expectedState = "adjacent-column:" + adjacent.Groups[1].Value.ToUpperInvariant();
        return true;
    }

    private static string? RelativeWeekdayFromIntent(string? note)
    {
        if (note is null) return null;
        var match = Regex.Match(note,
            @"\bnext\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string? NormalizeRelativeWeekday(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var match = Regex.Match(value.Trim(),
            @"^(?:next\s+)?(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$",
            RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : value;
    }
}
