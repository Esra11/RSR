# RSR (Record, Summarize, Replay)

RSR records a Windows workflow, generates a step-by-step PDF with screenshots, and creates a plan to replay the workflow later. You can add an intent note when an action needs more explanation or should adapt on future runs.

## What is included

- `DesktopSteps/`: C# source for the Windows Forms application.
- `bin/RSR/`: the latest framework-dependent Windows build.
- `.env.example`: a safe configuration template. Keep the local `.env` private and out of source control.

This package contains no recordings, screenshots, personal Foundry endpoint, or backup builds.

## Requirements

- Windows with an unlocked interactive desktop session.
- .NET 10 Windows Desktop Runtime to run the included executable. The .NET 10 SDK is needed only to build from source.
- Azure CLI and an account permitted to invoke your configured Microsoft Foundry agent.
- A Microsoft Foundry project endpoint and agent name for generating execution plans.

## Create the required Microsoft Foundry resources

RSR sends recorded action metadata and intent notes to a prompt agent. It needs a Microsoft Foundry project, a deployed chat model, and an agent that uses that model.

1. **Install and sign in to the prerequisites.**
   - Install the [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli-windows).
   - Sign in with `az login` using the account that will create or use the Foundry project.
   - If the account can access several subscriptions, list them with `az account list -o table`, then select the intended subscription with `az account set --subscription "<subscription-id-or-name>"`.
2. **Create or select a Foundry project.**
   - Open the [Microsoft Foundry portal](https://ai.azure.com/).
   - Confirm that you are signed into the intended Microsoft account and directory. Use the same account for Azure CLI sign-in. If a project is missing from the selector, check the account and directory before creating another project.
   - Click the **New Foundry** toggle at the top of the portal. If the **Select a project to continue** dialog appears, open **Select or search for a project**.
   - To reuse a project, select the intended project from the list. The dialog lists projects in supported regions; confirm its resource and location before selecting it.
   - To create a separate RSR project, select **Create a new project** at the top of that dropdown. Complete the creation form: choose your subscription, resource group, supported region, and Foundry resource as requested, and enter a project name such as `rsr-project`. Create a Foundry resource if you do not already have a suitable one. Finish creation and open the new project in the New Foundry experience.
   - RSR uses the new named-agent Responses API; do not substitute a classic hub-based assistant. If an existing project remains unavailable after checking the account and directory, verify its region and your access with the project administrator.
   - Your resource, project, and agent names can differ from the screenshot. For example, use `rsr-resource`, `rsr-project`, and `rsr-planner`.
3. **Deploy a chat model.**
   - After project creation, the New Foundry **Home** page shows **Build an agent**, **Use a model**, and **Model selection** cards.
   - Under **Model selection**, choose a compatible model from the displayed model cards, or browse under **Explore models**. For example, choose **GPT-6.1 Sol** (`gpt-6.1-sol`) if available in your subscription and region. The model is an example, not a fixed RSR requirement.
   - On the model details page within your project, open **Deploy** and choose **Default settings** for an initial pilot: the menu describes this as **Global Standard** with default quota. Choose **Custom settings** when you need an organization-approved deployment type, quota, or guardrails instead.
   - If you opened the standalone model catalog and see **Use this model**, select it and choose your existing project. Confirm the account and directory if the project is absent. Continue in that project's model deployment flow.
   - Wait for deployment to finish. The model **Playground** opens with an **Instructions** field and a **Save as agent** button.
   - Keep the deployment in the same Foundry project. RSR does not need the deployment name in `.env`; the agent stores that association.
4. **Create the prompt agent.**
   - In the deployed model's **Playground**, replace the default **Instructions** with:

     ```text
     Follow the task-specific instructions and output format supplied in each
     request. Preserve recorded action numbers and evidence. Do not invent
     actions or observations. Return the requested format without commentary.
     ```

   - RSR supplies the detailed planning schema and workflow rules in each request from `DesktopSteps/Foundry.cs`.
   - Under **Tools**, remove **Web search** using its three-dot menu if it is present. No tools, knowledge attachments, voice mode, or hosted-agent container are required.
   - Click **Save as agent**. In **Create an agent**, replace the generated **Agent name** with `rsr-planner`. If the name is already taken, choose a unique alternative and use that exact alternative in `.env`.
   - Click **Create and open playground**. The agent Playground should show `rsr-planner` as its heading and a saved version.
   - No Teams or Microsoft 365 publishing step is required. RSR uses the agent's **Responses** endpoint with **Microsoft Entra** authorization.
5. **Copy the project endpoint.**
   - Return to the project's **Home** page. Below the two main cards, click the copy icon beside **Project endpoint** (the middle field). Copy the complete value rather than the visibly truncated text.
   - The expected form is `https://<resource>.services.ai.azure.com/api/projects/<project>`.
   - Use the project endpoint, rather than the **Azure OpenAI endpoint** shown to its right, a model deployment endpoint, or the base resource URL.
   - **API key authentication is disabled for this project** is compatible with RSR: the application authenticates using your Azure CLI Microsoft Entra sign-in, not an API key.
6. **Grant access to the account that runs RSR.**
   - Assign **Foundry User** on the Foundry resource or project to users who create, configure, and invoke agents.
   - For an invoke-only account, **Foundry Agent Consumer** may be sufficient when that role is available for the project. The account still needs permission to invoke the selected agent.
   - Role changes can take several minutes to become effective. See [Foundry role-based access control](https://learn.microsoft.com/azure/foundry/concepts/rbac-foundry) for current role details.
7. **Record the subscription ID.**
   - Copy it from the Azure portal, or run `az account show --query id -o tsv` after selecting the correct subscription.

Microsoft's current end-to-end setup is also documented in [Get started with Microsoft Foundry](https://learn.microsoft.com/azure/foundry/quickstarts/get-started-code), the [prompt agent quickstart](https://learn.microsoft.com/azure/foundry/agents/quickstarts/prompt-agent), and [agent endpoint configuration](https://learn.microsoft.com/azure/foundry/agents/how-to/configure-agent).

RSR constructs this URL automatically from the two configuration values:

```text
{PROJECT_ENDPOINT}/agents/{AGENT_NAME}/endpoint/protocols/openai/responses?api-version=v1
```

Do not paste that complete URL into `PROJECT_ENDPOINT`; use only the project endpoint ending in `/api/projects/<project>`.

## Configure `.env`

The preferred location is the repository root, beside this README:

```text
C:\work\RSR-Github\RSR-main\.env
```

Copy `.env.example` to `.env`, then replace every placeholder:

```dotenv
PROJECT_ENDPOINT=https://YOUR-RESOURCE.services.ai.azure.com/api/projects/YOUR-PROJECT
AGENT_NAME=rsr-planner
AZURE_SUBSCRIPTION_ID=00000000-0000-0000-0000-000000000000
```

- `PROJECT_ENDPOINT` is required. Paste the Foundry **project endpoint**. A trailing slash is optional.
- `AGENT_NAME` is required. Use the exact saved or published agent name, including its capitalization.
- The setup example consistently uses `rsr-planner`. If you created the agent with a different name, put that actual name here; changing `.env` does not rename the Foundry agent.
- `AZURE_SUBSCRIPTION_ID` is optional but recommended. Set it to the subscription containing the Foundry project, especially when the Azure CLI default points to another subscription or tenant.
- Do not commit `.env`; it contains environment-specific configuration. Commit `.env.example` with placeholders only.

RSR first looks for `.env` beside the running executable and then checks each parent folder. For the included build, the repository-root `.env` is therefore found automatically. For a standalone copied build, place `.env` beside `RSR.exe`, for example `C:\path\to\RSR\.env`.

## Authenticate and run

1. Sign in and select the correct Azure subscription:

   ```powershell
   az login
   az account set --subscription "<subscription-id>"
   az account show --query "{subscription:name, subscriptionId:id, tenantId:tenantId}" -o table
   ```

2. Start the included build:

   ```powershell
   .\bin\RSR\RSR.exe
   ```

3. If plan generation returns `401` or `403`, sign in again and verify the role assignment. If it returns `404 WorkspaceNotFound`, verify the project endpoint, subscription, and tenant. If the agent is not found, verify `AGENT_NAME` against the agent shown in the same project.

To build from source instead, run:

```powershell
dotnet run --project DesktopSteps/DesktopSteps.csproj
```

The application searches its own folder and parent folders for `.env`. New recordings are saved in `Recordings` beside the executable being run. For the included build, that is `bin/RSR/Recordings`. This folder is created when needed and is ignored by Git.

Set `AZURE_SUBSCRIPTION_ID` to the subscription containing the Foundry project, especially when your CLI default is in another tenant. RSR requests that subscription's token without switching your CLI default. If omitted, Azure CLI's default subscription is used. A wrong tenant can cause `404 WorkspaceNotFound` even when the endpoint is correct.

Foundry request failures and incomplete responses stop plan generation with the actual error; captured events remain available for **Finish latest** after correcting configuration. RSR no longer silently substitutes a plan when Foundry fails. Keyboard replay supports Insert, Ctrl+Insert and Shift+Insert using the recorded modifiers.

Planning uses batches of 30 grouped actions, with two adjacent actions for context and at most two concurrent requests. Each batch has a 120-second timeout; the HTTP limit is 150 seconds so it does not cut that wait short. Timeout messages identify the affected action range and recommend **Finish latest**, rather than incorrectly suggesting a configuration failure. User cancellation remains cancellation. Completed responses must cover every requested action before a plan is generated.

Added actions must reference the actual numbered recorded intent, not an index into the list of notes. Invalid planning anchors trigger one correction request with the valid intent numbers; repeated invalid responses stop generation explicitly without guessing a target or changing recorded evidence. Recorded sheet-copy and rename input is handled by the existing local semantic conversion, not by requesting a second rename action from Foundry. Use **Finish latest** to regenerate a preserved recording after installing the correction.

## Use

1. Start recording and perform the task on your desktop. Optionally press **Ctrl+Alt+Shift+I** during recording to add a workflow intent note. Its timestamp is an anchor, not a restriction to the latest action.
2. Stop recording. RSR saves the captured events, execution plan, summary, screenshots, and `manual_steps.pdf` in a timestamped recording folder.
3. Review the plan and PDF, then choose **Replay selected** to repeat the workflow. Leave the desktop available while replay runs.

RSR captures actions and UI control information through Windows accessibility APIs. It saves screenshots for the PDF and human review; it does not run OCR or send screenshot pixels to Foundry. The planning request uses recorded action metadata and optional intent notes. Replay uses Windows UI Automation and stops or asks for help when it cannot safely identify an action. Elevated applications may require running RSR as administrator to record their controls.

Foundry receives all workflow intent notes and surrounding action batches. A note can describe several earlier or later steps. It can also explicitly request additional supported operations, including user-conditional replacements. These additions reference recorded control context, are tagged with `OriginIntent` in the JSON, and are labeled intent-generated in the PDF; they are not claimed to be captured clicks. Replay verifies the source note remains in the recorded events. Ambiguous/unsupported additions fail generation rather than silently inventing a locator. Rebuild existing recordings to apply this interpretation; their captured events do not need to be changed.

The project has been tried with parts of Excel, DebugDiag 2 Collection, and Microsoft Edge. Results depend on each application's accessibility support and current UI state; review a workflow before relying on unattended replay.

Menu replay waits for asynchronously opened dropdowns without reactivating their owner. If UI Automation does not expose a recorded menu item, replay can invoke an exact, unique, visible MSAA menu command in the foreground application's recorded window or associated popup. It does not use saved mouse coordinates or application-specific menu names. Keyboard context-menu recovery is reserved for recorded context-clicks, not ordinary dropdown-button clicks.

Optional dialog actions wait up to 20 seconds for their recorded window and automatically invoke the recorded control if it appears. When absent, replay proceeds only if the next recorded control is visible and enabled in an enabled window, or the plan contains an explicit conditional fallback. Otherwise it stops at the optional step rather than assuming the next view is ready. Leave the desktop untouched during this wait; manual dialog clicks can race replay.

Saved ordinary button clicks also honor an explicit conditional intent naming the exact quoted dialog title and button, saying "if present" and "otherwise". This supports conversational wording without making unrelated clicks optional. A missing unrecorded fallback command is never invented.

Keyboard modifiers are frozen in the keyboard hook before queued processing, so releasing Shift or Ctrl cannot change a recorded shortcut. When a list context action explicitly records `selection-intent:all`, replay establishes any missing all-row selection through UI Automation and verifies it before opening the context menu; it does not infer this intent from row text or an unmodified navigation key.

A plain Table/DataGrid background click without intent or expected state can be focus-only navigation before a button/menu command in the same application. Replay skips that redundant click because the command activates its own target; row/cell selections, context clicks and keyboard-dependent actions are retained.

A plain title-bar click immediately before a button/menu command in the same application is also redundant focus navigation, even when the next command belongs to an owned modal dialog. Compaction omits it and replay supports older plans by skipping it before activation. Title clicks carrying intent, expected state or keyboard dependencies are retained; the rule never crosses application processes.

Optional-dialog intent can explicitly name an unrecorded dialog using its exact quoted title and button name, with a recorded command as the absent-dialog fallback. Planning emits an attributed `optional-click` followed by `click-if-previous-absent`: replay waits up to 20 seconds for the dialog and runs the recorded command only if it does not appear. A note about another dialog does not make the fallback command itself optional.

The optional branch must be anchored to the absent-dialog command named by the intent, not the earlier command that launches the workflow. Planning rejects mismatched anchors; shared intent text alone never makes a normal command optional. A plain native static-label click without intent/state before an independently targeted command in the same window is treated as non-actionable navigation, not an invocation.

Keyboard and wheel identities are frozen before queued capture, and input timestamps use the Windows hook event time. Standalone modifier presses need no control lookup. Delayed input whose target cannot be trusted is saved as `unresolved-input` and blocks generation/replay explicitly. Rebuilding an older recording cannot recover identities already captured from the wrong application's focus; re-record that section.

Input hooks run on a separate STA message loop, not the recorder UI thread. Wheel hooks capture only the native application window, timestamp, delta and axis; accessibility scroll-bar resolution and screenshots happen later. Scroll bursts therefore do not trigger repeated synchronous UI Automation queries or lose their identity while UI processing is busy.

Replacement compaction preserves intent-generated find/replacement values instead of substituting earlier recorded field values. A generated replacement in the same consecutive replacement group already performed by the recorded action under the same user condition is removed. Consecutive replacements reacquire the visible, enabled native dialog after result acknowledgement.

Consecutive sort-header clicks compact into one `ensure-state` outcome, not a fixed number of toggles. If no final direction or explicit direction intent is available, replay asks for ascending/descending before changing applications and saves that choice in the JSON/PDF. Replacement result handling inspects owned native dialogs in the active Excel process using bounded Raw View traversal, reports verified no-match results explicitly, acknowledges them, then executes the saved Close step after replacements.

A Shift+End immediately after verified sorting, followed by explicit all-row context-selection evidence in the same view, uses verified `select-all-items` in that header's live list. An intervening focus-only click on that recorded row is redundant. This works when sorting was already correct and left no row focused, without using old row text or changing ordinary partial-range selections.

An ordinary Excel cell click directly after a known typed value is normalized to fill-down when its explicit intent names the matching destination column and says to copy below until before an empty adjacent cell in a named column. Filtered fills use the first visible data row, not the recorded endpoint. Focus-only clicks on the ExcelGrid pane before independently targeted same-workbook commands are skipped; clicks carrying intent or state are retained.

Filter actions bring their recorded header address into view through Excel Go To before resolving the dropdown, regardless of prior horizontal scrolling. Filtered fills navigate directly to the destination header, inspect adjacent visible data to establish the stopping point, then enter the value once across the visible-only selection using Ctrl+Enter. Each resulting visible value is still verified; hidden rows remain unchanged.

Intent-free scroll and Page left/right navigation immediately before a same-workbook filter action is redundant and removed, since that filter navigates to its own header.

Grid-background focus clicks before a verified filter are also redundant. An explicit conditional note naming the dialog and saying to click the refresh button otherwise produces an attributed `click-if-previous-absent` in the static plan. Its `unique-refresh-command` strategy accepts only one visible, enabled button named Refresh, Refresh View or Refresh Now in the exact resulting window and application. Missing or ambiguous commands stop replay; the fallback never runs when the dialog was handled.

Filter opening uses Alt+Down from the verified header cell, avoiding stale dropdown bounds after horizontal scrolling. Replay still requires the expected live checklist and verifies its selected values and closure. Refresh completion remains limited to 20 seconds: a sustained accessibility busy-state change followed by three restored ready samples can confirm completion even if hover/focus pixels differ. Merely finding an enabled button, without observing a transition, is not completion evidence.

Before replaying a plan with a Refresh button (including a conditional fallback), RSR asks how long refresh is expected to take. Enter positive whole seconds only (1 to 2147483645); signs, decimals, zero and invalid pasted text cannot start replay. The last expectation is saved as `RefreshExpectedSeconds` and offered again on the next replay. That duration plus 2 seconds is the completion timeout, not a mandatory delay: verified restoration of the Refresh button continues early. If completion is not verified by the deadline, the usual help dialog requires confirmation or stops replay. Cancel closes the initial prompt without changing the target application, and Stop replay interrupts the wait. Plans without a configured expectation retain the 20-second completion limit.

Dynamic filtered entry uses Go To from the verified workbook directly; it does not click a stale worksheet cell to establish focus. Intent-free recorded scrolls immediately before that entry are removed because the destination header and first visible row are navigated and verified independently.

A plain StatusBar background click without input, intent or expected state immediately before a named tab click in the same application window is redundant navigation and omitted. Status-bar buttons and meaningful status actions are retained.

A recorded Excel column-header click immediately followed by Delete in the same workbook selects the entire column through Go To and clears its contents without reselecting the active cell. The native worksheet object verifies the whole-column selection before Delete and verifies the column is empty afterward. Missing dated tab references can follow a sheet dynamically renamed during this replay only when the same-workbook name matches that explicit rename template; existing recorded tabs are not redirected.

Visible fills bulk-read adjacent values from the native worksheet object obtained from the verified workbook window, inspect row visibility without moving the cursor, and stop before the first empty visible adjacent cell. Filled values are bulk-read for verification. This removes K-by-K keyboard navigation while retaining hidden-row protection and the stopping rule; unavailable native object access stops explicitly.

The PDF uses the execution plan's exact step order and numbers, without merging or renumbering menu and paste steps, and is refreshed before replay. Explicit column-copy intent followed by a recorded cell selection and Ctrl+C becomes a verified range copy: use the named column and recorded starting row, stop before its first empty cell, and confirm Excel's copy mode. Plain addressed cell clicks use verified Go To navigation instead of stale cell bounds; preceding intent-free viewport scrolls are unnecessary. Generic Excel desktop-pane Ctrl+C preserves the verified worksheet selection rather than searching for an unnamed pane or scrolling it into view. Paste Values refuses to continue without an active Excel copied range.

Row-header context selection brings column A of the requested row into view through verified Go To before hit-testing its live header, even when the row itself is already visible. Move or Copy steps resolve only inside the active native copy dialog and verify the checkbox, destination selection and newly created sheet without repeated workbook-wide scans.

A recorded Paste dropdown followed by Values executes Excel's enabled `PasteValues` command directly, without relying on a transient ribbon popup. Replay requires an active copied rectangle and the verified destination, compares every pasted value with the source snapshot and confirms no formulas were pasted before advancing. Both recorded step numbers remain in the PDF and completion log.

Destination lookup can cross a bounded same-workbook ribbon-tab sequence and a paired File/Back visit, but never another worksheet selection or arbitrary command. The native adapter still verifies the active cell selection and copy mode immediately before pasting. Explicit two-column copy notes can accompany nearby clicks before a verified rectangular Ctrl+C selection; their columns must match that selection, while its endpoint is recalculated from live populated data.

A note on an Excel result-dialog OK can describe that following copy only when the recorded process ID and nearby workbook command corroborate its source workbook. Different users, processes, worksheets and conflicting column names prevent that association.

Table intent operations use application-neutral action names and a capability-adapter registry (`ITableReplayAdapter`). Copy-until-empty and extend-from-last-formula are dispatched by the live window's supported capabilities, not by workflow names, fixed columns or recorded row counts. The current native adapter supports Excel; other applications retain generic UI Automation/MSAA replay but require a registered table adapter for these advanced operations. Unsupported or ambiguous capabilities stop explicitly without input, rather than assuming Excel shortcuts work everywhere. This is not automatic formula support for every application.

Copy intent can name quoted stop values in addition to blanks. The plan preserves these as `copy-stop-values` and the PDF includes them. Excel evaluates the boundary against live values, including displayed native errors such as `#VALUE!`, and excludes the boundary cell. Stable-ID button lookup uses a direct, unique UI Automation query so ribbon commands are not lost beyond a large worksheet's bounded traversal.

Window activation resolves the target's nearest native ancestor and its native parent root, rather than climbing UIA to an owned dialog's disabled workbook owner. This preserves modal dialog activation even when the accessibility tree nests the dialog beneath its owner.

Explicit intent to copy two populated columns becomes a verified rectangular copy starting at the recorded cell row. Replay stops before a row with an empty cell in any selected column, snapshots both columns and verifies the subsequent values-only paste. A recorded ribbon-tab click between the destination cell and Paste Values does not discard the destination; the native selection is still verified before pasting.

Explicitly clarified chart-source replacement uses the last verified two-column paste on the same active sheet. The chart identity comes from the recording, not a fixed workflow name. The Excel adapter sets category/value references including the first data row and verifies every resulting category/value pair. This capability replaces the recorded Select Data menu/dialog sequence only when replacement intent, chart identity and dialog confirmation are present.

Chart legend drags capture verified before/after legend dimensions relative to the owning chart. Replay restores and verifies that layout through the Excel capability adapter, without recorded pointer coordinates. A drag without a verified supported outcome is saved as unresolved input and blocks planning rather than becoming a successful click-only replay. Older click-only recordings cannot recover a missing final resize size; re-record that adjustment.
A plain chart-selection click immediately before a verified legend-layout action for the same chart, workbook, process and user scope is redundant. Planning omits it, and replay recognizes it in existing plans without rewriting their JSON: the layout adapter targets and verifies the named chart/worksheet directly, so an offscreen UI Automation image is not a prerequisite. Clicks with separate intent, a different expected state, different identity or an intervening action are retained. Replay restores and verifies that layout through the Excel capability adapter, without recorded pointer coordinates. A drag without a verified supported outcome is saved as unresolved input and blocks planning rather than becoming a successful click-only replay. Older click-only recordings cannot recover a missing final resize size; re-record that adjustment.

Legend-drag detection is scoped to identified chart descendants, not ordinary worksheet or repeat-button clicks. Planning recovers only the known non-chart misclassification diagnostic from the earlier legend build without rewriting raw events: repeat-button clicks and selections followed by a verified copied range are retained; worksheet gestures lacking a final outcome remain explicit manual-review steps. Genuine delayed input and unresolved chart drags still block planning.
Mouse and keyboard hooks reserve a shared input-order slot before accessibility capture. Completion is delivered on the recorder dispatcher in that order, even if COM/UI Automation pumps nested hooks that finish in reverse order. Reentrant dispatcher callbacks cannot overtake an action currently being processed, and ignored events complete their slots without leaving gaps. This prevents a later mouse-up from being paired with an earlier, still-capturing mouse-down and falsely creating an unresolved chart drag. Raw recordings are not reordered or repaired afterward. Verified source consolidation may remove contiguous same-chart context-menu retries, but never another chart's action or an unresolved gesture. Planning recovers only the known non-chart misclassification diagnostic from the earlier legend build without rewriting raw events: repeat-button clicks and selections followed by a verified copied range are retained; worksheet gestures lacking a final outcome remain explicit manual-review steps. Genuine delayed input and unresolved chart drags still block planning.
Owned compact popup commands are frozen through a direct live MSAA hit before slower UI Automation lookup. The command must be enabled, visible, named, contain the input point, and belong to a popup with a same-process native owner. This generic fast path preserves transient confirmation buttons that disappear on release; it does not infer a missing filter Apply/OK click from subsequent worksheet activity. Unidentified input still blocks planning. Completion is delivered on the recorder dispatcher in that order, even if COM/UI Automation pumps nested hooks that finish in reverse order. Reentrant dispatcher callbacks cannot overtake an action currently being processed, and ignored events complete their slots without leaving gaps. This prevents a later mouse-up from being paired with an earlier, still-capturing mouse-down and falsely creating an unresolved chart drag. Raw recordings are not reordered or repaired afterward. Verified source consolidation may remove contiguous same-chart context-menu retries, but never another chart's action or an unresolved gesture. Planning recovers only the known non-chart misclassification diagnostic from the earlier legend build without rewriting raw events: repeat-button clicks and selections followed by a verified copied range are retained; worksheet gestures lacking a final outcome remain explicit manual-review steps. Genuine delayed input and unresolved chart drags still block planning.
Filter-item MSAA probing is limited to compact owned popup windows, not ordinary worksheet, ribbon or dialog clicks. Direct menu-item hits are captured before slower UIA queries when a native menu is active or the hit belongs to a popup. The existing source-dialog capture scope follows the foreground dialog during range-picker gestures, whose pointer can lie over the worksheet; these inputs clear ordinary drag tracking and require the verified final source outcome rather than being misclassified as legend resizing. The command must be enabled, visible, named, contain the input point, and belong to a popup with a same-process native owner. This generic fast path preserves transient confirmation buttons that disappear on release; it does not infer a missing filter Apply/OK click from subsequent worksheet activity. Unidentified input still blocks planning. Completion is delivered on the recorder dispatcher in that order, even if COM/UI Automation pumps nested hooks that finish in reverse order. Reentrant dispatcher callbacks cannot overtake an action currently being processed, and ignored events complete their slots without leaving gaps. This prevents a later mouse-up from being paired with an earlier, still-capturing mouse-down and falsely creating an unresolved chart drag. Raw recordings are not reordered or repaired afterward. Verified source consolidation may remove contiguous same-chart context-menu retries, but never another chart's action or an unresolved gesture. Planning recovers only the known non-chart misclassification diagnostic from the earlier legend build without rewriting raw events: repeat-button clicks and selections followed by a verified copied range are retained; worksheet gestures lacking a final outcome remain explicit manual-review steps. Genuine delayed input and unresolved chart drags still block planning.

Embedded Excel chart images can expose an unnamed Chart Area or a named legend rather than the chart identity. The recorder samples the hovered chart outside the input hook using native `RangeFromPoint`, with named accessibility ancestors corroborated against native chart objects for resize handles. Mouse-down freezes only a fresh, same-window, matching-bounds snapshot. Native reads are suppressed during dragging and source-dialog interaction. Hover briefly over the chart before opening its context menu or dragging the legend so identity and initial layout can be sampled. A missing initial layout leaves the drag unresolved, not guessed. Replay uses only the chart name and chart-relative final layout. After Select Data closes, the recorder reads the actual changed category/value references and saves a verified `set-chart-source-range` outcome without requiring an intent note. Capture supports a single series with adjacent category/value columns on the chart's worksheet; unsupported confirmed sources remain unresolved. Grouping replaces only the bounded same-process chart-edit sequence with that outcome, retaining raw evidence and unrelated or unowned unresolved input. Replay uses the entire last verified two-column paste when its worksheet and starting cell match, otherwise the captured range, and verifies every category/value. An older recording without chart identity or final legend dimensions must be re-recorded; replay rejects its unnamed chart context action before changing applications instead of guessing a chart or size.

Chart hover sampling rejects invalid accessibility rectangles (including negative-size provider exceptions) and continues checking containing chart ancestors. Native hit objects whose Name is a range object or a non-chart name can use an accessibility chart ancestor only when that name uniquely matches a native chart object. Keyboard capture accepts a hosted control from a different process only when its live accessibility ancestry reaches the foreground window; it retains the control's actual process identity rather than assuming the host owns the editor.
Chart hover sampling also handles a null UI Automation point hit explicitly, recording a diagnostic and publishing no snapshot. A missing element is a transient unavailable sample, not a reason to terminate the background thread or guess a control. The regression check injects repeated null hits on the STA sampling worker and verifies continued sampling and clean shutdown. Native hit objects whose Name is a range object or a non-chart name can use an accessibility chart ancestor only when that name uniquely matches a native chart object. Keyboard capture accepts a hosted control from a different process only when its live accessibility ancestry reaches the foreground window; it retains the control's actual process identity rather than assuming the host owns the editor.
Normal mouse-down control lookup now precedes the filter-specific fallback. A recognized popup control therefore does not wait for a filter ancestry probe; the filter fallback is used only when ordinary lookup returns no snapshot. This removes unnecessary filter probing from identifiable Paste controls without bypassing missing-identity validation. Native hit objects whose Name is a range object or a non-chart name can use an accessibility chart ancestor only when that name uniquely matches a native chart object. Keyboard capture accepts a hosted control from a different process only when its live accessibility ancestry reaches the foreground window; it retains the control's actual process identity rather than assuming the host owns the editor.
The filter-item mouse probe is limited to captionless native popup windows, excluding ordinary dialogs such as Find and Replace before invoking accessibility. The synchronous native dialog probe performs one direct accessibility hit instead of repeating a descendant search at each ancestor. Live dialog regression checks require the unrelated filter probe to reject each dialog command in under 250 ms; filter checklist and confirmation capture remain covered separately. These bounds do not guarantee every accessibility provider responds quickly, and missing input identities still block planning. Native hit objects whose Name is a range object or a non-chart name can use an accessibility chart ancestor only when that name uniquely matches a native chart object. Keyboard capture accepts a hosted control from a different process only when its live accessibility ancestry reaches the foreground window; it retains the control's actual process identity rather than assuming the host owns the editor.
Transient menu identities are captured before popup-button, filter and dialog probes. If a direct MSAA menu hit is unavailable, the popup-only UI Automation fallback runs immediately, while the menu is still open; it must verify a visible, enabled, named menu item at the input point. This avoids recording a newly opened dialog as the command that opened it. No dialog warning is suppressed and no application-specific command name is inferred. Native hit objects whose Name is a range object or a non-chart name can use an accessibility chart ancestor only when that name uniquely matches a native chart object. Keyboard capture accepts a hosted control from a different process only when its live accessibility ancestry reaches the foreground window; it retains the control's actual process identity rather than assuming the host owns the editor.
After a context-click is released, RSR snapshots visible, enabled, named items from compact same-process popup menus outside the low-level input hook. The next click can use that verified identity only within 1.5 s of the latest completed sample (the loop refreshes continuously while the menu stays open; Office search-box menus take about 0.5 s per UIA read), with no intervening input, the same foreground scope, and the same visible popup beneath the click. The foreground scope is the original window or a same-process captionless popup it owns, because Office search-box context menus (for example Excel row and column headers) take foreground themselves; any other foreground ends sampling and is logged in `capture_diagnostics.log`. The `--excel-row-menu-capture` regression right-clicks a scratch workbook's row header and selects Insert while recorder processing is blocked for three seconds. Stop and pause invalidate these snapshots. This handles fast press/release that previously identified the newly opened dialog instead of its menu command; the `--excel-copy-fast-capture` regression opens the menu, allows 250 ms for capture, then presses/releases immediately while recorder processing is blocked for three seconds. Selection before the initial snapshot is ready still uses the synchronous fallback and remains a limitation. No command is inferred from the dialog title or later actions. Native hit objects whose Name is a range object or a non-chart name can use an accessibility chart ancestor only when that name uniquely matches a native chart object. Keyboard capture accepts a hosted control from a different process only when its live accessibility ancestry reaches the foreground window; it retains the control's actual process identity rather than assuming the host owns the editor.

Chart hover sampling runs on a single background STA worker so slow accessibility/native reads do not block the recorder UI and its queued input processing. Requests are coalesced, and an in-progress refresh retains the previous snapshot without extending its 500 ms freshness limit. Samples are rejected if intervening input, a pause, a stop, pointer bounds or foreground ownership invalidate them. Pause discards cached chart identity; Stop cancels sampling and waits for the worker to finish. The regression check holds a sample for 2.6 seconds while checking the UI heartbeat and worker lifecycle; live source and legend checks verify native access from the worker. This scheduling correction adds no new application-specific capability.

The existing native chart adapter verifies the names supplied by the live hit or accessibility chart ancestors through named chart-object lookup rather than enumerating every chart in the worksheet. Legend sampling reuses the resolved chart object, avoiding repeated collection and chart-object reads; ambiguous ancestor matches are still rejected.

Chart-source completion waits for source-dialog input to settle, not subsequent worksheet scrolling or legend work. A verified source outcome replaces only the preceding chart-edit sequence through its last source-dialog input and retains later operations, including unresolved drags. A missing transient menu label does not invalidate an independently verified source outcome.
When no hover snapshot is available, mouse-down resolves a named containing image through bounded live UI Automation ancestry in the same native window/process, with valid bounds containing the input point. This fallback is generic and does not use a later click or legend event to guess an earlier identity. The already-supported chart adapter can arm source capture from this frozen chart/worksheet identity and verifies the actual named chart source after confirmation. The cold-start live check disables hover sampling and targets an unnamed inner Chart Area image before opening Select Data. Missing chart identity or an unverified source-dialog sequence blocks planning before contacting Foundry; a valid later legend outcome is retained, but cannot repair missing source evidence. Semantic chart explanations use the verified capture diagnostic rather than an AI description that might incorrectly claim the dimensions were unspecified. A verified source outcome replaces only the preceding chart-edit sequence through its last source-dialog input and retains later operations, including unresolved drags. A missing transient menu label does not invalidate an independently verified source outcome.

Fresh verified chart snapshots cover ordinary chart clicks as well as context clicks and drags, avoiding a slower accessibility lookup inside the input hook. A same-process tooltip may be sampled through the foreground workbook, but its own rectangle is never treated as chart bounds: the chart name is corroborated natively and the previously identified chart element's live bounds are re-read. Tooltip handling does not search the entire worksheet tree. Legend outcomes are read after mouse release has settled rather than before the application commits the gesture. The live chart check includes multiple charts, ordinary chart selection, deep worksheet scrolling, malformed legend accessibility rectangles, source replacement, dynamic pasted ranges, and an actual verified legend resize.

Recording writes a bounded `capture_diagnostics.log` on Stop, with slow input-capture stage timings and chart snapshot rejection reasons. Diagnostic entries are queued in memory, not written from input hooks, and contain no typed text or recorded pointer coordinates. A same-process tooltip may be sampled through the foreground workbook, but its own rectangle is never treated as chart bounds: the chart name is corroborated natively and the previously identified chart element's live bounds are re-read. Tooltip handling does not search the entire worksheet tree. Legend outcomes are read after mouse release has settled rather than before the application commits the gesture. The live chart check includes multiple charts, ordinary chart selection, deep worksheet scrolling, malformed legend accessibility rectangles, source replacement, dynamic pasted ranges, and an actual verified legend resize.

Pre-click chart sampling reads identity and initial legend layout only; chart source references are read after the source dialog closes, not repeatedly while the user is moving the pointer. Movement within the same verified chart bounds and hit window does not by itself invalidate a completed hover sample. Diagnostics are saved once per recording, including when Stop is followed by disposal. A same-process tooltip may be sampled through the foreground workbook, but its own rectangle is never treated as chart bounds: the chart name is corroborated natively and the previously identified chart element's live bounds are re-read. Tooltip handling does not search the entire worksheet tree. Legend outcomes are read after mouse release has settled rather than before the application commits the gesture. The live chart check includes multiple charts, ordinary chart selection, deep worksheet scrolling, malformed legend accessibility rectangles, source replacement, dynamic pasted ranges, and an actual verified legend resize.

Recovered worksheet gestures can become verified fill operations only with explicit column/boundary intent and a known source: a preceding completed text entry, an explicitly named formula cell, or the last populated formula in the intended column. Nearby formula notes are associated only across same-workbook scrolling and repeated clicks on that cell. A chart-selection gesture without a saved chart identity remains unresolved unless the user explicitly confirms the chart name and source replacement; an immediately subsequent semantic legend record must corroborate that identity. No recorded drag coordinates or endpoint row counts are used.

Explicit formula-fill intent also converts an ordinary cell click when the recorder did not capture the drag outcome. A note attached to the immediately preceding Paste Values result may describe the following formula fill. Conversion requires matching worksheet-column evidence and user scope, emits one fill per contiguous click sequence, and never crosses a worksheet-tab switch.

After a verified sort, a raw Shift+End targeting an old row is removed only when the immediately following operation explicitly selects every row under that same header. The verified all-row operation establishes focus and checks the current selection itself; partial selections and unrelated keyboard actions are retained.

An unannotated title-bar focus click before a same-window verified sort is redundant: the sort activates its own window. A queued End/row-click sequence can be replaced with verified all-row selection only when the subsequent same-view context action explicitly records all-row selection; End alone never implies select-all.

Continuous Excel cell typing that moves into the unnamed in-cell editor is reconstructed against the known starting cell, including Backspace corrections. The final text is replayed once instead of trying to resolve the transient editor. Cursor-navigation keys, other windows and unrelated controls stop reconstruction. An explicit filtered-fill intent following the completed entry replaces a coarse or unverified drag target with a verified visible-row fill; no endpoint coordinates are replayed.

An unidentified delayed mouse event still blocks planning. If the user confirms that it only revealed worksheet tabs, that clarification can be saved on the raw event; planning omits only that explicitly confirmed reveal when a named worksheet click was captured within five seconds. The original unresolved kind, diagnostic and timestamp remain intact. Tab capture avoids redundant UIA ancestor walks, reducing synchronous hook work; this does not guarantee that every transient control can be captured.

Filter confirmation always offers include, exclude and clear-filter choices. If named checkbox values were not captured, enter their exact labels (one per line); an empty include/exclude choice cannot silently become show-all. Explicitly confirmed filter notes may recover a bounded dropdown/checklist/OK sequence containing a misidentified image or worksheet click, but the recorder's raw control evidence remains unchanged. Excel checklist identities also use verified MSAA tree-item names and their Manual Filter ancestry at mouse-down, before a popup can disappear.

New Excel Ctrl+C captures the selected rectangular cell address, and replay restores and verifies that range instead of collapsing it to the active cell. An unavailable selection is saved as unresolved input, not a successful one-cell copy. Only a verified column-header class can start column-resize capture; dragging a worksheet cell edge is not classified as a column resize. Explicit formula-extension intent preserves existing formulas and extends from the last populated formula cell through the last contiguous populated adjacent row, with relative-reference verification.

When an optional action runs and its fallback is skipped, replay waits up to 120 seconds for the next actionable control in the resulting view, skipping inert status-label navigation. For automatic refresh, the saved refresh control must also be available and enabled; a list sort additionally requires an exposed row. Readiness must hold across three probes. Replay never clicks the skipped fallback to force readiness, and stops explicitly if the view fails to become ready.

Every verified list-sort step uses the same 120-second readiness wait, including after ordinary confirmation clicks that start an automatic refresh. The recorded header must be visible and enabled in an enabled window, with an exposed list row, across three probes before replay checks or changes the sort direction. A slow refresh is not treated as a missing header after the short ordinary control lookup.

Use **Intents** to view the selected recording's numbered workflow notes and modify them. Saving updates only the original events' intent text (not captured controls, inputs, timestamps or screenshots). After saving, click **Rebuild selected files**; replay is blocked until the edited notes have been rebuilt into the JSON and PDF. Cancel leaves the notes unchanged; clearing a note removes it.

Replace All consumes its own result acknowledgement for both successful and no-match results. The recorded result OK is not executed again after closing Find and Replace. If a recognized result's OK is missing or cannot dismiss the dialog, replay pauses with the exact message and identifies it as a no-match warning or success result; the user can dismiss it manually and continue only after closure is verified, or stop replay.

The recorder freezes menu command identity and `Target.ParentName` at mouse-down, before queued processing can see a closed or changed menu. A hover can open a submenu without producing a separate click event; replay opens its captured parent through UI Automation. Execution plans and PDFs are generated from recordings, not repaired manually. If an older recording captured the wrong command or omitted its parent, make a new recording with the corrected recorder and generate a new plan.

Transient popup commands that MSAA cannot identify use a popup-only UI Automation hit test at mouse-down. For worksheet tab context-clicks, replay first selects the named tab and reacquires it before hit testing: a tab provider can report `IsOffscreen=false` while its bounds are outside the visible tab strip.

Named controls are identified at mouse-down with physical screen coordinates and bound to their native window before a click can close a dialog or open another one. Verified grid headers retain that identity instead of being replaced by the later focused cell. These coordinates are capture evidence only, not execution-plan locators. A recording with previously misidentified controls must be replaced by a new recording; regeneration cannot recover a command identity that was never captured correctly.

Window-level dialog hits are resolved to the actual native/accessibility button or radio option at mouse-down. Frozen dialog commands are saved without querying their UIA element again after the dialog closes. Embedded grid dropdowns retain their immediate column/header parent so identical filter buttons are not confused. Replay rejects a filter lacking that parent before changing applications. Replacement intent is not evidence that Replace All was clicked: an explicit attributed intent-generated replacement may fulfill it instead. If neither a captured replacement nor an intent-generated replacement fulfills that operation, replay stops rather than silently moving on.

Unidentified dialog-click notices allow background command lookup to finish and recheck the event before prompting. A successfully repaired click no longer causes a stale pause warning. Native wheel capture retries transient window hit tests and follows same-process owner windows for untitled native surfaces; unresolved wheel events retain their delta and axis but still block planning rather than guessing an application. Pausing and discarding recorded actions does not undo changes in the target application. If those actions were not repeated, rebuilding cannot restore them.

Visible native modal dialogs are pre-captured outside input callbacks using a bounded Raw View lookup. A click uses that frozen command only while the same enabled native dialog and unchanged bounds are still under the pointer. This handles child dialogs such as selection warnings without traversing the workbook tree inside the mouse hook or identifying the newly opened dialog instead of the button that opened it.

The same capture covers small owned modeless dialogs, including dialogs whose native owner is hidden. Buttons, tabs, checkboxes and edit fields retain their dialog identity. Tab changes invalidate the old control snapshot immediately; older in-flight snapshots cannot restore it, and periodic refresh captures the new tab's controls. This prevents Find and Replace commands from being recorded as an unidentified window or as the worksheet revealed after closing.

Unidentified-click and pause messages name the raw recording action number, timestamp, application window, last retained action and discarded range. These are recording action numbers, not the later compacted plan's step numbers. Restoring application state remains a manual operation; discarding events does not undo edits. Replacement preflight distinguishes a preparatory dialog tab carrying contextual intent from actual unsubmitted field input: an identical note fulfilled by a later captured replacement in the same application does not require another invented Replace All command. Abandoned typed input still requires review.

Filter state capture probes the live checkbox and label as well as the row center, checking the exact accessible item name. Short labels in wide checklist rows can expose only the tree background at the center. Missing verified states still require confirmation; the message describes missing evidence rather than assuming the recording is old.

Excel filter checklist sequences are consolidated into a `filter-values` outcome: keep only the named values, exclude them, or show all. New recordings capture checked states through UIA/MSAA; older recordings without those states ask for a final outcome before replay and save the choice in the plan and PDF. Duplicate clicks are not treated as evidence of inclusion/exclusion. Replay verifies the baseline and final checkbox states, refuses disabled OK or an empty selection, and waits for the checklist to close before allowing worksheet input. Filter buttons resolve by their recorded column and stable ID, even when their caption changes from "No filter applied" to "Filter applied".

A coarse Excel grid click immediately before a precise recorded cell text entry is unnecessary when it has no selection state or intent: the text entry selects its own address through Go To and verifies it before typing. Such clicks are omitted instead of failing on stale/offscreen grid identities. Meaningful selection clicks remain. Address-based typing navigates across the live viewport without replaying saved scroll distances and refuses hidden, disabled, or unverifiable cells rather than substituting a different row.

For a filtered column entry immediately followed by explicit fill-down-to-adjacent-data intent, the plan uses `TargetStrategy: first-visible-filtered-row`. Replay navigates from the worksheet header horizontally to the destination column and down to the first visible data row, scrolling the live viewport as needed; it does not use the recorded row number or Go To. It verifies adjacent data, enters the value, and fills visible filtered rows until the first empty adjacent data cell. Each cell is verified, hidden rows are left unchanged, and unexpected navigation or changed row order stops replay. Ordinary fixed-address entries without this filtered fill context retain their address-based behavior.

Before that navigation, replay clicks a live, hit-tested visible worksheet cell and verifies cell focus. This prevents a preceding scrollbar/Page left click from receiving worksheet navigation keys. Header/column navigation waits for verified focus rather than assuming Excel updates its accessibility focus immediately.

If a visible data cell already has verified focus in the correct workbook (for example, immediately after the fill seed was committed), replay preserves that focus instead of reactivating the frame and rediscovering cells in the provider's descendant tree. Failure messages distinguish a stopped navigation from earlier cell input that has already completed.

Excel table headers can expose row-1 cells as HeaderItem rather than DataItem. Header navigation accepts either accessible type for row 1; data entry still requires a visible DataItem in a data row.

Dialog capture also resolves clicks on option labels through actionable Raw View ancestors and a bounded Raw View child search. MSAA radio/checkbox roles are retained, and the originating UIA dialog handle takes precedence over an application frame returned by native hit testing. This avoids recording the warning window itself when an owner-drawn button or option was clicked.

To run the interactive menu regression checks on an unlocked desktop:

```powershell
dotnet run --project .\DesktopSteps.Tests\DesktopSteps.Tests.csproj -c Release
```

To check Foundry connectivity and generation with synthetic metadata only:

```powershell
dotnet run --project .\DesktopSteps.Tests\DesktopSteps.Tests.csproj -c Release -- --foundry-check .env
```

The optional Excel dialog check creates and discards its own unsaved scratch workbook. It checks Remove Duplicates warning controls, Find and Replace controls and distinct filter parents without removing duplicates or replacing data:

```powershell
dotnet run --project .\DesktopSteps.Tests\DesktopSteps.Tests.csproj -c Release -- --excel-dialog-capture
```
