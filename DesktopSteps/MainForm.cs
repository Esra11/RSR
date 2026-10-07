using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace DesktopSteps;

internal sealed class MainForm : Form
{
    private readonly Recorder recorder = new();
    private readonly TextBox chat = new() { Multiline = true, ReadOnly = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };
    private readonly TextBox input = new() { Dock = DockStyle.Fill, PlaceholderText = "Ask to record, replay, or open the PDF" };
    private readonly Button record = new ReadableButton() { Text = "Start recording", AutoSize = true };
    private readonly Button pauseRecording = new ReadableButton() { Text = "Pause recording", AutoSize = true, Enabled = false };
    private readonly Button elevatedRecord = new ReadableButton() { Text = "Restart as administrator to record elevated apps", AutoSize = true };
    private readonly Button stop = new ReadableButton() { Text = "Stop Recording...", AutoSize = true, Enabled = false };
    private readonly Button cancelRecording = new ReadableButton() { Text = "Cancel recording", AutoSize = true, Enabled = false };
    private readonly Button finish = new ReadableButton() { Text = "Rebuild selected files", AutoSize = true };
    private readonly Button intents = new ReadableButton() { Text = "Intents", AutoSize = true };
    private readonly Button replay = new ReadableButton() { Text = "Replay selected", AutoSize = true };
    private readonly Button stopReplay = new ReadableButton() { Text = "Stop Replaying", AutoSize = true,
        Enabled = false, TextImageRelation = TextImageRelation.ImageBeforeText,
        ImageAlign = ContentAlignment.MiddleLeft };
    private readonly Button pdf = new ReadableButton() { Text = "Open PDF", AutoSize = true };
    private readonly Button recordingRules = new ReadableButton() { Text = "Add rule before recording. E.g: don't record interaction with notepad", AutoSize = true };
    private readonly Label activeRulesHint = new() { Text = "You have some configured rules in action", AutoSize = true, Visible = false };
    private readonly Button validationReport = new ReadableButton() { Text = "Plan Validation Report not available", AutoSize = true };
    private readonly Button feedbackReport = new ReadableButton() { Text = "Replay Feedback Report not available", AutoSize = true };
    private readonly Button send = new ReadableButton() { Text = "Send", AutoSize = true };
    private readonly Button theme = new ReadableButton() { Text = "Light mode", AutoSize = true };
    private readonly Button clear = new ReadableButton() { Text = "Clear chat", AutoSize = true };
    private readonly Button auth = new ReadableButton() { Text = "Checking sign-in...", AutoSize = true, Enabled = false };
    private readonly Label foundryAccount = new() { Text = "Foundry: checking...", AutoSize = true,
        Margin = new Padding(6, 8, 3, 3) };
    private readonly ComboBox recordings = new() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
    private readonly Button browseRecording = new ReadableButton() { Text = "Browse for execution_plan.json...", AutoSize = true };
    private readonly Label runningAs = new() { Dock = DockStyle.Top, Height = 26,
        Text = "Running as: checking...", TextAlign = ContentAlignment.MiddleLeft,
        Padding = new Padding(10, 0, 0, 0) };
    private readonly ComboBox deleteRecordings = new() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
    private readonly Button deleteRecording = new ReadableButton() { Text = "Delete specific recording", AutoSize = true };
    private readonly Button deleteAllRecordings = new ReadableButton() { Text = "Delete All Recordings", AutoSize = true };
    private readonly Label intentHint = new() { Dock = DockStyle.Top, Height = 44,
        Text = "During recording, please Cntrl+Alt+Shift+i to enter an intent if you think the action needs more explanation or isn't always consistent",
        TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(10, 0, 0, 0) };
    private readonly ToolStripMenuItem recordingHelp = new("Recording");
    private readonly ToolStripMenuItem recordingIntentHelp = new("During recording, please Cntrl+Alt+Shift+i to enter an intent if you think the action needs more explanation or isn't always consistent");
    private const string BrowseRecording = "Browse for execution_plan.json...";
    private readonly ToolTip buttonTips = new() { AutoPopDelay = 8000, InitialDelay = 450, ReshowDelay = 150, ShowAlways = true };
    private readonly ToolTip intentWarning = new() { IsBalloon = true, ShowAlways = true };
    private Form? activeIntentDialog;
    private readonly string sessions = Path.Combine(AppContext.BaseDirectory, "Recordings");
    private readonly string lastRecordingFile = Path.Combine(AppContext.BaseDirectory, ".last-recording");
    private readonly HashSet<string> ignoredProcesses = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<string> recordingRulesText = [];
    private string? activeDirectory;
    private string? selectedDirectory;
    private bool fillingRecordings;
    private bool intentDialogQueued;
    private bool dark = true;
    private bool? signedIn;
    private CancellationTokenSource? work;

    public MainForm()
    {
        Text = "RSR"; Width = 850; Height = 600;
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 200, Padding = new Padding(8) };
        var recordingMenu = new MenuStrip { Dock = DockStyle.Top };
        recordingHelp.DropDownItems.Add(recordingIntentHelp);
        recordingMenu.Items.Add(recordingHelp);
        MainMenuStrip = recordingMenu;
        stopReplay.Image = StopIndicator();
        var authGroup = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
        authGroup.Controls.Add(auth);
        authGroup.Controls.Add(foundryAccount);
        var rulesGroup = new FlowLayoutPanel { AutoSize = true, WrapContents = false,
            FlowDirection = FlowDirection.TopDown, Margin = Padding.Empty };
        rulesGroup.Controls.Add(activeRulesHint);
        rulesGroup.Controls.Add(recordingRules);
        buttons.Controls.AddRange([record, pauseRecording, elevatedRecord, stop, cancelRecording, finish, intents, replay, stopReplay, pdf,
            validationReport, feedbackReport, rulesGroup, clear, theme, authGroup]);
        buttonTips.SetToolTip(record, "Start capturing desktop clicks, typing, and scrolling.");
        buttonTips.SetToolTip(pauseRecording, "Temporarily exclude corrective actions, then resume and repeat the intended step.");
        buttonTips.SetToolTip(stop, "Stop capturing and create the execution plan, summary, and PDF.");
        buttonTips.SetToolTip(cancelRecording, "Stop capturing and discard this unfinished recording without creating files.");
        buttonTips.SetToolTip(finish, "Rebuild the selected recording's plan, summary, and PDF from saved events with help of Foundry.");
        buttonTips.SetToolTip(intents, "View and modify the selected recording's workflow intents, then rebuild selected files.");
        buttonTips.SetToolTip(replay, "Run the selected recording's execution plan on your desktop.");
        buttonTips.SetToolTip(stopReplay, "Stop the replay currently running on your desktop.");
        buttonTips.SetToolTip(pdf, "Open the selected recording's manual instructions PDF.");
        buttonTips.SetToolTip(validationReport, "Open the selected recording's plan check, including any steps that need review before replay.");
        buttonTips.SetToolTip(feedbackReport, "Open what happened during the selected recording's last replay and the suggested next steps.");
        buttonTips.SetToolTip(recordingRules, "Add or remove apps that the recorder should ignore, such as notepad.exe. Rules apply to future recordings.");
        buttonTips.SetToolTip(clear, "Clear messages shown in this window. Saved recordings remain available.");
        buttonTips.SetToolTip(theme, "Switch between dark and light mode.");
        buttonTips.SetToolTip(send, "Send the typed command to RSR.");
        buttonTips.SetToolTip(auth, "Sign in to or out of the shared Azure CLI session used for Foundry.");
        buttonTips.SetToolTip(recordings, "Choose a saved recording or browse to its execution_plan.json file.");
        var picker = new TableLayoutPanel { Dock = DockStyle.Top, Height = 38, ColumnCount = 3, Padding = new Padding(8, 2, 8, 2) };
        picker.ColumnStyles.Add(new(SizeType.AutoSize)); picker.ColumnStyles.Add(new(SizeType.Percent, 100));
        picker.ColumnStyles.Add(new(SizeType.AutoSize));
        picker.Controls.Add(new Label { Text = "Recording:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        picker.Controls.Add(recordings, 1, 0);
        picker.Controls.Add(browseRecording, 2, 0);
        var deletionPicker = new TableLayoutPanel { Dock = DockStyle.Top, Height = 42, ColumnCount = 3, Padding = new Padding(8, 2, 8, 2) };
        deletionPicker.ColumnStyles.Add(new(SizeType.AutoSize));
        deletionPicker.ColumnStyles.Add(new(SizeType.Percent, 100));
        deletionPicker.ColumnStyles.Add(new(SizeType.AutoSize));
        deletionPicker.Controls.Add(deleteRecording, 0, 0);
        deletionPicker.Controls.Add(deleteRecordings, 1, 0);
        deletionPicker.Controls.Add(deleteAllRecordings, 2, 0);
        var bottom = new TableLayoutPanel { Dock = DockStyle.Bottom, Height = 45, ColumnCount = 2 };
        bottom.ColumnStyles.Add(new(SizeType.Percent, 100)); bottom.ColumnStyles.Add(new(SizeType.AutoSize));
        bottom.Controls.Add(input, 0, 0); bottom.Controls.Add(send, 1, 0);
        Controls.Add(chat); Controls.Add(bottom); Controls.Add(intentHint); Controls.Add(deletionPicker); Controls.Add(picker); Controls.Add(buttons); Controls.Add(runningAs); Controls.Add(recordingMenu);
        record.Click += (_, _) => StartOrResumeRecording();
        pauseRecording.Click += (_, _) => PauseRecording();
        elevatedRecord.Click += (_, _) => RestartElevated();
        stop.Click += async (_, _) => await StopRecordingAsync();
        cancelRecording.Click += (_, _) => CancelRecording();
        finish.Click += async (_, _) => await FinishLatestAsync();
        intents.Click += async (_, _) => await EditSelectedIntentsAsync();
        replay.Click += async (_, _) => await ReplayAsync();
        stopReplay.Click += (_, _) => work?.Cancel();
        pdf.Click += (_, _) => OpenPdf();
        validationReport.Click += (_, _) => OpenSelectedReport("validation_report.md");
        feedbackReport.Click += (_, _) => OpenSelectedReport("replay_feedback.md");
        recordingRules.Click += (_, _) => EditRecordingRules();
        send.Click += async (_, _) => await ChatAsync();
        input.KeyDown += async (_, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; await ChatAsync(); } };
        theme.Click += (_, _) => { dark = !dark; ApplyTheme(); };
        clear.Click += (_, _) => chat.Clear();
        auth.Click += async (_, _) => await ToggleSignInAsync();
        recordings.SelectedIndexChanged += (_, _) => ChooseRecording();
        browseRecording.Click += (_, _) => BrowseForExecutionPlan();
        recordings.DropDown += async (_, _) =>
        {
            ShowBrowseFirst();
            await Task.Delay(50);
            ShowBrowseFirst();
        };
        deleteRecording.Click += (_, _) => DeleteSelectedRecording();
        deleteAllRecordings.Click += (_, _) => DeleteAllSavedRecordings();
        Shown += async (_, _) => { await RefreshRunningIdentityAsync(); await RefreshSignInAsync(); };
        recorder.Captured += e =>
        {
            // Key events are saved to the recording, but posting a UI message
            // for every keystroke can lag the recorder and the target program.
            if (e.Kind == "key") return;
            if (e.Kind == "unresolved-input")
            {
                Log("Recording needs attention: " + e.Diagnostic);
                return;
            }
            var message = e.Kind == "context-click" && e.AfterState?.StartsWith("selection-count:") == true &&
                int.TryParse(e.AfterState[16..], out var selected) && selected > 1
                ? "Captured context-click: current selection"
                : $"Captured {e.Kind}: {(string.IsNullOrWhiteSpace(e.Target?.Name) ? e.Key : e.Target.Name)}";
            // Hook callbacks must return promptly so Windows does not delay or
            // drop the user's input. UI logging can run after the hook returns.
            if (IsHandleCreated && !IsDisposed)
                BeginInvoke(new Action(() => { if (!IsDisposed) Log(message); }));
        };
        recorder.TargetCorrected += target =>
        {
            if (IsHandleCreated && !IsDisposed)
                BeginInvoke(new Action(() =>
                {
                    if (!IsDisposed) Log($"Recorder identified an earlier click as {target.Name}.");
                }));
        };
        recorder.UnresolvedDialogClick += (eventIndex, window) =>
        {
            if (!IsHandleCreated || IsDisposed) return;
            BeginInvoke(new Action(() =>
            {
                if (IsDisposed || !recorder.IsRecording || recorder.IsPaused ||
                    !recorder.IsUnresolvedDialogClick(eventIndex)) return;
                var recovery = recorder.DescribeRecovery(eventIndex);
                var pause = MessageBox.Show(this,
                    $"RSR could not identify a control clicked in '{window}'. The step may not replay automatically.\n\n" +
                    recovery + "\n\nPause and remove this action and everything recorded after it? " +
                    "Restore the application to just before the identified action, then resume and repeat that section. " +
                    "RSR does not undo application changes.",
                    "Recorded click needs attention", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (pause == DialogResult.Yes) PauseRecording(eventIndex);
                else Log("Recording continued with an unidentified dialog click. Replay may stop at this step.");
            }));
        };
        recorder.IntentRequested += priorWindow =>
        {
            if (intentDialogQueued) return;
            intentDialogQueued = true;
            BeginInvoke(new Action(() => ShowIntentNote(priorWindow)));
        };
        FormClosing += (_, _) => { work?.Cancel(); recorder.Dispose(); };
        LoadRecordingRules();
        RefreshRecordings(LoadLastRecording());
        ApplyTheme(); Log("Ready. Use the sign-in button if Foundry needs Azure authentication.");
    }

    private async Task RefreshSignInAsync()
    {
        auth.Enabled = false;
        auth.Text = "Checking sign-in...";
        foundryAccount.Text = "Foundry: checking...";
        try
        {
            signedIn = await AzureCli.IsSignedInAsync(CancellationToken.None);
            auth.Text = signedIn.Value ? "Sign out of Foundry" : "Sign in to Foundry";
            if (signedIn.Value)
            {
                try { foundryAccount.Text = "Foundry: " +
                    (await AzureCli.SignedInAccountAsync(CancellationToken.None) ?? "account unavailable"); }
                catch (Exception ex) { foundryAccount.Text = "Foundry: account unavailable"; Log("Azure account lookup failed: " + ex.Message); }
            }
            else foundryAccount.Text = "Foundry: signed out";
            auth.Enabled = true;
        }
        catch (Exception ex)
        {
            signedIn = null;
            auth.Text = "Check sign-in failed";
            foundryAccount.Text = "Foundry: account unavailable";
            Log("Azure sign-in check failed: " + ex.Message);
        }
    }

    private static async Task<string> WhoAmIAsync(CancellationToken token)
    {
        using var process = new Process { StartInfo = new ProcessStartInfo("whoami.exe")
        { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false,
          CreateNoWindow = true } };
        if (!process.Start()) throw new InvalidOperationException("Could not start whoami.exe.");
        var output = await process.StandardOutput.ReadToEndAsync(token);
        var error = await process.StandardError.ReadToEndAsync(token);
        await process.WaitForExitAsync(token);
        if (process.ExitCode != 0) throw new InvalidOperationException("whoami failed: " + error.Trim());
        return output.Trim();
    }

    private async Task RefreshRunningIdentityAsync()
    {
        try { runningAs.Text = "Running as: " + await WhoAmIAsync(CancellationToken.None); }
        catch (Exception ex) { runningAs.Text = "Running as: unavailable"; Log("Could not read whoami: " + ex.Message); }
    }

    private async Task ToggleSignInAsync()
    {
        if (signedIn is null) return;
        auth.Enabled = false;
        auth.Text = signedIn.Value ? "Signing out..." : "Signing in...";
        foundryAccount.Text = signedIn.Value ? "Foundry: signing out..." : "Foundry: signing in...";
        try
        {
            if (signedIn.Value)
            {
                await AzureCli.SignOutAsync(CancellationToken.None);
                Log("Signed out of Azure CLI. Foundry calls now require sign-in.");
            }
            else
            {
                Log("Azure CLI sign-in opened in a console window.");
                await AzureCli.SignInAsync(CancellationToken.None);
                Log("Azure CLI sign-in completed.");
            }
        }
        catch (Exception ex) { Log("Azure sign-in change failed: " + ex.Message); }
        await RefreshSignInAsync();
    }

    private void ApplyTheme()
    {
        var background = dark ? Color.FromArgb(25, 28, 35) : Color.WhiteSmoke;
        var foreground = dark ? Color.WhiteSmoke : Color.Black;
        BackColor = background; ForeColor = foreground;
        foreach (Control control in Controls.Cast<Control>()
            .SelectMany(c => new[] { c }.Concat(AllControls(c))).Prepend(this))
        {
            control.BackColor = background; control.ForeColor = foreground;
        }
        var identityColor = dark ? Color.FromArgb(255, 222, 89) : Color.FromArgb(125, 91, 0);
        runningAs.ForeColor = identityColor;
        foundryAccount.ForeColor = identityColor;
        var hintColor = dark ? Color.FromArgb(135, 206, 250) : Color.FromArgb(0, 85, 150);
        intentHint.ForeColor = hintColor;
        activeRulesHint.ForeColor = dark ? Color.LightGreen : Color.DarkGreen;
        recordingIntentHelp.ForeColor = hintColor;
        recordingHelp.ForeColor = hintColor;
        cancelRecording.ForeColor = Color.Red;
        stopReplay.ForeColor = Color.Red;
        theme.Text = dark ? "Light mode" : "Dark mode";
        UpdateReportButtons();
    }
    private static Bitmap StopIndicator()
    {
        var image = new Bitmap(14, 14);
        using var graphics = Graphics.FromImage(image);
        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);
        using var brush = new SolidBrush(Color.FromArgb(220, 45, 45));
        graphics.FillEllipse(brush, 1, 1, 12, 12);
        return image;
    }
    private static IEnumerable<Control> AllControls(Control parent)
    {
        foreach (Control child in parent.Controls) { yield return child; foreach (var nested in AllControls(child)) yield return nested; }
    }
    private void Log(string message) => chat.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");

    private void ShowIntentNote(nint previousWindow)
    {
        try
        {
            recorder.PrepareIntentDialog();
            using var dialog = new Form { Text = "Add recording intent",
                Width = 570, Height = 265,
                FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false, MinimizeBox = false, TopMost = true };
            activeIntentDialog = dialog;
            var label = new Label { Text = "Describe the intended workflow. A note can apply to several earlier or upcoming steps and request additional actions. Do not enter passwords or PINs.",
                Left = 16, Top = 14, Width = 525, Height = 42 };
            var note = new TextBox { Multiline = true, AcceptsReturn = true, Left = 16, Top = 62, Width = 525, Height = 90,
                PlaceholderText = "Example: Use the date of next Tuesday in the new name." };
            var add = new Button { Text = "Add intent", Left = 16, Top = 170, Width = 110, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Left = 136, Top = 170, Width = 90, DialogResult = DialogResult.Cancel };
            dialog.Controls.AddRange([label, note, add, cancel]);
            dialog.AcceptButton = add; dialog.CancelButton = cancel;
            // A multiline box normally consumes Enter. Recording intent is entered repeatedly,
            // so plain Enter must keep the original quick-submit behavior; Shift+Enter is left
            // untouched and therefore remains the explicit way to insert a newline.
            note.KeyDown += (_, eventArgs) =>
            {
                if (!IsIntentSubmitKey(eventArgs.KeyData)) return;
                eventArgs.Handled = true;
                eventArgs.SuppressKeyPress = true;
                add.PerformClick();
            };
            dialog.ActiveControl = note;
            dialog.Shown += (_, _) =>
            {
                recorder.IntentDialogOpened(dialog.Handle);
                dialog.BeginInvoke(new Action(() =>
                {
                    try { FocusIntentDialog(dialog, note); }
                    catch (Exception ex)
                    {
                        Log("Could not set intent box focus automatically: " + ex.Message);
                        if (!dialog.IsDisposed)
                        {
                            try { dialog.ActiveControl = note; note.Focus(); }
                            catch (Exception fallbackError) { Log("Intent box focus fallback failed: " + fallbackError.Message); }
                        }
                    }
                }));
            };
            dialog.Deactivate += (_, _) =>
            {
                if (dialog.IsDisposed || !dialog.Visible) return;
                try { intentWarning.Show("Close the intent box before continuing the recording.",
                    dialog, 18, dialog.ClientSize.Height - 18, 5000); }
                catch (InvalidOperationException) { }
            };
            if (dialog.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(note.Text))
                Log(recorder.AddIntentNote(note.Text)
                    ? $"Your workflow intent was saved exactly as: \"{note.Text.Trim()}\". Foundry will interpret it using surrounding steps."
                    : "Could not attach the explanation to the recorded action.");
        }
        finally
        {
            activeIntentDialog = null;
            recorder.IntentDialogClosed();
            intentDialogQueued = false;
            if (previousWindow != 0) SetForegroundWindow(previousWindow);
        }
    }

    internal static bool IsIntentSubmitKey(Keys keyData) =>
        // Comparing the complete key data intentionally excludes Shift+Enter and Ctrl+Enter.
        keyData == Keys.Enter || keyData == Keys.Return;

    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(nint window);
    [DllImport("user32.dll")] private static extern nint SendMessage(nint window, uint message, nint wParam, nint lParam);
    [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(nint window, ref ComboBoxInfo info);
    [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] private struct ComboBoxInfo
    {
        public int Size;
        public NativeRect Item, Button;
        public uint ButtonState;
        public nint Combo, Edit, List;
    }
    [DllImport("user32.dll")] private static extern nint GetForegroundWindow();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(nint window, nint processId);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint source, uint target, bool attach);

    private static void FocusIntentDialog(Form dialog, TextBox note)
    {
        if (dialog.IsDisposed) return;
        var foreground = GetForegroundWindow();
        var foregroundThread = GetWindowThreadProcessId(foreground, 0);
        var currentThread = GetCurrentThreadId();
        var attached = foregroundThread != 0 && foregroundThread != currentThread &&
            AttachThreadInput(currentThread, foregroundThread, true);
        try
        {
            SetForegroundWindow(dialog.Handle);
            dialog.Activate();
            dialog.ActiveControl = note;
            note.Focus();
            note.Select();
        }
        finally
        {
            if (attached) AttachThreadInput(currentThread, foregroundThread, false);
        }
    }

    private sealed record RecordingChoice(string Directory, string Label)
    {
        public override string ToString() => Label;
    }

    private void RefreshRecordings(string? preferred = null)
    {
        fillingRecordings = true;
        try
        {
            var wanted = preferred ?? selectedDirectory;
            var folders = Directory.Exists(sessions)
                ? Directory.GetDirectories(sessions).OrderByDescending(x => x).ToList()
                : [];
            if (wanted is not null && Directory.Exists(wanted) &&
                !folders.Contains(wanted, StringComparer.OrdinalIgnoreCase)) folders.Add(wanted);
            recordings.Items.Clear();
            deleteRecordings.Items.Clear();
            recordings.Items.Add(BrowseRecording);
            foreach (var folder in folders.Where(folder => IsOwnedRecordingDirectory(folder) ||
                (folder.Equals(wanted, StringComparison.OrdinalIgnoreCase) &&
                 (File.Exists(Path.Combine(folder, "recorded_events.json")) ||
                  File.Exists(Path.Combine(folder, "execution_plan.json"))))))
            {
                var name = Path.GetFileName(folder);
                var label = DateTime.TryParseExact(name, "yyyyMMdd-HHmmss", null,
                    System.Globalization.DateTimeStyles.None, out var date)
                    ? date.ToString("yyyy-MM-dd HH:mm:ss") : folder;
                var listed = new RecordingChoice(folder, label);
                recordings.Items.Add(listed);
                if (IsOwnedRecordingDirectory(folder)) deleteRecordings.Items.Add(listed);
            }
            var choice = recordings.Items.OfType<RecordingChoice>()
                .FirstOrDefault(item => item.Directory.Equals(wanted, StringComparison.OrdinalIgnoreCase))
                ?? recordings.Items.OfType<RecordingChoice>().FirstOrDefault();
            if (choice is not null) { recordings.SelectedItem = choice; selectedDirectory = choice.Directory; }
            else { recordings.SelectedIndex = 0; selectedDirectory = null; }
            if (deleteRecordings.Items.Count > 0) deleteRecordings.SelectedIndex = 0;
            deleteRecording.Enabled = deleteRecordings.Items.Count > 0;
            deleteAllRecordings.Enabled = deleteRecordings.Items.Count > 0;
        }
        finally { fillingRecordings = false; UpdateReportButtons(); }
    }

    private string? LoadLastRecording()
    {
        try
        {
            if (!File.Exists(lastRecordingFile)) return null;
            var saved = Path.GetFullPath(File.ReadAllText(lastRecordingFile).Trim());
            // Persisted state may only select an existing recording directory. A stale
            // or manually edited state file is ignored instead of escaping the picker.
            return Directory.Exists(saved) &&
                (File.Exists(Path.Combine(saved, "recorded_events.json")) ||
                 File.Exists(Path.Combine(saved, "execution_plan.json"))) ? saved : null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine("Could not restore the last recording: " + ex.Message);
            return null;
        }
    }

    private void SaveLastRecording(string directory)
    {
        try
        {
            var full = Path.GetFullPath(directory);
            if (!Directory.Exists(full)) return;
            File.WriteAllText(lastRecordingFile, full);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine("Could not save the last recording: " + ex.Message);
        }
    }

    private bool IsOwnedRecordingDirectory(string folder)
    {
        var full = Path.GetFullPath(folder);
        var root = Path.GetFullPath(sessions);
        var info = new DirectoryInfo(full);
        var timestampName = DateTime.TryParseExact(info.Name, "yyyyMMdd-HHmmss", null,
            System.Globalization.DateTimeStyles.None, out _);
        return info.Exists && info.Parent is not null &&
            info.Parent.FullName.Equals(root, StringComparison.OrdinalIgnoreCase) &&
            info.LinkTarget is null &&
            (timestampName || File.Exists(Path.Combine(full, "recorded_events.json")) ||
             File.Exists(Path.Combine(full, "execution_plan.json")));
    }

    private void DeleteOwnedRecording(string folder)
    {
        if (!IsOwnedRecordingDirectory(folder)) throw new IOException("Recording folder changed before deletion.");
        var full = Path.GetFullPath(folder);
        var attributes = File.GetAttributes(full);
        if ((attributes & FileAttributes.ReadOnly) != 0)
            File.SetAttributes(full, attributes & ~FileAttributes.ReadOnly);
        // Recheck after changing attributes; LinkTarget rejects actual links elsewhere.
        if (!IsOwnedRecordingDirectory(full)) throw new IOException("Recording folder changed before deletion.");
        Directory.Delete(full, recursive: true);
    }

    private void DeleteSelectedRecording()
    {
        if (work is not null || recorder.IsRecording) { Log("Stop the active operation before deleting a recording."); return; }
        if (deleteRecordings.SelectedItem is not RecordingChoice choice || !IsOwnedRecordingDirectory(choice.Directory)) return;
        if (MessageBox.Show(this, $"Are you sure? This will delete recording {choice.Label}.", "Delete recording",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        try
        {
            DeleteOwnedRecording(choice.Directory);
            Log("Deleted recording: " + choice.Directory);
            RefreshRecordings();
        }
        catch (Exception ex) { Log("Could not delete recording: " + ex); RefreshRecordings(); }
    }

    private void DeleteAllSavedRecordings()
    {
        if (work is not null || recorder.IsRecording) { Log("Stop the active operation before deleting recordings."); return; }
        var folders = Directory.Exists(sessions)
            ? Directory.GetDirectories(sessions).Where(IsOwnedRecordingDirectory).ToArray() : [];
        if (folders.Length == 0) { Log("No saved recordings to delete."); return; }
        if (MessageBox.Show(this, "Are you sure? This will delete all recording.",
            "Delete all recordings", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        var deleted = 0;
        foreach (var folder in folders)
        {
            try
            {
                DeleteOwnedRecording(folder);
                deleted++;
            }
            catch (Exception ex) { Log($"Could not delete {folder}: {ex}"); }
        }
        Log($"Deleted {deleted} of {folders.Length} recordings.");
        RefreshRecordings();
    }

    private void ChooseRecording()
    {
        if (fillingRecordings) return;
        if (recordings.SelectedItem is RecordingChoice choice)
        {
            selectedDirectory = choice.Directory;
            SaveLastRecording(choice.Directory);
            Log("Selected recording: " + choice.Directory);
            UpdateReportButtons();
            return;
        }
        if (recordings.SelectedItem is not string) return;
        BrowseForExecutionPlan();
    }

    private void BrowseForExecutionPlan()
    {
        using var dialog = new OpenFileDialog
        { Title = "Choose an RSR execution_plan.json file",
          Filter = "Execution plans (execution_plan.json)|execution_plan.json|JSON files (*.json)|*.json",
          InitialDirectory = selectedDirectory ?? sessions, CheckFileExists = true };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            var folder = Path.GetDirectoryName(dialog.FileName);
            if (folder is not null && Path.GetFileName(dialog.FileName).Equals("execution_plan.json", StringComparison.OrdinalIgnoreCase))
            {
                RefreshRecordings(folder);
                Log("Selected recording: " + folder);
                return;
            }
            Log("Choose the execution_plan.json file inside the recording folder.");
        }
        RefreshRecordings();
    }

    private void ShowBrowseFirst()
    {
        if (recordings.IsDisposed || !recordings.DroppedDown) return;
        SendMessage(recordings.Handle, 0x015B, 0, 0); // CB_SETTOPINDEX
        var info = new ComboBoxInfo { Size = Marshal.SizeOf<ComboBoxInfo>() };
        if (GetComboBoxInfo(recordings.Handle, ref info) && info.List != 0)
            SendMessage(info.List, 0x0197, 0, 0); // LB_SETTOPINDEX on the popup list.
    }

    private string? SelectedDirectory() => selectedDirectory;

    private void UpdateReportButtons()
    {
        var directory = SelectedDirectory();
        var validationAvailable = directory is not null && File.Exists(Path.Combine(directory, "validation_report.md"));
        var feedbackAvailable = directory is not null && File.Exists(Path.Combine(directory, "replay_feedback.md"));
        validationReport.Text = validationAvailable ? "Open Plan Validation Report" : "Plan Validation Report not available";
        validationReport.ForeColor = !validationAvailable && dark ? Color.LightSkyBlue : ForeColor;
        feedbackReport.Text = feedbackAvailable ? "Open Replay Feedback Report" : "Replay Feedback Report not available";
        feedbackReport.ForeColor = !feedbackAvailable && dark ? Color.LightSkyBlue : ForeColor;
    }

    private void OpenSelectedReport(string name)
    {
        var path = SelectedDirectory() is { } directory ? Path.Combine(directory, name) : null;
        if (path is null || !File.Exists(path)) { UpdateReportButtons(); return; }
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private string RecordingRulesPath => Path.Combine(sessions, "recording_rules.json");

    private void LoadRecordingRules()
    {
        var path = RecordingRulesPath;
        if (!File.Exists(path))
        {
            path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "RSR", "recording_rules.json");
            if (!File.Exists(path)) return;
        }
        try
        {
            var saved = System.Text.Json.JsonSerializer.Deserialize<string[]>(File.ReadAllText(path));
            foreach (var rule in saved ?? []) recordingRulesText.Add(rule);
            RebuildIgnoredProcesses();
            UpdateRulesHint();
            if (!path.Equals(RecordingRulesPath, StringComparison.OrdinalIgnoreCase))
                SaveRecordingRules();
        }
        catch (Exception ex) { Log("Could not load recording rules: " + ex.Message); }
    }

    private void SaveRecordingRules()
    {
        UpdateRulesHint();
        Directory.CreateDirectory(sessions);
        File.WriteAllText(RecordingRulesPath,
            System.Text.Json.JsonSerializer.Serialize(recordingRulesText.ToArray(), JsonFile.Options));
    }

    private void UpdateRulesHint() => activeRulesHint.Visible = recordingRulesText.Count > 0;

    private static string? ExcludedProcess(string rule)
    {
        var match = System.Text.RegularExpressions.Regex.Match(rule.Trim(),
            @"^(?:(?:don['’]t|do not)\s+record\s+(?:interactions?\s+with|activity\s+in)\s+)?(?<app>[A-Za-z0-9_.-]+(?:\.exe)?)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return match.Success ? Path.GetFileNameWithoutExtension(match.Groups["app"].Value) : null;
    }

    private void RebuildIgnoredProcesses()
    {
        ignoredProcesses.Clear();
        foreach (var rule in recordingRulesText)
            if (ExcludedProcess(rule) is { } process) ignoredProcesses.Add(process);
    }

    private void EditRecordingRules()
    {
        if (work is not null || recorder.IsRecording) return;
        using var dialog = new Form { Text = "Recording rules", ClientSize = new Size(700, 420),
            MinimumSize = new Size(560, 370), FormBorderStyle = FormBorderStyle.Sizable,
            StartPosition = FormStartPosition.CenterParent };
        var description = new Label { Dock = DockStyle.Fill, AutoSize = false,
            Text = "Add a standing rule for future recordings. For example: don't record interaction with notepad.exe; or: choose the second visible result after filtering. App exclusions apply during capture. Other rules guide plan creation.",
            Padding = new Padding(12, 10, 12, 4), TextAlign = ContentAlignment.TopLeft };
        var entry = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "Type any recording rule", Margin = new Padding(12, 4, 12, 4) };
        var list = new ListBox { Dock = DockStyle.Fill };
        void RefreshList()
        {
            list.Items.Clear();
            foreach (var rule in recordingRulesText) list.Items.Add(rule);
        }
        RefreshList();
        var add = new Button { Text = "Add rule", AutoSize = true };
        var remove = new Button { Text = "Remove selected", AutoSize = true };
        var close = new Button { Text = "Done", AutoSize = true, DialogResult = DialogResult.OK };
        add.Click += (_, _) =>
        {
            var value = entry.Text.Trim();
            if (value.Length is < 3 or > 1000)
            {
                MessageBox.Show(dialog, "Enter a rule between 3 and 1000 characters.",
                    "Rule needs text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!recordingRulesText.Contains(value, StringComparer.OrdinalIgnoreCase)) recordingRulesText.Add(value);
            RebuildIgnoredProcesses();
            SaveRecordingRules();
            RefreshList();
            entry.Clear();
        };
        remove.Click += (_, _) =>
        {
            if (list.SelectedItem is not string selected) return;
            recordingRulesText.Remove(selected);
            RebuildIgnoredProcesses();
            SaveRecordingRules();
            RefreshList();
        };
        var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10, 6, 0, 0) };
        actions.Controls.AddRange([add, remove, close]);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 88));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        layout.Controls.Add(description, 0, 0);
        layout.Controls.Add(entry, 0, 1);
        layout.Controls.Add(list, 0, 2);
        layout.Controls.Add(actions, 0, 3);
        dialog.Controls.Add(layout);
        dialog.AcceptButton = add;
        dialog.CancelButton = close;
        dialog.ShowDialog(this);
    }

    private void StartOrResumeRecording()
    {
        if (work is not null) { Log("Wait for the current operation before recording."); return; }
        if (recorder.IsPaused)
        {
            recorder.Resume();
            record.Text = "Start recording";
            record.Enabled = false;
            pauseRecording.Enabled = true;
            stop.Enabled = true;
            cancelRecording.Enabled = true;
            Text = "RSR - Recording";
            Log("Recording resumed. Repeat the intended step, then continue the workflow.");
            return;
        }
        StartRecording();
    }

    private void PauseRecording(int? discardFromEvent = null)
    {
        if (!recorder.IsRecording || recorder.IsPaused)
        {
            Log("No active recording is available to pause.");
            return;
        }
        var recovery = discardFromEvent is int index ? recorder.DescribeRecovery(index) : null;
        var discarded = recorder.Pause(discardFromEvent);
        record.Text = "Resume recording";
        record.Enabled = true;
        pauseRecording.Enabled = false;
        Text = "RSR - Recording paused";
        Log("Recording paused. Actions made while paused will not be included.");
        if (discarded > 0)
            Log($"Removed {discarded} recorded action(s) from the unidentified click onward. " +
                "Resume and repeat that part of the workflow.");
        MessageBox.Show(this,
            "Recording is paused.\n\n" +
            (recovery is null ? "" : recovery + "\n\n" +
                $"Removed {discarded} recorded action(s), starting at that action. Earlier actions were kept.\n\n") +
            "1. Return to the target application.\n" +
            "2. Revert the unwanted change and restore the point immediately before the step you want to replace.\n" +
            "3. Return to RSR and choose Resume recording.\n" +
            "4. Repeat the corrected step, then continue the rest of the workflow.\n\n" +
            "Actions performed while paused are excluded from the recording.",
            "Recording paused", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void StartRecording()
    {
        try
        {
            activeDirectory = Path.Combine(sessions, DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            recorder.SetIgnoredProcesses(ignoredProcesses);
            recorder.Start(activeDirectory);
            File.WriteAllText(Path.Combine(activeDirectory, "recording_rules.json"),
                System.Text.Json.JsonSerializer.Serialize(recordingRulesText.ToArray(), JsonFile.Options));
            record.Text = "Start recording"; record.Enabled = false; pauseRecording.Enabled = true;
            stop.Enabled = true; cancelRecording.Enabled = true; finish.Enabled = false;
            recordingRules.Enabled = false;
            Text = "RSR - Recording";
            Log("Recording started. Use the desktop, then return and press Stop.");
            if (!IsRunningAsAdministrator())
                Log("Elevated applications may not be captured. For a UAC-launched app such as DebugDiag, stop and use 'Restart as administrator to record elevated apps' before recording its controls.");
            if (recordingRulesText.Count > 0)
                Log($"{recordingRulesText.Count} standing recording rule(s) saved with this recording.");
            if (ignoredProcesses.Count > 0)
                Log("Recording rule active: ignoring " + string.Join(", ", ignoredProcesses.OrderBy(x => x).Select(x => x + ".exe")) + ".");
        }
        catch (Exception ex) { Log("Recording failed: " + ex.Message); }
    }

    private async Task StopRecordingAsync()
    {
        if (!recorder.IsRecording) { Log("No recording is in progress."); return; }
        recorder.Stop(); record.Text = "Start recording"; record.Enabled = true; pauseRecording.Enabled = false;
        stop.Enabled = false; cancelRecording.Enabled = false; finish.Enabled = true;
        recordingRules.Enabled = true;
        Text = "RSR";
        var directory = activeDirectory;
        activeDirectory = null;
        if (directory is null) return;
        var events = recorder.Events.ToList();
        if (events.Count == 0) { Log("No actions were captured."); return; }
        await BrowserIdentity.EnrichAsync(events);
        if (events.Any(item => item.Target?.Process?.Equals("msedge",
                StringComparison.OrdinalIgnoreCase) == true))
        {
            var browser = events.FirstOrDefault(item => item.Target?.BrowserLaunch is not null)
                ?.Target?.BrowserLaunch;
            var suggestedProfilePath = events.Select(item => item.Target?.BrowserLaunch?.ProfilePath)
                .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));
            var recordProfilePath = ChooseEdgeRecordingMode(suggestedProfilePath,
                out var recordedProfilePath);
            for (var index = 0; index < events.Count; index++)
            {
                if (events[index].Target?.Process?.Equals("msedge",
                    StringComparison.OrdinalIgnoreCase) != true) continue;
                var launch = events[index].Target!.BrowserLaunch ??
                    new BrowserLaunchIdentity("", null, null);
                if (!recordProfilePath)
                    launch = new BrowserLaunchIdentity(launch.ExecutablePath, null, null,
                        AskForFreshProcess: true);
                else
                    launch = new BrowserLaunchIdentity(launch.ExecutablePath, null, null,
                        ProfilePath: recordedProfilePath);
                events[index] = events[index] with
                {
                    Target = events[index].Target! with { BrowserLaunch = launch }
                };
            }
            Log(recordProfilePath
                ? "Recorded the Edge profile path for comparison during replay."
                : "Execution plan will launch a new isolated Edge process automatically.");
        }
        WarnIfRecordingEndsAfterSearchLaunch(events);
        await FinishAsync(directory, events);
    }

    private bool ChooseEdgeRecordingMode(string? suggestedProfilePath, out string? profilePath)
    {
        using var dialog = new Form
        {
            Text = "How should Edge replay?",
            Width = 760,
            Height = 240,
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ControlBox = false
        };
        var description = new Label
        {
            Left = 16, Top = 16, Width = 710, Height = 70,
            Text = "This recording used an already open Edge window. Check its Profile path at " +
                "edge://version and enter it below. That sensitive path will be saved in the plan. " +
                "The new-process option opens another Edge window with the default profile."
        };
        var command = new TextBox
        {
            Left = 16, Top = 95, Width = 710, Height = 28,
            Multiline = false,
            Text = suggestedProfilePath ?? ""
        };
        var recordCurrent = new Button
        {
            Left = 16, Top = 145, Width = 340, Height = 38,
            Text = "Record profile path", DialogResult = DialogResult.Yes
        };
        var openFresh = new Button
        {
            Left = 370, Top = 145, Width = 356, Height = 38,
            Text = "Record opening a new Edge process instead", DialogResult = DialogResult.No
        };
        dialog.Controls.AddRange([description, command, recordCurrent, openFresh]);
        dialog.AcceptButton = recordCurrent;
        while (true)
        {
            if (dialog.ShowDialog(this) != DialogResult.Yes)
            {
                profilePath = null;
                return false;
            }
            var selected = command.Text.Trim().Trim('"');
            if (Directory.Exists(selected))
            {
                profilePath = Path.GetFullPath(selected);
                return true;
            }
            MessageBox.Show(this, "Enter the existing Profile path shown at edge://version, " +
                "or choose to record opening a new Edge process.", "Profile path required",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private bool ApproveBrowserCommandLineChange(string message) =>
        MessageBox.Show(this, message, "Launch new Edge session?",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2) == DialogResult.Yes;

    private static bool IsRunningAsAdministrator()
    {
        using var identity = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
    }

    private void RestartElevated()
    {
        if (recorder.IsRecording || work is not null)
        {
            Log("Stop the active operation before restarting as administrator.");
            return;
        }
        if (IsRunningAsAdministrator())
        {
            Log("RSR is already running as administrator. Start a new recording and include the DebugDiag controls after approving UAC.");
            return;
        }
        try
        {
            var executable = Environment.ProcessPath ?? throw new InvalidOperationException("The executable path is unavailable.");
            var start = new ProcessStartInfo(executable) { UseShellExecute = true, Verb = "runas" };
            Process.Start(start);
            Close();
        }
        catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
        {
            Log("Administrator restart was cancelled. The current recorder remains open.");
        }
        catch (Exception ex)
        {
            Log("Could not restart as administrator: " + ex.Message);
        }
    }

    private void WarnIfRecordingEndsAfterSearchLaunch(IReadOnlyList<RecordedEvent> captured)
    {
        if (IsRunningAsAdministrator()) return;
        var launch = captured.LastOrDefault(item => item.Target?.Process == "SearchHost" &&
            item.Kind == "key" && item.Key == "Enter");
        if (launch is null || captured.Any(item => item.At > launch.At &&
            item.Target?.Process is not ("SearchHost" or "explorer"))) return;
        if (DateTimeOffset.Now - launch.At < TimeSpan.FromSeconds(10)) return;
        Log("Recording warning: no actions in the launched application were captured after Windows Search. If you used an elevated app, this plan will only replay its launch. Restart RSR as administrator, then record the application controls again.");
    }

    private void CancelRecording()
    {
        if (!recorder.IsRecording) { Log("No recording is in progress."); return; }
        recorder.Stop();
        record.Text = "Start recording"; record.Enabled = true; pauseRecording.Enabled = false;
        stop.Enabled = false; cancelRecording.Enabled = false; finish.Enabled = true;
        recordingRules.Enabled = true;
        Text = "RSR";
        var directory = activeDirectory;
        activeDirectory = null;
        try
        {
            if (directory is not null)
            {
                var full = Path.GetFullPath(directory);
                var root = Path.GetFullPath(sessions).TrimEnd(Path.DirectorySeparatorChar);
                if (Path.GetDirectoryName(full)?.Equals(root, StringComparison.OrdinalIgnoreCase) == true &&
                    Directory.Exists(full)) Directory.Delete(full, true);
            }
        }
        catch (Exception ex) { Log("Recording stopped, but its temporary files could not be removed: " + ex.Message); }
        Log("Recording cancelled. No plan or PDF was created.");
    }

    private async Task EditSelectedIntentsAsync()
    {
        if (recorder.IsRecording || work is not null)
        { Log("Stop recording or wait for the current operation before editing intents."); return; }
        var directory = SelectedDirectory();
        if (directory is null) { Log("Select a recording before editing its intents."); return; }
        var path = Path.Combine(directory, "recorded_events.json");
        if (!File.Exists(path)) { Log("This recording has no captured events to edit."); return; }
        using var editing = new CancellationTokenSource();
        work = editing;
        intents.Enabled = false;
        try
        {
            var original = await File.ReadAllTextAsync(path);
            var events = System.Text.Json.JsonSerializer.Deserialize<List<RecordedEvent>>(original, JsonFile.Options)
                ?? throw new InvalidDataException("The recording's captured events are empty.");
            var notes = events.Where(item => !string.IsNullOrWhiteSpace(item.Intent))
                .Select(item => item.Intent!).Distinct(StringComparer.Ordinal).ToArray();
            if (notes.Length == 0)
            { MessageBox.Show(this, "This recording has no saved intents.", "Intents", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using var dialog = new Form { Text = "Modify intents", ClientSize = new Size(800, 440),
                MinimumSize = new Size(650, 360), StartPosition = FormStartPosition.CenterParent };
            var list = new ListBox { Dock = DockStyle.Left, Width = 180 };
            for (var index = 0; index < notes.Length; index++) list.Items.Add($"Intent {index + 1}");
            var editor = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
            var hint = new Label { Dock = DockStyle.Top, Height = 55,
                Text = "Edit each intent, then Save intents. After you're done, click Rebuild selected files to regenerate the plan and PDF.\r\nClear an intent's text to remove it." };
            var save = new Button { Text = "Save intents", AutoSize = true, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 45 };
            footer.Controls.AddRange([save, cancel]);
            var active = -1;
            var edits = notes.ToArray();
            list.SelectedIndexChanged += (_, _) =>
            {
                if (active >= 0) edits[active] = editor.Text;
                active = list.SelectedIndex;
                editor.Text = active >= 0 ? edits[active] : "";
            };
            dialog.Controls.Add(editor); dialog.Controls.Add(list); dialog.Controls.Add(hint); dialog.Controls.Add(footer);
            dialog.CancelButton = cancel;
            list.SelectedIndex = 0;
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            if (active >= 0) edits[active] = editor.Text;
            var changed = notes.Select((note, index) => (note, value: edits[index].Trim()))
                .Where(pair => pair.note != pair.value).ToDictionary(pair => pair.note, pair => pair.value, StringComparer.Ordinal);
            if (changed.Count == 0) { Log("Intents were not changed."); return; }
            if (await File.ReadAllTextAsync(path) != original)
                throw new InvalidOperationException("The recording changed while the intents editor was open. Reopen Intents before saving.");
            var updated = ApplyIntentEdits(events, changed);
            var pending = path + ".intent-edit.tmp";
            try
            {
                await JsonFile.SaveAsync(pending, updated);
                await File.WriteAllTextAsync(Path.Combine(directory, "intents_need_rebuild.txt"),
                    "Workflow intents were edited. Click Rebuild selected files before replay.");
                File.Move(pending, path, overwrite: true);
            }
            finally { if (File.Exists(pending)) File.Delete(pending); }
            Log($"Saved changes to {changed.Count} intent(s). Click Rebuild selected files before replaying.");
        }
        catch (Exception ex)
        {
            Log("Could not save intents: " + ex.Message);
            MessageBox.Show(this, ex.Message, "Intents could not be saved", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally { intents.Enabled = true; work = null; }
    }

    internal static List<RecordedEvent> ApplyIntentEdits(IReadOnlyList<RecordedEvent> events,
        IReadOnlyDictionary<string, string> changes) =>
        events.Select(item => item.Intent is not null && changes.TryGetValue(item.Intent, out var edited)
            ? item with { Intent = string.IsNullOrWhiteSpace(edited) ? null : edited.Trim() } : item).ToList();

    private async Task FinishLatestAsync()
    {
        if (recorder.IsRecording) { await StopRecordingAsync(); return; }
        if (work is not null) { Log("Files are still being created. Please wait for completion."); return; }
        var directory = SelectedDirectory();
        if (directory is null) { Log("No recording found."); return; }
        var path = Path.Combine(directory, "recorded_events.json");
        if (!File.Exists(path)) { Log("The latest recording has no captured events file."); return; }
        List<RecordedEvent>? events;
        using (var loading = new CancellationTokenSource())
        {
            work = loading;
            try { events = await JsonFile.LoadAsync<List<RecordedEvent>>(path, loading.Token); }
            catch (Exception ex) { Log("Could not load captured events for rebuilding: " + ex.Message); return; }
            finally { work = null; }
        }
        if (events is null || events.Count == 0) { Log("The captured events file is empty."); return; }
        await FinishAsync(directory, events);
    }

    private async Task FinishAsync(string directory, List<RecordedEvent> events)
    {
        if (work is not null) { Log("An operation is already running."); return; }
        work = new CancellationTokenSource();
        var planSaved = false;
        try
        {
            var removedSyntheticClicks = events.RemoveAll(item => item.Kind == "click" &&
                item.AfterState == "dialog-command" && item.ClickX is null && item.ClickY is null);
            if (removedSyntheticClicks > 0)
                Log($"Removed {removedSyntheticClicks} accessibility events that had no recorded user click.");
            events = ExpandDialogFieldSnapshots(RecoverRecordedDialogCommands(events));
            await JsonFile.SaveAsync(Path.Combine(directory, "recorded_events.json"), events, work.Token);
            Log($"Captured {events.Count} actions. Asking Foundry to infer intent...");
            var rulesPath = Path.Combine(directory, "recording_rules.json");
            var rules = File.Exists(rulesPath)
                ? System.Text.Json.JsonSerializer.Deserialize<string[]>(await File.ReadAllTextAsync(rulesPath, work.Token)) ?? []
                : [];
            var plan = RefreshTiming.AddObservedWait(
                await new Foundry(EnvPath()).PlanAsync(events, work.Token, rules), events);
            await JsonFile.SaveAsync(Path.Combine(directory, "execution_plan.json"), plan, work.Token);
            planSaved = true;
            await File.WriteAllTextAsync(Path.Combine(directory, "summary.txt"), plan.Summary, work.Token);
            await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"), work.Token);
            await WriteValidationReportAsync(directory, plan, work.Token);
            File.Delete(Path.Combine(directory, "intents_need_rebuild.txt"));
            Log($"Created execution_plan.json, summary.txt, and manual_steps.pdf in {directory}");
            if (ReplacementPlanProblem(plan) is { } replacementProblem)
                Log("Plan needs review: " + replacementProblem);
        }
        catch (Exception ex)
        {
            Log("Planning failed; captured events were preserved: " + ex.Message);
            if (!planSaved && File.Exists(Path.Combine(directory, "execution_plan.json")))
                Log("The previous execution plan is still saved; replay would use that older plan, not this failed rebuild.");
        }
        finally { work.Dispose(); work = null; RefreshRecordings(directory); }
    }

    internal static List<RecordedEvent> ExpandDialogFieldSnapshots(List<RecordedEvent> source)
    {
        var expanded = new List<RecordedEvent>(source.Count);
        foreach (var item in source)
        {
            if (item.Kind == "click" && item.Target?.Name == "Replace All" &&
                item.BeforeState?.StartsWith("dialog-fields:", StringComparison.Ordinal) == true)
            {
                try
                {
                    var fields = System.Text.Json.JsonSerializer.Deserialize<List<DialogFieldSnapshot>>(
                        item.BeforeState["dialog-fields:".Length..]);
                    foreach (var field in fields ?? [])
                    {
                        if (field.Target.Window != item.Target.Window ||
                            string.IsNullOrWhiteSpace(field.Target.AutomationId)) continue;
                        if (expanded.Any(existing => existing.Kind == "type" &&
                            existing.Target?.Window == field.Target.Window &&
                            existing.Target.AutomationId == field.Target.AutomationId &&
                            existing.AfterState == "field-value:" + field.Value &&
                            item.At - existing.At >= TimeSpan.Zero &&
                            item.At - existing.At < TimeSpan.FromSeconds(1))) continue;
                        expanded.Add(new RecordedEvent(item.At.AddTicks(-1), "type", field.Target,
                            field.Value, null, null, null, "field-value:" + field.Value));
                    }
                }
                catch (System.Text.Json.JsonException) { /* Keep the original click for plan review. */ }
            }
            expanded.Add(item);
        }
        return expanded;
    }

    private async Task ReplayAsync()
    {
        var replayTrace = new List<string>();
        if (recorder.IsRecording) { Log("Stop recording before replaying a plan."); return; }
        if (work is not null) { Log("An operation is already running."); return; }
        var directory = SelectedDirectory();
        if (directory is null) { Log("No recording found."); return; }
        if (File.Exists(Path.Combine(directory, "intents_need_rebuild.txt")))
        {
            const string message = "This recording's intents were modified. Click Rebuild selected files before replaying; the saved plan still contains the previous intents.";
            Log(message);
            MessageBox.Show(this, message, "Rebuild required", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        var path = Path.Combine(directory, "execution_plan.json");
        if (!File.Exists(path)) { Log("This recording has no execution plan."); return; }
        // "Last used" means the plan that actually reached Replay, not merely the
        // last dropdown item highlighted. Persist it before any validation or UI
        // refresh can change the current selection.
        SaveLastRecording(directory);
        Log("Using execution plan: " + path);
        work = new CancellationTokenSource();
        stopReplay.Enabled = true;
        replay.Enabled = false;
        recordingRules.Enabled = false;
        try
        {
            var loaded = await JsonFile.LoadAsync<ExecutionPlan>(path, work.Token)
                ?? throw new InvalidDataException("Empty execution plan.");
            var plan = PlanCompactor.Compact(loaded);
            var repaired = !plan.Steps.SequenceEqual(loaded.Steps);
            for (var i = 0; i < plan.Steps.Count; i++)
            {
                var step = plan.Steps[i];
                if (step.Action is not ("click" or "context-click") ||
                    step.Target?.ControlType != "ControlType.ToolTip") continue;
                plan.Steps[i] = step with { Action = "manual-click" };
                repaired = true;
                Log($"Step {step.Number}: recorded a pop-up hint instead of the intended control; replay will ask you to make this click.");
            }
            var eventsPath = Path.Combine(directory, "recorded_events.json");
            if (File.Exists(eventsPath) &&
                await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath, work.Token) is { } recordedEvents)
            {
                if (RecoverOmittedWorksheetNavigation(plan, recordedEvents, out var recoveredTabs))
                {
                    repaired = true;
                    foreach (var recoveredTab in recoveredTabs)
                        Log($"Recovered recorded worksheet navigation to '{recoveredTab}' before the next grid action.");
                }
                for (var index = 0; index < plan.Steps.Count; index++)
                {
                    var step = plan.Steps[index];
                    if (step.Action != "manual-click" || step.Target?.ControlType != "ControlType.Window" ||
                        string.IsNullOrWhiteSpace(step.Screenshot)) continue;
                    var recordedClose = recordedEvents.FirstOrDefault(item =>
                        item.Kind == "click" && item.AfterState == "window-closed" &&
                        item.Screenshot == step.Screenshot &&
                        item.Target?.Process?.Equals(step.Target.Process,
                            StringComparison.OrdinalIgnoreCase) == true &&
                        item.Target.Window == step.Target.Window);
                    if (recordedClose is null) continue;
                    plan.Steps[index] = step with
                    {
                        Action = "close-window",
                        ExpectedState = "window-closed",
                        Explanation = $"Close {step.Target.Window ?? "the current window"} before continuing."
                    };
                    repaired = true;
                    Log($"Repaired step {step.Number} from its recorded window-closed result.");
                }
                var lostNavigationEnter = plan.Steps.Any(planned =>
                    planned.Action == "type" && planned.Target is
                        { ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } editId,
                            Process: { Length: > 0 } process } &&
                    Uri.TryCreate(planned.Value, UriKind.Absolute, out var address) &&
                    address.Query.Length > 0 &&
                    recordedEvents.Any(recorded => recorded is
                        { Kind: "key", Key: "Enter" or "Return", Target:
                            { ControlType: "ControlType.Edit" } } &&
                        recorded.Target.Process?.Equals(process, StringComparison.OrdinalIgnoreCase) == true &&
                        recorded.Target.AutomationId == editId) &&
                    !plan.Steps.Any(key => key.Action == "key" && key.Key is "Enter" or "Return" &&
                        key.Target?.Process?.Equals(process, StringComparison.OrdinalIgnoreCase) == true &&
                        key.Target.AutomationId == editId));
                if (lostNavigationEnter)
                {
                    const string message = "This saved plan replaced recorded browser typing and Enter keys " +
                        "with a navigation result URL. Select this recording and click Rebuild selected files " +
                        "before replaying it; the original events are still available.";
                    Log(message);
                    MessageBox.Show(this, message, "Plan needs rebuild",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var regroupedEvents = ActionGrouper.Group(recordedEvents);
                for (var groupedIndex = 0; groupedIndex + 1 < regroupedEvents.Count; groupedIndex++)
                {
                    var typed = regroupedEvents[groupedIndex];
                    var submitted = regroupedEvents[groupedIndex + 1];
                    if (typed is not { Kind: "type", Value: { Length: > 0 }, Target:
                            { ControlType: "ControlType.Edit", AutomationId: { Length: > 0 } editId,
                                Process: { Length: > 0 } process } } ||
                        submitted is not { Kind: "key", Key: "Enter" or "Return" } ||
                        submitted.Target?.Process?.Equals(process,
                            StringComparison.OrdinalIgnoreCase) != true ||
                        submitted.Target.AutomationId != editId ||
                        submitted.BeforeState != "field-value:" + typed.Value) continue;
                    var enterIndex = plan.Steps.FindIndex(step => step.Action == "key" &&
                        step.Key is "Enter" or "Return" &&
                        step.Target?.Process?.Equals(process,
                            StringComparison.OrdinalIgnoreCase) == true &&
                        step.Target.AutomationId == editId &&
                        step.Target.Window == submitted.Target.Window);
                    if (enterIndex < 0) continue;
                    var start = enterIndex;
                    while (start > 0 && plan.Steps[start - 1] is { Action: "type" or "key" } edit &&
                        edit.Target?.Process?.Equals(process,
                            StringComparison.OrdinalIgnoreCase) == true &&
                        edit.Target.AutomationId == editId) start--;
                    if (enterIndex - start < 2 ||
                        !plan.Steps.Skip(start).Take(enterIndex - start).Any(step =>
                            step.Action == "key" && step.Key is "Back" or "Backspace" or "Delete"))
                        continue;
                    var firstType = plan.Steps.Skip(start).Take(enterIndex - start)
                        .FirstOrDefault(step => step.Action == "type");
                    if (firstType is null) continue;
                    plan.Steps.RemoveRange(start, enterIndex - start);
                    plan.Steps.Insert(start, firstType with { Value = typed.Value,
                        Target = firstType.Target! with { Name = typed.Value },
                        ExpectedState = "field-value:" + typed.Value });
                    plan = plan with { Steps = plan.Steps.Select((step, position) =>
                        step with { Number = position + 1 }).ToList() };
                    repaired = true;
                    Log($"Repaired recorded address-bar corrections into one verified value: '{typed.Value}'.");
                }
                if (regroupedEvents.Any(item =>
                        item.Kind == "open-private-window") &&
                    !plan.Steps.Any(step => step.Action == "open-private-window"))
                {
                    const string message = "This saved plan omitted a recorded browser click that opened " +
                        "a new private window. Select this recording and click Rebuild selected files " +
                        "before replaying; the original events still contain the transition.";
                    Log(message);
                    MessageBox.Show(this, message, "Plan needs rebuild",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (plan.Steps is [{ Action: "type", Target:
                        { Process: "SearchHost", AutomationId: "SearchTextBox" }, Value: "" } emptySearch])
                {
                    var groupedSearch = ActionGrouper.Group(recordedEvents);
                    if (groupedSearch is [{ Kind: "type", Value: { Length: > 0 } query,
                            Target: { Process: "SearchHost", AutomationId: "SearchTextBox" } searchTarget },
                        { Kind: "key", Key: "Enter" or "Return",
                            Target: { Process: "SearchHost", AutomationId: "SearchTextBox" } } enter])
                    {
                        plan = plan with { Steps =
                        [
                            emptySearch with { Target = searchTarget with { Name = query }, Value = query,
                                ExpectedState = "field-value:" + query,
                                Explanation = "Enter the recorded Windows Search query." },
                            new PlanStep(2, "key", enter.Target, null, enter.Key,
                                enter.AfterState, enter.Screenshot,
                                "Activate the selected Windows Search result.")
                        ] };
                        repaired = true;
                        Log("Repaired Windows Search recording from its raw query and Enter key.");
        }
    }

                if (recordedEvents.Any(item => item.Kind == "click" &&
                    item.AfterState == "dialog-command" && item.ClickX is null && item.ClickY is null))
                {
                    const string message = "This recording contains app-generated accessibility events saved as clicks. " +
                        "Replay stopped before those steps. Use Rebuild selected files to remove them from the saved events and plan.";
                    Log(message);
                    MessageBox.Show(this, message, "Recording needs rebuild",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var timedPlan = RefreshTiming.AddObservedWait(plan, recordedEvents);
                if (!ReferenceEquals(timedPlan, plan)) { plan = timedPlan; repaired = true; }
            }
            var filterTotal = plan.Steps.Count(step => step.Action == "filter-values");
            var filterOrdinal = 0;
            var priorFilterValues = new Dictionary<string, (string[] Values, int Ordinal)>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < plan.Steps.Count; i++)
            {
                var step = plan.Steps[i];
                if (step.Action == "filter-values") filterOrdinal++;
                if (step.Action == "filter-values" && step.ExpectedState == "filter:unspecified")
                {
                    var filterKey = string.Join("\u001f", step.Target?.Window, step.Target?.ParentName);
                    priorFilterValues.TryGetValue(filterKey, out var prior);
                    var selection = AskFilterSelection(step, filterOrdinal, filterTotal,
                        prior.Values ?? [], prior.Ordinal);
                    if (selection is null) { Log("Replay cancelled: filter selection was not confirmed."); return; }
                    plan.Steps[i] = ExcelFilterPlan.Confirm(step, selection.Value.Mode, selection.Value.Values);
                    repaired = true;
                    Log(plan.Steps[i].Explanation!);
                    step = plan.Steps[i];
                }
                if (step.Action == "filter-values" && ExcelFilterPlan.Values(step) is { Length: > 0 } confirmedValues)
                    priorFilterValues[string.Join("\u001f", step.Target?.Window, step.Target?.ParentName)] =
                        (confirmedValues, filterOrdinal);
                if (step.Action == "filter-values") continue;
                if (step.Action != "ensure-state" || step.Target?.ControlType != "ControlType.HeaderItem" ||
                    step.ExpectedState?.Contains("sort", StringComparison.OrdinalIgnoreCase) != true ||
                    step.ExpectedState.Contains("ascending", StringComparison.OrdinalIgnoreCase) ||
                    step.ExpectedState.Contains("descending", StringComparison.OrdinalIgnoreCase)) continue;
                var direction = AskSortDirection(step.Target.Name ?? "this column");
                if (direction is null) { Log("Replay cancelled: sort direction was not selected."); return; }
                plan.Steps[i] = step with { ExpectedState = "sort:" + direction };
                repaired = true;
                Log($"Selected {direction} sort for {step.Target.Name}.");
            }
            if (repaired)
            {
                await JsonFile.SaveAsync(path, plan, work.Token);
                await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"), work.Token);
            }
            if (File.Exists(eventsPath) &&
                await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath, work.Token) is { } recoveryEvents)
            {
                var recovered = PlanCompactor.RecoverExplicitDuplicatedRowIntent(plan,
                    recoveryEvents.Select(item => item.Intent).Where(note => note is not null).Select(note => note!));
                recovered = ExcelFilterPlan.RecoverRecordedOutcomes(recovered, recoveryEvents);
                if (!recovered.Steps.SequenceEqual(plan.Steps))
                {
                    plan = recovered;
                    await JsonFile.SaveAsync(path, plan, work.Token);
                    await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"), work.Token);
                    Log("Recovered recorded filter outcomes and explicit worksheet intent from original capture evidence.");
                }
            }
            await WriteValidationReportAsync(directory, plan, work.Token);
            if (plan.Steps.Any(step => step.OriginIntent is not null))
            {
                var sourceEvents = File.Exists(eventsPath)
                    ? await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath, work.Token) : null;
                if (sourceEvents is null || plan.Steps.Any(step => step.OriginIntent is not null &&
                    !sourceEvents.Any(item => item.Intent == step.OriginIntent)))
                    throw new InvalidDataException("Intent-generated actions could not be linked to their original recorded notes. Replay stopped before changing applications.");
                Log($"Plan notice: {plan.Steps.Count(step => step.OriginIntent is not null)} action(s) were generated from explicit workflow intent, not captured clicks. Review them in the PDF.");
            }
            if (File.Exists(eventsPath) &&
                await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath, work.Token) is { } intentEvents)
            {
                var omittedIntents = intentEvents.Select(item => item.Intent)
                    .Where(note => !string.IsNullOrWhiteSpace(note)).Select(note => note!).Distinct(StringComparer.Ordinal)
                    .Where(note => !RecordedIntentIsRepresented(plan, note)).ToArray();
                if (omittedIntents.Length > 0)
                    throw new InvalidDataException("The generated plan omitted recorded workflow intent: " +
                        string.Join(" | ", omittedIntents.Select(note => $"\"{note}\"")) +
                        ". Edit the intent if it was saved incompletely, then rebuild the selected files before replay.");
            }
            if (ReplacementPlanProblem(plan) is { } replacementProblem)
            {
                Log(replacementProblem);
                MessageBox.Show(this, replacementProblem, "Replacement needs review",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (plan.Steps.Any(step => step.Action == "replace-all") && File.Exists(eventsPath) &&
                await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath, work.Token) is { } originalClicks &&
                !originalClicks.Any(item => item.Kind == "click" &&
                    (item.Target?.Name == "Replace All" || item.AfterState == "replace-all-result")) &&
                plan.Steps.Any(step => step.Action == "replace-all" &&
                    (step.OriginIntent is null || !originalClicks.Any(item => item.Intent == step.OriginIntent))))
            {
                var message = "This saved plan contains Replace All, but the recorded clicks do not. Replay stopped because the plan would press a button the recorder never captured. Please record the replacement again.";
                Log(message);
                MessageBox.Show(this, message, "Plan does not match recording",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (IncompleteReplacements(plan) is { Count: > 0 } missingReplacements)
            {
                var message = $"The recording entered text or specified replacement intent in Find and Replace near step {missingReplacements[0]}, but no Replace All click was saved before it moved on. Replay has stopped so it cannot silently skip your replacement. Record that part again; the app will not invent the missing click.";
                Log("Replay stopped: " + message);
                MessageBox.Show(this, message, "Recording missed Replace All",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (plan.Steps.Any(step => step.Action == "ensure-state" && step.ExpectedState == "sort:unspecified"))
            {
                foreach (var step in plan.Steps.Where(step => step.Action == "ensure-state" &&
                    step.ExpectedState == "sort:unspecified").ToArray())
                {
                    using var question = new Form { Text = "Choose recorded sort outcome", Width = 460, Height = 180,
                        StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                        MaximizeBox = false, MinimizeBox = false };
                    question.Controls.Add(new Label { Text = $"Sort '{step.Target?.Name}' in which direction?\nThe recording did not capture a final direction.",
                        Dock = DockStyle.Top, Height = 65 });
                    var descending = new Button { Text = "Descending", Left = 25, Top = 75, Width = 120, DialogResult = DialogResult.Yes };
                    var ascending = new Button { Text = "Ascending", Left = 155, Top = 75, Width = 120, DialogResult = DialogResult.No };
                    var cancel = new Button { Text = "Cancel replay", Left = 285, Top = 75, Width = 120, DialogResult = DialogResult.Cancel };
                    question.Controls.AddRange([descending, ascending, cancel]);
                    question.CancelButton = cancel;
                    var choice = question.ShowDialog(this);
                    if (choice is not (DialogResult.Yes or DialogResult.No))
                    { Log("Replay cancelled before changing applications: sort direction was not chosen."); return; }
                    var direction = choice == DialogResult.Yes ? "descending" : "ascending";
                    plan = plan with { Steps = plan.Steps.Select(candidate => candidate.Number == step.Number
                        ? candidate with { ExpectedState = "sort:" + direction,
                            Explanation = $"User selected {direction} order before replay." } : candidate).ToList() };
                    Log($"Saved replay choice: sort '{step.Target?.Name}' {direction}.");
                }
                await JsonFile.SaveAsync(Path.Combine(directory, "execution_plan.json"), plan, work.Token);
                await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"), work.Token);
            }
            var planWarnings = ValidatePlanTransitions(plan);
            var blockingWarnings = planWarnings.Where(IsBlockingPlanWarning).ToList();
            foreach (var notice in planWarnings.Except(blockingWarnings))
                Log("Plan notice: " + notice);
            if (blockingWarnings.Count > 0)
            {
                foreach (var warning in blockingWarnings) Log("Plan warning: " + warning);
                Log("Replay stopped before changing applications because the execution plan contains unresolved actions. " +
                    "Rebuild or re-record those actions; replay will not ask for approval to run an interactive plan.");
                return;
            }
            if (RefreshTiming.HasRefreshStep(plan))
            {
                var expectedSeconds = AskRefreshDuration(plan.RefreshExpectedSeconds ?? 18);
                if (expectedSeconds is null)
                { Log("Replay cancelled before changing applications: refresh duration was not confirmed."); return; }
                plan = plan with { RefreshExpectedSeconds = expectedSeconds };
                await JsonFile.SaveAsync(path, plan, work.Token);
                Log($"Refresh timeout: {expectedSeconds} seconds expected + 2 seconds buffer = {expectedSeconds + 2} seconds; verified completion continues early.");
            }
            var identity = await WhoAmIAsync(work.Token);
            await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"), work.Token);
            await File.AppendAllTextAsync(Path.Combine(directory, "replay_identity.txt"),
                $"{DateTimeOffset.Now:O} {identity}{Environment.NewLine}", work.Token);
            runningAs.Text = "Running as: " + identity;
            Log("Replay identity (whoami): " + identity);
            Log("Replay started. Leave the desktop available.");
            var incompleteReplacements = await Executor.ReplayAsync(plan, message =>
            {
                replayTrace.Add(message);
                Log(message);
            }, AskToHandleUnexpectedDialog, ApproveBrowserCommandLineChange, work.Token);
            Log(incompleteReplacements
                ? "Replay finished with earlier replacement mappings missing Replace All actions in the recording."
                : "Replay complete.");
            var manualActions = replayTrace.Where(line => line.Contains("user handled", StringComparison.OrdinalIgnoreCase)).ToList();
            if (manualActions.Count > 0)
                await File.WriteAllTextAsync(Path.Combine(directory, "replay_feedback.md"),
                    "# Replay feedback\n\nReplay reached the end, but these actions needed user help:\n\n" +
                    string.Join("\n", manualActions.Select(line => "- " + HumanReplayLine(line))) +
                    "\n\n## What to do next\n\n1. Open the PDF and check the screenshot for each listed step.\n" +
                    "2. If you want those steps to run automatically, record them again. Rebuilding the same saved recording cannot fill in a missing click.\n", work.Token);
        }
        catch (OperationCanceledException) { Log("Replay stopped by user."); }
        catch (Exception ex)
        {
            var diagnostic = Path.Combine(directory, "replay_diagnostics.log");
            await File.AppendAllTextAsync(diagnostic, $"{DateTimeOffset.Now:O}{Environment.NewLine}{ex}{Environment.NewLine}");
            var reason = string.IsNullOrWhiteSpace(ex.Message) ? ex.GetType().Name : ex.Message;
            await File.WriteAllTextAsync(Path.Combine(directory, "replay_feedback.md"),
                "# Replay feedback\n\n" +
                $"What stopped replay: {HumanReplayLine(reason)}\n\n" +
                "## What happened just before it stopped\n\n" +
                string.Join("\n", replayTrace.TakeLast(20).Select(line => "- " + HumanReplayLine(line))) + "\n\n" +
                "## What to do next\n\n" +
                ReplayRecoveryActions(reason), work.Token);
            Log($"Replay stopped: {reason} Details: {diagnostic}");
        }
        finally { stopReplay.Enabled = false; replay.Enabled = true; recordingRules.Enabled = true;
            work.Dispose(); work = null; UpdateReportButtons(); }
    }

    private static bool RecordedIntentIsRepresented(ExecutionPlan plan, string note)
    {
        if (plan.Steps.Any(step => step.Intent == note || step.OriginIntent == note)) return true;
        // A lone deictic word does not identify an action, target or outcome. The recorder keeps
        // it verbatim, but replay must not invent a mandatory operation or reject a safely
        // compacted plan because an incomplete note such as "previous" was attached to a click.
        if (note.Trim().Equals("previous", StringComparison.OrdinalIgnoreCase)) return true;
        var duplicatePreviousRow = System.Text.RegularExpressions.Regex.IsMatch(note,
            @"\bcopy\s+(?:the\s+)?(?:functions?|formulas?)\s+of\s+(?:the\s+)?(?:previous\s+row|last\s+row\s+containing\s+values?)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        var weekday = System.Text.RegularExpressions.Regex.Match(note,
            @"\bnext\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (duplicatePreviousRow && weekday.Success)
        {
            var duplicated = plan.Steps.FindIndex(step => step.Action == "duplicate-range-values-and-formulas");
            return duplicated >= 0 && duplicated + 1 < plan.Steps.Count &&
                plan.Steps[duplicated + 1] is { Action: "update-cell-to-relative-weekday", RelativeWeekday: { } day } &&
                day.Equals(weekday.Groups[1].Value, StringComparison.OrdinalIgnoreCase);
        }
        // Chart source capture is recorded after Excel closes Select Data, so the intent note
        // entered while that dialog is open may not be copied onto the resulting semantic step.
        // Treat a verified chart source as representing an instruction to use the cells pasted
        // in the preceding step. Keep all three wording checks so unrelated "add data" notes do
        // not bypass the omitted-intent safeguard.
        var pastedCellsAsData = System.Text.RegularExpressions.Regex.IsMatch(note,
            @"\badd\s+data\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase) &&
            System.Text.RegularExpressions.Regex.IsMatch(note,
                @"\bcells?\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase) &&
            System.Text.RegularExpressions.Regex.IsMatch(note,
                @"\bpast(?:e|ed|ing)\b.*\bprevious\s+step\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (pastedCellsAsData)
            return plan.Steps.Any(step => step is { Action: "set-chart-source-range",
                ExpectedState: "chart-source-verified", Value: { Length: > 0 } });
        // This note records that a collaboration dialog is intermittent; it does
        // not request a button choice. Its absence is represented by having no
        // mandatory action for either dialog button in the generated plan.
        var informationalOccasionalDialog = System.Text.RegularExpressions.Regex.IsMatch(note,
            @"\b(?:see everyone|see just mine).+\bdoesn['’]?t\s+always\s+appear\b|\bdoesn['’]?t\s+always\s+appear.+\b(?:see everyone|see just mine)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (informationalOccasionalDialog)
            return !plan.Steps.Any(step => step.Target?.Name is "See just mine" or "See everyone's" &&
                step.Action == "click");
        return false;
    }

    internal static bool RecoverOmittedWorksheetNavigation(ExecutionPlan plan,
        IReadOnlyList<RecordedEvent> recordedEvents, out List<string> recoveredTabs)
    {
        recoveredTabs = [];
        for (var planIndex = 0; planIndex < plan.Steps.Count; planIndex++)
        {
            var step = plan.Steps[planIndex];
            if (step.Action is not ("click" or "context-click") ||
                step.Target is not { Process: "EXCEL", ControlType: "ControlType.DataItem",
                    ClassName: { Length: > 0 } gridClass } ||
                !gridClass.StartsWith("XLGrid", StringComparison.Ordinal) &&
                !gridClass.Equals("XLSpreadsheetCell", StringComparison.Ordinal))
                continue;
            var eventIndex = -1;
            for (var candidate = 0; candidate < recordedEvents.Count; candidate++)
            {
                var recorded = recordedEvents[candidate];
                if (recorded.Kind != step.Action || recorded.Screenshot != step.Screenshot ||
                    recorded.Target?.Process != step.Target.Process ||
                    recorded.Target.ClassName != step.Target.ClassName ||
                    recorded.Target.Name != step.Target.Name) continue;
                if (string.IsNullOrWhiteSpace(step.Screenshot))
                {
                    if (planIndex + 1 >= plan.Steps.Count || candidate + 1 >= recordedEvents.Count)
                        continue;
                    var nextStep = plan.Steps[planIndex + 1];
                    var nextRecorded = recordedEvents[candidate + 1];
                    // A plain Excel cell click may have no screenshot. Tie it to its following
                    // captured action so repeated addresses such as A1 cannot recover a sheet
                    // tab from an unrelated part of the workflow.
                    if (nextRecorded.Kind != nextStep.Action || nextRecorded.Key != nextStep.Key ||
                        nextRecorded.Screenshot != nextStep.Screenshot ||
                        nextRecorded.Target?.Process != nextStep.Target?.Process ||
                        nextRecorded.Target?.ClassName != nextStep.Target?.ClassName ||
                        nextRecorded.Target?.Name != nextStep.Target?.Name)
                        continue;
                }
                eventIndex = candidate;
                break;
            }
            if (eventIndex <= 0 || recordedEvents[eventIndex - 1] is not
                { Kind: "click", Target: { Process: "EXCEL", ControlType: "ControlType.TabItem",
                    Name: { Length: > 0 } tabName } tabTarget } tabEvent ||
                tabTarget.Window != step.Target.Window)
                continue;
            if (IsFilteredWorksheetContinuation(plan.Steps, planIndex, step.Target.Window))
                continue;
            if (planIndex > 0 && plan.Steps[planIndex - 1] is
                { Action: "click", Target: { Process: "EXCEL", ControlType: "ControlType.TabItem" } existing } &&
                existing.Name == tabName && existing.Window == tabTarget.Window)
                continue;
            plan.Steps.Insert(planIndex, new PlanStep(0, "click", tabTarget, null, null,
                tabEvent.AfterState, tabEvent.Screenshot,
                $"Open the recorded worksheet '{tabName}' before continuing with its grid."));
            recoveredTabs.Add(tabName);
            planIndex++;
        }
        if (recoveredTabs.Count == 0) return false;
        var renumbered = plan.Steps.Select((step, index) => step with { Number = index + 1 }).ToList();
        plan.Steps.Clear();
        plan.Steps.AddRange(renumbered);
        return true;
    }

    private static bool IsFilteredWorksheetContinuation(IReadOnlyList<PlanStep> steps, int gridIndex,
        string? workbook)
    {
        if (gridIndex <= 0 || string.IsNullOrWhiteSpace(workbook))
            return false;
        var filterIndex = gridIndex - 1;
        while (filterIndex >= Math.Max(0, gridIndex - 3) &&
            steps[filterIndex] is { Action: "click", Intent: { Length: > 0 } intent,
                Target: { Process: "EXCEL", ControlType: "ControlType.Button",
                    ParentName: { Length: > 0 }, Window: null or "" } } &&
            System.Text.RegularExpressions.Regex.IsMatch(intent,
                @"\b(?:doesn['’]?t|does\s+not)\s+always\s+appear\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            filterIndex--;
        if (filterIndex < 0 || steps[filterIndex] is not
            { Action: "filter-values", Target: { Window: { } filterWorkbook } } ||
            !string.Equals(filterWorkbook, workbook, StringComparison.OrdinalIgnoreCase))
            return false;
        // A collaboration prompt can delay mouse processing: UIA may then report the sheet tab
        // underneath a blocked click even though Excel never activated that sheet. A semantic
        // filter followed by a visible-filtered-row edit proves the operation remains on the
        // filtered worksheet, so do not synthesize navigation from that stale point identity.
        return steps.Skip(gridIndex + 1).Take(4).Any(candidate =>
            candidate is { Action: "type", TargetStrategy: "first-visible-filtered-row",
                Target: { Window: { } candidateWorkbook } } &&
            string.Equals(candidateWorkbook, workbook, StringComparison.OrdinalIgnoreCase));
    }

    private static List<string> ValidatePlanTransitions(ExecutionPlan plan)
    {
        var warnings = new List<string>();
        foreach (var filter in plan.Steps.Where(step => step.Action == "click" &&
            step.Target is { AutomationId: "Dropdown", ControlType: "ControlType.MenuItem" } &&
            string.IsNullOrWhiteSpace(step.Target.ParentName)))
            warnings.Add($"Step {filter.Number}: the column parent for filter '{filter.Target!.Name}' " +
                "was not recorded. This plan cannot replay automatically; record the filter click again.");
        if (ReplacementPlanProblem(plan) is { } replacementProblem) warnings.Add(replacementProblem);
        foreach (var incomplete in IncompleteReplacements(plan))
            warnings.Add($"Step {incomplete}: text or replacement intent was recorded in Find and Replace, but the recording does not contain a Replace All click before moving on. Replay will not perform this replacement. Record that button click again.");
        foreach (var close in plan.Steps.Where(step => step.Action == "close-window" &&
            step.Target?.Window == "Find and Replace" && step.Explanation?.StartsWith("Close Find and Replace before", StringComparison.Ordinal) == true))
            warnings.Add($"Step {close.Number}: the recorder did not save the Find and Replace Close click. The plan added a Close step when it saw later worksheet work. Check this in the PDF; record the Close click again if you need the plan to match your actions exactly.");
        for (var i = 0; i + 1 < plan.Steps.Count; i++)
        {
            var first = plan.Steps[i];
            var next = plan.Steps[i + 1];
            if (first.Action == "type" && first.Target is
                { Process: "SearchHost", AutomationId: "SearchTextBox" } &&
                next.Action == "click" && next.Target is
                { Process: "explorer", ClassName: "UIProperty", Name: "Name" })
                warnings.Add($"Step {first.Number}: Windows Search text is followed by a File Explorer property click, not a recorded Search result or launch action. Replay will stop; record the result selection or Enter key.");
            if (first.Action == "context-click" && first.Target is
                { Process: "EXCEL", ClassName: "XLGridRowHeader" } &&
                next.Action == "click" && next.Target is
                { Process: "EXCEL", ClassName: "XLSpreadsheetCell" })
                warnings.Add($"Step {next.Number}: the click after opening row {first.Target.Name}'s menu was recorded as worksheet cell '{next.Target.Name}' instead of a menu command. Replay will stop before editing the sheet.");
            if (first.Action == "manual-click" && first.Target?.ControlType == "ControlType.ToolTip")
                warnings.Add($"Step {first.Number}: a pop-up hint covered the control you clicked. Replay will pause so you can make that click. Record it again if you need it to run automatically.");
            else if (first.Action == "manual-click")
                warnings.Add(ManualActionWarning(first));
            else if (first.ExpectedState == "unverified-click")
                warnings.Add($"Step {first.Number}: the saved control '{first.Target?.Name}' may be wrong. The app could not confirm it was under your pointer. Compare it with the PDF screenshot before replay.");
            else if (first.Action is "click" or "context-click" && first.Target?.ControlType == "ControlType.ToolTip")
                warnings.Add($"Step {first.Number}: the app saved a pop-up hint named '{first.Target.Name}' as the clicked control. That hint may have covered the control you intended. Record this action again before relying on replay.");
            if (first.Action != "click" || first.Target?.ControlType != "ControlType.DataItem" ||
                next.Action is not ("type" or "key") || next.Target?.ControlType != "ControlType.DataItem" ||
                first.Target == next.Target || first.Target.Process != next.Target.Process ||
                first.Target.Window != next.Target.Window) continue;
            warnings.Add($"Step {first.Number} clicks '{first.Target.Name}', then step {next.Number} sends keyboard input to '{next.Target.Name}'. Check that both targets are intended.");
        }
        if (plan.Steps.LastOrDefault() is { } last)
        {
            if (last.Action == "manual-click" && last.Target?.ControlType == "ControlType.ToolTip")
                warnings.Add($"Step {last.Number}: a pop-up hint covered the control you clicked. Replay will pause so you can make that click. Record it again if you need it to run automatically.");
            else if (last.Action == "manual-click")
                warnings.Add(ManualActionWarning(last));
            else if (last.ExpectedState == "unverified-click")
                warnings.Add($"Step {last.Number}: the saved control '{last.Target?.Name}' may be wrong. The app could not confirm it was under your pointer. Compare it with the PDF screenshot before replay.");
            else if (last.Action is "click" or "context-click" && last.Target?.ControlType == "ControlType.ToolTip")
                warnings.Add($"Step {last.Number}: the app saved a pop-up hint named '{last.Target.Name}' as the clicked control. Record this action again before relying on replay.");
        }
        return warnings;
    }

    private static string ManualActionWarning(PlanStep step) =>
        step.Target is { Process: "EXCEL", ControlType: "ControlType.DataItem", Name: { Length: > 0 } name }
            ? $"Step {step.Number}: the worksheet target '{name}' is known, but its final selection or fill outcome was not saved. This plan cannot replay automatically; supply explicit operation intent or re-record this gesture."
            : $"Step {step.Number}: the app recorded a click but could not identify its control. This plan cannot replay automatically; record that action again with this build.";

    private static bool IsBlockingPlanWarning(string warning) =>
        warning.Contains("cannot replay automatically", StringComparison.OrdinalIgnoreCase) ||
        warning.Contains("Replay will stop", StringComparison.OrdinalIgnoreCase) ||
        warning.Contains("does not contain a Replace All click", StringComparison.OrdinalIgnoreCase) ||
        warning.Contains("replacement", StringComparison.OrdinalIgnoreCase) &&
            warning.Contains("stopped", StringComparison.OrdinalIgnoreCase);

    private static string? ReplacementPlanProblem(ExecutionPlan plan)
    {
        string? findConfirmed = null;
        string? replaceConfirmed = null;
        foreach (var step in plan.Steps)
        {
            if (step.Action == "replace-all" && step.Target is not
                { Window: "Find and Replace", Name: "Replace All" })
                return $"Step {step.Number} is marked Replace All but targets '{step.Target?.Name}' in '{step.Target?.Window}'. Replay stopped before changing the worksheet because the plan does not match the recorded dialog action.";
            if (step.OriginIntent is not null && step.Action == "replace-all")
            {
                if (string.IsNullOrWhiteSpace(step.Value) || step.ExpectedState is null ||
                    step.Intent != step.OriginIntent ||
                    !step.OriginIntent.Contains(step.Value, StringComparison.Ordinal) ||
                    !step.OriginIntent.Contains(step.ExpectedState, StringComparison.Ordinal) ||
                    step.WhenUser is not null && !System.Text.RegularExpressions.Regex.IsMatch(
                        step.WhenUser, @"^(?:[\w.-]+\\)?[\w.-]+$"))
                    return $"Step {step.Number}: the intent-generated replacement lacks its source intent or replacement values. Replay stopped.";
                continue;
            }
            if (step.Target?.Window != "Find and Replace") continue;
            if (step.Action == "type" && step.Target.AutomationId == "18")
                findConfirmed = step.ExpectedState?.StartsWith("field-value:", StringComparison.Ordinal) == true
                    ? step.ExpectedState["field-value:".Length..] : null;
            if (step.Action == "type" && step.Target.AutomationId == "21")
                replaceConfirmed = step.ExpectedState?.StartsWith("field-value:", StringComparison.Ordinal) == true
                    ? step.ExpectedState["field-value:".Length..] : null;
            if (step.Action == "click" && step.Target.Name == "Replace All" &&
                (findConfirmed is null || replaceConfirmed is null))
                return $"Step {step.Number} presses Replace All before both text boxes have confirmed final values. " +
                    "The recording may have missed a field edit or confused another click with Replace All. " +
                    "Replay stopped before changing the worksheet. Record the replacement again and check the plan report.";
            if (step.Action == "replace-all" && step.Intent is null &&
                (findConfirmed is null || replaceConfirmed is null ||
                 step.Value != findConfirmed || step.ExpectedState != replaceConfirmed))
                return $"Step {step.Number} contains replacement text that the recorder did not confirm in both boxes. " +
                    "Replay stopped before changing the worksheet. Record the replacement again and check the plan report.";
        }
        return null;
    }

    private static List<int> IncompleteReplacements(ExecutionPlan plan)
    {
        var incomplete = new List<int>();
        int? firstField = null;
        var typed = false;
        string? pendingIntent = null;
        string? pendingProcess = null;
        var replaced = false;
        foreach (var step in plan.Steps)
        {
            var inDialog = step.Target?.Window == "Find and Replace";
            if (inDialog && (step.Action == "type" ||
                step.Action != "replace-all" && step.Target?.Name != "Replace All" &&
                System.Text.RegularExpressions.Regex.IsMatch(step.Intent ?? "", @"\breplace\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase)))
            {
                firstField ??= step.Number;
                typed |= step.Action == "type";
                pendingIntent ??= step.Intent;
                pendingProcess ??= step.Target?.Process;
                continue;
            }
            if (step.Action == "replace-all" ||
                step.Action == "click" && inDialog && step.Target?.Name == "Replace All")
            {
                replaced = true;
                continue;
            }
            if (!inDialog && firstField is { } number)
            {
                var repeatedIntentFulfilled = !typed && !string.IsNullOrWhiteSpace(pendingIntent) &&
                    plan.Steps.Any(later => later.Number > step.Number &&
                        later.Action == "replace-all" && later.Target?.Process == pendingProcess &&
                        later.Target?.Window == "Find and Replace" &&
                        (later.Intent == pendingIntent || later.OriginIntent == pendingIntent));
                if (!replaced && !repeatedIntentFulfilled) incomplete.Add(number);
                firstField = null;
                typed = false;
                pendingIntent = null;
                pendingProcess = null;
                replaced = false;
            }
        }
        if (firstField is { } last && !replaced) incomplete.Add(last);
        return incomplete;
    }

    internal static List<RecordedEvent> RecoverRecordedDialogCommands(List<RecordedEvent> source)
    {
        for (var index = 0; index + 1 < source.Count; index++)
        {
            var click = source[index];
            var next = source[index + 1];
            if (click.Kind != "click" || click.Target is not { Window: "Find and Replace" } target ||
                target.ControlType is not ("ControlType.Window" or "ControlType.Pane" or
                    "ControlType.Custom" or "ControlType.Tab" or "ControlType.Image") ||
                click.BeforeState?.Contains("\"AutomationId\":\"18\"") != true ||
                click.BeforeState.Contains("\"AutomationId\":\"21\"") != true ||
                next.Target?.Window is null || next.Target.Window == "Find and Replace") continue;
            // A new modal immediately after an unnamed command is observable evidence that
            // the dialog command ran. Preserve its accessible identity for coordinate-free replay.
            source[index] = click with { Target = target with { Name = "Replace All",
                ControlType = "ControlType.Button" }, AfterState = "dialog-command-result" };
        }
        return source;
    }

    private static Task WriteValidationReportAsync(string directory, ExecutionPlan plan, CancellationToken token)
    {
        var warnings = ValidatePlanTransitions(plan);
        var lines = new List<string>
        {
            "# Execution plan validation",
            "",
            $"Checked: {DateTimeOffset.Now:O}",
            $"Planned steps: {plan.Steps.Count}",
            $"Status: {(warnings.Count == 0 ? "No structural warnings" : "Review required")}",
            "",
            "The app checked whether each click has a saved control and whether nearby steps point to different controls. It cannot tell what a missing click was from a screenshot alone.",
            "",
            "## Findings",
            ""
        };
        if (warnings.Count == 0) lines.Add("- Every click has a saved control, and no nearby steps point to conflicting cells.");
        else lines.AddRange(warnings.Select(warning => "- " + warning));
        lines.AddRange(["", "## Action plan", "",
            "1. Open manual_steps.pdf at each step listed above and compare the screenshot with the action you intended.",
            "2. If the app did not save the clicked control, replay will ask you to make that click. Record that action again if it must run automatically.",
            "3. If the plan points to the wrong control, record that action again. The app keeps the saved plan for review and will not silently guess a replacement.",
            "4. After replay, open the Replay Feedback Report to see what happened and what to do next.", ""]);
        return File.WriteAllTextAsync(Path.Combine(directory, "validation_report.md"),
            string.Join("\n", lines), token);
    }

    private static string ReplayRecoveryActions(string reason)
    {
        if (reason.Contains("not found", StringComparison.OrdinalIgnoreCase) ||
            reason.Contains("no UI Automation target", StringComparison.OrdinalIgnoreCase))
            return "1. Open manual_steps.pdf and compare the failed step's screenshot with its target in execution_plan.json.\n" +
                "2. If the target is correct, expose it in the target program and retry the step. If the saved target is absent or wrong, record that action again; rebuilding unchanged events cannot reconstruct it.\n" +
                "3. Check validation_report.md for other steps requiring review.\n";
        if (reason.Contains("nonenabled", StringComparison.OrdinalIgnoreCase) ||
            reason.Contains("not enabled", StringComparison.OrdinalIgnoreCase))
            return "1. Check whether a dialog, menu, or pending operation is blocking the control.\n" +
                "2. Bring the intended control into its enabled state, then retry the step. Compare the saved target with manual_steps.pdf if it remains disabled.\n";
        if (reason.Contains("scroll", StringComparison.OrdinalIgnoreCase))
            return "1. Compare the recorded target and screenshot in manual_steps.pdf with the current view.\n" +
                "2. Use the target program's scroll controls to expose the intended control, then retry. Re-record if the saved target identifies a different control.\n";
        return "1. Review the failed step and recent actions above alongside manual_steps.pdf and execution_plan.json.\n" +
            "2. Resolve the target program state and retry the step. Re-record an action whose saved target is missing or incorrect.\n";
    }

    private static string HumanReplayLine(string line)
    {
        if (line.Contains("without a verifiable click target", StringComparison.OrdinalIgnoreCase))
            return System.Text.RegularExpressions.Regex.Replace(line, "was captured without a verifiable click target",
                "was recorded, but the app could not tell which control was clicked", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (line.Contains("UI Automation", StringComparison.OrdinalIgnoreCase))
            return line.Replace("UI Automation", "the target app's controls", StringComparison.OrdinalIgnoreCase);
        if (line.Contains("nonenabled element", StringComparison.OrdinalIgnoreCase))
            return "The target control was present but could not be used, possibly because a menu or dialog was blocking it.";
        return line;
    }

    private bool ConfirmPlanWarnings(IReadOnlyList<string> warnings)
    {
        using var dialog = new Form { Text = "Review execution plan", ClientSize = new System.Drawing.Size(760, 340),
            MinimumSize = new System.Drawing.Size(600, 300), FormBorderStyle = FormBorderStyle.Sizable,
            StartPosition = FormStartPosition.CenterParent, MaximizeBox = true };
        var details = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill, BorderStyle = BorderStyle.None,
            Text = "The saved plan has an unusual target change. Replay will perform every step as recorded.\r\n\r\n" +
                string.Join("\r\n", warnings) };
        var proceed = new Button { Text = "Proceed as recorded", AutoSize = true,
            MinimumSize = new System.Drawing.Size(170, 36), DialogResult = DialogResult.OK };
        var stop = new Button { Text = "Stop replay", AutoSize = true,
            MinimumSize = new System.Drawing.Size(125, 36), DialogResult = DialogResult.Cancel };
        var openPdf = new Button { Text = "Open recording PDF", AutoSize = true,
            MinimumSize = new System.Drawing.Size(170, 36) };
        openPdf.Click += (_, _) => OpenPdf();
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), RowCount = 2 };
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill };
        buttons.Controls.AddRange([proceed, stop, openPdf]);
        layout.Controls.Add(details, 0, 0);
        layout.Controls.Add(buttons, 0, 1);
        dialog.Controls.Add(layout);
        dialog.AcceptButton = proceed;
        dialog.CancelButton = stop;
        return dialog.ShowDialog(this) == DialogResult.OK;
    }

    private bool AskToHandleUnexpectedDialog(string title)
    {
        using var dialog = new Form
        {
            Text = title.StartsWith("Refresh completion", StringComparison.Ordinal)
                ? "Replay paused for refresh" :
                title.StartsWith("Replay needs help:", StringComparison.Ordinal)
                ? "Replay needs help" :
                title.StartsWith("After closing this prompt", StringComparison.Ordinal)
                ? "Replay waiting for context menu" : "Replay paused for an unexpected dialog",
            ClientSize = new System.Drawing.Size(760, 340),
            MinimumSize = new System.Drawing.Size(600, 300),
            FormBorderStyle = FormBorderStyle.Sizable, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = true, MinimizeBox = false
        };
        var message = new TextBox
        {
            Text = title.StartsWith("Refresh completion", StringComparison.Ordinal)
                ? title
                : title.StartsWith("Replay needs help:", StringComparison.Ordinal)
                ? title
                : title.StartsWith("After closing this prompt", StringComparison.Ordinal)
                ? title
                : title.StartsWith("Open the context menu", StringComparison.Ordinal)
                ? title
                : title.StartsWith("Excel could not finish", StringComparison.Ordinal)
                ? title + "\nHandle it in Excel, then return here."
                : $"An unrecorded dialog appeared: {title}\nSwitch to that app, handle the dialog, then return here.",
            Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill, BorderStyle = BorderStyle.None,
            TabStop = false
        };
        var resume = new Button { Text = title.StartsWith("After closing this prompt", StringComparison.Ordinal)
            ? "Begin waiting" : title.StartsWith("Replay needs help:", StringComparison.Ordinal)
            ? "I've handled it" : "I've handled it — continue",
            AutoSize = true, MinimumSize = new System.Drawing.Size(170, 36), DialogResult = DialogResult.OK };
        var stopReplay = new Button { Text = "Stop replay",
            AutoSize = true, MinimumSize = new System.Drawing.Size(125, 36), DialogResult = DialogResult.Cancel };
        var openPdf = new Button { Text = "Open recording PDF",
            AutoSize = true, MinimumSize = new System.Drawing.Size(170, 36) };
        openPdf.Click += (_, _) => OpenPdf();
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16),
            ColumnCount = 1, RowCount = 2 };
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
        buttons.Controls.AddRange([resume, stopReplay, openPdf]);
        layout.Controls.Add(message, 0, 0);
        layout.Controls.Add(buttons, 0, 1);
        dialog.Controls.Add(layout);
        dialog.AcceptButton = resume;
        dialog.CancelButton = stopReplay;
        return dialog.ShowDialog(this) == DialogResult.OK;
    }

    private static bool CanConfirmFilterSelection(int selectedIndex, string[] names) =>
        selectedIndex is 1 or 2 || selectedIndex == 0 && names.Length > 0 &&
        !names.Contains("(Select All)", StringComparer.Ordinal);

    private static string[] EffectiveFilterValues(string[] recordedValues, string[] priorValues) =>
        recordedValues.Length > 0 ? recordedValues : priorValues;

    private static string FilterLocation(PlanStep step)
    {
        var address = step.Target?.ParentName;
        var match = System.Text.RegularExpressions.Regex.Match(address ?? "", "^([A-Za-z]+)([0-9]+)$");
        return match.Success
            ? $"worksheet column {match.Groups[1].Value.ToUpperInvariant()} (header cell {address!.ToUpperInvariant()})"
            : !string.IsNullOrWhiteSpace(address) ? $"worksheet column {address}" : "the recorded worksheet column";
    }

    private (string Mode, string[] Values)? AskFilterSelection(PlanStep step, int ordinal, int total,
        string[] priorValues, int priorOrdinal)
    {
        using var dialog = new Form { Text = $"Confirm filter {ordinal} of {total}", ClientSize = new Size(720, 420),
            StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false, MinimizeBox = false };
        var recordedValues = ExcelFilterPlan.Values(step);
        var values = EffectiveFilterValues(recordedValues, priorValues);
        var message = new Label { Dock = DockStyle.Top, Height = 135, Padding = new Padding(12),
            Text = $"Filter {ordinal} of {total}: {FilterLocation(step)} (replay step {step.Number}).\r\n" +
                "The recorder could not verify all checkbox states in this recording. Replaying toggles could undo the filter.\r\n\r\n" +
                (recordedValues.Length > 0
                    ? $"Recorded values for this filter: {string.Join(", ", recordedValues)}\r\n"
                    : priorValues.Length > 0
                        ? $"No names were captured for this filter. Reusing the last verified values from filter {priorOrdinal} on this same column: {string.Join(", ", priorValues)}\r\n"
                        : "Recorded values: none were captured\r\n") +
                "Choose the intended final result. Add exact values, one per line, only when the result depends on named values. Your choice is saved for future replays." };
        var choices = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        choices.Items.AddRange(["Keep only the listed values", "Keep everyone except the listed values",
            "Show all values (clear this filter)"]);
        var choiceHelp = new Label { Dock = DockStyle.Top, Height = 55, Padding = new Padding(12, 8, 12, 4) };
        var input = new TextBox { Dock = DockStyle.Top, Multiline = true, Height = 110,
            ScrollBars = ScrollBars.Vertical, Text = string.Join(Environment.NewLine, values) };
        var confirm = new Button { Text = "Confirm filter", AutoSize = true, Enabled = false, DialogResult = DialogResult.OK };
        string[] Names() => input.Lines.Select(value => value.Trim()).Where(value => value.Length > 0)
            .Distinct(StringComparer.Ordinal).ToArray();
        void ValidateChoice()
        {
            var names = Names();
            confirm.Enabled = CanConfirmFilterSelection(choices.SelectedIndex, names);
            choiceHelp.Text = choices.SelectedIndex switch
            {
                0 => names.Length == 0
                    ? "Enter at least one value. Excel cannot apply a filter that keeps no values."
                    : "Only the values entered below will remain visible.",
                1 => names.Length == 0
                    ? "No exceptions are listed, so this will show all values and clear this column's filter."
                    : "All values remain visible except the values entered below.",
                2 => "All values will be selected and this column's filter will be cleared. No text is required.",
                _ => "Select the final filter result."
            };
        }
        choices.SelectedIndexChanged += (_, _) => { input.Enabled = choices.SelectedIndex != 2; ValidateChoice(); };
        input.TextChanged += (_, _) => ValidateChoice();
        var cancel = new Button { Text = "Cancel replay", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48 };
        buttons.Controls.AddRange([confirm, cancel]);
        dialog.Controls.Add(input); dialog.Controls.Add(choiceHelp); dialog.Controls.Add(choices); dialog.Controls.Add(message); dialog.Controls.Add(buttons);
        dialog.CancelButton = cancel;
        if (dialog.ShowDialog(this) != DialogResult.OK) return null;
        return (choices.SelectedIndex == 2 ? "all" : choices.SelectedIndex == 0 ? "only" : "exclude",
            choices.SelectedIndex == 2 ? [] : Names());
    }

    private string? AskSortDirection(string column)
    {
        using var dialog = new Form
        {
            Text = "Choose recorded sort direction", Width = 410, Height = 175,
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false
        };
        var message = new Label
        { Text = $"The saved plan does not say how {column} should be sorted. Choose the intended order:",
          Left = 16, Top = 16, Width = 370, Height = 48 };
        var ascending = new Button { Text = "Ascending", Left = 16, Top = 78, Width = 105, DialogResult = DialogResult.Yes };
        var descending = new Button { Text = "Descending", Left = 130, Top = 78, Width = 105, DialogResult = DialogResult.No };
        var cancel = new Button { Text = "Cancel replay", Left = 244, Top = 78, Width = 125, DialogResult = DialogResult.Cancel };
        dialog.Controls.AddRange([message, ascending, descending, cancel]);
        dialog.CancelButton = cancel;
        return dialog.ShowDialog(this) switch
        { DialogResult.Yes => "ascending", DialogResult.No => "descending", _ => null };
    }

    private int? AskRefreshDuration(int initialSeconds)
    {
        using var dialog = new Form { Text = "Refresh wait before replay", Width = 510, Height = 255,
            StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false, MinimizeBox = false };
        var explanation = new Label { Left = 18, Top = 15, Width = 465, Height = 65,
            Text = "This replay includes a Refresh button step (possibly conditional).\nHow many seconds do you expect refreshing to take?\nEnter a positive whole number; replay adds 2 seconds." };
        var input = new TextBox { Left = 18, Top = 85, Width = 180,
            Text = initialSeconds.ToString(System.Globalization.CultureInfo.InvariantCulture) };
        var result = new Label { Left = 18, Top = 120, Width = 465, Height = 35 };
        var start = new Button { Text = "Start replay", Left = 200, Top = 165, Width = 130,
            DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel replay", Left = 340, Top = 165, Width = 130,
            DialogResult = DialogResult.Cancel };
        void ValidateInput()
        {
            start.Enabled = RefreshTiming.TryExpectedSeconds(input.Text, out var seconds);
            result.Text = start.Enabled
                ? $"Completion timeout: {seconds + 2} seconds. Replay can finish waiting earlier."
                : $"Enter whole seconds from 1 to {int.MaxValue - 2}; no decimals or signs.";
        }
        input.KeyPress += (_, e) =>
            e.Handled = !char.IsControl(e.KeyChar) && (e.KeyChar < '0' || e.KeyChar > '9');
        input.TextChanged += (_, _) => ValidateInput();
        dialog.Controls.AddRange([explanation, input, result, start, cancel]);
        dialog.AcceptButton = start;
        dialog.CancelButton = cancel;
        dialog.Shown += (_, _) => { input.Focus(); input.SelectAll(); };
        ValidateInput();
        if (dialog.ShowDialog(this) != DialogResult.OK ||
            !RefreshTiming.TryExpectedSeconds(input.Text, out var value)) return null;
        return value;
    }

    private void OpenPdf()
    {
        var path = SelectedDirectory() is { } directory ? Path.Combine(directory, "manual_steps.pdf") : null;
        if (path is null || !File.Exists(path)) { Log("No PDF found."); return; }
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
    }

    private async Task ChatAsync()
    {
        var command = input.Text.Trim(); input.Clear();
        if (command.Length == 0) return;
        if (command.Equals("clear", StringComparison.OrdinalIgnoreCase) ||
            command.Equals("cls", StringComparison.OrdinalIgnoreCase)) { chat.Clear(); return; }
        Log("You: " + command);
        try
        {
            var response = await new Foundry(EnvPath()).AskAsync(
                "Classify this desktop app command. Reply with exactly one word: record, stop, replay, pdf, or unknown. Command: " + command,
                CancellationToken.None);
            switch (response.Trim().ToLowerInvariant())
            {
                case "record": StartOrResumeRecording(); break;
                case "stop": await StopRecordingAsync(); break;
                case "replay": await ReplayAsync(); break;
                case "pdf": OpenPdf(); break;
                default: Log("Choose Record, Stop, Replay, or Open PDF."); break;
            }
        }
        catch (Exception ex) { Log("Command interpretation failed: " + ex.Message); }
    }

    internal static string EnvPath()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        for (var depth = 0; directory is not null && depth < 7; depth++, directory = directory.Parent)
        {
            var candidate = Path.Combine(directory.FullName, ".env");
            if (File.Exists(candidate)) return candidate;
        }
        throw new FileNotFoundException("Place .env beside the app or in the project root.");
    }
}
