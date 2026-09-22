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
    private readonly Button elevatedRecord = new ReadableButton() { Text = "Restart as administrator to record elevated apps", AutoSize = true };
    private readonly Button stop = new ReadableButton() { Text = "Stop Recording...", AutoSize = true, Enabled = false };
    private readonly Button cancelRecording = new ReadableButton() { Text = "Cancel recording", AutoSize = true, Enabled = false };
    private readonly Button finish = new ReadableButton() { Text = "Rebuild selected files", AutoSize = true };
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
        Text = "Desktop Steps"; Width = 850; Height = 600;
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
        buttons.Controls.AddRange([record, elevatedRecord, stop, cancelRecording, finish, replay, stopReplay, pdf,
            validationReport, feedbackReport, rulesGroup, clear, theme, authGroup]);
        buttonTips.SetToolTip(record, "Start capturing desktop clicks, typing, and scrolling.");
        buttonTips.SetToolTip(stop, "Stop capturing and create the execution plan, summary, and PDF.");
        buttonTips.SetToolTip(cancelRecording, "Stop capturing and discard this unfinished recording without creating files.");
        buttonTips.SetToolTip(finish, "Rebuild the selected recording's plan, summary, and PDF from saved events with help of Foundry.");
        buttonTips.SetToolTip(replay, "Run the selected recording's execution plan on your desktop.");
        buttonTips.SetToolTip(stopReplay, "Stop the replay currently running on your desktop.");
        buttonTips.SetToolTip(pdf, "Open the selected recording's manual instructions PDF.");
        buttonTips.SetToolTip(validationReport, "Open the selected recording's plan check, including any steps that need review before replay.");
        buttonTips.SetToolTip(feedbackReport, "Open what happened during the selected recording's last replay and the suggested next steps.");
        buttonTips.SetToolTip(recordingRules, "Add or remove apps that the recorder should ignore, such as notepad.exe. Rules apply to future recordings.");
        buttonTips.SetToolTip(clear, "Clear messages shown in this window. Saved recordings remain available.");
        buttonTips.SetToolTip(theme, "Switch between dark and light mode.");
        buttonTips.SetToolTip(send, "Send the typed command to Desktop Steps.");
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
        record.Click += (_, _) => StartRecording();
        elevatedRecord.Click += (_, _) => RestartElevated();
        stop.Click += async (_, _) => await StopRecordingAsync();
        cancelRecording.Click += (_, _) => CancelRecording();
        finish.Click += async (_, _) => await FinishLatestAsync();
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
        recorder.IntentRequested += priorWindow =>
        {
            if (intentDialogQueued) return;
            intentDialogQueued = true;
            BeginInvoke(new Action(() => ShowIntentNote(priorWindow)));
        };
        recorder.IntentCancelRequested += () => BeginInvoke(new Action(() =>
        {
            if (activeIntentDialog is { IsDisposed: false } dialog)
            {
                dialog.DialogResult = DialogResult.Cancel;
                dialog.Close();
            }
        }));
        FormClosing += (_, _) => { work?.Cancel(); recorder.Dispose(); };
        LoadRecordingRules();
        RefreshRecordings();
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
            using var dialog = new Form { Text = "Add recording intent",
                Width = 570, Height = 265,
                FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false, MinimizeBox = false, TopMost = true };
            activeIntentDialog = dialog;
            var label = new Label { Text = "Describe what the last recorded action should achieve on future runs. Do not enter passwords or PINs.",
                Left = 16, Top = 14, Width = 525, Height = 42 };
            var note = new TextBox { Multiline = true, Left = 16, Top = 62, Width = 525, Height = 90,
                PlaceholderText = "Example: Use the date of next Tuesday in the new name." };
            var add = new Button { Text = "Add intent", Left = 16, Top = 170, Width = 110, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Left = 136, Top = 170, Width = 90, DialogResult = DialogResult.Cancel };
            dialog.Controls.AddRange([label, note, add, cancel]);
            dialog.AcceptButton = add; dialog.CancelButton = cancel;
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
                    ? "Your explanation was saved with that recorded action."
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
        { Title = "Choose a Desktop Steps execution_plan.json file",
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
                "DesktopSteps", "recording_rules.json");
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

    private void StartRecording()
    {
        try
        {
            activeDirectory = Path.Combine(sessions, DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            recorder.SetIgnoredProcesses(ignoredProcesses);
            recorder.Start(activeDirectory);
            File.WriteAllText(Path.Combine(activeDirectory, "recording_rules.json"),
                System.Text.Json.JsonSerializer.Serialize(recordingRulesText.ToArray(), JsonFile.Options));
            record.Enabled = false; stop.Enabled = true; cancelRecording.Enabled = true; finish.Enabled = false;
            recordingRules.Enabled = false;
            Text = "Desktop Steps - Recording";
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
        recorder.Stop(); record.Enabled = true; stop.Enabled = false; cancelRecording.Enabled = false; finish.Enabled = true;
        recordingRules.Enabled = true;
        Text = "Desktop Steps";
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
            Log("Desktop Steps is already running as administrator. Start a new recording and include the DebugDiag controls after approving UAC.");
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
        Log("Recording warning: no actions in the launched application were captured after Windows Search. If you used an elevated app, this plan will only replay its launch. Restart Desktop Steps as administrator, then record the application controls again.");
    }

    private void CancelRecording()
    {
        if (!recorder.IsRecording) { Log("No recording is in progress."); return; }
        recorder.Stop();
        record.Enabled = true; stop.Enabled = false; cancelRecording.Enabled = false; finish.Enabled = true;
        recordingRules.Enabled = true;
        Text = "Desktop Steps";
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

    private async Task FinishLatestAsync()
    {
        if (recorder.IsRecording) { await StopRecordingAsync(); return; }
        if (work is not null) { Log("Files are still being created. Please wait for completion."); return; }
        var directory = SelectedDirectory();
        if (directory is null) { Log("No recording found."); return; }
        var path = Path.Combine(directory, "recorded_events.json");
        if (!File.Exists(path)) { Log("The latest recording has no captured events file."); return; }
        var events = await JsonFile.LoadAsync<List<RecordedEvent>>(path);
        if (events is null || events.Count == 0) { Log("The captured events file is empty."); return; }
        await FinishAsync(directory, events);
    }

    private async Task FinishAsync(string directory, List<RecordedEvent> events)
    {
        if (work is not null) { Log("An operation is already running."); return; }
        work = new CancellationTokenSource();
        try
        {
            var removedSyntheticClicks = events.RemoveAll(item => item.Kind == "click" &&
                item.AfterState == "dialog-command" && item.ClickX is null && item.ClickY is null);
            if (removedSyntheticClicks > 0)
                Log($"Removed {removedSyntheticClicks} accessibility events that had no recorded user click.");
            events = ExpandDialogFieldSnapshots(events);
            await JsonFile.SaveAsync(Path.Combine(directory, "recorded_events.json"), events, work.Token);
            Log($"Captured {events.Count} actions. Asking Foundry to infer intent...");
            var rulesPath = Path.Combine(directory, "recording_rules.json");
            var rules = File.Exists(rulesPath)
                ? System.Text.Json.JsonSerializer.Deserialize<string[]>(await File.ReadAllTextAsync(rulesPath, work.Token)) ?? []
                : [];
            var plan = RefreshTiming.AddObservedWait(
                await new Foundry(EnvPath()).PlanAsync(events, work.Token, rules), events);
            await JsonFile.SaveAsync(Path.Combine(directory, "execution_plan.json"), plan, work.Token);
            await File.WriteAllTextAsync(Path.Combine(directory, "summary.txt"), plan.Summary, work.Token);
            await ManualPdf.CreateAsync(plan, directory, Path.Combine(directory, "manual_steps.pdf"), work.Token);
            await WriteValidationReportAsync(directory, plan, work.Token);
            Log($"Created execution_plan.json, summary.txt, and manual_steps.pdf in {directory}");
            if (ReplacementPlanProblem(plan) is { } replacementProblem)
                Log("Plan needs review: " + replacementProblem);
        }
        catch (Exception ex) { Log("Planning failed; captured events were preserved: " + ex.Message); }
        finally { work.Dispose(); work = null; RefreshRecordings(directory); }
    }

    private static List<RecordedEvent> ExpandDialogFieldSnapshots(List<RecordedEvent> source)
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
        var path = Path.Combine(directory, "execution_plan.json");
        if (!File.Exists(path)) { Log("This recording has no execution plan."); return; }
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
            for (var i = 0; i < plan.Steps.Count; i++)
            {
                var step = plan.Steps[i];
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
            if (repaired) await JsonFile.SaveAsync(path, plan, work.Token);
            await WriteValidationReportAsync(directory, plan, work.Token);
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
                    (item.Target?.Name == "Replace All" || item.AfterState == "replace-all-result")))
            {
                var message = "This saved plan contains Replace All, but the recorded clicks do not. Replay stopped because the plan would press a button the recorder never captured. Please record the replacement again.";
                Log(message);
                MessageBox.Show(this, message, "Plan does not match recording",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (IncompleteReplacements(plan) is { Count: > 0 } missingReplacements)
            {
                var message = $"The recording entered Find and Replace text near step {missingReplacements[0]}, but no Replace All click was saved before it moved on. Replay has stopped so it cannot silently skip your replacement. Record that part again; the app will not invent the missing click.";
                Log(message);
                MessageBox.Show(this, message, "Recording missed Replace All",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var planWarnings = ValidatePlanTransitions(plan);
            if (planWarnings.Count > 0)
            {
                foreach (var warning in planWarnings) Log("Plan warning: " + warning);
                if (!ConfirmPlanWarnings(planWarnings))
                {
                    Log("Replay cancelled during plan review.");
                    return;
                }
            }
            var identity = await WhoAmIAsync(work.Token);
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

    private static List<string> ValidatePlanTransitions(ExecutionPlan plan)
    {
        var warnings = new List<string>();
        if (ReplacementPlanProblem(plan) is { } replacementProblem) warnings.Add(replacementProblem);
        foreach (var incomplete in IncompleteReplacements(plan))
            warnings.Add($"Step {incomplete}: text was entered in Find and Replace, but the recording does not contain a Replace All click before moving on. Replay will not perform this replacement. Record that button click again.");
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
                warnings.Add($"Step {first.Number}: the app recorded a click but could not tell which control you clicked. Replay will pause so you can make this click yourself. Open the PDF to see what was on screen.");
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
                warnings.Add($"Step {last.Number}: the app recorded a click but could not tell which control you clicked. Replay will pause so you can make this click yourself. Open the PDF to see what was on screen.");
            else if (last.ExpectedState == "unverified-click")
                warnings.Add($"Step {last.Number}: the saved control '{last.Target?.Name}' may be wrong. The app could not confirm it was under your pointer. Compare it with the PDF screenshot before replay.");
            else if (last.Action is "click" or "context-click" && last.Target?.ControlType == "ControlType.ToolTip")
                warnings.Add($"Step {last.Number}: the app saved a pop-up hint named '{last.Target.Name}' as the clicked control. Record this action again before relying on replay.");
        }
        return warnings;
    }

    private static string? ReplacementPlanProblem(ExecutionPlan plan)
    {
        string? findConfirmed = null;
        string? replaceConfirmed = null;
        foreach (var step in plan.Steps)
        {
            if (step.Action == "replace-all" && step.Target is not
                { Window: "Find and Replace", Name: "Replace All" })
                return $"Step {step.Number} is marked Replace All but targets '{step.Target?.Name}' in '{step.Target?.Window}'. Replay stopped before changing the worksheet because the plan does not match the recorded dialog action.";
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
        var replaced = false;
        foreach (var step in plan.Steps)
        {
            var inDialog = step.Target?.Window == "Find and Replace";
            if (step.Action == "type" && inDialog)
            {
                firstField ??= step.Number;
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
                if (!replaced) incomplete.Add(number);
                firstField = null;
                replaced = false;
            }
        }
        if (firstField is { } last && !replaced) incomplete.Add(last);
        return incomplete;
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
                case "record": StartRecording(); break;
                case "stop": await StopRecordingAsync(); break;
                case "replay": await ReplayAsync(); break;
                case "pdf": OpenPdf(); break;
                default: Log("Choose Record, Stop, Replay, or Open PDF."); break;
            }
        }
        catch (Exception ex) { Log("Command interpretation failed: " + ex.Message); }
    }

    private static string EnvPath()
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
