using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;

namespace DesktopSteps;

internal static class BrowserIdentity
{
    public static async Task EnrichAsync(List<RecordedEvent> events, CancellationToken token = default)
    {
        var cache = new Dictionary<int, BrowserLaunchIdentity?>();
        for (var i = 0; i < events.Count; i++)
        {
            var target = events[i].Target;
            if (target?.Process?.Equals("msedge", StringComparison.OrdinalIgnoreCase) != true)
                continue;
            BrowserLaunchIdentity? identity = null;
            if (target.ProcessId is int id)
            {
                if (!cache.TryGetValue(id, out identity))
                    cache[id] = identity = await Task.Run(() => Read(id), token);
                if (identity is not null)
                    identity = identity with { ProfilePath = ResolveProfilePath(target.Window,
                        identity.UserDataDirectory) };
            }
            events[i] = events[i] with { Target = target with
            { BrowserLaunch = identity ?? new BrowserLaunchIdentity("", null, null) } };
        }
    }

    public static BrowserLaunchIdentity? Read(int processId, string? windowTitle = null)
    {
        string executable;
        try
        {
            using var process = Process.GetProcessById(processId);
            executable = process.MainModule?.FileName ?? "";
        }
        catch { return null; }
        if (string.IsNullOrWhiteSpace(executable)) return null;
        var commandLine = ReadCommandLine(processId);
        var userDataDirectory = Argument(commandLine, "user-data-dir");
        return new BrowserLaunchIdentity(executable, userDataDirectory,
            Argument(commandLine, "profile-directory"),
            string.IsNullOrWhiteSpace(commandLine) ? null :
                Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(commandLine.Trim()))),
            CommandLine: commandLine,
            ProfilePath: ResolveProfilePath(windowTitle, userDataDirectory));
    }

    public static string? ResolveProfilePath(string? windowTitle, string? userDataDirectory = null)
    {
        var root = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "Edge", "User Data");
        var stateFile = Path.Combine(root, "Local State");
        var marker = windowTitle?.LastIndexOf(" - Microsoft", StringComparison.OrdinalIgnoreCase) ?? -1;
        if (marker < 0) return null;
        var profileStart = windowTitle!.LastIndexOf(" - ", marker - 1, StringComparison.Ordinal);
        if (profileStart < 0) return null;
        var profileName = windowTitle[(profileStart + 3)..marker];
        try
        {
            if (File.Exists(stateFile))
            {
                using var document = JsonDocument.Parse(File.ReadAllText(stateFile));
                var profiles = document.RootElement.GetProperty("profile").GetProperty("info_cache");
                var matches = profiles.EnumerateObject().Where(entry =>
                {
                    var details = entry.Value;
                    return (details.TryGetProperty("name", out var name) &&
                            name.GetString()?.Equals(profileName, StringComparison.OrdinalIgnoreCase) == true) ||
                           (details.TryGetProperty("gaia_name", out var gaiaName) &&
                            gaiaName.GetString()?.Equals(profileName, StringComparison.OrdinalIgnoreCase) == true);
                }).Select(entry => Path.GetFullPath(Path.Combine(root, entry.Name))).ToList();
                if (matches.Count == 1) return matches[0];
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or
            JsonException or KeyNotFoundException or ArgumentException) { }
        // Edge's ordinary browser command line often omits --profile-directory.
        // In that case Default is the only path we can infer without profile metadata.
        var defaultPath = Path.Combine(root, "Default");
        return Directory.Exists(defaultPath) ? Path.GetFullPath(defaultPath) : null;
    }

    private static string? ReadCommandLine(int processId)
    {
        try
        {
            using var process = new Process();
            process.StartInfo = new ProcessStartInfo("powershell.exe")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            process.StartInfo.ArgumentList.Add("-NoProfile");
            process.StartInfo.ArgumentList.Add("-NonInteractive");
            process.StartInfo.ArgumentList.Add("-Command");
            process.StartInfo.ArgumentList.Add($"(Get-CimInstance Win32_Process -Filter 'ProcessId = {processId}' -ErrorAction Stop).CommandLine");
            if (!process.Start()) return null;
            var output = process.StandardOutput.ReadToEndAsync();
            if (!process.WaitForExit(3000))
            {
                process.Kill(entireProcessTree: true);
                return null;
            }
            return process.ExitCode == 0 ? output.GetAwaiter().GetResult().Trim() : null;
        }
        catch { return null; }
    }

    private static string? Argument(string? commandLine, string name)
    {
        if (string.IsNullOrWhiteSpace(commandLine)) return null;
        var match = Regex.Match(commandLine,
            @"(?:^|\s)--" + Regex.Escape(name) + @"(?:=|\s+)(?:""(?<quoted>[^""]*)""|(?<plain>\S+))",
            RegexOptions.IgnoreCase);
        return match.Success ? (match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value : match.Groups["plain"].Value) : null;
    }

    public static bool Matches(BrowserLaunchIdentity recorded, BrowserLaunchIdentity live,
        out string mismatch)
    {
        if (string.IsNullOrWhiteSpace(recorded.ExecutablePath) ||
            string.IsNullOrWhiteSpace(live.ExecutablePath))
        {
            mismatch = "executable path could not be read for both recording and replay";
            return false;
        }
        if (!Path.GetFullPath(recorded.ExecutablePath).Equals(
                Path.GetFullPath(live.ExecutablePath), StringComparison.OrdinalIgnoreCase))
        {
            mismatch = "executable path";
            return false;
        }
        if (recorded.ProfilePath is not null)
        {
            if (string.IsNullOrWhiteSpace(live.ProfilePath) ||
                !Path.GetFullPath(recorded.ProfilePath).Equals(
                    Path.GetFullPath(live.ProfilePath), StringComparison.OrdinalIgnoreCase))
            {
                mismatch = "Edge profile path";
                return false;
            }
            mismatch = "";
            return true;
        }
        if (recorded.UserDataDirectory is not null &&
            !string.Equals(recorded.UserDataDirectory, live.UserDataDirectory,
                StringComparison.OrdinalIgnoreCase))
        {
            mismatch = "user data directory";
            return false;
        }
        if (recorded.ProfileDirectory is not null &&
            !string.Equals(recorded.ProfileDirectory, live.ProfileDirectory,
                StringComparison.OrdinalIgnoreCase))
        {
            mismatch = "profile directory";
            return false;
        }
        if (recorded.CommandLineSha256 is null || live.CommandLineSha256 is null)
        {
            mismatch = "full command line could not be read for both recording and replay";
            return false;
        }
        if (!recorded.CommandLineSha256.Equals(live.CommandLineSha256,
            StringComparison.OrdinalIgnoreCase))
        {
            mismatch = "full command line fingerprint";
            return false;
        }
        mismatch = "";
        return true;
    }
}
