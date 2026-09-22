using System.Diagnostics;

namespace DesktopSteps;

internal static class AzureCli
{
    public static ProcessStartInfo Command(bool visible, params string[] args)
    {
        var az = (Environment.GetEnvironmentVariable("PATH") ?? "")
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Select(directory => Path.Combine(directory.Trim('"'), "az.cmd"))
            .FirstOrDefault(File.Exists)
            ?? throw new FileNotFoundException("Azure CLI was not found on PATH. Install Azure CLI and restart Desktop Steps.");
        var python = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(az)!, "..", "python.exe"));
        var info = new ProcessStartInfo(File.Exists(python) ? python : az)
        {
            UseShellExecute = visible,
            CreateNoWindow = !visible,
            RedirectStandardOutput = !visible,
            RedirectStandardError = !visible
        };
        if (File.Exists(python))
        {
            info.ArgumentList.Add("-IBm");
            info.ArgumentList.Add("azure.cli");
        }
        foreach (var arg in args) info.ArgumentList.Add(arg);
        return info;
    }

    public static async Task<bool> IsSignedInAsync(CancellationToken token)
    {
        using var process = Process.Start(Command(false, "account", "get-access-token",
            "--resource", "https://ai.azure.com", "--query", "accessToken", "-o", "tsv"))!;
        var output = process.StandardOutput.ReadToEndAsync(token);
        var error = process.StandardError.ReadToEndAsync(token);
        await process.WaitForExitAsync(token);
        return process.ExitCode == 0 && !string.IsNullOrWhiteSpace(await output);
    }

    public static async Task<string?> SignedInAccountAsync(CancellationToken token)
    {
        using var process = Process.Start(Command(false, "account", "show", "--query", "user.name", "-o", "tsv"))
            ?? throw new InvalidOperationException("Azure CLI account lookup could not start.");
        var output = process.StandardOutput.ReadToEndAsync(token);
        var error = process.StandardError.ReadToEndAsync(token);
        await process.WaitForExitAsync(token);
        var account = (await output).Trim();
        await error;
        return process.ExitCode == 0 && !string.IsNullOrWhiteSpace(account) ? account : null;
    }

    public static async Task SignInAsync(CancellationToken token)
    {
        using var process = Process.Start(Command(true, "login"))
            ?? throw new InvalidOperationException("Azure CLI sign-in could not start.");
        await process.WaitForExitAsync(token);
        if (process.ExitCode != 0) throw new InvalidOperationException("Azure CLI sign-in was not completed.");
    }

    public static async Task SignOutAsync(CancellationToken token)
    {
        using var process = Process.Start(Command(false, "logout"))!;
        var output = process.StandardOutput.ReadToEndAsync(token);
        var error = process.StandardError.ReadToEndAsync(token);
        await process.WaitForExitAsync(token);
        if (process.ExitCode != 0)
            throw new InvalidOperationException("Azure CLI sign-out failed: " + await error);
    }
}
