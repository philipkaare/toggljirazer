using System.Diagnostics;
using System.IO;

namespace ToggJirazer.Services;

/// <summary>
/// Registers/removes the daily Windows scheduled task that runs the app with --auto.
/// Uses PowerShell's Register-ScheduledTask because schtasks.exe cannot set
/// StartWhenAvailable (run as soon as possible after a missed start).
/// </summary>
public static class ScheduledTaskService
{
    public const string TaskName = "ToggJirazer AutoReport";

    public static bool IsEnabled()
    {
        var exitCode = RunPowerShell(
            $"if (Get-ScheduledTask -TaskName '{TaskName}' -ErrorAction SilentlyContinue) {{ exit 0 }} else {{ exit 1 }}",
            out _);
        return exitCode == 0;
    }

    public static void Enable()
    {
        var exePath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Could not determine the application executable path.");
        var workingDir = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;

        var command =
            "$ErrorActionPreference = 'Stop'; " +
            $"$action = New-ScheduledTaskAction -Execute '{Escape(exePath)}' -Argument '--auto' -WorkingDirectory '{Escape(workingDir)}'; " +
            "$trigger = New-ScheduledTaskTrigger -Daily -At 09:00; " +
            "$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable; " +
            $"Register-ScheduledTask -TaskName '{TaskName}' -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null";

        var exitCode = RunPowerShell(command, out var error);
        if (exitCode != 0)
            throw new InvalidOperationException($"Could not register the scheduled task: {error}");
    }

    public static void Disable()
    {
        var exitCode = RunPowerShell(
            $"$ErrorActionPreference = 'Stop'; Unregister-ScheduledTask -TaskName '{TaskName}' -Confirm:$false",
            out var error);
        if (exitCode != 0)
            throw new InvalidOperationException($"Could not remove the scheduled task: {error}");
    }

    private static string Escape(string value) => value.Replace("'", "''");

    private static int RunPowerShell(string command, out string error)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -NonInteractive -Command \"" + command.Replace("\"", "\\\"") + "\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardError = true
        };

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Could not start PowerShell.");
        error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return process.ExitCode;
    }
}
