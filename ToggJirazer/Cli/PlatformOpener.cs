using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ToggJirazer.Cli;

public enum OpenerOs { MacOs, Linux, Windows }

/// <summary>Opens a file with the operating system's default application.</summary>
public static class PlatformOpener
{
    public static (string FileName, string Arguments, bool UseShellExecute) CommandFor(OpenerOs os, string path)
        => os switch
        {
            OpenerOs.MacOs => ("open", Quote(path), false),
            OpenerOs.Linux => ("xdg-open", Quote(path), false),
            OpenerOs.Windows => (path, string.Empty, true),
            _ => (path, string.Empty, true),
        };

    public static void Open(string path)
    {
        var cmd = CommandFor(CurrentOs(), path);
        Process.Start(new ProcessStartInfo
        {
            FileName = cmd.FileName,
            Arguments = cmd.Arguments,
            UseShellExecute = cmd.UseShellExecute,
        });
    }

    private static OpenerOs CurrentOs()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return OpenerOs.MacOs;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return OpenerOs.Windows;
        return OpenerOs.Linux;
    }

    private static string Quote(string path) => $"\"{path}\"";
}
