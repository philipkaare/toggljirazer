using ToggJirazer.Cli;
using ToggJirazer.Services;
using ToggJirazer.Tui;

namespace ToggJirazer;

internal static class Program
{
    // Resolve appsettings.json from the current working directory so the file the
    // user edits (and Save writes) is the same one the app reads. Falls back to the
    // executable's directory when the app is launched from a directory without one.
    private static readonly string SettingsPath = ResolveSettingsPath();

    private static string ResolveSettingsPath()
    {
        var cwdPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
        if (File.Exists(cwdPath)) return cwdPath;
        return Path.Combine(AppContext.BaseDirectory, "appsettings.json");
    }

    private static async Task<int> Main(string[] args)
    {
        var mode = ArgParser.Resolve(args, Console.IsInputRedirected, Console.IsOutputRedirected);
        var settings = new SettingsService(SettingsPath);

        switch (mode)
        {
            case LaunchMode.Help:
                PrintUsage();
                return 0;

            case LaunchMode.Headless:
                return await HeadlessRunner.RunAsync(settings.Load());

            default:
                try
                {
                    AppTui.Run(settings, SettingsPath);
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Could not start the TUI ({ex.Message}). Running headless instead.");
                    return await HeadlessRunner.RunAsync(settings.Load());
                }
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("ToggJirazer — generate Toggl × Jira time reports.");
        Console.WriteLine();
        Console.WriteLine("Usage: ToggJirazer [options]");
        Console.WriteLine();
        Console.WriteLine("  (no options)     Launch the interactive TUI.");
        Console.WriteLine("  --headless,--run Run the report non-interactively using appsettings.json.");
        Console.WriteLine("  --help, -h       Show this help.");
        Console.WriteLine();
        Console.WriteLine("Configuration is read from appsettings.json in the current working directory");
        Console.WriteLine("(falling back to the executable's directory). The generated report is written");
        Console.WriteLine("to the OutputFile path, resolved relative to the working directory.");
    }
}
