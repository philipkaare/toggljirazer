using System.IO;
using Avalonia;
using ToggJirazer.Cli;
using ToggJirazer.Services;

namespace ToggJirazer;

public static class Program
{
    private static readonly string SettingsPath = SettingsLocator.UserSettingsPath();

    [STAThread]
    private static int Main(string[] args) => RunAsync(args).GetAwaiter().GetResult();

    private static async Task<int> RunAsync(string[] args)
    {
        var mode = ArgParser.Resolve(args, Console.IsInputRedirected, Console.IsOutputRedirected);
        if (mode is not (LaunchMode.Help or LaunchMode.AutoRemoved))
        {
            SettingsLocator.EnsureUserSettings(SettingsPath,
                SettingsLocator.LegacyCandidates(Directory.GetCurrentDirectory(), AppContext.BaseDirectory),
                Path.Combine(AppContext.BaseDirectory, "appsettings.json"), Console.WriteLine);
        }
        var settings = new SettingsService(SettingsPath);

        switch (mode)
        {
            case LaunchMode.Help:
                PrintUsage();
                return 0;

            case LaunchMode.AutoRemoved:
                Console.WriteLine("Automatic monthly email has been removed. Delete the 'ToggJirazer AutoReport' scheduled task.");
                return 0;

            case LaunchMode.Gui:
                App.SettingsPath = SettingsPath;
                return BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);

            default:
                return await HeadlessRunner.RunAsync(settings.Load());
        }
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>().UsePlatformDetect().WithInterFont().LogToTrace();

    private static void PrintUsage()
    {
        Console.WriteLine("ToggJirazer — generate Toggl × Jira time reports.");
        Console.WriteLine();
        Console.WriteLine("Usage: ToggJirazer [options]");
        Console.WriteLine();
        Console.WriteLine("  (no options)     Open the GUI.");
        Console.WriteLine("  --headless,--run Run the report non-interactively using appsettings.json.");
        Console.WriteLine("  --help, -h       Show this help.");
        Console.WriteLine();
        Console.WriteLine($"Settings are read from {SettingsPath}.");
        Console.WriteLine("The generated report is written to the OutputFile path, resolved relative to");
        Console.WriteLine("the working directory.");
    }
}
