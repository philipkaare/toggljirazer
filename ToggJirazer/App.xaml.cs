using System.Windows;
using ToggJirazer.Services;

namespace ToggJirazer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (e.Args.Any(a => a.Equals("--auto", StringComparison.OrdinalIgnoreCase)))
        {
            // Headless mode: no window; exit code signals success/failure.
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            RunAutoAndShutdown();
        }
        else
        {
            new MainWindow().Show();
        }
    }

    private async void RunAutoAndShutdown()
    {
        int exitCode;
        try
        {
            exitCode = await AutoSendRunner.RunAsync(AppContext.BaseDirectory);
        }
        catch
        {
            exitCode = 1;
        }
        Shutdown(exitCode);
    }
}
