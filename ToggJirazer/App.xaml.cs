using System.Windows;

namespace ToggJirazer;

public partial class App : Application
{
    /// <summary>Per-user settings file; set by Program before the app starts.</summary>
    public static string SettingsPath { get; set; } = string.Empty;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        new MainWindow(SettingsPath).Show();
    }
}
