using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;

namespace ToggJirazer;

public partial class App : Application
{
    /// <summary>Per-user settings file; set by Program before the app starts.</summary>
    public static string SettingsPath { get; set; } = string.Empty;

    public override void Initialize() => AvaloniaXamlLoader.Load(this);

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            desktop.MainWindow = new MainWindow(SettingsPath);
        base.OnFrameworkInitializationCompleted();
    }
}
