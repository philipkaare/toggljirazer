using Avalonia;
using Avalonia.Headless;

[assembly: AvaloniaTestApplication(typeof(ToggJirazer.Tests.TestAppBuilder))]

namespace ToggJirazer.Tests;

public class TestAppBuilder
{
    public static AppBuilder BuildAvaloniaApp() =>
        Program.BuildAvaloniaApp().UseHeadless(new AvaloniaHeadlessPlatformOptions());
}
