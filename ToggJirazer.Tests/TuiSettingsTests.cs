using System.Reflection;
using Terminal.Gui.Views;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Tui;
using Xunit;

namespace ToggJirazer.Tests;

public class TuiSettingsTests
{
    [Fact]
    public void Saving_tui_settings_preserves_email_configuration()
    {
        var path = Path.Combine(Path.GetTempPath(), $"tj-{Guid.NewGuid():N}.json");
        try
        {
            var settings = new SettingsService(path);
            settings.Save(new AppConfig { Email = new EmailConfig { SmtpHost = "smtp.example.com", Recipients = "team@example.com" } });
            var type = typeof(AppTui).Assembly.GetType("ToggJirazer.Tui.MainView")!;
            using var window = (Window)Activator.CreateInstance(type, settings)!;
            var edited = (AppConfig)type.GetMethod("ReadConfig", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(window, null)!;
            settings.Save(edited);
            var saved = settings.Load();
            Assert.Equal("smtp.example.com", saved.Email.SmtpHost);
            Assert.Equal("team@example.com", saved.Email.Recipients);
        }
        finally { File.Delete(path); }
    }
}
