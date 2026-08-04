using System.IO;
using ToggJirazer.Models;
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class SettingsServiceTests : IDisposable
{
    private readonly string _dir;
    private readonly string _settingsPath;

    public SettingsServiceTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "ToggJirazerTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
        _settingsPath = Path.Combine(_dir, "appsettings.json");
    }

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void EmailConfig_RoundTripsThroughSaveAndLoad()
    {
        var service = new SettingsService(_settingsPath);
        var config = new AppConfig
        {
            Email = new EmailConfig
            {
                SmtpHost = "smtp.example.com",
                SmtpPort = 465,
                UseSsl = true,
                Username = "philip@example.com",
                Password = "secret",
                FromAddress = "philip@example.com",
                Recipients = "a@example.com; b@example.com",
                Subject = "Timerapport {month} {year}",
                Body = "Hej,\n\nVedhæftet timerapport."
            }
        };

        service.Save(config);
        var loaded = service.Load();

        Assert.Equal("smtp.example.com", loaded.Email.SmtpHost);
        Assert.Equal(465, loaded.Email.SmtpPort);
        Assert.True(loaded.Email.UseSsl);
        Assert.Equal("philip@example.com", loaded.Email.Username);
        Assert.Equal("secret", loaded.Email.Password);
        Assert.Equal("philip@example.com", loaded.Email.FromAddress);
        Assert.Equal("a@example.com; b@example.com", loaded.Email.Recipients);
        Assert.Equal("Timerapport {month} {year}", loaded.Email.Subject);
        Assert.Equal("Hej,\n\nVedhæftet timerapport.", loaded.Email.Body);
    }

    [Fact]
    public void Load_FileWithoutEmailSection_ReturnsDefaults()
    {
        File.WriteAllText(_settingsPath, "{}");
        var loaded = new SettingsService(_settingsPath).Load();

        Assert.Equal(string.Empty, loaded.Email.SmtpHost);
        Assert.Equal(587, loaded.Email.SmtpPort);
        Assert.False(loaded.Email.UseSsl);
        Assert.Equal("Timerapport {month} {year}", loaded.Email.Subject);
        Assert.Contains("{month}", loaded.Email.Body);
    }
}
