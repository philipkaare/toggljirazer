using ToggJirazer.Models;
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class SettingsServiceTests
{
    [Fact]
    public void Save_then_Load_round_trips_config()
    {
        var path = Path.Combine(Path.GetTempPath(), $"tj-{Guid.NewGuid():N}.json");
        try
        {
            var service = new SettingsService(path);
            var original = new AppConfig
            {
                Toggl = new TogglConfig { ApiToken = "tok", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
                Jira = new JiraConfig { BaseUrl = "https://x.atlassian.net", UserEmail = "a@b.c", ApiToken = "jt", BudgetFieldName = "MyBudget", AccountFieldName = "MyAccount" },
                Report = new ReportConfig { StartDate = "2026-01-01", EndDate = "2026-06-30", OutputFile = "r.xlsx", Format = "xlsx" }
            };

            service.Save(original);
            var loaded = service.Load();

            Assert.Equal(original.Toggl.ApiToken, loaded.Toggl.ApiToken);
            Assert.Equal(original.Toggl.WorkspaceId, loaded.Toggl.WorkspaceId);
            Assert.Equal(original.Jira.BaseUrl, loaded.Jira.BaseUrl);
            Assert.Equal(original.Jira.BudgetFieldName, loaded.Jira.BudgetFieldName);
            Assert.Equal(original.Jira.AccountFieldName, loaded.Jira.AccountFieldName);
            Assert.Equal(original.Report.OutputFile, loaded.Report.OutputFile);
            Assert.Equal(original.Report.StartDate, loaded.Report.StartDate);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}

public class EmailSettingsServiceTests : IDisposable
{
    private readonly string _dir;
    private readonly string _settingsPath;

    public EmailSettingsServiceTests()
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

    [Fact]
    public void Save_does_not_persist_development_overrides()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"tj-dev-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "appsettings.json");
            File.WriteAllText(path, """{"Toggl":{"ApiToken":"base"}}""");
            File.WriteAllText(Path.Combine(dir, "appsettings.development.json"), """{"Toggl":{"ApiToken":"dev"}}""");
            var service = new SettingsService(path);

            service.Save(service.Load());

            File.Delete(Path.Combine(dir, "appsettings.development.json"));
            Assert.Equal("base", service.Load().Toggl.ApiToken);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Save_persists_user_change_to_dev_overridden_value()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"tj-dev-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "appsettings.json");
            File.WriteAllText(path, """{"Toggl":{"ApiToken":"base"}}""");
            File.WriteAllText(Path.Combine(dir, "appsettings.development.json"), """{"Toggl":{"ApiToken":"dev"}}""");
            var service = new SettingsService(path);
            var config = service.Load();
            config.Toggl.ApiToken = "edited";

            service.Save(config);

            File.Delete(Path.Combine(dir, "appsettings.development.json"));
            Assert.Equal("edited", service.Load().Toggl.ApiToken);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }
}
