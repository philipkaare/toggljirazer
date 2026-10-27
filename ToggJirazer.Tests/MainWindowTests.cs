using System.Globalization;
using Avalonia.Headless.XUnit;
using ToggJirazer.Models;
using Xunit;

namespace ToggJirazer.Tests;

public class MainWindowTests
{
    private static MainWindow CreateWindow()
        => new(Path.Combine(Path.GetTempPath(), $"tj-win-{Guid.NewGuid():N}", "appsettings.json"));

    [AvaloniaFact]
    public void ReadConfigFromUi_round_trips_populated_config()
    {
        var window = CreateWindow();
        var input = new AppConfig
        {
            Toggl = new TogglConfig { ApiToken = "tt", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
            Jira = new JiraConfig { CloudId = "cid-1", SiteUrl = "https://x.atlassian.net" },
            Report = new ReportConfig { StartDate = "2026-01-01", EndDate = "2026-06-30", OutputFile = "r.xlsx", Format = "csv" },
            Email = new EmailConfig
            {
                SmtpHost = "smtp.x", SmtpPort = 465, UseSsl = true, Username = "u", Password = "p",
                FromAddress = "f@x", Recipients = "a@x;b@x", Subject = "S {month}", Body = "B {year}"
            }
        };

        window.PopulateUiFromConfig(input);
        var output = window.ReadConfigFromUi();

        Assert.Equal("tt", output.Toggl.ApiToken);
        Assert.Equal(1, output.Toggl.OrganizationId);
        Assert.Equal(2, output.Toggl.WorkspaceId);
        Assert.Equal(3, output.Toggl.ProjectId);
        Assert.Equal("cid-1", output.Jira.CloudId);
        Assert.Equal("https://x.atlassian.net", output.Jira.SiteUrl);
        Assert.Equal("2026-01-01", output.Report.StartDate);
        Assert.Equal("2026-06-30", output.Report.EndDate);
        Assert.Equal("r.xlsx", output.Report.OutputFile);
        Assert.Equal("csv", output.Report.Format);
        Assert.Equal("smtp.x", output.Email.SmtpHost);
        Assert.Equal(465, output.Email.SmtpPort);
        Assert.True(output.Email.UseSsl);
        Assert.Equal("u", output.Email.Username);
        Assert.Equal("p", output.Email.Password);
        Assert.Equal("f@x", output.Email.FromAddress);
        Assert.Equal("a@x;b@x", output.Email.Recipients);
        Assert.Equal("S {month}", output.Email.Subject);
        Assert.Equal("B {year}", output.Email.Body);
    }

    [AvaloniaFact]
    public void ReadConfigFromUi_preserves_jira_field_names()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(new AppConfig
        {
            Jira = new JiraConfig { BudgetFieldName = "MyBudget", AccountFieldName = "MyAccount" }
        });

        var output = window.ReadConfigFromUi();

        Assert.Equal("MyBudget", output.Jira.BudgetFieldName);
        Assert.Equal("MyAccount", output.Jira.AccountFieldName);
    }

    [AvaloniaFact]
    public void Populate_handles_empty_and_invalid_dates()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(new AppConfig
        {
            Report = new ReportConfig { StartDate = null, EndDate = "not-a-date" }
        });

        var output = window.ReadConfigFromUi();

        Assert.Equal($"{DateTime.Today.Year}-01-01", output.Report.StartDate);
        Assert.Equal(DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), output.Report.EndDate);
    }

    [AvaloniaFact]
    public void Unknown_format_selects_first_item()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(new AppConfig { Report = new ReportConfig { Format = "pdf" } });

        Assert.Equal("xlsx", window.ReadConfigFromUi().Report.Format);
    }
}
