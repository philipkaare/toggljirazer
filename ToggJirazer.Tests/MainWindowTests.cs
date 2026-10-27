using System.Globalization;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using ToggJirazer.Models;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

public class MainWindowTests
{
    private static MainWindow CreateWindow()
        => new(Path.Combine(Path.GetTempPath(), $"tj-win-{Guid.NewGuid():N}", "appsettings.json"));

    private static AtlassianMcp NotSignedInMcp()
        => new(new AtlassianTokenStore(Path.Combine(Path.GetTempPath(), $"tj-win-{Guid.NewGuid():N}", "atlassian-auth.json")),
            _ => throw new InvalidOperationException("Tests never open a browser."));

    private static MainWindow CreateWindow(AtlassianMcp mcp)
        => new(Path.Combine(Path.GetTempPath(), $"tj-win-{Guid.NewGuid():N}", "appsettings.json"), mcp);

    [AvaloniaFact]
    public void Jira_tab_round_trips_cloud_id_and_site()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(new AppConfig
        {
            Jira = new JiraConfig { CloudId = "c1", SiteUrl = "https://abc.atlassian.net" }
        });

        var output = window.ReadConfigFromUi();

        Assert.Equal("c1", output.Jira.CloudId);
        Assert.Equal("https://abc.atlassian.net", output.Jira.SiteUrl);
    }

    [AvaloniaFact]
    public async Task Jira_tab_shows_not_signed_in_without_token()
    {
        var window = CreateWindow(NotSignedInMcp());

        await window.JiraSignIn.RefreshStatusAsync();

        Assert.Equal("Not signed in", window.JiraSignIn.StatusText.Text);
        Assert.Equal("Sign in with Atlassian", window.JiraSignIn.SignInButton.Content);
    }

    [AvaloniaFact]
    public async Task Run_stops_when_user_declines_sign_in()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(ValidConfig());
        var prompted = false;
        var reportRuns = 0;
        window.ConfirmSignInAgain = () => { prompted = true; return Task.FromResult(false); };
        window.RunReport = _ => { reportRuns++; return Task.FromResult(true); };

        await window.RunReportFlowAsync(sendEmail: false);
        Dispatcher.UIThread.RunJobs();

        Assert.True(prompted);
        Assert.Equal(0, reportRuns);
        Assert.Contains("[ERROR]", window.LogTextBox.Text);
    }

    [AvaloniaFact]
    public async Task Run_logs_error_without_atlassian_connection()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(ValidConfig());
        var reportRuns = 0;
        window.RunReport = _ => { reportRuns++; return Task.FromResult(true); };

        await window.RunReportFlowAsync(sendEmail: false);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(0, reportRuns);
        Assert.Contains("[ERROR]", window.LogTextBox.Text);
    }

    private static AppConfig ValidConfig() => new()
    {
        Toggl = new TogglConfig { ApiToken = "tt", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
        Jira = new JiraConfig { CloudId = "c1", SiteUrl = "https://abc.atlassian.net" },
        Report = new ReportConfig { StartDate = "2026-01-01", EndDate = "2026-01-31", OutputFile = "r.xlsx", Format = "xlsx" }
    };

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
