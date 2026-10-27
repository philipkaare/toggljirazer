using System.Globalization;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;
using ToggJirazer.Setup;
using ModelContextProtocol.Authentication;
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
    public async Task Accepted_sign_in_again_signs_in_then_runs_the_report()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(ValidConfig());
        var order = new List<string>();
        window.ConfirmSignInAgain = () => { order.Add("confirm"); return Task.FromResult(true); };
        window.SignInAgain = _ => { order.Add("sign-in"); return Task.CompletedTask; };
        window.RunReport = _ => { order.Add("run"); return Task.FromResult(false); };

        await window.RunReportFlowAsync(sendEmail: false);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(["confirm", "sign-in", "run"], order);
        Assert.Contains("Signed in to Atlassian.", window.LogTextBox.Text);
        Assert.DoesNotContain("[ERROR]", window.LogTextBox.Text);
    }

    [AvaloniaFact]
    public async Task Failed_sign_in_again_logs_error_and_skips_the_report()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(ValidConfig());
        window.ConfirmSignInAgain = () => Task.FromResult(true);
        window.SignInAgain = _ => throw new AtlassianAuthException("Browser closed");
        var reportRuns = 0;
        window.RunReport = _ => { reportRuns++; return Task.FromResult(true); };

        await window.RunReportFlowAsync(sendEmail: false);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(0, reportRuns);
        Assert.Contains("[ERROR] Atlassian sign-in failed: Browser closed", window.LogTextBox.Text);
        Assert.False(window.CancelSignInButton.IsVisible);
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

    [AvaloniaFact]
    public async Task Sign_in_control_is_disabled_while_a_report_runs()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(ValidConfig());
        window.ConfirmSignInAgain = () => Task.FromResult(true);
        window.SignInAgain = _ => Task.CompletedTask;
        bool? enabledDuringRun = null;
        window.RunReport = _ => { enabledDuringRun = window.JiraSignIn.IsEnabled; return Task.FromResult(false); };

        await window.RunReportFlowAsync(sendEmail: false);
        Dispatcher.UIThread.RunJobs();

        Assert.False(enabledDuringRun);
        Assert.True(window.JiraSignIn.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Sign_in_status_is_refreshed_after_the_run_not_during_it()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(ValidConfig());
        window.ConfirmSignInAgain = () => Task.FromResult(true);
        window.SignInAgain = _ => { window.JiraSignIn.StatusText.Text = "signed in by test"; return Task.CompletedTask; };
        string? statusDuringRun = null;
        window.RunReport = _ => { statusDuringRun = window.JiraSignIn.StatusText.Text; return Task.FromResult(false); };

        await window.RunReportFlowAsync(sendEmail: false);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal("signed in by test", statusDuringRun);
        // The refresh after the run finds no stored tokens.
        Assert.Equal("Not signed in", window.JiraSignIn.StatusText.Text);
    }

    [AvaloniaFact]
    public async Task Pre_run_sign_in_can_be_cancelled_from_the_log_area()
    {
        var window = CreateWindow(NotSignedInMcp());
        window.PopulateUiFromConfig(ValidConfig());
        window.ConfirmSignInAgain = () => Task.FromResult(true);
        var signInCancelled = false;
        window.SignInAgain = async ct =>
        {
            try { await Task.Delay(Timeout.Infinite, ct); }
            catch (OperationCanceledException) { signInCancelled = true; throw new AtlassianAuthException("Sign-in was cancelled or timed out.", abortsSignIn: true); }
        };
        var reportRuns = 0;
        window.RunReport = _ => { reportRuns++; return Task.FromResult(true); };
        Assert.False(window.CancelSignInButton.IsVisible);

        var flow = window.RunReportFlowAsync(sendEmail: false);
        for (var i = 0; i < 50 && !window.CancelSignInButton.IsVisible; i++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
        Assert.True(window.CancelSignInButton.IsVisible);

        window.CancelSignInButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Avalonia.Controls.Button.ClickEvent));
        await flow.WaitAsync(TimeSpan.FromSeconds(5));
        Dispatcher.UIThread.RunJobs();

        Assert.True(signInCancelled);
        Assert.Equal(0, reportRuns);
        Assert.False(window.CancelSignInButton.IsVisible);
        Assert.Contains("Atlassian sign-in was cancelled", window.LogTextBox.Text);
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

    // -------------------------------------------------------------------------
    // Setup wizard
    // -------------------------------------------------------------------------

    private static string NewSettingsPath()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"tj-win-{Guid.NewGuid():N}"));
        return Path.Combine(dir.FullName, "appsettings.json");
    }

    [AvaloniaFact]
    public async Task Setup_button_opens_wizard_and_saves_result()
    {
        var path = NewSettingsPath();
        var window = new MainWindow(path);
        var modified = ValidConfig();
        modified.Toggl.ProjectId = 42;
        modified.Report.OutputFile = "from-wizard.xlsx";
        var calls = 0;
        window.ShowWizard = _ => { calls++; return Task.FromResult<AppConfig?>(modified); };

        window.SetupButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Avalonia.Controls.Button.ClickEvent));
        Dispatcher.UIThread.RunJobs();
        await window.OpenSetupAsync();

        Assert.Equal(2, calls);
        var saved = new SettingsService(path).Load();
        Assert.Equal(42, saved.Toggl.ProjectId);
        Assert.Equal("from-wizard.xlsx", saved.Report.OutputFile);
        Assert.Equal("42", window.TogglProjectId.Text);
        Assert.Equal("from-wizard.xlsx", window.ReportOutputFile.Text);
    }

    [AvaloniaFact]
    public async Task Wizard_cancel_changes_nothing()
    {
        var path = NewSettingsPath();
        new SettingsService(path).Save(ValidConfig());
        var before = File.ReadAllBytes(path);
        var window = new MainWindow(path);
        window.ShowWizard = _ => Task.FromResult<AppConfig?>(null);

        await window.OpenSetupAsync();

        Assert.Equal(before, File.ReadAllBytes(path));
        Assert.Equal("3", window.TogglProjectId.Text);
    }

    [AvaloniaFact]
    public async Task Wizard_draft_starts_from_window_state()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(ValidConfig());
        window.EmailSubject.Text = "Unsaved subject";
        SetupWizardState? captured = null;
        window.ShowWizard = state => { captured = state; return Task.FromResult<AppConfig?>(null); };

        await window.OpenSetupAsync();

        Assert.NotNull(captured);
        Assert.Equal("Unsaved subject", captured!.Draft.Email.Subject);
    }

    [AvaloniaFact]
    public void First_run_opens_wizard_when_incomplete()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(new AppConfig());
        window.HasAtlassianToken = () => false;
        var calls = 0;
        window.ShowWizard = _ => { calls++; return Task.FromResult<AppConfig?>(null); };

        window.Show();
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(1, calls);
        window.Close();
    }

    [AvaloniaFact]
    public void Complete_config_does_not_open_wizard()
    {
        var window = CreateWindow();
        window.PopulateUiFromConfig(ValidConfig());
        window.HasAtlassianToken = () => true;
        var calls = 0;
        window.ShowWizard = _ => { calls++; return Task.FromResult<AppConfig?>(null); };

        window.Show();
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(0, calls);
        window.Close();
    }

    [Fact]
    public async Task HasStoredSignIn_reflects_the_token_file()
    {
        var store = new AtlassianTokenStore(Path.Combine(Path.GetTempPath(), $"tj-win-{Guid.NewGuid():N}", "atlassian-auth.json"));
        await using var mcp = new AtlassianMcp(store, _ => throw new InvalidOperationException("Tests never open a browser."));
        Assert.False(mcp.HasStoredSignIn);

        await store.StoreTokensAsync(new TokenContainer { TokenType = "Bearer", AccessToken = "fake", ObtainedAt = DateTimeOffset.UtcNow }, CancellationToken.None);

        Assert.True(mcp.HasStoredSignIn);
    }
}
