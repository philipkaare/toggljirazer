using System.Net;
using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;
using ToggJirazer.Setup;
using ToggJirazer.Tests.Fakes;
using Xunit;

namespace ToggJirazer.Tests;

public class SetupWizardWindowTests
{
    private const string WrongPassword =
        "Wrong email or password. If you sign in to Toggl with Google, Apple or SSO, use an API token instead.";

    private static readonly IReadOnlyList<AtlassianSite> OneSite = [new("c1", "https://abc.atlassian.net", "ABC")];

    /// <summary>Canned Toggl API: one organization, one workspace in it, one project.</summary>
    private static HttpResponseMessage TogglApi(HttpRequestMessage request)
    {
        var path = request.RequestUri!.AbsolutePath;
        return path switch
        {
            "/api/v9/me" => FakeHttpHandler.Json("""{"api_token":"tok","fullname":"Ann Example","email":"ann@x"}"""),
            "/api/v9/me/organizations" => FakeHttpHandler.Json("""[{"id":1,"name":"Org"}]"""),
            "/api/v9/me/workspaces" => FakeHttpHandler.Json("""[{"id":2,"name":"WS","organization_id":1},{"id":9,"name":"Other","organization_id":8}]"""),
            "/api/v9/workspaces/2/projects" => FakeHttpHandler.Json("""[{"id":3,"name":"Proj"}]"""),
            _ => FakeHttpHandler.Json("{}", HttpStatusCode.NotFound)
        };
    }

    private static AtlassianMcp NotSignedInMcp()
        => new(new AtlassianTokenStore(Path.Combine(Path.GetTempPath(), $"tj-wiz-{Guid.NewGuid():N}", "atlassian-auth.json")),
            _ => throw new InvalidOperationException("Tests never open a browser."));

    private static SetupWizardWindow CreateWizard(
        AppConfig draft,
        Func<HttpRequestMessage, HttpResponseMessage>? toggl = null,
        Func<IJiraTools, string, Task<IReadOnlyList<string>>>? listFieldNames = null)
    {
        var wizard = new SetupWizardWindow(
            new SetupWizardState(draft),
            NotSignedInMcp(),
            () => new TogglSetupService(new FakeHttpHandler(toggl ?? TogglApi)),
            listFieldNames ?? ((_, _) => Task.FromResult<IReadOnlyList<string>>(["Account", "Budget"])));
        wizard.OpenUrl = _ => throw new InvalidOperationException("Tests never open a browser.");
        wizard.JiraSignIn.LoadSitesIfSignedIn = () => Task.FromResult<IReadOnlyList<AtlassianSite>?>(OneSite);
        return wizard;
    }

    private static AppConfig CompleteConfig() => new()
    {
        Toggl = new TogglConfig { ApiToken = "tok", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
        Jira = new JiraConfig { CloudId = "c1", SiteUrl = "https://abc.atlassian.net", BudgetFieldName = "Old budget", AccountFieldName = "" },
        Report = new ReportConfig { OutputFile = "r.xlsx", Format = "xlsx" },
        Email = new EmailConfig { Subject = "Keep me" }
    };

    private static async Task AdvanceToAsync(SetupWizardWindow wizard, SetupStep step)
    {
        while (wizard.State.Step != step)
            await wizard.NextAsync();
    }

    [AvaloniaFact]
    public async Task Wizard_window_next_disabled_on_empty_toggl_step()
    {
        var wizard = CreateWizard(new AppConfig());
        wizard.Show();

        await wizard.NextAsync();

        Assert.Equal(SetupStep.TogglSignIn, wizard.State.Step);
        Assert.False(wizard.NextButton.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Toggl_sign_in_failure_shows_message_and_clears_password()
    {
        var wizard = CreateWizard(new AppConfig(), _ => FakeHttpHandler.Json("{}", HttpStatusCode.Forbidden));
        wizard.Show();
        await wizard.NextAsync();
        wizard.TogglEmail.Text = "ann@x";
        wizard.TogglPassword.Text = "secret";

        await wizard.SignInToTogglAsync();

        Assert.Equal(WrongPassword, wizard.TogglStatusText.Text);
        Assert.True(string.IsNullOrEmpty(wizard.TogglPassword.Text));
        Assert.Equal("", wizard.State.Draft.Toggl.ApiToken);
        Assert.False(wizard.NextButton.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Toggl_sign_in_success_sets_token_and_shows_name()
    {
        var wizard = CreateWizard(new AppConfig());
        wizard.Show();
        await wizard.NextAsync();
        wizard.TogglEmail.Text = "ann@x";
        wizard.TogglPassword.Text = "secret";

        await wizard.SignInToTogglAsync();

        Assert.Equal("Signed in as Ann Example", wizard.TogglSignedInText.Text);
        Assert.True(wizard.TogglSignedInPanel.IsVisible);
        Assert.True(string.IsNullOrEmpty(wizard.TogglPassword.Text));
        Assert.Equal("tok", wizard.State.Draft.Toggl.ApiToken);
        Assert.True(wizard.NextButton.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Api_token_link_opens_profile_and_checked_token_is_used()
    {
        var wizard = CreateWizard(new AppConfig());
        string? opened = null;
        wizard.OpenUrl = url => opened = url;
        wizard.Show();
        await wizard.NextAsync();

        wizard.ShowTokenEntry();
        wizard.TogglTokenBox.Text = " pasted ";
        await wizard.CheckTogglTokenAsync();

        Assert.Equal("https://track.toggl.com/profile", opened);
        Assert.Equal("Signed in as Ann Example", wizard.TogglSignedInText.Text);
        Assert.Equal("tok", wizard.State.Draft.Toggl.ApiToken);
    }

    [AvaloniaFact]
    public async Task Project_step_auto_selects_single_options()
    {
        var wizard = CreateWizard(new AppConfig { Toggl = new TogglConfig { ApiToken = "tok" } });
        wizard.Show();

        await AdvanceToAsync(wizard, SetupStep.TogglProject);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(1, wizard.State.Draft.Toggl.OrganizationId);
        Assert.Equal(2, wizard.State.Draft.Toggl.WorkspaceId);
        Assert.Equal(3, wizard.State.Draft.Toggl.ProjectId);
        Assert.True(wizard.NextButton.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Project_step_shows_loading_errors_inline()
    {
        var wizard = CreateWizard(new AppConfig { Toggl = new TogglConfig { ApiToken = "tok" } },
            r => r.RequestUri!.AbsolutePath == "/api/v9/me/organizations"
                ? FakeHttpHandler.Json("""{"error":"down"}""", HttpStatusCode.InternalServerError)
                : TogglApi(r));
        wizard.Show();

        await AdvanceToAsync(wizard, SetupStep.TogglProject);
        Dispatcher.UIThread.RunJobs();

        Assert.Contains("Could not load", wizard.ProjectStatusText.Text);
        Assert.False(wizard.NextButton.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Fields_step_keeps_current_name_and_defaults_by_name()
    {
        var calls = 0;
        var wizard = CreateWizard(CompleteConfig(), listFieldNames: (_, cloudId) =>
        {
            calls++;
            Assert.Equal("c1", cloudId);
            return Task.FromResult<IReadOnlyList<string>>(["Account", "Budget", "Story Points"]);
        });
        wizard.Show();

        await AdvanceToAsync(wizard, SetupStep.JiraFields);

        Assert.Equal(1, calls);
        Assert.Equal("Old budget (current)", wizard.BudgetCombo.SelectedItem);
        Assert.Equal("Account", wizard.AccountCombo.SelectedItem);
        Assert.Equal("Old budget", wizard.State.Draft.Jira.BudgetFieldName);
        Assert.Equal("Account", wizard.State.Draft.Jira.AccountFieldName);

        wizard.BudgetCombo.SelectedItem = "(none)";
        Assert.Equal("", wizard.State.Draft.Jira.BudgetFieldName);
    }

    [AvaloniaFact]
    public async Task Fields_step_falls_back_to_text_when_listing_fails()
    {
        var wizard = CreateWizard(CompleteConfig(), listFieldNames: (_, _) => throw new AtlassianAuthException("Not signed in"));
        wizard.Show();

        await AdvanceToAsync(wizard, SetupStep.JiraFields);

        Assert.True(wizard.BudgetText.IsVisible);
        Assert.False(wizard.BudgetCombo.IsVisible);
        Assert.Equal("Old budget", wizard.BudgetText.Text);
        Assert.False(string.IsNullOrEmpty(wizard.FieldsNoteText.Text));
        Assert.True(wizard.NextButton.IsEnabled);

        wizard.AccountText.Text = "Customer";
        Assert.Equal("Customer", wizard.State.Draft.Jira.AccountFieldName);
    }

    [AvaloniaFact]
    public async Task Finish_returns_the_draft_and_cancel_returns_null()
    {
        var owner = new Window();
        owner.Show();

        var wizard = CreateWizard(CompleteConfig());
        var finished = wizard.ShowDialog<AppConfig?>(owner);
        await AdvanceToAsync(wizard, SetupStep.ReportOutput);
        wizard.ReportFormat.SelectedIndex = 1;
        await wizard.NextAsync();
        Assert.True(wizard.FinishButton.IsVisible);
        wizard.FinishButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        var result = await finished;

        Assert.NotNull(result);
        Assert.Equal("csv", result!.Report.Format);
        Assert.Equal("Keep me", result.Email.Subject);
        Assert.Equal("c1", result.Jira.CloudId);

        var cancelled = CreateWizard(CompleteConfig());
        var cancelTask = cancelled.ShowDialog<AppConfig?>(owner);
        cancelled.CancelButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        Assert.Null(await cancelTask);
    }
}
