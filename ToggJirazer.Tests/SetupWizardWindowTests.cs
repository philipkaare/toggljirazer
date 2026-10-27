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
        Func<IJiraTools, string, Task<IReadOnlyList<string>>>? listFieldNames = null,
        HeldHttpHandler? held = null)
    {
        var wizard = new SetupWizardWindow(
            new SetupWizardState(draft),
            NotSignedInMcp(),
            held != null
                ? () => new TogglSetupService(held.NewHandler())
                : () => new TogglSetupService(new FakeHttpHandler(toggl ?? TogglApi)),
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
        // An empty saved name means "(none)" was chosen; it is not replaced by the "Account" default.
        Assert.Equal("(none)", wizard.AccountCombo.SelectedItem);
        Assert.Equal("Old budget", wizard.State.Draft.Jira.BudgetFieldName);
        Assert.Equal("", wizard.State.Draft.Jira.AccountFieldName);

        wizard.BudgetCombo.SelectedItem = "(none)";
        Assert.Equal("", wizard.State.Draft.Jira.BudgetFieldName);
    }

    [AvaloniaFact]
    public async Task Fields_step_none_choice_survives_next_and_back()
    {
        var wizard = CreateWizard(CompleteConfig());
        wizard.Show();
        await AdvanceToAsync(wizard, SetupStep.JiraFields);

        wizard.BudgetCombo.SelectedItem = "(none)";
        await wizard.NextAsync();
        await wizard.BackAsync();

        Assert.Equal("(none)", wizard.BudgetCombo.SelectedItem);
        Assert.Equal("", wizard.State.Draft.Jira.BudgetFieldName);
    }

    private static string AuthUser(HttpRequestMessage r) =>
        System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(r.Headers.Authorization!.Parameter!)).Split(':')[0];

    /// <summary>Toggl API with two accounts: token "tokA" (Ann) and the email "bob@x" (Bob, token "tokB").</summary>
    private static HttpResponseMessage TwoAccountApi(HttpRequestMessage r) =>
        r.RequestUri!.AbsolutePath == "/api/v9/me"
            ? AuthUser(r) == "bob@x"
                ? FakeHttpHandler.Json("""{"api_token":"tokB","fullname":"Bob","email":"bob@x"}""")
                : FakeHttpHandler.Json("""{"api_token":"tokA","fullname":"Ann","email":"ann@x"}""")
            : TogglApi(r);

    [AvaloniaFact]
    public async Task Navigation_and_account_link_disabled_while_toggl_check_is_pending()
    {
        var held = new HeldHttpHandler(TwoAccountApi) { Hold = _ => true };
        var draft = CompleteConfig();
        draft.Toggl.ApiToken = "tokA";
        var wizard = CreateWizard(draft, held: held);
        wizard.Show();

        // The saved-token check: no navigation and no "Use a different account" while it runs.
        var entering = wizard.NextAsync();
        Assert.Equal(1, held.PendingCount);
        Assert.False(wizard.NextButton.IsEnabled);
        Assert.False(wizard.BackButton.IsEnabled);
        Assert.False(wizard.UseDifferentAccountLink.IsEnabled);
        held.ReleaseAll();
        await entering;
        Assert.Equal("Signed in as Ann", wizard.TogglSignedInText.Text);
        Assert.True(wizard.NextButton.IsEnabled);
        Assert.True(wizard.UseDifferentAccountLink.IsEnabled);

        // Signing in as someone else: Next must wait for Toggl's answer.
        wizard.UseDifferentAccountLink.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        wizard.TogglEmail.Text = "bob@x";
        wizard.TogglPassword.Text = "pw";
        var signingIn = wizard.SignInToTogglAsync();
        Assert.Equal(1, held.PendingCount);
        Assert.False(wizard.NextButton.IsEnabled);
        Assert.False(wizard.BackButton.IsEnabled);
        held.ReleaseAll();
        await signingIn;

        Assert.Equal("tokB", wizard.State.Draft.Toggl.ApiToken);
        Assert.True(wizard.NextButton.IsEnabled);
        Assert.True(wizard.BackButton.IsEnabled);
    }

    [AvaloniaFact]
    public async Task Toggl_result_arriving_after_step_change_is_dropped()
    {
        var held = new HeldHttpHandler(TwoAccountApi);
        var draft = CompleteConfig();
        draft.Toggl.ApiToken = "tokA";
        var wizard = CreateWizard(draft, held: held);
        wizard.Show();
        await wizard.NextAsync();
        Assert.Equal("Signed in as Ann", wizard.TogglSignedInText.Text);

        held.Hold = _ => true;
        wizard.TogglEmail.Text = "bob@x";
        wizard.TogglPassword.Text = "pw";
        var signingIn = wizard.SignInToTogglAsync();
        await wizard.BackAsync();
        held.ReleaseAll();
        await signingIn;

        Assert.Equal(SetupStep.Welcome, wizard.State.Step);
        Assert.Equal("tokA", wizard.State.Draft.Toggl.ApiToken);
    }

    /// <summary>Two organizations; organization 1 has two workspaces.</summary>
    private static HttpResponseMessage TwoOrgApi(HttpRequestMessage r) => r.RequestUri!.AbsolutePath switch
    {
        "/api/v9/me/organizations" => FakeHttpHandler.Json("""[{"id":1,"name":"Org"},{"id":8,"name":"Org 8"}]"""),
        "/api/v9/me/workspaces" => FakeHttpHandler.Json(
            """[{"id":2,"name":"WS","organization_id":1},{"id":4,"name":"WS 4","organization_id":1},{"id":9,"name":"WS 9","organization_id":8}]"""),
        "/api/v9/workspaces/4/projects" => FakeHttpHandler.Json("""[{"id":5,"name":"Proj 5"}]"""),
        "/api/v9/workspaces/9/projects" => FakeHttpHandler.Json("""[{"id":10,"name":"Proj 10"}]"""),
        _ => TogglApi(r)
    };

    private static ComboBoxItem ItemWithId(ComboBox combo, long id) =>
        combo.Items.Cast<ComboBoxItem>().Single(i => (long)i.Tag! == id);

    private static async Task RunUntilAsync(Func<bool> done)
    {
        for (var i = 0; i < 50 && !done(); i++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
    }

    [AvaloniaFact]
    public async Task Changing_organization_clears_workspace_and_project_until_loaded()
    {
        var held = new HeldHttpHandler(TwoOrgApi);
        var wizard = CreateWizard(CompleteConfig(), held: held);
        wizard.Show();
        await AdvanceToAsync(wizard, SetupStep.TogglProject);
        Dispatcher.UIThread.RunJobs();
        Assert.Equal(3, wizard.State.Draft.Toggl.ProjectId);

        held.Hold = r => r.RequestUri!.AbsolutePath == "/api/v9/me/workspaces";
        wizard.OrganizationCombo.SelectedItem = ItemWithId(wizard.OrganizationCombo, 8);

        Assert.Equal(1, held.PendingCount);
        Assert.Equal(8, wizard.State.Draft.Toggl.OrganizationId);
        Assert.Equal(0, wizard.State.Draft.Toggl.WorkspaceId);
        Assert.Equal(0, wizard.State.Draft.Toggl.ProjectId);
        Assert.False(wizard.NextButton.IsEnabled);

        held.ReleaseAll();
        await RunUntilAsync(() => wizard.State.Draft.Toggl.ProjectId != 0);
        Assert.Equal(9, wizard.State.Draft.Toggl.WorkspaceId);
        Assert.Equal(10, wizard.State.Draft.Toggl.ProjectId);
    }

    [AvaloniaFact]
    public async Task Changing_workspace_clears_project_until_loaded()
    {
        var held = new HeldHttpHandler(TwoOrgApi);
        var wizard = CreateWizard(CompleteConfig(), held: held);
        wizard.Show();
        await AdvanceToAsync(wizard, SetupStep.TogglProject);
        Dispatcher.UIThread.RunJobs();
        Assert.Equal(3, wizard.State.Draft.Toggl.ProjectId);

        held.Hold = r => r.RequestUri!.AbsolutePath == "/api/v9/workspaces/4/projects";
        wizard.WorkspaceCombo.SelectedItem = ItemWithId(wizard.WorkspaceCombo, 4);

        Assert.Equal(1, held.PendingCount);
        Assert.Equal(4, wizard.State.Draft.Toggl.WorkspaceId);
        Assert.Equal(0, wizard.State.Draft.Toggl.ProjectId);
        Assert.False(wizard.NextButton.IsEnabled);

        held.ReleaseAll();
        await RunUntilAsync(() => wizard.State.Draft.Toggl.ProjectId != 0);
        Assert.Equal(5, wizard.State.Draft.Toggl.ProjectId);
    }

    private static void PressEnter(Control box) =>
        box.RaiseEvent(new Avalonia.Input.KeyEventArgs
        {
            RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent,
            Key = Avalonia.Input.Key.Enter
        });

    [AvaloniaFact]
    public async Task Enter_in_password_box_signs_in()
    {
        var wizard = CreateWizard(new AppConfig());
        wizard.Show();
        await wizard.NextAsync();
        wizard.TogglEmail.Text = "ann@x";
        wizard.TogglPassword.Text = "secret";

        PressEnter(wizard.TogglPassword);
        await RunUntilAsync(() => wizard.State.Draft.Toggl.ApiToken.Length > 0);

        Assert.Equal("tok", wizard.State.Draft.Toggl.ApiToken);
    }

    [AvaloniaFact]
    public async Task Enter_in_token_box_checks_token()
    {
        var wizard = CreateWizard(new AppConfig());
        wizard.OpenUrl = _ => { };
        wizard.Show();
        await wizard.NextAsync();
        wizard.ShowTokenEntry();
        wizard.TogglTokenBox.Text = "pasted";

        PressEnter(wizard.TogglTokenBox);
        await RunUntilAsync(() => wizard.State.Draft.Toggl.ApiToken.Length > 0);

        Assert.Equal("tok", wizard.State.Draft.Toggl.ApiToken);
    }

    [AvaloniaFact]
    public void Secret_boxes_have_undo_disabled()
    {
        var wizard = CreateWizard(new AppConfig());

        Assert.False(wizard.TogglPassword.IsUndoEnabled);
        Assert.False(wizard.TogglTokenBox.IsUndoEnabled);
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
