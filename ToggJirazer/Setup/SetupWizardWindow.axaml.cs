using Avalonia.Controls;
using Avalonia.Interactivity;
using ToggJirazer.Cli;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;
using ToggJirazer.Setup.Steps;

namespace ToggJirazer.Setup;

/// <summary>
/// The setup wizard. Works on <see cref="SetupWizardState.Draft"/> only; the caller saves the
/// returned draft. <c>ShowDialog&lt;AppConfig?&gt;</c> returns the draft on Finish and null on
/// Cancel or close. Atlassian sign-in and sign-out take effect immediately (session state).
/// A browser opens only from the Atlassian sign-in button and the "Use an API token instead" link.
/// The window owns navigation; steps 2–6 are controls under <c>Setup/Steps</c>.
/// </summary>
public partial class SetupWizardWindow : Window
{
    private static readonly IReadOnlyDictionary<SetupStep, string> Titles = new Dictionary<SetupStep, string>
    {
        [SetupStep.Welcome] = "Welcome to ToggJirazer",
        [SetupStep.TogglSignIn] = "Sign in to Toggl",
        [SetupStep.TogglProject] = "Choose your Toggl project",
        [SetupStep.JiraSignIn] = "Sign in to Atlassian",
        [SetupStep.JiraFields] = "Jira fields",
        [SetupStep.ReportOutput] = "Report file",
        [SetupStep.Done] = "All set",
    };

    private readonly WizardContext _context;

    // Required by Avalonia's runtime XAML loader and designer.
    public SetupWizardWindow() : this(new SetupWizardState(new AppConfig()), null, () => new TogglSetupService()) { }

    public SetupWizardWindow(
        SetupWizardState state,
        AtlassianMcp? mcp,
        Func<TogglSetupService> togglFactory,
        Func<IJiraTools, string, Task<IReadOnlyList<string>>>? listFieldNames = null)
    {
        State = state;
        OpenUrl = PlatformOpener.Open;
        _context = new WizardContext(state, togglFactory, url => OpenUrl(url), UpdateButtons);

        InitializeComponent();

        TogglSignInStep.Attach(_context);
        TogglProjectStep.Attach(_context);
        JiraSignInStep.Attach(_context, mcp);
        JiraFieldsStep.Attach(_context, mcp, listFieldNames ?? ListFieldNamesFromJiraAsync);
        ReportOutputStep.Attach(_context);

        _ = ShowStepAsync();
    }

    internal SetupWizardState State { get; }

    /// <summary>Opens a URL in the browser. Replaceable by tests.</summary>
    internal Action<string> OpenUrl { get; set; }

    private AppConfig Draft => State.Draft;

    /// <summary>The default field-name lookup: the site's Jira field names through the Atlassian MCP tools.</summary>
    private static async Task<IReadOnlyList<string>> ListFieldNamesFromJiraAsync(IJiraTools tools, string cloudId)
    {
        var fields = await new JiraService(new JiraConfig { CloudId = cloudId }, tools).ListFieldNamesAsync();
        return fields.Select(f => f.Name).ToList();
    }

    // -------------------------------------------------------------------------
    // Navigation
    // -------------------------------------------------------------------------

    internal async Task NextAsync()
    {
        State.Next();
        await ShowStepAsync();
    }

    internal async Task BackAsync()
    {
        State.Back();
        await ShowStepAsync();
    }

    private Task ShowStepAsync()
    {
        _context.NextLoadGeneration();
        var step = State.Step;

        WelcomePanel.IsVisible = step == SetupStep.Welcome;
        TogglSignInStep.IsVisible = step == SetupStep.TogglSignIn;
        TogglProjectStep.IsVisible = step == SetupStep.TogglProject;
        JiraSignInStep.IsVisible = step == SetupStep.JiraSignIn;
        JiraFieldsStep.IsVisible = step == SetupStep.JiraFields;
        ReportOutputStep.IsVisible = step == SetupStep.ReportOutput;
        DonePanel.IsVisible = step == SetupStep.Done;

        StepCounterText.Text = $"Step {(int)step + 1} of {Titles.Count}";
        StepTitleText.Text = Titles[step];
        UpdateButtons();

        switch (step)
        {
            case SetupStep.TogglSignIn: return TogglSignInStep.EnterAsync();
            case SetupStep.TogglProject: return TogglProjectStep.EnterAsync();
            case SetupStep.JiraSignIn: return JiraSignInStep.EnterAsync();
            case SetupStep.JiraFields: return JiraFieldsStep.EnterAsync();
            case SetupStep.ReportOutput: ReportOutputStep.Enter(); break;
            case SetupStep.Done: SummaryText.Text = BuildSummary(); break;
        }
        return Task.CompletedTask;
    }

    // A Toggl check or an Atlassian sign-in is waiting for an answer: navigation waits too.
    private bool Busy => TogglSignInStep.IsBusy || JiraSignInStep.IsSigningIn;

    private void UpdateButtons()
    {
        var done = State.Step == SetupStep.Done;
        BackButton.IsEnabled = State.CanGoBack && !Busy;
        NextButton.IsVisible = !done;
        NextButton.IsEnabled = State.CanGoNext && !Busy;
        FinishButton.IsVisible = done;
        CancelButton.IsEnabled = !JiraSignInStep.IsSigningIn;
    }

    private async void Next_Click(object? sender, RoutedEventArgs e)
    {
        if (State.CanGoNext && !Busy) await NextAsync();
    }

    private async void Back_Click(object? sender, RoutedEventArgs e)
    {
        if (!Busy) await BackAsync();
    }

    private void Finish_Click(object? sender, RoutedEventArgs e) => Close(Draft);

    private void Cancel_Click(object? sender, RoutedEventArgs e)
    {
        if (!JiraSignInStep.IsSigningIn) Close(null);
    }

    // Closing with the title bar's X while an Atlassian sign-in waits must not leave its browser-callback
    // listener running (it would hold the port and could store tokens later without anyone seeing it).
    protected override void OnClosing(WindowClosingEventArgs e)
    {
        base.OnClosing(e);
        if (!e.Cancel) JiraSignInStep.CancelSignIn();
    }

    // -------------------------------------------------------------------------
    // Done
    // -------------------------------------------------------------------------

    private string BuildSummary()
    {
        static string Field(string name) => name.Length == 0 ? FieldChoices.None : name;

        return string.Join(Environment.NewLine,
            $"Toggl: {TogglSignInStep.Summary}",
            $"Toggl project: {TogglProjectStep.Summary}",
            $"Jira site: {(Draft.Jira.SiteUrl.Length > 0 ? Draft.Jira.SiteUrl : Draft.Jira.CloudId)}",
            $"Budget field: {Field(Draft.Jira.BudgetFieldName)}",
            $"Account field: {Field(Draft.Jira.AccountFieldName)}",
            $"Report: {Draft.Report.OutputFile} ({Draft.Report.Format})");
    }
}
