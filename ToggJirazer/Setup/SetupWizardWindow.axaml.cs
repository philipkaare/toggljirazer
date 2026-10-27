using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using ToggJirazer.Cli;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Setup;

/// <summary>
/// The setup wizard. Works on <see cref="SetupWizardState.Draft"/> only; the caller saves the
/// returned draft. <c>ShowDialog&lt;AppConfig?&gt;</c> returns the draft on Finish and null on
/// Cancel or close. Atlassian sign-in and sign-out take effect immediately (session state).
/// A browser opens only from the Atlassian sign-in button and the "Use an API token instead" link.
/// </summary>
public partial class SetupWizardWindow : Window
{
    private const string TogglProfileUrl = "https://track.toggl.com/profile";
    private const string DefaultBudgetField = "Budget";
    private const string DefaultAccountField = "Account";

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

    private readonly AtlassianMcp? _mcp;
    private readonly Func<TogglSetupService> _togglFactory;
    private readonly Func<IJiraTools, string, Task<IReadOnlyList<string>>> _listFieldNames;

    // Set while controls are filled from code, so their change handlers don't treat it as a user choice.
    private bool _filling;
    private bool _savedTokenChecked;
    // True while a Toggl sign-in or token check runs: navigation waits for its answer.
    private bool _togglBusy;
    // Bumped whenever the step changes or a list load starts; a load whose generation is stale is dropped.
    private int _loadGeneration;

    // Required by Avalonia's runtime XAML loader and designer.
    public SetupWizardWindow() : this(new SetupWizardState(new AppConfig()), null, () => new TogglSetupService()) { }

    public SetupWizardWindow(
        SetupWizardState state,
        AtlassianMcp? mcp,
        Func<TogglSetupService> togglFactory,
        Func<IJiraTools, string, Task<IReadOnlyList<string>>>? listFieldNames = null)
    {
        State = state;
        _mcp = mcp;
        _togglFactory = togglFactory;
        _listFieldNames = listFieldNames ?? ListFieldNamesFromJiraAsync;
        OpenUrl = PlatformOpener.Open;

        InitializeComponent();

        // Property changes (not TextChanged, which is raised later) keep the draft in step with the boxes.
        BudgetText.PropertyChanged += FieldText_PropertyChanged;
        AccountText.PropertyChanged += FieldText_PropertyChanged;
        ReportOutputFile.PropertyChanged += ReportOutputFile_PropertyChanged;

        if (mcp != null) JiraSignIn.Initialize(mcp);
        JiraSignIn.SetSelection(Draft.Jira.CloudId, Draft.Jira.SiteUrl);
        JiraSignIn.SignInChanged += (_, _) => SyncJiraFromControl();
        JiraSignIn.SiteComboBox.SelectionChanged += (_, _) => SyncJiraFromControl();
        TogglPassword.AddHandler(KeyDownEvent, TogglPassword_KeyDown, RoutingStrategies.Tunnel);
        TogglTokenBox.AddHandler(KeyDownEvent, TogglTokenBox_KeyDown, RoutingStrategies.Tunnel);

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
        _loadGeneration++;
        var step = State.Step;

        WelcomePanel.IsVisible = step == SetupStep.Welcome;
        TogglSignInPanel.IsVisible = step == SetupStep.TogglSignIn;
        TogglProjectPanel.IsVisible = step == SetupStep.TogglProject;
        JiraSignInPanel.IsVisible = step == SetupStep.JiraSignIn;
        JiraFieldsPanel.IsVisible = step == SetupStep.JiraFields;
        ReportOutputPanel.IsVisible = step == SetupStep.ReportOutput;
        DonePanel.IsVisible = step == SetupStep.Done;

        StepCounterText.Text = $"Step {(int)step + 1} of {Titles.Count}";
        StepTitleText.Text = Titles[step];
        UpdateButtons();

        switch (step)
        {
            case SetupStep.TogglSignIn: return EnterTogglSignInAsync();
            case SetupStep.TogglProject: return LoadOrganizationsAsync();
            case SetupStep.JiraSignIn: return EnterJiraSignInAsync();
            case SetupStep.JiraFields: return LoadFieldNamesAsync();
            case SetupStep.ReportOutput: EnterReportOutput(); break;
            case SetupStep.Done: SummaryText.Text = BuildSummary(); break;
        }
        return Task.CompletedTask;
    }

    private void UpdateButtons()
    {
        var done = State.Step == SetupStep.Done;
        BackButton.IsEnabled = State.CanGoBack && !_togglBusy;
        NextButton.IsVisible = !done;
        NextButton.IsEnabled = State.CanGoNext && !_togglBusy;
        FinishButton.IsVisible = done;
    }

    private async void Next_Click(object? sender, RoutedEventArgs e)
    {
        if (State.CanGoNext && !_togglBusy) await NextAsync();
    }

    private async void Back_Click(object? sender, RoutedEventArgs e)
    {
        if (!_togglBusy) await BackAsync();
    }

    private void Finish_Click(object? sender, RoutedEventArgs e) => Close(Draft);

    private void Cancel_Click(object? sender, RoutedEventArgs e) => Close(null);

    // -------------------------------------------------------------------------
    // Toggl sign-in
    // -------------------------------------------------------------------------

    private async Task EnterTogglSignInAsync()
    {
        if (State.CanGoNext && !_savedTokenChecked)
        {
            // Show who the saved token belongs to; a failed check keeps the token and offers the sign-in form.
            _savedTokenChecked = true;
            ShowTogglSignedIn("Checking your saved Toggl token…");
            var answered = await RunTogglCheckAsync(s => s.ValidateTokenAsync(Draft.Toggl.ApiToken), busyText: string.Empty,
                failurePrefix: "Your saved Toggl token could not be checked: ");
            // The step changed before Toggl answered: check again when the step is shown next time.
            if (!answered) _savedTokenChecked = false;
        }
        else if (!TogglSignedInPanel.IsVisible)
        {
            ShowTogglForm();
        }
    }

    internal async Task SignInToTogglAsync()
    {
        var email = (TogglEmail.Text ?? string.Empty).Trim();
        var password = TogglPassword.Text ?? string.Empty;
        TogglPassword.Text = string.Empty;
        if (email.Length == 0 || password.Length == 0)
        {
            TogglStatusText.Text = "Enter your Toggl email and password.";
            return;
        }

        await RunTogglCheckAsync(s => s.SignInAsync(email, password), "Signing in…");
    }

    internal async Task CheckTogglTokenAsync()
    {
        var token = (TogglTokenBox.Text ?? string.Empty).Trim();
        if (token.Length == 0)
        {
            TogglStatusText.Text = "Paste your API token from your Toggl profile.";
            return;
        }

        await RunTogglCheckAsync(s => s.ValidateTokenAsync(token), "Checking token…");
    }

    /// <summary>
    /// Runs a Toggl account check; on success uses its API token. Errors are shown inline, never thrown.
    /// Navigation is disabled until Toggl answers. Returns false when the step changed before the answer
    /// arrived; that answer is dropped.
    /// </summary>
    private async Task<bool> RunTogglCheckAsync(Func<TogglSetupService, Task<TogglAccount>> check, string busyText,
        string failurePrefix = "")
    {
        var generation = _loadGeneration;
        SetTogglBusy(true);
        TogglStatusText.Text = busyText;
        try
        {
            using var service = _togglFactory();
            var account = await check(service);
            if (generation != _loadGeneration) return false;
            Draft.Toggl.ApiToken = account.ApiToken;
            ShowTogglSignedIn($"Signed in as {(account.FullName.Length > 0 ? account.FullName : account.Email)}");
            TogglStatusText.Text = string.Empty;
        }
        catch (Exception ex)
        {
            if (generation != _loadGeneration) return false;
            if (failurePrefix.Length > 0) ShowTogglForm();
            TogglStatusText.Text = failurePrefix + (ex is TogglSetupException ? ex.Message : $"Could not reach Toggl: {ex.Message}");
        }
        finally
        {
            SetTogglBusy(false);
        }
        return true;
    }

    private void SetTogglBusy(bool busy)
    {
        _togglBusy = busy;
        TogglSignInButton.IsEnabled = !busy;
        CheckTokenButton.IsEnabled = !busy;
        UseDifferentAccountLink.IsEnabled = !busy;
        UpdateButtons();
    }

    private void ShowTogglSignedIn(string text)
    {
        TogglSignedInText.Text = text;
        TogglSignedInPanel.IsVisible = true;
        TogglFormPanel.IsVisible = false;
        TogglTokenPanel.IsVisible = false;
    }

    private void ShowTogglForm()
    {
        TogglSignedInPanel.IsVisible = false;
        TogglFormPanel.IsVisible = true;
    }

    internal void ShowTokenEntry()
    {
        TogglTokenPanel.IsVisible = true;
        try
        {
            OpenUrl(TogglProfileUrl);
        }
        catch (Exception)
        {
            TogglStatusText.Text = $"Open {TogglProfileUrl} in your browser to copy your API token.";
        }
        TogglTokenBox.Focus();
    }

    private async void TogglSignIn_Click(object? sender, RoutedEventArgs e) => await SignInToTogglAsync();

    private async void CheckToken_Click(object? sender, RoutedEventArgs e) => await CheckTogglTokenAsync();

    private void UseToken_Click(object? sender, RoutedEventArgs e) => ShowTokenEntry();

    // Tunnel handlers: they run before the TextBox's own key handling.
    private async void TogglPassword_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || !TogglSignInButton.IsEnabled) return;
        e.Handled = true;
        await SignInToTogglAsync();
    }

    private async void TogglTokenBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || !CheckTokenButton.IsEnabled) return;
        e.Handled = true;
        await CheckTogglTokenAsync();
    }

    private void UseDifferentAccount_Click(object? sender, RoutedEventArgs e)
    {
        ShowTogglForm();
        TogglStatusText.Text = string.Empty;
        TogglEmail.Focus();
    }

    // -------------------------------------------------------------------------
    // Toggl project: organization → workspace → project
    // -------------------------------------------------------------------------

    private async Task LoadOrganizationsAsync()
    {
        ClearLists(OrganizationCombo, WorkspaceCombo, ProjectCombo);
        var organizations = await LoadListAsync("organizations", s => s.GetOrganizationsAsync(Draft.Toggl.ApiToken));
        if (organizations == null) return;

        var id = Fill(OrganizationCombo, organizations.Select(o => (o.Id, o.Name)), Draft.Toggl.OrganizationId,
            "No organizations found for this Toggl account.");
        Draft.Toggl.OrganizationId = id;
        await OnOrganizationChosenAsync(id);
    }

    private async Task OnOrganizationChosenAsync(long organizationId)
    {
        ClearLists(WorkspaceCombo, ProjectCombo);
        if (organizationId == 0)
        {
            Draft.Toggl.WorkspaceId = 0;
            Draft.Toggl.ProjectId = 0;
            UpdateButtons();
            return;
        }

        UpdateButtons();
        var workspaces = await LoadListAsync("workspaces", s => s.GetWorkspacesAsync(Draft.Toggl.ApiToken, organizationId));
        if (workspaces == null) return;

        var id = Fill(WorkspaceCombo, workspaces.Select(w => (w.Id, w.Name)), Draft.Toggl.WorkspaceId,
            "No workspaces in this organization.");
        Draft.Toggl.WorkspaceId = id;
        await OnWorkspaceChosenAsync(id);
    }

    private async Task OnWorkspaceChosenAsync(long workspaceId)
    {
        ClearLists(ProjectCombo);
        if (workspaceId == 0)
        {
            Draft.Toggl.ProjectId = 0;
            UpdateButtons();
            return;
        }

        UpdateButtons();
        var projects = await LoadListAsync("projects", s => s.GetActiveProjectsAsync(Draft.Toggl.ApiToken, workspaceId));
        if (projects == null) return;

        Draft.Toggl.ProjectId = Fill(ProjectCombo, projects.Select(p => (p.Id, p.Name)), Draft.Toggl.ProjectId,
            "No active projects in this workspace.");
        UpdateButtons();
    }

    /// <summary>Loads one Toggl list with inline loading/error text. Null when it failed or was superseded.</summary>
    private async Task<IReadOnlyList<T>?> LoadListAsync<T>(string what, Func<TogglSetupService, Task<IReadOnlyList<T>>> load)
    {
        var generation = ++_loadGeneration;
        ProjectStatusText.Text = $"Loading {what}…";
        try
        {
            using var service = _togglFactory();
            var items = await load(service);
            if (generation != _loadGeneration) return null;
            ProjectStatusText.Text = string.Empty;
            return items;
        }
        catch (Exception ex)
        {
            if (generation == _loadGeneration)
                ProjectStatusText.Text = $"Could not load your Toggl {what}: {ex.Message}";
            return null;
        }
    }

    /// <summary>Fills a dropdown, selecting <paramref name="preferredId"/> or a single option. Returns the selected ID or 0.</summary>
    private long Fill(ComboBox combo, IEnumerable<(long Id, string Name)> options, long preferredId, string emptyText)
    {
        var items = options.Select(o => new ComboBoxItem { Content = o.Name, Tag = o.Id }).ToList();
        _filling = true;
        try
        {
            combo.ItemsSource = items;
            combo.SelectedItem = items.FirstOrDefault(i => (long)i.Tag! == preferredId) ?? (items.Count == 1 ? items[0] : null);
        }
        finally
        {
            _filling = false;
        }
        if (items.Count == 0) ProjectStatusText.Text = emptyText;
        return SelectedId(combo);
    }

    private void ClearLists(params ComboBox[] combos)
    {
        _filling = true;
        try
        {
            foreach (var combo in combos) combo.ItemsSource = null;
        }
        finally
        {
            _filling = false;
        }
    }

    private static long SelectedId(ComboBox combo) => (combo.SelectedItem as ComboBoxItem)?.Tag as long? ?? 0;

    private async void OrganizationCombo_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;
        Draft.Toggl.OrganizationId = SelectedId(OrganizationCombo);
        // The old workspace and project belong to the old organization.
        Draft.Toggl.WorkspaceId = 0;
        Draft.Toggl.ProjectId = 0;
        await OnOrganizationChosenAsync(Draft.Toggl.OrganizationId);
    }

    private async void WorkspaceCombo_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;
        Draft.Toggl.WorkspaceId = SelectedId(WorkspaceCombo);
        Draft.Toggl.ProjectId = 0;
        await OnWorkspaceChosenAsync(Draft.Toggl.WorkspaceId);
    }

    private void ProjectCombo_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;
        Draft.Toggl.ProjectId = SelectedId(ProjectCombo);
        UpdateButtons();
    }

    // -------------------------------------------------------------------------
    // Jira sign-in and fields
    // -------------------------------------------------------------------------

    private async Task EnterJiraSignInAsync()
    {
        // Never prompts or opens a browser; a browser opens only from the control's sign-in button.
        await JiraSignIn.RefreshStatusAsync();
        SyncJiraFromControl();
    }

    private void SyncJiraFromControl()
    {
        State.JiraSignedIn = JiraSignIn.IsSignedIn;
        Draft.Jira.CloudId = JiraSignIn.CloudId;
        Draft.Jira.SiteUrl = JiraSignIn.SiteUrl;
        UpdateButtons();
    }

    private async Task LoadFieldNamesAsync()
    {
        var generation = _loadGeneration;
        ShowFieldCombos([]);
        BudgetCombo.IsEnabled = AccountCombo.IsEnabled = false;
        FieldsNoteText.Text = "Loading the field names from Jira…";

        IReadOnlyList<string> names;
        string? error = null;
        try
        {
            names = _mcp == null
                ? throw new InvalidOperationException("the Atlassian connection is not available")
                : await _listFieldNames(_mcp.Tools, Draft.Jira.CloudId);
        }
        catch (Exception ex)
        {
            names = [];
            error = ex.Message;
        }
        if (generation != _loadGeneration) return;

        if (names.Count > 0)
        {
            ShowFieldCombos(names);
            FieldsNoteText.Text = "Pick \"(none)\" to leave that column empty in the report.";
            return;
        }

        ShowFieldText(error == null
            ? "No Jira issues were updated in the last 90 days, so the field names could not be listed. Type them instead; leave a box empty to skip that column."
            : $"Could not list the Jira field names ({error}). Type them instead; leave a box empty to skip that column.");
    }

    private void ShowFieldCombos(IReadOnlyList<string> names)
    {
        BudgetText.IsVisible = AccountText.IsVisible = false;
        BudgetCombo.IsVisible = AccountCombo.IsVisible = true;
        BudgetCombo.IsEnabled = AccountCombo.IsEnabled = true;
        if (names.Count == 0)
        {
            ClearLists(BudgetCombo, AccountCombo);
            return;
        }

        Draft.Jira.BudgetFieldName = FillFieldCombo(BudgetCombo, names, Draft.Jira.BudgetFieldName, DefaultBudgetField);
        Draft.Jira.AccountFieldName = FillFieldCombo(AccountCombo, names, Draft.Jira.AccountFieldName, DefaultAccountField);
    }

    private string FillFieldCombo(ComboBox combo, IReadOnlyList<string> names, string current, string defaultName)
    {
        var (items, selectedIndex) = FieldChoices.Build(names, current, defaultName);
        _filling = true;
        try
        {
            combo.ItemsSource = items;
            combo.SelectedIndex = selectedIndex;
        }
        finally
        {
            _filling = false;
        }
        return FieldChoices.ToSetting(items[selectedIndex]);
    }

    private void ShowFieldText(string note)
    {
        BudgetCombo.IsVisible = AccountCombo.IsVisible = false;
        BudgetText.IsVisible = AccountText.IsVisible = true;
        _filling = true;
        try
        {
            BudgetText.Text = Draft.Jira.BudgetFieldName;
            AccountText.Text = Draft.Jira.AccountFieldName;
        }
        finally
        {
            _filling = false;
        }
        FieldsNoteText.Text = note;
    }

    private void FieldCombo_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_filling || sender is not ComboBox { SelectedItem: string item } combo) return;
        if (combo == BudgetCombo) Draft.Jira.BudgetFieldName = FieldChoices.ToSetting(item);
        else Draft.Jira.AccountFieldName = FieldChoices.ToSetting(item);
    }

    private void FieldText_PropertyChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (_filling || e.Property != TextBox.TextProperty || sender is not TextBox box) return;
        var value = (box.Text ?? string.Empty).Trim();
        if (box == BudgetText) Draft.Jira.BudgetFieldName = value;
        else Draft.Jira.AccountFieldName = value;
    }

    // -------------------------------------------------------------------------
    // Report output
    // -------------------------------------------------------------------------

    private void EnterReportOutput()
    {
        _filling = true;
        try
        {
            ReportOutputFile.Text = Draft.Report.OutputFile;
            ReportFormat.SelectedItem = ReportFormat.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(i => string.Equals(i.Content?.ToString(), Draft.Report.Format, StringComparison.OrdinalIgnoreCase))
                ?? ReportFormat.Items[0];
        }
        finally
        {
            _filling = false;
        }
        Draft.Report.Format = SelectedFormat();
        UpdateButtons();
    }

    // Exactly "xlsx" or "csv": the items' own text.
    private string SelectedFormat() => (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx";

    private void ReportOutputFile_PropertyChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (_filling || e.Property != TextBox.TextProperty) return;
        Draft.Report.OutputFile = (ReportOutputFile.Text ?? string.Empty).Trim();
        UpdateButtons();
    }

    private void ReportFormat_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;
        Draft.Report.Format = SelectedFormat();
        UpdateButtons();
    }

    private async void Browse_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var isXlsx = SelectedFormat() == "xlsx";
            var current = (ReportOutputFile.Text ?? string.Empty).Trim();
            string? initialDir = null;
            try { initialDir = Path.GetDirectoryName(Path.GetFullPath(current)); }
            catch { /* ignore invalid path */ }

            var fileType = isXlsx
                ? new FilePickerFileType("Excel files") { Patterns = ["*.xlsx"] }
                : new FilePickerFileType("CSV files") { Patterns = ["*.csv"] };
            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Select output file",
                FileTypeChoices = [fileType, FilePickerFileTypes.All],
                SuggestedFileName = Path.GetFileName(current),
                SuggestedStartLocation = initialDir == null ? null : await StorageProvider.TryGetFolderFromPathAsync(initialDir)
            });

            if (file?.TryGetLocalPath() is { } path)
                ReportOutputFile.Text = path;
        }
        catch (Exception)
        {
            // The picker is optional: the path can still be typed.
        }
    }

    // -------------------------------------------------------------------------
    // Done
    // -------------------------------------------------------------------------

    private string BuildSummary()
    {
        static string Named(ComboBox combo, long id) =>
            combo.SelectedItem is ComboBoxItem { Content: { } name } && SelectedId(combo) == id ? $"{name}" : $"ID {id}";
        static string Field(string name) => name.Length == 0 ? FieldChoices.None : name;

        var toggl = TogglSignedInPanel.IsVisible && TogglSignedInText.Text?.StartsWith("Signed in as ") == true
            ? TogglSignedInText.Text
            : "API token set";

        return string.Join(Environment.NewLine,
            $"Toggl: {toggl}",
            $"Toggl project: {Named(ProjectCombo, Draft.Toggl.ProjectId)} " +
                $"({Named(WorkspaceCombo, Draft.Toggl.WorkspaceId)}, {Named(OrganizationCombo, Draft.Toggl.OrganizationId)})",
            $"Jira site: {(Draft.Jira.SiteUrl.Length > 0 ? Draft.Jira.SiteUrl : Draft.Jira.CloudId)}",
            $"Budget field: {Field(Draft.Jira.BudgetFieldName)}",
            $"Account field: {Field(Draft.Jira.AccountFieldName)}",
            $"Report: {Draft.Report.OutputFile} ({Draft.Report.Format})");
    }
}
