using Avalonia.Controls;
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Setup.Steps;

/// <summary>Wizard step 3: Toggl organization → workspace → project.</summary>
public partial class TogglProjectStep : UserControl
{
    private WizardContext _ctx = null!;
    // Set while controls are filled from code, so their change handlers don't treat it as a user choice.
    private bool _filling;

    public TogglProjectStep()
    {
        InitializeComponent();
    }

    internal void Attach(WizardContext context) => _ctx = context;

    private AppConfig Draft => _ctx.Draft;

    /// <summary>The Toggl project line of the Done summary: "Project (Workspace, Organization)".</summary>
    internal string Summary =>
        $"{Named(ProjectCombo, Draft.Toggl.ProjectId)} " +
        $"({Named(WorkspaceCombo, Draft.Toggl.WorkspaceId)}, {Named(OrganizationCombo, Draft.Toggl.OrganizationId)})";

    private static string Named(ComboBox combo, long id) =>
        combo.SelectedItem is ComboBoxItem { Content: { } name } && SelectedId(combo) == id ? $"{name}" : $"ID {id}";

    internal async Task EnterAsync()
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
            _ctx.UpdateButtons();
            return;
        }

        _ctx.UpdateButtons();
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
            _ctx.UpdateButtons();
            return;
        }

        _ctx.UpdateButtons();
        var projects = await LoadListAsync("projects", s => s.GetActiveProjectsAsync(Draft.Toggl.ApiToken, workspaceId));
        if (projects == null) return;

        Draft.Toggl.ProjectId = Fill(ProjectCombo, projects.Select(p => (p.Id, p.Name)), Draft.Toggl.ProjectId,
            "No active projects in this workspace.");
        _ctx.UpdateButtons();
    }

    /// <summary>Loads one Toggl list with inline loading/error text. Null when it failed or was superseded.</summary>
    private async Task<IReadOnlyList<T>?> LoadListAsync<T>(string what, Func<TogglSetupService, Task<IReadOnlyList<T>>> load)
    {
        var generation = _ctx.NextLoadGeneration();
        ProjectStatusText.Text = $"Loading {what}…";
        try
        {
            using var service = _ctx.TogglFactory();
            var items = await load(service);
            if (generation != _ctx.LoadGeneration) return null;
            ProjectStatusText.Text = string.Empty;
            return items;
        }
        catch (Exception ex)
        {
            if (generation == _ctx.LoadGeneration)
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
        _ctx.UpdateButtons();
    }
}
