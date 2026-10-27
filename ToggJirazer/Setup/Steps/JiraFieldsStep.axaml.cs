using Avalonia;
using Avalonia.Controls;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Setup.Steps;

/// <summary>
/// Wizard step 5: the Jira Budget and Account field names, picked from the site's field names or,
/// when they can't be listed, typed.
/// </summary>
public partial class JiraFieldsStep : UserControl
{
    private const string DefaultBudgetField = "Budget";
    private const string DefaultAccountField = "Account";

    private WizardContext _ctx = null!;
    private AtlassianMcp? _mcp;
    private Func<IJiraTools, string, Task<IReadOnlyList<string>>> _listFieldNames = null!;
    // Set while controls are filled from code, so their change handlers don't treat it as a user choice.
    private bool _filling;

    public JiraFieldsStep()
    {
        InitializeComponent();
        // Property changes (not TextChanged, which is raised later) keep the draft in step with the boxes.
        BudgetText.PropertyChanged += FieldText_PropertyChanged;
        AccountText.PropertyChanged += FieldText_PropertyChanged;
    }

    internal void Attach(WizardContext context, AtlassianMcp? mcp, Func<IJiraTools, string, Task<IReadOnlyList<string>>> listFieldNames)
    {
        _ctx = context;
        _mcp = mcp;
        _listFieldNames = listFieldNames;
    }

    private AppConfig Draft => _ctx.Draft;

    internal async Task EnterAsync()
    {
        var generation = _ctx.LoadGeneration;
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
        if (generation != _ctx.LoadGeneration) return;

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
}
