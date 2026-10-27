using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;

namespace ToggJirazer.Setup.Steps;

/// <summary>Wizard step 6: where the report is saved and in which format.</summary>
public partial class ReportOutputStep : UserControl
{
    private WizardContext _ctx = null!;
    // Set while controls are filled from code, so their change handlers don't treat it as a user choice.
    private bool _filling;

    public ReportOutputStep()
    {
        InitializeComponent();
        // A property change (not TextChanged, which is raised later) keeps the draft in step with the box.
        ReportOutputFile.PropertyChanged += ReportOutputFile_PropertyChanged;
    }

    internal void Attach(WizardContext context) => _ctx = context;

    internal void Enter()
    {
        _filling = true;
        try
        {
            ReportOutputFile.Text = _ctx.Draft.Report.OutputFile;
            ReportFormat.SelectedItem = ReportFormat.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(i => string.Equals(i.Content?.ToString(), _ctx.Draft.Report.Format, StringComparison.OrdinalIgnoreCase))
                ?? ReportFormat.Items[0];
        }
        finally
        {
            _filling = false;
        }
        _ctx.Draft.Report.Format = SelectedFormat();
        _ctx.UpdateButtons();
    }

    // Exactly "xlsx" or "csv": the items' own text.
    private string SelectedFormat() => (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx";

    private void ReportOutputFile_PropertyChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (_filling || e.Property != TextBox.TextProperty) return;
        _ctx.Draft.Report.OutputFile = (ReportOutputFile.Text ?? string.Empty).Trim();
        _ctx.UpdateButtons();
    }

    private void ReportFormat_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_filling) return;
        _ctx.Draft.Report.Format = SelectedFormat();
        _ctx.UpdateButtons();
    }

    private async void Browse_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var storage = TopLevel.GetTopLevel(this)?.StorageProvider;
            if (storage == null) return;

            var isXlsx = SelectedFormat() == "xlsx";
            var current = (ReportOutputFile.Text ?? string.Empty).Trim();
            string? initialDir = null;
            try { initialDir = Path.GetDirectoryName(Path.GetFullPath(current)); }
            catch { /* ignore invalid path */ }

            var fileType = isXlsx
                ? new FilePickerFileType("Excel files") { Patterns = ["*.xlsx"] }
                : new FilePickerFileType("CSV files") { Patterns = ["*.csv"] };
            var file = await storage.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Select output file",
                FileTypeChoices = [fileType, FilePickerFileTypes.All],
                SuggestedFileName = Path.GetFileName(current),
                SuggestedStartLocation = initialDir == null ? null : await storage.TryGetFolderFromPathAsync(initialDir)
            });

            if (file?.TryGetLocalPath() is { } path)
                ReportOutputFile.Text = path;
        }
        catch (Exception)
        {
            // The picker is optional: the path can still be typed.
        }
    }
}
