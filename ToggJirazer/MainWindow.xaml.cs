using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer;

public partial class MainWindow : Window
{
    private readonly SettingsService _settingsService;
    private bool _isRunning;

    private static readonly string SettingsPath =
        Path.Combine(AppContext.BaseDirectory, "appsettings.json");

    public MainWindow()
    {
        InitializeComponent();
        _settingsService = new SettingsService(SettingsPath);
        PopulateUiFromConfig(_settingsService.Load());
    }

    // -------------------------------------------------------------------------
    // Settings helpers
    // -------------------------------------------------------------------------

    private void PopulateUiFromConfig(AppConfig config)
    {
        TogglApiToken.Password = config.Toggl.ApiToken;
        TogglOrgId.Text = config.Toggl.OrganizationId == 0 ? string.Empty : config.Toggl.OrganizationId.ToString();
        TogglWorkspaceId.Text = config.Toggl.WorkspaceId == 0 ? string.Empty : config.Toggl.WorkspaceId.ToString();
        TogglProjectId.Text = config.Toggl.ProjectId == 0 ? string.Empty : config.Toggl.ProjectId.ToString();

        JiraBaseUrl.Text = config.Jira.BaseUrl;
        JiraUserEmail.Text = config.Jira.UserEmail;
        JiraApiToken.Password = config.Jira.ApiToken;

        ReportStartDate.Text = config.Report.StartDate ?? new DateTime(DateTime.Today.Year, 1, 1).ToString("yyyy-MM-dd");
        ReportEndDate.Text = config.Report.EndDate ?? DateTime.Today.ToString("yyyy-MM-dd");
        ReportOutputFile.Text = config.Report.OutputFile;

        foreach (var item in ReportFormat.Items.Cast<ComboBoxItem>())
        {
            if (string.Equals(item.Content?.ToString(), config.Report.Format, StringComparison.OrdinalIgnoreCase))
            {
                ReportFormat.SelectedItem = item;
                break;
            }
        }
        if (ReportFormat.SelectedIndex < 0)
            ReportFormat.SelectedIndex = 0;
    }

    private AppConfig ReadConfigFromUi()
    {
        long.TryParse(TogglOrgId.Text, out var orgId);
        long.TryParse(TogglWorkspaceId.Text, out var wsId);
        long.TryParse(TogglProjectId.Text, out var projId);

        return new AppConfig
        {
            Toggl = new TogglConfig
            {
                ApiToken = TogglApiToken.Password,
                OrganizationId = orgId,
                WorkspaceId = wsId,
                ProjectId = projId
            },
            Jira = new JiraConfig
            {
                BaseUrl = JiraBaseUrl.Text.Trim(),
                UserEmail = JiraUserEmail.Text.Trim(),
                ApiToken = JiraApiToken.Password
            },
            Report = new ReportConfig
            {
                StartDate = string.IsNullOrWhiteSpace(ReportStartDate.Text) ? null : ReportStartDate.Text.Trim(),
                EndDate = string.IsNullOrWhiteSpace(ReportEndDate.Text) ? null : ReportEndDate.Text.Trim(),
                OutputFile = ReportOutputFile.Text.Trim(),
                Format = (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx"
            }
        };
    }

    private static string? ValidateConfig(AppConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.Toggl.ApiToken))
            return "Toggl API Token is required (Toggl tab).";
        if (config.Toggl.OrganizationId == 0)
            return "Toggl Organization ID is required (Toggl tab).";
        if (config.Toggl.WorkspaceId == 0)
            return "Toggl Workspace ID is required (Toggl tab).";
        if (config.Toggl.ProjectId == 0)
            return "Toggl Project ID is required (Toggl tab).";
        if (string.IsNullOrWhiteSpace(config.Jira.BaseUrl))
            return "Jira Base URL is required (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Jira.UserEmail))
            return "Jira User Email is required (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Jira.ApiToken))
            return "Jira API Token is required (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Report.OutputFile))
            return "Output File path is required (Report tab).";
        return null;
    }

    // -------------------------------------------------------------------------
    // Button handlers
    // -------------------------------------------------------------------------

    private void SaveSettings_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _settingsService.Save(ReadConfigFromUi());
            AppendLog("Settings saved.");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not save settings: {ex.Message}");
        }
    }

    private void BrowseOutputFile_Click(object sender, RoutedEventArgs e)
    {
        var format = (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx";
        var isXlsx = format.Equals("xlsx", StringComparison.OrdinalIgnoreCase);

        string? initialDir = null;
        try { initialDir = Path.GetDirectoryName(Path.GetFullPath(ReportOutputFile.Text)); }
        catch { /* ignore invalid path */ }

        var dlg = new SaveFileDialog
        {
            Title = "Select output file",
            Filter = isXlsx
                ? "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                : "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            FileName = Path.GetFileName(ReportOutputFile.Text),
            InitialDirectory = initialDir ?? string.Empty
        };

        if (dlg.ShowDialog(this) == true)
        {
            ReportOutputFile.Text = dlg.FileName;
            // Persist immediately so the path is not lost
            try { _settingsService.Save(ReadConfigFromUi()); } catch { /* ignore */ }
        }
    }

    private async void RunButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isRunning) return;

        var config = ReadConfigFromUi();
        var error = ValidateConfig(config);
        if (error != null)
        {
            AppendLog($"[ERROR] {error}");
            return;
        }

        // Auto-save settings before running so the output path is persisted
        try { _settingsService.Save(config); } catch { /* ignore */ }

        _isRunning = true;
        RunButton.IsEnabled = false;
        SaveButton.IsEnabled = false;
        LogTextBox.Clear();
        ProgressBar.IsIndeterminate = true;
        ProgressBar.Value = 0;

        // Redirect Console.Out / Console.Error to the log window for the duration of the run
        var originalOut = Console.Out;
        var originalError = Console.Error;
        var logWriter = new LogTextWriter(AppendLog);
        var errorWriter = new LogTextWriter(line => AppendLog($"[ERROR] {line}"));
        Console.SetOut(logWriter);
        Console.SetError(errorWriter);

        try
        {
            // Run on a thread-pool thread so the UI stays responsive
            await Task.Run(async () => await RunReportAsync(config));
            AppendLog(string.Empty);
            AppendLog("Done.");
            ProgressBar.IsIndeterminate = false;
            ProgressBar.Value = 100;
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
            ProgressBar.IsIndeterminate = false;
            ProgressBar.Value = 0;
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
            logWriter.Dispose();
            errorWriter.Dispose();
            RunButton.IsEnabled = true;
            SaveButton.IsEnabled = true;
            _isRunning = false;
        }
    }

    // -------------------------------------------------------------------------
    // Report generation (called on a thread-pool thread)
    // -------------------------------------------------------------------------

    private static async Task RunReportAsync(AppConfig appConfig)
    {
        // Determine reporting period
        DateTime startDate;
        DateTime endDate;

        if (!string.IsNullOrWhiteSpace(appConfig.Report.StartDate) &&
            !string.IsNullOrWhiteSpace(appConfig.Report.EndDate))
        {
            if (!DateTime.TryParse(appConfig.Report.StartDate, out startDate))
                throw new InvalidOperationException(
                    $"'Start Date' value '{appConfig.Report.StartDate}' is not a valid date. Use yyyy-MM-dd.");
            if (!DateTime.TryParse(appConfig.Report.EndDate, out endDate))
                throw new InvalidOperationException(
                    $"'End Date' value '{appConfig.Report.EndDate}' is not a valid date. Use yyyy-MM-dd.");
        }
        else
        {
            var now = DateTime.Today;
            startDate = new DateTime(now.Year, 1, 1);
            endDate = now;
            Console.WriteLine($"No period specified — defaulting to Jan 1 to today: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
        }

        Console.WriteLine($"Report period: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
        Console.WriteLine($"Toggl workspace: {appConfig.Toggl.WorkspaceId}, project: {appConfig.Toggl.ProjectId}");
        Console.WriteLine($"Jira base URL: {appConfig.Jira.BaseUrl}");
        Console.WriteLine();

        // Fetch Toggl entries for the period
        using var togglService = new TogglService(appConfig.Toggl);
        var entries = await togglService.GetDetailedReportAsync(startDate, endDate);

        if (entries.Count == 0)
        {
            Console.WriteLine("No Toggl time entries found for the specified period and project.");
            return;
        }

        Console.WriteLine();

        // Fetch all-time entries for version totals
        Console.WriteLine("Fetching all-time Toggl entries for version totals...");
        var allEntries = await togglService.GetAllEntriesAsync();
        Console.WriteLine();

        // Build the report
        using var jiraService = new JiraService(appConfig.Jira);
        var reportService = new ReportService(jiraService);
        var (reportSheets, versionRows, leverances, categoryConsumedHours, weeklyCategoryConsumedHours) =
            await reportService.BuildReportAsync(entries, allEntries);

        Console.WriteLine();
        Console.WriteLine($"Report contains {reportSheets.Sum(s => s.Rows.Count)} rows.");
        Console.WriteLine($"Version report contains {versionRows.Count} rows.");
        Console.WriteLine($"Leverances: {leverances.Count}.");
        Console.WriteLine();

        // Write output
        var format = appConfig.Report.Format?.Trim().ToLowerInvariant();
        if (format == "xlsx")
        {
            var extraColumns = reportService.ReadVersionReportExtraColumns(appConfig.Report.OutputFile);
            reportService.WriteXlsx(
                reportSheets,
                versionRows,
                leverances,
                categoryConsumedHours,
                weeklyCategoryConsumedHours,
                appConfig.Report.OutputFile,
                extraColumns);
        }
        else
        {
            if (!string.IsNullOrEmpty(format) && format != "csv")
                Console.Error.WriteLine($"Warning: Unrecognized format '{appConfig.Report.Format}'. Defaulting to CSV.");

            var rows = reportSheets.SelectMany(s => s.Rows).ToList();
            reportService.WriteCsv(rows, appConfig.Report.OutputFile);

            var outputPath = appConfig.Report.OutputFile;
            var versionOutputPath = Path.Combine(
                Path.GetDirectoryName(outputPath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputPath) + "_versions" + Path.GetExtension(outputPath));
            var extraColumns = reportService.ReadVersionReportExtraColumns(versionOutputPath);
            reportService.WriteVersionCsv(versionRows, versionOutputPath, extraColumns);
        }
    }

    // -------------------------------------------------------------------------
    // Log helpers
    // -------------------------------------------------------------------------

    /// <summary>Appends a line to the log TextBox from any thread.</summary>
    private void AppendLog(string line)
    {
        Dispatcher.BeginInvoke(() =>
        {
            LogTextBox.AppendText(line + Environment.NewLine);
            LogTextBox.ScrollToEnd();
        });
    }
}
