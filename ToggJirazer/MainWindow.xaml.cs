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
    private string? _lastOutputPath;

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

        // Read the cache toggle on the UI thread before work moves to a background thread.
        var cacheEnabled = EnableCacheCheckBox.IsChecked == true;

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
            await Task.Run(async () => await RunReportAsync(config, cacheEnabled, ShowCacheWritePermissionDialog));
            AppendLog(string.Empty);
            AppendLog("Done.");
            ProgressBar.IsIndeterminate = false;
            ProgressBar.Value = 100;
            _lastOutputPath = ReadConfigFromUi().Report.OutputFile;
            OpenReportButton.IsEnabled = File.Exists(_lastOutputPath);
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

    private void OpenReport_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastOutputPath) || !File.Exists(_lastOutputPath))
        {
            AppendLog("[ERROR] Report file not found.");
            return;
        }

        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = _lastOutputPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not open report: {ex.Message}");
        }
    }

    // -------------------------------------------------------------------------
    // Report generation (called on a thread-pool thread)
    // -------------------------------------------------------------------------

    private static async Task RunReportAsync(
        AppConfig appConfig,
        bool cacheEnabled,
        Func<Exception, bool> cacheWritePermissionPrompt)
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

        // Set up the request cache (shared by both services). Disabled when the user unchecks it.
        HttpResponseCache? cache = null;
        if (cacheEnabled)
        {
            var cacheDir = Path.Combine(AppContext.BaseDirectory, "cache");
            cache = new HttpResponseCache(cacheDir, TimeSpan.FromDays(1), cacheWritePermissionPrompt);
            Console.WriteLine($"Request cache enabled (folder: {cacheDir}, expires after 1 day).");
        }
        else
        {
            Console.WriteLine("Request cache disabled.");
        }
        Console.WriteLine();

        // Fetch Toggl entries for the period
        using var togglService = new TogglService(appConfig.Toggl, cache);
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
        using var jiraService = new JiraService(appConfig.Jira, cache);
        var reportService = new ReportService(jiraService);
        var (reportSheets, versionRows, leverances, categoryConsumedHours, weeklyCategoryConsumedHours, monthlyCategoryConsumedHours, noBudgetRows) =
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
                monthlyCategoryConsumedHours,
                appConfig.Report.OutputFile,
                extraColumns,
                noBudgetRows);
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
    // Cache permission dialog
    // -------------------------------------------------------------------------

    /// <summary>
    /// Shown (on the UI thread) when a cache entry cannot be saved because of a permission
    /// problem. Returns <c>true</c> if the user wants to try again, <c>false</c> to cancel.
    /// May be called from a background thread.
    /// </summary>
    private bool ShowCacheWritePermissionDialog(Exception ex)
    {
        return Dispatcher.Invoke(() =>
        {
            var message = new TextBlock
            {
                Text = "The request cache could not be saved because of a permission problem:\n\n" +
                       ex.Message +
                       "\n\nDo you want to try again?",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 16)
            };

            var tryAgainButton = new Button { Content = "Try Again", Width = 90, Height = 28, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancelButton = new Button { Content = "Cancel", Width = 90, Height = 28, IsCancel = true };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            buttons.Children.Add(tryAgainButton);
            buttons.Children.Add(cancelButton);

            var panel = new StackPanel { Margin = new Thickness(16) };
            panel.Children.Add(message);
            panel.Children.Add(buttons);

            var dialog = new Window
            {
                Title = "Could not save cache",
                Content = panel,
                Width = 420,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this
            };

            var retry = false;
            tryAgainButton.Click += (_, _) => { retry = true; dialog.DialogResult = true; };
            cancelButton.Click += (_, _) => { retry = false; dialog.DialogResult = false; };

            dialog.ShowDialog();
            return retry;
        });
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
