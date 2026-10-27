using System.Globalization;
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

        ReportStartDate.SelectedDate = ParseConfigDate(config.Report.StartDate) ?? new DateTime(DateTime.Today.Year, 1, 1);
        ReportEndDate.SelectedDate = ParseConfigDate(config.Report.EndDate) ?? DateTime.Today;
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

        EmailSmtpHost.Text = config.Email.SmtpHost;
        EmailSmtpPort.Text = config.Email.SmtpPort.ToString();
        EmailUseSsl.IsChecked = config.Email.UseSsl;
        EmailUsername.Text = config.Email.Username;
        EmailPassword.Password = config.Email.Password;
        EmailFromAddress.Text = config.Email.FromAddress;
        EmailRecipients.Text = config.Email.Recipients;
        EmailSubject.Text = config.Email.Subject;
        EmailBody.Text = config.Email.Body;
    }

    private static DateTime? ParseConfigDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (DateTime.TryParseExact(value.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var exact))
            return exact;
        return DateTime.TryParse(value, out var loose) ? loose : null;
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
                StartDate = ReportStartDate.SelectedDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                EndDate = ReportEndDate.SelectedDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                OutputFile = ReportOutputFile.Text.Trim(),
                Format = (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx"
            },
            Email = new EmailConfig
            {
                SmtpHost = EmailSmtpHost.Text.Trim(),
                SmtpPort = int.TryParse(EmailSmtpPort.Text, out var smtpPort) ? smtpPort : 587,
                UseSsl = EmailUseSsl.IsChecked == true,
                Username = EmailUsername.Text.Trim(),
                Password = EmailPassword.Password,
                FromAddress = EmailFromAddress.Text.Trim(),
                Recipients = EmailRecipients.Text.Trim(),
                Subject = EmailSubject.Text,
                Body = EmailBody.Text
            }
        };
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
        => await RunReportFlowAsync(sendEmail: false);

    private async void SendEmailNow_Click(object sender, RoutedEventArgs e)
        => await RunReportFlowAsync(sendEmail: true);

    private async Task RunReportFlowAsync(bool sendEmail)
    {
        if (_isRunning) return;

        var config = ReadConfigFromUi();
        var error = ReportRunner.ValidateConfig(config);
        if (error == null && sendEmail)
        {
            error = EmailService.ValidateEmailConfig(config.Email);
            if (error == null && !string.Equals(config.Report.Format?.Trim(), "xlsx", StringComparison.OrdinalIgnoreCase))
                error = "Sending by email requires the xlsx report format (Report tab).";
        }
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
        SendEmailNowButton.IsEnabled = false;
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
            var (startDate, endDate) = ReportRunner.ResolveConfiguredPeriod(config.Report);
            var reportWritten = await Task.Run(async () => await ReportRunner.RunAsync(
                config, startDate, endDate, cacheEnabled, ShowCacheWritePermissionDialog));

            if (reportWritten)
            {
                _lastOutputPath = config.Report.OutputFile;
                OpenReportButton.IsEnabled = File.Exists(_lastOutputPath);
            }

            if (sendEmail)
            {
                if (!reportWritten)
                {
                    AppendLog("[ERROR] No report was generated; email not sent.");
                }
                else
                {
                    var subject = EmailTemplate.ResolvePlaceholders(config.Email.Subject, endDate);
                    var body = EmailTemplate.ResolvePlaceholders(config.Email.Body, endDate);
                    AppendLog($"Sending email to {string.Join(", ", EmailService.ParseRecipients(config.Email.Recipients))}...");
                    await Task.Run(async () => await EmailService.SendReportAsync(
                        config.Email, subject, body, config.Report.OutputFile));
                    AppendLog("Email sent.");
                }
            }

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
            SendEmailNowButton.IsEnabled = true;
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
