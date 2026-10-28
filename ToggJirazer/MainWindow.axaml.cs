using System.Globalization;
using System.IO;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;
using ToggJirazer.Setup;

namespace ToggJirazer;

public partial class MainWindow : Window
{
    private const string SignInAgainMessage = "Your Atlassian sign-in has expired. Sign in again?";

    private readonly SettingsService _settingsService;
    private readonly AtlassianMcp? _mcp;
    private readonly Func<TogglSetupService> _togglFactory;
    private bool _isRunning;
    private bool _isSetupOpen;
    private CancellationTokenSource? _preRunSignInCts; // set while the pre-run sign-in waits for the browser
    private string? _lastOutputPath;
    // Fields without UI controls (Jira custom field names) are carried through from here on Save.
    private AppConfig _loadedConfig = new();

    // Required by Avalonia's runtime XAML loader and designer.
    public MainWindow() : this(SettingsLocator.UserSettingsPath()) { }

    /// <param name="togglFactory">Toggl client for the Toggl tab's sign-in; without it the saved token isn't checked.</param>
    public MainWindow(string settingsPath, AtlassianMcp? mcp = null, Func<TogglSetupService>? togglFactory = null)
    {
        InitializeComponent();
        _settingsService = new SettingsService(settingsPath);
        _mcp = mcp;
        _togglFactory = togglFactory ?? (() => new TogglSetupService());
        ConfirmSignInAgain = ShowSignInAgainDialogAsync;
        RunReport = RunReportPipelineAsync;
        SignInAgain = ct => _mcp!.SignInAsync(ct, uri => Dispatcher.UIThread.Post(() =>
            AppendLog($"If your browser didn't open, copy this address into it: {uri.AbsoluteUri}")));
        ShowWizard = ShowSetupWizardAsync;
        HasAtlassianToken = () => _mcp?.HasStoredSignIn == true;
        PopulateUiFromConfig(_settingsService.Load());
        Opened += MainWindow_Opened;
        JiraSignIn.SigningInChanged += (_, _) => UpdateButtonsForJiraSignIn();
        TogglSignIn.TokenChanged += (_, _) => SaveAfterTogglSignIn();

        if (togglFactory != null)
        {
            TogglSignIn.Initialize(togglFactory);
            // Background check of who the saved token belongs to; never throws.
            _ = TogglSignIn.RefreshStatusAsync();
        }

        if (mcp != null)
        {
            JiraSignIn.Initialize(mcp);
            // Background status check: never prompts, opens a browser or throws.
            _ = JiraSignIn.RefreshStatusAsync();
        }
    }

    /// <summary>Asks whether to sign in again before a run. Replaceable by tests.</summary>
    internal Func<Task<bool>> ConfirmSignInAgain { get; set; }

    /// <summary>The interactive Atlassian sign-in offered before a run; cancelled by "Cancel sign-in". Replaceable by tests.</summary>
    internal Func<CancellationToken, Task> SignInAgain { get; set; }

    /// <summary>Runs the report pipeline; true when a file was written. Replaceable by tests.</summary>
    internal Func<AppConfig, Task<bool>> RunReport { get; set; }

    /// <summary>Shows the setup wizard; the draft on Finish, null on Cancel. Replaceable by tests.</summary>
    internal Func<SetupWizardState, Task<AppConfig?>> ShowWizard { get; set; }

    /// <summary>Whether an Atlassian sign-in is stored (no network). Replaceable by tests.</summary>
    internal Func<bool> HasAtlassianToken { get; set; }

    // -------------------------------------------------------------------------
    // Setup wizard
    // -------------------------------------------------------------------------

    private async void MainWindow_Opened(object? sender, EventArgs e)
    {
        if (SetupWizardState.IsSetupIncomplete(_loadedConfig, HasAtlassianToken()))
            await OpenSetupAsync();
    }

    /// <summary>
    /// Opens the wizard on the window's current (possibly unsaved) settings. Finish saves the result and
    /// refreshes the fields; Cancel changes nothing. The Jira tab's sign-in status is refreshed either way,
    /// because signing in or out inside the wizard takes effect immediately.
    /// </summary>
    internal async Task OpenSetupAsync()
    {
        if (_isSetupOpen || _isRunning || JiraSignIn.IsSigningIn) return;
        _isSetupOpen = true;
        try
        {
            var result = await ShowWizard(new SetupWizardState(ReadConfigFromUi()));
            if (result != null)
            {
                try
                {
                    _settingsService.Save(result);
                    AppendLog("Settings saved.");
                }
                catch (Exception ex)
                {
                    AppendLog($"[ERROR] Could not save settings: {ex.Message}");
                }
                PopulateUiFromConfig(result);
                _ = TogglSignIn.RefreshStatusAsync();
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Setup failed: {ex.Message}");
        }
        finally
        {
            _isSetupOpen = false;
        }

        await JiraSignIn.RefreshStatusAsync();
    }

    private Task<AppConfig?> ShowSetupWizardAsync(SetupWizardState state)
        => new SetupWizardWindow(state, _mcp, _togglFactory).ShowDialog<AppConfig?>(this);

    // -------------------------------------------------------------------------
    // Settings helpers
    // -------------------------------------------------------------------------

    internal void PopulateUiFromConfig(AppConfig config)
    {
        _loadedConfig = config;
        TogglSignIn.SetToken(config.Toggl.ApiToken);
        TogglOrgId.Text = config.Toggl.OrganizationId == 0 ? string.Empty : config.Toggl.OrganizationId.ToString();
        TogglWorkspaceId.Text = config.Toggl.WorkspaceId == 0 ? string.Empty : config.Toggl.WorkspaceId.ToString();
        TogglProjectId.Text = config.Toggl.ProjectId == 0 ? string.Empty : config.Toggl.ProjectId.ToString();
        JiraSignIn.SetSelection(config.Jira.CloudId, config.Jira.SiteUrl);

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
        EmailPassword.Text = config.Email.Password;
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

    internal AppConfig ReadConfigFromUi()
    {
        long.TryParse(TogglOrgId.Text, out var orgId);
        long.TryParse(TogglWorkspaceId.Text, out var wsId);
        long.TryParse(TogglProjectId.Text, out var projId);

        return new AppConfig
        {
            Toggl = new TogglConfig
            {
                ApiToken = TogglSignIn.ApiToken,
                OrganizationId = orgId,
                WorkspaceId = wsId,
                ProjectId = projId
            },
            Jira = new JiraConfig
            {
                CloudId = JiraSignIn.CloudId,
                SiteUrl = JiraSignIn.SiteUrl,
                BudgetFieldName = _loadedConfig.Jira.BudgetFieldName,
                AccountFieldName = _loadedConfig.Jira.AccountFieldName
            },
            Report = new ReportConfig
            {
                StartDate = ReportStartDate.SelectedDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                EndDate = ReportEndDate.SelectedDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                OutputFile = Trimmed(ReportOutputFile),
                Format = (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx"
            },
            Email = new EmailConfig
            {
                SmtpHost = Trimmed(EmailSmtpHost),
                SmtpPort = int.TryParse(EmailSmtpPort.Text, out var smtpPort) ? smtpPort : 587,
                UseSsl = EmailUseSsl.IsChecked == true,
                Username = Trimmed(EmailUsername),
                Password = EmailPassword.Text ?? string.Empty,
                FromAddress = Trimmed(EmailFromAddress),
                Recipients = Trimmed(EmailRecipients),
                Subject = EmailSubject.Text ?? string.Empty,
                Body = EmailBody.Text ?? string.Empty
            }
        };
    }

    private static string Trimmed(TextBox box) => (box.Text ?? string.Empty).Trim();

    // -------------------------------------------------------------------------
    // Button handlers
    // -------------------------------------------------------------------------

    private async void Setup_Click(object? sender, RoutedEventArgs e) => await OpenSetupAsync();

    private void SaveSettings_Click(object? sender, RoutedEventArgs e)
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

    /// <summary>Signing in or out on the Toggl tab takes effect immediately, like the Jira tab's sign-in.</summary>
    private void SaveAfterTogglSignIn()
    {
        try
        {
            _settingsService.Save(ReadConfigFromUi());
            AppendLog(TogglSignIn.ApiToken.Length > 0 ? "Signed in to Toggl; settings saved." : "Signed out of Toggl; settings saved.");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not save settings: {ex.Message}");
        }
    }

    private async void BrowseOutputFile_Click(object? sender, RoutedEventArgs e)
    {
        var format = (ReportFormat.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "xlsx";
        var isXlsx = format.Equals("xlsx", StringComparison.OrdinalIgnoreCase);

        string? initialDir = null;
        try { initialDir = Path.GetDirectoryName(Path.GetFullPath(Trimmed(ReportOutputFile))); }
        catch { /* ignore invalid path */ }

        var fileType = isXlsx
            ? new FilePickerFileType("Excel files") { Patterns = ["*.xlsx"] }
            : new FilePickerFileType("CSV files") { Patterns = ["*.csv"] };
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Select output file",
            FileTypeChoices = [fileType, FilePickerFileTypes.All],
            SuggestedFileName = Path.GetFileName(Trimmed(ReportOutputFile)),
            SuggestedStartLocation = initialDir == null ? null : await StorageProvider.TryGetFolderFromPathAsync(initialDir)
        });

        var path = file?.TryGetLocalPath();
        if (path != null)
        {
            ReportOutputFile.Text = path;
            // Persist immediately so the path is not lost
            try { _settingsService.Save(ReadConfigFromUi()); } catch { /* ignore */ }
        }
    }

    private async void RunButton_Click(object? sender, RoutedEventArgs e)
        => await RunReportFlowAsync(sendEmail: false);

    private async void SendEmailNow_Click(object? sender, RoutedEventArgs e)
        => await RunReportFlowAsync(sendEmail: true);

    /// <summary>
    /// While the Jira tab's sign-in waits for the browser, Setup… and the report buttons are disabled: the
    /// wizard and a run would use (or replace) the Atlassian connection mid sign-in. The control itself stays
    /// enabled so its "Cancel sign-in" works.
    /// </summary>
    private void UpdateButtonsForJiraSignIn()
    {
        if (_isRunning) return; // a run restores the buttons itself when it ends
        var enabled = !JiraSignIn.IsSigningIn;
        SetupButton.IsEnabled = enabled;
        RunButton.IsEnabled = enabled;
        SendEmailNowButton.IsEnabled = enabled;
    }

    internal async Task RunReportFlowAsync(bool sendEmail)
    {
        if (_isRunning || JiraSignIn.IsSigningIn) return;

        var config = ReadConfigFromUi();
        var error = ReportRunner.ValidateConfig(config);
        if (error == null && sendEmail)
        {
            error = EmailService.ValidateEmailConfig(config.Email);
            if (error == null && !string.Equals(config.Report.Format?.Trim(), "xlsx", StringComparison.OrdinalIgnoreCase))
                error = "Sending by email requires the xlsx report format (Report tab).";
        }
        if (error == null && _mcp == null)
            error = "The Atlassian connection is not available. Restart ToggJirazer.";
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
        SetupButton.IsEnabled = false;
        SendEmailNowButton.IsEnabled = false;
        // Signing in or out mid-run would drop the connection the run uses.
        JiraSignIn.IsEnabled = false;
        TogglSignIn.IsEnabled = false;
        LogTextBox.Text = string.Empty;
        ProgressBar.IsIndeterminate = true;
        ProgressBar.Value = 0;

        // Redirect Console.Out / Console.Error to the log window for the duration of the run
        var originalOut = Console.Out;
        var originalError = Console.Error;
        var logWriter = new LogTextWriter(AppendLog);
        var errorWriter = new LogTextWriter(line => AppendLog($"[ERROR] {line}"));
        Console.SetOut(logWriter);
        Console.SetError(errorWriter);

        var signInAttempted = false;
        try
        {
            if (!await EnsureSignedInAsync(_mcp!, () => signInAttempted = true))
            {
                ProgressBar.IsIndeterminate = false;
                ProgressBar.Value = 0;
                return;
            }

            var reportWritten = await RunReport(config);

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
                    var (_, endDate) = ReportRunner.ResolveConfiguredPeriod(config.Report);
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

            // A sign-in check during the run could reset the connection the run uses, so the status is
            // refreshed once the run is over (before a new run can start). RefreshStatusAsync never throws.
            if (signInAttempted)
                await JiraSignIn.RefreshStatusAsync();

            RunButton.IsEnabled = true;
            SaveButton.IsEnabled = true;
            SetupButton.IsEnabled = true;
            SendEmailNowButton.IsEnabled = true;
            JiraSignIn.IsEnabled = true;
            TogglSignIn.IsEnabled = true;
            _isRunning = false;
        }
    }

    /// <summary>
    /// Checks the Atlassian sign-in before a run and, when it is missing or expired, asks to sign in
    /// again. Returns false, after logging an error, when the run should stop. Server trouble during the check
    /// (a 5xx tool or transport error) isn't a missing sign-in: the run stops without offering a sign-in, which
    /// would drop working tokens. Any other tool error (e.g. a refresh rejected as invalid_grant) offers it.
    /// </summary>
    private async Task<bool> EnsureSignedInAsync(AtlassianMcp mcp, Action onSignInAttempt)
    {
        SignInStatus status;
        try
        {
            status = await mcp.CheckSignInAsync(CancellationToken.None);
        }
        catch (Exception ex) when (AtlassianMcp.ServerStatus(ex) != null)
        {
            LogServerTrouble(ex.Message);
            return false;
        }

        if (status.IsSignedIn)
            return true;

        if (status.IsServerTrouble)
        {
            LogServerTrouble(status.ToolError);
            return false;
        }

        if (!await ConfirmSignInAgain())
        {
            AppendLog("[ERROR] Not signed in to Atlassian; the report was not run. Use 'Sign in with Atlassian' on the Jira tab.");
            return false;
        }

        onSignInAttempt();
        using var cts = new CancellationTokenSource();
        _preRunSignInCts = cts;
        CancelSignInButton.IsVisible = true;
        try
        {
            AppendLog("Waiting for you to finish signing in in your browser… (use 'Cancel sign-in' to stop)");
            await SignInAgain(cts.Token);
        }
        catch (Exception) when (cts.IsCancellationRequested)
        {
            AppendLog("[ERROR] Atlassian sign-in was cancelled; the report was not run.");
            return false;
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Atlassian sign-in failed: {ex.Message}");
            return false;
        }
        finally
        {
            _preRunSignInCts = null;
            CancelSignInButton.IsVisible = false;
        }

        AppendLog("Signed in to Atlassian.");
        return true;
    }

    private void LogServerTrouble(string? details) =>
        AppendLog($"[ERROR] {AtlassianMcp.ServerTroubleMessage} Details: {details}");

    private void CancelSignIn_Click(object? sender, RoutedEventArgs e) => _preRunSignInCts?.Cancel();

    /// <summary>The real report pipeline: Toggl fetch, Jira lookup through MCP, report file.</summary>
    private Task<bool> RunReportPipelineAsync(AppConfig config)
    {
        // Read the cache toggle on the UI thread before work moves to a background thread.
        var cacheEnabled = EnableCacheCheckBox.IsChecked == true;
        var jiraTools = _mcp!.Tools;
        return Task.Run(async () =>
        {
            var (startDate, endDate) = ReportRunner.ResolveConfiguredPeriod(config.Report);
            return await ReportRunner.RunAsync(
                config, jiraTools, startDate, endDate, cacheEnabled, ShowCacheWritePermissionDialog);
        });
    }

    private void OpenReport_Click(object? sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastOutputPath) || !File.Exists(_lastOutputPath))
        {
            AppendLog("[ERROR] Report file not found.");
            return;
        }

        try
        {
            Cli.PlatformOpener.Open(_lastOutputPath);
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not open report: {ex.Message}");
        }
    }

    // -------------------------------------------------------------------------
    // Dialogs
    // -------------------------------------------------------------------------

    /// <summary>
    /// Shown (on the UI thread) when a cache entry cannot be saved because of a permission
    /// problem. Returns <c>true</c> if the user wants to try again, <c>false</c> to cancel.
    /// May be called from a background thread.
    /// </summary>
    private bool ShowCacheWritePermissionDialog(Exception ex)
        => Dispatcher.UIThread.InvokeAsync(() => ShowConfirmDialogAsync(
            "Could not save cache",
            "The request cache could not be saved because of a permission problem:\n\n" +
            ex.Message +
            "\n\nDo you want to try again?",
            "Try Again")).GetAwaiter().GetResult();

    private Task<bool> ShowSignInAgainDialogAsync()
        => ShowConfirmDialogAsync("Atlassian sign-in", SignInAgainMessage, "Sign in");

    /// <summary>A modal dialog with a confirm button and Cancel; returns <c>true</c> on confirm.</summary>
    private async Task<bool> ShowConfirmDialogAsync(string title, string text, string confirmText)
    {
        var message = new TextBlock
        {
            Text = text,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin = new Avalonia.Thickness(0, 0, 0, 16)
        };

        var confirmButton = new Button { Content = confirmText, Width = 90, Margin = new Avalonia.Thickness(0, 0, 8, 0), IsDefault = true };
        var cancelButton = new Button { Content = "Cancel", Width = 90, IsCancel = true };

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttons.Children.Add(confirmButton);
        buttons.Children.Add(cancelButton);

        var panel = new StackPanel { Margin = new Avalonia.Thickness(16) };
        panel.Children.Add(message);
        panel.Children.Add(buttons);

        var dialog = new Window
        {
            Title = title,
            Content = panel,
            Width = 420,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };

        confirmButton.Click += (_, _) => dialog.Close(true);
        cancelButton.Click += (_, _) => dialog.Close(false);

        return await dialog.ShowDialog<bool>(this);
    }

    // -------------------------------------------------------------------------
    // Log helpers
    // -------------------------------------------------------------------------

    /// <summary>Appends a line to the log TextBox from any thread.</summary>
    private void AppendLog(string line)
    {
        Dispatcher.UIThread.Post(() =>
        {
            LogTextBox.Text += line + Environment.NewLine;
            LogTextBox.CaretIndex = LogTextBox.Text.Length;
        });
    }
}
