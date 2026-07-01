// Terminal.Gui 2.4.16 API notes (vs. brief which targeted 2.x conceptually):
//
//   1. TabView + Tab       → Tabs (Terminal.Gui.Views.Tabs) with InsertTab(int, View)
//                            Tab label = inserted View.Title; Value = currently selected View
//   2. RadioGroup          → NOT in 2.4.16; replaced by two CheckBox controls with RadioStyle=true
//   3. Application.Init/Run/Shutdown/Invoke  → static API is [Obsolete] in 2.4.x;
//                            suppressed at file scope. Signatures:
//                            Init(string? driverName = null); Run(IRunnable, Func<Exception,bool>?);
//                            Shutdown(); Invoke(Action)
//   4. Button label        → Button.Text (not .Title) per reference usage in uRobot TuiShell.cs
//   5. TextView            → [Obsolete] in 2.4.16 but still functional; suppressed below
//   6. Window.RequestStop()→ Runnable.RequestStop() — stops the run loop; replaces Application.RequestStop()
//   7. Key / KeyCode       → Terminal.Gui.Input.Key, Terminal.Gui.Drivers.KeyCode
//   8. CheckBox.RadioStyle → makes checkbox render as a radio button (mutual exclusivity is manual)
//   9. ProgressBarStyle    → Terminal.Gui.Views.ProgressBarStyle  (MarqueeBlocks exists in 2.4.16)
//  10. Application.Invoke  → works as static shorthand even though Obsolete (suppressed)

#pragma warning disable CS0618  // Suppress [Obsolete] on static Application API and TextView

using System.Reflection;
using Terminal.Gui.App;
using Terminal.Gui.Drivers;
using Terminal.Gui.ViewBase;
using Terminal.Gui.Views;
using ToggJirazer.Cli;
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Tui;

/// <summary>The interactive Terminal.Gui screen: edit settings, generate the
/// report, watch progress, open the result.</summary>
public static class AppTui
{
    public static void Run(SettingsService settings, string settingsPath)
    {
        Application.Init();
        try
        {
            var ui = new MainView(settings);
            Application.Run(ui, null!);
        }
        finally
        {
            Application.Shutdown();
        }
    }
}

internal sealed class MainView : Window
{
    private readonly SettingsService _settings;

    // --- Toggl fields ---
    private readonly TextField _togglToken = new() { Secret = true };
    private readonly TextField _togglOrg = new();
    private readonly TextField _togglWorkspace = new();
    private readonly TextField _togglProject = new();

    // --- Jira fields ---
    private readonly TextField _jiraBaseUrl = new();
    private readonly TextField _jiraEmail = new();
    private readonly TextField _jiraToken = new() { Secret = true };

    // --- Report fields ---
    private readonly TextField _startDate = new();
    private readonly TextField _endDate = new();
    private readonly TextField _outputFile = new();

    // Format selector: CheckBox pair with RadioStyle=true replaces RadioGroup (gone in 2.4.16)
    private readonly CheckBox _xlsxBox = new() { Text = "xlsx", RadioStyle = true };
    private readonly CheckBox _csvBox = new() { Text = "csv", RadioStyle = true };
    private bool _suppressFormatChange; // guard against recursive mutual-exclusivity loop

    // --- Retained Jira field names (not exposed in UI; preserved across read/save) ---
    private string _jiraBudgetFieldName = "Budget";
    private string _jiraAccountFieldName = "Account";

    // --- Log pane ---
    private readonly TextView _log = new()
    {
        ReadOnly = true,
        WordWrap = false,
        ScrollBars = true,
    };

    // --- Progress indicator ---
    private readonly ProgressBar _progress = new()
    {
        ProgressBarStyle = ProgressBarStyle.MarqueeBlocks,
        Visible = false,
    };

    // --- Action buttons ---
    private readonly Button _saveBtn = new() { Text = "Save (F2)" };
    private readonly Button _runBtn = new() { Text = "Generate (F5)" };
    private readonly Button _openBtn = new() { Text = "Open report", Enabled = false };

    private bool _isRunning;
    private string? _lastOutputPath;

    public MainView(SettingsService settings)
    {
        _settings = settings;
        Title = "ToggJirazer  (Ctrl+Q to quit)";
        Width = Dim.Fill();
        Height = Dim.Fill();

        var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.0";
        var banner = new Label { Text = Banner.Text(version), X = 0, Y = 0, Height = 6, Width = Dim.Fill() };

        var tabs = BuildTabs();
        tabs.X = 0; tabs.Y = 6; tabs.Width = Dim.Fill(); tabs.Height = 12;

        _saveBtn.X = 0;  _saveBtn.Y = Pos.Bottom(tabs) + 1;
        _runBtn.X = Pos.Right(_saveBtn) + 2; _runBtn.Y = Pos.Bottom(tabs) + 1;
        _openBtn.X = Pos.Right(_runBtn) + 2; _openBtn.Y = Pos.Bottom(tabs) + 1;

        _progress.X = 0; _progress.Y = Pos.Bottom(_saveBtn) + 1;
        _progress.Width = Dim.Fill(); _progress.Height = 1;

        var logFrame = new FrameView
        {
            Title = "Log",
            X = 0, Y = Pos.Bottom(_progress) + 1,
            Width = Dim.Fill(), Height = Dim.Fill(),
        };
        _log.X = 0; _log.Y = 0; _log.Width = Dim.Fill(); _log.Height = Dim.Fill();
        logFrame.Add(_log);

        Add(banner, tabs, _saveBtn, _runBtn, _openBtn, _progress, logFrame);

        // Wire button actions
        _saveBtn.Accepting += (_, _) => SaveSettings();
        _runBtn.Accepting  += (_, _) => _ = GenerateAsync();
        _openBtn.Accepting += (_, _) => OpenReport();

        // Wire format checkbox mutual exclusivity
        _xlsxBox.ValueChanged += (_, _) =>
        {
            if (!_suppressFormatChange && _xlsxBox.Value == CheckState.Checked)
            {
                _suppressFormatChange = true;
                _csvBox.Value = CheckState.UnChecked;
                _suppressFormatChange = false;
            }
        };
        _csvBox.ValueChanged += (_, _) =>
        {
            if (!_suppressFormatChange && _csvBox.Value == CheckState.Checked)
            {
                _suppressFormatChange = true;
                _xlsxBox.Value = CheckState.UnChecked;
                _suppressFormatChange = false;
            }
        };

        // Ctrl+Q quits (explicit, because focused views swallow it otherwise)
        KeyDown += (_, key) =>
        {
            if (key.IsCtrl && key.NoCtrl.NoShift.NoAlt.KeyCode == KeyCode.Q)
            {
                RequestStop();
                key.Handled = true;
            }
            else if (key.KeyCode == KeyCode.F2)
            {
                SaveSettings();
                key.Handled = true;
            }
            else if (key.KeyCode == KeyCode.F5)
            {
                _ = GenerateAsync();
                key.Handled = true;
            }
        };

        PopulateFromConfig(_settings.Load());
    }

    // ── Tab layout ─────────────────────────────────────────────────────────────

    private Tabs BuildTabs()
    {
        var tabs = new Tabs();

        // Toggl tab: each tab is a View whose Title becomes the tab-strip label
        var toggl = new View { Title = "Toggl", Width = Dim.Fill(), Height = Dim.Fill() };
        AddField(toggl, "API Token:", _togglToken, 0);
        AddField(toggl, "Organization:", _togglOrg, 2);
        AddField(toggl, "Workspace:", _togglWorkspace, 4);
        AddField(toggl, "Project:", _togglProject, 6);
        tabs.InsertTab(0, toggl);

        // Jira tab
        var jira = new View { Title = "Jira", Width = Dim.Fill(), Height = Dim.Fill() };
        AddField(jira, "Base URL:", _jiraBaseUrl, 0);
        AddField(jira, "User Email:", _jiraEmail, 2);
        AddField(jira, "API Token:", _jiraToken, 4);
        tabs.InsertTab(1, jira);

        // Report tab (includes xlsx/csv selector)
        var report = new View { Title = "Report", Width = Dim.Fill(), Height = Dim.Fill() };
        AddField(report, "Start Date:", _startDate, 0);
        AddField(report, "End Date:", _endDate, 2);
        AddField(report, "Output File:", _outputFile, 4);
        var fmtLabel = new Label { Text = "Format:", X = 0, Y = 6 };
        _xlsxBox.X = 16; _xlsxBox.Y = 6;
        _csvBox.X = 24; _csvBox.Y = 6;
        report.Add(fmtLabel, _xlsxBox, _csvBox);
        tabs.InsertTab(2, report);

        // Select the first tab initially
        tabs.Value = toggl;

        return tabs;
    }

    private static void AddField(View parent, string label, TextField field, int row)
    {
        var lbl = new Label { Text = label, X = 0, Y = row };
        field.X = 16; field.Y = row; field.Width = Dim.Fill(1);
        parent.Add(lbl, field);
    }

    // ── Config ↔ UI ────────────────────────────────────────────────────────────

    private void PopulateFromConfig(AppConfig c)
    {
        _togglToken.Text = c.Toggl.ApiToken;
        _togglOrg.Text = c.Toggl.OrganizationId == 0 ? "" : c.Toggl.OrganizationId.ToString();
        _togglWorkspace.Text = c.Toggl.WorkspaceId == 0 ? "" : c.Toggl.WorkspaceId.ToString();
        _togglProject.Text = c.Toggl.ProjectId == 0 ? "" : c.Toggl.ProjectId.ToString();
        _jiraBaseUrl.Text = c.Jira.BaseUrl;
        _jiraEmail.Text = c.Jira.UserEmail;
        _jiraToken.Text = c.Jira.ApiToken;
        _jiraBudgetFieldName = c.Jira.BudgetFieldName;
        _jiraAccountFieldName = c.Jira.AccountFieldName;
        _startDate.Text = c.Report.StartDate
            ?? new DateTime(DateTime.Today.Year, 1, 1).ToString("yyyy-MM-dd");
        _endDate.Text = c.Report.EndDate ?? DateTime.Today.ToString("yyyy-MM-dd");
        _outputFile.Text = c.Report.OutputFile;

        var isCsv = string.Equals(c.Report.Format, "csv", StringComparison.OrdinalIgnoreCase);
        _suppressFormatChange = true;
        _xlsxBox.Value = isCsv ? CheckState.UnChecked : CheckState.Checked;
        _csvBox.Value  = isCsv ? CheckState.Checked   : CheckState.UnChecked;
        _suppressFormatChange = false;
    }

    private AppConfig ReadConfig()
    {
        long.TryParse(_togglOrg.Text, out var org);
        long.TryParse(_togglWorkspace.Text, out var ws);
        long.TryParse(_togglProject.Text, out var proj);
        return new AppConfig
        {
            Toggl = new TogglConfig
            {
                ApiToken = _togglToken.Text,
                OrganizationId = org,
                WorkspaceId = ws,
                ProjectId = proj,
            },
            Jira = new JiraConfig
            {
                BaseUrl = _jiraBaseUrl.Text.Trim(),
                UserEmail = _jiraEmail.Text.Trim(),
                ApiToken = _jiraToken.Text,
                BudgetFieldName = _jiraBudgetFieldName,
                AccountFieldName = _jiraAccountFieldName,
            },
            Report = new ReportConfig
            {
                StartDate = string.IsNullOrWhiteSpace(_startDate.Text) ? null : _startDate.Text.Trim(),
                EndDate   = string.IsNullOrWhiteSpace(_endDate.Text)   ? null : _endDate.Text.Trim(),
                OutputFile = _outputFile.Text.Trim(),
                Format = _csvBox.Value == CheckState.Checked ? "csv" : "xlsx",
            },
        };
    }

    // ── Actions ────────────────────────────────────────────────────────────────

    private void SaveSettings()
    {
        try
        {
            _settings.Save(ReadConfig());
            AppendLog("Settings saved.");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not save settings: {ex.Message}");
        }
    }

    private async Task GenerateAsync()
    {
        if (_isRunning) return;

        var config = ReadConfig();
        var validationError = ConfigValidator.Validate(config);
        if (validationError != null) { AppendLog($"[ERROR] {validationError}"); return; }

        // Persist config before running so a partial save is better than none
        try { _settings.Save(config); } catch { /* swallow — running is more important */ }

        _isRunning = true;
        SetButtons(false);
        _log.Text = "";
        _progress.Visible = true;

        // Redirect Console.Out/Error through LogTextWriter so service-internal Console writes
        // appear in the TUI log (mirrors the old WPF approach).
        var originalOut   = Console.Out;
        var originalError = Console.Error;
        var logWriter   = new LogTextWriter(AppendLog);
        var errorWriter = new LogTextWriter(line => AppendLog($"[ERROR] {line}"));
        Console.SetOut(logWriter);
        Console.SetError(errorWriter);

        try
        {
            // Run the report on a background task; never block the UI thread.
            await Task.Run(() => ReportRunner.RunAsync(config, s => Console.WriteLine(s)));

            _lastOutputPath = config.Report.OutputFile;
            var exists = File.Exists(_lastOutputPath);
            // Marshal UI update back onto the Terminal.Gui main loop
            Application.Invoke(() => _openBtn.Enabled = exists);
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
            logWriter.Dispose();
            errorWriter.Dispose();
            Application.Invoke(() =>
            {
                _progress.Visible = false;
                SetButtons(true);
                _isRunning = false;
            });
        }
    }

    private void OpenReport()
    {
        if (string.IsNullOrWhiteSpace(_lastOutputPath) || !File.Exists(_lastOutputPath))
        {
            AppendLog("[ERROR] Report file not found.");
            return;
        }
        try { PlatformOpener.Open(_lastOutputPath); }
        catch (Exception ex) { AppendLog($"[ERROR] Could not open report: {ex.Message}"); }
    }

    private void SetButtons(bool enabled)
    {
        _runBtn.Enabled  = enabled;
        _saveBtn.Enabled = enabled;
    }

    /// <summary>Appends a line to the log pane. Always marshaled onto the main loop so it is safe
    /// to call from background tasks.</summary>
    private void AppendLog(string line)
    {
        Application.Invoke(() =>
        {
            _log.Text += line + "\n";
            _log.MoveEnd();
        });
    }
}
