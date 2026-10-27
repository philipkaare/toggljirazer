# Avalonia Port Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the WPF window with an Avalonia window that runs on Windows and macOS. Remove the terminal UI and the automatic monthly email. Move settings to a per-user folder.

**Architecture:** A single `net10.0` exe. `Program.Main` resolves a launch mode: no arguments starts Avalonia's classic desktop lifetime, while `--headless` or redirected I/O runs the existing `HeadlessRunner`. Settings live in a per-user `appsettings.json` that `SettingsLocator` creates on first launch by copying the newest legacy file that holds real settings. The window keeps the WPF code-behind structure and names.

**Tech Stack:** .NET 10, Avalonia 11.3.x (`Avalonia`, `Avalonia.Desktop`, `Avalonia.Themes.Fluent`, `Avalonia.Fonts.Inter`), `Avalonia.Headless.XUnit` for window tests, xUnit, ClosedXML, CsvHelper, MailKit.

**Spec:** `docs/superpowers/specs/2026-10-27-avalonia-port-design.md`

## Global Constraints

- Target `net10.0` only, in both `ToggJirazer.csproj` and `ToggJirazer.Tests.csproj`. No `-windows` TFM, no `UseWPF`, no `#if WINDOWS`.
- Nothing may reference WPF (`System.Windows`), `Terminal.Gui`, `Tui/`, `ScheduledTaskService`, `AutoSendRunner`, `AutomationStateStore`, or `automation-state.json`.
- The `appsettings.json` schema stays unchanged: same sections, keys and defaults.
- The app never writes `ToggJirazer/appsettings.json`, the tracked repo file. It stays a placeholder template.
- User settings path: `%APPDATA%\ToggJirazer\appsettings.json` on Windows and `~/Library/Application Support/ToggJirazer/appsettings.json` on macOS.
- `--auto` prints exactly `Automatic monthly email has been removed. Delete the 'ToggJirazer AutoReport' scheduled task.` and exits 0 without opening a window.
- Dates are saved as `yyyy-MM-dd` (invariant culture).
- Passwords and tokens are masked in the UI (`PasswordChar="●"`).
- Log pane: dark background `#1E1E1E`, foreground `#D4D4D4`, font `Consolas, Menlo, monospace`, size 12.
- Keep the code-behind and existing control names. No MVVM.
- Commit messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

1. **The user settings file already exists, so first-run migration must not run.** It must never overwrite it. Tested in Task 2 (`EnsureUserSettings_does_nothing_when_user_file_exists`).
2. **Philip's real tokens sit in the sibling `appsettings.development.json` next to the repo template, not in the template itself.** The migration must count the merged values as real settings. Tested in Task 2 (`EnsureUserSettings_copies_merged_dev_values_from_candidate`).
3. **A `appsettings.development.json` next to the user file overrides values on load.** Save must not write those overrides back. Tested in Task 2 (`Save_does_not_persist_development_overrides`).
4. **Custom Jira `BudgetFieldName` / `AccountFieldName` survive a Save from the window.** There are no UI fields for them. Tested in Task 3 (`ReadConfigFromUi_preserves_jira_field_names`).
5. **Empty or unparseable Start/End dates in the file load as the defaults** (Jan 1 of this year to today) without crashing the window. Tested in Task 3 (`Populate_handles_empty_and_invalid_dates`).

---

### Task 0: Branch and preserve uncommitted work

**Files:**
- Commit as-is: `ToggJirazer/MainWindow.xaml`, `ToggJirazer/MainWindow.xaml.cs` (the date-picker change, already in the working tree)
- Commit: `docs/superpowers/specs/2026-10-27-avalonia-port-design.md`, `docs/superpowers/plans/2026-10-27-avalonia-port.md`

- [ ] **Step 1: Create the branch.** Run `git checkout -b feat/avalonia-port`. Expected: `Switched to a new branch 'feat/avalonia-port'`.
- [ ] **Step 2: Verify that the date-picker change builds and the tests pass.** Run `dotnet build ToggJirazer.slnx` and `dotnet test ToggJirazer.slnx`. Expected: build succeeds and all tests pass.
- [ ] **Step 3: Commit.** Do not add `ToggJirazer/report.xlsx` or `.claude/`.

```bash
git add ToggJirazer/MainWindow.xaml ToggJirazer/MainWindow.xaml.cs
git commit -m "feat(gui): use date pickers for report period"
git add docs/superpowers/specs/2026-10-27-avalonia-port-design.md docs/superpowers/plans/2026-10-27-avalonia-port.md
git commit -m "docs: Avalonia port design and plan"
```

---

### Task 1: Remove the TUI and the automatic email, and reshape launch modes

**Files:**
- Delete: `ToggJirazer/Tui/AppTui.cs`, `ToggJirazer/Tui/Banner.cs`, `ToggJirazer/Services/AutoSendRunner.cs`, `ToggJirazer/Services/AutomationStateStore.cs`, `ToggJirazer/Services/ScheduledTaskService.cs`, `ToggJirazer/Services/AutomationLogic.cs`
- Delete tests: `ToggJirazer.Tests/TuiLayoutTests.cs`, `TuiSettingsTests.cs`, `BannerTests.cs`, `AutomationStateStoreTests.cs`, `AutomationLogicTests.cs`
- Create: `ToggJirazer/Services/EmailTemplate.cs`, `ToggJirazer.Tests/EmailTemplateTests.cs`
- Modify: `ToggJirazer/Cli/ArgParser.cs`, `ToggJirazer.Tests/ArgParserTests.cs`, `ToggJirazer/Program.cs`, `ToggJirazer/ToggJirazer.csproj` (remove the `Terminal.Gui` package), `ToggJirazer/MainWindow.xaml(.cs)` (remove the `AutoSendCheckBox`, its handlers, `MainWindow_Loaded` and `_suppressAutoSendEvents`; switch to `EmailTemplate`), `ToggJirazer/App.xaml.cs` (remove the `--auto` branch so it always shows `MainWindow`)

**Interfaces:**
- Produces: `public static class EmailTemplate { public static string ResolvePlaceholders(string template, DateTime periodEnd); }` in `ToggJirazer.Services`. It behaves exactly like the old `AutomationLogic.ResolvePlaceholders`: `{month}` is the da-DK `MMMM` name and `{year}` is the invariant-culture year.
- Produces: `public enum LaunchMode { Gui, Headless, Help, AutoRemoved }` and `public static LaunchMode Resolve(string[] args, bool inputRedirected, bool outputRedirected)` in `ToggJirazer.Cli.ArgParser`. The `windowsGuiAvailable` parameter is removed.

- [ ] **Step 1: Write the failing tests.** In `EmailTemplateTests`, move the old `ResolvePlaceholders` assertions from `AutomationLogicTests` and add this one:

```csharp
[Fact]
public void ResolvePlaceholders_uses_danish_month_and_year()
    => Assert.Equal("Timerapport juli 2026",
        EmailTemplate.ResolvePlaceholders("Timerapport {month} {year}", new DateTime(2026, 7, 31)));
```

In `ArgParserTests`, keep the help, headless and redirection tests, drop every `--tui` and `windowsGuiAvailable` case, and add these:

```csharp
[Fact] public void No_args_returns_gui() => Assert.Equal(LaunchMode.Gui, ArgParser.Resolve([], false, false));
[Fact] public void Tui_flag_falls_through_to_gui() => Assert.Equal(LaunchMode.Gui, ArgParser.Resolve(["--tui"], false, false));
[Fact] public void Auto_flag_returns_auto_removed()
{
    Assert.Equal(LaunchMode.AutoRemoved, ArgParser.Resolve(["--auto"], false, false));
    Assert.Equal(LaunchMode.AutoRemoved, ArgParser.Resolve(["--AUTO"], true, true)); // even when redirected (Task Scheduler)
}
[Fact] public void Help_beats_auto() => Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["--auto", "--help"], false, false));
```

- [ ] **Step 2: Run the tests and confirm they fail.** Run `dotnet test ToggJirazer.slnx --filter "EmailTemplateTests|ArgParserTests"`. Expected: compile errors, because `EmailTemplate` and `LaunchMode.AutoRemoved` don't exist yet.
- [ ] **Step 3: Implement the changes.**
  - Create `EmailTemplate`.
  - Reshape `ArgParser` with this precedence: Help, then AutoRemoved (case-insensitive), then Headless (flag or redirected), then Gui.
  - Delete the files listed above and remove `Terminal.Gui` from the csproj.
  - Update `Program`:
    - `AutoRemoved` prints the exact message from Global Constraints and returns 0.
    - The `Gui` case stays `#if WINDOWS` with `new App().Run()` for now. On other platforms it falls back to `HeadlessRunner` (temporary; Task 3 replaces it).
    - Remove `--tui` and `--auto` from `PrintUsage`.
  - Update `MainWindow` to use `EmailTemplate.ResolvePlaceholders`.
- [ ] **Step 4: Verify.** Run `dotnet build ToggJirazer.slnx` and `dotnet test ToggJirazer.slnx`. Expected: success. Then run `grep -rEn "Terminal.Gui|AppTui|AutoSend|AutomationState|ScheduledTask|AutomationLogic" ToggJirazer ToggJirazer.Tests --include=*.cs --include=*.csproj --include=*.xaml`. Expected: no output (ignore `bin/` and `obj/`).
- [ ] **Step 5: Commit** with the message `refactor: remove terminal UI and automatic monthly email`.

---

### Task 2: Per-user settings location and first-run migration

**Files:**
- Create: `ToggJirazer/Services/SettingsLocator.cs`, `ToggJirazer.Tests/SettingsLocatorTests.cs`
- Modify: `ToggJirazer/Services/SettingsService.cs` (`Save` only), `ToggJirazer.Tests/SettingsServiceTests.cs`, `ToggJirazer/Program.cs`

**Interfaces:**
- Produces, in `ToggJirazer.Services`:

```csharp
public static class SettingsLocator
{
    // %APPDATA%\ToggJirazer\appsettings.json (Windows),
    // ~/Library/Application Support/ToggJirazer/appsettings.json (macOS),
    // Environment.SpecialFolder.ApplicationData/ToggJirazer/appsettings.json otherwise.
    public static string UserSettingsPath();

    // Legacy files, in no particular order: cwd/appsettings.json,
    // AppContext.BaseDirectory/appsettings.json, and
    // <BaseDirectory>/../*-windows/appsettings.json (e.g. bin/Debug/net10.0-windows).
    public static IEnumerable<string> LegacyCandidates(string currentDirectory, string baseDirectory);

    // If userPath exists: do nothing, return null.
    // Else: pick the qualifying candidate with the newest write time (max of the file and
    // its sibling appsettings.development.json); qualifying = new SettingsService(c).Load()
    // has a Toggl.ApiToken that is non-empty and not "your-toggl-api-token".
    // Save that loaded (dev-merged) config to userPath via SettingsService. If none qualify,
    // copy templatePath verbatim. Creates the directory. Returns the source path used and
    // calls log($"Settings copied from {source} to {userPath}").
    public static string? EnsureUserSettings(string userPath, IEnumerable<string> candidates,
        string templatePath, Action<string> log);
}
```

- Consumes: `SettingsService(string path)`, `Load()`, `Save(AppConfig)`.
- Produces a change in the behaviour of `SettingsService.Save`. Before writing, it loads the sibling `appsettings.development.json` (the same path rule `Load` uses). For every leaf in the dev file whose value equals the value being saved, it writes the base file's existing value at that path instead. If the base file has no value there, it omits the key. Values the user changed in the UI are still saved.

- [ ] **Step 1: Write the failing tests.** Use temp directories, as the existing `SettingsServiceTests` do.

```csharp
// SettingsLocatorTests
[Fact] public void EnsureUserSettings_does_nothing_when_user_file_exists()
// user file contains {"Toggl":{"ApiToken":"mine"}}; a qualifying candidate exists
// → returns null; user file text unchanged.

[Fact] public void EnsureUserSettings_copies_newest_qualifying_candidate()
// candidate A token "old" (write time T), candidate B token "new" (T+1 min), candidate C template token
// → user file loads ApiToken "new"; return value == B's path.

[Fact] public void EnsureUserSettings_copies_merged_dev_values_from_candidate()
// candidate appsettings.json has token "your-toggl-api-token", sibling appsettings.development.json
// has {"Toggl":{"ApiToken":"real","WorkspaceId":42}} → user file loads ApiToken "real", WorkspaceId 42.

[Fact] public void EnsureUserSettings_falls_back_to_template()
// only non-qualifying candidates → user file text equals template text; returns templatePath.

[Fact] public void EnsureUserSettings_creates_missing_directory()
// userPath in a not-yet-existing subdirectory → file exists afterwards.

[Fact] public void LegacyCandidates_includes_windows_build_output()
// baseDirectory = <tmp>/bin/Debug/net10.0/, with <tmp>/bin/Debug/net10.0-windows/appsettings.json present
// → result contains that path.

// SettingsServiceTests
[Fact] public void Save_does_not_persist_development_overrides()
// base {"Toggl":{"ApiToken":"base"}}, dev {"Toggl":{"ApiToken":"dev"}}; Load() then Save() unchanged
// → base file's Toggl.ApiToken is still "base".

[Fact] public void Save_persists_user_change_to_dev_overridden_value()
// same files; Load(), set ApiToken = "edited", Save() → base file's ApiToken == "edited".
```

- [ ] **Step 2: Run the tests and confirm they fail.** Run `dotnet test ToggJirazer.slnx --filter "SettingsLocatorTests|SettingsServiceTests"`. Expected: compile errors, because `SettingsLocator` is missing, and the two new `SettingsServiceTests` fail.
- [ ] **Step 3: Implement `SettingsLocator` and the `Save` change.** On macOS (`OperatingSystem.IsMacOS()`), build the path from `Environment.SpecialFolder.UserProfile` + `Library/Application Support/ToggJirazer`. Then wire it into `Program`:
  - Replace `ResolveSettingsPath()` with `SettingsLocator.UserSettingsPath()`.
  - Call `EnsureUserSettings(path, LegacyCandidates(Directory.GetCurrentDirectory(), AppContext.BaseDirectory), Path.Combine(AppContext.BaseDirectory, "appsettings.json"), Console.WriteLine)` before the mode switch, except for `Help` and `AutoRemoved`.
  - Pass the path to `MainWindow` through a constructor parameter, `MainWindow(string settingsPath)`, and remove its own `SettingsPath` field. `App` gets the path from a static `App.SettingsPath` property that `Program` sets.
  - Update `PrintUsage` so its configuration text names the user settings path.
- [ ] **Step 4: Verify.** Run `dotnet test ToggJirazer.slnx`. Expected: all pass. Then run `git status --short ToggJirazer/appsettings.json`. Expected: no output.
- [ ] **Step 5: Commit** with the message `feat: store settings per user and migrate legacy settings on first run`.

---

### Task 3: Avalonia window replaces WPF

**Files:**
- Delete: `ToggJirazer/App.xaml`, `App.xaml.cs`, `MainWindow.xaml`, `MainWindow.xaml.cs`
- Create: `ToggJirazer/App.axaml`, `App.axaml.cs`, `MainWindow.axaml`, `MainWindow.axaml.cs`, `ToggJirazer/CacheRetryDialog.axaml(.cs)` (or build it in code, as WPF did), `ToggJirazer.Tests/MainWindowTests.cs`, `ToggJirazer.Tests/TestAppBuilder.cs`
- Modify: `ToggJirazer/ToggJirazer.csproj`, `ToggJirazer.Tests/ToggJirazer.Tests.csproj`, `ToggJirazer/Program.cs`, `ToggJirazer/Cli/ArgParser.cs` (only if anything still mentions Windows)

**Interfaces:**
- Consumes: `SettingsService`, `EmailTemplate.ResolvePlaceholders`, `Services.ReportRunner.ValidateConfig/ResolveConfiguredPeriod/RunAsync`, `EmailService.ValidateEmailConfig/ParseRecipients/SendReportAsync`, `PlatformOpener.Open(string)`, `App.SettingsPath` (Task 2).
- Produces: `public MainWindow(string settingsPath)` with `internal void PopulateUiFromConfig(AppConfig config)` and `internal AppConfig ReadConfigFromUi()`. Make them internal and add `<InternalsVisibleTo Include="ToggJirazer.Tests" />` to the csproj.
- Produces: `public static AppBuilder BuildAvaloniaApp()` on `Program`, which the tests reuse.

**Port mapping (WPF → Avalonia):**

| WPF | Avalonia |
|---|---|
| `PasswordBox` / `.Password` | `TextBox PasswordChar="●"` / `.Text` |
| `DatePicker` | `CalendarDatePicker` (`SelectedDate`) |
| `ToolTip="…"` | `ToolTip.Tip="…"` |
| `Label` styles in `Window.Resources` | `Window.Styles` with `Style Selector="Label"` etc. |
| `SaveFileDialog` | `StorageProvider.SaveFilePickerAsync` with `FilePickerFileType` xlsx/csv plus "All files", `SuggestedFileName`, `SuggestedStartLocation` from the current output directory |
| `Dispatcher.Invoke/BeginInvoke` | `Dispatcher.UIThread.InvokeAsync(...).GetAwaiter().GetResult()` for the cache dialog callback (called from a background thread) / `Dispatcher.UIThread.Post` for `AppendLog` |
| `LogTextBox.ScrollToEnd()` | set `CaretIndex = Text.Length` after appending |
| `Process.Start` with `UseShellExecute` | `PlatformOpener.Open(path)` |

Rename the button to **"Open Report"** (control name `OpenReportButton`, as before). Keep window size 760×720, min 640×600, the four tabs with the same fields and tooltips, the cache checkbox, the progress bar and the log pane. The "Send automatically every month" checkbox was removed in Task 1. Window icon: `Assets/toggjirazer.ico` as an `AvaloniaResource`.

- [ ] **Step 1: Swap the project files.**
  - `ToggJirazer.csproj`:
    - Set `TargetFramework` to `net10.0` only.
    - Remove `UseWPF`, `EnableWindowsTargeting`, the conditional `Compile Remove` and the WPF `Resource` item.
    - Add `OutputType WinExe`, keeping `DisableWinExeOutputInference`, so Windows doesn't open a console window behind the GUI. **Check: headless output must still reach the console.** If `WinExe` breaks `run.ps1 --headless` output, keep `Exe` and note it in the commit.
    - Add `<AvaloniaUseCompiledBindingsByDefault>true</AvaloniaUseCompiledBindingsByDefault>`.
    - Add the Avalonia packages at the same 11.3.x version, `<AvaloniaResource Include="Assets/**" />`, and `InternalsVisibleTo`.
  - `ToggJirazer.Tests.csproj`: set `net10.0` only, remove `EnableWindowsTargeting`, and add `Avalonia.Headless.XUnit` at the same version.
- [ ] **Step 2: Write the failing window tests** in `MainWindowTests`. Use `[AvaloniaFact]`, with `TestAppBuilder` declaring `[assembly: AvaloniaTestApplication(typeof(TestAppBuilder))]` and returning `Program.BuildAvaloniaApp().UseHeadless(new AvaloniaHeadlessPlatformOptions())`. Each test gets a temp settings path.

```csharp
[AvaloniaFact] public void ReadConfigFromUi_round_trips_populated_config()
// Populate with tokens, IDs, BaseUrl, email, StartDate "2026-01-01", EndDate "2026-06-30",
// OutputFile "r.xlsx", Format "csv", all Email fields → ReadConfigFromUi() equals input field-by-field.

[AvaloniaFact] public void ReadConfigFromUi_preserves_jira_field_names()
// Populate with BudgetFieldName "MyBudget", AccountFieldName "MyAccount" → read-back has both.

[AvaloniaFact] public void Populate_handles_empty_and_invalid_dates()
// StartDate null, EndDate "not-a-date" → no exception; read-back StartDate == $"{DateTime.Today.Year}-01-01",
// EndDate == DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture).

[AvaloniaFact] public void Unknown_format_selects_first_item()
// Format "pdf" → read-back Format == "xlsx".
```

- [ ] **Step 3: Run the tests and confirm they fail.** Run `dotnet test ToggJirazer.slnx --filter MainWindowTests`. Expected: compile errors, because the Avalonia `MainWindow` and `BuildAvaloniaApp` don't exist yet.
- [ ] **Step 4: Implement the Avalonia `App`, `MainWindow` and cache dialog** following the mapping table. Port the handlers' logic unchanged, apart from the dispatcher and picker calls.
  - `ReadConfigFromUi` starts from a copy of the last loaded or populated config's `Jira.BudgetFieldName` and `AccountFieldName` (held in a `_loadedConfig` field), so they survive a Save.
  - Keep `ParseConfigDate`.
  - `App.OnFrameworkInitializationCompleted` sets `desktop.MainWindow = new MainWindow(SettingsPath)`.
  - `Program`:
    - `BuildAvaloniaApp() => AppBuilder.Configure<App>().UsePlatformDetect().WithInterFont().LogToTrace()`.
    - The `Gui` case returns `BuildAvaloniaApp().StartWithClassicDesktopLifetime(args)`.
    - Remove `#if WINDOWS` and the temporary non-Windows fallback.
- [ ] **Step 5: Verify the build and tests.** Run `dotnet build ToggJirazer.slnx` and `dotnet test ToggJirazer.slnx`. Expected: all pass. Then run `grep -rn "System.Windows\|UseWPF\|#if WINDOWS" ToggJirazer ToggJirazer.Tests --include=*.cs --include=*.csproj --include=*.axaml`. Expected: no output.
- [ ] **Step 6: Run the manual smoke test on Windows** with `.\run.ps1`. Check each of these:
  - The window opens with Philip's migrated settings, and the console shows the "Settings copied from …" line on the first run only.
  - Each tab shows its values.
  - Save Settings writes `%APPDATA%\ToggJirazer\appsettings.json`.
  - Browse opens a save picker.
  - Generate Report fills the log and progress, then enables Open Report.
  - Open Report opens the file.
  - Send Report by Email Now sends, if SMTP is configured.
  - The window icon shows.

  Report what was and wasn't checked.
- [ ] **Step 7: Verify the macOS build** with `dotnet publish ToggJirazer/ToggJirazer.csproj -r osx-arm64 -c Release --self-contained false`. Expected: success. If no Mac is available, say so plainly in the hand-off. A Mac run is still required before calling the port done.
- [ ] **Step 8: Commit** with the message `feat(gui): port window from WPF to Avalonia`.

---

### Task 4: Documentation and helper scripts

**Files:**
- Modify: `README.md`, `run.ps1`, `run.sh`

- [ ] **Step 1: Update `README.md`.**
  - The overview says the windowed app opens by default on Windows and macOS, and `--headless` is for scripted runs. Remove all TUI and Terminal.Gui text.
  - The Configuration section says settings live in `%APPDATA%\ToggJirazer\appsettings.json` / `~/Library/Application Support/ToggJirazer/appsettings.json`. It also says they are created on first launch from existing settings or the template, can be edited in the window, and that the repo's `appsettings.json` is only a template.
  - Replace "Automatic monthly email" with an "Email" section describing **Send Report by Email Now** and the `{month}`/`{year}` placeholders. Add one line: "Automatic monthly sending was removed. If you enabled it earlier, delete the 'ToggJirazer AutoReport' task in Task Scheduler."
  - In Building & Running, remove the `--tui` examples.
- [ ] **Step 2: Update `run.ps1` and `run.sh`.**
  - Header comments: the scripts run from the project directory so relative report paths land in `ToggJirazer/`, and settings live in the per-user folder.
  - Examples: list only the no-argument GUI, `--headless` and `--help`.
- [ ] **Step 3: Verify.** Run `grep -n "tui\|TUI\|Terminal.Gui\|--auto\|every month" README.md run.ps1 run.sh`. Expected: only the README's "Automatic monthly sending was removed…" line. Run `.\run.ps1 --help`. Expected: usage without `--tui` or `--auto`. Run `.\run.ps1 --headless`. Expected: the report is written, or there's a validation error naming the missing setting.
- [ ] **Step 4: Commit** with the message `docs: describe Avalonia GUI, per-user settings, removed TUI and auto email`.
