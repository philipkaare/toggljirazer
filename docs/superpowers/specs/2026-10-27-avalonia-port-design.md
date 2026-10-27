# Port the GUI to Avalonia — Design

Date: 2026-10-27
Status: Draft for approval

This is project 1 of 3 in the "make ToggJirazer easier for others to use" effort:

1. **Port the GUI from WPF to Avalonia** (this document)
2. Switch data fetching to MCP servers (Atlassian's official Jira MCP server; a community Toggl MCP server, since we use legacy Toggl Track)
3. First-run setup wizard (depends on how sign-in works after project 2)

## Intent

Philip wants colleagues to be able to use ToggJirazer, on both Windows and macOS, through **one** windowed app. Today the windowed app is WPF (Windows-only) and macOS users only get a terminal UI. Colleagues will clone the repo and run it with the helper scripts (they will have the .NET 10 SDK).

The point of this project is a single, cross-platform windowed app that projects 2 and 3 can build on. It is a port, not a redesign.

## Goals

- The windowed app runs on Windows and macOS from the same code, built from a single `net10.0` target.
- Launching with no arguments opens the windowed app on every platform.
- Behaviour matches today's WPF window. That covers the Toggl, Jira, Report and Email tabs, Save Settings, Generate Report, Open report, the "Cache requests (1 day)" toggle, the progress bar, the live log pane, Browse for output file, Send Report by Email Now, and the cache-permission "Try Again / Cancel" dialog.
- Headless mode (`--headless` / `--run`, and automatic headless selection when input/output is redirected) keeps working for scripted runs.

## Non-goals

- MCP, sign-in changes and the setup wizard. Those are projects 2 and 3.
- Visual redesign. Keep the same layout and fields. Avalonia's default Fluent look is fine.
- Linux support as a target. It will probably just work with Avalonia, but it isn't tested or promised.
- Packaging and installers (signed .app, MSI and so on). Users clone the repo and run it.

## Anti-goals (would make this a failure even if it technically works)

- Losing a setting or behaviour that works in today's WPF window, other than the removals listed below.
- Breaking the existing `appsettings.json` format, so a colleague's or Philip's existing file stops loading.
- Ending up with two windowed UIs to maintain. WPF must be fully gone.

## Decisions made with Philip

- **UI framework: Avalonia.** Chosen over MAUI because it's closest to WPF and covers Windows and Mac.
- **Remove automatic monthly email entirely, on all platforms.** This removes:
  - the "Send automatically every month" checkbox and its handlers
  - `ScheduledTaskService`, `AutoSendRunner`, `AutomationStateStore`, and the `--auto` launch mode
  - `automation-state.json` handling and the `logs\auto.log` writing tied to it
  - `AutomationLogicTests`, `AutomationStateStoreTests`, and the README section "Automatic monthly email" (rewrite it to describe only Send Now)
- **Remove the terminal UI entirely.** This removes:
  - `Tui/` (`AppTui`, `Banner`), the `--tui` flag, and the `Terminal.Gui` package
  - `TuiLayoutTests`, `TuiSettingsTests`, and `BannerTests`
- **Keep headless mode.**
- **Settings move to a per-user folder outside the repo.**
  - **Location:** `%APPDATA%\ToggJirazer\appsettings.json` on Windows and `~/Library/Application Support/ToggJirazer/appsettings.json` on macOS. Use `Environment.SpecialFolder.ApplicationData`, which maps to the right folder on each OS. If needed, special-case macOS to `~/Library/Application Support`, because .NET maps ApplicationData to `~/.config` there.
  - **Who uses it:** the GUI and headless mode read and write only this file.
  - **The repo file:** `ToggJirazer/appsettings.json` stays a committed **placeholder template** and is never written by the app.
  - **First launch (file missing):** copy existing settings once, in this order:
    1. the newest legacy file in `ToggJirazer/bin/Debug/*-windows/appsettings.json` and `AppContext.BaseDirectory/appsettings.json` whose values differ from the template, so Philip's real settings in `bin\Debug\net10.0-windows\` carry over
    2. otherwise the repo template

    Log which file was copied.
  - **`appsettings.development.json`:** it can still be merged on load if present, but Save must not write values merged in from it back into the user file.
  - Update the README configuration section to point at the new location. This is also where project 3's setup wizard will write.
- **Jira `BudgetFieldName` / `AccountFieldName` are preserved on Save.** WPF currently resets them to "Budget"/"Account". The port carries the loaded values through. No new UI fields.

## Constraints and things that must survive

- **`AutomationLogic.ResolvePlaceholders`** (`{month}` / `{year}` with Danish month names) is used by Send Now. It must stay. Move it somewhere that fits, such as `EmailService` or a small `EmailTemplate` helper, keep its tests, and delete the rest of `AutomationLogic` only if nothing else uses it.
- **Uncommitted work in `MainWindow.xaml` / `MainWindow.xaml.cs`.** It replaces the Start/End date text boxes with date pickers and adds `ParseConfigDate`. This must not be lost. Commit it as-is before the port starts, and carry it into the Avalonia window (Avalonia's `CalendarDatePicker` is the closest equivalent). Dates are still saved as `yyyy-MM-dd`.
- The `appsettings.json` schema stays unchanged. `SettingsService` changes only as needed for the new path, the first-run copy, and not saving development overrides.
- Passwords and tokens stay masked in the UI.
- The log pane keeps its dark monospace look. Use a font that exists on both platforms, for example `Consolas, Menlo, monospace`.
- The window icon still comes from `Assets/toggjirazer.ico`. A `.png` alongside it is fine if Avalonia on macOS needs one.

## My calls (not discussed with Philip — flag if wrong)

- **"Open in Excel" becomes "Open Report"**, because on Mac it may open in Numbers. It uses the existing `Cli/PlatformOpener`.
- **Browse for output file** uses Avalonia's `StorageProvider` save picker, with the same xlsx/csv filter behaviour.
- **Code-behind, not MVVM.** Keep the code-behind structure and the existing naming so the diff stays easy to follow. A ViewModel rewrite isn't needed for this project.
- **`run.sh` and `run.ps1`** drop `--tui` from their comments and examples. `run.sh` opening the GUI on Mac is the new default.
- **`ArgParser`**: `LaunchMode` becomes `{ Gui, Headless, Help }`. The `windowsGuiAvailable` parameter goes away, the GUI is the default, and redirected input/output or `--headless` / `--run` select headless. `--tui` is no longer recognised and falls through to the GUI. `--auto` prints "Automatic monthly email has been removed. Delete the 'ToggJirazer AutoReport' scheduled task." and exits with code 0, without opening a window, so a leftover scheduled task doesn't pop the window up every morning. Update `ArgParserTests` to match.
- **Test project** drops its Windows-conditional target framework and targets `net10.0` only.

## Left to the builder

- Avalonia package versions. Use a current stable 11.x release (`Avalonia`, `Avalonia.Desktop`, `Avalonia.Themes.Fluent`, plus `Avalonia.Fonts.Inter` if needed).
- How to rebuild the cache-permission dialog. A small Avalonia `Window` shown with `ShowDialog` is the obvious route. Note that the callback is invoked from a background thread, so use `Dispatcher.UIThread.InvokeAsync` and wait for the result.
- Whether `Cli/ReportRunner` and `Services/ReportRunner` should be merged. Only do it if the port makes it natural, otherwise leave them alone.
- How to verify on macOS. If no Mac is available, say so plainly. At minimum the app must build for `osx-arm64` (`dotnet build -r osx-arm64`), and someone should run it on a Mac before this is called done.

## Done when

- `dotnet build` and `dotnet test` pass on Windows, and nothing references WPF, `Terminal.Gui`, the TUI, or the automation code any more.
- `run.ps1` opens the Avalonia window on Windows and `run.sh` opens it on macOS. Each of the four tabs loads and saves settings, a report generates and opens, and Send Now sends.
- `run.ps1 --headless` still produces a report.
- On Philip's PC, the first launch copies his settings from `bin\Debug\net10.0-windows\` into `%APPDATA%\ToggJirazer\`, and Save never modifies the tracked repo `appsettings.json`.
- README is updated: the GUI is the default on all platforms, there are no TUI or auto-email sections, and the prerequisites are the same .NET 10 SDK.
