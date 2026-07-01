# ToggJirazer → .NET 10 Console TUI — Design

**Date:** 2026-07-01
**Status:** Approved for planning

## Summary

Convert ToggJirazer from a Windows-only WPF desktop app into a cross-platform
.NET 10 console application with a Terminal.Gui text user interface (TUI),
inspired by the uRobot (`~/git/RoboCoder`) project's TUI stack. The report
generation logic and all API/service code are preserved and reused; only the
presentation layer changes. A non-interactive `--headless` mode is added for
automation.

## Motivation

The project is described as a "console application" but is actually a WPF app
(`OutputType=WinExe`, `net9.0-windows`, `UseWPF=true`). It cannot run on
macOS/Linux. The `Services/` and `Models/` code is already platform-agnostic
(pure JSON, `HttpClient`, `ClosedXML`, `CsvHelper`), and the report pipeline in
`MainWindow.RunReportAsync` is already written against `Console.WriteLine`. This
makes a port to a genuine cross-platform console TUI low-risk and high-value.

## Goals

- Run on macOS, Linux, and Windows from a single `net10.0` build.
- Full feature parity with the WPF UI: edit Toggl/Jira/Report settings in-app,
  Save, Generate report, Open report, progress indication, scrolling log.
- Add a non-interactive `--headless` / `--run` mode for automation/CI.
- Match the uRobot aesthetic: a static ASCII banner and graceful non-TUI
  fallback when the terminal can't host the TUI.

## Non-Goals

- No animated splash screen (static banner only).
- No change to report contents, columns, or the Toggl/Jira integration logic.
- No retention of the WPF project or any Windows-only code paths.
- No live theme switching or other uRobot features beyond the banner + fallback.

## Technology Stack

- **Terminal.Gui 2.4.x** — the only new dependency. Provides layout, widgets,
  colors, and the banner rendering. Spectre.Console is explicitly **not**
  used (headless mode writes plain lines to `Console.Out`; the banner is a
  static ASCII string in a Terminal.Gui view).
- Retained: `ClosedXML` (xlsx), `CsvHelper` (csv), `Microsoft.Extensions.Configuration.*`.
- Target framework: **`net10.0`**, `OutputType=Exe`.

## Architecture

### Project changes

Retarget `ToggJirazer/ToggJirazer.csproj`:

- `TargetFramework`: `net9.0-windows` → `net10.0`.
- Remove `UseWPF`, `EnableWindowsTargeting`; `OutputType`: `WinExe` → `Exe`.
- Add `<PackageReference Include="Terminal.Gui" Version="2.4.*" />`.
- Keep existing `ClosedXML`, `CsvHelper`, and configuration package refs.
- Keep `appsettings.json` content-copy rules.

Remove WPF files: `App.xaml`, `App.xaml.cs`, `MainWindow.xaml`,
`MainWindow.xaml.cs`. Keep everything under `Services/` and `Models/`
unchanged. Keep and reuse `LogTextWriter.cs` (already generic — routes lines to
an `Action<string>`).

### New components

| File | Responsibility |
|---|---|
| `Program.cs` | Real entry point. Parse args, detect TUI vs headless, load config, launch the chosen mode. |
| `Cli/ReportRunner.cs` | The report pipeline extracted verbatim from `MainWindow.RunReportAsync`: period resolution → Toggl fetch → Jira enrich → write xlsx/csv. Takes an `AppConfig` and an `Action<string>` log sink. Single source of truth for both modes. |
| `Cli/ConfigValidator.cs` | The `ValidateConfig` rules moved out of `MainWindow` into a shared static helper returning `string?` (null = valid). |
| `Cli/HeadlessRunner.cs` | Validate config, run `ReportRunner` streaming logs to `Console.Out`, set process exit code. |
| `Cli/PlatformOpener.cs` | Cross-platform "open file with OS default app": `open` (macOS), `xdg-open` (Linux), `Process.Start(UseShellExecute=true)` (Windows). |
| `Tui/Banner.cs` | Static `ToggJirazer` ASCII wordmark + version/tagline string. |
| `Tui/AppTui.cs` | Terminal.Gui screen: builds layout, owns Save/Generate/Open actions, routes logs to the log view via `Application.Invoke`. |

### Data flow

```
appsettings.json ──SettingsService.Load()──> AppConfig
                                                 │
                        ┌────────────────────────┴───────────────────────┐
                        │ TUI mode                                        │ Headless mode
                        ▼                                                 ▼
         AppTui (edit fields) ──Save──> SettingsService.Save()   HeadlessRunner
                        │                                                 │
                    Generate                                              │
                        ▼                                                 ▼
        ConfigValidator.Validate ── error ──> log pane / stderr + exit code
                        │ ok                                              │ ok
                        ▼                                                 ▼
              ReportRunner.Run(config, log) ── same pipeline for both ──> report.xlsx/csv
                        │                                                 │
                   log ──> TextView (Application.Invoke)             log ──> Console.Out
```

## Mode Selection

At startup `Program.cs` chooses a mode:

1. `--help` / `-h` → print usage, exit 0.
2. `--headless` or `--run` flag present → headless.
3. `Console.IsInputRedirected` or `Console.IsOutputRedirected` → headless
   (protects CI, pipes, and non-interactive shells; mirrors uRobot's inline
   fallback).
4. Otherwise → launch the Terminal.Gui TUI.

If Terminal.Gui initialization fails for any reason, print a message and fall
back to headless rather than crashing.

## TUI Screen

Full-screen Terminal.Gui window reproducing the WPF form:

```
┌ ToggJirazer v0.x · Toggl × Jira time reports ────────────────┐
│ [ Toggl ] [ Jira ] [ Report ]        ← TabView               │
│   (Toggl tab)                                                 │
│     API Token:     [•••••••••]   (Secret field)              │
│     Organization:  [________]                                 │
│     Workspace:     [________]                                 │
│     Project:       [________]                                 │
│   (Jira tab)  Base URL / User Email / API Token (secret)      │
│   (Report tab) Start Date / End Date / Output File / Format   │
│                                                               │
│ [ Save (F2) ]  [ Generate (F5) ]  [ Open report ]            │
│ Progress: ▓▓▓▓▓░░░░░  (indeterminate pulse while running)    │
│ ┌ Log ───────────────────────────────────────────────────┐  │
│ │ Report period: 2026-01-01 to 2026-07-01                 │  │
│ │ ...                                                      │  │
│ └─────────────────────────────────────────────────────────┘  │
└───────────────────────────────────────────────────────────────┘  Ctrl+Q quit
```

- **Fields:** `TabView` with Toggl / Jira / Report tabs; a `TextField` per WPF
  field. Secret fields (`TogglApiToken`, `JiraApiToken`) use `Secret = true`.
  Format is a two-option selector (xlsx / csv) reproducing the WPF combo.
- **Populate/read:** on launch, populate fields from `SettingsService.Load()`
  (same defaults as WPF: start-of-year / today for dates). Reading fields back
  into an `AppConfig` mirrors `MainWindow.ReadConfigFromUi` (int parsing,
  trimming, empty-date → null).
- **Save:** persist via `SettingsService.Save()`; log "Settings saved."
- **Generate:** read + `ConfigValidator.Validate`; on error, log it and stop.
  Otherwise auto-save, disable buttons, start the indeterminate progress pulse,
  and run `ReportRunner.Run` on a background `Task`. Logs are routed through
  `LogTextWriter` → `Application.Invoke(() => append to TextView)` (the
  Terminal.Gui equivalent of WPF's `Dispatcher.BeginInvoke`). On completion,
  stop the pulse, re-enable buttons, enable **Open report** if the file exists.
- **Open report:** uses `PlatformOpener` on the last output path; failures are
  logged and non-fatal.
- **Quit:** `Ctrl+Q`.

## Headless Mode

- Loads `appsettings.json` via `SettingsService` (edit the file by hand).
- `ConfigValidator.Validate` failure → write to `Console.Error`, exit non-zero.
- Runs `ReportRunner.Run` with the log sink writing to `Console.Out`.
- Run exception → log to stderr, exit non-zero. Success → exit 0.

## Error Handling

- Config validation errors are surfaced (log pane in TUI, stderr in headless)
  and never crash the app.
- Report-run exceptions are caught per run (as the WPF app did): logged,
  progress reset, TUI stays alive; headless exits non-zero.
- `PlatformOpener` failures are logged, non-fatal.

## Testing

Terminal.Gui views are awkward to unit-test, so the testable logic is isolated
into plain classes and covered by a new `ToggJirazer.Tests` xunit project:

- `ConfigValidator` — each required-field rule and the valid case.
- `PlatformOpener` — selects the correct command/mechanism per `OSPlatform`
  (inject an OS discriminator so the selection is testable without launching a
  process).
- `SettingsService` round-trip — Save then Load returns equivalent config
  (guards the config surface the TUI depends on).

`ReportRunner` is not fully unit-tested (it calls live Toggl/Jira APIs) but is
structured to accept a log sink so its non-network branches (period resolution,
format selection, validation) are exercisable. The TUI wiring itself is verified
by manual launch (`dotnet run`) and driven after implementation.

## Documentation

- Update `README.md`: describe it as a cross-platform console/TUI app; net10
  prerequisite; `dotnet run` launches the TUI; `--headless` for automation;
  cross-platform notes; remove the stale "WPF/Windows" implications.
- Add a `run.sh` convenience script (mirrors uRobot's) that runs
  `dotnet run --project ToggJirazer`.

## Rollout / Migration

Single change set. After implementation: `dotnet build` must succeed on macOS,
`dotnet run` must launch the TUI, and `dotnet run -- --headless` must run the
report against `appsettings.json`. Real Toggl/Jira credentials are required in
`appsettings.json` for an end-to-end report; without them the app still launches
and validation reports the missing fields.
