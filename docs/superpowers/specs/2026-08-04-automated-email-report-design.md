# Automated Monthly Report Email — Design

**Date:** 2026-08-04
**Status:** Approved

## Goal

Automate the generation and emailing of the monthly Toggl/Jira report. A daily
Windows scheduled task runs the app headless; the first run on or after the 1st
of a month (i.e. the first day the PC is on) generates the report and emails it
to the configured recipients. No manual steps once configured.

## Decisions (from brainstorming)

- **Runs on the user's Windows PC** via Windows Task Scheduler (daily trigger),
  not in the cloud.
- **Sends on the 1st when possible**: the daily task also runs on the 1st; the
  send-state check makes later days a catch-up, not the norm.
- **Email via the user's own mailbox** over SMTP (credentials in appsettings.json,
  same trust model as the existing Toggl/Jira tokens).
- **Recipients are clients/colleagues** — configurable list, short fixed
  (Danish) text in the body, the .xlsx report attached.
- **Report period:** Jan 1 (of the previous month's year) → last day of the
  previous month. One sheet per month, same shape as manual runs, but the
  just-started month gets no near-empty sheet.
- **Manual xlsx edits are preserved** exactly as in manual runs: the automated
  run writes to the same configured output file, so the existing
  read-previous-file behavior (extra version-report columns, Plan-sheet cells)
  applies unchanged.

## Architecture

### 1. Config & state

New `EmailConfig` section on `AppConfig`, persisted in appsettings.json like the
existing sections:

| Field | Notes |
|---|---|
| `SmtpHost`, `SmtpPort`, `UseSsl` | SMTP server settings |
| `Username`, `Password` | SMTP credentials |
| `FromAddress` | Sender address |
| `Recipients` | Semicolon-separated list |
| `Subject`, `Body` | Support `{month}` and `{year}` placeholders (month name in Danish, e.g. "Timerapport {month} {year}") |

Send-state lives in a **separate** `automation-state.json` next to the exe
(`AppContext.BaseDirectory`), NOT in appsettings.json, so the settings UI can
never clobber it. Content: the last report period sent, e.g.
`{ "LastSentPeriod": "2026-07" }`.

**Decision rule (pure function, unit-tested):** target period = the calendar
month before today. Send iff `LastSentPeriod` is absent or earlier than the
target period. On success, set `LastSentPeriod` to the target period.

**Period computation (pure function, unit-tested):**
`end = first day of current month minus 1 day; start = Jan 1 of end.Year`.
(A run in January therefore covers the whole previous year.)

### 2. EmailService

New `Services/EmailService.cs` using **MailKit** (NuGet). Single responsibility:
send one message (resolved subject/body, recipient list, xlsx attachment) via
the configured SMTP server. `System.Net.Mail.SmtpClient` is not used
(deprecated).

### 3. Headless auto mode

- `App.xaml.cs` inspects startup args: `--auto` → run headless (no window),
  otherwise show `MainWindow` as today.
- The report pipeline currently in `MainWindow.RunReportAsync` moves to a shared
  `ReportRunner` class; both `MainWindow` and auto mode call it. `MainWindow`
  keeps its UI concerns (log textbox redirection, progress bar, dialogs).
- Auto mode:
  1. Load settings; validate report + email config.
  2. Apply decision rule; exit 0 silently if nothing to send.
  3. Compute period, run `ReportRunner` (xlsx format), send email, update state.
  4. Log to `logs\auto.log` (append, in `AppContext.BaseDirectory`); no console.
  5. Cache write-permission prompt: auto-answer "no retry" (skip cache write).
  6. Exit non-zero on failure; state is NOT advanced, so the next daily run
     retries. Transient Toggl/Jira/SMTP failures self-heal with at most a
     one-day delay.

### 4. UI: Email tab + automation toggle

New tab in `MainWindow.xaml` following the existing tab pattern:

- Fields for all `EmailConfig` values (password in a `PasswordBox`), saved via
  the existing Save flow.
- **"Send report by email now"** button: generates the report and sends it
  immediately, ignoring the state file (manual/test path). Runs through the
  same `ReportRunner` + `EmailService` code as auto mode; logs to the log
  textbox as usual.
- **"Enable daily automatic send"** checkbox reflecting actual Task Scheduler
  state (queried at startup). Checking it registers a scheduled task via
  PowerShell `Register-ScheduledTask`: daily trigger at 09:00,
  `-Settings (New-ScheduledTaskSettingsSet -StartWhenAvailable)` so a missed
  run fires as soon as the PC is next on; action = current exe path with
  `--auto`. Unchecking removes the task. Task name: `ToggJirazer AutoReport`.

### 5. Error handling

- UI validation mirrors the existing `ValidateConfig` pattern: missing SMTP
  host/from/recipients produce clear `[ERROR]` log lines before any work runs.
- Auto mode writes errors to `logs\auto.log`; no dialogs, no toasts.
- Scheduled-task registration errors surface in the log textbox.

### 6. Testing

- New small test project (`ToggJirazer.Tests`, xunit) — the repo has none today.
- Unit tests for the pure logic: period computation (incl. January), decision
  rule (never sent / already sent / older period / failure not advancing state),
  and `{month}`/`{year}` placeholder resolution.
- Email sending and task registration are verified manually via the
  "Send now" button and the checkbox.

## Out of scope

- Cloud/CI execution (explicitly rejected in favor of local scheduling).
- Summary numbers in the email body (fixed text chosen).
- Preserving manual edits anywhere other than the existing output file.
