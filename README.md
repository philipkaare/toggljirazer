# toggljirazer
A console application to generate reports based on Jira and Toggl Track.

## Overview

ToggJirazer fetches detailed time entries from the [Toggl Track API](https://developers.track.toggl.com/docs/) for a specific project and time period, cross-references them with [Jira](https://developer.atlassian.com/cloud/jira/platform/rest/v3/) issues (matched by issue key found in the time entry description), and produces a report in CSV or XLSX format. ToggJirazer opens the Windows GUI by default on Windows and runs a console report on macOS/Linux. Use `--tui` for the interactive terminal UI or `--headless` for an explicit non-interactive run.

## Report Columns

| Column | Description |
|---|---|
| Issue Type | Jira issue type (e.g. Story, Bug, Task) |
| Key | Jira issue key (e.g. PROJ-123) |
| Summary | Jira issue summary |
| Budget | Budget field from Jira (customfield_10016) |
| Account | Account field from Jira (customfield_10014) |
| Person | Toggl user name |
| Start Date | Earliest time entry start date for this task/user |
| Time Used (HH:MM) | Total time spent in HH:MM format |
| Time Used (Decimal) | Total time spent as decimal hours |

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- A [Toggl Track](https://track.toggl.com) account with API access
- A [Jira](https://www.atlassian.com/software/jira) account with API access

## Configuration

Edit `ToggJirazer/appsettings.json` before running:

```json
{
  "Toggl": {
    "ApiToken": "your-toggl-api-token",
    "WorkspaceId": 123456,
    "ProjectId": 789012
  },
  "Jira": {
    "BaseUrl": "https://yourcompany.atlassian.net",
    "UserEmail": "your-email@example.com",
    "ApiToken": "your-jira-api-token"
  },
  "Report": {
    "StartDate": "2024-01-01",
    "EndDate": "2024-01-31",
    "OutputFile": "report.csv",
    "Format": "csv"
  }
}
```

- **Toggl:ApiToken** – Found at https://track.toggl.com/profile
- **Toggl:WorkspaceId** – Found in your Toggl workspace settings URL
- **Toggl:ProjectId** – The Toggl project ID to report on
- **Jira:BaseUrl** – Your Jira instance base URL
- **Jira:UserEmail** – Email address for Jira authentication
- **Jira:ApiToken** – Create at https://id.atlassian.com/manage-profile/security/api-tokens
- **Report:StartDate / EndDate** – Optional. Defaults to the current calendar month if not set.
- **Report:OutputFile** – Output file path (default: `report.csv`)
- **Report:Format** – Output format: `csv` (default) or `xlsx`

## How It Works

Time entry descriptions in Toggl must begin with a Jira issue key (e.g. `PROJ-123 - some work done`). The application extracts the key from each entry and looks it up in Jira to enrich the report.

## Building & Running

```bash
# Interactive TUI
dotnet run --project ToggJirazer -- --tui

# Or use the helper script
./run.sh --tui

# Show available options
./run.sh --help

# Non-interactive run (reads appsettings.json, writes the report, exits)
./run.sh --headless
```

The helper script works from any working directory when invoked by its path.
It uses `ToggJirazer/appsettings.json` and writes relative report paths inside
`ToggJirazer/`. Any arguments are forwarded to the application.

The app is a cross-platform .NET 10 console application (Terminal.Gui). It runs
on macOS, Linux, and Windows. In the TUI you can edit the Toggl/Jira/Report
settings, save them, generate the report, and open it in your default
spreadsheet app. Use `--tui` in a terminal to open the TUI. Redirected input/output selects headless mode.
Use `--headless` for automation; the
process exits non-zero on validation or run errors.

## Automatic monthly email

The Windows GUI’s Email tab configures SMTP settings and recipients for sending the report by email.

- **Send Report by Email Now** generates the report with the configured period and emails
  it immediately (requires xlsx format).
- **Send automatically every month** registers a daily Windows scheduled task
  ("ToggJirazer AutoReport", 09:00, runs when the PC is next on if missed) that starts
  the app with `--auto`. On each run the app checks `automation-state.json` next to the
  exe: if the report covering last month (period: Jan 1 → end of previous month) has not
  been sent yet, it is generated and emailed, and the state is updated. Failures are
  logged to `logs\auto.log` and retried on the next daily run.

