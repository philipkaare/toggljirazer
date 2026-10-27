# toggljirazer
A console application to generate reports based on Jira and Toggl Track.

## Overview

ToggJirazer fetches detailed time entries from the [Toggl Track API](https://developers.track.toggl.com/docs/) for a specific project and time period, cross-references them with [Jira](https://developer.atlassian.com/cloud/jira/platform/rest/v3/) issues (matched by issue key found in the time entry description), and produces a report in CSV or XLSX format. ToggJirazer opens a windowed app (Avalonia) by default on Windows and macOS. Use `--headless` for a scripted, non-interactive run.

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

Settings are stored per user, outside the repository:

- **Windows:** `%APPDATA%\ToggJirazer\appsettings.json`
- **macOS:** `~/Library/Application Support/ToggJirazer/appsettings.json`

The file is created on first launch, either from your existing settings (an older
`appsettings.json` from a previous version) or from the template in
`ToggJirazer/appsettings.json`. Edit the values in the app and click **Save Settings**,
or edit the file directly. The `appsettings.json` in the repository is only a template;
the app never writes to it, so your tokens stay out of git.

The settings file looks like this:

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

On macOS/Linux use the helper script:

```bash
./run.sh              # open the app
./run.sh --headless   # non-interactive run: reads your settings, writes the report, exits
./run.sh --help
```

On Windows use the PowerShell helper script:

```powershell
.\run.ps1             # open the app
.\run.ps1 --headless
.\run.ps1 --help
```

The helper scripts work from any working directory when invoked by their path.
They run from `ToggJirazer/`, so relative report paths are written inside
`ToggJirazer/`. Any arguments are forwarded to the application.

The app is a cross-platform .NET 10 Avalonia application. In the window you can
edit the Toggl/Jira/Report/Email settings, save them, generate the report, and open
it in your default spreadsheet app. Redirected input/output selects headless mode.
Use `--headless` for automation; the process exits non-zero on validation or run errors.

## Email

The Email tab configures SMTP settings and recipients for sending the report by email.

- **Send Report by Email Now** generates the report with the configured period and emails
  it immediately (requires xlsx format).
- `{month}` and `{year}` in the subject and body are replaced with the report month
  (Danish month name) and year.

Automatic monthly sending was removed. If you enabled it earlier, delete the 'ToggJirazer AutoReport' task in Task Scheduler.
