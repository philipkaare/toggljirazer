# toggljirazer
A console application to generate reports based on Jira and Toggl Track.

## Overview

ToggJirazer fetches detailed time entries from the [Toggl Track API](https://developers.track.toggl.com/docs/) for a specific project and time period, cross-references them with [Jira](https://developer.atlassian.com/cloud/jira/platform/rest/v3/) issues (matched by issue key found in the time entry description), and produces a CSV report.

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

- [.NET 9 SDK](https://dotnet.microsoft.com/download)
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
    "OutputFile": "report.csv"
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
- **Report:OutputFile** – Output CSV file path (default: `report.csv`)

## How It Works

Time entry descriptions in Toggl must begin with a Jira issue key (e.g. `PROJ-123 - some work done`). The application extracts the key from each entry and looks it up in Jira to enrich the report.

## Building & Running

```bash
cd ToggJirazer
dotnet build
dotnet run
```

Or publish a self-contained executable:

```bash
dotnet publish -c Release -r linux-x64 --self-contained
```

