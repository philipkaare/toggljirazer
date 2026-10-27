# Setup Wizard — Design

Date: 2026-10-27
Status: Approved for build (Philip asked for designs to be built without a further review round)

This is project 3 of 3. It depends on project 2 (`2026-10-27-jira-mcp-design.md`) for Atlassian sign-in, the site list, and the `getJiraIssue`-based field-name lookup.

## Intent

New colleagues clone the repo, run `run.ps1` or `run.sh`, and get walked through setup without hunting for tokens or numeric IDs:
- **Toggl:** they sign in with their Toggl email and password, and the app fetches their API token.
- **Toggl IDs:** organization, workspace and project are picked from dropdowns.
- **Jira:** sign-in is the browser flow from project 2.

The same steps stay available later for changing projects, re-signing in, or fixing settings.

## Goals

- **When it opens:** the wizard opens automatically on launch when setup is incomplete. That means any of these:
  - no real Toggl token (empty or `your-toggl-api-token`)
  - `Toggl.OrganizationId`, `WorkspaceId` or `ProjectId` is 0
  - `Jira.CloudId` is empty
  - no Atlassian token file
- **Opening it later:** a **"Setup…"** button in the main window's button row opens it at any time, pre-filled with the current settings.
- **Steps, in order:**
  1. **Welcome:** one short paragraph on what is needed (a Toggl account and an Atlassian account) and what the wizard will do.
  2. **Toggl sign-in:**
     - **Fields:** email and password, plus a **"Sign in"** button. The app calls Toggl, gets the API token, shows "Signed in as <fullname>", and never stores or logs the password.
     - **Already signed in:** the step shows "Signed in as …" with a **"Use a different account"** link.
     - **Token alternative:** "Use an API token instead" opens `https://track.toggl.com/profile` in the browser, then shows a paste box and validates the token. This is for Google, Apple or SSO accounts that have no password.
  3. **Toggl project:** dropdowns for Organization, then Workspace (filtered by organization), then Project (active projects of that workspace). They are pre-selected from current settings, and a single choice is auto-selected.
  4. **Jira:** the project-2 sign-in control (sign in, status, site dropdown).
  5. **Jira fields:** Budget field and Account field dropdowns listing the site's field names.
     - **Pre-selection:** the current setting, otherwise a field named "Budget"/"Account", otherwise "(none)".
     - **"(none)"** is allowed, and the report then leaves that column empty, as today when a field isn't found.
  6. **Report output:** output file with **Browse…** and a format choice (xlsx/csv), matching the Report tab.
  7. **Done:** a summary of the choices. **Finish** saves everything in one go and returns to the main window, with its fields refreshed.
- **Navigation:** **Back**, **Next** and **Cancel** on every step. Next is disabled until the step is valid. Cancel, or closing the window, leaves settings exactly as they were. On the automatic first-run, cancelling just shows the main window.

## Non-goals

- An email step. The Email tab in the main window covers it, and Philip excluded it.
- Report period dates. They stay on the Report tab.
- Replacing the main window's tabs. They stay as the "advanced" view, with the same fields.

## Anti-goals

- The Toggl password is stored anywhere: settings, logs or cache. It may only be held in memory for the single request.
- Cancel changes any setting.
- A wizard that can't be completed by someone without a Toggl password. The API-token path must always be there.

## Decisions made with Philip

- **Toggl:** stays on the direct REST API. The wizard fetches the token from email and password, and offers dropdowns instead of IDs.
- **Steps:** the wizard covers report output and Jira field names in addition to Toggl and Jira, with no email step.
- **Reachable later:** setup steps are available after first use too.

## The how (decided)

- **Toggl calls** live in a new `TogglSetupService` that takes an `HttpMessageHandler` for tests. All calls use Basic auth.
  - **Token from login:** `GET https://api.track.toggl.com/api/v9/me` with `email:password` returns `api_token`, `fullname` and `email`. 401 or 403 shows "Wrong email or password. If you sign in to Toggl with Google, Apple or SSO, use an API token instead." Other failures show the HTTP status and Toggl's message.
  - **Validate a token:** the same endpoint with `<token>:api_token`.
  - **Organizations:** `GET /api/v9/me/organizations` returns `[{id, name}]`.
  - **Workspaces:** `GET /api/v9/me/workspaces` returns `[{id, name, organization_id}]`, filtered client-side by organization.
  - **Projects:** `GET /api/v9/workspaces/{wid}/projects?active=true&per_page=200&page=N`, paginating until a page has fewer than 200 items. It returns `[{id, name, active}]`, sorted by name.
- **Jira fields:** `JiraService` (project 2) gets `ListFieldNamesAsync(cloudId) → IReadOnlyList<(string Id, string Name)>`.
  - It runs JQL `updated >= -90d order by updated desc` with `maxResults` 50 and fields `["key"]`, then `getJiraIssue(<first key>, ["*all"], expand "names")`, and returns the `names` entries sorted by name. Unbounded JQL is rejected by the server, so it must be bounded.
  - If no issue was updated in 90 days, the step shows the current values as free text with a note.
  - *Builder-check call:* bounded JQL of this form (`updated >= -30d order by updated desc`) was verified as accepted by the live server during design.
  - *Builder-check call:* a single issue only lists the custom fields configured for its project. So take the first issue of each distinct project key among the 50 results, up to 5 `getJiraIssue` calls, and merge their `names` maps.
  - *Builder-check call:* **the current non-empty setting is always offered and pre-selected, even if it isn't in the list.** It is shown as "<name> (current)". An existing non-empty setting is never defaulted to "(none)".
  - The settings keep storing field **names**, as today (`BudgetFieldName`, `AccountFieldName`). "(none)" is stored as an empty string, and `JiraService` treats an empty name as "skip this field" without a warning.
- **Wizard logic** sits in a UI-free `SetupWizardState` class:
  - current step, per-step validity, the draft `AppConfig`, and `Next`/`Back`/`CanGoNext`
  - `static bool IsSetupIncomplete(AppConfig config, bool hasAtlassianToken)`

  Unit tests cover it. The Avalonia `SetupWizardWindow` binds to it and is shown with `ShowDialog` over the main window. On first run, it is shown right after the main window opens.
- **Saving:** **Finish** saves through `SettingsService` and then calls `MainWindow.PopulateUiFromConfig(saved)`. Settings without a wizard step, such as Email and the report dates, are carried through unchanged.
  - *Builder-check call:* the wizard's draft starts from the main window's **current** state (`ReadConfigFromUi()`), not the file. Unsaved edits in the tabs therefore survive Finish and are saved with it.
- *Builder-check call:* **Atlassian sign-in and sign-out inside the wizard take effect immediately.** The token file is session state, not a setting. Only `CloudId` and `SiteUrl` sit in the draft until Finish. After the wizard closes, Finish or Cancel, the main window's Jira tab refreshes its signed-in status. The Cancel anti-goal covers `appsettings.json`.
- **"Setup…" button:** goes in the main window's button row, before Save Settings.
- **Docs:** the README "Configuration" section leads with "On first launch a setup wizard walks you through…; reopen it with Setup…", and keeps the file-location notes.

## Constraints

- **No live Toggl or Atlassian calls in tests.** Use canned responses through a fake `HttpMessageHandler` or `IJiraTools`. Philip is away during the build, so no sign-ins can be completed.
- **Never run the app against Philip's real settings.** GUI smoke tests use `TOGGJIRAZER_SETTINGS_DIR`, from project 2, pointing at a throwaway folder. With an empty folder the wizard should appear on launch, which doubles as the first-run smoke test.

## Left to the builder

- Wizard window layout, size and visual style. It should be consistent with the main window's Fluent compact look, using a simple header per step, the content, and a button row.
- Whether the step controls are separate `UserControl`s or one window with panels.
- Exact copy beyond the strings quoted above. Keep it short and plain.

## Done when

- `dotnet test` passes, including new tests for:
  - **`TogglSetupService`:** token from login success, 401 message, token validation, organization/workspace filtering, project pagination
  - **`SetupWizardState`:** `IsSetupIncomplete` cases, Next disabled until valid, Cancel leaves the draft unapplied, Finish carries through non-wizard settings
  - **`ListFieldNamesAsync`:** parsing, using canned JSON
- With an empty `TOGGJIRAZER_SETTINGS_DIR`, the app opens and the wizard appears. Verified by screenshot.
- With a complete throwaway settings file, plus a fake token file if feasible, the wizard does not appear on launch, and "Setup…" opens it pre-filled.
- Not verified: real Toggl sign-in and real Atlassian sign-in, which need Philip. The hand-off says so.
