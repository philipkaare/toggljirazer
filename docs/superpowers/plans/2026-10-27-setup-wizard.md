# Setup Wizard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a setup wizard that opens on first launch when setup is incomplete and is always available through a "Setup…" button. It covers:
- Toggl sign-in (email and password fetch the API token, or paste a token)
- Toggl organization, workspace and project dropdowns
- Atlassian sign-in and site choice
- Jira Budget and Account field dropdowns
- report output

**Architecture:**
- **Toggl calls:** a testable `TogglSetupService` (injected `HttpMessageHandler`) talks to Toggl API v9.
- **Jira field names:** `JiraService.ListFieldNamesAsync` lists field names over MCP.
- **Wizard logic:** a UI-free `SetupWizardState` holds the steps, per-step validity, the draft config, and the "setup incomplete" rule.
- **Wizard window:** `SetupWizardWindow` (Avalonia) binds to the state. It reuses `AtlassianSignInControl` from project 2 and is shown as a dialog over `MainWindow`.

**Tech Stack:** .NET 10, Avalonia 11.3.22, xUnit, Avalonia.Headless.XUnit.

**Spec:** `docs/superpowers/specs/2026-10-27-setup-wizard-design.md`. Read it, including its "Builder-check call" lines. The project 2 spec (`2026-10-27-jira-mcp-design.md`) describes the MCP tool shapes.

## Global Constraints

- **No network calls in tests.** Toggl gets canned responses through a fake `HttpMessageHandler`. Jira gets canned JSON through an `IJiraTools` fake; see `ToggJirazer.Tests/Fakes/FakeJiraTools.cs`.
- **Never launch the app against real settings.** GUI smoke launches set `TOGGJIRAZER_SETTINGS_DIR` to a fresh temp folder, never redirect stdout (that selects headless mode), and kill only their own PID. Never click a sign-in button.
- **Build output:** build and test with `--artifacts-path <scratch dir>`, because the developer's running app may lock `bin/`.
- **The Toggl password is never stored, logged, cached or put in settings.** It lives only in memory for the single request.
- **Cancel changes nothing:** Cancel or closing the window leaves `appsettings.json` unchanged. Atlassian sign-in and sign-out take effect immediately, because they are session state.
- **Exact strings:**
  - `Wrong email or password. If you sign in to Toggl with Google, Apple or SSO, use an API token instead.`
  - `Use an API token instead` (link), which opens `https://track.toggl.com/profile`
  - `Use a different account` (link)
  - `Signed in as {fullname}`
  - `(none)` for the field dropdowns, stored as an empty string
  - `{name} (current)` for a current field name missing from the list
  - Button labels `Back`, `Next`, `Cancel`, `Finish`, and `Setup…` in the main window
- **Toggl endpoints** (Basic auth): `GET https://api.track.toggl.com/api/v9/me`, `/api/v9/me/organizations`, `/api/v9/me/workspaces`, `/api/v9/workspaces/{wid}/projects?active=true&per_page=200&page={n}`.
- **Commits:** messages end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## Review Focus

1. **An existing non-empty Budget or Account setting (for example "Konto") is missing from the fetched field list.** It stays selected as "Konto (current)" and is never saved as empty. Tested in Task 2 (`Current_field_name_kept_when_not_listed`) and Task 3 (`Field_step_keeps_current_name`).
2. **The user cancels after changing several steps.** The settings file is byte-for-byte unchanged. Tested in Task 3 (`Cancel_leaves_settings_file_unchanged`).
3. **Unsaved edits in the main window's Email tab before opening Setup….** They survive Finish. Tested in Task 4 (`Wizard_draft_starts_from_window_state`).
4. **Toggl returns more than 200 projects.** All pages are fetched. Tested in Task 1 (`Projects_paginate_until_short_page`).
5. **First launch with an empty settings folder.** The wizard opens automatically. Verified in Task 4's smoke screenshot. A complete config plus a token file does not open it (`IsSetupIncomplete` tests in Task 3).

---

### Task 1: TogglSetupService

**Files:** Create `ToggJirazer/Services/TogglSetupService.cs`. Test: `ToggJirazer.Tests/TogglSetupServiceTests.cs`, plus a small `FakeHttpHandler` in `ToggJirazer.Tests/Fakes/` that records requests and returns scripted responses.

**Interfaces (Produces), namespace `ToggJirazer.Services`:**

```csharp
public sealed record TogglAccount(string ApiToken, string FullName, string Email);
public sealed record TogglOrganization(long Id, string Name);
public sealed record TogglWorkspace(long Id, string Name, long OrganizationId);
public sealed record TogglProject(long Id, string Name);
public sealed class TogglSetupException(string message) : Exception(message);

public sealed class TogglSetupService : IDisposable
{
    public TogglSetupService(HttpMessageHandler? handler = null);
    public Task<TogglAccount> SignInAsync(string email, string password, CancellationToken ct = default);   // 401/403 → TogglSetupException(exact wrong-password string)
    public Task<TogglAccount> ValidateTokenAsync(string apiToken, CancellationToken ct = default);            // 401/403 → TogglSetupException("That API token was not accepted by Toggl.")
    public Task<IReadOnlyList<TogglOrganization>> GetOrganizationsAsync(string apiToken, CancellationToken ct = default);
    public Task<IReadOnlyList<TogglWorkspace>> GetWorkspacesAsync(string apiToken, long organizationId, CancellationToken ct = default); // filtered client-side
    public Task<IReadOnlyList<TogglProject>> GetActiveProjectsAsync(string apiToken, long workspaceId, CancellationToken ct = default);  // paginated, sorted by name
}
```

Other non-success statuses throw `TogglSetupException($"Toggl returned {(int)status}: {body-or-reason}")`.

- [ ] **Step 1: Write failing tests.**
  - `SignIn_returns_token_and_name`: `/me` returns `{"api_token":"t1","fullname":"Ann","email":"a@x"}`. Assert the result, and that the request's Basic header decodes to `a@x:pw`.
  - `SignIn_401_throws_exact_message`
  - `ValidateToken_uses_api_token_basic_auth`: the header decodes to `t1:api_token`.
  - `Workspaces_filtered_by_organization`
  - `Projects_paginate_until_short_page`: page 1 has 200 items and page 2 has 3, giving 203 results sorted by name and exactly 2 requests.
  - `Password_not_in_exception_messages`: a 500 whose body echoes the request does not leak the password in the message. The message uses status and reason, and the body only when it doesn't contain the password.
- [ ] **Step 2: Run the tests** with `dotnet test ... --filter FullyQualifiedName~TogglSetupServiceTests`. Expected: compile errors.
- [ ] **Step 3: Implement.** Use one `HttpClient` per service instance, and set auth per request through `HttpRequestMessage.Headers.Authorization`.
- [ ] **Step 4: Run the tests and confirm they pass**, then run the full suite.
- [ ] **Step 5: Commit** with the message `feat(setup): Toggl setup service (token from login, orgs/workspaces/projects)`.

---

### Task 2: Jira field-name listing and optional fields

**Files:** Modify `ToggJirazer/Services/JiraService.cs`. Test: `ToggJirazer.Tests/JiraServiceTests.cs`, plus `ToggJirazer.Tests/FieldChoicesTests.cs`. Create `ToggJirazer/Services/FieldChoices.cs`.

**Interfaces (Produces):**
- **`JiraService.ListFieldNamesAsync()`:** signature `public Task<IReadOnlyList<(string Id, string Name)>> ListFieldNamesAsync(CancellationToken ct = default)`.
  - Search with JQL `updated >= -90d order by updated desc`, `maxResults` 50 and `fields: ["key"]`.
  - For the first key of each distinct project, up to 5 projects, call `getJiraIssue(key, ["*all"], expand "names")`.
  - Merge the `names` maps, de-duplicate by Id, and sort by Name, ignoring case.
  - No issues found gives an empty list.
- **`FieldChoices.Build(...)`:** signature `public static (IReadOnlyList<string> Items, int SelectedIndex) Build(IReadOnlyList<string> siteFieldNames, string currentName, string defaultName)`.
  - Items are `"(none)"` followed by the distinct site names.
  - Selection order:
    - a non-empty `currentName` that is in the list gets selected;
    - a non-empty `currentName` that is not in the list is appended as `"{currentName} (current)"` and selected;
    - an empty `currentName` selects `defaultName` (Budget or Account) if listed, otherwise `(none)`.
- **`FieldChoices.ToSetting(string item)`:** signature `public static string ToSetting(string item)`. `"(none)"` maps to `""`, `"X (current)"` maps to `"X"`, and anything else stays as it is.
- **Empty field names:** `JiraService` treats an empty `BudgetFieldName` or `AccountFieldName` as "skip this field". It requests no ID for it and prints no "not found" warning.

- [ ] **Step 1: Write failing tests.**
  - `ListFieldNames_merges_names_across_projects`: two projects, whose maps overlap in one field.
  - `ListFieldNames_jql_is_bounded`: the JQL contains `updated >=`.
  - `ListFieldNames_empty_when_no_recent_issues`
  - `Empty_field_name_skips_field_without_warning`: capture `Console.Out`. Restore it in `finally`.
  - `FieldChoices` cases:
    - `Current_field_name_kept_when_not_listed`: current "Konto" not in the list gives a selected item `Konto (current)`.
    - Current name listed.
    - Empty current name with the default listed.
    - Empty current name with the default missing, giving `(none)`.
    - `ToSetting` round-trips.
- [ ] **Step 2: Run the tests and confirm they fail.**
- [ ] **Step 3: Implement.**
- [ ] **Step 4: Run the tests and confirm they pass**, then run the full suite.
- [ ] **Step 5: Commit** with the message `feat(setup): list Jira field names and allow skipping Budget/Account`.

---

### Task 3: SetupWizardState

**Files:** Create `ToggJirazer/Setup/SetupWizardState.cs`. Test: `ToggJirazer.Tests/SetupWizardStateTests.cs`.

**Interfaces (Produces), namespace `ToggJirazer.Setup`:**

```csharp
public enum SetupStep { Welcome, TogglSignIn, TogglProject, JiraSignIn, JiraFields, ReportOutput, Done }

public sealed class SetupWizardState
{
    public SetupWizardState(AppConfig startingConfig);     // deep-copies into Draft; startingConfig is never mutated
    public AppConfig Draft { get; }
    public SetupStep Step { get; }
    public bool CanGoBack { get; }                          // false on Welcome
    public bool CanGoNext { get; }                          // per-step validity (below); false on Done
    public void Next();                                     // throws InvalidOperationException if !CanGoNext
    public void Back();
    public bool JiraSignedIn { get; set; }                  // set by the window from the sign-in control
    public static bool IsSetupIncomplete(AppConfig config, bool hasAtlassianToken);
}
```

**Validity rules:**
- **Welcome:** always valid.
- **TogglSignIn:** `Draft.Toggl.ApiToken` is non-empty and not `your-toggl-api-token`.
- **TogglProject:** organization, workspace and project IDs are all non-zero.
- **JiraSignIn:** `JiraSignedIn` is true and `Draft.Jira.CloudId` is non-empty.
- **JiraFields:** always valid, because empty means none.
- **ReportOutput:** `OutputFile` is non-empty and `Format` is `xlsx` or `csv`.

`IsSetupIncomplete` is true when any of these hold: the Toggl token is empty or the placeholder, any Toggl ID is 0, `CloudId` is empty, or `!hasAtlassianToken`.

The window owns saving: Finish saves `Draft` through `SettingsService`. Cancel saves nothing.

- [ ] **Step 1: Write failing tests.**
  - `IsSetupIncomplete` truth-table cases, including the placeholder token and a complete config with a token, which gives false.
  - `Next_disabled_until_toggl_token_present`
  - `Project_step_requires_all_three_ids`
  - `Jira_step_requires_signed_in_and_site`
  - `Back_from_toggl_returns_to_welcome`
  - `Starting_config_not_mutated`
  - `Field_step_keeps_current_name`: a draft with `AccountFieldName` "Konto" stays "Konto" after walking through the steps with no field change.
  - `Cancel_leaves_settings_file_unchanged`: write a settings file, build the state from it, change several Draft values, don't save, and assert the file bytes are identical.
- [ ] **Step 2: Run the tests and confirm they fail.**
- [ ] **Step 3: Implement.** Deep-copy by serialize/deserialize, or by explicit copies.
- [ ] **Step 4: Run the tests and confirm they pass**, then run the full suite.
- [ ] **Step 5: Commit** with the message `feat(setup): wizard state and setup-incomplete rule`.

---

### Task 4: Wizard window, Setup… button, first-run launch, README

**Files:** Create `ToggJirazer/Setup/SetupWizardWindow.axaml(.cs)`. Modify `ToggJirazer/MainWindow.axaml(.cs)`, `ToggJirazer/App.axaml.cs` if needed, and `README.md`. Test: `ToggJirazer.Tests/SetupWizardWindowTests.cs` and `MainWindowTests.cs`.

**Interfaces:**
- **Consumes:**
  - Task 1: `TogglSetupService`
  - Task 2: `JiraService.ListFieldNamesAsync`, `FieldChoices`
  - Task 3: `SetupWizardState`
  - project 2: `AtlassianSignInControl`, with `Initialize`, `SetSelection`, `CloudId`, `SiteUrl`, `RefreshStatusAsync` and `SignInChanged`
  - `AtlassianMcp` with `IsSignedInAsync` and `Tools`
  - `SettingsService`
- **Produces:**
  - **Window constructor:** `public SetupWizardWindow(SetupWizardState state, AtlassianMcp? mcp, Func<TogglSetupService> togglFactory, Func<IJiraTools, string, Task<IReadOnlyList<string>>> listFieldNames)`. The default `listFieldNames` builds a `JiraService` on the given tools and cloud ID and returns the names.
  - **Dialog result:** `ShowDialog<AppConfig?>(owner)` returns the Draft on Finish and null on Cancel or close.
  - **`MainWindow.OpenSetupAsync()`** (internal):
    1. Build the state from `ReadConfigFromUi()`, the window's current state, not the file.
    2. Show the wizard.
    3. On a non-null result, `_settingsService.Save(result)` then `PopulateUiFromConfig(result)`.
    4. Always refresh the Jira tab's sign-in status afterwards.
  - **`internal Func<SetupWizardState, Task<AppConfig?>> ShowWizard`:** a seam for tests. The default shows the real window.
  - **Setup… button:** goes in the button row before Save Settings and calls `OpenSetupAsync`.
  - **First run:** in `MainWindow`'s `Opened` handler, if `SetupWizardState.IsSetupIncomplete(loadedConfig, tokenStoreHasTokens)`, call `OpenSetupAsync()`. Expose `bool hasAtlassianToken` through `AtlassianMcp` or the token store; add a small `public bool HasStoredSignIn` on `AtlassianMcp` if needed.
- **Step contents:** these follow the spec, using the exact strings from Global Constraints.
  - **Toggl sign-in:** email and password boxes with a "Sign in" button. Show the result as "Signed in as {fullname}" plus the "Use a different account" link. The "Use an API token instead" link opens the profile page and shows a paste box with a "Check token" button. Errors appear inline. Set `Draft.Toggl.ApiToken` only on success, and clear the password box after each attempt.
  - **Project step:** load organizations, then workspaces, then projects as the selection changes. Pre-select from Draft and auto-select a single option. Loading or error text appears inline.
  - **Jira step:** host `AtlassianSignInControl`. Update `state.JiraSignedIn` and `Draft.Jira.CloudId`/`SiteUrl` on `SignInChanged` and on site selection.
  - **Fields step:** call `listFieldNames` once per entry to the step, build two dropdowns through `FieldChoices.Build`, and write back through `ToSetting`. If loading fails or the list is empty, show text boxes holding the current values plus a one-line note.
  - **Report step:** output file with a Browse… save picker and an xlsx/csv choice.
  - **Done step:** a short summary.

- [ ] **Step 1: Write failing tests** (`[AvaloniaFact]`, with fakes; no network).
  - `Setup_button_opens_wizard_and_saves_result`: `ShowWizard` returns a modified config. Assert that the file was saved and the window fields updated.
  - `Wizard_cancel_changes_nothing`: `ShowWizard` returns null, and the file bytes are unchanged.
  - `Wizard_draft_starts_from_window_state`: type an unsaved Email subject in the window, then let `ShowWizard` capture the state. The `Draft.Email.Subject` equals the typed value.
  - `First_run_opens_wizard_when_incomplete`: an empty config plus no token, with `ShowWizard` counting calls, gives 1 after `Opened`.
  - `Complete_config_does_not_open_wizard`: a complete config plus a fake token gives 0.
  - `Wizard_window_next_disabled_on_empty_toggl_step`: build `SetupWizardWindow` with an empty draft, advance from Welcome, and assert the Next button is disabled.
- [ ] **Step 2: Run the tests and confirm they fail.**
- [ ] **Step 3: Implement**, and update the README.
  - The "Configuration" section leads with: "On first launch a setup wizard walks you through signing in to Toggl and Atlassian and picking your project, site, Jira fields and report file. Reopen it any time with **Setup…**."
  - Keep the file-location notes.
- [ ] **Step 4: Verify.**
  - Run the full suite. Expected: green.
  - Smoke launch with `TOGGJIRAZER_SETTINGS_DIR` set to a new empty temp folder, without redirecting stdout. The wizard should appear on its own. Screenshot it (the window or the screen region) and Read the screenshot to verify. Kill only that PID.
  - Report the screenshot path.
- [ ] **Step 5: Commit** with the message `feat(setup): setup wizard window, Setup… button and first-run launch`.
