# Jira through Atlassian's MCP server — Design

Date: 2026-10-27
Status: Approved for build (Philip asked for designs to be built without a further review round)

This is project 2 of 3 in the "make ToggJirazer easier for others to use" effort. Project 1, the Avalonia port with per-user settings, is merged. Project 3 is the setup wizard (`2026-10-27-setup-wizard-design.md`).

## Intent

Colleagues should sign in to Atlassian in their browser instead of creating, copying and pasting a Jira API token. Atlassian's official remote MCP server (`https://mcp.atlassian.com/v2/mcp`) supports OAuth 2.1 with dynamic client registration, so a desktop app can sign people in without us registering an OAuth app or shipping a client secret. Toggl is out of scope and keeps its direct REST connection.

## Goals

- **Data source:** all Jira data the report uses comes through the Atlassian MCP server's tools.
- **Sign-in:** people sign in with **"Sign in with Atlassian"** in the browser, and the sign-in persists between app launches.
- **Site choice:** users pick their Atlassian site from a dropdown of the sites they can access, instead of typing a URL.
- **Expired sign-in:** before each report run the app checks that the sign-in still works, refreshing it quietly when it can.
  - In the GUI, if the sign-in is expired or revoked, the user is asked to sign in again and the run continues afterwards.
  - Headless mode stops with a clear message.
- **Report contents stay identical**, including fix-version inheritance through the parent chain, the Budget and Account custom fields looked up by name, the version-total queries, and the 1-day cache.

## Non-goals

- Toggl changes. Toggl stays on its direct REST connection.
- The setup wizard (project 3). Project 3 reuses this project's sign-in and site list.
- Confluence or any other Atlassian product.

## Anti-goals (would make this a failure even if it technically works)

- Any report value differs from what the direct-REST version produced for the same data.
- Tokens or the OAuth client registration end up in the repository or in `appsettings.json`.
- The app silently produces a report with missing Jira data when the sign-in has expired. It must ask, or stop.

## Decisions made with Philip

- **MCP is the only way to reach Jira.** There is no API-token fallback, and the direct Jira REST code and its settings (`BaseUrl`, `UserEmail`, `ApiToken`) are removed.
- **Check the sign-in before each run.** If the user is not signed in, ask them to sign in again.
- If a site's admin blocks localhost sign-in redirects, the user can't connect until an admin allows it. Philip expects it to work, and people in the organization can change the setting.

## Facts established during design (verified against the live server, read-only)

- **Site list:** `getAccessibleAtlassianResources` (no arguments) returns a JSON array of `{id, url, name, scopes, avatarUrl}`. The same site appears once per product, so de-duplicate by `id`. `id` is the cloud ID, and every other tool takes it as `cloudId`.
- **Search:** `searchJiraIssuesUsingJql` takes `{cloudId, jql, fields: string[], maxResults (50–100), nextPageToken?, responseContentFormat}`. It returns `{"issues": {"nodes": [<issue>], "pageInfo": {"hasNextPage": bool, "endCursor": string}, "webUrl": ...}, "context": {...}}`.
  - Each `<issue>` has the standard Jira REST shape: `{id, key, fields: {summary, issuetype: {name, ...}, parent: {key, ...}, fixVersions: [...], timeoriginalestimate, customfield_NNNNN: ...}}`.
  - To get the next page, pass `pageInfo.endCursor` as `nextPageToken`.
- **Unbounded JQL is rejected:** "Unbounded JQL queries are not allowed here". Every JQL must contain a restriction such as `key in (...)` or `fixVersion = ...`.
- **Single issue:** `getJiraIssue` takes `{cloudId, issueIdOrKey, fields: string[], expand?}`. It returns the same `{"issues": {"nodes": [<issue>]}}` wrapper.
  - With `fields: ["*all"]` and `expand: "names"`, the node carries a `names` map from field ID to display name. On the real site this gives `customfield_10073 → "Budget"` and `customfield_10111 → "Account"`.
- **Rate limit:** the server reports a limit of 200 per window. Batch with JQL at 100 issues per call, and never make one call per issue in normal operation.
- **Response format:** tool results arrive as MCP text content holding this JSON. Use `responseContentFormat: "adf"` so field values aren't rewritten to markdown.

- **Value formats match today's REST responses** (checked on real data during the builder check):
  - `fixVersions[]` items are `{self, id, description, name, archived, released}`.
  - `timeoriginalestimate` is an integer number of seconds, for example `26640`.
  - The Budget custom field is a select option `{self, value, id}`.
  - Issues without a value give `null`.

  `ExtractStringField` already handles all of these. Fixtures must use exactly these shapes.

## Builder-check calls (made by Claude while settling the reviewer's questions)

- **Rate limit:**
  - **Window length:** unknown, so keep the call count low.
  - **Version queries:** batch them as `fixVersion in ("v1", "v2", …)`, up to 50 versions per JQL, paginated. Group issues by each of their own fix versions on our side. An issue with several matching fix versions counts toward each one, exactly as running one query per version did. A test proves the batched grouping equals per-version results.
  - **When the limit is hit:** a tool error with status 429, or a `rateLimit.remaining` of 0, stops the run with "Atlassian rate limit reached — try again later (the 1-day cache keeps what was fetched)." The run never continues with partial Jira data.
- **Field-ID resolution:** try the first key of each distinct project in the run's key set, up to 5 `getJiraIssue` calls, stopping as soon as both fields are found. If a field is still not found, log a warning exactly as today. This matches today's behaviour for a field that doesn't exist.
- **Auth errors stop the run:** an authorization failure in the middle of a run stops it, even in code paths that warn and carry on for other errors.
- **Headless is run by hand only:** the automatic monthly email was removed in project 1, so there's no scheduled headless run. DPAPI CurrentUser is therefore fine.
  - **Concurrent refresh:** the token cache re-reads the file in every `GetTokensAsync` and writes atomically (temp file plus rename), so a GUI and a manual headless run don't overwrite each other's refreshed tokens.
  - **No browser outside an explicit sign-in:** headless runs and the startup status check never open a browser. The callback handler is only active during an explicit sign-in, and otherwise fails fast as "not signed in".

## The how (decided)

- **Package:** `ModelContextProtocol.Core`, the client-only package. Use the newest stable 2.x (2.3.0 at design time) and connect with `HttpClientTransport` over Streamable HTTP to `https://mcp.atlassian.com/v2/mcp`.
- **OAuth:** use the SDK's `ClientOAuthOptions`.
  - **No client ID:** leave `ClientId` unset so dynamic client registration is used. Once registration succeeds, persist the registered client ID and secret so we don't register on every launch.
  - **Redirect URI:** a fixed `http://localhost:33418/callback`. It must be fixed because the dynamic registration is tied to it. If the port is busy, show an error that names it.
  - **Authorization callback:** implement `AuthorizationCallbackHandler`, not the obsolete delegate. It starts a one-shot `HttpListener` on the redirect URI, opens the authorization URL in the default browser (reuse the `PlatformOpener` logic so it works on macOS too), and waits for the callback with a 5-minute timeout.
    - It returns `code`, `state` and `iss` from the query string, and serves a small "Signed in — you can close this tab" page.
    - Cancel and timeout surface as a clear "Sign-in was cancelled or timed out" error.
  - **Token cache:** implement `ITokenCache` backed by `atlassian-auth.json` in the per-user settings folder, next to `appsettings.json`. The file holds the token container and the client registration. On Windows it is encrypted with DPAPI (`ProtectedData`, CurrentUser scope). On macOS it is written with file mode 600.
- **Settings:** `JiraConfig` becomes `{ CloudId, SiteUrl, BudgetFieldName, AccountFieldName }`.
  - **Legacy migration:** when `SettingsService.Load` finds a legacy `Jira.BaseUrl` and no `CloudId`, it sets `SiteUrl` from it. The site dropdown preselects the matching site after sign-in, and its `id` becomes `CloudId`.
  - **Removed keys:** `Save` no longer writes `BaseUrl`, `UserEmail` or `ApiToken`, so the old Jira token disappears from the user's file on the next save. That is intended.
- **Window:** the Jira tab replaces the Base URL, User Email and API Token rows with:
  - a **"Sign in with Atlassian"** button, which reads **"Sign out"** when signed in
  - a status line, either "Signed in" plus the site name or "Not signed in"
  - a **Site** dropdown filled from `getAccessibleAtlassianResources` after sign-in

  The rows are a reusable `UserControl` so project 3 can embed them. "Sign out" deletes `atlassian-auth.json`.
- **Sign-in check:** before a run, call `getAccessibleAtlassianResources`. The SDK refreshes the token automatically when it can.
  - **GUI, not signed in:** if the call fails because of authorization, or there is no token file, show "Your Atlassian sign-in has expired. Sign in again?" with **Sign in** and **Cancel**. Sign in runs the browser flow and the report then continues. Cancel stops the run with a log line.
  - **Headless, not signed in:** exit code 2 with "Not signed in to Atlassian. Open ToggJirazer and use 'Sign in with Atlassian' on the Jira tab."
  - **On app start:** the Jira tab status runs the same check in the background, without prompting.
- **`JiraService`:** keeps its public methods so `ReportService` doesn't change. Internally it talks to an `IJiraTools` abstraction, essentially `CallToolAsync(name, args) → JsonElement` with the wrapper already parsed. Tests feed it canned JSON, so tests never touch the network.
  - **Field IDs:** call `getJiraIssue(firstKey, ["*all"], expand "names")` once per run and match `BudgetFieldName` and `AccountFieldName` against `names`, ignoring case. This replaces `rest/api/3/field`.
  - **Bulk fetch:** use JQL `key in (K1, K2, …)` in chunks of 100, requesting `summary, issuetype, fixVersions, timeoriginalestimate, parent` plus the two custom field IDs.
    - **Missing keys:** if Jira rejects a chunk because a key doesn't exist or isn't visible, retry that chunk with the offending keys removed. Parse them from the error text, and fall back to splitting the chunk in half. Report missing keys exactly as today.
  - **Parent chain:** batched by JQL, level by level, as today with bulkfetch.
  - **Version queries:** JQL `fixVersion = "<v>"`, paginated with `nextPageToken`.
- **Cache:** the 1-day cache works at the tool-call level. The key is the tool name plus canonical JSON arguments, and the value is the parsed result. It is stored through the existing cache directory, expiry and permission-prompt behaviour. The HTTP-level `HttpResponseCache` handler is not put on the MCP transport, because the JSON-RPC request IDs would defeat it. Toggl keeps the HTTP-level cache.
- **Validation:** both validators (`ReportRunner.ValidateConfig` and `Cli/ConfigValidator`) require a non-empty `CloudId`, with the message "Pick your Atlassian site (Jira tab)." They stop checking the URL, email and token.
- **Docs:** README Jira prerequisites and configuration move to "Sign in with Atlassian". It explains the admin allowlist case: "If sign-in shows a blocked consent screen, ask your Atlassian organization admin to allow `localhost` redirects for the Rovo MCP server."

## Constraints

- **No live calls in tests.** The build runs while Philip is away, so nobody can complete a browser sign-in. All Jira behaviour is proven with canned JSON in the shapes above. Fixtures are synthetic, use made-up keys like `ABC-1`, and contain no data from the real site.
- **Never run the app against Philip's real settings.** That means neither the GUI nor `--headless`, because a run writes the report to a shared drive. Smoke tests that need to launch the GUI do so only with a throwaway settings file. Do this by adding an environment override `TOGGJIRAZER_SETTINGS_DIR` that `SettingsLocator.UserSettingsPath()` honours. It is a test seam, documented in the README under a developer note.
- **Report output unchanged:** existing tests (`BudgetInheritanceTests` and others) keep passing.

## Left to the builder

- The exact SDK type names and wiring, such as `McpClient.CreateAsync` and the transport options. Follow the SDK version's docs.
- How to persist the dynamic client registration with the SDK. Use its registration options or callbacks if it has them. Otherwise cache the issued client ID and secret and pass them as `ClientId`/`ClientSecret` on later launches.
- Exact error-text parsing for missing keys. The halving fallback must guarantee progress.
- Layout details of the Jira tab rows.

## Done when

- `dotnet test` passes, including new tests for:
  - **Parsing:** search and issue wrapper parsing, and pagination through `endCursor`
  - **Field IDs:** resolution from `names`
  - **Missing keys:** chunk retry
  - **Settings:** migration from legacy `BaseUrl` to `SiteUrl`, and Save no longer writing the token
  - **Token cache:** round-trip, with DPAPI on Windows
  - **Sign-in check:** behaviour in headless mode (exit code 2 and the message)
- No code references `rest/api/3`, `Jira.ApiToken`, `UserEmail` or `BaseUrl`, except the legacy read in `SettingsService.Load`.
- The GUI launches with a throwaway settings directory and the Jira tab shows "Not signed in" with the sign-in button. Not verified: an actual sign-in, which needs Philip in front of a browser. The hand-off says so plainly.
