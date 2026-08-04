# Automated Monthly Report Email — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A daily Windows scheduled task runs the app headless (`--auto`); the first run on/after the 1st of a month generates the report (Jan 1 → end of previous month) and emails it via SMTP to configured recipients, tracked by a state file so it sends exactly once per month.

**Architecture:** The report pipeline moves from `MainWindow` into a shared static `ReportRunner`. New services: `AutomationLogic` (pure decision/period/placeholder functions, unit-tested), `AutomationStateStore` (json state file), `EmailService` (MailKit SMTP), `ScheduledTaskService` (PowerShell `Register-ScheduledTask`), `AutoSendRunner` (headless orchestration). `App.xaml.cs` dispatches on `--auto`. A new Email tab in the UI configures SMTP/recipients and offers "send now" + an automation toggle.

**Tech Stack:** .NET 9 WPF (net9.0-windows), MailKit 4.8.0, xunit (new test project), Windows Task Scheduler via PowerShell.

**Spec:** `docs/superpowers/specs/2026-08-04-automated-email-report-design.md`

## Global Constraints

- **NEVER run the application** — no `dotnet run`, and absolutely never with `--auto`. The bin folder's appsettings.json may contain the user's real API tokens and SMTP credentials; running the app can call real APIs and **send real email to the user's clients**. Verify with `dotnet build` and `dotnet test` only.
- The working tree has pre-existing uncommitted changes in `ToggJirazer/Models/Leverance.cs`, `ToggJirazer/Services/JiraService.cs`, `ToggJirazer/Services/ReportService.cs`, `ToggJirazer/Services/TogglService.cs`. **Never stage these.** Always `git add` only the exact files your task created/modified. Never use `git add -A`, `git add .`, or `git commit -a`.
- Follow existing code style: file-scoped namespaces, `ImplicitUsings` (no `using System;` etc. needed for System/Linq/Threading.Tasks), XML doc comments on classes, `string.Empty` over `""`.
- All user-facing email default text is Danish; UI labels are English (matches existing UI).
- Build command: `dotnet build ToggJirazer/ToggJirazer.csproj -v q --nologo` — must end with `0 Error(s)`.
- Test command: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo` (from Task 1 onward).
- There is no .sln file; build/test the csproj files directly.

---

### Task 1: Test project + AutomationLogic (pure functions)

**Files:**
- Create: `ToggJirazer.Tests/ToggJirazer.Tests.csproj`
- Create: `ToggJirazer.Tests/AutomationLogicTests.cs`
- Create: `ToggJirazer/Services/AutomationLogic.cs`

**Interfaces:**
- Consumes: nothing.
- Produces (used by Tasks 7 and 8):
  - `static (DateTime Start, DateTime End) AutomationLogic.ComputeReportPeriod(DateTime today)`
  - `static string AutomationLogic.GetTargetPeriod(DateTime today)` — returns `"yyyy-MM"` of the month before `today`
  - `static bool AutomationLogic.ShouldSend(string? lastSentPeriod, DateTime today)`
  - `static string AutomationLogic.ResolvePlaceholders(string template, DateTime periodEnd)`

- [ ] **Step 1: Create the test project**

Write `ToggJirazer.Tests/ToggJirazer.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net9.0-windows</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <UseWPF>true</UseWPF>
    <EnableWindowsTargeting>true</EnableWindowsTargeting>
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.12.0" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\ToggJirazer\ToggJirazer.csproj" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the failing tests**

Write `ToggJirazer.Tests/AutomationLogicTests.cs`:

```csharp
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class AutomationLogicTests
{
    // --- ComputeReportPeriod -------------------------------------------------

    [Fact]
    public void ComputeReportPeriod_MidYear_CoversJan1ToEndOfPreviousMonth()
    {
        var (start, end) = AutomationLogic.ComputeReportPeriod(new DateTime(2026, 8, 4));
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 7, 31), end);
    }

    [Fact]
    public void ComputeReportPeriod_OnTheFirst_StillCoversPreviousMonth()
    {
        var (start, end) = AutomationLogic.ComputeReportPeriod(new DateTime(2026, 9, 1));
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 8, 31), end);
    }

    [Fact]
    public void ComputeReportPeriod_InJanuary_CoversWholePreviousYear()
    {
        var (start, end) = AutomationLogic.ComputeReportPeriod(new DateTime(2027, 1, 2));
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 12, 31), end);
    }

    // --- GetTargetPeriod -----------------------------------------------------

    [Fact]
    public void GetTargetPeriod_MidYear_IsPreviousMonth()
    {
        Assert.Equal("2026-07", AutomationLogic.GetTargetPeriod(new DateTime(2026, 8, 4)));
    }

    [Fact]
    public void GetTargetPeriod_InJanuary_IsDecemberOfPreviousYear()
    {
        Assert.Equal("2026-12", AutomationLogic.GetTargetPeriod(new DateTime(2027, 1, 15)));
    }

    // --- ShouldSend ----------------------------------------------------------

    [Fact]
    public void ShouldSend_NeverSent_ReturnsTrue()
    {
        Assert.True(AutomationLogic.ShouldSend(null, new DateTime(2026, 8, 4)));
        Assert.True(AutomationLogic.ShouldSend(string.Empty, new DateTime(2026, 8, 4)));
    }

    [Fact]
    public void ShouldSend_AlreadySentForTargetPeriod_ReturnsFalse()
    {
        Assert.False(AutomationLogic.ShouldSend("2026-07", new DateTime(2026, 8, 4)));
        Assert.False(AutomationLogic.ShouldSend("2026-07", new DateTime(2026, 8, 31)));
    }

    [Fact]
    public void ShouldSend_SentForOlderPeriod_ReturnsTrue()
    {
        Assert.True(AutomationLogic.ShouldSend("2026-06", new DateTime(2026, 8, 4)));
        Assert.True(AutomationLogic.ShouldSend("2025-12", new DateTime(2026, 8, 4)));
    }

    [Fact]
    public void ShouldSend_NewMonthStarts_ReturnsTrueAgain()
    {
        // Sent for July; on Sep 1 the target is August, so it must send again.
        Assert.True(AutomationLogic.ShouldSend("2026-07", new DateTime(2026, 9, 1)));
    }

    // --- ResolvePlaceholders -------------------------------------------------

    [Fact]
    public void ResolvePlaceholders_ReplacesMonthAndYearInDanish()
    {
        var result = AutomationLogic.ResolvePlaceholders(
            "Timerapport {month} {year}", new DateTime(2026, 7, 31));
        Assert.Equal("Timerapport juli 2026", result);
    }

    [Fact]
    public void ResolvePlaceholders_NoPlaceholders_ReturnsTemplateUnchanged()
    {
        Assert.Equal("Hej", AutomationLogic.ResolvePlaceholders("Hej", new DateTime(2026, 7, 31)));
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo`
Expected: build FAILS with `The type or namespace name 'AutomationLogic' does not exist` (compile error is the failing state for a new type).

- [ ] **Step 4: Implement AutomationLogic**

Write `ToggJirazer/Services/AutomationLogic.cs`:

```csharp
using System.Globalization;

namespace ToggJirazer.Services;

/// <summary>
/// Pure decision logic for the automated monthly report email:
/// which period to report on, whether a send is due, and placeholder
/// resolution for the email subject/body.
/// </summary>
public static class AutomationLogic
{
    private static readonly CultureInfo Danish = CultureInfo.GetCultureInfo("da-DK");

    /// <summary>
    /// The automated report period: Jan 1 of the previous month's year through
    /// the last day of the previous month. A run in January therefore covers
    /// the whole previous year.
    /// </summary>
    public static (DateTime Start, DateTime End) ComputeReportPeriod(DateTime today)
    {
        var end = new DateTime(today.Year, today.Month, 1).AddDays(-1);
        var start = new DateTime(end.Year, 1, 1);
        return (start, end);
    }

    /// <summary>The period key ("yyyy-MM") a run on <paramref name="today"/> should have sent.</summary>
    public static string GetTargetPeriod(DateTime today)
    {
        var (_, end) = ComputeReportPeriod(today);
        return end.ToString("yyyy-MM", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// True when the report for the target period has not been sent yet.
    /// "yyyy-MM" keys compare correctly with ordinal string comparison.
    /// </summary>
    public static bool ShouldSend(string? lastSentPeriod, DateTime today)
    {
        if (string.IsNullOrWhiteSpace(lastSentPeriod)) return true;
        return string.CompareOrdinal(lastSentPeriod, GetTargetPeriod(today)) < 0;
    }

    /// <summary>Replaces {month} (Danish month name) and {year} using the period end date.</summary>
    public static string ResolvePlaceholders(string template, DateTime periodEnd)
    {
        return template
            .Replace("{month}", periodEnd.ToString("MMMM", Danish))
            .Replace("{year}", periodEnd.Year.ToString(CultureInfo.InvariantCulture));
    }
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo`
Expected: all 11 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add ToggJirazer.Tests/ToggJirazer.Tests.csproj ToggJirazer.Tests/AutomationLogicTests.cs ToggJirazer/Services/AutomationLogic.cs
git commit -m "feat: add test project and automation period/send-decision logic"
```

---

### Task 2: EmailConfig model + settings persistence

**Files:**
- Modify: `ToggJirazer/Models/AppConfig.cs`
- Modify: `ToggJirazer/Services/SettingsService.cs`
- Create: `ToggJirazer.Tests/SettingsServiceTests.cs`

**Interfaces:**
- Consumes: nothing.
- Produces (used by Tasks 5, 7, 8): `EmailConfig` class with properties
  `string SmtpHost`, `int SmtpPort` (default 587), `bool UseSsl` (default false),
  `string Username`, `string Password`, `string FromAddress`, `string Recipients`
  (semicolon-separated), `string Subject`, `string Body`; and
  `AppConfig.Email` of type `EmailConfig`.

- [ ] **Step 1: Write the failing test**

Write `ToggJirazer.Tests/SettingsServiceTests.cs`:

```csharp
using System.IO;
using ToggJirazer.Models;
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class SettingsServiceTests : IDisposable
{
    private readonly string _dir;
    private readonly string _settingsPath;

    public SettingsServiceTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "ToggJirazerTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
        _settingsPath = Path.Combine(_dir, "appsettings.json");
    }

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void EmailConfig_RoundTripsThroughSaveAndLoad()
    {
        var service = new SettingsService(_settingsPath);
        var config = new AppConfig
        {
            Email = new EmailConfig
            {
                SmtpHost = "smtp.example.com",
                SmtpPort = 465,
                UseSsl = true,
                Username = "philip@example.com",
                Password = "secret",
                FromAddress = "philip@example.com",
                Recipients = "a@example.com; b@example.com",
                Subject = "Timerapport {month} {year}",
                Body = "Hej,\n\nVedhæftet timerapport."
            }
        };

        service.Save(config);
        var loaded = service.Load();

        Assert.Equal("smtp.example.com", loaded.Email.SmtpHost);
        Assert.Equal(465, loaded.Email.SmtpPort);
        Assert.True(loaded.Email.UseSsl);
        Assert.Equal("philip@example.com", loaded.Email.Username);
        Assert.Equal("secret", loaded.Email.Password);
        Assert.Equal("philip@example.com", loaded.Email.FromAddress);
        Assert.Equal("a@example.com; b@example.com", loaded.Email.Recipients);
        Assert.Equal("Timerapport {month} {year}", loaded.Email.Subject);
        Assert.Equal("Hej,\n\nVedhæftet timerapport.", loaded.Email.Body);
    }

    [Fact]
    public void Load_FileWithoutEmailSection_ReturnsDefaults()
    {
        File.WriteAllText(_settingsPath, "{}");
        var loaded = new SettingsService(_settingsPath).Load();

        Assert.Equal(string.Empty, loaded.Email.SmtpHost);
        Assert.Equal(587, loaded.Email.SmtpPort);
        Assert.False(loaded.Email.UseSsl);
        Assert.Equal("Timerapport {month} {year}", loaded.Email.Subject);
        Assert.Contains("{month}", loaded.Email.Body);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo`
Expected: build FAILS with `'AppConfig' does not contain a definition for 'Email'`.

- [ ] **Step 3: Add EmailConfig to the models**

In `ToggJirazer/Models/AppConfig.cs`, add to the `AppConfig` class (after the `Report` property):

```csharp
    public EmailConfig Email { get; set; } = new();
```

And append this class at the end of the file (after `ReportConfig`):

```csharp
public class EmailConfig
{
    public string SmtpHost { get; set; } = string.Empty;
    public int SmtpPort { get; set; } = 587;

    /// <summary>True = implicit SSL (typically port 465); false = STARTTLS (typically port 587).</summary>
    public bool UseSsl { get; set; }

    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string FromAddress { get; set; } = string.Empty;

    /// <summary>Semicolon-separated list of recipient addresses.</summary>
    public string Recipients { get; set; } = string.Empty;

    /// <summary>Supports {month} (Danish month name) and {year} placeholders.</summary>
    public string Subject { get; set; } = "Timerapport {month} {year}";

    /// <summary>Supports {month} (Danish month name) and {year} placeholders.</summary>
    public string Body { get; set; } =
        "Hej,\n\nVedhæftet timerapport for {month} {year}.\n\nMed venlig hilsen";
}
```

- [ ] **Step 4: Extend SettingsService.Load**

In `ToggJirazer/Services/SettingsService.cs`, inside the `return new AppConfig { ... }` object initializer in `Load()`, add after the `Report = new ReportConfig { ... }` entry (add a comma after it):

```csharp
                Email = new EmailConfig
                {
                    SmtpHost = node["Email"]?["SmtpHost"]?.GetValue<string>() ?? string.Empty,
                    SmtpPort = node["Email"]?["SmtpPort"]?.GetValue<int>() ?? 587,
                    UseSsl = node["Email"]?["UseSsl"]?.GetValue<bool>() ?? false,
                    Username = node["Email"]?["Username"]?.GetValue<string>() ?? string.Empty,
                    Password = node["Email"]?["Password"]?.GetValue<string>() ?? string.Empty,
                    FromAddress = node["Email"]?["FromAddress"]?.GetValue<string>() ?? string.Empty,
                    Recipients = node["Email"]?["Recipients"]?.GetValue<string>() ?? string.Empty,
                    Subject = node["Email"]?["Subject"]?.GetValue<string>() ?? new EmailConfig().Subject,
                    Body = node["Email"]?["Body"]?.GetValue<string>() ?? new EmailConfig().Body
                }
```

- [ ] **Step 5: Extend SettingsService.Save**

In `Save()`, add to the `node` JsonObject after the `["Report"]` entry (add a comma after it):

```csharp
            ["Email"] = new JsonObject
            {
                ["SmtpHost"] = config.Email.SmtpHost,
                ["SmtpPort"] = config.Email.SmtpPort,
                ["UseSsl"] = config.Email.UseSsl,
                ["Username"] = config.Email.Username,
                ["Password"] = config.Email.Password,
                ["FromAddress"] = config.Email.FromAddress,
                ["Recipients"] = config.Email.Recipients,
                ["Subject"] = config.Email.Subject,
                ["Body"] = config.Email.Body
            }
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo`
Expected: all tests PASS (11 from Task 1 + 2 new).

- [ ] **Step 7: Commit**

```bash
git add ToggJirazer/Models/AppConfig.cs ToggJirazer/Services/SettingsService.cs ToggJirazer.Tests/SettingsServiceTests.cs
git commit -m "feat: add EmailConfig section to settings"
```

---

### Task 3: AutomationStateStore

**Files:**
- Create: `ToggJirazer/Services/AutomationStateStore.cs`
- Create: `ToggJirazer.Tests/AutomationStateStoreTests.cs`

**Interfaces:**
- Consumes: nothing.
- Produces (used by Task 7):
  - `class AutomationState { string? LastSentPeriod }`
  - `class AutomationStateStore` with ctor `AutomationStateStore(string directory)`
    (file is `<directory>\automation-state.json`), `AutomationState Load()`,
    `void Save(AutomationState state)`.

- [ ] **Step 1: Write the failing tests**

Write `ToggJirazer.Tests/AutomationStateStoreTests.cs`:

```csharp
using System.IO;
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class AutomationStateStoreTests : IDisposable
{
    private readonly string _dir;

    public AutomationStateStoreTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "ToggJirazerTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
    }

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void Load_NoFile_ReturnsEmptyState()
    {
        var state = new AutomationStateStore(_dir).Load();
        Assert.Null(state.LastSentPeriod);
    }

    [Fact]
    public void SaveThenLoad_RoundTripsLastSentPeriod()
    {
        var store = new AutomationStateStore(_dir);
        store.Save(new AutomationState { LastSentPeriod = "2026-07" });

        var loaded = store.Load();
        Assert.Equal("2026-07", loaded.LastSentPeriod);
        Assert.True(File.Exists(Path.Combine(_dir, "automation-state.json")));
    }

    [Fact]
    public void Load_CorruptFile_ReturnsEmptyState()
    {
        File.WriteAllText(Path.Combine(_dir, "automation-state.json"), "not json {{{");
        var state = new AutomationStateStore(_dir).Load();
        Assert.Null(state.LastSentPeriod);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo`
Expected: build FAILS with `The type or namespace name 'AutomationStateStore' does not exist`.

- [ ] **Step 3: Implement AutomationStateStore**

Write `ToggJirazer/Services/AutomationStateStore.cs`:

```csharp
using System.IO;
using System.Text.Json;

namespace ToggJirazer.Services;

/// <summary>State persisted between automatic runs (which period was last emailed).</summary>
public class AutomationState
{
    /// <summary>Last report period sent, as "yyyy-MM" (the previous month at send time).</summary>
    public string? LastSentPeriod { get; set; }
}

/// <summary>
/// Reads/writes automation-state.json. Kept separate from appsettings.json so the
/// settings UI can never overwrite the send-state.
/// </summary>
public class AutomationStateStore
{
    private readonly string _path;

    public AutomationStateStore(string directory)
        => _path = Path.Combine(directory, "automation-state.json");

    public AutomationState Load()
    {
        if (!File.Exists(_path)) return new AutomationState();
        try
        {
            return JsonSerializer.Deserialize<AutomationState>(File.ReadAllText(_path))
                   ?? new AutomationState();
        }
        catch
        {
            return new AutomationState();
        }
    }

    public void Save(AutomationState state)
    {
        var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_path, json);
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo`
Expected: all tests PASS.

- [ ] **Step 5: Commit**

```bash
git add ToggJirazer/Services/AutomationStateStore.cs ToggJirazer.Tests/AutomationStateStoreTests.cs
git commit -m "feat: add automation send-state store"
```

---

### Task 4: Extract ReportRunner from MainWindow

**Files:**
- Create: `ToggJirazer/Services/ReportRunner.cs`
- Modify: `ToggJirazer/MainWindow.xaml.cs`

**Interfaces:**
- Consumes: existing `TogglService`, `JiraService`, `ReportService`, `HttpResponseCache`, `AppConfig`.
- Produces (used by Tasks 7 and 8):
  - `static string? ReportRunner.ValidateConfig(AppConfig config)` — returns error message or null (moved verbatim from `MainWindow.ValidateConfig`)
  - `static (DateTime Start, DateTime End) ReportRunner.ResolveConfiguredPeriod(ReportConfig report)` — parses StartDate/EndDate, defaults to Jan 1 → today; throws `InvalidOperationException` on unparsable dates
  - `static Task ReportRunner.RunAsync(AppConfig config, DateTime startDate, DateTime endDate, bool cacheEnabled, Func<Exception, bool> cacheWritePermissionPrompt)` — the full generate-and-write pipeline

This is a pure refactor: no behavior change. The existing `MainWindow.RunReportAsync` (lines ~241-346) and `MainWindow.ValidateConfig` (lines ~88-107) move out; the date-defaulting logic moves into `ResolveConfiguredPeriod`.

- [ ] **Step 1: Create ReportRunner**

Write `ToggJirazer/Services/ReportRunner.cs`. The body of `RunAsync` is the existing `MainWindow.RunReportAsync` with the period-resolution block removed (dates are now parameters); `ValidateConfig` is moved unchanged:

```csharp
using System.IO;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

/// <summary>
/// The report pipeline (Toggl fetch → Jira lookup → report build → file write),
/// shared by the UI "Generate Report" button and the headless --auto mode.
/// Progress is written to Console.Out/Error; callers redirect those as needed.
/// </summary>
public static class ReportRunner
{
    public static string? ValidateConfig(AppConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.Toggl.ApiToken))
            return "Toggl API Token is required (Toggl tab).";
        if (config.Toggl.OrganizationId == 0)
            return "Toggl Organization ID is required (Toggl tab).";
        if (config.Toggl.WorkspaceId == 0)
            return "Toggl Workspace ID is required (Toggl tab).";
        if (config.Toggl.ProjectId == 0)
            return "Toggl Project ID is required (Toggl tab).";
        if (string.IsNullOrWhiteSpace(config.Jira.BaseUrl))
            return "Jira Base URL is required (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Jira.UserEmail))
            return "Jira User Email is required (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Jira.ApiToken))
            return "Jira API Token is required (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Report.OutputFile))
            return "Output File path is required (Report tab).";
        return null;
    }

    /// <summary>
    /// Resolves the period from the Report settings: both dates set → parse them
    /// (throws on invalid values); otherwise Jan 1 of the current year → today.
    /// </summary>
    public static (DateTime Start, DateTime End) ResolveConfiguredPeriod(ReportConfig report)
    {
        if (!string.IsNullOrWhiteSpace(report.StartDate) &&
            !string.IsNullOrWhiteSpace(report.EndDate))
        {
            if (!DateTime.TryParse(report.StartDate, out var start))
                throw new InvalidOperationException(
                    $"'Start Date' value '{report.StartDate}' is not a valid date. Use yyyy-MM-dd.");
            if (!DateTime.TryParse(report.EndDate, out var end))
                throw new InvalidOperationException(
                    $"'End Date' value '{report.EndDate}' is not a valid date. Use yyyy-MM-dd.");
            return (start, end);
        }

        var now = DateTime.Today;
        Console.WriteLine($"No period specified — defaulting to Jan 1 to today: {new DateTime(now.Year, 1, 1):yyyy-MM-dd} to {now:yyyy-MM-dd}");
        return (new DateTime(now.Year, 1, 1), now);
    }

    public static async Task RunAsync(
        AppConfig appConfig,
        DateTime startDate,
        DateTime endDate,
        bool cacheEnabled,
        Func<Exception, bool> cacheWritePermissionPrompt)
    {
        Console.WriteLine($"Report period: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
        Console.WriteLine($"Toggl workspace: {appConfig.Toggl.WorkspaceId}, project: {appConfig.Toggl.ProjectId}");
        Console.WriteLine($"Jira base URL: {appConfig.Jira.BaseUrl}");

        // Set up the request cache (shared by both services). Disabled when the user unchecks it.
        HttpResponseCache? cache = null;
        if (cacheEnabled)
        {
            var cacheDir = Path.Combine(AppContext.BaseDirectory, "cache");
            cache = new HttpResponseCache(cacheDir, TimeSpan.FromDays(1), cacheWritePermissionPrompt);
            Console.WriteLine($"Request cache enabled (folder: {cacheDir}, expires after 1 day).");
        }
        else
        {
            Console.WriteLine("Request cache disabled.");
        }
        Console.WriteLine();

        // Fetch Toggl entries for the period
        using var togglService = new TogglService(appConfig.Toggl, cache);
        var entries = await togglService.GetDetailedReportAsync(startDate, endDate);

        if (entries.Count == 0)
        {
            Console.WriteLine("No Toggl time entries found for the specified period and project.");
            return;
        }

        Console.WriteLine();

        // Fetch all-time entries for version totals
        Console.WriteLine("Fetching all-time Toggl entries for version totals...");
        var allEntries = await togglService.GetAllEntriesAsync();
        Console.WriteLine();

        // Build the report
        using var jiraService = new JiraService(appConfig.Jira, cache);
        var reportService = new ReportService(jiraService);
        var (reportSheets, versionRows, leverances, categoryConsumedHours, weeklyCategoryConsumedHours, monthlyCategoryConsumedHours, noBudgetRows) =
            await reportService.BuildReportAsync(entries, allEntries);

        Console.WriteLine();
        Console.WriteLine($"Report contains {reportSheets.Sum(s => s.Rows.Count)} rows.");
        Console.WriteLine($"Version report contains {versionRows.Count} rows.");
        Console.WriteLine($"Leverances: {leverances.Count}.");
        Console.WriteLine();

        // Write output
        var format = appConfig.Report.Format?.Trim().ToLowerInvariant();
        if (format == "xlsx")
        {
            var extraColumns = reportService.ReadVersionReportExtraColumns(appConfig.Report.OutputFile);
            reportService.WriteXlsx(
                reportSheets,
                versionRows,
                leverances,
                categoryConsumedHours,
                weeklyCategoryConsumedHours,
                monthlyCategoryConsumedHours,
                appConfig.Report.OutputFile,
                extraColumns,
                noBudgetRows);
        }
        else
        {
            if (!string.IsNullOrEmpty(format) && format != "csv")
                Console.Error.WriteLine($"Warning: Unrecognized format '{appConfig.Report.Format}'. Defaulting to CSV.");

            var rows = reportSheets.SelectMany(s => s.Rows).ToList();
            reportService.WriteCsv(rows, appConfig.Report.OutputFile);

            var outputPath = appConfig.Report.OutputFile;
            var versionOutputPath = Path.Combine(
                Path.GetDirectoryName(outputPath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputPath) + "_versions" + Path.GetExtension(outputPath));
            var extraColumns = reportService.ReadVersionReportExtraColumns(versionOutputPath);
            reportService.WriteVersionCsv(versionRows, versionOutputPath, extraColumns);
        }
    }
}
```

**Note:** the `WriteXlsx` call above was copied from the current `MainWindow.RunReportAsync`. If the user's uncommitted changes to `ReportService.cs` have altered that method's signature, adapt the call to match the current `ReportService` — copy the exact call from the current `MainWindow.RunReportAsync` rather than from this plan.

- [ ] **Step 2: Update MainWindow to use ReportRunner**

In `ToggJirazer/MainWindow.xaml.cs`:

1. Delete the entire `private static async Task RunReportAsync(...)` method (the "Report generation" region) and the `private static string? ValidateConfig(AppConfig config)` method.
2. In `RunButton_Click`, replace `var error = ValidateConfig(config);` with:

```csharp
        var error = ReportRunner.ValidateConfig(config);
```

3. In `RunButton_Click`, replace the line
   `await Task.Run(async () => await RunReportAsync(config, cacheEnabled, ShowCacheWritePermissionDialog));`
   with:

```csharp
            var (startDate, endDate) = ReportRunner.ResolveConfiguredPeriod(config.Report);
            await Task.Run(async () => await ReportRunner.RunAsync(
                config, startDate, endDate, cacheEnabled, ShowCacheWritePermissionDialog));
```

(`ResolveConfiguredPeriod` throws `InvalidOperationException` for bad dates; the existing `catch` in `RunButton_Click` already logs it. Its `Console.WriteLine` goes to the log textbox because the redirection is already active at this point.)

- [ ] **Step 3: Build and run tests**

Run: `dotnet build ToggJirazer/ToggJirazer.csproj -v q --nologo` — expected `0 Error(s)`.
Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo` — expected all PASS.

- [ ] **Step 4: Commit**

```bash
git add ToggJirazer/Services/ReportRunner.cs ToggJirazer/MainWindow.xaml.cs
git commit -m "refactor: extract report pipeline into ReportRunner"
```

---

### Task 5: EmailService (MailKit)

**Files:**
- Modify: `ToggJirazer/ToggJirazer.csproj`
- Create: `ToggJirazer/Services/EmailService.cs`

**Interfaces:**
- Consumes: `EmailConfig` (Task 2).
- Produces (used by Tasks 7 and 8):
  - `static string? EmailService.ValidateEmailConfig(EmailConfig config)` — error message or null
  - `static List<string> EmailService.ParseRecipients(string recipients)`
  - `static Task EmailService.SendReportAsync(EmailConfig config, string subject, string body, string attachmentPath)`

- [ ] **Step 1: Add the MailKit package**

In `ToggJirazer/ToggJirazer.csproj`, add to the existing `<ItemGroup>` with `PackageReference` entries:

```xml
    <PackageReference Include="MailKit" Version="4.8.0" />
```

(If restore fails on this version, use the latest available 4.x.)

- [ ] **Step 2: Implement EmailService**

Write `ToggJirazer/Services/EmailService.cs`:

```csharp
using System.IO;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

/// <summary>Sends the generated report as an email attachment via SMTP (MailKit).</summary>
public static class EmailService
{
    public static string? ValidateEmailConfig(EmailConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.SmtpHost))
            return "SMTP Host is required (Email tab).";
        if (config.SmtpPort <= 0)
            return "SMTP Port must be a positive number (Email tab).";
        if (string.IsNullOrWhiteSpace(config.FromAddress))
            return "From Address is required (Email tab).";
        if (ParseRecipients(config.Recipients).Count == 0)
            return "At least one recipient is required (Email tab).";
        return null;
    }

    public static List<string> ParseRecipients(string recipients) =>
        (recipients ?? string.Empty)
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();

    public static async Task SendReportAsync(EmailConfig config, string subject, string body, string attachmentPath)
    {
        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException($"Report file not found: {attachmentPath}");

        var message = new MimeMessage();
        message.From.Add(MailboxAddress.Parse(config.FromAddress));
        foreach (var recipient in ParseRecipients(config.Recipients))
            message.To.Add(MailboxAddress.Parse(recipient));
        message.Subject = subject;

        var builder = new BodyBuilder { TextBody = body };
        await builder.Attachments.AddAsync(attachmentPath);
        message.Body = builder.ToMessageBody();

        using var client = new SmtpClient();
        var socketOptions = config.UseSsl
            ? SecureSocketOptions.SslOnConnect
            : SecureSocketOptions.StartTlsWhenAvailable;
        await client.ConnectAsync(config.SmtpHost, config.SmtpPort, socketOptions);
        if (!string.IsNullOrWhiteSpace(config.Username))
            await client.AuthenticateAsync(config.Username, config.Password);
        await client.SendAsync(message);
        await client.DisconnectAsync(true);
    }
}
```

- [ ] **Step 3: Build and run tests**

Run: `dotnet build ToggJirazer/ToggJirazer.csproj -v q --nologo` — expected `0 Error(s)`.
Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo` — expected all PASS.
(No unit tests for SMTP itself — network I/O; the user verifies manually with the "send now" button.)

- [ ] **Step 4: Commit**

```bash
git add ToggJirazer/ToggJirazer.csproj ToggJirazer/Services/EmailService.cs
git commit -m "feat: add SMTP email service using MailKit"
```

---

### Task 6: ScheduledTaskService

**Files:**
- Create: `ToggJirazer/Services/ScheduledTaskService.cs`

**Interfaces:**
- Consumes: nothing.
- Produces (used by Task 8):
  - `const string ScheduledTaskService.TaskName` = `"ToggJirazer AutoReport"`
  - `static bool ScheduledTaskService.IsEnabled()`
  - `static void ScheduledTaskService.Enable()` — throws `InvalidOperationException` on failure
  - `static void ScheduledTaskService.Disable()` — throws `InvalidOperationException` on failure

- [ ] **Step 1: Implement ScheduledTaskService**

Write `ToggJirazer/Services/ScheduledTaskService.cs`:

```csharp
using System.Diagnostics;
using System.IO;

namespace ToggJirazer.Services;

/// <summary>
/// Registers/removes the daily Windows scheduled task that runs the app with --auto.
/// Uses PowerShell's Register-ScheduledTask because schtasks.exe cannot set
/// StartWhenAvailable (run as soon as possible after a missed start).
/// </summary>
public static class ScheduledTaskService
{
    public const string TaskName = "ToggJirazer AutoReport";

    public static bool IsEnabled()
    {
        var exitCode = RunPowerShell(
            $"if (Get-ScheduledTask -TaskName '{TaskName}' -ErrorAction SilentlyContinue) {{ exit 0 }} else {{ exit 1 }}",
            out _);
        return exitCode == 0;
    }

    public static void Enable()
    {
        var exePath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Could not determine the application executable path.");
        var workingDir = Path.GetDirectoryName(exePath) ?? AppContext.BaseDirectory;

        var command =
            $"$action = New-ScheduledTaskAction -Execute '{Escape(exePath)}' -Argument '--auto' -WorkingDirectory '{Escape(workingDir)}'; " +
            "$trigger = New-ScheduledTaskTrigger -Daily -At 09:00; " +
            "$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable; " +
            $"Register-ScheduledTask -TaskName '{TaskName}' -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null";

        var exitCode = RunPowerShell(command, out var error);
        if (exitCode != 0)
            throw new InvalidOperationException($"Could not register the scheduled task: {error}");
    }

    public static void Disable()
    {
        var exitCode = RunPowerShell(
            $"Unregister-ScheduledTask -TaskName '{TaskName}' -Confirm:$false",
            out var error);
        if (exitCode != 0)
            throw new InvalidOperationException($"Could not remove the scheduled task: {error}");
    }

    private static string Escape(string value) => value.Replace("'", "''");

    private static int RunPowerShell(string command, out string error)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = "-NoProfile -NonInteractive -Command \"" + command.Replace("\"", "\\\"") + "\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardError = true
        };

        using var process = Process.Start(psi)
            ?? throw new InvalidOperationException("Could not start PowerShell.");
        error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return process.ExitCode;
    }
}
```

- [ ] **Step 2: Build and run tests**

Run: `dotnet build ToggJirazer/ToggJirazer.csproj -v q --nologo` — expected `0 Error(s)`.
Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo` — expected all PASS.
(Do NOT invoke `Enable()`/`Disable()` in tests or manually — they mutate the user's Task Scheduler.)

- [ ] **Step 3: Commit**

```bash
git add ToggJirazer/Services/ScheduledTaskService.cs
git commit -m "feat: add scheduled task registration service"
```

---

### Task 7: Headless --auto mode

**Files:**
- Create: `ToggJirazer/Services/AutoSendRunner.cs`
- Modify: `ToggJirazer/App.xaml`
- Modify: `ToggJirazer/App.xaml.cs`

**Interfaces:**
- Consumes: `AutomationLogic` (Task 1), `SettingsService`/`EmailConfig` (Task 2), `AutomationStateStore` (Task 3), `ReportRunner` (Task 4), `EmailService` (Task 5), existing `LogTextWriter`.
- Produces: `static Task<int> AutoSendRunner.RunAsync(string baseDirectory)` — exit code 0 (sent or nothing due) / 1 (failure). `App` now shows `MainWindow` explicitly in normal mode.

- [ ] **Step 1: Implement AutoSendRunner**

Write `ToggJirazer/Services/AutoSendRunner.cs`:

```csharp
using System.IO;

namespace ToggJirazer.Services;

/// <summary>
/// Headless flow behind the --auto flag: decide whether this month's report is due,
/// generate it, email it, and advance the send-state. Logs to logs\auto.log.
/// On failure the state is NOT advanced, so the next daily run retries.
/// </summary>
public static class AutoSendRunner
{
    public static async Task<int> RunAsync(string baseDirectory)
    {
        var logDir = Path.Combine(baseDirectory, "logs");
        Directory.CreateDirectory(logDir);
        using var logFile = new StreamWriter(Path.Combine(logDir, "auto.log"), append: true) { AutoFlush = true };
        void Log(string line) => logFile.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {line}");

        // Route the services' Console progress output into the log file.
        var originalOut = Console.Out;
        var originalError = Console.Error;
        using var outWriter = new LogTextWriter(Log);
        using var errorWriter = new LogTextWriter(line => Log($"[ERROR] {line}"));
        Console.SetOut(outWriter);
        Console.SetError(errorWriter);

        try
        {
            var today = DateTime.Today;
            var stateStore = new AutomationStateStore(baseDirectory);
            var state = stateStore.Load();

            // Already sent this month's report — exit silently so the daily run doesn't bloat the log.
            if (!AutomationLogic.ShouldSend(state.LastSentPeriod, today))
                return 0;

            Log($"Starting automatic report run (last sent: {state.LastSentPeriod ?? "never"}).");

            var config = new SettingsService(Path.Combine(baseDirectory, "appsettings.json")).Load();

            var error = ReportRunner.ValidateConfig(config) ?? EmailService.ValidateEmailConfig(config.Email);
            if (error == null && !string.Equals(config.Report.Format?.Trim(), "xlsx", StringComparison.OrdinalIgnoreCase))
                error = "Automatic email sending requires the xlsx report format.";
            if (error != null)
            {
                Log($"[ERROR] {error}");
                return 1;
            }

            var (startDate, endDate) = AutomationLogic.ComputeReportPeriod(today);
            Log($"Generating report for {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}.");

            await ReportRunner.RunAsync(config, startDate, endDate,
                cacheEnabled: true, cacheWritePermissionPrompt: _ => false);

            var subject = AutomationLogic.ResolvePlaceholders(config.Email.Subject, endDate);
            var body = AutomationLogic.ResolvePlaceholders(config.Email.Body, endDate);
            await EmailService.SendReportAsync(config.Email, subject, body, config.Report.OutputFile);

            state.LastSentPeriod = AutomationLogic.GetTargetPeriod(today);
            stateStore.Save(state);
            Log($"Report emailed to {string.Join(", ", EmailService.ParseRecipients(config.Email.Recipients))}; state advanced to {state.LastSentPeriod}.");
            return 0;
        }
        catch (Exception ex)
        {
            Log($"[ERROR] {ex}");
            return 1;
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }
    }
}
```

- [ ] **Step 2: Remove StartupUri from App.xaml**

`ToggJirazer/App.xaml` becomes (only the `StartupUri` attribute is removed):

```xml
<Application x:Class="ToggJirazer.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Application.Resources>
    </Application.Resources>
</Application>
```

- [ ] **Step 3: Dispatch on --auto in App.xaml.cs**

`ToggJirazer/App.xaml.cs` becomes:

```csharp
using System.Windows;
using ToggJirazer.Services;

namespace ToggJirazer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (e.Args.Any(a => a.Equals("--auto", StringComparison.OrdinalIgnoreCase)))
        {
            // Headless mode: no window; exit code signals success/failure.
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            RunAutoAndShutdown();
        }
        else
        {
            new MainWindow().Show();
        }
    }

    private async void RunAutoAndShutdown()
    {
        int exitCode;
        try
        {
            exitCode = await AutoSendRunner.RunAsync(AppContext.BaseDirectory);
        }
        catch
        {
            exitCode = 1;
        }
        Shutdown(exitCode);
    }
}
```

- [ ] **Step 4: Build and run tests**

Run: `dotnet build ToggJirazer/ToggJirazer.csproj -v q --nologo` — expected `0 Error(s)`.
Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo` — expected all PASS.
**Do NOT run the app to test --auto** (see Global Constraints — it would send real email).

- [ ] **Step 5: Commit**

```bash
git add ToggJirazer/Services/AutoSendRunner.cs ToggJirazer/App.xaml ToggJirazer/App.xaml.cs
git commit -m "feat: add headless --auto mode that emails the monthly report once"
```

---

### Task 8: Email tab, send-now button, automation toggle, README

**Files:**
- Modify: `ToggJirazer/MainWindow.xaml`
- Modify: `ToggJirazer/MainWindow.xaml.cs`
- Modify: `README.md`

**Interfaces:**
- Consumes: `EmailConfig` (Task 2), `ReportRunner` (Task 4), `EmailService` (Task 5), `ScheduledTaskService` (Task 6), `AutomationLogic` (Task 1).
- Produces: UI only — nothing downstream.

- [ ] **Step 1: Add the Email tab to MainWindow.xaml**

In `ToggJirazer/MainWindow.xaml`, after the closing `</TabItem>` of the "Report settings" tab (before `</TabControl>`), insert:

```xml
            <!-- Email settings -->
            <TabItem Header="Email">
                <Grid Margin="10,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="160"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Label Grid.Row="0" Grid.Column="0" Content="SMTP Host:"/>
                    <TextBox Grid.Row="0" Grid.Column="1" x:Name="EmailSmtpHost"/>

                    <Label Grid.Row="1" Grid.Column="0" Content="SMTP Port:"/>
                    <TextBox Grid.Row="1" Grid.Column="1" x:Name="EmailSmtpPort"
                             ToolTip="587 for STARTTLS, 465 for implicit SSL"/>

                    <Label Grid.Row="2" Grid.Column="0" Content="Implicit SSL:"/>
                    <CheckBox Grid.Row="2" Grid.Column="1" x:Name="EmailUseSsl"
                              VerticalAlignment="Center" Margin="0,3"
                              ToolTip="Check for implicit SSL (port 465); uncheck for STARTTLS (port 587)."/>

                    <Label Grid.Row="3" Grid.Column="0" Content="Username:"/>
                    <TextBox Grid.Row="3" Grid.Column="1" x:Name="EmailUsername"/>

                    <Label Grid.Row="4" Grid.Column="0" Content="Password:"/>
                    <PasswordBox Grid.Row="4" Grid.Column="1" x:Name="EmailPassword" PasswordChar="●"/>

                    <Label Grid.Row="5" Grid.Column="0" Content="From Address:"/>
                    <TextBox Grid.Row="5" Grid.Column="1" x:Name="EmailFromAddress"/>

                    <Label Grid.Row="6" Grid.Column="0" Content="Recipients:"/>
                    <TextBox Grid.Row="6" Grid.Column="1" x:Name="EmailRecipients"
                             ToolTip="Separate multiple addresses with semicolons"/>

                    <Label Grid.Row="7" Grid.Column="0" Content="Subject:"/>
                    <TextBox Grid.Row="7" Grid.Column="1" x:Name="EmailSubject"
                             ToolTip="{}{month} and {year} are replaced with the report month (Danish) and year"/>

                    <Label Grid.Row="8" Grid.Column="0" Content="Body:" VerticalAlignment="Top" Margin="0,6,4,0"/>
                    <TextBox Grid.Row="8" Grid.Column="1" x:Name="EmailBody"
                             Height="70" TextWrapping="Wrap" AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"
                             ToolTip="{}{month} and {year} are replaced with the report month (Danish) and year"/>

                    <StackPanel Grid.Row="9" Grid.Column="1" Orientation="Horizontal" Margin="0,8,0,0">
                        <Button x:Name="SendEmailNowButton" Content="Send Report by Email Now"
                                Margin="0,0,16,0" Click="SendEmailNow_Click"/>
                        <CheckBox x:Name="AutoSendCheckBox" Content="Send automatically every month"
                                  VerticalAlignment="Center"
                                  Checked="AutoSend_Checked" Unchecked="AutoSend_Unchecked"
                                  ToolTip="Registers a daily 09:00 scheduled task; the report is emailed once per month, on the first day this PC is on."/>
                    </StackPanel>
                </Grid>
            </TabItem>
```

- [ ] **Step 2: Wire up the code-behind**

In `ToggJirazer/MainWindow.xaml.cs`:

1. Add a field next to `_isRunning`:

```csharp
    private bool _suppressAutoSendEvents;
```

2. In the constructor, after `PopulateUiFromConfig(...)`, add:

```csharp
        Loaded += MainWindow_Loaded;
```

3. Add the Loaded handler (queries actual Task Scheduler state off the UI thread):

```csharp
    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            var enabled = await Task.Run(ScheduledTaskService.IsEnabled);
            _suppressAutoSendEvents = true;
            AutoSendCheckBox.IsChecked = enabled;
            _suppressAutoSendEvents = false;
        }
        catch
        {
            // Leave unchecked if the query fails; toggling will surface any real error.
        }
    }
```

4. In `PopulateUiFromConfig`, add at the end:

```csharp
        EmailSmtpHost.Text = config.Email.SmtpHost;
        EmailSmtpPort.Text = config.Email.SmtpPort.ToString();
        EmailUseSsl.IsChecked = config.Email.UseSsl;
        EmailUsername.Text = config.Email.Username;
        EmailPassword.Password = config.Email.Password;
        EmailFromAddress.Text = config.Email.FromAddress;
        EmailRecipients.Text = config.Email.Recipients;
        EmailSubject.Text = config.Email.Subject;
        EmailBody.Text = config.Email.Body;
```

5. In `ReadConfigFromUi`, add to the returned `AppConfig` initializer (after `Report = ...`):

```csharp
            Email = new EmailConfig
            {
                SmtpHost = EmailSmtpHost.Text.Trim(),
                SmtpPort = int.TryParse(EmailSmtpPort.Text, out var smtpPort) ? smtpPort : 587,
                UseSsl = EmailUseSsl.IsChecked == true,
                Username = EmailUsername.Text.Trim(),
                Password = EmailPassword.Password,
                FromAddress = EmailFromAddress.Text.Trim(),
                Recipients = EmailRecipients.Text.Trim(),
                Subject = EmailSubject.Text,
                Body = EmailBody.Text
            }
```

6. Replace `RunButton_Click` with a shared flow plus two thin handlers. The body is the existing `RunButton_Click` with the marked additions:

```csharp
    private async void RunButton_Click(object sender, RoutedEventArgs e)
        => await RunReportFlowAsync(sendEmail: false);

    private async void SendEmailNow_Click(object sender, RoutedEventArgs e)
        => await RunReportFlowAsync(sendEmail: true);

    private async Task RunReportFlowAsync(bool sendEmail)
    {
        if (_isRunning) return;

        var config = ReadConfigFromUi();
        var error = ReportRunner.ValidateConfig(config);
        if (error == null && sendEmail)
        {
            error = EmailService.ValidateEmailConfig(config.Email);
            if (error == null && !string.Equals(config.Report.Format?.Trim(), "xlsx", StringComparison.OrdinalIgnoreCase))
                error = "Sending by email requires the xlsx report format (Report tab).";
        }
        if (error != null)
        {
            AppendLog($"[ERROR] {error}");
            return;
        }

        // Auto-save settings before running so the output path is persisted
        try { _settingsService.Save(config); } catch { /* ignore */ }

        // Read the cache toggle on the UI thread before work moves to a background thread.
        var cacheEnabled = EnableCacheCheckBox.IsChecked == true;

        _isRunning = true;
        RunButton.IsEnabled = false;
        SaveButton.IsEnabled = false;
        SendEmailNowButton.IsEnabled = false;
        LogTextBox.Clear();
        ProgressBar.IsIndeterminate = true;
        ProgressBar.Value = 0;

        // Redirect Console.Out / Console.Error to the log window for the duration of the run
        var originalOut = Console.Out;
        var originalError = Console.Error;
        var logWriter = new LogTextWriter(AppendLog);
        var errorWriter = new LogTextWriter(line => AppendLog($"[ERROR] {line}"));
        Console.SetOut(logWriter);
        Console.SetError(errorWriter);

        try
        {
            var (startDate, endDate) = ReportRunner.ResolveConfiguredPeriod(config.Report);
            await Task.Run(async () => await ReportRunner.RunAsync(
                config, startDate, endDate, cacheEnabled, ShowCacheWritePermissionDialog));

            if (sendEmail)
            {
                var subject = AutomationLogic.ResolvePlaceholders(config.Email.Subject, endDate);
                var body = AutomationLogic.ResolvePlaceholders(config.Email.Body, endDate);
                AppendLog($"Sending email to {string.Join(", ", EmailService.ParseRecipients(config.Email.Recipients))}...");
                await Task.Run(async () => await EmailService.SendReportAsync(
                    config.Email, subject, body, config.Report.OutputFile));
                AppendLog("Email sent.");
            }

            AppendLog(string.Empty);
            AppendLog("Done.");
            ProgressBar.IsIndeterminate = false;
            ProgressBar.Value = 100;
            _lastOutputPath = config.Report.OutputFile;
            OpenReportButton.IsEnabled = File.Exists(_lastOutputPath);
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
            ProgressBar.IsIndeterminate = false;
            ProgressBar.Value = 0;
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
            logWriter.Dispose();
            errorWriter.Dispose();
            RunButton.IsEnabled = true;
            SaveButton.IsEnabled = true;
            SendEmailNowButton.IsEnabled = true;
            _isRunning = false;
        }
    }
```

(Note: `_lastOutputPath` is now taken from the validated `config` instead of re-reading the UI — same value, avoids a second `ReadConfigFromUi()` call.)

7. Add the automation toggle handlers:

```csharp
    private async void AutoSend_Checked(object sender, RoutedEventArgs e)
    {
        if (_suppressAutoSendEvents) return;
        try
        {
            await Task.Run(ScheduledTaskService.Enable);
            AppendLog($"Daily automatic send enabled (task '{ScheduledTaskService.TaskName}', 09:00; a missed run fires when the PC is next on).");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
            _suppressAutoSendEvents = true;
            AutoSendCheckBox.IsChecked = false;
            _suppressAutoSendEvents = false;
        }
    }

    private async void AutoSend_Unchecked(object sender, RoutedEventArgs e)
    {
        if (_suppressAutoSendEvents) return;
        try
        {
            await Task.Run(ScheduledTaskService.Disable);
            AppendLog("Daily automatic send disabled.");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
            _suppressAutoSendEvents = true;
            AutoSendCheckBox.IsChecked = true;
            _suppressAutoSendEvents = false;
        }
    }
```

- [ ] **Step 3: Build and run tests**

Run: `dotnet build ToggJirazer/ToggJirazer.csproj -v q --nologo` — expected `0 Error(s)`.
Run: `dotnet test ToggJirazer.Tests/ToggJirazer.Tests.csproj -v q --nologo` — expected all PASS.
Do NOT launch the UI; the user performs the manual verification below after implementation.

- [ ] **Step 4: Document the feature in README.md**

Append a section to `README.md` (adapt heading level to the existing document structure):

```markdown
## Automatic monthly email

The Email tab configures SMTP settings and recipients for sending the report by email.

- **Send Report by Email Now** generates the report with the configured period and emails
  it immediately (requires xlsx format).
- **Send automatically every month** registers a daily Windows scheduled task
  ("ToggJirazer AutoReport", 09:00, runs when the PC is next on if missed) that starts
  the app with `--auto`. On each run the app checks `automation-state.json` next to the
  exe: if the report covering last month (period: Jan 1 → end of previous month) has not
  been sent yet, it is generated and emailed, and the state is updated. Failures are
  logged to `logs\auto.log` and retried on the next daily run.
```

- [ ] **Step 5: Commit**

```bash
git add ToggJirazer/MainWindow.xaml ToggJirazer/MainWindow.xaml.cs README.md
git commit -m "feat: add Email tab with send-now and monthly auto-send toggle"
```

---

## Manual verification (performed by the user after implementation)

1. Launch the app, fill in the Email tab (SMTP host/port/credentials, from, recipients — use your own address first), Save Settings.
2. Click **Send Report by Email Now** — verify the report generates and the email arrives with the xlsx attached and Danish subject/body with the month resolved.
3. Tick **Send automatically every month** — verify the task appears in Task Scheduler (`ToggJirazer AutoReport`) with a daily 09:00 trigger and "Run task as soon as possible after a scheduled start is missed" enabled.
4. Optionally run the task once manually from Task Scheduler and check `logs\auto.log` and `automation-state.json` in the app folder.
5. Untick the checkbox — verify the task is removed.
