# ToggJirazer .NET 10 Console TUI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convert ToggJirazer from a Windows-only WPF app into a cross-platform .NET 10 console application with a Terminal.Gui TUI plus a `--headless` automation mode, reusing all existing service/report code.

**Architecture:** Remove the WPF presentation layer; keep `Services/` and `Models/` untouched. Extract the report pipeline from `MainWindow.RunReportAsync` into a shared `Cli/ReportRunner`. Two front-ends consume it: a Terminal.Gui screen (`Tui/AppTui`) for interactive use and a `Cli/HeadlessRunner` for automation. `Program.cs` selects the mode from args and stdin/stdout redirection.

**Tech Stack:** .NET 10 (`net10.0`, `OutputType=Exe`), Terminal.Gui 2.4.x, ClosedXML, CsvHelper, xunit for tests.

**Spec:** `docs/superpowers/specs/2026-07-01-toggljirazer-console-tui-design.md`

**Reference implementation:** uRobot at `~/git/RoboCoder/src/uRobot.Cli/Tui/` uses Terminal.Gui 2.4.7. When exact Terminal.Gui v2 widget API is uncertain, cross-check against those files (`UiLauncher.Run.cs`, `TerminalGuiDispatcher.cs`, `TerminalGuiRenderer.cs`) — they are known-good against the same 2.4.x line.

## Global Constraints

- Target framework: `net10.0`. `OutputType` = `Exe`. No `UseWPF`, no `WinExe`, no `EnableWindowsTargeting`.
- Only one new runtime dependency: `Terminal.Gui` (2.4.x). **Do not** add Spectre.Console.
- All existing files under `ToggJirazer/Services/` and `ToggJirazer/Models/` are reused unchanged.
- Namespace root stays `ToggJirazer` (sub-namespaces `ToggJirazer.Cli`, `ToggJirazer.Tui`).
- The report pipeline is written once (`ReportRunner`) and used by both the TUI and headless modes — do not duplicate it.
- Background report runs must marshal UI/log updates onto the Terminal.Gui main loop via `Application.Invoke(...)` (the equivalent of the old WPF `Dispatcher.BeginInvoke`).
- Commit after every task.

---

## Task 1: Retarget project to net10 console and remove WPF

**Files:**
- Modify: `ToggJirazer/ToggJirazer.csproj`
- Delete: `ToggJirazer/App.xaml`, `ToggJirazer/App.xaml.cs`, `ToggJirazer/MainWindow.xaml`, `ToggJirazer/MainWindow.xaml.cs`
- Modify: `ToggJirazer/Program.cs` (temporary stub so the project compiles)

**Interfaces:**
- Consumes: nothing.
- Produces: a buildable `net10.0` console project. `Program.cs` has a temporary `Main` returning 0; later tasks replace it.

- [ ] **Step 1: Replace the csproj contents**

Replace the whole file with:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <RootNamespace>ToggJirazer</RootNamespace>
  </PropertyGroup>

  <ItemGroup>
    <None Remove="appsettings.development.json" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="ClosedXML" Version="0.104.2" />
    <PackageReference Include="CsvHelper" Version="33.0.1" />
    <PackageReference Include="Microsoft.Extensions.Configuration" Version="9.0.1" />
    <PackageReference Include="Microsoft.Extensions.Configuration.Binder" Version="9.0.1" />
    <PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="9.0.1" />
    <PackageReference Include="Terminal.Gui" Version="2.4.16" />
  </ItemGroup>

  <ItemGroup>
    <Content Include="appsettings.development.json" Condition="Exists('appsettings.development.json')">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </Content>
    <Content Include="appsettings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </Content>
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Delete the WPF files**

```bash
cd /Users/philipkaare/git/toggljirazer
git rm ToggJirazer/App.xaml ToggJirazer/App.xaml.cs ToggJirazer/MainWindow.xaml ToggJirazer/MainWindow.xaml.cs
```

- [ ] **Step 3: Replace Program.cs with a temporary stub**

```csharp
namespace ToggJirazer;

internal static class Program
{
    private static int Main(string[] args) => 0;
}
```

- [ ] **Step 4: Restore and build**

Run: `cd /Users/philipkaare/git/toggljirazer/ToggJirazer && dotnet build`
Expected: `Build succeeded. 0 Error(s)`. (The Terminal.Gui package restores; no WPF errors.)

- [ ] **Step 5: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Retarget to net10.0 console, remove WPF, add Terminal.Gui"
```

---

## Task 2: Create test project with SettingsService round-trip test

This task stands up the xunit harness and proves it by testing the existing `SettingsService`.

**Files:**
- Create: `ToggJirazer.Tests/ToggJirazer.Tests.csproj`
- Create: `ToggJirazer.Tests/SettingsServiceTests.cs`
- Modify: `ToggJirazer.slnx` (add the test project)

**Interfaces:**
- Consumes: `ToggJirazer.Services.SettingsService`, `ToggJirazer.Models.AppConfig`.
- Produces: a runnable `dotnet test` harness for later tasks.

- [ ] **Step 1: Create the test csproj**

`ToggJirazer.Tests/ToggJirazer.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\ToggJirazer\ToggJirazer.csproj" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the failing test**

`ToggJirazer.Tests/SettingsServiceTests.cs`:

```csharp
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Tests;

public class SettingsServiceTests
{
    [Fact]
    public void Save_then_Load_round_trips_config()
    {
        var path = Path.Combine(Path.GetTempPath(), $"tj-{Guid.NewGuid():N}.json");
        try
        {
            var service = new SettingsService(path);
            var original = new AppConfig
            {
                Toggl = new TogglConfig { ApiToken = "tok", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
                Jira = new JiraConfig { BaseUrl = "https://x.atlassian.net", UserEmail = "a@b.c", ApiToken = "jt" },
                Report = new ReportConfig { StartDate = "2026-01-01", EndDate = "2026-06-30", OutputFile = "r.xlsx", Format = "xlsx" }
            };

            service.Save(original);
            var loaded = service.Load();

            Assert.Equal(original.Toggl.ApiToken, loaded.Toggl.ApiToken);
            Assert.Equal(original.Toggl.WorkspaceId, loaded.Toggl.WorkspaceId);
            Assert.Equal(original.Jira.BaseUrl, loaded.Jira.BaseUrl);
            Assert.Equal(original.Report.OutputFile, loaded.Report.OutputFile);
            Assert.Equal(original.Report.StartDate, loaded.Report.StartDate);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
```

- [ ] **Step 3: Add the test project to the solution**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet sln ToggJirazer.slnx add ToggJirazer.Tests/ToggJirazer.Tests.csproj`
(If `slnx` add is unsupported by the installed SDK, edit `ToggJirazer.slnx` by hand to add `<Project Path="ToggJirazer.Tests/ToggJirazer.Tests.csproj" />` alongside the existing project entry.)

- [ ] **Step 4: Run the test**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test`
Expected: PASS (1 test passed).

- [ ] **Step 5: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Add xunit test project with SettingsService round-trip test"
```

---

## Task 3: ConfigValidator (shared validation)

**Files:**
- Create: `ToggJirazer/Cli/ConfigValidator.cs`
- Create: `ToggJirazer.Tests/ConfigValidatorTests.cs`

**Interfaces:**
- Consumes: `ToggJirazer.Models.AppConfig`.
- Produces: `static string? ToggJirazer.Cli.ConfigValidator.Validate(AppConfig config)` — returns `null` when valid, otherwise a human-readable error message. Consumed by Task 4 (headless) and Task 7 (TUI).

- [ ] **Step 1: Write the failing tests**

`ToggJirazer.Tests/ConfigValidatorTests.cs`:

```csharp
using ToggJirazer.Cli;
using ToggJirazer.Models;

namespace ToggJirazer.Tests;

public class ConfigValidatorTests
{
    private static AppConfig Valid() => new()
    {
        Toggl = new TogglConfig { ApiToken = "t", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
        Jira = new JiraConfig { BaseUrl = "https://x.atlassian.net", UserEmail = "a@b.c", ApiToken = "j" },
        Report = new ReportConfig { OutputFile = "r.xlsx", Format = "xlsx" }
    };

    [Fact]
    public void Valid_config_returns_null()
        => Assert.Null(ConfigValidator.Validate(Valid()));

    [Fact]
    public void Missing_toggl_token_is_reported()
    {
        var c = Valid();
        c.Toggl.ApiToken = "";
        Assert.Contains("Toggl API Token", ConfigValidator.Validate(c));
    }

    [Fact]
    public void Zero_workspace_id_is_reported()
    {
        var c = Valid();
        c.Toggl.WorkspaceId = 0;
        Assert.Contains("Workspace ID", ConfigValidator.Validate(c));
    }

    [Fact]
    public void Missing_jira_base_url_is_reported()
    {
        var c = Valid();
        c.Jira.BaseUrl = "";
        Assert.Contains("Jira Base URL", ConfigValidator.Validate(c));
    }

    [Fact]
    public void Missing_output_file_is_reported()
    {
        var c = Valid();
        c.Report.OutputFile = "";
        Assert.Contains("Output File", ConfigValidator.Validate(c));
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter ConfigValidatorTests`
Expected: FAIL — `ConfigValidator` does not exist (compile error).

- [ ] **Step 3: Implement ConfigValidator**

`ToggJirazer/Cli/ConfigValidator.cs` (rules copied verbatim from the old `MainWindow.ValidateConfig`, with tab hints removed since there are no tabs in headless):

```csharp
using ToggJirazer.Models;

namespace ToggJirazer.Cli;

/// <summary>
/// Validates an <see cref="AppConfig"/> before a report run.
/// Returns <c>null</c> when the config is complete, otherwise a message
/// naming the first missing required field.
/// </summary>
public static class ConfigValidator
{
    public static string? Validate(AppConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.Toggl.ApiToken))
            return "Toggl API Token is required.";
        if (config.Toggl.OrganizationId == 0)
            return "Toggl Organization ID is required.";
        if (config.Toggl.WorkspaceId == 0)
            return "Toggl Workspace ID is required.";
        if (config.Toggl.ProjectId == 0)
            return "Toggl Project ID is required.";
        if (string.IsNullOrWhiteSpace(config.Jira.BaseUrl))
            return "Jira Base URL is required.";
        if (string.IsNullOrWhiteSpace(config.Jira.UserEmail))
            return "Jira User Email is required.";
        if (string.IsNullOrWhiteSpace(config.Jira.ApiToken))
            return "Jira API Token is required.";
        if (string.IsNullOrWhiteSpace(config.Report.OutputFile))
            return "Output File path is required.";
        return null;
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter ConfigValidatorTests`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Add shared ConfigValidator with tests"
```

---

## Task 4: ReportRunner (extract the report pipeline)

Ports `MainWindow.RunReportAsync` into a reusable class. The reporting-period logic is extracted into a separately testable method; the rest of the pipeline (which calls live Toggl/Jira) is preserved verbatim and writes progress via a log sink.

**Files:**
- Create: `ToggJirazer/Cli/ReportRunner.cs`
- Create: `ToggJirazer.Tests/ReportPeriodTests.cs`

**Interfaces:**
- Consumes: `TogglService`, `JiraService`, `ReportService` (existing, unchanged), `AppConfig`.
- Produces:
  - `static (DateTime Start, DateTime End) ToggJirazer.Cli.ReportRunner.ResolvePeriod(ReportConfig report, DateTime today, Action<string> log)` — parses configured dates or defaults to Jan-1→today. Throws `InvalidOperationException` on unparseable dates.
  - `static Task ToggJirazer.Cli.ReportRunner.RunAsync(AppConfig config, Action<string> log)` — the full pipeline. All progress goes to `log`. Consumed by Task 5 (headless) and Task 7 (TUI).

- [ ] **Step 1: Write the failing tests for ResolvePeriod**

`ToggJirazer.Tests/ReportPeriodTests.cs`:

```csharp
using ToggJirazer.Cli;
using ToggJirazer.Models;

namespace ToggJirazer.Tests;

public class ReportPeriodTests
{
    private static void Sink(string _) { }

    [Fact]
    public void Explicit_dates_are_parsed()
    {
        var report = new ReportConfig { StartDate = "2026-02-01", EndDate = "2026-02-28" };
        var (start, end) = ReportRunner.ResolvePeriod(report, new DateTime(2026, 7, 1), Sink);
        Assert.Equal(new DateTime(2026, 2, 1), start);
        Assert.Equal(new DateTime(2026, 2, 28), end);
    }

    [Fact]
    public void Empty_dates_default_to_year_start_and_today()
    {
        var report = new ReportConfig { StartDate = null, EndDate = null };
        var (start, end) = ReportRunner.ResolvePeriod(report, new DateTime(2026, 7, 1), Sink);
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 7, 1), end);
    }

    [Fact]
    public void Unparseable_start_date_throws()
    {
        var report = new ReportConfig { StartDate = "not-a-date", EndDate = "2026-02-28" };
        Assert.Throws<InvalidOperationException>(
            () => ReportRunner.ResolvePeriod(report, new DateTime(2026, 7, 1), Sink));
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter ReportPeriodTests`
Expected: FAIL — `ReportRunner` does not exist.

- [ ] **Step 3: Implement ReportRunner**

`ToggJirazer/Cli/ReportRunner.cs`. `RunAsync` is a direct port of the old `MainWindow.RunReportAsync`, with every `Console.WriteLine(x)` replaced by `log(x)` and `Console.Error.WriteLine(x)` by `log("[ERROR] " + x)`:

```csharp
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Cli;

/// <summary>
/// Runs the end-to-end report pipeline: resolve period, fetch Toggl entries,
/// enrich with Jira, and write the xlsx/csv output. All progress is emitted
/// through the supplied <paramref name="log"/> sink so both the TUI and the
/// headless front-end can consume it.
/// </summary>
public static class ReportRunner
{
    public static (DateTime Start, DateTime End) ResolvePeriod(
        ReportConfig report, DateTime today, Action<string> log)
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

        var startDate = new DateTime(today.Year, 1, 1);
        log($"No period specified — defaulting to Jan 1 to today: {startDate:yyyy-MM-dd} to {today:yyyy-MM-dd}");
        return (startDate, today);
    }

    public static async Task RunAsync(AppConfig config, Action<string> log)
    {
        var (startDate, endDate) = ResolvePeriod(config.Report, DateTime.Today, log);

        log($"Report period: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
        log($"Toggl workspace: {config.Toggl.WorkspaceId}, project: {config.Toggl.ProjectId}");
        log($"Jira base URL: {config.Jira.BaseUrl}");
        log(string.Empty);

        using var togglService = new TogglService(config.Toggl);
        var entries = await togglService.GetDetailedReportAsync(startDate, endDate);

        if (entries.Count == 0)
        {
            log("No Toggl time entries found for the specified period and project.");
            return;
        }

        log(string.Empty);
        log("Fetching all-time Toggl entries for version totals...");
        var allEntries = await togglService.GetAllEntriesAsync();
        log(string.Empty);

        using var jiraService = new JiraService(config.Jira);
        var reportService = new ReportService(jiraService);
        var (reportSheets, versionRows, leverances,
             categoryConsumedHours, weeklyCategoryConsumedHours, monthlyCategoryConsumedHours) =
            await reportService.BuildReportAsync(entries, allEntries);

        log(string.Empty);
        log($"Report contains {reportSheets.Sum(s => s.Rows.Count)} rows.");
        log($"Version report contains {versionRows.Count} rows.");
        log($"Leverances: {leverances.Count}.");
        log(string.Empty);

        var format = config.Report.Format?.Trim().ToLowerInvariant();
        if (format == "xlsx")
        {
            var extraColumns = reportService.ReadVersionReportExtraColumns(config.Report.OutputFile);
            reportService.WriteXlsx(
                reportSheets, versionRows, leverances,
                categoryConsumedHours, weeklyCategoryConsumedHours, monthlyCategoryConsumedHours,
                config.Report.OutputFile, extraColumns);
        }
        else
        {
            if (!string.IsNullOrEmpty(format) && format != "csv")
                log($"[ERROR] Warning: Unrecognized format '{config.Report.Format}'. Defaulting to CSV.");

            var rows = reportSheets.SelectMany(s => s.Rows).ToList();
            reportService.WriteCsv(rows, config.Report.OutputFile);

            var outputPath = config.Report.OutputFile;
            var versionOutputPath = Path.Combine(
                Path.GetDirectoryName(outputPath) ?? string.Empty,
                Path.GetFileNameWithoutExtension(outputPath) + "_versions" + Path.GetExtension(outputPath));
            var extraColumns = reportService.ReadVersionReportExtraColumns(versionOutputPath);
            reportService.WriteVersionCsv(versionRows, versionOutputPath, extraColumns);
        }

        log(string.Empty);
        log("Done.");
    }
}
```

Note: add `using System.IO;` is unnecessary (ImplicitUsings covers `System.IO`), but `Path`/`File` resolve via implicit usings.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter ReportPeriodTests`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Add ReportRunner with period-resolution tests"
```

---

## Task 5: PlatformOpener (cross-platform open)

**Files:**
- Create: `ToggJirazer/Cli/PlatformOpener.cs`
- Create: `ToggJirazer.Tests/PlatformOpenerTests.cs`

**Interfaces:**
- Produces:
  - `enum ToggJirazer.Cli.OpenerOs { MacOs, Linux, Windows }`
  - `static (string FileName, string Arguments, bool UseShellExecute) ToggJirazer.Cli.PlatformOpener.CommandFor(OpenerOs os, string path)` — pure, testable command selection.
  - `static void ToggJirazer.Cli.PlatformOpener.Open(string path)` — detects the current OS and launches the opener. Consumed by Task 7 (TUI "Open report").

- [ ] **Step 1: Write the failing tests**

`ToggJirazer.Tests/PlatformOpenerTests.cs`:

```csharp
using ToggJirazer.Cli;

namespace ToggJirazer.Tests;

public class PlatformOpenerTests
{
    [Fact]
    public void MacOs_uses_open()
    {
        var cmd = PlatformOpener.CommandFor(OpenerOs.MacOs, "/tmp/r.xlsx");
        Assert.Equal("open", cmd.FileName);
        Assert.Contains("/tmp/r.xlsx", cmd.Arguments);
        Assert.False(cmd.UseShellExecute);
    }

    [Fact]
    public void Linux_uses_xdg_open()
    {
        var cmd = PlatformOpener.CommandFor(OpenerOs.Linux, "/tmp/r.xlsx");
        Assert.Equal("xdg-open", cmd.FileName);
        Assert.False(cmd.UseShellExecute);
    }

    [Fact]
    public void Windows_uses_shell_execute_on_the_path()
    {
        var cmd = PlatformOpener.CommandFor(OpenerOs.Windows, @"C:\r.xlsx");
        Assert.Equal(@"C:\r.xlsx", cmd.FileName);
        Assert.True(cmd.UseShellExecute);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter PlatformOpenerTests`
Expected: FAIL — `PlatformOpener` does not exist.

- [ ] **Step 3: Implement PlatformOpener**

`ToggJirazer/Cli/PlatformOpener.cs`:

```csharp
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ToggJirazer.Cli;

public enum OpenerOs { MacOs, Linux, Windows }

/// <summary>Opens a file with the operating system's default application.</summary>
public static class PlatformOpener
{
    public static (string FileName, string Arguments, bool UseShellExecute) CommandFor(OpenerOs os, string path)
        => os switch
        {
            OpenerOs.MacOs => ("open", Quote(path), false),
            OpenerOs.Linux => ("xdg-open", Quote(path), false),
            OpenerOs.Windows => (path, string.Empty, true),
            _ => (path, string.Empty, true),
        };

    public static void Open(string path)
    {
        var cmd = CommandFor(CurrentOs(), path);
        Process.Start(new ProcessStartInfo
        {
            FileName = cmd.FileName,
            Arguments = cmd.Arguments,
            UseShellExecute = cmd.UseShellExecute,
        });
    }

    private static OpenerOs CurrentOs()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return OpenerOs.MacOs;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return OpenerOs.Windows;
        return OpenerOs.Linux;
    }

    private static string Quote(string path) => $"\"{path}\"";
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter PlatformOpenerTests`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Add cross-platform PlatformOpener with tests"
```

---

## Task 6: HeadlessRunner and Program mode selection

**Files:**
- Create: `ToggJirazer/Cli/HeadlessRunner.cs`
- Create: `ToggJirazer/Cli/ArgParser.cs`
- Create: `ToggJirazer.Tests/ArgParserTests.cs`
- Modify: `ToggJirazer/Program.cs`

**Interfaces:**
- Consumes: `ConfigValidator`, `ReportRunner`, `SettingsService`, `AppConfig`.
- Produces:
  - `enum ToggJirazer.Cli.LaunchMode { Tui, Headless, Help }`
  - `static LaunchMode ToggJirazer.Cli.ArgParser.Resolve(string[] args, bool inputRedirected, bool outputRedirected)`
  - `static Task<int> ToggJirazer.Cli.HeadlessRunner.RunAsync(AppConfig config)` — validates, runs, returns process exit code.
  - `Program.Main` wiring (the real entry point). Task 7 provides `AppTui.Run(...)`; until then Program calls a placeholder for the Tui branch (replaced in Task 7).

- [ ] **Step 1: Write the failing tests for ArgParser**

`ToggJirazer.Tests/ArgParserTests.cs`:

```csharp
using ToggJirazer.Cli;

namespace ToggJirazer.Tests;

public class ArgParserTests
{
    [Fact]
    public void Help_flag_returns_help()
    {
        Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["--help"], false, false));
        Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["-h"], false, false));
    }

    [Fact]
    public void Headless_flag_returns_headless()
    {
        Assert.Equal(LaunchMode.Headless, ArgParser.Resolve(["--headless"], false, false));
        Assert.Equal(LaunchMode.Headless, ArgParser.Resolve(["--run"], false, false));
    }

    [Fact]
    public void Redirected_output_forces_headless()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve([], false, true));

    [Fact]
    public void Redirected_input_forces_headless()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve([], true, false));

    [Fact]
    public void Interactive_no_flags_returns_tui()
        => Assert.Equal(LaunchMode.Tui, ArgParser.Resolve([], false, false));
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter ArgParserTests`
Expected: FAIL — `ArgParser` does not exist.

- [ ] **Step 3: Implement ArgParser**

`ToggJirazer/Cli/ArgParser.cs`:

```csharp
namespace ToggJirazer.Cli;

public enum LaunchMode { Tui, Headless, Help }

public static class ArgParser
{
    public static LaunchMode Resolve(string[] args, bool inputRedirected, bool outputRedirected)
    {
        foreach (var a in args)
        {
            if (a is "--help" or "-h") return LaunchMode.Help;
        }
        foreach (var a in args)
        {
            if (a is "--headless" or "--run") return LaunchMode.Headless;
        }
        if (inputRedirected || outputRedirected) return LaunchMode.Headless;
        return LaunchMode.Tui;
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter ArgParserTests`
Expected: PASS (5 tests).

- [ ] **Step 5: Implement HeadlessRunner**

`ToggJirazer/Cli/HeadlessRunner.cs`:

```csharp
using ToggJirazer.Models;

namespace ToggJirazer.Cli;

/// <summary>Non-interactive report run: validate, run, return an exit code.</summary>
public static class HeadlessRunner
{
    public static async Task<int> RunAsync(AppConfig config)
    {
        var error = ConfigValidator.Validate(config);
        if (error != null)
        {
            Console.Error.WriteLine($"[ERROR] {error}");
            return 2;
        }

        try
        {
            await ReportRunner.RunAsync(config, Console.WriteLine);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[ERROR] {ex.Message}");
            return 1;
        }
    }
}
```

- [ ] **Step 6: Rewrite Program.cs as the real entry point**

`ToggJirazer/Program.cs`:

```csharp
using ToggJirazer.Cli;
using ToggJirazer.Services;
using ToggJirazer.Tui;

namespace ToggJirazer;

internal static class Program
{
    private static readonly string SettingsPath =
        Path.Combine(AppContext.BaseDirectory, "appsettings.json");

    private static async Task<int> Main(string[] args)
    {
        var mode = ArgParser.Resolve(args, Console.IsInputRedirected, Console.IsOutputRedirected);
        var settings = new SettingsService(SettingsPath);

        switch (mode)
        {
            case LaunchMode.Help:
                PrintUsage();
                return 0;

            case LaunchMode.Headless:
                return await HeadlessRunner.RunAsync(settings.Load());

            default:
                try
                {
                    AppTui.Run(settings, SettingsPath);
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Could not start the TUI ({ex.Message}). Running headless instead.");
                    return await HeadlessRunner.RunAsync(settings.Load());
                }
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("ToggJirazer — generate Toggl × Jira time reports.");
        Console.WriteLine();
        Console.WriteLine("Usage: ToggJirazer [options]");
        Console.WriteLine();
        Console.WriteLine("  (no options)     Launch the interactive TUI.");
        Console.WriteLine("  --headless,--run Run the report non-interactively using appsettings.json.");
        Console.WriteLine("  --help, -h       Show this help.");
        Console.WriteLine();
        Console.WriteLine("Configuration is read from appsettings.json next to the executable.");
    }
}
```

Note: this references `ToggJirazer.Tui.AppTui.Run(SettingsService, string)`, created in Task 7. The project will not compile until Task 7 adds `AppTui`. To keep Task 6 independently green, add a minimal placeholder now and let Task 7 replace it:

Create `ToggJirazer/Tui/AppTui.cs` (placeholder):

```csharp
using ToggJirazer.Services;

namespace ToggJirazer.Tui;

public static class AppTui
{
    public static void Run(SettingsService settings, string settingsPath)
        => throw new NotSupportedException("TUI not yet implemented.");
}
```

- [ ] **Step 7: Build and verify headless works end-to-end**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet build && dotnet test`
Expected: build succeeds, all tests pass.

Run: `cd /Users/philipkaare/git/toggljirazer/ToggJirazer && dotnet run -- --help`
Expected: usage text prints, exit 0.

Run: `cd /Users/philipkaare/git/toggljirazer/ToggJirazer && dotnet run -- --headless; echo "exit=$?"`
Expected: with the default placeholder `appsettings.json` (zeros/placeholders), prints `[ERROR] Toggl API Token is required.` (or the first missing field) and `exit=2`.

- [ ] **Step 8: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Add headless mode, arg parsing, and real Program entry point"
```

---

## Task 7: Banner and Terminal.Gui screen (AppTui)

Replaces the `AppTui` placeholder with the full interactive screen. This is the largest task and the widget API must match Terminal.Gui 2.4.x — **cross-check every widget against the uRobot reference files** (`~/git/RoboCoder/src/uRobot.Cli/Tui/UiLauncher.Run.cs`, `TerminalGuiDispatcher.cs`, `TerminalGuiRenderer.cs`, `TuiShell.cs`) which compile against 2.4.7. Verify the build after each sub-area rather than writing the whole file blind.

**Files:**
- Create: `ToggJirazer/Tui/Banner.cs`
- Create: `ToggJirazer.Tests/BannerTests.cs`
- Replace: `ToggJirazer/Tui/AppTui.cs` (placeholder → full screen)
- Reuse: `ToggJirazer/LogTextWriter.cs` (unchanged)

**Interfaces:**
- Consumes: `SettingsService`, `ConfigValidator`, `ReportRunner`, `PlatformOpener`, `AppConfig`, `LogTextWriter`.
- Produces: `static void ToggJirazer.Tui.AppTui.Run(SettingsService settings, string settingsPath)` — blocks until the user quits. `static string ToggJirazer.Tui.Banner.Text(string version)`.

- [ ] **Step 1: Write the failing Banner test**

`ToggJirazer.Tests/BannerTests.cs`:

```csharp
using ToggJirazer.Tui;

namespace ToggJirazer.Tests;

public class BannerTests
{
    [Fact]
    public void Text_includes_product_name_and_version()
    {
        var text = Banner.Text("1.2.3");
        Assert.Contains("ToggJirazer", text);
        Assert.Contains("1.2.3", text);
    }
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter BannerTests`
Expected: FAIL — `Banner` does not exist.

- [ ] **Step 3: Implement Banner**

`ToggJirazer/Tui/Banner.cs`:

```csharp
using System.Text;

namespace ToggJirazer.Tui;

/// <summary>Static ASCII masthead shown at the top of the TUI.</summary>
public static class Banner
{
    public static string Text(string version)
    {
        var sb = new StringBuilder();
        sb.AppendLine(@" _____                 _ _ _                     ");
        sb.AppendLine(@"|_   _|__  __ _ __ _   | (_) |_ __ _ _______ _ __ ");
        sb.AppendLine(@"  | |/ _ \/ _` / _` |  | | | __/ _` |_  / _ \ '__|");
        sb.AppendLine(@"  | | (_) \__, \__, |__| | | || (_| |/ /  __/ |   ");
        sb.AppendLine(@"  |_|\___/|___/|___/|__/|_| |\__\__,_/___\___|_|   ");
        sb.AppendLine($"ToggJirazer v{version}  ·  Toggl × Jira time reports");
        return sb.ToString();
    }
}
```

- [ ] **Step 4: Run to verify it passes**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test --filter BannerTests`
Expected: PASS.

- [ ] **Step 5: Implement AppTui**

Replace `ToggJirazer/Tui/AppTui.cs` with the full screen below. It reproduces the WPF form: three field groups (Toggl / Jira / Report) inside a `TabView`, Save / Generate / Open buttons, an indeterminate `ProgressBar`, and a scrolling `TextView` log. The report runs on a background `Task`; `Console.Out`/`Console.Error` are redirected through `LogTextWriter` whose sink marshals onto the main loop via `Application.Invoke`, mirroring the old WPF `Dispatcher.BeginInvoke`.

> **API note:** The code below targets Terminal.Gui v2 (`Terminal.Gui.App`, `Terminal.Gui.Views`, `Terminal.Gui.ViewBase` namespaces; `Application.Init/Run/Shutdown/Invoke`; `Pos`/`Dim` layout; `TextField.Secret`; `TabView`/`Tab`). If any member name differs in 2.4.16, correct it against the uRobot reference files and the package's public API — keep the behavior identical.

```csharp
using System.Reflection;
using Terminal.Gui.App;
using Terminal.Gui.ViewBase;
using Terminal.Gui.Views;
using ToggJirazer.Cli;
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Tui;

/// <summary>The interactive Terminal.Gui screen: edit settings, generate the
/// report, watch progress, open the result.</summary>
public static class AppTui
{
    public static void Run(SettingsService settings, string settingsPath)
    {
        Application.Init();
        try
        {
            var ui = new MainView(settings);
            Application.Run(ui);
        }
        finally
        {
            Application.Shutdown();
        }
    }
}

internal sealed class MainView : Window
{
    private readonly SettingsService _settings;

    // Toggl
    private readonly TextField _togglToken = new() { Secret = true };
    private readonly TextField _togglOrg = new();
    private readonly TextField _togglWorkspace = new();
    private readonly TextField _togglProject = new();
    // Jira
    private readonly TextField _jiraBaseUrl = new();
    private readonly TextField _jiraEmail = new();
    private readonly TextField _jiraToken = new() { Secret = true };
    // Report
    private readonly TextField _startDate = new();
    private readonly TextField _endDate = new();
    private readonly TextField _outputFile = new();
    private readonly RadioGroup _format = new() { RadioLabels = ["xlsx", "csv"] };

    private readonly TextView _log = new()
    {
        ReadOnly = true,
        WordWrap = false,
    };
    private readonly ProgressBar _progress = new()
    {
        ProgressBarStyle = ProgressBarStyle.MarqueeBlocks,
        Visible = false,
    };
    private readonly Button _saveBtn = new() { Title = "Save (F2)" };
    private readonly Button _runBtn = new() { Title = "Generate (F5)" };
    private readonly Button _openBtn = new() { Title = "Open report", Enabled = false };

    private bool _isRunning;
    private string? _lastOutputPath;

    public MainView(SettingsService settings)
    {
        _settings = settings;
        Title = "ToggJirazer  (Ctrl+Q to quit)";

        var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.0";
        var banner = new Label { Text = Banner.Text(version), X = 0, Y = 0, Height = 6 };

        var tabs = BuildTabs();
        tabs.X = 0; tabs.Y = 6; tabs.Width = Dim.Fill(); tabs.Height = 9;

        _saveBtn.X = 0; _saveBtn.Y = 15;
        _runBtn.X = Pos.Right(_saveBtn) + 2; _runBtn.Y = 15;
        _openBtn.X = Pos.Right(_runBtn) + 2; _openBtn.Y = 15;

        _progress.X = 0; _progress.Y = 17; _progress.Width = Dim.Fill(); _progress.Height = 1;

        var logFrame = new FrameView { Title = "Log", X = 0, Y = 18, Width = Dim.Fill(), Height = Dim.Fill() };
        _log.X = 0; _log.Y = 0; _log.Width = Dim.Fill(); _log.Height = Dim.Fill();
        logFrame.Add(_log);

        Add(banner, tabs, _saveBtn, _runBtn, _openBtn, _progress, logFrame);

        _saveBtn.Accepting += (_, _) => SaveSettings();
        _runBtn.Accepting += (_, _) => _ = GenerateAsync();
        _openBtn.Accepting += (_, _) => OpenReport();

        PopulateFromConfig(_settings.Load());
    }

    private TabView BuildTabs()
    {
        var tabs = new TabView();

        var toggl = new View { Width = Dim.Fill(), Height = Dim.Fill() };
        AddField(toggl, "API Token:", _togglToken, 0);
        AddField(toggl, "Organization:", _togglOrg, 2);
        AddField(toggl, "Workspace:", _togglWorkspace, 4);
        AddField(toggl, "Project:", _togglProject, 6);

        var jira = new View { Width = Dim.Fill(), Height = Dim.Fill() };
        AddField(jira, "Base URL:", _jiraBaseUrl, 0);
        AddField(jira, "User Email:", _jiraEmail, 2);
        AddField(jira, "API Token:", _jiraToken, 4);

        var report = new View { Width = Dim.Fill(), Height = Dim.Fill() };
        AddField(report, "Start Date:", _startDate, 0);
        AddField(report, "End Date:", _endDate, 2);
        AddField(report, "Output File:", _outputFile, 4);
        var fmtLabel = new Label { Text = "Format:", X = 0, Y = 6 };
        _format.X = 16; _format.Y = 6;
        report.Add(fmtLabel, _format);

        tabs.AddTab(new Tab { DisplayText = "Toggl", View = toggl }, true);
        tabs.AddTab(new Tab { DisplayText = "Jira", View = jira }, false);
        tabs.AddTab(new Tab { DisplayText = "Report", View = report }, false);
        return tabs;
    }

    private static void AddField(View parent, string label, TextField field, int row)
    {
        var lbl = new Label { Text = label, X = 0, Y = row };
        field.X = 16; field.Y = row; field.Width = Dim.Fill(1);
        parent.Add(lbl, field);
    }

    // ---- config <-> UI -------------------------------------------------------

    private void PopulateFromConfig(AppConfig c)
    {
        _togglToken.Text = c.Toggl.ApiToken;
        _togglOrg.Text = c.Toggl.OrganizationId == 0 ? "" : c.Toggl.OrganizationId.ToString();
        _togglWorkspace.Text = c.Toggl.WorkspaceId == 0 ? "" : c.Toggl.WorkspaceId.ToString();
        _togglProject.Text = c.Toggl.ProjectId == 0 ? "" : c.Toggl.ProjectId.ToString();
        _jiraBaseUrl.Text = c.Jira.BaseUrl;
        _jiraEmail.Text = c.Jira.UserEmail;
        _jiraToken.Text = c.Jira.ApiToken;
        _startDate.Text = c.Report.StartDate ?? new DateTime(DateTime.Today.Year, 1, 1).ToString("yyyy-MM-dd");
        _endDate.Text = c.Report.EndDate ?? DateTime.Today.ToString("yyyy-MM-dd");
        _outputFile.Text = c.Report.OutputFile;
        _format.SelectedItem = string.Equals(c.Report.Format, "csv", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
    }

    private AppConfig ReadConfig()
    {
        long.TryParse(_togglOrg.Text, out var org);
        long.TryParse(_togglWorkspace.Text, out var ws);
        long.TryParse(_togglProject.Text, out var proj);
        return new AppConfig
        {
            Toggl = new TogglConfig
            {
                ApiToken = _togglToken.Text, OrganizationId = org, WorkspaceId = ws, ProjectId = proj
            },
            Jira = new JiraConfig
            {
                BaseUrl = _jiraBaseUrl.Text.Trim(), UserEmail = _jiraEmail.Text.Trim(), ApiToken = _jiraToken.Text
            },
            Report = new ReportConfig
            {
                StartDate = string.IsNullOrWhiteSpace(_startDate.Text) ? null : _startDate.Text.Trim(),
                EndDate = string.IsNullOrWhiteSpace(_endDate.Text) ? null : _endDate.Text.Trim(),
                OutputFile = _outputFile.Text.Trim(),
                Format = _format.SelectedItem == 1 ? "csv" : "xlsx"
            }
        };
    }

    // ---- actions -------------------------------------------------------------

    private void SaveSettings()
    {
        try
        {
            _settings.Save(ReadConfig());
            AppendLog("Settings saved.");
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] Could not save settings: {ex.Message}");
        }
    }

    private async Task GenerateAsync()
    {
        if (_isRunning) return;

        var config = ReadConfig();
        var error = ConfigValidator.Validate(config);
        if (error != null) { AppendLog($"[ERROR] {error}"); return; }

        try { _settings.Save(config); } catch { /* ignore */ }

        _isRunning = true;
        SetButtons(false);
        _log.Text = "";
        _progress.Visible = true;

        var originalOut = Console.Out;
        var originalError = Console.Error;
        var logWriter = new LogTextWriter(AppendLog);
        var errorWriter = new LogTextWriter(line => AppendLog($"[ERROR] {line}"));
        Console.SetOut(logWriter);
        Console.SetError(errorWriter);

        try
        {
            await Task.Run(() => ReportRunner.RunAsync(config, s => Console.WriteLine(s)));
            _lastOutputPath = config.Report.OutputFile;
            var exists = File.Exists(_lastOutputPath);
            Application.Invoke(() => _openBtn.Enabled = exists);
        }
        catch (Exception ex)
        {
            AppendLog($"[ERROR] {ex.Message}");
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
            logWriter.Dispose();
            errorWriter.Dispose();
            Application.Invoke(() =>
            {
                _progress.Visible = false;
                SetButtons(true);
            });
            _isRunning = false;
        }
    }

    private void OpenReport()
    {
        if (string.IsNullOrWhiteSpace(_lastOutputPath) || !File.Exists(_lastOutputPath))
        {
            AppendLog("[ERROR] Report file not found.");
            return;
        }
        try { PlatformOpener.Open(_lastOutputPath); }
        catch (Exception ex) { AppendLog($"[ERROR] Could not open report: {ex.Message}"); }
    }

    private void SetButtons(bool enabled)
    {
        _runBtn.Enabled = enabled;
        _saveBtn.Enabled = enabled;
    }

    private void AppendLog(string line)
    {
        Application.Invoke(() =>
        {
            _log.Text += line + "\n";
            _log.MoveEnd();
        });
    }
}
```

- [ ] **Step 6: Build**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet build`
Expected: `Build succeeded`. If any Terminal.Gui member does not resolve, fix the name/namespace against the uRobot reference files and rebuild. Do not proceed until the build is green.

- [ ] **Step 7: Run the full test suite**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet test`
Expected: all tests pass.

- [ ] **Step 8: Manual launch smoke test**

Run: `cd /Users/philipkaare/git/toggljirazer/ToggJirazer && dotnet run`
Expected: the TUI renders with the banner, three tabs, the buttons, and the log frame. Tab between fields, type in one, press the Save button (should log "Settings saved."), then quit with Ctrl+Q. (Driving this under tmux is fine; a blank screen means launch failure — investigate before marking done.)

- [ ] **Step 9: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Add Terminal.Gui interactive screen and ASCII banner"
```

---

## Task 8: README and run.sh, final verification

**Files:**
- Modify: `README.md`
- Create: `run.sh`

**Interfaces:**
- Consumes: the finished app.
- Produces: accurate docs and a launch script.

- [ ] **Step 1: Update the README**

In `README.md`, replace the "Prerequisites", "Building & Running", and any WPF/Windows implications:

- Prerequisites: change ".NET 9 SDK" → ".NET 10 SDK".
- Replace the "Building & Running" section with:

```markdown
## Building & Running

```bash
# Interactive TUI
dotnet run --project ToggJirazer

# Or use the helper script
./run.sh

# Non-interactive run (reads appsettings.json, writes the report, exits)
dotnet run --project ToggJirazer -- --headless
```

The app is a cross-platform .NET 10 console application (Terminal.Gui). It runs
on macOS, Linux, and Windows. In the TUI you can edit the Toggl/Jira/Report
settings, save them, generate the report, and open it in your default
spreadsheet app. Use `--headless` (or pipe/redirect output) for automation; the
process exits non-zero on validation or run errors.
```

- Add a one-line note under Overview: "ToggJirazer runs as an interactive terminal UI, or non-interactively with `--headless`."

- [ ] **Step 2: Create run.sh**

`run.sh`:

```bash
#!/usr/bin/env bash
# Run ToggJirazer. Any arguments are forwarded to the program.
set -euo pipefail

cd "$(dirname "$0")"

dotnet run --project ToggJirazer/ToggJirazer.csproj -- "$@"
```

- [ ] **Step 3: Make it executable**

Run: `cd /Users/philipkaare/git/toggljirazer && chmod +x run.sh`

- [ ] **Step 4: Final full verification**

Run: `cd /Users/philipkaare/git/toggljirazer && dotnet build && dotnet test`
Expected: build succeeds, all tests pass.

Run: `cd /Users/philipkaare/git/toggljirazer && ./run.sh --help`
Expected: usage text prints.

Run: `cd /Users/philipkaare/git/toggljirazer && ./run.sh --headless; echo "exit=$?"`
Expected: prints the first missing-field validation error and `exit=2` (with placeholder credentials).

- [ ] **Step 5: Commit**

```bash
cd /Users/philipkaare/git/toggljirazer
git add -A
git commit -m "Update README for .NET 10 TUI and add run.sh"
```

---

## Self-Review Notes

- **Spec coverage:** framework/net10 retarget (T1), WPF removal (T1), Terminal.Gui-only dep (T1), reused Services/Models (all), shared ReportRunner (T4), full-parity TUI form (T7), Save/Generate/Open (T7), progress + log marshaling (T7), headless + redirect detection (T6), cross-platform opener (T5), static banner (T7), xunit tests for ConfigValidator/PlatformOpener/SettingsService/period/argparser/banner (T2–T7), README + run.sh (T8). All spec sections map to a task.
- **Known risk:** exact Terminal.Gui 2.4.16 widget API (T7). Mitigated by grounding in the uRobot reference and a mandatory build-green gate before proceeding.
```
