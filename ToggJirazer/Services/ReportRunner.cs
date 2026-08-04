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
