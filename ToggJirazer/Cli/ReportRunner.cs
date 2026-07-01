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
