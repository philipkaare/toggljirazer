using ToggJirazer.Models;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Cli;

/// <summary>
/// Runs the end-to-end report pipeline: resolve period, fetch Toggl entries,
/// enrich with Jira, and write the xlsx/csv output using the shared pipeline.
/// Pipeline progress goes to Console, which the TUI redirects to its log pane.
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

    public static async Task<bool> RunAsync(AppConfig config, IJiraTools jiraTools, Action<string> log)
    {
        var (startDate, endDate) = ResolvePeriod(config.Report, DateTime.Today, log);

        var reportWritten = await Services.ReportRunner.RunAsync(config, jiraTools, startDate, endDate,
            cacheEnabled: false, cacheWritePermissionPrompt: _ => false);
        log("Done.");
        return reportWritten;
    }
}
