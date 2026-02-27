using Microsoft.Extensions.Configuration;
using ToggJirazer.Models;
using ToggJirazer.Services;

// Load configuration
var config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
    .AddJsonFile("appsettings.development.json", optional: true, reloadOnChange: false)
    .Build();

var appConfig = new AppConfig();
config.Bind(appConfig);

// Validate required config
if (string.IsNullOrWhiteSpace(appConfig.Toggl.ApiToken))
{
    Console.Error.WriteLine("Error: 'Toggl:ApiToken' is missing in appsettings.json.");
    Console.Error.WriteLine("Remedy: Set your Toggl API token. You can find it at https://track.toggl.com/profile.");
    return 1;
}

if (appConfig.Toggl.WorkspaceId == 0)
{
    Console.Error.WriteLine("Error: 'Toggl:WorkspaceId' is missing or zero in appsettings.json.");
    Console.Error.WriteLine("Remedy: Set your Toggl workspace ID. You can find it in your Toggl workspace settings.");
    return 1;
}

if (appConfig.Toggl.OrganizationId == 0)
{
    Console.Error.WriteLine("Error: 'Toggl:OrganizationId' is missing or zero in appsettings.json.");
    Console.Error.WriteLine("Remedy: Set your Toggl organization ID. You can find it in your Toggl organization settings.");
    return 1;
}

if (appConfig.Toggl.ProjectId == 0)
{
    Console.Error.WriteLine("Error: 'Toggl:ProjectId' is missing or zero in appsettings.json.");
    Console.Error.WriteLine("Remedy: Set the Toggl project ID you want to report on.");
    return 1;
}

if (string.IsNullOrWhiteSpace(appConfig.Jira.BaseUrl))
{
    Console.Error.WriteLine("Error: 'Jira:BaseUrl' is missing in appsettings.json.");
    Console.Error.WriteLine("Remedy: Set your Jira base URL, e.g. https://yourcompany.atlassian.net");
    return 1;
}

if (string.IsNullOrWhiteSpace(appConfig.Jira.UserEmail))
{
    Console.Error.WriteLine("Error: 'Jira:UserEmail' is missing in appsettings.json.");
    Console.Error.WriteLine("Remedy: Set the email address associated with your Jira account.");
    return 1;
}

if (string.IsNullOrWhiteSpace(appConfig.Jira.ApiToken))
{
    Console.Error.WriteLine("Error: 'Jira:ApiToken' is missing in appsettings.json.");
    Console.Error.WriteLine("Remedy: Create an API token at https://id.atlassian.com/manage-profile/security/api-tokens");
    return 1;
}

// Determine the reporting period
DateTime startDate;
DateTime endDate;

if (!string.IsNullOrWhiteSpace(appConfig.Report.StartDate) &&
    !string.IsNullOrWhiteSpace(appConfig.Report.EndDate))
{
    if (!DateTime.TryParse(appConfig.Report.StartDate, out startDate))
    {
        Console.Error.WriteLine($"Error: 'Report:StartDate' value '{appConfig.Report.StartDate}' is not a valid date.");
        Console.Error.WriteLine("Remedy: Use the format yyyy-MM-dd, e.g. 2024-01-01");
        return 1;
    }
    if (!DateTime.TryParse(appConfig.Report.EndDate, out endDate))
    {
        Console.Error.WriteLine($"Error: 'Report:EndDate' value '{appConfig.Report.EndDate}' is not a valid date.");
        Console.Error.WriteLine("Remedy: Use the format yyyy-MM-dd, e.g. 2024-01-31");
        return 1;
    }
}
else
{
    // Default to current month
    var now = DateTime.Today;
    startDate = new DateTime(now.Year, now.Month, 1);
    endDate = startDate.AddMonths(1).AddDays(-1);
    Console.WriteLine($"No period specified in config. Defaulting to current month: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
}

Console.WriteLine($"Report period: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");
Console.WriteLine($"Toggl workspace: {appConfig.Toggl.WorkspaceId}, project: {appConfig.Toggl.ProjectId}");
Console.WriteLine($"Jira base URL: {appConfig.Jira.BaseUrl}");
Console.WriteLine();

try
{
    // Fetch Toggl data
    using var togglService = new TogglService(appConfig.Toggl);
    var entries = await togglService.GetDetailedReportAsync(startDate, endDate);

    if (entries.Count == 0)
    {
        Console.WriteLine("No Toggl time entries found for the specified period and project.");
        return 0;
    }

    Console.WriteLine();

    // Fetch all-time Toggl entries for the version report totals
    Console.WriteLine("Fetching all-time Toggl entries for version totals...");
    var allEntries = await togglService.GetAllEntriesAsync();
    Console.WriteLine();

    // Build report
    using var jiraService = new JiraService(appConfig.Jira);
    var reportService = new ReportService(jiraService);
    var (rows, versionRows) = await reportService.BuildReportAsync(entries, allEntries);

    Console.WriteLine();
    Console.WriteLine($"Report contains {rows.Count} rows.");
    Console.WriteLine($"Version report contains {versionRows.Count} rows.");
    Console.WriteLine();

    // Write report in configured format
    var format = appConfig.Report.Format?.Trim().ToLowerInvariant();
    if (format == "xlsx")
    {
        reportService.WriteXlsx(rows, versionRows, appConfig.Report.OutputFile);
    }
    else
    {
        if (!string.IsNullOrEmpty(format) && format != "csv")
        {
            Console.Error.WriteLine($"Warning: Unrecognized format '{appConfig.Report.Format}'. Defaulting to CSV.");
        }
        reportService.WriteCsv(rows, appConfig.Report.OutputFile);

        // Write version report to a second CSV file
        var outputPath = appConfig.Report.OutputFile;
        var versionOutputPath = Path.Combine(
            Path.GetDirectoryName(outputPath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(outputPath) + "_versions" + Path.GetExtension(outputPath));
        reportService.WriteVersionCsv(versionRows, versionOutputPath);
    }
}
catch (InvalidOperationException ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Unexpected error: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace);
    return 1;
}

return 0;
