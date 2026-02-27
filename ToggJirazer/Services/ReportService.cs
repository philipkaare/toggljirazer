using System.Globalization;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

public class ReportService
{
    private readonly JiraService _jiraService;

    public ReportService(JiraService jiraService)
    {
        _jiraService = jiraService;
    }

    public async Task<(List<ReportRow> Rows, List<VersionReportRow> VersionRows)> BuildReportAsync(
        List<TogglTimeEntry> periodEntries,
        List<TogglTimeEntry> allEntries)
    {
        // Collect all unique JIRA keys from both entry sets
        var allKeys = periodEntries.Concat(allEntries)
            .Select(e => ExtractJiraKey(e.Description))
            .Where(k => !string.IsNullOrEmpty(k))
            .Select(k => k!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        Console.WriteLine($"Found {allKeys.Count} unique Jira issue keys to look up.");

        // Fetch JIRA issues (with caching)
        var jiraIssues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in allKeys)
        {
            Console.WriteLine($"  Looking up Jira issue: {key}");
            jiraIssues[key] = await _jiraService.GetIssueAsync(key);
        }

        // Build main report rows (grouped by JIRA key and user, using period entries)
        var groups = periodEntries
            .Select(e => new { Entry = e, Key = ExtractJiraKey(e.Description) })
            .Where(x => !string.IsNullOrEmpty(x.Key))
            .GroupBy(x => new { x.Key, x.Entry.User, x.Entry.Email });

        var rows = new List<ReportRow>();
        foreach (var group in groups)
        {
            var jiraKey = group.Key.Key!;
            var issue = jiraIssues.TryGetValue(jiraKey, out var ji) ? ji : null;

            var totalMs = group.Sum(x => x.Entry.Duration);
            var totalSeconds = totalMs / 1000;
            var timeSpan = TimeSpan.FromSeconds(totalSeconds);

            var startDate = group.Min(x => x.Entry.Start).ToString("yyyy-MM-dd");

            rows.Add(new ReportRow
            {
                IssueType = issue?.IssueType ?? string.Empty,
                Key = jiraKey,
                Summary = issue?.Summary ?? string.Empty,
                Budget = issue?.Budget ?? string.Empty,
                Account = issue?.Account ?? string.Empty,
                Person = group.Key.User,
                StartDate = startDate,
                TimeUsedHHMM = $"{(int)timeSpan.TotalHours:D2}:{timeSpan.Minutes:D2}",
                TimeUsedDecimal = Math.Round(timeSpan.TotalHours, 2).ToString("F2", CultureInfo.InvariantCulture)
            });
        }

        rows.Sort((a, b) =>
        {
            int cmp = string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase);
            return cmp != 0 ? cmp : string.Compare(a.Person, b.Person, StringComparison.OrdinalIgnoreCase);
        });

        // Build version report rows
        var versionRows = await BuildVersionRowsAsync(periodEntries, allEntries, jiraIssues);

        return (rows, versionRows);
    }

    private async Task<List<VersionReportRow>> BuildVersionRowsAsync(
        List<TogglTimeEntry> periodEntries,
        List<TogglTimeEntry> allEntries,
        Dictionary<string, JiraIssue?> jiraIssues)
    {
        // Collect fix versions referenced by jira issues linked from toggl
        var fixVersionSet = jiraIssues.Values
            .Where(i => i != null)
            .SelectMany(i => i!.FixVersions)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v)
            .ToList();

        if (fixVersionSet.Count == 0)
            return new List<VersionReportRow>();

        // Build a lookup: jira key -> fix versions (from period-linked issues)
        var keyToVersions = jiraIssues
            .Where(kv => kv.Value != null && kv.Value.FixVersions.Count > 0)
            .ToDictionary(kv => kv.Key, kv => kv.Value!.FixVersions, StringComparer.OrdinalIgnoreCase);

        // Sum period hours per fix version
        var periodHoursByVersion = SumHoursByVersion(periodEntries, keyToVersions);

        // Sum all-time hours per fix version
        var totalHoursByVersion = SumHoursByVersion(allEntries, keyToVersions);

        Console.WriteLine($"Building version report for {fixVersionSet.Count} fix version(s).");

        var versionRows = new List<VersionReportRow>();
        foreach (var version in fixVersionSet)
        {
            Console.WriteLine($"  Fetching Jira issues for fix version: {version}");
            var versionIssues = await _jiraService.GetIssuesByFixVersionAsync(version);
            var estimateSum = versionIssues.Sum(i => i.Estimate ?? 0.0);

            periodHoursByVersion.TryGetValue(version, out var periodHours);
            totalHoursByVersion.TryGetValue(version, out var totalHours);
            var difference = estimateSum - totalHours;

            versionRows.Add(new VersionReportRow
            {
                Version = version,
                TotalEstimateSum = Math.Round(estimateSum, 2),
                WorkedHoursInPeriod = Math.Round(periodHours, 2),
                TotalWorkedHours = Math.Round(totalHours, 2),
                Difference = Math.Round(difference, 2)
            });
        }

        return versionRows;
    }

    private static Dictionary<string, double> SumHoursByVersion(
        List<TogglTimeEntry> entries,
        Dictionary<string, List<string>> keyToVersions)
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            var key = ExtractJiraKey(entry.Description);
            if (key == null) continue;
            if (!keyToVersions.TryGetValue(key, out var versions)) continue;
            var hours = entry.Duration / 1000.0 / 3600.0;
            foreach (var version in versions)
            {
                result.TryGetValue(version, out var existing);
                result[version] = existing + hours;
            }
        }
        return result;
    }

    public void WriteCsv(List<ReportRow> rows, string outputPath)
    {
        using var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8);
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true
        };
        using var csv = new CsvWriter(writer, config);
        csv.WriteRecords(rows);
        Console.WriteLine($"Report written to: {Path.GetFullPath(outputPath)}");
    }

    public void WriteVersionCsv(List<VersionReportRow> versionRows, string outputPath)
    {
        using var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8);
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true
        };
        using var csv = new CsvWriter(writer, config);
        csv.WriteRecords(versionRows);
        Console.WriteLine($"Version report written to: {Path.GetFullPath(outputPath)}");
    }

    /// <summary>
    /// Writes the report rows to an Excel (.xlsx) file at the specified path,
    /// with a second sheet containing the version summary report.
    /// </summary>
    public void WriteXlsx(List<ReportRow> rows, List<VersionReportRow> versionRows, string outputPath)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Report");

        // Write header row
        var headers = new[]
        {
            "Issue Type", "Key", "Summary", "Budget", "Account",
            "Person", "Start Date", "Time Used (HH:MM)", "Time Used (Decimal)"
        };
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cell(1, i + 1).Value = headers[i];
        }

        // Write data rows
        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            worksheet.Cell(r + 2, 1).Value = row.IssueType;
            worksheet.Cell(r + 2, 2).Value = row.Key;
            worksheet.Cell(r + 2, 3).Value = row.Summary;
            worksheet.Cell(r + 2, 4).Value = row.Budget;
            worksheet.Cell(r + 2, 5).Value = row.Account;
            worksheet.Cell(r + 2, 6).Value = row.Person;
            worksheet.Cell(r + 2, 7).Value = row.StartDate;
            worksheet.Cell(r + 2, 8).Value = row.TimeUsedHHMM;
            worksheet.Cell(r + 2, 9).Value = row.TimeUsedDecimal;
        }

        // Second sheet: version report
        var versionSheet = workbook.Worksheets.Add("Version Report");
        var versionHeaders = new[]
        {
            "Version", "Total Estimate Sum", "Worked Hours in Period",
            "Total Worked Hours", "Difference"
        };
        for (int i = 0; i < versionHeaders.Length; i++)
        {
            versionSheet.Cell(1, i + 1).Value = versionHeaders[i];
        }

        for (int r = 0; r < versionRows.Count; r++)
        {
            var row = versionRows[r];
            versionSheet.Cell(r + 2, 1).Value = row.Version;
            versionSheet.Cell(r + 2, 2).Value = row.TotalEstimateSum;
            versionSheet.Cell(r + 2, 3).Value = row.WorkedHoursInPeriod;
            versionSheet.Cell(r + 2, 4).Value = row.TotalWorkedHours;
            versionSheet.Cell(r + 2, 5).Value = row.Difference;
        }

        workbook.SaveAs(outputPath);
        Console.WriteLine($"Report written to: {Path.GetFullPath(outputPath)}");
    }

    /// <summary>
    /// Extracts a JIRA issue key (e.g. PROJECT-123) from a time entry description.
    /// The key must match the pattern LETTERS-DIGITS optionally followed by whitespace/punctuation.
    /// </summary>
    private static string? ExtractJiraKey(string description)
    {
        if (string.IsNullOrWhiteSpace(description))
            return null;

        var match = System.Text.RegularExpressions.Regex.Match(
            description.Trim(),
            @"^([A-Z][A-Z0-9]+-\d+)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        return match.Success ? match.Groups[1].Value.ToUpperInvariant() : null;
    }
}
