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

    public async Task<List<ReportRow>> BuildReportAsync(List<TogglTimeEntry> entries)
    {
        // Group entries by JIRA key (extracted from description) and user
        var groups = entries
            .Select(e => new { Entry = e, Key = ExtractJiraKey(e.Description) })
            .Where(x => !string.IsNullOrEmpty(x.Key))
            .GroupBy(x => new { x.Key, x.Entry.User, x.Entry.Email });

        // Collect all unique JIRA keys to look up
        var jiraKeys = groups.Select(g => g.Key.Key!).Distinct().ToList();
        Console.WriteLine($"Found {jiraKeys.Count} unique Jira issue keys to look up.");

        // Fetch JIRA issues (with caching)
        var jiraIssues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in jiraKeys)
        {
            Console.WriteLine($"  Looking up Jira issue: {key}");
            jiraIssues[key] = await _jiraService.GetIssueAsync(key);
        }

        var rows = new List<ReportRow>();
        foreach (var group in groups)
        {
            var jiraKey = group.Key.Key!;
            var issue = jiraIssues.TryGetValue(jiraKey, out var ji) ? ji : null;

            // Total duration in milliseconds
            var totalMs = group.Sum(x => x.Entry.Duration);
            var totalSeconds = totalMs / 1000;
            var timeSpan = TimeSpan.FromSeconds(totalSeconds);

            // Earliest start date in the group
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

        // Sort by Key then Person
        rows.Sort((a, b) =>
        {
            int cmp = string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase);
            return cmp != 0 ? cmp : string.Compare(a.Person, b.Person, StringComparison.OrdinalIgnoreCase);
        });

        return rows;
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

    /// <summary>
    /// Writes the report rows to an Excel (.xlsx) file at the specified path.
    /// </summary>
    public void WriteXlsx(List<ReportRow> rows, string outputPath)
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
