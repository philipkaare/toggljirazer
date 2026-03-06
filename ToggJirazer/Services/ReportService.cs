using System.Globalization;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

public class ReportService
{
    private static readonly HashSet<string> StandardVersionColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "Version", "Total Estimate Sum", "Worked Hours in Period", "Total Worked Hours", "Difference"
    };

    private readonly JiraService _jiraService;

    public ReportService(JiraService jiraService)
    {
        _jiraService = jiraService;
    }

    private static List<string> GetExtraColumnNames(Dictionary<string, Dictionary<string, string>>? extraColumns)
    {
        if (extraColumns == null || extraColumns.Count == 0)
            return [];
        return extraColumns.Values
            .SelectMany(d => d.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<(List<(string SheetName, List<ReportRow> Rows)> ReportSheets, List<VersionReportRow> VersionRows)> BuildReportAsync(
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

        // Fetch JIRA issues in bulk to avoid rate limiting
        var jiraIssues = await _jiraService.GetIssuesBulkAsync(allKeys);

        // Group period entries by year-month to detect multi-month periods
        var byMonth = periodEntries
            .GroupBy(e => new { e.Start.Year, e.Start.Month })
            .OrderBy(g => g.Key.Year).ThenBy(g => g.Key.Month)
            .ToList();

        List<(string SheetName, List<ReportRow> Rows)> reportSheets;
        if (byMonth.Count <= 1)
        {
            // Single month (or no entries) — use a single "Report" sheet
            reportSheets = [("Report", BuildRowsFromEntries(periodEntries, jiraIssues))];
        }
        else
        {
            // Multiple months — create one sheet per month
            reportSheets = byMonth
                .Select(g => (
                    new DateTime(g.Key.Year, g.Key.Month, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture),
                    BuildRowsFromEntries(g.ToList(), jiraIssues)
                ))
                .ToList();
        }

        // Build version report rows
        var versionRows = await BuildVersionRowsAsync(periodEntries, allEntries, jiraIssues);

        return (reportSheets, versionRows);
    }

    private static List<ReportRow> BuildRowsFromEntries(
        List<TogglTimeEntry> entries,
        Dictionary<string, JiraIssue?> jiraIssues)
    {
        var groups = entries
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

        return rows;
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

    /// <summary>
    /// Reads any extra (non-standard) columns from an existing version report file so they can be
    /// preserved when the report is regenerated. The file format is inferred from its extension
    /// (.xlsx or .csv). Returns a dictionary mapping version name to a dictionary of extra column
    /// name/value pairs. Returns an empty dictionary if the file does not exist or cannot be read.
    /// </summary>
    public Dictionary<string, Dictionary<string, string>> ReadVersionReportExtraColumns(string outputPath)
    {
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        if (!File.Exists(outputPath))
            return result;

        var ext = Path.GetExtension(outputPath).ToLowerInvariant();

        if (ext == ".xlsx")
            ReadVersionExtraColumnsFromXlsx(outputPath, result);
        else if (ext == ".csv")
            ReadVersionExtraColumnsFromCsv(outputPath, result);

        return result;
    }

    private static void ReadVersionExtraColumnsFromXlsx(string path, Dictionary<string, Dictionary<string, string>> result)
    {
        try
        {
            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.Equals("Version Report", StringComparison.OrdinalIgnoreCase));
            if (sheet == null) return;

            var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            if (lastCol == 0) return;

            var extraColIndices = new Dictionary<int, string>();
            for (int col = 1; col <= lastCol; col++)
            {
                var header = sheet.Cell(1, col).GetString();
                if (!string.IsNullOrWhiteSpace(header) && !StandardVersionColumns.Contains(header))
                    extraColIndices[col] = header;
            }

            if (extraColIndices.Count == 0) return;

            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
            for (int row = 2; row <= lastRow; row++)
            {
                var version = sheet.Cell(row, 1).GetString();
                if (string.IsNullOrWhiteSpace(version)) continue;

                var extras = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var (col, colName) in extraColIndices)
                    extras[colName] = sheet.Cell(row, col).GetString();

                result[version] = extras;
            }
        }
        catch (Exception ex) { Console.Error.WriteLine($"Warning: Could not read extra columns from '{path}': {ex.Message}"); }
    }

    private static void ReadVersionExtraColumnsFromCsv(string path, Dictionary<string, Dictionary<string, string>> result)
    {
        try
        {
            using var reader = new StreamReader(path, System.Text.Encoding.UTF8);
            var config = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true };
            using var csv = new CsvReader(reader, config);

            csv.Read();
            csv.ReadHeader();

            var extraHeaders = csv.HeaderRecord!
                .Where(h => !StandardVersionColumns.Contains(h))
                .ToList();

            if (extraHeaders.Count == 0) return;

            while (csv.Read())
            {
                var version = csv.GetField("Version") ?? string.Empty;
                if (string.IsNullOrWhiteSpace(version)) continue;

                var extras = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var header in extraHeaders)
                    extras[header] = csv.GetField(header) ?? string.Empty;

                result[version] = extras;
            }
        }
        catch (Exception ex) { Console.Error.WriteLine($"Warning: Could not read extra columns from '{path}': {ex.Message}"); }
    }

    public void WriteVersionCsv(List<VersionReportRow> versionRows, string outputPath, Dictionary<string, Dictionary<string, string>>? extraColumns = null)
    {
        using var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8);
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true
        };
        using var csv = new CsvWriter(writer, config);

        var extraColumnNames = GetExtraColumnNames(extraColumns);

        if (extraColumnNames.Count == 0)
        {
            csv.WriteRecords(versionRows);
        }
        else
        {
            // Write headers manually so extra columns follow the standard ones
            csv.WriteField("Version");
            csv.WriteField("Total Estimate Sum");
            csv.WriteField("Worked Hours in Period");
            csv.WriteField("Total Worked Hours");
            csv.WriteField("Difference");
            foreach (var colName in extraColumnNames)
                csv.WriteField(colName);
            csv.NextRecord();

            foreach (var row in versionRows)
            {
                csv.WriteField(row.Version);
                csv.WriteField(row.TotalEstimateSum);
                csv.WriteField(row.WorkedHoursInPeriod);
                csv.WriteField(row.TotalWorkedHours);
                csv.WriteField(row.Difference);

                Dictionary<string, string>? rowExtras = null;
                var hasExtras = extraColumns != null && extraColumns.TryGetValue(row.Version, out rowExtras);
                foreach (var colName in extraColumnNames)
                    csv.WriteField(hasExtras && rowExtras != null && rowExtras.TryGetValue(colName, out var v) ? v : string.Empty);

                csv.NextRecord();
            }
        }

        Console.WriteLine($"Version report written to: {Path.GetFullPath(outputPath)}");
    }

    /// <summary>
    /// Writes the report rows to an Excel (.xlsx) file at the specified path.
    /// When <paramref name="reportSheets"/> contains more than one entry (multi-month period),
    /// each entry is written to its own named worksheet. A final "Version Report" sheet is
    /// always appended. Sheets are formatted as tables with headers; numeric columns use a
    /// fixed two-decimal numeric format ("0.00"), and Excel displays the decimal separator
    /// according to the user's locale. Sums are appended at the bottom of numeric columns.
    /// When <paramref name="extraColumns"/> is provided, any extra version-report columns are
    /// appended after the standard columns and their values are merged by version name.
    /// </summary>
    public void WriteXlsx(List<(string SheetName, List<ReportRow> Rows)> reportSheets, List<VersionReportRow> versionRows, string outputPath, Dictionary<string, Dictionary<string, string>>? extraColumns = null)
    {
        using var workbook = new XLWorkbook();

        foreach (var (sheetName, rows) in reportSheets)
        {
            WriteReportSheet(workbook, sheetName, rows);
        }

        var extraColumnNames = GetExtraColumnNames(extraColumns);

        // Version Report sheet
        var versionSheet = workbook.Worksheets.Add("Version Report");
        var standardHeaders = new[] { "Version", "Total Estimate Sum", "Worked Hours in Period", "Total Worked Hours", "Difference" };
        var versionHeaders = standardHeaders.Concat(extraColumnNames).ToArray();

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

            if (extraColumnNames.Count > 0 && extraColumns != null)
            {
                var hasExtras = extraColumns.TryGetValue(row.Version, out var rowExtras);
                for (int c = 0; c < extraColumnNames.Count; c++)
                {
                    var val = hasExtras && rowExtras!.TryGetValue(extraColumnNames[c], out var v) ? v : string.Empty;
                    versionSheet.Cell(r + 2, 6 + c).Value = val;
                }
            }
        }

        // Format version sheet as Excel table with SUM totals for all numeric columns
        if (versionRows.Count > 0)
        {
            var versionTable = versionSheet.Range(1, 1, versionRows.Count + 1, versionHeaders.Length).CreateTable();
            versionTable.ShowTotalsRow = true;
            versionTable.Field("Total Estimate Sum").TotalsRowFunction = XLTotalsRowFunction.Sum;
            versionTable.Field("Worked Hours in Period").TotalsRowFunction = XLTotalsRowFunction.Sum;
            versionTable.Field("Total Worked Hours").TotalsRowFunction = XLTotalsRowFunction.Sum;
            versionTable.Field("Difference").TotalsRowFunction = XLTotalsRowFunction.Sum;
        }

        // Apply number format with 2 decimal places to numeric columns on version sheet
        for (int col = 2; col <= 5; col++)
            versionSheet.Column(col).Style.NumberFormat.Format = "0.00";

        workbook.SaveAs(outputPath);
        Console.WriteLine($"Report written to: {Path.GetFullPath(outputPath)}");
    }

    private static void WriteReportSheet(XLWorkbook workbook, string sheetName, List<ReportRow> rows)
    {
        var worksheet = workbook.Worksheets.Add(sheetName);

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
            // Store as an actual number so Excel displays the decimal separator according to the user's locale
            if (double.TryParse(row.TimeUsedDecimal, NumberStyles.Float, CultureInfo.InvariantCulture, out var decimalHours))
                worksheet.Cell(r + 2, 9).Value = decimalHours;
            else
                worksheet.Cell(r + 2, 9).Value = row.TimeUsedDecimal;
        }

        // Format as Excel table and add a totals row with SUM for the decimal hours column
        if (rows.Count > 0)
        {
            var table = worksheet.Range(1, 1, rows.Count + 1, headers.Length).CreateTable();
            table.ShowTotalsRow = true;
            table.Field("Time Used (Decimal)").TotalsRowFunction = XLTotalsRowFunction.Sum;
        }

        // Apply locale-agnostic number format with 2 decimal places to the decimal hours column
        worksheet.Column(9).Style.NumberFormat.Format = "0.00";
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
