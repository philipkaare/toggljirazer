using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

public class ReportService
{
    private static readonly string[] StandardVersionHeaders =
    [
        "Version",
        "Total Estimate Sum",
        "Worked Hours in Period",
        "Total Worked Hours",
        "Difference"
    ];

    private static readonly HashSet<string> StandardVersionColumns = new(StandardVersionHeaders, StringComparer.OrdinalIgnoreCase);
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
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<(
        List<(string SheetName, List<ReportRow> Rows)> ReportSheets,
        List<VersionReportRow> VersionRows,
        List<Leverance> Leverances,
        Dictionary<string, double> CategoryConsumedHours,
        Dictionary<(int Year, int Week), Dictionary<string, double>> WeeklyCategoryConsumedHours)> BuildReportAsync(
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
            // Single month (or no entries) — use a single "Rapport" sheet
            reportSheets = [("Rapport", BuildRowsFromEntries(periodEntries, jiraIssues))];
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

        // Build leverances from all-time rows
        var allTimeRows = BuildRowsFromEntries(allEntries, jiraIssues);
        var leverances = BuildLeverances(allTimeRows, jiraIssues);
        Console.WriteLine($"Built {leverances.Count} leverance(s).");

        var (categoryConsumedHours, weeklyCategoryConsumedHours) = BuildCategoryConsumption(periodEntries, jiraIssues);
        return (reportSheets, versionRows, leverances, categoryConsumedHours, weeklyCategoryConsumedHours);
    }

    /// <summary>
    /// Builds leverance groupings. A leverance is anchored by tasks of type "Leverance".
    /// All tasks that share the same first fix version are grouped under that leverance.
    /// </summary>
    private static List<Leverance> BuildLeverances(List<ReportRow> allRows, Dictionary<string, JiraIssue?> jiraIssues)
    {
        // Find all fix versions that have an anchor issue of type "Leverance"
        var leveranceVersions = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in jiraIssues)
        {
            var issue = kv.Value;
            if (issue == null) continue;
            if (!issue.IssueType.Equals("Leverance", StringComparison.OrdinalIgnoreCase)) continue;
            foreach (var fv in issue.FixVersions)
            {
                if (!leveranceVersions.ContainsKey(fv))
                    leveranceVersions[fv] = issue.Budget;
            }
        }

        // Group all rows by their first fix version
        var rowsByVersion = new Dictionary<string, List<ReportRow>>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in allRows)
        {
            var firstVersion = row.FixVersions.FirstOrDefault();
            if (string.IsNullOrEmpty(firstVersion)) continue;
            if (!leveranceVersions.ContainsKey(firstVersion)) continue;

            if (!rowsByVersion.TryGetValue(firstVersion, out var list))
            {
                list = new List<ReportRow>();
                rowsByVersion[firstVersion] = list;
            }
            list.Add(row);
        }

        var leverances = new List<Leverance>();
        foreach (var (version, budget) in leveranceVersions)
        {
            var lev = new Leverance
            {
                FixVersion = version,
                Budget = budget,
                Issues = rowsByVersion.TryGetValue(version, out var rows) ? rows : new List<ReportRow>()
            };
            leverances.Add(lev);
        }

        leverances.Sort((a, b) => string.Compare(a.FixVersion, b.FixVersion, StringComparison.OrdinalIgnoreCase));
        return leverances;
    }

    private static List<ReportRow> BuildRowsFromEntries(
        List<TogglTimeEntry> entries,
        Dictionary<string, JiraIssue?> jiraIssues)
    {
        var versionToBudget = BuildLeveranceBudgetByVersion(jiraIssues);

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
                Budget = ResolveIssueBudget(issue, versionToBudget, string.Empty),
                Account = issue?.Account ?? string.Empty,
                Person = group.Key.User,
                StartDate = startDate,
                TimeUsedHHMM = $"{(int)timeSpan.TotalHours:D2}:{timeSpan.Minutes:D2}",
                TimeUsedDecimal = Math.Round(timeSpan.TotalHours, 2).ToString("F2", CultureInfo.InvariantCulture),
                FixVersions = issue?.FixVersions ?? new List<string>()
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
        else
            Console.Error.WriteLine($"Warning: Unrecognized report file extension '{ext}' for path '{outputPath}'. Extra columns will not be preserved.");

        return result;
    }

    private static void ReadVersionExtraColumnsFromXlsx(string path, Dictionary<string, Dictionary<string, string>> result)
    {
        try
        {
            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.Equals("Versionsrapport", StringComparison.OrdinalIgnoreCase)
                    || ws.Name.Equals("Version Report", StringComparison.OrdinalIgnoreCase));
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
    public void WriteXlsx(
        List<(string SheetName, List<ReportRow> Rows)> reportSheets,
        List<VersionReportRow> versionRows,
        List<Leverance> leverances,
        Dictionary<string, double> categoryConsumedHours,
        Dictionary<(int Year, int Week), Dictionary<string, double>> weeklyCategoryConsumedHours,
        string outputPath,
        Dictionary<string, Dictionary<string, string>>? extraColumns = null)
    {
        // Read existing data before creating new workbook
        var existingYearlyBudgets = ReadExistingOpsummeringYearlyBudgets(outputPath);
        var existingPlanData = ReadExistingPlanData(outputPath);

        using var workbook = new XLWorkbook();

        foreach (var (sheetName, rows) in reportSheets)
        {
            WriteReportSheet(workbook, sheetName, rows);
        }

        var extraColumnNames = GetExtraColumnNames(extraColumns);

        // Versionsrapport sheet (Danish title)
        var versionSheet = workbook.Worksheets.Add("Versionsrapport");
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
        versionSheet.Columns().AdjustToContents();

        // Opsummering sheet
        WriteOpsummeringSheet(workbook, categoryConsumedHours, weeklyCategoryConsumedHours, existingYearlyBudgets);

        // Plan sheet
        WritePlanSheet(workbook, leverances, existingPlanData);

        workbook.SaveAs(outputPath);
        Console.WriteLine($"Report written to: {Path.GetFullPath(outputPath)}");
    }

    /// <summary>
    /// Reads existing "Budget pr. år" values from the Opsummering sheet of an existing XLSX file.
    /// Returns a dictionary mapping category name → yearly budget hours.
    /// </summary>
    private static Dictionary<string, double> ReadExistingOpsummeringYearlyBudgets(string path)
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return result;

        try
        {
            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheets
                .FirstOrDefault(s => s.Name.Equals("Opsummering", StringComparison.OrdinalIgnoreCase));
            if (ws == null) return result;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            if (lastCol == 0) return result;

            int headerRow = -1;
            int categoryCol = -1;
            int yearlyBudgetCol = -1;

            for (int row = 1; row <= Math.Min(lastRow, 10); row++)
            {
                int foundCategoryCol = -1;
                int foundYearlyBudgetCol = -1;

                for (int col = 1; col <= lastCol; col++)
                {
                    var header = ws.Cell(row, col).GetString();
                    if (header.Equals("Kategori", StringComparison.OrdinalIgnoreCase))
                        foundCategoryCol = col;
                    else if (header.Equals("Budget pr. år", StringComparison.OrdinalIgnoreCase))
                        foundYearlyBudgetCol = col;
                }

                if (foundCategoryCol > 0 && foundYearlyBudgetCol > 0)
                {
                    headerRow = row;
                    categoryCol = foundCategoryCol;
                    yearlyBudgetCol = foundYearlyBudgetCol;
                    break;
                }
            }

            if (headerRow < 0) return result;

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                var category = ws.Cell(row, categoryCol).GetString();
                if (string.IsNullOrWhiteSpace(category)) break;
                if (category.Equals("Uge", StringComparison.OrdinalIgnoreCase)) break;

                var cell = ws.Cell(row, yearlyBudgetCol);
                if (cell.TryGetValue<double>(out var yearlyBudget))
                    result[category] = yearlyBudget;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: Could not read existing Opsummering yearly budget data: {ex.Message}");
        }

        return result;
    }

    private static (Dictionary<string, double> CategoryConsumedHours, Dictionary<(int Year, int Week), Dictionary<string, double>> WeeklyCategoryConsumedHours)
        BuildCategoryConsumption(List<TogglTimeEntry> entries, Dictionary<string, JiraIssue?> jiraIssues)
    {
        var versionToBudget = BuildLeveranceBudgetByVersion(jiraIssues);

        string ResolveCategory(JiraIssue? issue)
        {
            return ResolveIssueBudget(issue, versionToBudget, "(Intet budget)");
        }

        var categoryConsumedHours = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        var weeklyCategoryConsumedHours = new Dictionary<(int Year, int Week), Dictionary<string, double>>();

        foreach (var entry in entries)
        {
            var jiraKey = ExtractJiraKey(entry.Description);
            var issue = jiraKey != null && jiraIssues.TryGetValue(jiraKey, out var ji) ? ji : null;
            var category = ResolveCategory(issue);
            var hours = entry.Duration / 1000.0 / 3600.0;

            categoryConsumedHours.TryGetValue(category, out var currentCategoryHours);
            categoryConsumedHours[category] = currentCategoryHours + hours;

            var weekKey = (
                Year: ISOWeek.GetYear(entry.Start.Date),
                Week: ISOWeek.GetWeekOfYear(entry.Start.Date));
            if (!weeklyCategoryConsumedHours.TryGetValue(weekKey, out var byCategory))
            {
                byCategory = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                weeklyCategoryConsumedHours[weekKey] = byCategory;
            }

            byCategory.TryGetValue(category, out var currentWeekCategoryHours);
            byCategory[category] = currentWeekCategoryHours + hours;
        }

        return (categoryConsumedHours, weeklyCategoryConsumedHours);
    }

    private static Dictionary<string, string> BuildLeveranceBudgetByVersion(Dictionary<string, JiraIssue?> jiraIssues)
    {
        var versionToBudget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var issue in jiraIssues.Values)
        {
            if (issue == null) continue;
            if (!issue.IssueType.Equals("Leverance", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.IsNullOrWhiteSpace(issue.Budget)) continue;

            foreach (var version in issue.FixVersions.Where(v => !string.IsNullOrWhiteSpace(v)))
            {
                if (!versionToBudget.ContainsKey(version))
                    versionToBudget[version] = issue.Budget;
            }
        }

        return versionToBudget;
    }

    private static string ResolveIssueBudget(
        JiraIssue? issue,
        Dictionary<string, string> versionToBudget,
        string fallbackWhenMissing)
    {
        var firstVersion = issue?.FixVersions.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        if (!string.IsNullOrWhiteSpace(firstVersion) && versionToBudget.TryGetValue(firstVersion, out var overriddenBudget))
            return overriddenBudget;

        if (!string.IsNullOrWhiteSpace(issue?.Budget))
            return issue.Budget!;

        return fallbackWhenMissing;
    }

    private static void WriteOpsummeringSheet(
        XLWorkbook workbook,
        Dictionary<string, double> categoryConsumedHours,
        Dictionary<(int Year, int Week), Dictionary<string, double>> weeklyCategoryConsumedHours,
        Dictionary<string, double> existingYearlyBudgets)
    {
        var ws = workbook.Worksheets.Add("Opsummering");

        var categories = categoryConsumedHours.Keys
            .Union(existingYearlyBudgets.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(c => c, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (categories.Count == 0)
            categories.Add("(Intet budget)");

        // Tabel 1
        var table1Headers = new[] { "Kategori", "Budget pr. år", "Budget pr. måned", "Forbrugt", "Rest", "Rest pr. måned" };
        for (int i = 0; i < table1Headers.Length; i++)
            ws.Cell(1, i + 1).Value = table1Headers[i];

        for (int i = 0; i < categories.Count; i++)
        {
            var row = i + 2;
            var category = categories[i];
            ws.Cell(row, 1).Value = category;

            if (existingYearlyBudgets.TryGetValue(category, out var yearlyBudget))
                ws.Cell(row, 2).Value = Math.Round(yearlyBudget, 2);

            categoryConsumedHours.TryGetValue(category, out var consumedHours);
            ws.Cell(row, 4).Value = Math.Round(consumedHours, 2);

            ws.Cell(row, 3).FormulaA1 = $"IF(B{row}=\"\",\"\",B{row}/12)";
            ws.Cell(row, 5).FormulaA1 = $"IF(B{row}=\"\",\"\",B{row}-D{row})";
            ws.Cell(row, 6).FormulaA1 = $"IF(E{row}=\"\",\"\",E{row}/12)";
        }

        var table1LastRow = categories.Count + 1;
        var table1 = ws.Range(1, 1, table1LastRow, table1Headers.Length).CreateTable("OpsummeringKategori");
        table1.Theme = XLTableTheme.TableStyleMedium2;
        ws.Range(2, 2, table1LastRow, 6).Style.NumberFormat.Format = "0.00";

        // Tabel 2
        var startRowTable2 = table1LastRow + 3;
        ws.Cell(startRowTable2, 1).Value = "Uge";
        for (int i = 0; i < categories.Count; i++)
            ws.Cell(startRowTable2, i + 2).Value = categories[i];

        var weeks = weeklyCategoryConsumedHours.Keys
            .OrderBy(w => w.Year)
            .ThenBy(w => w.Week)
            .ToList();
        var hasMultipleYears = weeks.Select(w => w.Year).Distinct().Skip(1).Any();

        for (int i = 0; i < weeks.Count; i++)
        {
            var row = startRowTable2 + i + 1;
            var week = weeks[i];
            ws.Cell(row, 1).Value = hasMultipleYears ? $"{week.Year}-{week.Week:D2}" : week.Week;

            weeklyCategoryConsumedHours.TryGetValue(week, out var weekByCategory);
            for (int c = 0; c < categories.Count; c++)
            {
                var category = categories[c];
                var value = 0.0;
                if (weekByCategory != null && weekByCategory.TryGetValue(category, out var weeklyHours))
                    value = weeklyHours;
                ws.Cell(row, c + 2).Value = Math.Round(value, 2);
            }
        }

        var table2LastRow = startRowTable2 + weeks.Count;
        var table2LastCol = categories.Count + 1;
        var table2 = ws.Range(startRowTable2, 1, table2LastRow, table2LastCol).CreateTable("OpsummeringUgeforbrug");
        table2.Theme = XLTableTheme.TableStyleMedium9;
        if (table2LastCol >= 2)
            ws.Range(startRowTable2 + 1, 2, table2LastRow, table2LastCol).Style.NumberFormat.Format = "0.00";

        ws.Columns().AdjustToContents();
    }

    /// <summary>
    /// Reads existing Plan sheet data from an XLSX file so that user-edited week cells
    /// can be preserved across regenerations. Returns a dictionary mapping leverance name
    /// to a dictionary of week number → cell value.
    /// </summary>
    private static Dictionary<string, Dictionary<int, string>> ReadExistingPlanData(string path)
    {
        var result = new Dictionary<string, Dictionary<int, string>>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return result;

        try
        {
            using var workbook = new XLWorkbook(path);
            var ws = workbook.Worksheets
                .FirstOrDefault(s => s.Name.Equals("Plan", StringComparison.OrdinalIgnoreCase));
            if (ws == null) return result;

            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

            // Read week numbers from header row (row 1, starting at col 3)
            var weekByCol = new Dictionary<int, int>();
            for (int col = 3; col <= lastCol; col++)
            {
                var header = ws.Cell(1, col).GetString();
                if (int.TryParse(header, out var weekNum))
                    weekByCol[col] = weekNum;
            }

            // Read each leverance row
            for (int row = 2; row <= lastRow; row++)
            {
                var name = ws.Cell(row, 1).GetString();
                if (string.IsNullOrWhiteSpace(name)) continue;

                var weekValues = new Dictionary<int, string>();
                foreach (var (col, weekNum) in weekByCol)
                {
                    var val = ws.Cell(row, col).GetString();
                    if (!string.IsNullOrEmpty(val))
                        weekValues[weekNum] = val;
                }

                result[name] = weekValues;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: Could not read existing Plan data: {ex.Message}");
        }

        return result;
    }

    private static void WritePlanSheet(XLWorkbook workbook, List<Leverance> leverances, Dictionary<string, Dictionary<int, string>> existingPlanData)
    {
        var ws = workbook.Worksheets.Add("Plan");

        // Determine week range
        var today = DateTime.Today;
        var currentWeek = ISOWeek.GetWeekOfYear(today);

        // Collect all week numbers from existing plan data
        var existingWeeks = existingPlanData.Values
            .SelectMany(d => d.Keys)
            .Distinct()
            .ToHashSet();

        // Build week list: start from week 1, go through at least current week
        int maxWeek = Math.Max(currentWeek, existingWeeks.Count > 0 ? existingWeeks.Max() : 0);

        var weeks = Enumerable.Range(1, maxWeek).ToList();

        // Headers
        ws.Cell(1, 1).Value = "Leverance";
        ws.Cell(1, 2).Value = "Total timer";
        ws.Range(1, 1, 1, 2).Style.Font.Bold = true;

        for (int w = 0; w < weeks.Count; w++)
        {
            var col = w + 3;
            ws.Cell(1, col).Value = weeks[w];
            ws.Cell(1, col).Style.Font.Bold = true;

            // Mark current week column with green background
            if (weeks[w] == currentWeek)
            {
                ws.Cell(1, col).Style.Fill.BackgroundColor = XLColor.LightGreen;
            }
        }

        // Write leverance rows
        for (int i = 0; i < leverances.Count; i++)
        {
            var lev = leverances[i];
            var row = i + 2;

            ws.Cell(row, 1).Value = lev.FixVersion;
            ws.Cell(row, 2).Value = Math.Round(lev.TotalHours, 2);

            // Restore existing week cell values
            if (existingPlanData.TryGetValue(lev.FixVersion, out var weekValues))
            {
                for (int w = 0; w < weeks.Count; w++)
                {
                    if (weekValues.TryGetValue(weeks[w], out var val) && !string.IsNullOrEmpty(val))
                    {
                        var col = w + 3;
                        // Try to preserve numeric values as numbers
                        if (double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var numVal))
                            ws.Cell(row, col).Value = numVal;
                        else
                            ws.Cell(row, col).Value = val;
                    }
                }
            }

            // Mark current week column cells with green background
            var currentWeekColIndex = weeks.IndexOf(currentWeek);
            if (currentWeekColIndex >= 0)
            {
                ws.Cell(row, currentWeekColIndex + 3).Style.Fill.BackgroundColor = XLColor.LightGreen;
            }
        }

        ws.Column(2).Style.NumberFormat.Format = "0.00";
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
        worksheet.Columns().AdjustToContents();
    }

    /// <summary>
    /// Extracts a JIRA issue key (e.g. PROJECT-123) from a time entry description.
    /// Only the first space-delimited token is considered when matching.
    /// </summary>
    private static string? ExtractJiraKey(string description)
    {
        if (string.IsNullOrWhiteSpace(description))
            return null;

        var tokens = description.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0)
            return null;

        var firstToken = tokens[0];

        var match = System.Text.RegularExpressions.Regex.Match(
            firstToken,
            @"^([A-Z][A-Z0-9]+-\d+)",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        return match.Success ? match.Groups[1].Value.ToUpperInvariant() : null;
    }
}
