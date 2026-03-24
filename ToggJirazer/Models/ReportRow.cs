using CsvHelper.Configuration.Attributes;

namespace ToggJirazer.Models;

public class ReportRow
{
    [Name("Issue Type")]
    public string IssueType { get; set; } = string.Empty;

    /// <summary>Fix versions (Jira releases) for this issue. Not included in CSV output.</summary>
    [Ignore]
    public List<string> FixVersions { get; set; } = new();

    [Name("Key")]
    public string Key { get; set; } = string.Empty;

    [Name("Summary")]
    public string Summary { get; set; } = string.Empty;

    [Name("Budget")]
    public string Budget { get; set; } = string.Empty;

    [Name("Account")]
    public string Account { get; set; } = string.Empty;

    [Name("Person")]
    public string Person { get; set; } = string.Empty;

    [Name("Start Date")]
    public string StartDate { get; set; } = string.Empty;

    [Name("Time Used (HH:MM)")]
    public string TimeUsedHHMM { get; set; } = string.Empty;

    [Name("Time Used (Decimal)")]
    public string TimeUsedDecimal { get; set; } = string.Empty;
}
