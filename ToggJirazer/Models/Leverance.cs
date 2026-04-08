namespace ToggJirazer.Models;

/// <summary>
/// Represents a "Leverance" (delivery) — a grouping of Jira issues that share
/// a common fix version. The leverance is anchored by a Jira issue of type
/// "Leverance" and includes all issues with the same fix version.
/// </summary>
public class Leverance
{
    /// <summary>Fix version that defines this leverance.</summary>
    public string FixVersion { get; set; } = string.Empty;

    /// <summary>The budget value from the anchor "Leverance" issue (if any).</summary>
    public string? Budget { get; set; }

    /// <summary>All issues belonging to this leverance (including the anchor).</summary>
    public List<ReportRow> Issues { get; set; } = new();

    /// <summary>Total worked hours (decimal) across all issues in this leverance.</summary>
    public double TotalHours => Issues.Sum(r =>
        double.TryParse(r.TimeUsedDecimal, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var h) ? h : 0);
}
