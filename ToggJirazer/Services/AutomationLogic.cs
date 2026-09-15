using System.Globalization;

namespace ToggJirazer.Services;

/// <summary>
/// Pure decision logic for the automated monthly report email:
/// which period to report on, whether a send is due, and placeholder
/// resolution for the email subject/body.
/// </summary>
public static class AutomationLogic
{
    private static readonly CultureInfo Danish = CultureInfo.GetCultureInfo("da-DK");

    /// <summary>
    /// The automated report period: Jan 1 of the previous month's year through
    /// the last day of the previous month. A run in January therefore covers
    /// the whole previous year.
    /// </summary>
    public static (DateTime Start, DateTime End) ComputeReportPeriod(DateTime today)
    {
        var end = new DateTime(today.Year, today.Month, 1).AddDays(-1);
        var start = new DateTime(end.Year, 1, 1);
        return (start, end);
    }

    /// <summary>The period key ("yyyy-MM") a run on <paramref name="today"/> should have sent.</summary>
    public static string GetTargetPeriod(DateTime today)
    {
        var (_, end) = ComputeReportPeriod(today);
        return end.ToString("yyyy-MM", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// True when the report for the target period has not been sent yet.
    /// "yyyy-MM" keys compare correctly with ordinal string comparison.
    /// </summary>
    public static bool ShouldSend(string? lastSentPeriod, DateTime today)
    {
        if (string.IsNullOrWhiteSpace(lastSentPeriod)) return true;
        return string.CompareOrdinal(lastSentPeriod, GetTargetPeriod(today)) < 0;
    }

    /// <summary>Replaces {month} (Danish month name) and {year} using the period end date.</summary>
    public static string ResolvePlaceholders(string template, DateTime periodEnd)
    {
        return template
            .Replace("{month}", periodEnd.ToString("MMMM", Danish))
            .Replace("{year}", periodEnd.Year.ToString(CultureInfo.InvariantCulture));
    }
}
