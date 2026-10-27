using System.Globalization;

namespace ToggJirazer.Services;

/// <summary>Placeholder resolution for the report email subject/body.</summary>
public static class EmailTemplate
{
    private static readonly CultureInfo Danish = CultureInfo.GetCultureInfo("da-DK");

    /// <summary>Replaces {month} (Danish month name) and {year} using the period end date.</summary>
    public static string ResolvePlaceholders(string template, DateTime periodEnd)
    {
        return template
            .Replace("{month}", periodEnd.ToString("MMMM", Danish))
            .Replace("{year}", periodEnd.Year.ToString(CultureInfo.InvariantCulture));
    }
}
