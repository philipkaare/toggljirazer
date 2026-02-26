using CsvHelper.Configuration.Attributes;

namespace ToggJirazer.Models;

public class VersionReportRow
{
    [Name("Version")]
    public string Version { get; set; } = string.Empty;

    [Name("Total Estimate Sum")]
    public double TotalEstimateSum { get; set; }

    [Name("Worked Hours in Period")]
    public double WorkedHoursInPeriod { get; set; }

    [Name("Total Worked Hours")]
    public double TotalWorkedHours { get; set; }

    [Name("Difference")]
    public double Difference { get; set; }
}
