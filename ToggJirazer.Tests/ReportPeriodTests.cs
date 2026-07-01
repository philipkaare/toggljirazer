using Xunit;
using ToggJirazer.Cli;
using ToggJirazer.Models;

namespace ToggJirazer.Tests;

public class ReportPeriodTests
{
    private static void Sink(string _) { }

    [Fact]
    public void Explicit_dates_are_parsed()
    {
        var report = new ReportConfig { StartDate = "2026-02-01", EndDate = "2026-02-28" };
        var (start, end) = ReportRunner.ResolvePeriod(report, new DateTime(2026, 7, 1), Sink);
        Assert.Equal(new DateTime(2026, 2, 1), start);
        Assert.Equal(new DateTime(2026, 2, 28), end);
    }

    [Fact]
    public void Empty_dates_default_to_year_start_and_today()
    {
        var report = new ReportConfig { StartDate = null, EndDate = null };
        var (start, end) = ReportRunner.ResolvePeriod(report, new DateTime(2026, 7, 1), Sink);
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 7, 1), end);
    }

    [Fact]
    public void Unparseable_start_date_throws()
    {
        var report = new ReportConfig { StartDate = "not-a-date", EndDate = "2026-02-28" };
        Assert.Throws<InvalidOperationException>(
            () => ReportRunner.ResolvePeriod(report, new DateTime(2026, 7, 1), Sink));
    }
}
