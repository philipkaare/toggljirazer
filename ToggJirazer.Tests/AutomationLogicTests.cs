using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class AutomationLogicTests
{
    // --- ComputeReportPeriod -------------------------------------------------

    [Fact]
    public void ComputeReportPeriod_MidYear_CoversJan1ToEndOfPreviousMonth()
    {
        var (start, end) = AutomationLogic.ComputeReportPeriod(new DateTime(2026, 8, 4));
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 7, 31), end);
    }

    [Fact]
    public void ComputeReportPeriod_OnTheFirst_StillCoversPreviousMonth()
    {
        var (start, end) = AutomationLogic.ComputeReportPeriod(new DateTime(2026, 9, 1));
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 8, 31), end);
    }

    [Fact]
    public void ComputeReportPeriod_InJanuary_CoversWholePreviousYear()
    {
        var (start, end) = AutomationLogic.ComputeReportPeriod(new DateTime(2027, 1, 2));
        Assert.Equal(new DateTime(2026, 1, 1), start);
        Assert.Equal(new DateTime(2026, 12, 31), end);
    }

    // --- GetTargetPeriod -----------------------------------------------------

    [Fact]
    public void GetTargetPeriod_MidYear_IsPreviousMonth()
    {
        Assert.Equal("2026-07", AutomationLogic.GetTargetPeriod(new DateTime(2026, 8, 4)));
    }

    [Fact]
    public void GetTargetPeriod_InJanuary_IsDecemberOfPreviousYear()
    {
        Assert.Equal("2026-12", AutomationLogic.GetTargetPeriod(new DateTime(2027, 1, 15)));
    }

    // --- ShouldSend ----------------------------------------------------------

    [Fact]
    public void ShouldSend_NeverSent_ReturnsTrue()
    {
        Assert.True(AutomationLogic.ShouldSend(null, new DateTime(2026, 8, 4)));
        Assert.True(AutomationLogic.ShouldSend(string.Empty, new DateTime(2026, 8, 4)));
    }

    [Fact]
    public void ShouldSend_AlreadySentForTargetPeriod_ReturnsFalse()
    {
        Assert.False(AutomationLogic.ShouldSend("2026-07", new DateTime(2026, 8, 4)));
        Assert.False(AutomationLogic.ShouldSend("2026-07", new DateTime(2026, 8, 31)));
    }

    [Fact]
    public void ShouldSend_SentForOlderPeriod_ReturnsTrue()
    {
        Assert.True(AutomationLogic.ShouldSend("2026-06", new DateTime(2026, 8, 4)));
        Assert.True(AutomationLogic.ShouldSend("2025-12", new DateTime(2026, 8, 4)));
    }

    [Fact]
    public void ShouldSend_NewMonthStarts_ReturnsTrueAgain()
    {
        // Sent for July; on Sep 1 the target is August, so it must send again.
        Assert.True(AutomationLogic.ShouldSend("2026-07", new DateTime(2026, 9, 1)));
    }

    // --- ResolvePlaceholders -------------------------------------------------

    [Fact]
    public void ResolvePlaceholders_ReplacesMonthAndYearInDanish()
    {
        var result = AutomationLogic.ResolvePlaceholders(
            "Timerapport {month} {year}", new DateTime(2026, 7, 31));
        Assert.Equal("Timerapport juli 2026", result);
    }

    [Fact]
    public void ResolvePlaceholders_NoPlaceholders_ReturnsTemplateUnchanged()
    {
        Assert.Equal("Hej", AutomationLogic.ResolvePlaceholders("Hej", new DateTime(2026, 7, 31)));
    }
}
