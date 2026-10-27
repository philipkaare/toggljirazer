using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class EmailTemplateTests
{
    [Fact]
    public void ResolvePlaceholders_uses_danish_month_and_year()
        => Assert.Equal("Timerapport juli 2026",
            EmailTemplate.ResolvePlaceholders("Timerapport {month} {year}", new DateTime(2026, 7, 31)));

    [Fact]
    public void ResolvePlaceholders_without_placeholders_returns_template_unchanged()
        => Assert.Equal("Hej", EmailTemplate.ResolvePlaceholders("Hej", new DateTime(2026, 7, 31)));
}
