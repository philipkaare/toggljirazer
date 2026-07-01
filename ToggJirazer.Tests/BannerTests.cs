using Xunit;
using ToggJirazer.Tui;

namespace ToggJirazer.Tests;

public class BannerTests
{
    [Fact]
    public void Text_includes_product_name_and_version()
    {
        var text = Banner.Text("1.2.3");
        Assert.Contains("ToggJirazer", text);
        Assert.Contains("1.2.3", text);
    }
}
