using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class FieldChoicesTests
{
    private static readonly string[] Site = ["Account", "Budget", "Summary"];

    [Fact]
    public void Current_field_name_kept_when_not_listed()
    {
        var (items, index) = FieldChoices.Build(Site, "Konto", "Account");

        Assert.Equal("Konto (current)", items[index]);
        Assert.Equal("(none)", items[0]);
        Assert.Equal(5, items.Count);
    }

    [Fact]
    public void Current_name_listed_is_selected()
    {
        var (items, index) = FieldChoices.Build(Site, "Budget", "Account");

        Assert.Equal("Budget", items[index]);
        Assert.Equal(4, items.Count);
    }

    [Fact]
    public void Null_current_selects_listed_default()
    {
        var (items, index) = FieldChoices.Build(Site, null, "Account");

        Assert.Equal("Account", items[index]);
    }

    [Fact]
    public void Null_current_with_missing_default_selects_none()
    {
        var (items, index) = FieldChoices.Build(["Summary"], null, "Account");

        Assert.Equal(0, index);
        Assert.Equal("(none)", items[index]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public void Empty_current_means_none_was_chosen(string current)
    {
        var (items, index) = FieldChoices.Build(Site, current, "Account");

        Assert.Equal(0, index);
        Assert.Equal("(none)", items[index]);
    }

    [Theory]
    [InlineData("(none)", "")]
    [InlineData("Konto (current)", "Konto")]
    [InlineData("Budget", "Budget")]
    public void ToSetting_maps_items(string item, string expected)
    {
        Assert.Equal(expected, FieldChoices.ToSetting(item));
    }

    [Fact]
    public void ToSetting_round_trips_selection()
    {
        var (items, index) = FieldChoices.Build(Site, "Konto", "Account");
        Assert.Equal("Konto", FieldChoices.ToSetting(items[index]));
    }
}
