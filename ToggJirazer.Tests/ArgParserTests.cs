using Xunit;
using ToggJirazer.Cli;

namespace ToggJirazer.Tests;

public class ArgParserTests
{
    [Fact]
    public void Help_flag_returns_help()
    {
        Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["--help"], false, false));
        Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["-h"], false, false));
    }

    [Fact]
    public void Headless_flag_returns_headless()
    {
        Assert.Equal(LaunchMode.Headless, ArgParser.Resolve(["--headless"], false, false));
        Assert.Equal(LaunchMode.Headless, ArgParser.Resolve(["--run"], false, false));
    }

    [Fact]
    public void Redirected_output_forces_headless()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve([], false, true));

    [Fact]
    public void Redirected_input_forces_headless()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve([], true, false));

    [Fact]
    public void Interactive_no_flags_returns_headless()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve([], false, false));
    [Fact]
    public void Tui_flag_selects_terminal_interface()
        => Assert.Equal(LaunchMode.Tui, ArgParser.Resolve(["--tui"], false, false));

    [Fact]
    public void Help_takes_precedence_over_tui()
        => Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["--tui", "--help"], false, false));

    [Fact]
    public void Headless_takes_precedence_over_tui()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve(["--tui", "--headless"], false, false));

    [Fact]
    public void Tui_with_redirected_input_uses_headless()
        => Assert.Equal(LaunchMode.Headless, ArgParser.Resolve(["--tui"], true, false));
    [Fact]
    public void Windows_defaults_to_gui()
        => Assert.Equal(LaunchMode.Gui, ArgParser.Resolve([], false, false, true));

    [Fact]
    public void Windows_tui_flag_overrides_gui()
        => Assert.Equal(LaunchMode.Tui, ArgParser.Resolve(["--tui"], false, false, true));

    [Theory]
    [InlineData("--auto")]
    [InlineData("--AUTO")]
    public void Auto_remains_available_with_redirected_streams(string flag)
        => Assert.Equal(LaunchMode.Auto, ArgParser.Resolve([flag], true, true));

    [Fact]
    public void Help_prevents_auto_run()
        => Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["--auto", "--help"], false, false));
}
