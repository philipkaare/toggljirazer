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
    public void Interactive_no_flags_returns_tui()
        => Assert.Equal(LaunchMode.Tui, ArgParser.Resolve([], false, false));
}
