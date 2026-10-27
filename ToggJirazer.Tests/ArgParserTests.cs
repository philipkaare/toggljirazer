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
    public void No_args_returns_gui()
        => Assert.Equal(LaunchMode.Gui, ArgParser.Resolve([], false, false));

    [Fact]
    public void Tui_flag_falls_through_to_gui()
        => Assert.Equal(LaunchMode.Gui, ArgParser.Resolve(["--tui"], false, false));

    [Fact]
    public void Auto_flag_returns_auto_removed()
    {
        Assert.Equal(LaunchMode.AutoRemoved, ArgParser.Resolve(["--auto"], false, false));
        Assert.Equal(LaunchMode.AutoRemoved, ArgParser.Resolve(["--AUTO"], true, true)); // even when redirected (Task Scheduler)
    }

    [Fact]
    public void Help_beats_auto()
        => Assert.Equal(LaunchMode.Help, ArgParser.Resolve(["--auto", "--help"], false, false));
}
