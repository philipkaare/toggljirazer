using Xunit;
using ToggJirazer.Cli;

namespace ToggJirazer.Tests;

public class PlatformOpenerTests
{
    [Fact]
    public void MacOs_uses_open()
    {
        var cmd = PlatformOpener.CommandFor(OpenerOs.MacOs, "/tmp/r.xlsx");
        Assert.Equal("open", cmd.FileName);
        Assert.Contains("/tmp/r.xlsx", cmd.Arguments);
        Assert.False(cmd.UseShellExecute);
    }

    [Fact]
    public void Linux_uses_xdg_open()
    {
        var cmd = PlatformOpener.CommandFor(OpenerOs.Linux, "/tmp/r.xlsx");
        Assert.Equal("xdg-open", cmd.FileName);
        Assert.False(cmd.UseShellExecute);
    }

    [Fact]
    public void Windows_uses_shell_execute_on_the_path()
    {
        var cmd = PlatformOpener.CommandFor(OpenerOs.Windows, @"C:\r.xlsx");
        Assert.Equal(@"C:\r.xlsx", cmd.FileName);
        Assert.True(cmd.UseShellExecute);
    }
}
