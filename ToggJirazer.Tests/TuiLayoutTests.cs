using System.Drawing;
using Terminal.Gui.ViewBase;
using Terminal.Gui.Views;
using ToggJirazer.Services;
using ToggJirazer.Tui;
using Xunit;

#pragma warning disable CS0618

namespace ToggJirazer.Tests;

public class TuiLayoutTests
{
    [Fact]
    public void Read_only_log_text_has_contrasting_colors()
    {
        var type = typeof(AppTui).Assembly.GetType("ToggJirazer.Tui.MainView")!;
        using var window = (Window)Activator.CreateInstance(type,
            new SettingsService(Path.Combine(Path.GetTempPath(), $"tj-{Guid.NewGuid():N}.json")))!;
        var log = window.SubViews.OfType<FrameView>().Single(v => v.Title == "Log")
            .SubViews.OfType<TextView>().Single();

        log.Text = "Settings saved.\n";
        var colors = log.GetAttributeForRole(Terminal.Gui.Drawing.VisualRole.ReadOnly);
        Assert.True(log.ReadOnly);
        Assert.NotEqual(colors.Foreground, colors.Background);
    }

    [Theory]
    [InlineData(24)]
    [InlineData(30)]
    [InlineData(40)]
    public void Log_has_readable_space_when_idle_running_and_resized(int rows)
    {
        var type = typeof(AppTui).Assembly.GetType("ToggJirazer.Tui.MainView")!;
        using var window = (Window)Activator.CreateInstance(type,
            new SettingsService(Path.Combine(Path.GetTempPath(), $"tj-{Guid.NewGuid():N}.json")))!;
        window.BeginInit();
        window.EndInit();

        foreach (var height in new[] { rows, 40, 24, rows })
        foreach (var running in new[] { false, true })
        {
            window.Frame = new Rectangle(0, 0, 80, height);
            window.SubViews.OfType<ProgressBar>().Single().Visible = running;
            window.SetNeedsLayout();
            window.Layout();

            var frame = window.SubViews.OfType<FrameView>().Single(v => v.Title == "Log");
            var log = frame.SubViews.OfType<TextView>().Single();
            Assert.True(log.Viewport.Height >= 4,
                $"At 80x{height}, running={running}: log has {log.Viewport.Height} rows (frame {frame.Frame}).");
            Assert.True(frame.Frame.Bottom <= window.Viewport.Height);
            var tabs = window.SubViews.OfType<Tabs>().Single();
            Assert.True(tabs.Frame.Bottom <= frame.Frame.Top);
        }
    }
}
