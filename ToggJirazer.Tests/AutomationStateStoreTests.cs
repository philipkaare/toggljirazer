using System.IO;
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class AutomationStateStoreTests : IDisposable
{
    private readonly string _dir;

    public AutomationStateStoreTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "ToggJirazerTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
    }

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    [Fact]
    public void Load_NoFile_ReturnsEmptyState()
    {
        var state = new AutomationStateStore(_dir).Load();
        Assert.Null(state.LastSentPeriod);
    }

    [Fact]
    public void SaveThenLoad_RoundTripsLastSentPeriod()
    {
        var store = new AutomationStateStore(_dir);
        store.Save(new AutomationState { LastSentPeriod = "2026-07" });

        var loaded = store.Load();
        Assert.Equal("2026-07", loaded.LastSentPeriod);
        Assert.True(File.Exists(Path.Combine(_dir, "automation-state.json")));
    }

    [Fact]
    public void Load_CorruptFile_ReturnsEmptyState()
    {
        File.WriteAllText(Path.Combine(_dir, "automation-state.json"), "not json {{{");
        var state = new AutomationStateStore(_dir).Load();
        Assert.Null(state.LastSentPeriod);
    }
}
