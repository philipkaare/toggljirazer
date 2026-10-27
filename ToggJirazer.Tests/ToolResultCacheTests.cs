using System.Text.Json;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

public class ToolResultCacheTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"tj-toolcache-{Guid.NewGuid():N}");

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    /// <summary>Returns a fresh payload per call and counts the calls.</summary>
    private sealed class CountingTools : IJiraTools
    {
        public int Calls { get; private set; }

        public Task<JsonElement> CallToolAsync(string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct = default)
        {
            Calls++;
            using var doc = JsonDocument.Parse($$"""{"tool":"{{toolName}}","call":{{Calls}}}""");
            return Task.FromResult(doc.RootElement.Clone());
        }
    }

    private static Dictionary<string, object?> Args(params (string Key, object? Value)[] pairs) =>
        pairs.ToDictionary(p => p.Key, p => p.Value);

    [Fact]
    public async Task Cache_returns_stored_result_without_calling_inner()
    {
        var inner = new CountingTools();
        var tools = new ToolResultCache(_dir, TimeSpan.FromDays(1)).Wrap(inner);
        var args = Args(("cloudId", "cloud-1"), ("jql", "key in (ABC-1)"), ("fields", new[] { "summary" }));

        var first = await tools.CallToolAsync("searchJiraIssuesUsingJql", args);
        var second = await tools.CallToolAsync("searchJiraIssuesUsingJql", args);

        Assert.Equal(1, inner.Calls);
        Assert.Equal(1, second.GetProperty("call").GetInt32());
        Assert.Equal(first.GetRawText(), second.GetRawText());
    }

    [Fact]
    public async Task Cache_key_ignores_argument_order()
    {
        var inner = new CountingTools();
        var tools = new ToolResultCache(_dir, TimeSpan.FromDays(1)).Wrap(inner);

        await tools.CallToolAsync("getJiraIssue", Args(("cloudId", "cloud-1"), ("issueIdOrKey", "ABC-1"), ("fields", new[] { "*all" })));
        await tools.CallToolAsync("getJiraIssue", Args(("fields", new[] { "*all" }), ("issueIdOrKey", "ABC-1"), ("cloudId", "cloud-1")));
        await tools.CallToolAsync("getJiraIssue", Args(("cloudId", "cloud-1"), ("issueIdOrKey", "ABC-2"), ("fields", new[] { "*all" })));

        Assert.Equal(2, inner.Calls);
    }

    [Fact]
    public async Task Cache_expires()
    {
        var inner = new CountingTools();
        var args = Args(("cloudId", "cloud-1"));
        await new ToolResultCache(_dir, TimeSpan.FromDays(1)).Wrap(inner).CallToolAsync("getAccessibleAtlassianResources", args);

        var expired = new ToolResultCache(_dir, TimeSpan.Zero).Wrap(inner);
        var result = await expired.CallToolAsync("getAccessibleAtlassianResources", args);

        Assert.Equal(2, inner.Calls);
        Assert.Equal(2, result.GetProperty("call").GetInt32());
    }
}
