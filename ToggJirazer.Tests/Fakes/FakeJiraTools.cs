using System.Text.Json;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Tests.Fakes;

/// <summary>
/// Records every tool call and answers through a scripted responder. The responder may throw
/// (e.g. <see cref="AtlassianToolException"/>) to script a tool error.
/// </summary>
public sealed class FakeJiraTools(Func<string, IReadOnlyDictionary<string, object?>, JsonElement> responder) : IJiraTools
{
    public List<(string Tool, IReadOnlyDictionary<string, object?> Args)> Calls { get; } = new();

    public IEnumerable<(string Tool, IReadOnlyDictionary<string, object?> Args)> CallsTo(string tool) =>
        Calls.Where(c => c.Tool == tool);

    public Task<JsonElement> CallToolAsync(string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct = default)
    {
        Calls.Add((toolName, arguments));
        return Task.FromResult(responder(toolName, arguments));
    }
}
