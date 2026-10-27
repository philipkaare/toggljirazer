using System.Text.Json;

namespace ToggJirazer.Services.Atlassian;

/// <summary>Calls Atlassian MCP tools. Implemented over the live MCP client and wrapped by <see cref="ToolResultCache"/>.</summary>
public interface IJiraTools
{
    /// <summary>
    /// Calls an Atlassian MCP tool and returns the parsed JSON payload (the text content parsed as JSON).
    /// Throws <see cref="AtlassianAuthException"/>, <see cref="AtlassianRateLimitException"/>,
    /// or <see cref="AtlassianToolException"/>.
    /// </summary>
    Task<JsonElement> CallToolAsync(string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct = default);
}
