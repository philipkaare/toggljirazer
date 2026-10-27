using System.Text.Json;

namespace ToggJirazer.Services.Atlassian;

/// <summary>The Atlassian sign-in is missing, expired or revoked.</summary>
public sealed class AtlassianAuthException(string message, Exception? inner = null, bool abortsSignIn = false) : Exception(message, inner)
{
    /// <summary>
    /// True when the interactive sign-in ended for a reason a new client registration can't fix: the user
    /// cancelled, declined or let it time out, or the callback port was busy.
    /// </summary>
    public bool AbortsSignIn { get; } = abortsSignIn;
}

/// <summary>The Atlassian MCP server's rate limit was reached. The run must stop rather than continue with partial data.</summary>
public sealed class AtlassianRateLimitException()
    : Exception("Atlassian rate limit reached — try again later (the 1-day cache keeps what was fetched).");

/// <summary>An Atlassian MCP tool reported an error, e.g. Jira rejecting a JQL query.</summary>
public sealed class AtlassianToolException(string message, int? statusCode) : Exception(message)
{
    public int? StatusCode { get; } = statusCode;
}

/// <summary>Unwraps Atlassian MCP tool text content and classifies tool errors.</summary>
public static class AtlassianToolResult
{
    /// <summary>
    /// Parses tool text content as JSON. Error payloads such as
    /// <c>{"error":true,"message":...,"statusCode":400,"rateLimit":{"remaining":N}}</c> map to:
    /// 401/403 → <see cref="AtlassianAuthException"/>; 429 or <c>rateLimit.remaining == 0</c> →
    /// <see cref="AtlassianRateLimitException"/>; anything else → <see cref="AtlassianToolException"/>.
    /// </summary>
    public static JsonElement Parse(string text, bool isError)
    {
        JsonElement root;
        try
        {
            using var doc = JsonDocument.Parse(text);
            root = doc.RootElement.Clone();
        }
        catch (JsonException)
        {
            if (isError) throw new AtlassianToolException(text, statusCode: null);
            throw new AtlassianToolException($"Atlassian returned a result that is not JSON: {Truncate(text)}", statusCode: null);
        }

        var errorFlag = root.ValueKind == JsonValueKind.Object
            && root.TryGetProperty("error", out var error)
            && error.ValueKind == JsonValueKind.True;
        if (!isError && !errorFlag) return root;

        throw Classify(root, text);
    }

    private static Exception Classify(JsonElement root, string text)
    {
        if (root.ValueKind != JsonValueKind.Object)
            return new AtlassianToolException(text, statusCode: null);

        var statusCode = ReadInt(root, "statusCode");
        var message = root.TryGetProperty("message", out var m) && m.ValueKind == JsonValueKind.String
            ? m.GetString()!
            : text;
        var remaining = root.TryGetProperty("rateLimit", out var rateLimit) && rateLimit.ValueKind == JsonValueKind.Object
            ? ReadInt(rateLimit, "remaining")
            : null;

        if (statusCode == 429 || remaining == 0) return new AtlassianRateLimitException();
        if (statusCode is 401 or 403) return new AtlassianAuthException(message);
        return new AtlassianToolException(message, statusCode);
    }

    private static int? ReadInt(JsonElement obj, string name)
    {
        if (!obj.TryGetProperty(name, out var value)) return null;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number)) return number;
        if (value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out var parsed)) return parsed;
        return null;
    }

    private static string Truncate(string text) => text.Length <= 200 ? text : text[..200] + "…";
}
