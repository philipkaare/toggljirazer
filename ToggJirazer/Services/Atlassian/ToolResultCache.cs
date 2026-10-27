using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ToggJirazer.Services.Atlassian;

/// <summary>
/// On-disk cache for Atlassian MCP tool results, so repeated runs don't re-request the same data.
/// Caching happens per tool call (the HTTP-level <see cref="HttpResponseCache"/> can't be used on the
/// MCP transport because JSON-RPC request IDs make every request unique). Only successful results
/// are stored; entries expire after the configured duration.
/// </summary>
public sealed class ToolResultCache
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = false };

    private readonly string _directory;
    private readonly TimeSpan _expiration;
    private readonly Func<Exception, bool>? _onWritePermissionError;

    public ToolResultCache(string directory, TimeSpan expiration, Func<Exception, bool>? onWritePermissionError = null)
    {
        _directory = directory;
        _expiration = expiration;
        _onWritePermissionError = onWritePermissionError;
    }

    /// <summary>Returns tools that answer from this cache and fall back to <paramref name="inner"/> on a miss.</summary>
    public IJiraTools Wrap(IJiraTools inner) => new CachingTools(this, inner);

    /// <summary>SHA256 of the tool name plus the arguments as JSON with object keys sorted, so argument order doesn't matter.</summary>
    private static string ComputeKey(string toolName, IReadOnlyDictionary<string, object?> arguments)
    {
        using var buffer = new MemoryStream();
        buffer.Write(Encoding.UTF8.GetBytes(toolName + "\n"));
        using (var writer = new Utf8JsonWriter(buffer))
            WriteCanonical(writer, JsonSerializer.SerializeToElement(arguments, JsonOptions));

        return Convert.ToHexString(SHA256.HashData(buffer.ToArray()));
    }

    private static void WriteCanonical(Utf8JsonWriter writer, JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                writer.WriteStartObject();
                foreach (var property in element.EnumerateObject().OrderBy(p => p.Name, StringComparer.Ordinal))
                {
                    writer.WritePropertyName(property.Name);
                    WriteCanonical(writer, property.Value);
                }
                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (var item in element.EnumerateArray())
                    WriteCanonical(writer, item);
                writer.WriteEndArray();
                break;
            default:
                element.WriteTo(writer);
                break;
        }
    }

    private string PathForKey(string key) => Path.Combine(_directory, key + ".json");

    /// <summary>Returns the cached payload, or <c>null</c> on a miss, an expired entry or a read problem.</summary>
    private JsonElement? TryRead(string key)
    {
        var path = PathForKey(key);
        try
        {
            if (!File.Exists(path)) return null;

            var entry = JsonSerializer.Deserialize<CacheEntry>(File.ReadAllText(path), JsonOptions);
            if (entry == null) return null;

            if (DateTime.UtcNow - entry.TimestampUtc >= _expiration)
            {
                // Expired — remove it so the directory does not grow unbounded.
                try { File.Delete(path); } catch { /* best effort */ }
                return null;
            }

            using var doc = JsonDocument.Parse(entry.Json);
            return doc.RootElement.Clone();
        }
        catch
        {
            return null;
        }
    }

    private void Write(string key, JsonElement payload)
    {
        var entry = new CacheEntry { TimestampUtc = DateTime.UtcNow, Json = payload.GetRawText() };
        HttpResponseCache.WriteEntryFile(_directory, PathForKey(key), JsonSerializer.Serialize(entry, JsonOptions), _onWritePermissionError);
    }

    private sealed class CacheEntry
    {
        public DateTime TimestampUtc { get; set; }
        public string Json { get; set; } = string.Empty;
    }

    private sealed class CachingTools(ToolResultCache cache, IJiraTools inner) : IJiraTools
    {
        public async Task<JsonElement> CallToolAsync(string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct = default)
        {
            var key = ComputeKey(toolName, arguments);

            if (cache.TryRead(key) is { } cached)
            {
                Console.WriteLine($"  [cache] hit: {toolName}");
                return cached;
            }

            // Errors propagate from the inner call, so failures are never cached.
            var payload = await inner.CallToolAsync(toolName, arguments, ct);
            cache.Write(key, payload);
            return payload;
        }
    }
}
