using System.IO;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ToggJirazer.Services;

/// <summary>
/// A simple on-disk cache for HTTP responses. Each unique request (method + URI + body)
/// is stored as a single file so repeated runs do not re-request the same "chunks" from
/// the Toggl/Jira APIs. Entries expire after a configurable duration (1 day by default).
/// </summary>
public class HttpResponseCache
{
    private readonly string _directory;
    private readonly TimeSpan _expiration;

    /// <summary>
    /// Invoked when an entry cannot be written because of a permission problem.
    /// Return <c>true</c> to retry the write, <c>false</c> to give up caching that entry.
    /// </summary>
    private readonly Func<Exception, bool>? _onWritePermissionError;

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = false };

    public HttpResponseCache(string directory, TimeSpan expiration, Func<Exception, bool>? onWritePermissionError = null)
    {
        _directory = directory;
        _expiration = expiration;
        _onWritePermissionError = onWritePermissionError;
    }

    /// <summary>
    /// Builds a <see cref="DelegatingHandler"/> that consults this cache before hitting the
    /// network. A fresh inner <see cref="HttpClientHandler"/> is created per call so each
    /// service can own (and dispose) its own handler chain.
    /// </summary>
    public DelegatingHandler CreateHandler() =>
        new CachingHttpHandler(this) { InnerHandler = new HttpClientHandler() };

    private string PathForKey(string key) => Path.Combine(_directory, key + ".json");

    /// <summary>Computes a stable cache key from the request method, URI and body.</summary>
    public static async Task<string> ComputeKeyAsync(HttpRequestMessage request)
    {
        var body = request.Content != null
            ? await request.Content.ReadAsByteArrayAsync()
            : Array.Empty<byte>();

        using var sha = SHA256.Create();
        sha.TransformBlock(Encoding.UTF8.GetBytes(request.Method.Method + "\n"), 0, request.Method.Method.Length + 1, null, 0);
        var uriBytes = Encoding.UTF8.GetBytes((request.RequestUri?.ToString() ?? string.Empty) + "\n");
        sha.TransformBlock(uriBytes, 0, uriBytes.Length, null, 0);
        sha.TransformFinalBlock(body, 0, body.Length);
        return Convert.ToHexString(sha.Hash!);
    }

    /// <summary>Returns a cached response for the key, or <c>null</c> on a miss or if expired.</summary>
    public HttpResponseMessage? TryRead(string key, HttpRequestMessage request)
    {
        var path = PathForKey(key);
        try
        {
            if (!File.Exists(path)) return null;

            var entry = JsonSerializer.Deserialize<CacheEntry>(File.ReadAllText(path), JsonOptions);
            if (entry == null) return null;

            if (DateTime.UtcNow - entry.TimestampUtc > _expiration)
            {
                // Expired — remove it so the directory does not grow unbounded.
                try { File.Delete(path); } catch { /* best effort */ }
                return null;
            }

            var bytes = Convert.FromBase64String(entry.ContentBase64);
            var response = new HttpResponseMessage((HttpStatusCode)entry.StatusCode)
            {
                Content = new ByteArrayContent(bytes),
                RequestMessage = request
            };
            if (!string.IsNullOrEmpty(entry.ContentType))
                response.Content.Headers.TryAddWithoutValidation("Content-Type", entry.ContentType);
            return response;
        }
        catch
        {
            // Treat any read problem as a cache miss.
            return null;
        }
    }

    /// <summary>
    /// Stores the response for the key. The response content is buffered and replaced so that
    /// the caller can still read it afterwards. If writing fails because of a permission problem
    /// the configured callback decides whether to retry or cancel.
    /// </summary>
    public async Task WriteAsync(string key, HttpResponseMessage response)
    {
        // Buffer the content and put a fresh copy back so the caller can still read the body.
        var bytes = await response.Content.ReadAsByteArrayAsync();
        var contentType = response.Content.Headers.ContentType?.ToString();
        var newContent = new ByteArrayContent(bytes);
        if (!string.IsNullOrEmpty(contentType))
            newContent.Headers.TryAddWithoutValidation("Content-Type", contentType);
        response.Content = newContent;

        var entry = new CacheEntry
        {
            TimestampUtc = DateTime.UtcNow,
            StatusCode = (int)response.StatusCode,
            ContentType = contentType,
            ContentBase64 = Convert.ToBase64String(bytes)
        };
        var json = JsonSerializer.Serialize(entry, JsonOptions);
        var path = PathForKey(key);

        while (true)
        {
            try
            {
                Directory.CreateDirectory(_directory);
                File.WriteAllText(path, json);
                return;
            }
            catch (Exception ex) when (IsPermissionProblem(ex))
            {
                if (_onWritePermissionError != null && _onWritePermissionError(ex))
                    continue; // user chose "Try again"
                Console.Error.WriteLine($"  [cache] Skipped saving (permission denied): {ex.Message}");
                return;
            }
            catch (Exception ex)
            {
                // Non-permission failures should never break report generation.
                Console.Error.WriteLine($"  [cache] Could not save entry: {ex.Message}");
                return;
            }
        }
    }

    private static bool IsPermissionProblem(Exception ex) =>
        ex is UnauthorizedAccessException
        || (ex is IOException && (ex.HResult & 0xFFFF) == 5 /* ERROR_ACCESS_DENIED */);

    private sealed class CacheEntry
    {
        public DateTime TimestampUtc { get; set; }
        public int StatusCode { get; set; }
        public string? ContentType { get; set; }
        public string ContentBase64 { get; set; } = string.Empty;
    }

    /// <summary>
    /// A delegating handler that returns cached responses on a hit and stores successful
    /// responses on a miss.
    /// </summary>
    private sealed class CachingHttpHandler : DelegatingHandler
    {
        private readonly HttpResponseCache _cache;

        public CachingHttpHandler(HttpResponseCache cache)
        {
            _cache = cache;
        }

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var key = await ComputeKeyAsync(request);

            var cached = _cache.TryRead(key, request);
            if (cached != null)
            {
                Console.WriteLine($"  [cache] hit: {request.Method} {request.RequestUri}");
                return cached;
            }

            var response = await base.SendAsync(request, cancellationToken);

            // Only cache successful responses so transient errors/auth failures are not persisted.
            if (response.IsSuccessStatusCode)
                await _cache.WriteAsync(key, response);

            return response;
        }
    }
}
