using System.Collections.Concurrent;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ModelContextProtocol.Authentication;

namespace ToggJirazer.Services.Atlassian;

/// <summary>
/// Persists the Atlassian OAuth tokens and the dynamic client registration in one file
/// (e.g. <c>atlassian-auth.json</c> next to the per-user appsettings.json). On Windows the file is
/// encrypted with DPAPI (CurrentUser); elsewhere it is written with file mode 600.
/// Every read goes to disk and every write is atomic (temp file plus rename), so a GUI and a
/// manual headless run don't overwrite each other's refreshed tokens.
/// </summary>
public sealed class AtlassianTokenStore : ITokenCache
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ToggJirazer.AtlassianTokenStore");
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = false };

    private const int ReadRetries = 3;
    private static readonly TimeSpan ReadRetryDelay = TimeSpan.FromMilliseconds(50);

    // One lock per file, shared by every store instance in this process that uses it.
    private static readonly ConcurrentDictionary<string, object> WriteLocks =
        new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);

    private readonly string _filePath;
    private readonly object _writeLock;

    public AtlassianTokenStore(string filePath)
    {
        _filePath = filePath;
        _writeLock = WriteLocks.GetOrAdd(Path.GetFullPath(filePath), _ => new object());
    }

    /// <summary>The lock guarding writes to this store's file (shared per full path). Exposed for tests.</summary>
    internal object WriteLock => _writeLock;

    /// <summary>True when the file exists and decrypts to a token container.</summary>
    public bool HasTokens => Read()?.Tokens != null;

    /// <summary>Signs out by deleting the file (tokens and client registration).</summary>
    public void Clear()
    {
        lock (_writeLock)
        {
            if (File.Exists(_filePath)) File.Delete(_filePath);
        }
    }

    /// <summary>Deletes the stored tokens but keeps the client registration.</summary>
    public void ClearTokens()
    {
        if (HasTokens) Update(file => file.Tokens = null);
    }

    /// <summary>The dynamically registered OAuth client, persisted alongside the tokens.</summary>
    public (string ClientId, string? ClientSecret)? ClientRegistration
    {
        get => Read() is { ClientId: { Length: > 0 } id } file ? (id, file.ClientSecret) : null;
        set => Update(file =>
        {
            file.ClientId = value?.ClientId;
            file.ClientSecret = value?.ClientSecret;
        });
    }

    public ValueTask<TokenContainer?> GetTokensAsync(CancellationToken cancellationToken) =>
        ValueTask.FromResult(Read()?.Tokens);

    public ValueTask StoreTokensAsync(TokenContainer tokens, CancellationToken cancellationToken)
    {
        Update(file => file.Tokens = tokens);
        return ValueTask.CompletedTask;
    }

    private void Update(Action<AuthFile> change)
    {
        lock (_writeLock)
        {
            var file = Read() ?? new AuthFile();
            change(file);
            Write(file);
        }
    }

    /// <summary>
    /// Returns the file's contents, or <c>null</c> when it is missing or cannot be decrypted/parsed.
    /// A file that can't be opened for a moment (another process replacing it) is retried a few times first.
    /// </summary>
    private AuthFile? Read()
    {
        for (var retry = 0; ; retry++)
        {
            byte[]? bytes;
            try
            {
                bytes = ReadBytes();
            }
            catch (IOException) when (retry < ReadRetries)
            {
                Thread.Sleep(ReadRetryDelay);
                continue;
            }
            catch
            {
                return null; // still unreadable: counts as "not signed in"
            }

            return bytes == null ? null : Decode(bytes);
        }
    }

    private byte[]? ReadBytes()
    {
        if (!File.Exists(_filePath)) return null;

        // Share delete/write so a concurrent atomic replace by another process is never blocked.
        using var stream = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static AuthFile? Decode(byte[] bytes)
    {
        try
        {
            if (OperatingSystem.IsWindows())
                bytes = ProtectedData.Unprotect(bytes, Entropy, DataProtectionScope.CurrentUser);

            return JsonSerializer.Deserialize<AuthFile>(bytes, JsonOptions);
        }
        catch
        {
            // A corrupt or foreign file counts as "not signed in".
            return null;
        }
    }

    private void Write(AuthFile file)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(file, JsonOptions);
        if (OperatingSystem.IsWindows())
            bytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);

        var directory = Path.GetDirectoryName(Path.GetFullPath(_filePath))!;
        Directory.CreateDirectory(directory);
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(_filePath)}.{Guid.NewGuid():N}.tmp");

        var options = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write };
        if (!OperatingSystem.IsWindows())
            options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;

        try
        {
            using (var stream = new FileStream(tempPath, options))
                stream.Write(bytes);
            File.Move(tempPath, _filePath, overwrite: true);
        }
        finally
        {
            if (File.Exists(tempPath)) File.Delete(tempPath);
        }
    }

    private sealed class AuthFile
    {
        public TokenContainer? Tokens { get; set; }
        public string? ClientId { get; set; }
        public string? ClientSecret { get; set; }
    }
}
