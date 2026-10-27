using System.Text;
using ModelContextProtocol.Authentication;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

public class AtlassianTokenStoreTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"tj-auth-{Guid.NewGuid():N}");
    private string FilePath => Path.Combine(_dir, "atlassian-auth.json");

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    private static TokenContainer SampleTokens() => new()
    {
        TokenType = "Bearer",
        AccessToken = "access-token-SECRET-123",
        RefreshToken = "refresh-token-456",
        ExpiresIn = 3600,
        ObtainedAt = new DateTimeOffset(2026, 10, 27, 8, 0, 0, TimeSpan.Zero),
        ClientId = "client-1",
        ClientSecret = "client-secret-1",
        TokenEndpointAuthMethod = "client_secret_post",
        AuthorizationServer = "https://auth.example.test",
    };

    [Fact]
    public async Task TokenStore_round_trips_tokens_and_registration()
    {
        var store = new AtlassianTokenStore(FilePath);
        Assert.False(store.HasTokens);

        store.ClientRegistration = ("client-1", "client-secret-1");
        await store.StoreTokensAsync(SampleTokens(), CancellationToken.None);

        var reopened = new AtlassianTokenStore(FilePath);
        var tokens = await reopened.GetTokensAsync(CancellationToken.None);

        Assert.True(reopened.HasTokens);
        Assert.NotNull(tokens);
        Assert.Equal("access-token-SECRET-123", tokens.AccessToken);
        Assert.Equal("refresh-token-456", tokens.RefreshToken);
        Assert.Equal(3600, tokens.ExpiresIn);
        Assert.Equal(SampleTokens().ObtainedAt, tokens.ObtainedAt);
        Assert.Equal("client-1", tokens.ClientId);
        Assert.Equal("client-secret-1", tokens.ClientSecret);
        Assert.Equal("client_secret_post", tokens.TokenEndpointAuthMethod);
        Assert.Equal("https://auth.example.test", tokens.AuthorizationServer);
        Assert.Equal(("client-1", (string?)"client-secret-1"), reopened.ClientRegistration);
    }

    [WindowsFact]
    public async Task TokenStore_file_is_not_plaintext_on_windows()
    {
        var store = new AtlassianTokenStore(FilePath);
        await store.StoreTokensAsync(SampleTokens(), CancellationToken.None);

        var content = Encoding.UTF8.GetString(File.ReadAllBytes(FilePath));
        Assert.DoesNotContain("access-token-SECRET-123", content);
    }

    [Fact]
    public async Task TokenStore_corrupt_file_reads_as_null()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllBytes(FilePath, [1, 2, 3, 4, 5]);

        var store = new AtlassianTokenStore(FilePath);

        Assert.False(store.HasTokens);
        Assert.Null(await store.GetTokensAsync(CancellationToken.None));
    }

    [WindowsFact]
    public async Task TokenStore_read_retries_while_the_file_is_briefly_locked()
    {
        var store = new AtlassianTokenStore(FilePath);
        await store.StoreTokensAsync(SampleTokens(), CancellationToken.None);

        // Another process mid-replace: the file can't be opened for a moment.
        var exclusive = new FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var release = Task.Run(async () =>
        {
            await Task.Delay(60);
            exclusive.Dispose();
        });

        var hasTokens = new AtlassianTokenStore(FilePath).HasTokens;
        await release;

        Assert.True(hasTokens);
    }

    [Fact]
    public void Stores_for_the_same_file_share_one_write_lock()
    {
        var a = new AtlassianTokenStore(FilePath);
        var b = new AtlassianTokenStore(Path.Combine(_dir, ".", "atlassian-auth.json"));
        var other = new AtlassianTokenStore(Path.Combine(_dir, "other.json"));

        Assert.Same(a.WriteLock, b.WriteLock);
        Assert.NotSame(a.WriteLock, other.WriteLock);
    }

    [Fact]
    public async Task Clear_deletes_file()
    {
        var store = new AtlassianTokenStore(FilePath);
        await store.StoreTokensAsync(SampleTokens(), CancellationToken.None);
        Assert.True(File.Exists(FilePath));

        store.Clear();

        Assert.False(File.Exists(FilePath));
        Assert.False(store.HasTokens);
    }
}

/// <summary>A fact that is skipped unless the tests run on Windows.</summary>
public sealed class WindowsFactAttribute : FactAttribute
{
    public WindowsFactAttribute()
    {
        if (!OperatingSystem.IsWindows()) Skip = "Only applies on Windows.";
    }
}
