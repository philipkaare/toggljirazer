using System.Net.Http;
using ModelContextProtocol.Authentication;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

/// <summary>
/// The decision logic of <see cref="AtlassianMcp.SignInAsync"/>: start from cleared tokens and, when a
/// persisted client registration fails, register a new client and try once more.
/// </summary>
public class AtlassianSignInRecoveryTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"tj-recover-{Guid.NewGuid():N}");
    private string FilePath => Path.Combine(_dir, "atlassian-auth.json");

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    private async Task<AtlassianTokenStore> StoreWithTokensAndRegistrationAsync()
    {
        var store = new AtlassianTokenStore(FilePath);
        await store.StoreTokensAsync(new TokenContainer { TokenType = "Bearer", AccessToken = "dead", ObtainedAt = DateTimeOffset.UtcNow }, CancellationToken.None);
        store.ClientRegistration = ("old-client", "old-secret");
        return store;
    }

    [Fact]
    public async Task Sign_in_starts_without_stored_tokens_but_keeps_the_registration()
    {
        var store = await StoreWithTokensAndRegistrationAsync();
        var attempts = new List<(bool HasTokens, string? ClientId)>();

        await AtlassianMcp.SignInWithRecoveryAsync(store, () =>
        {
            attempts.Add((store.HasTokens, store.ClientRegistration?.ClientId));
            return Task.CompletedTask;
        });

        Assert.Equal([(false, "old-client")], attempts);
    }

    public static TheoryData<Exception> RegistrationFailures() => new()
    {
        new AtlassianAuthException("Sign-in failed: invalid_client"),
        new AtlassianToolException("Protocol error", statusCode: null),
    };

    [Theory]
    [MemberData(nameof(RegistrationFailures))]
    public async Task Failure_with_persisted_registration_retries_once_with_a_new_registration(Exception failure)
    {
        var store = await StoreWithTokensAndRegistrationAsync();
        var attempts = new List<string?>();

        await AtlassianMcp.SignInWithRecoveryAsync(store, async () =>
        {
            attempts.Add(store.ClientRegistration?.ClientId);
            if (attempts.Count == 1)
            {
                // The failed attempt may have stored tokens for the dead client; the retry must not use them.
                await store.StoreTokensAsync(new TokenContainer { TokenType = "Bearer", AccessToken = "x", ObtainedAt = DateTimeOffset.UtcNow }, CancellationToken.None);
                throw failure;
            }
            Assert.False(store.HasTokens);
        });

        Assert.Equal(["old-client", null], attempts);
    }

    [Fact]
    public async Task Retry_failure_propagates_after_two_attempts()
    {
        var store = await StoreWithTokensAndRegistrationAsync();
        var attempts = 0;

        await Assert.ThrowsAsync<AtlassianAuthException>(() => AtlassianMcp.SignInWithRecoveryAsync(store, () =>
        {
            attempts++;
            throw new AtlassianAuthException("Sign-in failed: invalid_client");
        }));

        Assert.Equal(2, attempts);
    }

    [Fact]
    public async Task Failure_without_persisted_registration_is_not_retried()
    {
        var store = new AtlassianTokenStore(FilePath);
        var attempts = 0;

        await Assert.ThrowsAsync<AtlassianAuthException>(() => AtlassianMcp.SignInWithRecoveryAsync(store, () =>
        {
            attempts++;
            throw new AtlassianAuthException("Sign-in failed: invalid_client");
        }));

        Assert.Equal(1, attempts);
    }

    public static TheoryData<Exception> NonRegistrationFailures() => new()
    {
        new AtlassianAuthException("Sign-in was cancelled or timed out.", abortsSignIn: true),
        new HttpRequestException("offline"),
        new OperationCanceledException(),
    };

    [Theory]
    [MemberData(nameof(NonRegistrationFailures))]
    public async Task User_abort_or_network_failure_is_not_retried_and_keeps_the_registration(Exception failure)
    {
        var store = await StoreWithTokensAndRegistrationAsync();
        var attempts = 0;

        var thrown = await Assert.ThrowsAnyAsync<Exception>(() => AtlassianMcp.SignInWithRecoveryAsync(store, () =>
        {
            attempts++;
            throw failure;
        }));

        Assert.Same(failure, thrown);
        Assert.Equal(1, attempts);
        Assert.Equal("old-client", store.ClientRegistration?.ClientId);
    }

    public static TheoryData<Exception> ServerFailures() => new()
    {
        new AtlassianToolException("Internal server error", statusCode: 500),
        new AtlassianToolException("Bad gateway", statusCode: 502),
        new HttpRequestException("Service unavailable", null, System.Net.HttpStatusCode.ServiceUnavailable),
    };

    [Theory]
    [MemberData(nameof(ServerFailures))]
    public async Task Server_or_transport_5xx_is_not_retried_and_reports_trouble(Exception failure)
    {
        var store = await StoreWithTokensAndRegistrationAsync();
        var attempts = 0;

        var thrown = await Assert.ThrowsAnyAsync<Exception>(() => AtlassianMcp.SignInWithRecoveryAsync(store, () =>
        {
            attempts++;
            throw failure;
        }));

        Assert.Equal(1, attempts); // no second browser open
        Assert.Equal("Atlassian is having trouble right now — try again in a moment.", thrown.Message);
        Assert.Equal("old-client", store.ClientRegistration?.ClientId);
    }

    [Fact]
    public void Callback_access_denied_aborts_sign_in()
    {
        var ex = Assert.Throws<AtlassianAuthException>(() => OAuthCallbackListener.ParseQuery("?error=access_denied"));
        Assert.True(ex.AbortsSignIn);
    }

    [Fact]
    public void Callback_invalid_client_does_not_abort_sign_in()
    {
        var ex = Assert.Throws<AtlassianAuthException>(() => OAuthCallbackListener.ParseQuery("?error=invalid_client"));
        Assert.False(ex.AbortsSignIn);
    }
}
