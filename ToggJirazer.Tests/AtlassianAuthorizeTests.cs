using System.Net;
using System.Net.Http;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

/// <summary>
/// The browser step of the interactive sign-in (<see cref="AtlassianMcp.AuthorizeAsync"/>). These tests bind
/// the local callback port (localhost only, no network); they live in one class so they never run in parallel.
/// </summary>
public class AtlassianAuthorizeTests
{
    private static readonly Uri AuthorizationUri = new("https://auth.example.test/authorize?client_id=x");
    private static readonly TimeSpan Timeout = TimeSpan.FromMinutes(5);

    /// <summary>True when the callback port can be bound right now, i.e. no listener holds it.</summary>
    private static bool CallbackPortIsFree()
    {
        var listener = new HttpListener();
        listener.Prefixes.Add($"{OAuthCallbackListener.RedirectUri.Scheme}://{OAuthCallbackListener.RedirectUri.Authority}/");
        try
        {
            listener.Start();
            return true;
        }
        catch (HttpListenerException)
        {
            return false;
        }
        finally
        {
            listener.Close();
        }
    }

    [Fact]
    public async Task Browser_open_failure_without_a_way_to_cancel_stops_the_listener()
    {
        var failure = new InvalidOperationException("no browser");

        var thrown = await Assert.ThrowsAsync<InvalidOperationException>(() => AtlassianMcp.AuthorizeAsync(
            AuthorizationUri, _ => throw failure, showUri: null, Timeout, CancellationToken.None));

        Assert.Same(failure, thrown);
        Assert.True(CallbackPortIsFree());
    }

    [Fact]
    public async Task Browser_open_failure_keeps_listening_while_cancellable_so_the_shown_url_can_be_used()
    {
        using var cts = new CancellationTokenSource();
        Uri? shown = null;

        var authorize = AtlassianMcp.AuthorizeAsync(
            AuthorizationUri, _ => throw new InvalidOperationException("no browser"), uri => shown = uri, Timeout, cts.Token);

        Assert.Equal(AuthorizationUri, shown);
        Assert.False(authorize.IsCompleted);

        // The user pastes the URL into a browser; Atlassian redirects back to the listener.
        using var http = new HttpClient();
        using var response = await http.GetAsync(new Uri(OAuthCallbackListener.RedirectUri, "?code=abc&state=s1"));
        var (code, state, _) = await authorize;

        Assert.Equal("abc", code);
        Assert.Equal("s1", state);
        Assert.True(CallbackPortIsFree());
    }

    [Fact]
    public async Task Cancel_stops_the_wait_and_frees_the_port_immediately()
    {
        using var cts = new CancellationTokenSource();
        var authorize = AtlassianMcp.AuthorizeAsync(AuthorizationUri, _ => { }, showUri: null, Timeout, cts.Token);
        Assert.False(CallbackPortIsFree());

        cts.Cancel();
        var ex = await Assert.ThrowsAsync<AtlassianAuthException>(() => authorize);

        Assert.True(ex.AbortsSignIn);
        Assert.True(CallbackPortIsFree());
    }

    [Fact]
    public async Task Cancel_after_a_browser_open_failure_frees_the_port()
    {
        using var cts = new CancellationTokenSource();
        var authorize = AtlassianMcp.AuthorizeAsync(
            AuthorizationUri, _ => throw new InvalidOperationException("no browser"), showUri: null, Timeout, cts.Token);

        cts.Cancel();
        await Assert.ThrowsAsync<AtlassianAuthException>(() => authorize);

        Assert.True(CallbackPortIsFree());
    }
}
