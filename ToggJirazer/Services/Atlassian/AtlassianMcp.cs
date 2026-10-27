using System.Net;
using System.Net.Http;
using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Authentication;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using ToggJirazer.Cli;

namespace ToggJirazer.Services.Atlassian;

/// <summary>An Atlassian site the signed-in user can access. <see cref="CloudId"/> is passed to every Jira tool as <c>cloudId</c>.</summary>
public sealed record AtlassianSite(string CloudId, string Url, string Name);

/// <summary>
/// Connection to Atlassian's remote MCP server with OAuth. A browser opens only in
/// <see cref="SignInAsync"/>; every other path uses a non-interactive client that fails fast with
/// <see cref="AtlassianAuthException"/> when the stored sign-in can't be used or refreshed.
/// </summary>
public sealed class AtlassianMcp : IAsyncDisposable
{
    public static readonly Uri Endpoint = new("https://mcp.atlassian.com/v2/mcp");

    private const string SitesTool = "getAccessibleAtlassianResources";
    private static readonly TimeSpan SignInTimeout = TimeSpan.FromMinutes(5);
    private static readonly IReadOnlyDictionary<string, object?> NoArguments = new Dictionary<string, object?>();

    private readonly AtlassianTokenStore _store;
    private readonly Action<Uri> _openBrowser;
    private readonly SemaphoreSlim _connectGate = new(1, 1);
    private Connection? _connection; // the non-interactive client, created on first use

    public AtlassianMcp(AtlassianTokenStore store, Action<Uri>? openBrowser = null)
    {
        _store = store;
        _openBrowser = openBrowser ?? (uri => PlatformOpener.Open(uri.AbsoluteUri));
        Tools = new McpJiraTools(this);
    }

    /// <summary>Whether a sign-in is stored on disk. Does not check that it still works and makes no network call.</summary>
    public bool HasStoredSignIn => _store.HasTokens;

    /// <summary>Non-interactive tools: never open a browser.</summary>
    public IJiraTools Tools { get; }

    /// <summary>
    /// Explicit interactive sign-in — the only place a browser opens. Recovers from a stale stored sign-in
    /// (see <see cref="SignInWithRecoveryAsync"/>).
    /// </summary>
    public async Task SignInAsync(CancellationToken ct)
    {
        await SignInWithRecoveryAsync(_store, async () =>
        {
            await using var connection = await ConnectAsync(interactive: true, ct);
            // Connecting authorizes the session; a tool call proves the sign-in actually works.
            await CallAsync(connection.Client, SitesTool, NoArguments, ct);
        });

        // Drop the non-interactive client so it starts fresh with the new tokens.
        await ResetConnectionAsync();
    }

    /// <summary>
    /// Sign-in is only offered when the stored sign-in doesn't work, so its tokens are dropped first.
    /// When the attempt then fails with an Atlassian error while a persisted client registration was
    /// used (Atlassian may have dropped that client), the registration is dropped too and the attempt
    /// is repeated once with a newly registered client. A user abort or a network error is not retried.
    /// </summary>
    internal static async Task SignInWithRecoveryAsync(AtlassianTokenStore store, Func<Task> attempt)
    {
        store.ClearTokens();
        var usedPersistedRegistration = store.ClientRegistration != null;
        try
        {
            await attempt();
        }
        catch (Exception ex) when (usedPersistedRegistration && IsRegistrationSuspect(ex))
        {
            store.ClearTokens();
            store.ClientRegistration = null;
            await attempt();
        }
    }

    private static bool IsRegistrationSuspect(Exception ex) =>
        ex is AtlassianAuthException { AbortsSignIn: false } or AtlassianToolException;

    /// <summary>Signs out: deletes the stored tokens and client registration and drops the cached connection.</summary>
    public async Task SignOutAsync()
    {
        _store.Clear();
        await ResetConnectionAsync();
    }

    /// <summary>The user's Atlassian sites, de-duplicated by cloud ID. Throws <see cref="AtlassianAuthException"/> when not signed in.</summary>
    public async Task<IReadOnlyList<AtlassianSite>> GetSitesAsync(CancellationToken ct) =>
        ParseSites(await Tools.CallToolAsync(SitesTool, NoArguments, ct));

    /// <summary>Whether the stored sign-in works (refreshing it quietly when it can). Never opens a browser.</summary>
    public async Task<bool> IsSignedInAsync(CancellationToken ct)
    {
        if (!_store.HasTokens) return false;
        return await ProbeSignInAsync(() => GetSitesAsync(ct), ResetConnectionAsync);
    }

    /// <summary>
    /// Runs the sites call and reports whether it worked. A failed token refresh (e.g. <c>invalid_grant</c>)
    /// may surface as <see cref="AtlassianToolException"/> rather than <see cref="AtlassianAuthException"/>,
    /// so both count as "not signed in". Only an auth failure drops the cached connection: a tool error
    /// says nothing certain about the connection, which a running report may be using. Network errors propagate.
    /// </summary>
    internal static async Task<bool> ProbeSignInAsync(Func<Task> getSites, Func<Task> resetConnection)
    {
        try
        {
            await getSites();
            return true;
        }
        catch (AtlassianAuthException)
        {
            await resetConnection();
            return false;
        }
        catch (AtlassianToolException)
        {
            return false;
        }
    }

    /// <summary>Parses the <c>getAccessibleAtlassianResources</c> array; the same site appears once per product.</summary>
    internal static IReadOnlyList<AtlassianSite> ParseSites(JsonElement resources)
    {
        if (resources.ValueKind != JsonValueKind.Array)
            throw new AtlassianToolException("Unexpected response from getAccessibleAtlassianResources.", statusCode: null);

        return resources.EnumerateArray()
            .Select(r => new AtlassianSite(
                r.GetProperty("id").GetString() ?? string.Empty,
                r.TryGetProperty("url", out var url) ? url.GetString() ?? string.Empty : string.Empty,
                r.TryGetProperty("name", out var name) ? name.GetString() ?? string.Empty : string.Empty))
            .Where(s => s.CloudId.Length > 0)
            .DistinctBy(s => s.CloudId)
            .ToList();
    }

    public async ValueTask DisposeAsync()
    {
        await ResetConnectionAsync();
        _connectGate.Dispose();
    }

    private async Task<JsonElement> CallNonInteractiveAsync(string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct)
    {
        // Without stored tokens the SDK would register a new OAuth client before failing; skip that.
        if (!_store.HasTokens) throw new AtlassianAuthException("Not signed in");

        try
        {
            var connection = await GetConnectionAsync(ct);
            return await CallAsync(connection.Client, toolName, arguments, ct);
        }
        catch (AtlassianAuthException)
        {
            await ResetConnectionAsync();
            throw;
        }
    }

    private static async Task<JsonElement> CallAsync(McpClient client, string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct)
    {
        CallToolResult result;
        try
        {
            result = await client.CallToolAsync(toolName, arguments, cancellationToken: ct);
        }
        catch (Exception ex) when (ClassifyFailure(ex) is { } classified)
        {
            throw classified;
        }

        var text = string.Concat(result.Content.OfType<TextContentBlock>().Select(b => b.Text));
        return AtlassianToolResult.Parse(text, result.IsError == true);
    }

    private async Task<Connection> GetConnectionAsync(CancellationToken ct)
    {
        await _connectGate.WaitAsync(ct);
        try
        {
            return _connection ??= await ConnectAsync(interactive: false, ct);
        }
        finally
        {
            _connectGate.Release();
        }
    }

    private async Task ResetConnectionAsync()
    {
        await _connectGate.WaitAsync();
        try
        {
            if (_connection != null) await _connection.DisposeAsync();
            _connection = null;
        }
        finally
        {
            _connectGate.Release();
        }
    }

    private async Task<Connection> ConnectAsync(bool interactive, CancellationToken ct)
    {
        var registration = _store.ClientRegistration;
        var oauth = new ClientOAuthOptions
        {
            RedirectUri = OAuthCallbackListener.RedirectUri,
            TokenCache = _store,
            // A persisted registration avoids registering a new client on every launch.
            ClientId = registration?.ClientId,
            ClientSecret = registration?.ClientSecret,
            DynamicClientRegistration = new DynamicClientRegistrationOptions
            {
                ClientName = "ToggJirazer",
                ResponseDelegate = (response, _) =>
                {
                    _store.ClientRegistration = (response.ClientId, response.ClientSecret);
                    return Task.CompletedTask;
                },
            },
            AuthorizationCallbackHandler = interactive
                ? AuthorizeInBrowserAsync
                : (_, _) => throw new AtlassianAuthException("Not signed in"),
        };

        var transport = new HttpClientTransport(new HttpClientTransportOptions
        {
            Endpoint = Endpoint,
            TransportMode = HttpTransportMode.StreamableHttp,
            Name = "Atlassian",
            OAuth = oauth,
        });

        try
        {
            var client = await McpClient.CreateAsync(transport, new McpClientOptions
            {
                ClientInfo = new Implementation { Name = "ToggJirazer", Version = "1.0" },
                // The OAuth flow runs inside the initialize request, so allow time for the browser sign-in.
                InitializationTimeout = interactive ? SignInTimeout + TimeSpan.FromMinutes(1) : TimeSpan.FromSeconds(60),
            }, cancellationToken: ct);
            return new Connection(client, transport);
        }
        catch (Exception ex)
        {
            await transport.DisposeAsync();
            if (ClassifyFailure(ex) is { } classified) throw classified;
            throw;
        }
    }

    private async Task<AuthorizationResult?> AuthorizeInBrowserAsync(AuthorizationCallbackContext context, CancellationToken ct)
    {
        // Start listening before the browser opens so the redirect can't arrive too early.
        var callback = new OAuthCallbackListener().WaitForCallbackAsync(SignInTimeout, ct);
        _openBrowser(context.AuthorizationUri);
        var (code, state, iss) = await callback;
        return new AuthorizationResult { Code = code, State = state, Iss = iss };
    }

    /// <summary>
    /// Maps a failure from the SDK onto the <see cref="IJiraTools"/> exception types, or returns
    /// <c>null</c> when it should propagate unchanged (cancellation, plain network errors).
    /// The whole exception chain is searched because the SDK may wrap our callback's exception.
    /// </summary>
    internal static Exception? ClassifyFailure(Exception ex)
    {
        McpException? protocolError = null;
        for (var current = ex; current != null; current = current.InnerException)
        {
            switch (current)
            {
                case AtlassianAuthException or AtlassianRateLimitException or AtlassianToolException:
                    return current;
                case HttpRequestException { StatusCode: HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden } http:
                    return new AtlassianAuthException("Your Atlassian sign-in is no longer valid.", http);
                case HttpRequestException { StatusCode: HttpStatusCode.TooManyRequests }:
                    return new AtlassianRateLimitException();
                case OperationCanceledException or HttpRequestException:
                    return null;
                case McpException mcp:
                    protocolError ??= mcp;
                    break;
            }
        }

        // JSON-RPC errors (invalid params, unknown tool, ...) and other SDK failures.
        return protocolError is null ? null : new AtlassianToolException(protocolError.Message, statusCode: null);
    }

    private sealed record Connection(McpClient Client, HttpClientTransport Transport) : IAsyncDisposable
    {
        public async ValueTask DisposeAsync()
        {
            await Client.DisposeAsync();
            await Transport.DisposeAsync();
        }
    }

    private sealed class McpJiraTools(AtlassianMcp owner) : IJiraTools
    {
        public Task<JsonElement> CallToolAsync(string toolName, IReadOnlyDictionary<string, object?> arguments, CancellationToken ct = default) =>
            owner.CallNonInteractiveAsync(toolName, arguments, ct);
    }
}
