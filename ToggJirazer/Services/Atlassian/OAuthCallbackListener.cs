using System.Net;
using System.Text;
using System.Web;

namespace ToggJirazer.Services.Atlassian;

/// <summary>
/// One-shot local HTTP listener that receives the OAuth authorization redirect.
/// The redirect URI is fixed because the dynamic client registration is tied to it.
/// </summary>
public sealed class OAuthCallbackListener
{
    public static readonly Uri RedirectUri = new("http://localhost:33418/callback");

    private const string SuccessPage =
        "<!doctype html><html><head><meta charset=\"utf-8\"><title>ToggJirazer</title></head>" +
        "<body style=\"font-family:sans-serif\"><p>Signed in — you can close this tab.</p></body></html>";

    /// <summary>
    /// Listens on <see cref="RedirectUri"/> until the callback arrives, serves a small confirmation
    /// page and returns <c>code</c>, <c>state</c> and <c>iss</c>. The listener is started before this
    /// method returns, so the browser can be opened right after calling it.
    /// </summary>
    public Task<(string Code, string? State, string? Iss)> WaitForCallbackAsync(TimeSpan timeout, CancellationToken ct)
    {
        var listener = new HttpListener();
        listener.Prefixes.Add($"{RedirectUri.Scheme}://{RedirectUri.Authority}/");
        try
        {
            listener.Start();
        }
        catch (HttpListenerException ex)
        {
            listener.Close();
            throw new AtlassianAuthException(
                $"Could not listen for the sign-in callback on port {RedirectUri.Port} — another program may be using it. Close it and try again.", ex);
        }

        return ReceiveAsync(listener, timeout, ct);
    }

    private static async Task<(string Code, string? State, string? Iss)> ReceiveAsync(HttpListener listener, TimeSpan timeout, CancellationToken ct)
    {
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeoutCts.CancelAfter(timeout);
        // GetContextAsync has no cancellation token; stopping the listener aborts it.
        var registration = timeoutCts.Token.Register(listener.Stop);

        try
        {
            while (true)
            {
                HttpListenerContext context;
                try
                {
                    context = await listener.GetContextAsync();
                }
                catch (Exception ex) when (ex is HttpListenerException or ObjectDisposedException && timeoutCts.IsCancellationRequested)
                {
                    throw new AtlassianAuthException("Sign-in was cancelled or timed out.", ex);
                }

                if (!string.Equals(context.Request.Url?.AbsolutePath, RedirectUri.AbsolutePath, StringComparison.Ordinal))
                {
                    // e.g. the browser asking for /favicon.ico
                    context.Response.StatusCode = 404;
                    context.Response.Close();
                    continue;
                }

                try
                {
                    var result = ParseQuery(context.Request.Url!.Query);
                    await RespondAsync(context.Response, SuccessPage);
                    return result;
                }
                catch (AtlassianAuthException ex)
                {
                    await RespondAsync(context.Response, $"<!doctype html><html><head><meta charset=\"utf-8\"></head><body><p>{WebUtility.HtmlEncode(ex.Message)}</p></body></html>");
                    throw;
                }
            }
        }
        finally
        {
            registration.Dispose(); // before Close, so a late cancellation can't Stop a closed listener
            listener.Close();
        }
    }

    private static async Task RespondAsync(HttpListenerResponse response, string html)
    {
        var bytes = Encoding.UTF8.GetBytes(html);
        response.ContentType = "text/html; charset=utf-8";
        response.ContentLength64 = bytes.Length;
        await response.OutputStream.WriteAsync(bytes);
        response.Close();
    }

    /// <summary>
    /// Extracts <c>code</c>, <c>state</c> and <c>iss</c> from the callback query string.
    /// Throws <see cref="AtlassianAuthException"/> when the query carries <c>error=</c> or no code.
    /// </summary>
    public static (string Code, string? State, string? Iss) ParseQuery(string query)
    {
        var values = HttpUtility.ParseQueryString(query);

        var error = values["error"];
        if (!string.IsNullOrEmpty(error))
        {
            var description = values["error_description"];
            throw new AtlassianAuthException(
                $"Sign-in failed: {(string.IsNullOrEmpty(description) ? error : $"{error} ({description})")}");
        }

        var code = values["code"];
        if (string.IsNullOrEmpty(code))
            throw new AtlassianAuthException("Sign-in failed: Atlassian did not return an authorization code.");

        return (code, values["state"], values["iss"]);
    }
}
