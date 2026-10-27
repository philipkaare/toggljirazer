using System.Text.Json;
using Avalonia.Headless.XUnit;
using ToggJirazer.Tests.Fakes;
using ModelContextProtocol.Authentication;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

public class AtlassianSignInControlTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"tj-signin-{Guid.NewGuid():N}");
    private string AuthPath => Path.Combine(_dir, "atlassian-auth.json");

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    private static readonly AtlassianSite[] Sites =
    [
        new("c1", "https://abc.atlassian.net", "ABC"),
        new("c2", "https://xyz.atlassian.net", "XYZ"),
    ];

    [AvaloniaFact]
    public async Task Shows_not_signed_in_without_token()
    {
        var control = new AtlassianSignInControl();
        control.Initialize(new AtlassianMcp(new AtlassianTokenStore(AuthPath), _ => throw new InvalidOperationException("no browser")));

        await control.RefreshStatusAsync();

        Assert.Equal("Not signed in", control.StatusText.Text);
        Assert.Equal("Sign in with Atlassian", control.SignInButton.Content);
    }

    [AvaloniaFact]
    public async Task Shows_not_signed_in_without_mcp()
    {
        var control = new AtlassianSignInControl();

        await control.RefreshStatusAsync();

        Assert.Equal("Not signed in", control.StatusText.Text);
        Assert.Equal("Sign in with Atlassian", control.SignInButton.Content);
    }

    [AvaloniaFact]
    public void Selection_is_kept_before_sites_load()
    {
        var control = new AtlassianSignInControl();
        control.SetSelection("c1", "https://abc.atlassian.net");

        Assert.Equal("c1", control.CloudId);
        Assert.Equal("https://abc.atlassian.net", control.SiteUrl);
    }

    [AvaloniaFact]
    public void Site_is_preselected_by_cloud_id()
    {
        var control = new AtlassianSignInControl();
        control.SetSelection("c2", "https://abc.atlassian.net");

        control.ShowSites(Sites);

        Assert.Equal("c2", control.CloudId);
        Assert.Equal("https://xyz.atlassian.net", control.SiteUrl);
    }

    [AvaloniaFact]
    public void Site_is_preselected_by_normalized_url_when_cloud_id_is_empty()
    {
        var control = new AtlassianSignInControl();
        control.SetSelection("", "HTTPS://XYZ.atlassian.net/");

        control.ShowSites(Sites);

        Assert.Equal("c2", control.CloudId);
        Assert.Equal("https://xyz.atlassian.net", control.SiteUrl);
    }

    [AvaloniaFact]
    public void Picking_a_site_updates_cloud_id_and_url()
    {
        var control = new AtlassianSignInControl();
        control.ShowSites(Sites);

        control.SiteComboBox.SelectedIndex = 0;

        Assert.Equal("c1", control.CloudId);
        Assert.Equal("https://abc.atlassian.net", control.SiteUrl);
    }

    [AvaloniaFact]
    public void Single_site_is_auto_selected_when_none_is_configured()
    {
        var control = new AtlassianSignInControl();

        control.ShowSites([Sites[1]]);

        Assert.Equal("c2", control.CloudId);
        Assert.Equal("https://xyz.atlassian.net", control.SiteUrl);
    }

    [AvaloniaFact]
    public void Single_site_is_auto_selected_when_the_configured_site_is_not_accessible()
    {
        var control = new AtlassianSignInControl();
        control.SetSelection("gone", "https://gone.atlassian.net");

        control.ShowSites([Sites[1]]);

        Assert.Equal("c2", control.CloudId);
    }

    [AvaloniaFact]
    public void No_site_is_selected_among_several_when_none_is_configured()
    {
        var control = new AtlassianSignInControl();

        control.ShowSites(Sites);

        Assert.Null(control.SiteComboBox.SelectedItem);
        Assert.Equal("", control.CloudId);
    }

    [AvaloniaFact]
    public async Task Sign_in_failure_is_not_overwritten_by_an_in_flight_status_check()
    {
        var control = new AtlassianSignInControl();
        var check = new TaskCompletionSource<IReadOnlyList<AtlassianSite>?>();
        control.LoadSitesIfSignedIn = () => check.Task;

        var refresh = control.RefreshStatusAsync();
        control.ShowSignInFailed("Sign-in failed: boom");
        check.SetResult(Sites);
        await refresh;

        Assert.Equal("Sign-in failed: boom", control.StatusText.Text);
        Assert.Equal("Sign in with Atlassian", control.SignInButton.Content);
    }

    private static void Click(Avalonia.Controls.Button button) =>
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Avalonia.Controls.Button.ClickEvent));

    [AvaloniaFact]
    public async Task Sign_in_shows_wait_message_url_and_cancel_button_and_cancel_stops_it()
    {
        var control = new AtlassianSignInControl();
        var url = new Uri("https://auth.example.test/authorize?client_id=x");
        var cancelled = new TaskCompletionSource();
        control.SignIn = async (ct, showUri) =>
        {
            showUri(url);
            try { await Task.Delay(Timeout.Infinite, ct); }
            catch (OperationCanceledException) { cancelled.SetResult(); throw new AtlassianAuthException("Sign-in was cancelled or timed out.", abortsSignIn: true); }
        };

        Click(control.SignInButton);
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        Assert.Equal("Cancel sign-in", control.SignInButton.Content);
        Assert.True(control.SignInButton.IsEnabled);
        Assert.Equal("Waiting for you to finish signing in in your browser…", control.StatusText.Text);
        Assert.True(control.AuthorizationUrlPanel.IsVisible);
        Assert.Equal(url.AbsoluteUri, control.AuthorizationUrlText.Text);

        Click(control.SignInButton);
        await cancelled.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        Assert.Equal("Sign-in cancelled", control.StatusText.Text);
        Assert.Equal("Sign in with Atlassian", control.SignInButton.Content);
        Assert.False(control.AuthorizationUrlPanel.IsVisible);
    }

    [AvaloniaFact]
    public async Task Status_refresh_during_a_waiting_sign_in_leaves_the_sign_in_state_alone()
    {
        var control = new AtlassianSignInControl();
        var checks = 0;
        control.LoadSitesIfSignedIn = () => { checks++; return Task.FromResult<IReadOnlyList<AtlassianSite>?>(null); };
        control.SignIn = async (ct, showUri) =>
        {
            showUri(new Uri("https://auth.example.test/authorize"));
            await Task.Delay(Timeout.Infinite, ct);
        };
        Click(control.SignInButton);
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();

        await control.RefreshStatusAsync();

        Assert.Equal(0, checks);
        Assert.True(control.IsSigningIn);
        Assert.Equal("Cancel sign-in", control.SignInButton.Content);
        Assert.Equal("Waiting for you to finish signing in in your browser…", control.StatusText.Text);
        Assert.True(control.AuthorizationUrlPanel.IsVisible);
        control.CancelSignIn();
    }

    [AvaloniaFact]
    public async Task Status_check_loads_sites_with_a_single_sites_call()
    {
        var store = new AtlassianTokenStore(AuthPath);
        await store.StoreTokensAsync(new TokenContainer { TokenType = "Bearer", AccessToken = "a", ObtainedAt = DateTimeOffset.UtcNow }, CancellationToken.None);
        var tools = new FakeJiraTools((tool, _) => tool == "getAccessibleAtlassianResources"
            ? JsonSerializer.SerializeToElement(new[] { new { id = "c1", url = "https://abc.atlassian.net", name = "ABC" } })
            : throw new InvalidOperationException($"Unexpected tool call: {tool}"));
        await using var mcp = new AtlassianMcp(store, _ => throw new InvalidOperationException("no browser"), tools);
        var control = new AtlassianSignInControl();
        control.Initialize(mcp);

        await control.RefreshStatusAsync();

        Assert.Single(tools.CallsTo("getAccessibleAtlassianResources"));
        Assert.Equal("Signed in (ABC)", control.StatusText.Text);
    }

    [Fact]
    public async Task SignOutAsync_deletes_stored_tokens()
    {
        var store = new AtlassianTokenStore(AuthPath);
        await store.StoreTokensAsync(new TokenContainer { TokenType = "Bearer", AccessToken = "a", ObtainedAt = DateTimeOffset.UtcNow }, CancellationToken.None);
        Assert.True(store.HasTokens);
        await using var mcp = new AtlassianMcp(store, _ => throw new InvalidOperationException("no browser"));

        await mcp.SignOutAsync();

        Assert.False(store.HasTokens);
    }
}
