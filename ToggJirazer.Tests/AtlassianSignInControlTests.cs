using Avalonia.Headless.XUnit;
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
