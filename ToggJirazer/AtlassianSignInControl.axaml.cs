using Avalonia.Controls;
using Avalonia.Interactivity;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer;

/// <summary>
/// The Atlassian sign-in rows: a "Sign in with Atlassian" / "Sign out" button, a status line and
/// the Site dropdown. A browser opens only when the user clicks "Sign in with Atlassian".
/// </summary>
public partial class AtlassianSignInControl : UserControl
{
    private const string SignInText = "Sign in with Atlassian";
    private const string SignOutText = "Sign out";
    private const string NotSignedInText = "Not signed in";

    private AtlassianMcp? _mcp;
    private bool _signedIn;
    private int _refreshGeneration;
    // The configured site; used until (and unless) it matches a loaded site, so it survives a save before sign-in.
    private string _cloudId = string.Empty;
    private string _siteUrl = string.Empty;

    public AtlassianSignInControl()
    {
        InitializeComponent();
        LoadSitesIfSignedIn = LoadSitesFromMcpAsync;
    }

    /// <summary>The check behind <see cref="RefreshStatusAsync"/>: the sites when signed in, otherwise null. Replaceable by tests.</summary>
    internal Func<Task<IReadOnlyList<AtlassianSite>?>> LoadSitesIfSignedIn { get; set; }

    /// <summary>Raised after the user signs in or out from this control.</summary>
    public event EventHandler? SignInChanged;

    /// <summary>Whether the last status check (or sign-in) found a working sign-in.</summary>
    public bool IsSignedIn => _signedIn;

    public string CloudId => SelectedSite?.CloudId ?? _cloudId;

    public string SiteUrl => SelectedSite?.Url ?? _siteUrl;

    private AtlassianSite? SelectedSite => (SiteComboBox.SelectedItem as ComboBoxItem)?.Tag as AtlassianSite;

    public void Initialize(AtlassianMcp mcp) => _mcp = mcp;

    /// <summary>Pre-selects a site by cloud ID, or by site URL when the cloud ID is empty (legacy settings).</summary>
    public void SetSelection(string cloudId, string siteUrl)
    {
        _cloudId = cloudId ?? string.Empty;
        _siteUrl = siteUrl ?? string.Empty;
        SiteComboBox.SelectedItem = FindItem(_cloudId, _siteUrl);
    }

    /// <summary>
    /// Checks the stored sign-in without opening a browser and loads the sites when signed in.
    /// Never throws: a failed check is shown in the status line.
    /// </summary>
    public async Task RefreshStatusAsync()
    {
        var generation = ++_refreshGeneration;
        try
        {
            var sites = await LoadSitesIfSignedIn();

            if (generation != _refreshGeneration) return; // a newer check has taken over

            if (sites == null)
            {
                ShowSignedOut(NotSignedInText);
                return;
            }

            _signedIn = true;
            SignInButton.Content = SignOutText;
            ShowSites(sites);
        }
        catch (Exception ex)
        {
            if (generation != _refreshGeneration) return;
            ShowSignedOut($"Could not check sign-in: {ex.Message}");
        }
    }

    private Task<IReadOnlyList<AtlassianSite>?> LoadSitesFromMcpAsync() =>
        _mcp?.TryGetSitesAsync(CancellationToken.None) ?? Task.FromResult<IReadOnlyList<AtlassianSite>?>(null);

    /// <summary>
    /// Fills the Site dropdown and re-applies the configured selection. When the configured site isn't
    /// among them (or none is configured) and there is exactly one site, that site is selected.
    /// </summary>
    internal void ShowSites(IReadOnlyList<AtlassianSite> sites)
    {
        var items = sites
            .Select(site => new ComboBoxItem { Content = $"{site.Name} ({site.Url})", Tag = site })
            .ToList();
        SiteComboBox.ItemsSource = items;
        SiteComboBox.SelectedItem = FindItem(_cloudId, _siteUrl) ?? (items.Count == 1 ? items[0] : null);
        UpdateSignedInStatus();
    }

    private ComboBoxItem? FindItem(string cloudId, string siteUrl)
    {
        var items = SiteComboBox.ItemsSource?.Cast<ComboBoxItem>() ?? [];
        if (cloudId.Length > 0)
            return items.FirstOrDefault(i => ((AtlassianSite)i.Tag!).CloudId == cloudId);

        var url = NormalizeUrl(siteUrl);
        return url.Length == 0 ? null : items.FirstOrDefault(i => NormalizeUrl(((AtlassianSite)i.Tag!).Url) == url);
    }

    private static string NormalizeUrl(string url) => url.Trim().TrimEnd('/').ToLowerInvariant();

    /// <summary>Shows a failed sign-in or sign-out; a status check still in flight can no longer overwrite it.</summary>
    internal void ShowSignInFailed(string status)
    {
        _refreshGeneration++;
        ShowSignedOut(status);
    }

    private void ShowSignedOut(string status)
    {
        _signedIn = false;
        SignInButton.Content = SignInText;
        StatusText.Text = status;
        SiteComboBox.ItemsSource = null;
    }

    private void UpdateSignedInStatus()
    {
        if (!_signedIn) return;
        StatusText.Text = SelectedSite is { } site ? $"Signed in ({site.Name})" : "Signed in — pick your site";
    }

    private void SiteComboBox_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (SelectedSite is { } site)
        {
            _cloudId = site.CloudId;
            _siteUrl = site.Url;
        }
        UpdateSignedInStatus();
    }

    private async void SignInButton_Click(object? sender, RoutedEventArgs e)
    {
        if (_mcp == null)
        {
            ShowSignedOut(NotSignedInText);
            return;
        }

        var signingOut = _signedIn;
        SignInButton.IsEnabled = false;
        try
        {
            if (signingOut)
            {
                await _mcp.SignOutAsync();
            }
            else
            {
                StatusText.Text = "Signing in — finish in your browser…";
                await _mcp.SignInAsync(CancellationToken.None);
            }

            await RefreshStatusAsync();
            SignInChanged?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            ShowSignInFailed(signingOut ? $"Sign-out failed: {ex.Message}" : $"Sign-in failed: {ex.Message}");
        }
        finally
        {
            SignInButton.IsEnabled = true;
        }
    }
}
