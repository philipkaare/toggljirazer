using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using ToggJirazer.Cli;
using ToggJirazer.Services;
using ToggJirazer.Setup;

namespace ToggJirazer;

/// <summary>
/// The Toggl sign-in rows: while signed out, email and password (or an API token) and "Sign in"; while
/// signed in, who the token belongs to and "Sign out". Signing in or out changes <see cref="ApiToken"/>
/// and raises <see cref="TokenChanged"/>.
/// </summary>
public partial class TogglSignInControl : UserControl
{
    private const string TogglProfileUrl = "https://track.toggl.com/profile";
    private const string NotSignedInText = "Not signed in";

    private Func<TogglSetupService>? _togglFactory;
    private string _apiToken = string.Empty;
    private int _generation; // bumped on every token change, so a late answer about an older token is dropped

    public TogglSignInControl()
    {
        InitializeComponent();
        // Tunnel handlers: they run before the TextBox's own key handling.
        PasswordBox.AddHandler(KeyDownEvent, PasswordBox_KeyDown, RoutingStrategies.Tunnel);
        TokenBox.AddHandler(KeyDownEvent, TokenBox_KeyDown, RoutingStrategies.Tunnel);
        ShowSignedOut(NotSignedInText);
    }

    /// <summary>Raised after the user signs in or out from this control.</summary>
    public event EventHandler? TokenChanged;

    /// <summary>The API token in use; empty when signed out.</summary>
    public string ApiToken => _apiToken;

    /// <summary>Opens a URL in the browser. Replaceable by tests.</summary>
    internal Action<string> OpenUrl { get; set; } = PlatformOpener.Open;

    /// <summary>True while a sign-in or token check waits for Toggl.</summary>
    public bool IsBusy { get; private set; }

    /// <summary>Enables talking to Toggl. Without it the control only shows whether a token is set.</summary>
    public void Initialize(Func<TogglSetupService> togglFactory) => _togglFactory = togglFactory;

    /// <summary>Shows a saved token (e.g. from the settings file) without checking it. The sample placeholder counts as none.</summary>
    public void SetToken(string? apiToken)
    {
        _generation++;
        _apiToken = (apiToken ?? string.Empty).Trim();
        if (_apiToken == SetupWizardState.TokenPlaceholder) _apiToken = string.Empty;
        if (_apiToken.Length == 0) ShowSignedOut(NotSignedInText);
        else ShowSignedIn("Signed in (API token set)");
    }

    /// <summary>
    /// Asks Toggl who the current token belongs to. Never throws: a failed check keeps the token and says so
    /// in the status line. Does nothing when there is no token or no Toggl connection.
    /// </summary>
    public async Task RefreshStatusAsync()
    {
        if (_apiToken.Length == 0 || _togglFactory == null) return;
        var token = _apiToken;
        StatusText.Text = "Checking your Toggl token…";
        await RunCheckAsync(s => s.ValidateTokenAsync(token), onFailure: message =>
            ShowSignedIn($"Your saved Toggl token could not be checked: {message}"));
    }

    internal async Task SignInAsync()
    {
        var email = (EmailBox.Text ?? string.Empty).Trim();
        var password = PasswordBox.Text ?? string.Empty;
        PasswordBox.Text = string.Empty;
        if (email.Length == 0 || password.Length == 0)
        {
            StatusText.Text = "Enter your Toggl email and password.";
            return;
        }

        StatusText.Text = "Signing in…";
        await RunCheckAsync(s => s.SignInAsync(email, password), onFailure: ShowSignedOut, raiseChanged: true);
    }

    internal async Task CheckTokenAsync()
    {
        var token = (TokenBox.Text ?? string.Empty).Trim();
        if (token.Length == 0)
        {
            StatusText.Text = "Paste your API token from your Toggl profile.";
            return;
        }

        StatusText.Text = "Checking token…";
        await RunCheckAsync(s => s.ValidateTokenAsync(token), onFailure: ShowSignedOut, raiseChanged: true);
    }

    internal void SignOut()
    {
        SetToken(string.Empty);
        TokenChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Runs a Toggl account check; on success uses its token. Errors go to <paramref name="onFailure"/>.</summary>
    private async Task RunCheckAsync(Func<TogglSetupService, Task<TogglAccount>> check, Action<string> onFailure,
        bool raiseChanged = false)
    {
        if (_togglFactory == null)
        {
            onFailure("Toggl is not available. Restart ToggJirazer.");
            return;
        }

        var generation = ++_generation;
        SetBusy(true);
        try
        {
            using var service = _togglFactory();
            var account = await check(service);
            if (generation != _generation) return;
            _apiToken = account.ApiToken;
            ShowSignedIn($"Signed in as {(account.FullName.Length > 0 ? account.FullName : account.Email)}");
            if (raiseChanged) TokenChanged?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            if (generation != _generation) return;
            onFailure(ex is TogglSetupException ? ex.Message : $"Could not reach Toggl: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void SetBusy(bool busy)
    {
        IsBusy = busy;
        SignInButton.IsEnabled = !busy;
        CheckTokenButton.IsEnabled = !busy;
        SignOutButton.IsEnabled = !busy;
    }

    private void ShowSignedIn(string status)
    {
        StatusText.Text = status;
        SignOutButton.IsVisible = true;
        SetFormVisible(false);
    }

    private void ShowSignedOut(string status)
    {
        StatusText.Text = status;
        SignOutButton.IsVisible = false;
        SetFormVisible(true);
    }

    private void SetFormVisible(bool visible)
    {
        EmailLabel.IsVisible = EmailBox.IsVisible = visible;
        PasswordLabel.IsVisible = PasswordBox.IsVisible = SignInButton.IsVisible = visible;
        UseTokenLink.IsVisible = visible;
        if (!visible) SetTokenEntryVisible(false);
    }

    private void SetTokenEntryVisible(bool visible)
    {
        TokenLabel.IsVisible = TokenBox.IsVisible = CheckTokenButton.IsVisible = visible;
        if (!visible) TokenBox.Text = string.Empty;
    }

    internal void ShowTokenEntry()
    {
        SetTokenEntryVisible(true);
        try
        {
            OpenUrl(TogglProfileUrl);
        }
        catch (Exception)
        {
            StatusText.Text = $"Open {TogglProfileUrl} in your browser to copy your API token.";
        }
        TokenBox.Focus();
    }

    private async void SignInButton_Click(object? sender, RoutedEventArgs e) => await SignInAsync();

    private async void CheckTokenButton_Click(object? sender, RoutedEventArgs e) => await CheckTokenAsync();

    private void SignOutButton_Click(object? sender, RoutedEventArgs e) => SignOut();

    private void UseTokenLink_Click(object? sender, RoutedEventArgs e) => ShowTokenEntry();

    private async void PasswordBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || !SignInButton.IsEnabled) return;
        e.Handled = true;
        await SignInAsync();
    }

    private async void TokenBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || !CheckTokenButton.IsEnabled) return;
        e.Handled = true;
        await CheckTokenAsync();
    }
}
