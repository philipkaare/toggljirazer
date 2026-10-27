using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using ToggJirazer.Services;

namespace ToggJirazer.Setup.Steps;

/// <summary>Wizard step 2: sign in to Toggl with email and password, or with an API token.</summary>
public partial class TogglSignInStep : UserControl
{
    private const string TogglProfileUrl = "https://track.toggl.com/profile";

    private WizardContext _ctx = null!;
    private bool _savedTokenChecked;

    public TogglSignInStep()
    {
        InitializeComponent();
        // Tunnel handlers: they run before the TextBox's own key handling.
        TogglPassword.AddHandler(KeyDownEvent, TogglPassword_KeyDown, RoutingStrategies.Tunnel);
        TogglTokenBox.AddHandler(KeyDownEvent, TogglTokenBox_KeyDown, RoutingStrategies.Tunnel);
    }

    internal void Attach(WizardContext context) => _ctx = context;

    /// <summary>True while a Toggl sign-in or token check runs: navigation waits for its answer.</summary>
    internal bool IsBusy { get; private set; }

    /// <summary>The Toggl line of the Done summary.</summary>
    internal string Summary =>
        TogglSignedInPanel.IsVisible && TogglSignedInText.Text?.StartsWith("Signed in as ") == true
            ? TogglSignedInText.Text
            : "API token set";

    internal async Task EnterAsync()
    {
        if (_ctx.State.CanGoNext && !_savedTokenChecked)
        {
            // Show who the saved token belongs to; a failed check keeps the token and offers the sign-in form.
            _savedTokenChecked = true;
            ShowTogglSignedIn("Checking your saved Toggl token…");
            var answered = await RunTogglCheckAsync(s => s.ValidateTokenAsync(_ctx.Draft.Toggl.ApiToken), busyText: string.Empty,
                failurePrefix: "Your saved Toggl token could not be checked: ");
            // The step changed before Toggl answered: check again when the step is shown next time.
            if (!answered) _savedTokenChecked = false;
        }
        else if (!TogglSignedInPanel.IsVisible)
        {
            ShowTogglForm();
        }
    }

    internal async Task SignInToTogglAsync()
    {
        var email = (TogglEmail.Text ?? string.Empty).Trim();
        var password = TogglPassword.Text ?? string.Empty;
        TogglPassword.Text = string.Empty;
        if (email.Length == 0 || password.Length == 0)
        {
            TogglStatusText.Text = "Enter your Toggl email and password.";
            return;
        }

        await RunTogglCheckAsync(s => s.SignInAsync(email, password), "Signing in…");
    }

    internal async Task CheckTogglTokenAsync()
    {
        var token = (TogglTokenBox.Text ?? string.Empty).Trim();
        if (token.Length == 0)
        {
            TogglStatusText.Text = "Paste your API token from your Toggl profile.";
            return;
        }

        await RunTogglCheckAsync(s => s.ValidateTokenAsync(token), "Checking token…");
    }

    /// <summary>
    /// Runs a Toggl account check; on success uses its API token. Errors are shown inline, never thrown.
    /// Navigation is disabled until Toggl answers. Returns false when the step changed before the answer
    /// arrived; that answer is dropped.
    /// </summary>
    private async Task<bool> RunTogglCheckAsync(Func<TogglSetupService, Task<TogglAccount>> check, string busyText,
        string failurePrefix = "")
    {
        var generation = _ctx.LoadGeneration;
        SetTogglBusy(true);
        TogglStatusText.Text = busyText;
        try
        {
            using var service = _ctx.TogglFactory();
            var account = await check(service);
            if (generation != _ctx.LoadGeneration) return false;
            _ctx.Draft.Toggl.ApiToken = account.ApiToken;
            ShowTogglSignedIn($"Signed in as {(account.FullName.Length > 0 ? account.FullName : account.Email)}");
            TogglStatusText.Text = string.Empty;
        }
        catch (Exception ex)
        {
            if (generation != _ctx.LoadGeneration) return false;
            if (failurePrefix.Length > 0) ShowTogglForm();
            TogglStatusText.Text = failurePrefix + (ex is TogglSetupException ? ex.Message : $"Could not reach Toggl: {ex.Message}");
        }
        finally
        {
            SetTogglBusy(false);
        }
        return true;
    }

    private void SetTogglBusy(bool busy)
    {
        IsBusy = busy;
        TogglSignInButton.IsEnabled = !busy;
        CheckTokenButton.IsEnabled = !busy;
        UseDifferentAccountLink.IsEnabled = !busy;
        _ctx.UpdateButtons();
    }

    private void ShowTogglSignedIn(string text)
    {
        TogglSignedInText.Text = text;
        TogglSignedInPanel.IsVisible = true;
        TogglFormPanel.IsVisible = false;
        TogglTokenPanel.IsVisible = false;
    }

    private void ShowTogglForm()
    {
        TogglSignedInPanel.IsVisible = false;
        TogglFormPanel.IsVisible = true;
    }

    internal void ShowTokenEntry()
    {
        TogglTokenPanel.IsVisible = true;
        try
        {
            _ctx.OpenUrl(TogglProfileUrl);
        }
        catch (Exception)
        {
            TogglStatusText.Text = $"Open {TogglProfileUrl} in your browser to copy your API token.";
        }
        TogglTokenBox.Focus();
    }

    private async void TogglSignIn_Click(object? sender, RoutedEventArgs e) => await SignInToTogglAsync();

    private async void CheckToken_Click(object? sender, RoutedEventArgs e) => await CheckTogglTokenAsync();

    private void UseToken_Click(object? sender, RoutedEventArgs e) => ShowTokenEntry();

    private async void TogglPassword_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || !TogglSignInButton.IsEnabled) return;
        e.Handled = true;
        await SignInToTogglAsync();
    }

    private async void TogglTokenBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || !CheckTokenButton.IsEnabled) return;
        e.Handled = true;
        await CheckTogglTokenAsync();
    }

    private void UseDifferentAccount_Click(object? sender, RoutedEventArgs e)
    {
        ShowTogglForm();
        TogglStatusText.Text = string.Empty;
        TogglEmail.Focus();
    }
}
