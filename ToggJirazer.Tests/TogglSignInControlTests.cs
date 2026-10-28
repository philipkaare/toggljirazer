using System.Net;
using Avalonia.Headless.XUnit;
using ToggJirazer.Services;
using ToggJirazer.Tests.Fakes;
using Xunit;

namespace ToggJirazer.Tests;

public class TogglSignInControlTests
{
    private const string Me = """{"api_token":"t1","fullname":"Ann","email":"a@x"}""";

    private static TogglSignInControl Control(Func<HttpRequestMessage, HttpResponseMessage> responder)
    {
        var control = new TogglSignInControl { OpenUrl = _ => { } };
        control.Initialize(() => new TogglSetupService(new FakeHttpHandler(responder)));
        return control;
    }

    [AvaloniaFact]
    public void Shows_signed_out_form_without_token()
    {
        var control = new TogglSignInControl();
        control.SetToken("");

        Assert.Equal("Not signed in", control.StatusText.Text);
        Assert.False(control.SignOutButton.IsVisible);
        Assert.True(control.EmailBox.IsVisible);
        Assert.True(control.SignInButton.IsVisible);
    }

    [AvaloniaFact]
    public void Sample_placeholder_token_counts_as_signed_out()
    {
        var control = new TogglSignInControl();
        control.SetToken("your-toggl-api-token");

        Assert.Equal("", control.ApiToken);
        Assert.False(control.SignOutButton.IsVisible);
    }

    [AvaloniaFact]
    public async Task Saved_token_is_checked_and_shows_who_is_signed_in()
    {
        var control = Control(_ => FakeHttpHandler.Json(Me));
        control.SetToken("t1");

        await control.RefreshStatusAsync();

        Assert.Equal("Signed in as Ann", control.StatusText.Text);
        Assert.True(control.SignOutButton.IsVisible);
        Assert.False(control.EmailBox.IsVisible);
    }

    [AvaloniaFact]
    public async Task Failed_check_keeps_the_saved_token()
    {
        var control = Control(_ => FakeHttpHandler.Json("{}", HttpStatusCode.Unauthorized));
        control.SetToken("old");

        await control.RefreshStatusAsync();

        Assert.Equal("old", control.ApiToken);
        Assert.StartsWith("Your saved Toggl token could not be checked:", control.StatusText.Text);
        Assert.True(control.SignOutButton.IsVisible);
    }

    [AvaloniaFact]
    public async Task Sign_in_with_email_and_password_uses_the_account_token()
    {
        var control = Control(_ => FakeHttpHandler.Json(Me));
        var changed = 0;
        control.TokenChanged += (_, _) => changed++;
        control.EmailBox.Text = "a@x";
        control.PasswordBox.Text = "secret";

        await control.SignInAsync();

        Assert.Equal("t1", control.ApiToken);
        Assert.Equal("Signed in as Ann", control.StatusText.Text);
        Assert.Equal(1, changed);
        Assert.True(string.IsNullOrEmpty(control.PasswordBox.Text));
    }

    [AvaloniaFact]
    public async Task Wrong_password_stays_signed_out()
    {
        var control = Control(_ => FakeHttpHandler.Json("{}", HttpStatusCode.Forbidden));
        var changed = 0;
        control.TokenChanged += (_, _) => changed++;
        control.EmailBox.Text = "a@x";
        control.PasswordBox.Text = "wrong";

        await control.SignInAsync();

        Assert.Equal("", control.ApiToken);
        Assert.StartsWith("Wrong email or password.", control.StatusText.Text);
        Assert.Equal(0, changed);
    }

    [AvaloniaFact]
    public async Task Api_token_can_be_used_instead()
    {
        var control = Control(_ => FakeHttpHandler.Json(Me));
        control.ShowTokenEntry();
        Assert.True(control.TokenBox.IsVisible);
        control.TokenBox.Text = "t1";

        await control.CheckTokenAsync();

        Assert.Equal("t1", control.ApiToken);
        Assert.False(control.TokenBox.IsVisible);
    }

    [AvaloniaFact]
    public void Sign_out_clears_the_token()
    {
        var control = new TogglSignInControl();
        var changed = 0;
        control.TokenChanged += (_, _) => changed++;
        control.SetToken("t1");

        control.SignOut();

        Assert.Equal("", control.ApiToken);
        Assert.Equal("Not signed in", control.StatusText.Text);
        Assert.True(control.EmailBox.IsVisible);
        Assert.Equal(1, changed);
    }
}
