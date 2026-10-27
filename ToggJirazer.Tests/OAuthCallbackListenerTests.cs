using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

public class OAuthCallbackListenerTests
{
    [Fact]
    public void ParseQuery_extracts_code_state_iss()
    {
        var (code, state, iss) = OAuthCallbackListener.ParseQuery("?code=abc123&state=xyz%3D&iss=https%3A%2F%2Fauth.example.test");

        Assert.Equal("abc123", code);
        Assert.Equal("xyz=", state);
        Assert.Equal("https://auth.example.test", iss);
    }

    [Fact]
    public void ParseQuery_error_throws_auth()
    {
        Assert.Throws<AtlassianAuthException>(() => OAuthCallbackListener.ParseQuery("?error=access_denied"));
    }

    [Fact]
    public void ParseQuery_missing_code_throws_auth()
    {
        Assert.Throws<AtlassianAuthException>(() => OAuthCallbackListener.ParseQuery("?state=xyz"));
    }
}
