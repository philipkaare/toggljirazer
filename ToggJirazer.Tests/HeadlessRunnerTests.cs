using System.Text.Json;
using ToggJirazer.Cli;
using ToggJirazer.Models;
using ToggJirazer.Services.Atlassian;
using ToggJirazer.Tests.Fakes;
using Xunit;

namespace ToggJirazer.Tests;

public class HeadlessRunnerTests
{
    private const string NotSignedIn =
        "Not signed in to Atlassian. Open ToggJirazer and use 'Sign in with Atlassian' on the Jira tab.";

    private static AppConfig Valid() => new()
    {
        Toggl = new TogglConfig { ApiToken = "t", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
        Jira = new JiraConfig { CloudId = "cid-1", SiteUrl = "https://x.atlassian.net" },
        Report = new ReportConfig { OutputFile = "r.xlsx", Format = "xlsx" }
    };

    private static FakeJiraTools UnusedTools() =>
        new((tool, _) => throw new InvalidOperationException($"Unexpected tool call: {tool}"));

    [Fact]
    public async Task Headless_not_signed_in_exits_2()
    {
        var stderr = new StringWriter();
        var tools = UnusedTools();

        var exitCode = await HeadlessRunner.RunAsync(Valid(), () => Task.FromResult(false), tools, stderr);

        Assert.Equal(2, exitCode);
        Assert.Equal(NotSignedIn, stderr.ToString().Trim());
        Assert.Empty(tools.Calls);
    }

    [Fact]
    public async Task Headless_invalid_config_exits_2_before_sign_in_check()
    {
        var stderr = new StringWriter();
        var config = Valid();
        config.Jira.CloudId = "";
        var checkedSignIn = false;

        var exitCode = await HeadlessRunner.RunAsync(config,
            () => { checkedSignIn = true; return Task.FromResult(true); }, UnusedTools(), stderr);

        Assert.Equal(2, exitCode);
        Assert.Contains("Pick your Atlassian site (Jira tab).", stderr.ToString());
        Assert.False(checkedSignIn);
    }
}

public class AtlassianSignInProbeTests
{
    [Fact]
    public async Task Probe_returns_true_when_sites_load()
    {
        var resets = 0;
        var signedIn = await AtlassianMcp.ProbeSignInAsync(() => Task.CompletedTask, () => { resets++; return Task.CompletedTask; });

        Assert.True(signedIn);
        Assert.Equal(0, resets);
    }

    [Fact]
    public async Task Probe_auth_failure_is_not_signed_in_and_resets()
    {
        var resets = 0;
        var signedIn = await AtlassianMcp.ProbeSignInAsync(
            () => throw new AtlassianAuthException("Not signed in"),
            () => { resets++; return Task.CompletedTask; });

        Assert.False(signedIn);
        Assert.Equal(1, resets);
    }

    [Fact]
    public async Task Probe_tool_failure_such_as_refresh_invalid_grant_is_not_signed_in_and_resets()
    {
        var resets = 0;
        var signedIn = await AtlassianMcp.ProbeSignInAsync(
            () => throw new AtlassianToolException("invalid_grant", statusCode: null),
            () => { resets++; return Task.CompletedTask; });

        Assert.False(signedIn);
        Assert.Equal(1, resets);
    }

    [Fact]
    public async Task Probe_network_failure_propagates()
    {
        await Assert.ThrowsAsync<HttpRequestException>(() => AtlassianMcp.ProbeSignInAsync(
            () => throw new HttpRequestException("offline"),
            () => Task.CompletedTask));
    }
}
