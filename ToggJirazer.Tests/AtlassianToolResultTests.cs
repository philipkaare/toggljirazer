using System.Net;
using System.Text.Json;
using ModelContextProtocol;
using ToggJirazer.Services.Atlassian;
using Xunit;

namespace ToggJirazer.Tests;

public class AtlassianToolResultTests
{
    private const string RateLimitMessage =
        "Atlassian rate limit reached — try again later (the 1-day cache keeps what was fetched).";

    [Fact]
    public void Parse_returns_payload_for_success()
    {
        const string text = """
            {"issues":{"nodes":[{"id":"10001","key":"ABC-1","fields":{"summary":"First"}}],
             "pageInfo":{"hasNextPage":false,"endCursor":null}},
             "context":{"cloudId":"cloud-1"}}
            """;

        var payload = AtlassianToolResult.Parse(text, isError: false);

        var node = payload.GetProperty("issues").GetProperty("nodes")[0];
        Assert.Equal("ABC-1", node.GetProperty("key").GetString());
    }

    [Fact]
    public void Parse_401_throws_auth()
    {
        const string text = """{"error":true,"message":"Unauthorized","statusCode":401}""";

        Assert.Throws<AtlassianAuthException>(() => AtlassianToolResult.Parse(text, isError: true));
    }

    [Fact]
    public void Rate_limit_throws_with_message()
    {
        const string status429 = """{"error":true,"message":"Too many requests","statusCode":429}""";
        const string remainingZero = """{"error":true,"message":"Slow down","statusCode":400,"rateLimit":{"remaining":0}}""";

        var byStatus = Assert.Throws<AtlassianRateLimitException>(() => AtlassianToolResult.Parse(status429, isError: true));
        var byRemaining = Assert.Throws<AtlassianRateLimitException>(() => AtlassianToolResult.Parse(remainingZero, isError: true));

        Assert.Equal(RateLimitMessage, byStatus.Message);
        Assert.Equal(RateLimitMessage, byRemaining.Message);
    }

    [Fact]
    public void Parse_400_throws_tool_exception_with_status()
    {
        const string text = """{"error":true,"message":"Unbounded JQL queries are not allowed here.","statusCode":400,"rateLimit":{"remaining":150}}""";

        var ex = Assert.Throws<AtlassianToolException>(() => AtlassianToolResult.Parse(text, isError: true));

        Assert.Equal(400, ex.StatusCode);
        Assert.Contains("Unbounded JQL queries are not allowed here.", ex.Message);
    }

    [Fact]
    public void Sites_are_deduped_by_id()
    {
        using var doc = JsonDocument.Parse("""
            [
              {"id":"cloud-1","url":"https://abc.atlassian.net","name":"abc","scopes":["read:jira-work"],"avatarUrl":"x"},
              {"id":"cloud-1","url":"https://abc.atlassian.net","name":"abc","scopes":["read:confluence"],"avatarUrl":"x"},
              {"id":"cloud-2","url":"https://def.atlassian.net","name":"def","scopes":[],"avatarUrl":"y"}
            ]
            """);

        var sites = AtlassianMcp.ParseSites(doc.RootElement);

        Assert.Equal(
            [new AtlassianSite("cloud-1", "https://abc.atlassian.net", "abc"), new AtlassianSite("cloud-2", "https://def.atlassian.net", "def")],
            sites);
    }

    [Fact]
    public void Sites_are_read_from_the_v2_data_envelope()
    {
        // The shape mcp.atlassian.com/v2/mcp returns: no "name", so it is taken from the site's host.
        using var doc = JsonDocument.Parse("""
            {"data":{"resources":[
              {"cloudId":"cloud-1","url":"https://abc.atlassian.net","products":[{"id":"jira","access":"read-write"}]},
              {"cloudId":"cloud-2","url":"https://def.atlassian.net","name":"Def","products":[]}
            ]}}
            """);

        var sites = AtlassianMcp.ParseSites(doc.RootElement);

        Assert.Equal(
            [new AtlassianSite("cloud-1", "https://abc.atlassian.net", "abc"), new AtlassianSite("cloud-2", "https://def.atlassian.net", "Def")],
            sites);
    }

    [Fact]
    public void Transport_429_classifies_as_rate_limit()
    {
        var ex = new HttpRequestException("Too Many Requests", null, HttpStatusCode.TooManyRequests);

        var classified = Assert.IsType<AtlassianRateLimitException>(AtlassianMcp.ClassifyFailure(ex));
        Assert.Equal(RateLimitMessage, classified.Message);
    }

    [Fact]
    public void Transport_401_classifies_as_auth()
    {
        var ex = new McpException("wrapped", new HttpRequestException("Unauthorized", null, HttpStatusCode.Unauthorized));

        Assert.IsType<AtlassianAuthException>(AtlassianMcp.ClassifyFailure(ex));
    }

    [Fact]
    public void Wrapped_auth_exception_is_returned_as_is()
    {
        var auth = new AtlassianAuthException("Not signed in");

        Assert.Same(auth, AtlassianMcp.ClassifyFailure(new McpException("Failed to handle unauthorized response", auth)));
    }

    [Fact]
    public void Protocol_error_classifies_as_tool_exception()
    {
        var ex = new McpProtocolException("Unknown tool: nope", McpErrorCode.InvalidParams);

        var classified = Assert.IsType<AtlassianToolException>(AtlassianMcp.ClassifyFailure(ex));
        Assert.Contains("Unknown tool: nope", classified.Message);
        Assert.Null(classified.StatusCode);
    }

    [Fact]
    public void Cancellation_and_network_errors_propagate()
    {
        Assert.Null(AtlassianMcp.ClassifyFailure(new OperationCanceledException()));
        Assert.Null(AtlassianMcp.ClassifyFailure(new HttpRequestException("No such host is known.")));
    }
}
