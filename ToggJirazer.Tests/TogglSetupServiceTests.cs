using System.Net;
using System.Text;
using ToggJirazer.Services;
using ToggJirazer.Tests.Fakes;
using Xunit;

namespace ToggJirazer.Tests;

public class TogglSetupServiceTests
{
    private const string WrongPassword =
        "Wrong email or password. If you sign in to Toggl with Google, Apple or SSO, use an API token instead.";

    private static string Decode(HttpRequestMessage r) =>
        Encoding.UTF8.GetString(Convert.FromBase64String(r.Headers.Authorization!.Parameter!));

    [Fact]
    public async Task SignIn_returns_token_and_name()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("""{"api_token":"t1","fullname":"Ann","email":"a@x"}"""));
        using var svc = new TogglSetupService(handler);

        var account = await svc.SignInAsync("a@x", "pw");

        Assert.Equal(new TogglAccount("t1", "Ann", "a@x"), account);
        var req = Assert.Single(handler.Requests);
        Assert.Equal("https://api.track.toggl.com/api/v9/me", req.RequestUri!.ToString());
        Assert.Equal("Basic", req.Headers.Authorization!.Scheme);
        Assert.Equal("a@x:pw", Decode(req));
    }

    [Theory]
    [InlineData(HttpStatusCode.Unauthorized)]
    [InlineData(HttpStatusCode.Forbidden)]
    public async Task SignIn_401_throws_exact_message(HttpStatusCode status)
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("{}", status));
        using var svc = new TogglSetupService(handler);

        var ex = await Assert.ThrowsAsync<TogglSetupException>(() => svc.SignInAsync("a@x", "pw"));
        Assert.Equal(WrongPassword, ex.Message);
    }

    [Fact]
    public async Task ValidateToken_uses_api_token_basic_auth()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("""{"api_token":"t1","fullname":"Ann","email":"a@x"}"""));
        using var svc = new TogglSetupService(handler);

        var account = await svc.ValidateTokenAsync("t1");

        Assert.Equal("Ann", account.FullName);
        Assert.Equal("t1:api_token", Decode(Assert.Single(handler.Requests)));
    }

    [Fact]
    public async Task ValidateToken_401_throws_token_message()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("{}", HttpStatusCode.Unauthorized));
        using var svc = new TogglSetupService(handler);

        var ex = await Assert.ThrowsAsync<TogglSetupException>(() => svc.ValidateTokenAsync("bad"));
        Assert.Equal("That API token was not accepted by Toggl.", ex.Message);
    }

    [Fact]
    public async Task Organizations_are_mapped()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("""[{"id":1,"name":"Org A"},{"id":2,"name":"Org B"}]"""));
        using var svc = new TogglSetupService(handler);

        var orgs = await svc.GetOrganizationsAsync("t1");

        Assert.Equal([new TogglOrganization(1, "Org A"), new TogglOrganization(2, "Org B")], orgs);
        Assert.Equal("https://api.track.toggl.com/api/v9/me/organizations", handler.Requests[0].RequestUri!.ToString());
    }

    [Fact]
    public async Task Workspaces_filtered_by_organization()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json(
            """[{"id":10,"name":"W1","organization_id":1},{"id":11,"name":"W2","organization_id":2},{"id":12,"name":"W3","organization_id":1}]"""));
        using var svc = new TogglSetupService(handler);

        var ws = await svc.GetWorkspacesAsync("t1", 1);

        Assert.Equal([new TogglWorkspace(10, "W1", 1), new TogglWorkspace(12, "W3", 1)], ws);
        Assert.Equal("https://api.track.toggl.com/api/v9/me/workspaces", handler.Requests[0].RequestUri!.ToString());
    }

    [Fact]
    public async Task Projects_paginate_until_short_page()
    {
        static string Page(int from, int count) =>
            "[" + string.Join(",", Enumerable.Range(from, count).Select(i => $$"""{"id":{{i}},"name":"P{{i:D4}}"}""")) + "]";
        var handler = new FakeHttpHandler(r => r.RequestUri!.Query.Contains("page=1")
            ? FakeHttpHandler.Json(Page(1000, 200))
            : FakeHttpHandler.Json(Page(1, 3)));
        using var svc = new TogglSetupService(handler);

        var projects = await svc.GetActiveProjectsAsync("t1", 77);

        Assert.Equal(203, projects.Count);
        Assert.Equal(2, handler.Requests.Count);
        Assert.Equal("P0001", projects[0].Name);
        Assert.Equal(projects.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).Select(p => p.Id), projects.Select(p => p.Id));
        Assert.Equal("https://api.track.toggl.com/api/v9/workspaces/77/projects?active=true&per_page=200&page=1",
            handler.Requests[0].RequestUri!.ToString());
        Assert.EndsWith("page=2", handler.Requests[1].RequestUri!.ToString());
    }

    [Fact]
    public async Task Projects_stop_when_toggl_ignores_page()
    {
        var page = "[" + string.Join(",", Enumerable.Range(1, 200).Select(i => $$"""{"id":{{i}},"name":"P{{i}}"}""")) + "]";
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json(page));
        using var svc = new TogglSetupService(handler);

        var projects = await svc.GetActiveProjectsAsync("t1", 77);

        Assert.Equal(2, handler.Requests.Count);
        Assert.Equal(200, projects.Count);
    }

    [Fact]
    public async Task Projects_pagination_is_capped_at_50_pages()
    {
        var handler = new FakeHttpHandler(r =>
        {
            var n = int.Parse(System.Web.HttpUtility.ParseQueryString(r.RequestUri!.Query)["page"]!);
            return FakeHttpHandler.Json("[" + string.Join(",", Enumerable.Range(n * 1000, 200)
                .Select(i => $$"""{"id":{{i}},"name":"P{{i}}"}""")) + "]");
        });
        using var svc = new TogglSetupService(handler);

        var projects = await svc.GetActiveProjectsAsync("t1", 77);

        Assert.Equal(50, handler.Requests.Count);
        Assert.Equal(50 * 200, projects.Count);
    }

    [Fact]
    public async Task Echoed_api_token_not_in_exception_messages()
    {
        var handler = new FakeHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.InternalServerError)
        {
            ReasonPhrase = "Server Error",
            Content = new StringContent("bad request for token s3cr3t-tok")
        });
        using var svc = new TogglSetupService(handler);

        var ex = await Assert.ThrowsAsync<TogglSetupException>(() => svc.GetOrganizationsAsync("s3cr3t-tok"));

        Assert.DoesNotContain("s3cr3t-tok", ex.Message);
        Assert.Equal("Toggl returned 500: Server Error", ex.Message);
    }

    [Fact]
    public async Task Body_mentioning_api_token_word_is_shown_for_token_calls()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("api_token quota exceeded", HttpStatusCode.TooManyRequests));
        using var svc = new TogglSetupService(handler);

        var ex = await Assert.ThrowsAsync<TogglSetupException>(() => svc.GetOrganizationsAsync("t1"));

        Assert.Equal("Toggl returned 429: api_token quota exceeded", ex.Message);
    }

    [Fact]
    public async Task Password_not_in_exception_messages()
    {
        var handler = new FakeHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.InternalServerError)
        {
            ReasonPhrase = "Server Error",
            Content = new StringContent("boom for user a@x password hunter2")
        });
        using var svc = new TogglSetupService(handler);

        var ex = await Assert.ThrowsAsync<TogglSetupException>(() => svc.SignInAsync("a@x", "hunter2"));

        Assert.DoesNotContain("hunter2", ex.Message);
        Assert.Contains("500", ex.Message);
        Assert.Contains("Server Error", ex.Message);
    }

    [Fact]
    public async Task Other_errors_include_body_when_safe()
    {
        var handler = new FakeHttpHandler(_ => FakeHttpHandler.Json("rate limited", HttpStatusCode.TooManyRequests));
        using var svc = new TogglSetupService(handler);

        var ex = await Assert.ThrowsAsync<TogglSetupException>(() => svc.GetOrganizationsAsync("t1"));
        Assert.Equal("Toggl returned 429: rate limited", ex.Message);
    }
}
