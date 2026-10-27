using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Services.Atlassian;
using ToggJirazer.Tests.Fakes;
using Xunit;

namespace ToggJirazer.Tests;

public class JiraServiceTests
{
    private const string BudgetId = "customfield_10073";
    private const string AccountId = "customfield_10111";
    private const string Search = "searchJiraIssuesUsingJql";
    private const string GetIssue = "getJiraIssue";

    private static readonly JiraConfig Config = new() { CloudId = "cloud-1", BudgetFieldName = "Budget", AccountFieldName = "Account" };

    // ---- Fixtures in the shapes the Atlassian MCP server returns ----

    private static JsonObject Issue(
        string key,
        string type = "Task",
        string? parentKey = null,
        (string Name, bool Released)[]? fixVersions = null,
        long? estimateSeconds = null,
        string? budgetOption = null,
        string? account = null)
    {
        var versions = new JsonArray();
        foreach (var (name, released) in fixVersions ?? [])
        {
            versions.Add(new JsonObject
            {
                ["self"] = $"https://example.atlassian.net/rest/api/3/version/{name}",
                ["id"] = "1" + Math.Abs(name.GetHashCode() % 1000),
                ["description"] = "",
                ["name"] = name,
                ["archived"] = false,
                ["released"] = released,
            });
        }

        var fields = new JsonObject
        {
            ["summary"] = $"Summary of {key}",
            ["issuetype"] = new JsonObject { ["name"] = type, ["subtask"] = false },
            ["fixVersions"] = versions,
            ["timeoriginalestimate"] = estimateSeconds is { } s ? JsonValue.Create(s) : null,
            [BudgetId] = budgetOption is null ? null : new JsonObject { ["self"] = "https://example.atlassian.net/rest/api/3/customFieldOption/1", ["value"] = budgetOption, ["id"] = "1" },
            [AccountId] = account is null ? null : new JsonObject { ["self"] = "https://example.atlassian.net/rest/api/3/customFieldOption/2", ["value"] = account, ["id"] = "2" },
        };
        if (parentKey != null)
            fields["parent"] = new JsonObject { ["id"] = "99", ["key"] = parentKey };

        return new JsonObject { ["id"] = "10" + key.GetHashCode() % 1000, ["key"] = key, ["fields"] = fields };
    }

    private static JsonElement Page(IEnumerable<JsonObject> nodes, string? nextCursor = null)
    {
        var array = new JsonArray();
        foreach (var node in nodes) array.Add(node.DeepClone());
        var root = new JsonObject
        {
            ["issues"] = new JsonObject
            {
                ["nodes"] = array,
                ["pageInfo"] = new JsonObject { ["hasNextPage"] = nextCursor != null, ["endCursor"] = nextCursor },
                ["webUrl"] = "https://example.atlassian.net/issues",
            },
            ["context"] = new JsonObject(),
        };
        return JsonSerializer.SerializeToElement(root);
    }

    private static JsonElement NamesResponse(JsonObject issue, Dictionary<string, string> names)
    {
        var node = (JsonObject)issue.DeepClone();
        var namesObj = new JsonObject();
        foreach (var kv in names) namesObj[kv.Key] = kv.Value;
        node["names"] = namesObj;
        return Page([node]);
    }

    private static readonly Dictionary<string, string> AllNames = new()
    {
        ["summary"] = "Summary",
        [BudgetId] = "Budget",
        [AccountId] = "Account",
    };

    private static List<string> KeysIn(string jql)
    {
        var m = Regex.Match(jql, @"key in \(([^)]*)\)");
        return m.Success ? m.Groups[1].Value.Split(',').Select(k => k.Trim().Trim('"')).ToList() : [];
    }

    private static List<string> VersionsIn(string jql)
    {
        var m = Regex.Match(jql, @"fixVersion in \(([^)]*)\)");
        return m.Success ? m.Groups[1].Value.Split(',').Select(v => v.Trim().Trim('"')).ToList() : [];
    }

    private static string Jql(IReadOnlyDictionary<string, object?> args) => (string)args["jql"]!;

    private static IEnumerable<string> Fields(IReadOnlyDictionary<string, object?> args) => (IEnumerable<string>)args["fields"]!;

    /// <summary>A fake Jira backed by a set of issues: answers getJiraIssue with names, and key/fixVersion JQL in one page.</summary>
    private static FakeJiraTools Backed(IEnumerable<JsonObject> issues, Func<string, Dictionary<string, string>>? namesForKey = null)
    {
        var store = issues.ToDictionary(i => (string)i["key"]!, StringComparer.OrdinalIgnoreCase);
        return new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue)
            {
                var key = (string)args["issueIdOrKey"]!;
                if (!store.TryGetValue(key, out var issue))
                    throw new AtlassianToolException($"Issue does not exist or you do not have permission to see it.", 404);
                return NamesResponse(issue, namesForKey?.Invoke(key) ?? AllNames);
            }

            var jql = Jql(args);
            var keys = KeysIn(jql);
            if (keys.Count > 0)
            {
                var missing = keys.Where(k => !store.ContainsKey(k)).ToList();
                if (missing.Count > 0)
                    throw new AtlassianToolException($"An issue with key '{missing[0]}' does not exist for field 'key'.", 400);
                return Page(keys.Select(k => store[k]));
            }

            var versions = VersionsIn(jql);
            var matches = store.Values.Where(i => ((JsonArray)i["fields"]!["fixVersions"]!)
                .Any(v => versions.Contains((string)v!["name"]!, StringComparer.OrdinalIgnoreCase)));
            if (jql.Contains("issuetype = Delivery"))
                matches = matches.Where(i => (string)i["fields"]!["issuetype"]!["name"]! == "Delivery");
            return Page(matches);
        });
    }

    // ---- Tests ----

    [Fact]
    public async Task Field_ids_resolve_from_names_map()
    {
        var tools = Backed([Issue("ABC-1", budgetOption: "Fondsdrift", account: "Kunde A")]);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1"]);

        var getCall = Assert.Single(tools.CallsTo(GetIssue)).Args;
        Assert.Equal("cloud-1", getCall["cloudId"]);
        Assert.Equal("ABC-1", getCall["issueIdOrKey"]);
        Assert.Equal(["*all"], Fields(getCall));
        Assert.Equal("names", getCall["expand"]);
        Assert.Equal("adf", getCall["responseContentFormat"]);

        var searchCall = Assert.Single(tools.CallsTo(Search)).Args;
        Assert.Contains(BudgetId, Fields(searchCall));
        Assert.Contains(AccountId, Fields(searchCall));

        Assert.Equal("Fondsdrift", result["ABC-1"]!.Budget);
        Assert.Equal("Kunde A", result["ABC-1"]!.Account);
    }

    [Fact]
    public async Task Field_ids_try_next_project_when_missing()
    {
        var tools = Backed(
            [Issue("ABC-1"), Issue("ABC-2"), Issue("XYZ-1", budgetOption: "Drift")],
            key => key.StartsWith("ABC") ? new() { [AccountId] = "Account" } : AllNames);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1", "ABC-2", "XYZ-1"]);

        var getCalls = tools.CallsTo(GetIssue).ToList();
        Assert.Equal(2, getCalls.Count);
        Assert.Equal("ABC-1", getCalls[0].Args["issueIdOrKey"]);
        Assert.Equal("XYZ-1", getCalls[1].Args["issueIdOrKey"]);
        Assert.True(tools.Calls.Count <= 5, $"Expected at most 5 calls, got {tools.Calls.Count}");
        Assert.Equal("Drift", result["XYZ-1"]!.Budget);
    }

    [Fact]
    public async Task Field_ids_try_next_key_of_same_project_when_probe_fails()
    {
        // ABC-1 (e.g. a mistyped key in a Toggl description) doesn't exist; ABC-2 does.
        var tools = Backed([Issue("ABC-2", budgetOption: "Fondsdrift"), Issue("ABC-3")]);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1", "ABC-2", "ABC-3"]);

        var getCalls = tools.CallsTo(GetIssue).ToList();
        Assert.Equal(2, getCalls.Count);
        Assert.Equal("ABC-1", getCalls[0].Args["issueIdOrKey"]);
        Assert.Equal("ABC-2", getCalls[1].Args["issueIdOrKey"]);
        Assert.Contains(BudgetId, Fields(tools.CallsTo(Search).First().Args));
        Assert.Equal("Fondsdrift", result["ABC-2"]!.Budget);
    }

    [Theory]
    [InlineData(500, "Internal server error")]
    [InlineData(null, "Atlassian returned a result that is not JSON: <html>")]
    public async Task Bulk_non_key_tool_error_propagates_without_retries(int? status, string message)
    {
        var tools = new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue) return NamesResponse(Issue("ABC-1"), AllNames);
            throw new AtlassianToolException(message, status);
        });
        var service = new JiraService(Config, tools);

        var keys = Enumerable.Range(1, 100).Select(n => $"ABC-{n}").ToList();
        var ex = await Assert.ThrowsAsync<AtlassianToolException>(() => service.GetIssuesBulkAsync(keys));

        Assert.Equal(status, ex.StatusCode);
        Assert.Single(tools.CallsTo(Search));
    }

    [Fact]
    public async Task Bulk_rate_limit_from_search_propagates()
    {
        var tools = new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue) return NamesResponse(Issue("ABC-1"), AllNames);
            throw new AtlassianRateLimitException();
        });
        var service = new JiraService(Config, tools);

        await Assert.ThrowsAsync<AtlassianRateLimitException>(() => service.GetIssuesBulkAsync(["ABC-1", "ABC-2"]));
        Assert.Single(tools.CallsTo(Search));
    }

    [Fact]
    public async Task Bulk_uses_key_in_jql_chunks_of_100()
    {
        var issues = Enumerable.Range(1, 150).Select(n => Issue($"ABC-{n}")).ToList();
        var tools = Backed(issues);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(issues.Select(i => (string)i["key"]!));

        var searches = tools.CallsTo(Search).ToList();
        Assert.Equal(2, searches.Count);
        Assert.All(searches, s => Assert.StartsWith("key in (", Jql(s.Args)));
        Assert.Equal(100, KeysIn(Jql(searches[0].Args)).Count);
        Assert.Equal(50, KeysIn(Jql(searches[1].Args)).Count);
        Assert.Equal(150, result.Values.Count(v => v != null));
    }

    [Fact]
    public async Task Bulk_retries_chunk_without_missing_keys()
    {
        var keys = Enumerable.Range(1, 10).Select(n => $"ABC-{n}").ToList();
        var tools = Backed(keys.Where(k => k != "ABC-9").Select(k => Issue(k)));
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(keys);

        var searches = tools.CallsTo(Search).ToList();
        Assert.Equal(2, searches.Count);
        Assert.Contains("ABC-9", KeysIn(Jql(searches[0].Args)));
        Assert.DoesNotContain("ABC-9", KeysIn(Jql(searches[1].Args)));
        Assert.True(result.ContainsKey("ABC-9"));
        Assert.Null(result["ABC-9"]);
        Assert.All(keys.Where(k => k != "ABC-9"), k => Assert.NotNull(result[k]));
    }

    [Fact]
    public async Task Bulk_halving_terminates_when_error_text_unparseable()
    {
        var keys = Enumerable.Range(1, 10).Select(n => $"ABC-{n}").ToList();
        var store = keys.ToDictionary(k => k, k => Issue(k));
        var tools = new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue) return NamesResponse(store["ABC-1"], AllNames);
            var requested = KeysIn(Jql(args));
            if (requested.Contains("ABC-3")) throw new AtlassianToolException("Something went wrong.", 400);
            return Page(requested.Select(k => store[k]));
        });
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(keys);

        // 10 -> 5 + 5 -> 2 + 3 -> 1 + 2: seven searches; anything near one per key would be a regression.
        Assert.True(tools.CallsTo(Search).Count() <= 9, $"Expected at most 9 searches, got {tools.CallsTo(Search).Count()}");
        Assert.Null(result["ABC-3"]);
        Assert.All(keys.Where(k => k != "ABC-3"), k => Assert.NotNull(result[k]));
    }

    [Fact]
    public async Task Parent_chain_inherits_fix_version()
    {
        var tools = Backed(
        [
            Issue("ABC-1", parentKey: "ABC-2"),
            Issue("ABC-2", type: "Epic", fixVersions: [("v1", false)]),
        ]);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1"]);

        var searches = tools.CallsTo(Search).ToList();
        Assert.Equal(2, searches.Count);
        Assert.Equal(["ABC-2"], KeysIn(Jql(searches[1].Args)));
        Assert.Equal(["v1"], result["ABC-1"]!.FixVersions);
    }

    [Fact]
    public async Task Batched_version_grouping_equals_per_version()
    {
        var abc1 = Issue("ABC-1", fixVersions: [("v1", false), ("v2", true)], estimateSeconds: 3600);
        var abc2 = Issue("ABC-2", fixVersions: [("v2", true)], estimateSeconds: 26640);
        var tools = new FakeJiraTools((tool, args) =>
        {
            var token = args.TryGetValue("nextPageToken", out var t) ? t as string : null;
            return token == null ? Page([abc1], nextCursor: "cursor-1") : Page([abc2]);
        });
        var service = new JiraService(Config, tools);

        var grouped = await service.GetIssuesByFixVersionsAsync(["v1", "v2", "v3"]);

        var searches = tools.CallsTo(Search).ToList();
        Assert.Equal(2, searches.Count);
        Assert.StartsWith("fixVersion in (", Jql(searches[0].Args));
        Assert.Equal("cursor-1", searches[1].Args["nextPageToken"]);
        Assert.Equal(["ABC-1"], grouped["v1"].Select(i => i.Key));
        Assert.Equal(["ABC-1", "ABC-2"], grouped["v2"].Select(i => i.Key));
        Assert.Empty(grouped["v3"]);
        Assert.Equal(1.0 + 26640 / 3600.0, grouped["v2"].Sum(i => i.Estimate ?? 0));

        // The single-version method gives the same issues for v2.
        var perVersionTools = Backed([abc1, abc2]);
        var perVersion = await new JiraService(Config, perVersionTools).GetIssuesByFixVersionAsync("v2");
        Assert.Equal(grouped["v2"].Select(i => i.Key).Order(), perVersion.Select(i => i.Key).Order());
    }

    [Fact]
    public async Task Auth_error_mid_run_throws()
    {
        var tools = new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue) return NamesResponse(Issue("ABC-1"), AllNames);
            throw new AtlassianAuthException("Your Atlassian sign-in is no longer valid.");
        });
        var service = new JiraService(Config, tools);

        await Assert.ThrowsAsync<AtlassianAuthException>(() => service.GetIssuesBulkAsync(["ABC-1", "ABC-2"]));
        await Assert.ThrowsAsync<AtlassianAuthException>(() =>
            service.GetFixVersionsWithRecentlyUpdatedIssuesAsync(["v1"], new DateTime(2026, 1, 1)));
    }

    [Fact]
    public async Task Rate_limit_propagates()
    {
        var tools = new FakeJiraTools((_, _) => throw new AtlassianRateLimitException());
        var service = new JiraService(Config, tools);

        await Assert.ThrowsAsync<AtlassianRateLimitException>(() => service.GetIssuesBulkAsync(["ABC-1"]));
        await Assert.ThrowsAsync<AtlassianRateLimitException>(() =>
            service.GetFixVersionsWithRecentlyUpdatedIssuesAsync(["v1"], new DateTime(2026, 1, 1)));
    }

    [Fact]
    public async Task Every_jql_is_bounded()
    {
        var tools = Backed(
        [
            Issue("ABC-1", parentKey: "ABC-2", fixVersions: [("v1", false)]),
            Issue("ABC-2", type: "Epic", parentKey: "ABC-3"),
            Issue("ABC-3", type: "Initiative", budgetOption: "Fondsdrift"),
            Issue("ABC-4", type: "Delivery", fixVersions: [("v1", false)], budgetOption: "Drift"),
        ]);
        var service = new JiraService(Config, tools);

        await service.GetIssueAsync("ABC-1");
        var bulk = await service.GetIssuesBulkAsync(["ABC-1", "ABC-4"]);
        await service.GetIssuesByFixVersionAsync("v1");
        await service.GetIssuesByFixVersionsAsync(["v1", "v2"]);
        var inherited = await service.GetInheritedParentBudgetsAsync(bulk.Values.OfType<JiraIssue>());
        var leverance = await service.GetLeveranceBudgetsByFixVersionAsync(["v1"]);
        await service.GetFixVersionsWithRecentlyUpdatedIssuesAsync(["v1"], new DateTime(2026, 1, 1));

        Assert.Equal("Fondsdrift", inherited["ABC-1"]);
        Assert.Equal("Drift", leverance["v1"]);

        var searches = tools.CallsTo(Search).ToList();
        Assert.NotEmpty(searches);
        Assert.All(searches, s =>
        {
            var jql = Jql(s.Args);
            Assert.True(jql.Contains("key in (") || jql.Contains("fixVersion") || jql.Contains("updated >="), $"Unbounded JQL: {jql}");
            Assert.Equal("cloud-1", s.Args["cloudId"]);
            Assert.Equal("adf", s.Args["responseContentFormat"]);
        });
    }
}
