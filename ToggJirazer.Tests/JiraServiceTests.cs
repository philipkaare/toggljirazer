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
    private const string ExecuteRead = "executeRead";
    private const string FieldMetaOperation = "getJiraIssueTypeMetaWithFields";

    private static readonly JiraConfig Config = new() { CloudId = "cloud-1", BudgetFieldName = "Budget", AccountFieldName = "Account" };

    // ---- Fixtures in the shapes the Atlassian MCP server (v2) returns, after its "data" envelope is unwrapped ----

    private static string TypeId(string type) => "type-" + type;

    private static string TypeName(string typeId) => typeId["type-".Length..];

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

        // Custom fields come keyed by label with their ID; an empty one has no "value"; null system fields are left out.
        var customFields = new JsonObject
        {
            ["Budget"] = budgetOption is null
                ? new JsonObject { ["id"] = BudgetId }
                : new JsonObject { ["id"] = BudgetId, ["value"] = new JsonObject { ["value"] = budgetOption, ["id"] = "10177" } },
            ["Account"] = account is null
                ? new JsonObject { ["id"] = AccountId }
                : new JsonObject { ["id"] = AccountId, ["value"] = new JsonArray(account) },
        };
        var fields = new JsonObject
        {
            ["summary"] = $"Summary of {key}",
            ["issuetype"] = new JsonObject { ["id"] = TypeId(type), ["name"] = type, ["subtask"] = false, ["hierarchyLevel"] = 0 },
            ["fixVersions"] = versions,
            ["customFields"] = customFields,
        };
        if (estimateSeconds is { } s)
            fields["timeoriginalestimate"] = s;
        if (parentKey != null)
            fields["parent"] = new JsonObject { ["id"] = "99", ["key"] = parentKey, ["fields"] = new JsonObject { ["summary"] = "Parent" } };

        return new JsonObject { ["id"] = "10" + key.GetHashCode() % 1000, ["key"] = key, ["fields"] = fields };
    }

    private static JsonElement Page(IEnumerable<JsonObject> nodes, string? nextCursor = null)
    {
        var array = new JsonArray();
        foreach (var node in nodes) array.Add(node.DeepClone());
        var root = new JsonObject { ["issues"] = array, ["isLast"] = nextCursor == null, ["appliedContentFormat"] = "markdown" };
        if (nextCursor != null) root["nextPageToken"] = nextCursor;
        return JsonSerializer.SerializeToElement(root);
    }

    private static JsonElement Single(JsonObject issue) => JsonSerializer.SerializeToElement(issue);

    /// <summary>The <c>getJiraIssueTypeMetaWithFields</c> result: every field of a project's issue type.</summary>
    private static JsonElement FieldMeta(Dictionary<string, string> names)
    {
        var fields = new JsonArray();
        foreach (var kv in names)
            fields.Add(new JsonObject { ["fieldId"] = kv.Key, ["name"] = kv.Value, ["required"] = false, ["schema"] = new JsonObject { ["type"] = "string" } });
        return JsonSerializer.SerializeToElement(new JsonObject { ["total"] = names.Count, ["startAt"] = 0, ["maxResults"] = 200, ["fields"] = fields });
    }

    private static (string Project, string IssueTypeId) MetaTarget(IReadOnlyDictionary<string, object?> args)
    {
        Assert.Equal(FieldMetaOperation, args["name"]);
        var inputs = (IReadOnlyDictionary<string, object?>)args["inputs"]!;
        Assert.Equal("cloud-1", inputs["cloudId"]);
        Assert.Equal(false, inputs["requiredFieldsOnly"]);
        return ((string)inputs["projectIdOrKey"]!, (string)inputs["issueTypeId"]!);
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

    /// <summary>
    /// A fake Jira backed by a set of issues: answers getJiraIssue, the field metadata of a project's issue type
    /// (<paramref name="namesFor"/>(project, issue type), all names by default), and key/fixVersion JQL in one page.
    /// </summary>
    private static FakeJiraTools Backed(IEnumerable<JsonObject> issues, Func<string, string, Dictionary<string, string>>? namesFor = null)
    {
        var store = issues.ToDictionary(i => (string)i["key"]!, StringComparer.OrdinalIgnoreCase);
        return new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue)
            {
                var key = (string)args["issueIdOrKey"]!;
                if (!store.TryGetValue(key, out var issue))
                    throw new AtlassianToolException($"Issue \"{key}\" not found", 404);
                return Single(issue);
            }

            if (tool == ExecuteRead)
            {
                var (project, typeId) = MetaTarget(args);
                return FieldMeta(namesFor?.Invoke(project, TypeName(typeId)) ?? AllNames);
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
    public async Task Field_ids_resolve_from_issue_type_metadata()
    {
        var tools = Backed([Issue("ABC-1", budgetOption: "Fondsdrift", account: "Kunde A")]);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1"]);

        var getCall = Assert.Single(tools.CallsTo(GetIssue)).Args;
        Assert.Equal("cloud-1", getCall["cloudId"]);
        Assert.Equal("ABC-1", getCall["issueIdOrKey"]);
        Assert.Equal(["issuetype"], Fields(getCall));
        Assert.Equal("full", getCall["view"]);

        Assert.Equal(("ABC", TypeId("Task")), MetaTarget(Assert.Single(tools.CallsTo(ExecuteRead)).Args));

        var searchCall = Assert.Single(tools.CallsTo(Search)).Args;
        Assert.Contains(BudgetId, Fields(searchCall));
        Assert.Contains(AccountId, Fields(searchCall));

        Assert.Equal("Fondsdrift", result["ABC-1"]!.Budget);
        Assert.Equal("Kunde A", result["ABC-1"]!.Account);
    }

    [Fact]
    public async Task Custom_field_values_are_read_by_id_whatever_their_label()
    {
        // The site shows the Account field as "Konto": the value is still found by its ID.
        var issue = Issue("ABC-1", account: "Kunde A");
        var custom = (JsonObject)issue["fields"]!["customFields"]!;
        var account = custom["Account"]!.DeepClone();
        custom.Remove("Account");
        custom["Konto"] = account;
        var tools = Backed([issue]);

        var result = await new JiraService(Config, tools).GetIssuesBulkAsync(["ABC-1"]);

        Assert.Equal("Kunde A", result["ABC-1"]!.Account);
    }

    [Fact]
    public async Task Field_ids_are_read_once_per_project_and_issue_type()
    {
        var tools = Backed([Issue("ABC-1"), Issue("ABC-2")], (_, _) => new() { [AccountId] = "Account" });
        var service = new JiraService(Config, tools);

        await service.GetIssuesBulkAsync(["ABC-1", "ABC-2"]);

        Assert.Equal(2, tools.CallsTo(GetIssue).Count());
        Assert.Single(tools.CallsTo(ExecuteRead));
    }

    [Fact]
    public async Task Field_ids_try_next_project_when_missing()
    {
        var tools = Backed(
            [Issue("ABC-1"), Issue("ABC-2"), Issue("XYZ-1", budgetOption: "Drift")],
            (project, _) => project == "ABC" ? new() { [AccountId] = "Account" } : AllNames);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1", "ABC-2", "XYZ-1"]);

        // The first key of each project is tried before a second key of any project.
        var getCalls = tools.CallsTo(GetIssue).Select(c => c.Args["issueIdOrKey"]).ToList();
        Assert.Equal(["ABC-1", "XYZ-1"], getCalls);
        Assert.Equal(2, tools.CallsTo(ExecuteRead).Count());
        Assert.Equal("Drift", result["XYZ-1"]!.Budget);
    }

    [Fact]
    public async Task Field_ids_probe_other_projects_before_more_keys_of_a_big_project()
    {
        // Six ABC keys whose project lacks Budget must not use up the five probes before XYZ is tried.
        var abcKeys = Enumerable.Range(1, 6).Select(n => $"ABC-{n}").ToList();
        var tools = Backed(
            abcKeys.Select(k => Issue(k)).Append(Issue("XYZ-1", budgetOption: "Drift")),
            (project, _) => project == "ABC" ? new() { [AccountId] = "Account" } : AllNames);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync([.. abcKeys, "XYZ-1"]);

        var getCalls = tools.CallsTo(GetIssue).Select(c => c.Args["issueIdOrKey"]).ToList();
        Assert.Equal(["ABC-1", "XYZ-1"], getCalls);
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

    [Fact]
    public async Task Field_id_probes_stop_at_five_calls()
    {
        var keys = Enumerable.Range(1, 8).Select(n => $"ABC-{n}").ToList();
        var tools = Backed(keys.Select(k => Issue(k)), (_, _) => new() { [AccountId] = "Account" });
        var service = new JiraService(Config, tools);

        await service.GetIssuesBulkAsync(keys);

        Assert.Equal(5, tools.CallsTo(GetIssue).Count());
    }

    [Theory]
    [InlineData(500, "Internal server error")]
    [InlineData(null, "Atlassian returned a result that is not JSON: <html>")]
    public async Task Field_id_probe_non_key_tool_error_propagates(int? status, string message)
    {
        var tools = new FakeJiraTools((tool, _) => tool == GetIssue
            ? throw new AtlassianToolException(message, status)
            : Page([Issue("ABC-1")]));
        var service = new JiraService(Config, tools);

        var ex = await Assert.ThrowsAsync<AtlassianToolException>(() => service.GetIssuesBulkAsync(["ABC-1", "ABC-2"]));

        Assert.Equal(message, ex.Message);
        Assert.Single(tools.CallsTo(GetIssue));
        Assert.Empty(tools.CallsTo(Search));
    }

    [Fact]
    public async Task Field_ids_keep_probing_the_same_project_when_a_field_is_still_missing()
    {
        // ABC-1's issue type (Task) has no Budget field on its screen; ABC-2's (Epic) has.
        var tools = Backed(
            [Issue("ABC-1"), Issue("ABC-2", type: "Epic", budgetOption: "Fondsdrift")],
            (_, type) => type == "Task" ? new() { [AccountId] = "Account" } : AllNames);
        var service = new JiraService(Config, tools);

        var result = await service.GetIssuesBulkAsync(["ABC-1", "ABC-2"]);

        var getCalls = tools.CallsTo(GetIssue).Select(c => c.Args["issueIdOrKey"]).ToList();
        Assert.Equal(["ABC-1", "ABC-2"], getCalls);
        Assert.True(getCalls.Count <= 5);
        Assert.Equal("Fondsdrift", result["ABC-2"]!.Budget);
    }

    [Theory]
    [InlineData(500, "Internal server error")]
    [InlineData(null, "Atlassian returned a result that is not JSON: <html>")]
    [InlineData(null, "You do not have permission to use this tool.")]
    public async Task Bulk_non_key_tool_error_propagates_without_retries(int? status, string message)
    {
        var tools = new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue) return Single(Issue("ABC-1"));
            if (tool == ExecuteRead) return FieldMeta(AllNames);
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
            if (tool == GetIssue) return Single(Issue("ABC-1"));
            if (tool == ExecuteRead) return FieldMeta(AllNames);
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
            if (tool == GetIssue) return Single(store["ABC-1"]);
            if (tool == ExecuteRead) return FieldMeta(AllNames);
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
    public async Task Bulk_unparseable_rejection_of_every_search_throws_instead_of_all_null()
    {
        var keys = Enumerable.Range(1, 10).Select(n => $"ABC-{n}").ToList();
        var tools = new FakeJiraTools((tool, args) =>
        {
            if (tool == GetIssue) return Single(Issue("ABC-1"));
            if (tool == ExecuteRead) return FieldMeta(AllNames);
            throw new AtlassianToolException("Something went wrong.", 400);
        });
        var service = new JiraService(Config, tools);

        var ex = await Assert.ThrowsAsync<AtlassianToolException>(() => service.GetIssuesBulkAsync(keys));

        Assert.Equal(400, ex.StatusCode);
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
            if (tool == GetIssue) return Single(Issue("ABC-1"));
            if (tool == ExecuteRead) return FieldMeta(AllNames);
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

    /// <summary>Recent issues for the field listing; the field metadata of each issue's project and type is its names.</summary>
    private static FakeJiraTools FieldListing(params (string Key, string Type, Dictionary<string, string> Names)[] recent) =>
        new((tool, args) =>
        {
            if (tool == ExecuteRead)
            {
                var (project, typeId) = MetaTarget(args);
                return FieldMeta(recent.First(r => r.Key.StartsWith(project + "-") && r.Type == TypeName(typeId)).Names);
            }
            return Page(recent.Select(r => Issue(r.Key, type: r.Type)));
        });

    private static IEnumerable<(string Project, string IssueTypeId)> MetaTargets(FakeJiraTools tools) =>
        tools.CallsTo(ExecuteRead).Select(c => MetaTarget(c.Args));

    [Fact]
    public async Task ListFieldNames_merges_names_across_projects()
    {
        var tools = FieldListing(
            ("ABC-1", "Task", new() { ["summary"] = "Summary", [BudgetId] = "Budget" }),
            ("ABC-2", "Task", new() { ["summary"] = "Summary" }),
            ("XYZ-1", "Task", new() { ["summary"] = "Summary", [AccountId] = "account" }));
        var service = new JiraService(Config, tools);

        var fields = await service.ListFieldNamesAsync();

        Assert.Equal(
            [(AccountId, "account"), (BudgetId, "Budget"), ("summary", "Summary")],
            fields.ToArray());
        Assert.Equal([("ABC", TypeId("Task")), ("XYZ", TypeId("Task"))], MetaTargets(tools));
        Assert.Empty(tools.CallsTo(GetIssue));
    }

    [Fact]
    public async Task ListFieldNames_merges_names_across_issue_types()
    {
        // Budget is only on the Epic screen.
        var tools = FieldListing(
            ("ABC-1", "Task", new() { ["summary"] = "Summary" }),
            ("ABC-2", "Epic", new() { ["summary"] = "Summary", [BudgetId] = "Budget" }));

        var fields = await new JiraService(Config, tools).ListFieldNamesAsync();

        Assert.Equal([(BudgetId, "Budget"), ("summary", "Summary")], fields.ToArray());
        Assert.Equal([("ABC", TypeId("Task")), ("ABC", TypeId("Epic"))], MetaTargets(tools));
    }

    [Fact]
    public async Task ListFieldNames_jql_is_bounded()
    {
        var tools = FieldListing(("ABC-1", "Task", AllNames));

        await new JiraService(Config, tools).ListFieldNamesAsync();

        var search = Assert.Single(tools.CallsTo(Search)).Args;
        Assert.Contains("updated >=", Jql(search));
        Assert.Equal(50, search["maxResults"]);
        Assert.Equal(["issuetype"], Fields(search));
        Assert.Equal("full", search["view"]);
    }

    [Fact]
    public async Task ListFieldNames_reads_at_most_five_issue_types_trying_every_project_first()
    {
        var tools = FieldListing(
            ("AAA-1", "Task", new() { ["summary"] = "Summary" }),
            ("AAA-2", "Epic", new() { ["summary"] = "Summary" }),
            ("BBB-1", "Task", new() { ["summary"] = "Summary" }),
            ("CCC-1", "Task", new() { ["summary"] = "Summary" }),
            ("DDD-1", "Task", new() { ["summary"] = "Summary" }),
            ("EEE-1", "Task", new() { ["summary"] = "Summary" }),
            ("FFF-1", "Task", new() { [BudgetId] = "Budget" }));

        var fields = await new JiraService(Config, tools).ListFieldNamesAsync();

        Assert.Equal(["AAA", "BBB", "CCC", "DDD", "EEE"], MetaTargets(tools).Select(t => t.Project));
        Assert.Equal([("summary", "Summary")], fields.ToArray());
    }

    [Fact]
    public async Task ListFieldNames_empty_when_no_recent_issues()
    {
        var tools = FieldListing();

        var fields = await new JiraService(Config, tools).ListFieldNamesAsync();

        Assert.Empty(fields);
        Assert.Empty(tools.CallsTo(ExecuteRead));
    }

    [Fact]
    public async Task Empty_field_name_skips_field_without_warning()
    {
        var config = new JiraConfig { CloudId = "cloud-1", BudgetFieldName = "", AccountFieldName = "Nonexistent" };
        var tools = Backed([Issue("ABC-1", budgetOption: "Fondsdrift", account: "Kunde A")]);
        var original = Console.Out;
        var captured = new StringWriter();
        Console.SetOut(captured);
        try
        {
            var result = await new JiraService(config, tools).GetIssuesBulkAsync(["ABC-1"]);

            Assert.Null(result["ABC-1"]!.Budget);
            Assert.DoesNotContain(BudgetId, Fields(Assert.Single(tools.CallsTo(Search)).Args));
        }
        finally
        {
            Console.SetOut(original);
        }

        var output = captured.ToString();
        Assert.DoesNotContain("Jira field ''", output);
        Assert.Contains("Jira field 'Nonexistent' not found", output);
    }

    [Fact]
    public async Task Both_field_names_empty_make_no_field_lookup()
    {
        var config = new JiraConfig { CloudId = "cloud-1", BudgetFieldName = "", AccountFieldName = "" };
        var tools = Backed([Issue("ABC-1")]);

        await new JiraService(config, tools).GetIssuesBulkAsync(["ABC-1"]);

        Assert.Empty(tools.CallsTo(GetIssue));
        Assert.Empty(tools.CallsTo(ExecuteRead));
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
            Assert.Equal("full", s.Args["view"]);
            Assert.False(s.Args.ContainsKey("responseContentFormat"));
        });
    }
}
