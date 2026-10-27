using System.Runtime.ExceptionServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using ToggJirazer.Models;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Services;

/// <summary>
/// Reads Jira data through the Atlassian MCP server's tools. Every search is a bounded JQL query
/// (<c>key in (...)</c>, <c>fixVersion in (...)</c>), batched to keep the call count low.
/// <see cref="AtlassianAuthException"/> and <see cref="AtlassianRateLimitException"/> always propagate.
/// </summary>
public class JiraService
{
    private const string SearchTool = "searchJiraIssuesUsingJql";
    private const string GetIssueTool = "getJiraIssue";
    private const string ContentFormat = "adf";
    private const int KeyChunkSize = 100;
    private const int VersionBatchSize = 50;
    private const int PageSize = 100;
    private const int MaxFieldIdLookups = 5;

    private readonly JiraConfig _config;
    private readonly IJiraTools _tools;

    private string? _budgetFieldId;
    private string? _accountFieldId;
    private bool _fieldIdsResolved;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public JiraService(JiraConfig config, IJiraTools tools)
    {
        _config = config;
        _tools = tools;
    }

    /// <summary>
    /// Finds the Budget and Account custom field IDs by name, once per service instance. Tries the first
    /// readable key of each distinct project in <paramref name="candidateKeys"/> (at most
    /// <see cref="MaxFieldIdLookups"/> <c>getJiraIssue</c> calls in total), stopping as soon as both are found.
    /// Does nothing when there is no candidate key yet.
    /// </summary>
    private async Task ResolveFieldIdsAsync(IEnumerable<string> candidateKeys)
    {
        if (_fieldIdsResolved) return;

        var keysByProject = candidateKeys
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .GroupBy(ProjectOf, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.ToList())
            .ToList();
        if (keysByProject.Count == 0) return;

        _fieldIdsResolved = true;
        var lookups = 0;
        foreach (var projectKeys in keysByProject)
        {
            // One successful lookup per project; a key that can't be read (deleted, mistyped) moves on
            // to the next key of the same project.
            foreach (var key in projectKeys)
            {
                if (_budgetFieldId != null && _accountFieldId != null) break;
                if (lookups >= MaxFieldIdLookups) break;
                lookups++;

                JsonElement root;
                try
                {
                    root = await _tools.CallToolAsync(GetIssueTool, new Dictionary<string, object?>
                    {
                        ["cloudId"] = _config.CloudId,
                        ["issueIdOrKey"] = key,
                        ["fields"] = new[] { "*all" },
                        ["expand"] = "names",
                        ["responseContentFormat"] = ContentFormat,
                    });
                }
                catch (AtlassianToolException)
                {
                    continue;
                }

                var names = ReadFieldNames(root);
                _budgetFieldId ??= names.FirstOrDefault(n =>
                    string.Equals(n.Name, _config.BudgetFieldName, StringComparison.OrdinalIgnoreCase)).Id;
                _accountFieldId ??= names.FirstOrDefault(n =>
                    string.Equals(n.Name, _config.AccountFieldName, StringComparison.OrdinalIgnoreCase)).Id;
                break;
            }
        }

        if (_budgetFieldId == null)
            Console.WriteLine($"  Warning: Jira field '{_config.BudgetFieldName}' not found in field metadata.");
        if (_accountFieldId == null)
            Console.WriteLine($"  Warning: Jira field '{_config.AccountFieldName}' not found in field metadata.");
    }

    private static string ProjectOf(string issueKey)
    {
        var dash = issueKey.LastIndexOf('-');
        return dash > 0 ? issueKey[..dash] : issueKey;
    }

    /// <summary>Reads the field ID → display name map that <c>expand: "names"</c> adds to each issue node.</summary>
    private static List<(string Id, string Name)> ReadFieldNames(JsonElement root)
    {
        var result = new List<(string Id, string Name)>();
        foreach (var node in IssueNodes(root))
        {
            if (!node.TryGetProperty("names", out var names) || names.ValueKind != JsonValueKind.Object) continue;
            foreach (var property in names.EnumerateObject())
            {
                if (property.Value.ValueKind == JsonValueKind.String)
                    result.Add((property.Name, property.Value.GetString()!));
            }
        }
        return result;
    }

    private static IEnumerable<JsonElement> IssueNodes(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Object
            && root.TryGetProperty("issues", out var issues)
            && issues.ValueKind == JsonValueKind.Object
            && issues.TryGetProperty("nodes", out var nodes)
            && nodes.ValueKind == JsonValueKind.Array)
        {
            return nodes.EnumerateArray();
        }
        return [];
    }

    private static string? NextPageToken(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Object
            && root.TryGetProperty("issues", out var issues)
            && issues.ValueKind == JsonValueKind.Object
            && issues.TryGetProperty("pageInfo", out var pageInfo)
            && pageInfo.ValueKind == JsonValueKind.Object
            && pageInfo.TryGetProperty("hasNextPage", out var hasNext)
            && hasNext.ValueKind == JsonValueKind.True
            && pageInfo.TryGetProperty("endCursor", out var cursor)
            && cursor.ValueKind == JsonValueKind.String
            && !string.IsNullOrEmpty(cursor.GetString()))
        {
            return cursor.GetString();
        }
        return null;
    }

    /// <summary>
    /// Runs a JQL search through all its pages, handing each issue to <paramref name="onIssue"/> as it
    /// arrives, so issues from pages before a failing page are kept.
    /// </summary>
    private async Task SearchAsync(string jql, IReadOnlyList<string> fields, Action<JiraIssueResponse> onIssue)
    {
        var fieldArray = fields.ToArray();
        string? nextPageToken = null;
        do
        {
            var args = new Dictionary<string, object?>
            {
                ["cloudId"] = _config.CloudId,
                ["jql"] = jql,
                ["fields"] = fieldArray,
                ["maxResults"] = PageSize,
                ["responseContentFormat"] = ContentFormat,
            };
            if (nextPageToken != null) args["nextPageToken"] = nextPageToken;

            var root = await _tools.CallToolAsync(SearchTool, args);
            foreach (var node in IssueNodes(root))
            {
                var issue = node.Deserialize<JiraIssueResponse>(JsonOptions);
                if (issue != null) onIssue(issue);
            }

            var next = NextPageToken(root);
            nextPageToken = next == nextPageToken ? null : next; // Guard against a cursor that doesn't advance.
        }
        while (nextPageToken != null);
    }

    private static string FixVersionInJql(IEnumerable<string> versions)
    {
        var quoted = versions.Select(v => "\"" + v.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"");
        return $"fixVersion in ({string.Join(", ", quoted)})";
    }

    private List<string> IssueFields()
    {
        var fields = new List<string> { "summary", "issuetype", "fixVersions", "timeoriginalestimate", "parent" };
        if (_budgetFieldId != null) fields.Add(_budgetFieldId);
        if (_accountFieldId != null) fields.Add(_accountFieldId);
        return fields;
    }

    public async Task<JiraIssue?> GetIssueAsync(string issueKey)
    {
        await ResolveFieldIdsAsync([issueKey]);

        var jiraIssue = await FetchRawIssueAsync(issueKey);
        if (jiraIssue == null) return null;

        // Iteratively walk up the parent chain to inherit fix versions.
        // A visited set prevents infinite loops from circular parent references.
        if (jiraIssue.FixVersions.Count == 0 && !string.IsNullOrEmpty(jiraIssue.ParentKey))
        {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { issueKey };
            var current = jiraIssue;
            while (current.FixVersions.Count == 0
                   && !string.IsNullOrEmpty(current.ParentKey)
                   && visited.Add(current.ParentKey))
            {
                var parent = await FetchRawIssueAsync(current.ParentKey);
                if (parent == null) break;
                if (parent.FixVersions.Count > 0)
                {
                    Console.WriteLine($"  Issue '{jiraIssue.Key}' has no fix version; inheriting from '{current.ParentKey}': {string.Join(", ", parent.FixVersions)}");
                    jiraIssue.FixVersions = parent.FixVersions.ToList();
                    break;
                }
                current = parent;
            }
        }

        return jiraIssue;
    }

    /// <summary>
    /// Fetches a single Jira issue with <c>getJiraIssue</c> and parses it into a <see cref="JiraIssue"/>
    /// without applying any parent-inheritance logic.
    /// </summary>
    private async Task<JiraIssue?> FetchRawIssueAsync(string issueKey)
    {
        JsonElement root;
        try
        {
            root = await _tools.CallToolAsync(GetIssueTool, new Dictionary<string, object?>
            {
                ["cloudId"] = _config.CloudId,
                ["issueIdOrKey"] = issueKey,
                ["fields"] = IssueFields().ToArray(),
                ["responseContentFormat"] = ContentFormat,
            });
        }
        catch (AtlassianToolException ex) when (ex.StatusCode == 404)
        {
            Console.WriteLine($"  Warning: Jira issue '{issueKey}' not found.");
            return null;
        }

        var node = IssueNodes(root).Cast<JsonElement?>().FirstOrDefault();
        var issue = node?.Deserialize<JiraIssueResponse>(JsonOptions);
        return issue == null ? null : ToJiraIssue(issue, issueKey);
    }

    private JiraIssue ToJiraIssue(JiraIssueResponse issue, string fallbackKey)
    {
        var budgetValue = _budgetFieldId != null
            ? ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(_budgetFieldId))
            : null;
        var accountValue = _accountFieldId != null
            ? ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(_accountFieldId))
            : null;

        return new JiraIssue
        {
            Key = issue.Key ?? fallbackKey,
            IssueType = issue.Fields?.Issuetype?.Name ?? string.Empty,
            Summary = issue.Fields?.Summary ?? string.Empty,
            Budget = budgetValue,
            Account = accountValue,
            FixVersions = issue.Fields?.FixVersions?.Select(v => v.Name ?? string.Empty)
                              .Where(n => !string.IsNullOrEmpty(n)).ToList() ?? new(),
            ReleasedFixVersionNames = new HashSet<string>(
                issue.Fields?.FixVersions?.Where(v => v.Released && !string.IsNullOrEmpty(v.Name))
                    .Select(v => v.Name!) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase),
            Estimate = issue.Fields?.TimeOriginalEstimate.HasValue == true
                ? issue.Fields.TimeOriginalEstimate.Value / 3600.0
                : null,
            ParentKey = issue.Fields?.Parent?.Key
        };
    }

    private static string? ExtractStringField(object? field)
    {
        if (field is JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.GetRawText(),
                JsonValueKind.Array => element.EnumerateArray().Aggregate("", (acc, e) => acc + ExtractStringField(e) + ";").TrimEnd(';'),
                JsonValueKind.Object => element.TryGetProperty("value", out var val)
                    ? val.GetString()
                    : element.TryGetProperty("name", out var name)
                        ? name.GetString()
                        : null,
                _ => null
            };
        }
        return field?.ToString();
    }

    public async Task<Dictionary<string, JiraIssue?>> GetIssuesBulkAsync(IEnumerable<string> issueKeys)
    {
        var keyList = issueKeys.ToList();
        await ResolveFieldIdsAsync(keyList);

        var result = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        var errorStatus = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int fetched = 0;

        PrintProgressBar(fetched, keyList.Count);

        for (int i = 0; i < keyList.Count; i += KeyChunkSize)
        {
            var batch = keyList.Skip(i).Take(KeyChunkSize).ToList();
            await FetchRawIssuesBulkAsync(batch, result, errorStatus);
            fetched += batch.Count;
            PrintProgressBar(fetched, keyList.Count);
        }

        Console.WriteLine();

        foreach (var kv in errorStatus)
            Console.WriteLine($"  Warning: Jira issue '{kv.Key}' could not be fetched (status {kv.Value}).");

        // Iteratively resolve the full parent chain for fix version inheritance.
        // Uses a visited set to prevent re-fetching and guard against circular references.
        // Note: TotalEstimateSum in version reports is computed via JQL fixVersion in (..), so only
        // issues with a Jira-assigned fix version contribute to estimates; inherited children do not.
        var parentCache = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        var visitedParentKeys = new HashSet<string>(result.Keys, StringComparer.OrdinalIgnoreCase);

        var pendingParentKeys = result.Values
            .Where(i => i != null && i.FixVersions.Count == 0 && !string.IsNullOrEmpty(i.ParentKey))
            .Select(i => i!.ParentKey!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(pk => !visitedParentKeys.Contains(pk))
            .ToList();

        while (pendingParentKeys.Count > 0)
        {
            Console.WriteLine($"  Fetching {pendingParentKeys.Count} parent issue(s) to inherit fix versions.");
            var batchParents = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
            var parentErrors = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < pendingParentKeys.Count; i += KeyChunkSize)
                await FetchRawIssuesBulkAsync(pendingParentKeys.Skip(i).Take(KeyChunkSize).ToList(), batchParents, parentErrors);
            foreach (var kv in parentErrors)
                Console.WriteLine($"  Warning: Parent issue '{kv.Key}' could not be fetched (status {kv.Value}); fix version will not be inherited.");

            foreach (var kv in batchParents)
                parentCache[kv.Key] = kv.Value;
            foreach (var key in pendingParentKeys)
                visitedParentKeys.Add(key);

            // Discover the next level of parents that still need to be resolved
            pendingParentKeys = batchParents.Values
                .Where(i => i != null && i.FixVersions.Count == 0 && !string.IsNullOrEmpty(i.ParentKey))
                .Select(i => i!.ParentKey!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(pk => !visitedParentKeys.Contains(pk))
                .ToList();
        }

        // Walk the parent chain for each issue that still has no fix version and apply inheritance
        foreach (var issue in result.Values.Where(i => i != null && i.FixVersions.Count == 0 && !string.IsNullOrEmpty(i.ParentKey)))
        {
            var walked = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { issue!.Key };
            var current = issue;
            while (!string.IsNullOrEmpty(current.ParentKey) && walked.Add(current.ParentKey))
            {
                JiraIssue? parent;
                if (!result.TryGetValue(current.ParentKey, out parent))
                    parentCache.TryGetValue(current.ParentKey, out parent);
                if (parent == null) break;
                if (parent.FixVersions.Count > 0)
                {
                    Console.WriteLine($"  Issue '{issue.Key}' has no fix version; inheriting from '{current.ParentKey}': {string.Join(", ", parent.FixVersions)}");
                    issue.FixVersions = parent.FixVersions.ToList();
                    break;
                }
                current = parent;
            }
        }

        return result;
    }

    /// <summary>
    /// Fetches up to <see cref="KeyChunkSize"/> issues with one <c>key in (...)</c> search and adds the parsed
    /// results to <paramref name="target"/>. No parent-inheritance logic is applied.
    /// When Jira rejects the chunk, the keys named in the error are recorded as missing and the rest is
    /// retried; when the error names none of them, the chunk is split in half, so the recursion always ends.
    /// Keys that end up unfetched map to <c>null</c> and get a status in <paramref name="errorStatus"/>.
    /// If every search of the chunk was rejected without naming a key, the rejection is not about the keys
    /// (a report with every key missing must not be written), so the first rejection is rethrown.
    /// </summary>
    private async Task FetchRawIssuesBulkAsync(
        List<string> keys,
        Dictionary<string, JiraIssue?> target,
        Dictionary<string, int> errorStatus)
    {
        var rejections = new ChunkRejections();
        await FetchKeysAsync(keys, target, errorStatus, rejections);

        if (rejections.First is { } first && !rejections.KeyNamed && keys.All(k => target.GetValueOrDefault(k) == null))
            ExceptionDispatchInfo.Throw(first);
    }

    private async Task FetchKeysAsync(
        List<string> keys,
        Dictionary<string, JiraIssue?> target,
        Dictionary<string, int> errorStatus,
        ChunkRejections rejections)
    {
        if (keys.Count == 0) return;

        var fetched = new List<JiraIssueResponse>();
        try
        {
            await SearchAsync($"key in ({string.Join(", ", keys)})", IssueFields(), fetched.Add);
        }
        catch (AtlassianToolException ex) when (IsKeyRejection(ex))
        {
            rejections.First ??= ex;
            var status = ex.StatusCode ?? 400;
            var rejected = KeysNamedIn(ex.Message, keys);
            if (rejected.Count > 0)
            {
                rejections.KeyNamed = true;
                foreach (var key in rejected) MarkMissing(key, status);
                await FetchKeysAsync(keys.Where(k => !rejected.Contains(k)).ToList(), target, errorStatus, rejections);
            }
            else if (keys.Count == 1)
            {
                MarkMissing(keys[0], status);
            }
            else
            {
                var half = keys.Count / 2;
                await FetchKeysAsync(keys.Take(half).ToList(), target, errorStatus, rejections);
                await FetchKeysAsync(keys.Skip(half).ToList(), target, errorStatus, rejections);
            }
            return;
        }

        foreach (var issue in fetched)
        {
            if (issue.Key == null) continue;
            target[issue.Key] = ToJiraIssue(issue, issue.Key);
        }

        // A key Jira accepted but didn't return (e.g. not visible) is reported like a missing one.
        foreach (var key in keys.Where(k => !target.ContainsKey(k)))
            MarkMissing(key, 404);

        void MarkMissing(string key, int status)
        {
            target[key] = null;
            errorStatus[key] = status;
        }
    }

    /// <summary>What the key rejections of one chunk looked like: the first one, and whether any named a key.</summary>
    private sealed class ChunkRejections
    {
        public AtlassianToolException? First { get; set; }
        public bool KeyNamed { get; set; }
    }

    /// <summary>
    /// True when Jira rejected the JQL because of the keys in it (400/404, or a "does not exist" message;
    /// Atlassian reports keys the user can't see as not existing). Other tool errors (5xx, non-JSON results,
    /// permission/auth messages, protocol errors) are not key problems: they propagate and stop the run
    /// instead of being retried key by key.
    /// </summary>
    private static bool IsKeyRejection(AtlassianToolException ex) =>
        ex.StatusCode is 400 or 404
        || Regex.IsMatch(ex.Message, "does not exist", RegexOptions.IgnoreCase);

    /// <summary>The keys of <paramref name="chunk"/> that a Jira error message names in single quotes, e.g. "An issue with key 'ABC-9' does not exist".</summary>
    private static HashSet<string> KeysNamedIn(string message, IEnumerable<string> chunk)
    {
        var named = Regex.Matches(message, "'([^']+)'")
            .Select(m => m.Groups[1].Value)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return chunk.Where(named.Contains).ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    // Internal response models (the MCP issue nodes have the Jira REST shape)
    private sealed class JiraIssueResponse
    {
        public string? Key { get; set; }
        public JiraIssueFields? Fields { get; set; }
    }

    private sealed class JiraIssueFields
    {
        public string? Summary { get; set; }
        public JiraIssueType? Issuetype { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("fixVersions")]
        public List<JiraVersionRef>? FixVersions { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("timeoriginalestimate")]
        public long? TimeOriginalEstimate { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("parent")]
        public JiraParentRef? Parent { get; set; }

        [System.Text.Json.Serialization.JsonExtensionData]
        public Dictionary<string, object?>? CustomFields { get; set; }
    }

    private sealed class JiraParentRef
    {
        public string? Key { get; set; }
    }

    private sealed class JiraVersionRef
    {
        public string? Name { get; set; }
        public bool Released { get; set; }
    }

    private sealed class JiraIssueType
    {
        public string? Name { get; set; }
    }

    public async Task<List<JiraIssue>> GetIssuesByFixVersionAsync(string version)
    {
        var grouped = await GetIssuesByFixVersionsAsync([version]);
        return grouped.TryGetValue(version, out var issues) ? issues : new List<JiraIssue>();
    }

    /// <summary>
    /// Fetches the issues of many fix versions with <c>fixVersion in (...)</c> searches of up to
    /// <see cref="VersionBatchSize"/> versions, and groups each issue under every one of its own fix
    /// versions that was requested — the same as one query per version. Every requested version is a
    /// key, possibly with an empty list.
    /// </summary>
    public async Task<Dictionary<string, List<JiraIssue>>> GetIssuesByFixVersionsAsync(IEnumerable<string> versions)
    {
        var result = new Dictionary<string, List<JiraIssue>>(StringComparer.OrdinalIgnoreCase);
        foreach (var version in versions)
        {
            if (version != null) result.TryAdd(version, new List<JiraIssue>());
        }

        var queryable = result.Keys.Where(v => !string.IsNullOrWhiteSpace(v)).ToList();
        var added = new HashSet<(string Version, string Key)>();
        var fields = new List<string> { "summary", "issuetype", "fixVersions", "timeoriginalestimate" };

        for (int i = 0; i < queryable.Count; i += VersionBatchSize)
        {
            // Only credit versions that were part of this batch's query, so an issue whose fix versions
            // span two batches is counted once per version.
            var batchSet = new HashSet<string>(queryable.Skip(i).Take(VersionBatchSize), StringComparer.OrdinalIgnoreCase);

            await SearchAsync(FixVersionInJql(batchSet), fields, issue =>
            {
                if (issue.Key == null) return;
                var jiraIssue = new JiraIssue
                {
                    Key = issue.Key,
                    IssueType = issue.Fields?.Issuetype?.Name ?? string.Empty,
                    Summary = issue.Fields?.Summary ?? string.Empty,
                    Estimate = issue.Fields?.TimeOriginalEstimate.HasValue == true
                        ? issue.Fields.TimeOriginalEstimate.Value / 3600.0
                        : null
                };

                foreach (var name in issue.Fields?.FixVersions?.Select(v => v.Name) ?? [])
                {
                    if (name == null || !batchSet.TryGetValue(name, out var requested)) continue;
                    if (added.Add((requested, issue.Key)))
                        result[requested].Add(jiraIssue);
                }
            });
        }

        return result;
    }

    /// <summary>
    /// Resolves, for each supplied issue, the budget of the nearest ancestor in its parent chain.
    /// Ancestors are fetched level by level and only as far as needed: the walk stops climbing a
    /// branch as soon as an ancestor carries a budget. Returns a map of issue key -> inherited
    /// budget, containing only the issues where an ancestor budget was actually found.
    /// </summary>
    public async Task<Dictionary<string, string>> GetInheritedParentBudgetsAsync(
        IEnumerable<JiraIssue> issues)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var issueList = issues.ToList();

        await ResolveFieldIdsAsync(issueList.Select(i => i.Key)
            .Concat(issueList.Select(i => i.ParentKey ?? string.Empty)));
        if (_budgetFieldId == null)
            return result; // Budget field unknown - nothing to inherit.

        var starting = issueList.Where(i => !string.IsNullOrEmpty(i.ParentKey)).ToList();
        if (starting.Count == 0) return result;

        var ancestors = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var pending = starting
            .Select(i => i.ParentKey!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        while (pending.Count > 0)
        {
            Console.WriteLine($"  Fetching {pending.Count} parent issue(s) to inherit budgets.");

            var batchParents = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
            var parentErrors = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < pending.Count; i += KeyChunkSize)
                await FetchRawIssuesBulkAsync(pending.Skip(i).Take(KeyChunkSize).ToList(), batchParents, parentErrors);

            foreach (var kv in parentErrors)
                Console.WriteLine($"  Warning: Parent issue '{kv.Key}' could not be fetched (status {kv.Value}); budget will not be inherited.");

            foreach (var kv in batchParents)
                ancestors[kv.Key] = kv.Value;
            foreach (var key in pending)
                visited.Add(key);

            // Climb only the branches that still have no budget to offer.
            pending = batchParents.Values
                .Where(i => i != null
                            && string.IsNullOrWhiteSpace(i.Budget)
                            && !string.IsNullOrEmpty(i.ParentKey))
                .Select(i => i!.ParentKey!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(pk => !visited.Contains(pk))
                .ToList();
        }

        foreach (var issue in starting)
        {
            var inherited = NearestAncestorBudget(issue, ancestors);
            if (string.IsNullOrWhiteSpace(inherited)) continue;

            result[issue.Key] = inherited!;
            Console.WriteLine($"  Issue '{issue.Key}' has no budget; inheriting from parent chain: {inherited}");
        }

        return result;
    }

    /// <summary>
    /// Walks up the parent chain of <paramref name="issue"/> and returns the budget of the nearest
    /// ancestor that has one. Returns <c>null</c> when no ancestor carries a budget, when an
    /// ancestor was not fetched, or when the chain is circular.
    /// </summary>
    public static string? NearestAncestorBudget(
        JiraIssue issue,
        IReadOnlyDictionary<string, JiraIssue?> issuesByKey)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { issue.Key };
        var current = issue;
        while (!string.IsNullOrEmpty(current.ParentKey) && visited.Add(current.ParentKey))
        {
            if (!issuesByKey.TryGetValue(current.ParentKey, out var parent) || parent == null)
                return null;
            if (!string.IsNullOrWhiteSpace(parent.Budget))
                return parent.Budget;
            current = parent;
        }

        return null;
    }

    /// <summary>
    /// For each supplied fix version, looks up the "Delivery" (leverance) issues that belong to it
    /// and returns a map of fix version name → budget. This lets issues inherit the budget of the
    /// leverance task that shares their fix version, even when that leverance task has no Toggl time
    /// logged and is therefore not part of the bulk-fetched issue set.
    /// Needs the Budget field ID, which is resolved from issue keys by an earlier call
    /// (<see cref="GetIssuesBulkAsync"/>); without it the map is empty.
    /// </summary>
    public async Task<Dictionary<string, string>> GetLeveranceBudgetsByFixVersionAsync(IEnumerable<string> versions)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (_budgetFieldId == null)
            return result; // Budget field unknown — nothing to inherit.
        var budgetFieldId = _budgetFieldId;

        var distinctVersions = versions
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Query in batches with `fixVersion in (...)` so many versions resolve in few requests.
        for (int i = 0; i < distinctVersions.Count; i += VersionBatchSize)
        {
            var batch = distinctVersions.Skip(i).Take(VersionBatchSize).ToList();
            var jql = $"{FixVersionInJql(batch)} AND issuetype = Delivery";

            try
            {
                await SearchAsync(jql, ["issuetype", "fixVersions", budgetFieldId], issue =>
                {
                    var budget = ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(budgetFieldId));
                    if (string.IsNullOrWhiteSpace(budget)) return;

                    // Map every fix version of the leverance task (first one wins per version).
                    var fixVersionNames = issue.Fields?.FixVersions?
                        .Select(v => v.Name ?? string.Empty)
                        .Where(n => !string.IsNullOrWhiteSpace(n)) ?? Enumerable.Empty<string>();

                    foreach (var name in fixVersionNames)
                    {
                        if (!result.ContainsKey(name))
                            result[name] = budget!;
                    }
                });
            }
            catch (AtlassianToolException ex)
            {
                Console.Error.WriteLine(
                    $"  Warning: Jira returned {StatusText(ex)} when looking up leverance budgets: {ex.Message}");
            }
        }

        return result;
    }

    /// <summary>
    /// Returns the subset of the supplied fix versions that have at least one issue updated in Jira
    /// on or after <paramref name="updatedSince"/>. Used to decide which leverances are still active
    /// enough to appear on the Plan sheet.
    /// </summary>
    public async Task<HashSet<string>> GetFixVersionsWithRecentlyUpdatedIssuesAsync(
        IEnumerable<string> versions, DateTime updatedSince)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var distinctVersions = versions
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (distinctVersions.Count == 0)
            return result;

        var sinceStr = updatedSince.ToString("yyyy-MM-dd");

        for (int i = 0; i < distinctVersions.Count; i += VersionBatchSize)
        {
            var batch = distinctVersions.Skip(i).Take(VersionBatchSize).ToList();
            var batchSet = new HashSet<string>(batch, StringComparer.OrdinalIgnoreCase);
            var jql = $"{FixVersionInJql(batch)} AND updated >= \"{sinceStr}\"";

            try
            {
                await SearchAsync(jql, ["fixVersions"], issue =>
                {
                    var fixVersionNames = issue.Fields?.FixVersions?
                        .Select(v => v.Name ?? string.Empty)
                        .Where(n => !string.IsNullOrWhiteSpace(n)) ?? Enumerable.Empty<string>();

                    // Only credit versions that were part of this batch's query.
                    foreach (var name in fixVersionNames)
                    {
                        if (batchSet.Contains(name))
                            result.Add(name);
                    }
                });
            }
            catch (AtlassianToolException ex)
            {
                Console.Error.WriteLine(
                    $"  Warning: Jira returned {StatusText(ex)} when checking recently updated issues: {ex.Message}");
            }
        }

        return result;
    }

    private static string StatusText(AtlassianToolException ex) =>
        ex.StatusCode?.ToString() ?? "an error";

    private static void PrintProgressBar(int current, int total)
    {
        const int barWidth = 30;
        int filled = total > 0 ? (int)((double)current / total * barWidth) : barWidth;
        var bar = new string('█', filled) + new string('░', barWidth - filled);
        Console.WriteLine($"  Fetching Jira issues: [{bar}] {current}/{total}");
    }
}
