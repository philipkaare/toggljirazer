using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

public class JiraService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly JiraConfig _config;

    private string? _budgetFieldId;
    private string? _accountFieldId;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public JiraService(JiraConfig config)
    {
        _config = config;
        _httpClient = new HttpClient();
        var credentials = Convert.ToBase64String(
            Encoding.ASCII.GetBytes($"{config.UserEmail}:{config.ApiToken}"));
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        _httpClient.BaseAddress = new Uri(config.BaseUrl.TrimEnd('/') + "/");
    }

    private async Task ResolveFieldIdsAsync()
    {
        if (_budgetFieldId != null && _accountFieldId != null) return;

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.GetAsync("rest/api/3/field");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to fetch Jira field metadata: {ex.Message}", ex);
        }

        if (!response.IsSuccessStatusCode) return;

        var json = await response.Content.ReadAsStringAsync();
        var fields = JsonSerializer.Deserialize<List<JiraFieldResponse>>(json, JsonOptions);
        if (fields == null) return;

        _budgetFieldId = fields.FirstOrDefault(f =>
            string.Equals(f.Name, _config.BudgetFieldName, StringComparison.OrdinalIgnoreCase))?.Id;
        _accountFieldId = fields.FirstOrDefault(f =>
            string.Equals(f.Name, _config.AccountFieldName, StringComparison.OrdinalIgnoreCase))?.Id;

        if (_budgetFieldId == null)
            Console.WriteLine($"  Warning: Jira field '{_config.BudgetFieldName}' not found in field metadata.");
        if (_accountFieldId == null)
            Console.WriteLine($"  Warning: Jira field '{_config.AccountFieldName}' not found in field metadata.");
    }

    public async Task<JiraIssue?> GetIssueAsync(string issueKey)
    {
        await ResolveFieldIdsAsync();

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
    /// Fetches a single Jira issue from the API and parses it into a <see cref="JiraIssue"/>
    /// without applying any parent-inheritance logic.
    /// </summary>
    private async Task<JiraIssue?> FetchRawIssueAsync(string issueKey)
    {
        var fieldsParam = "summary,issuetype,fixVersions,timeoriginalestimate,parent";
        if (_budgetFieldId != null) fieldsParam += $",{_budgetFieldId}";
        if (_accountFieldId != null) fieldsParam += $",{_accountFieldId}";

        var url = $"rest/api/3/issue/{Uri.EscapeDataString(issueKey)}?fields={fieldsParam}";

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.GetAsync(url);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to connect to Jira API: {ex.Message}. " +
                "Please check 'Jira:BaseUrl' in appsettings.json and your network connection.", ex);
        }

        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            Console.WriteLine($"  Warning: Jira issue '{issueKey}' not found.");
            return null;
        }

        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
            response.StatusCode == System.Net.HttpStatusCode.Forbidden)
        {
            throw new InvalidOperationException(
                "Jira authentication failed. Please verify 'Jira:UserEmail' and 'Jira:ApiToken' " +
                "in appsettings.json are correct.");
        }

        if (!response.IsSuccessStatusCode)
        {
            var error = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException(
                $"Jira API returned {(int)response.StatusCode} for issue '{issueKey}': {error}");
        }

        var json = await response.Content.ReadAsStringAsync();
        var issue = JsonSerializer.Deserialize<JiraIssueResponse>(json, JsonOptions);
        if (issue == null) return null;

        var budgetValue = _budgetFieldId != null
            ? ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(_budgetFieldId))
            : null;
        var accountValue = _accountFieldId != null
            ? ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(_accountFieldId))
            : null;

        return new JiraIssue
        {
            Key = issue.Key ?? issueKey,
            IssueType = issue.Fields?.Issuetype?.Name ?? string.Empty,
            Summary = issue.Fields?.Summary ?? string.Empty,
            Budget = budgetValue,
            Account = accountValue,
            FixVersions = issue.Fields?.FixVersions?.Select(v => v.Name ?? string.Empty)
                              .Where(n => !string.IsNullOrEmpty(n)).ToList() ?? new(),
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
        await ResolveFieldIdsAsync();

        var result = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        var errorStatus = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var keyList = issueKeys.ToList();
        int fetched = 0;

        PrintProgressBar(fetched, keyList.Count);

        const int batchSize = 100;
        for (int i = 0; i < keyList.Count; i += batchSize)
        {
            var batch = keyList.Skip(i).Take(batchSize).ToList();
            await FetchRawIssuesBulkAsync(batch, result, errorStatus);
            fetched += batch.Count;
            PrintProgressBar(fetched, keyList.Count);
        }

        Console.WriteLine();

        foreach (var kv in errorStatus)
            Console.WriteLine($"  Warning: Jira issue '{kv.Key}' could not be fetched (status {kv.Value}).");

        // Iteratively resolve the full parent chain for fix version inheritance.
        // Uses a visited set to prevent re-fetching and guard against circular references.
        // Note: TotalEstimateSum in version reports is computed via JQL fixVersion="..", so only
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
            await FetchRawIssuesBulkAsync(pendingParentKeys, batchParents, parentErrors);
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
    /// Fetches issues in bulk from the Jira API and adds the parsed results to <paramref name="target"/>.
    /// No parent-inheritance logic is applied. Errors are recorded in <paramref name="errorStatus"/>
    /// when provided; otherwise they are silently stored as <c>null</c> entries.
    /// </summary>
    private async Task FetchRawIssuesBulkAsync(
        IEnumerable<string> issueKeys,
        Dictionary<string, JiraIssue?> target,
        Dictionary<string, int>? errorStatus)
    {
        var fields = new List<string> { "summary", "issuetype", "fixVersions", "timeoriginalestimate", "parent" };
        if (_budgetFieldId != null) fields.Add(_budgetFieldId);
        if (_accountFieldId != null) fields.Add(_accountFieldId);

        var requestBody = new JiraBulkFetchRequest { IssueIdsOrKeys = issueKeys.ToList(), Fields = fields };
        var json = JsonSerializer.Serialize(requestBody, JsonOptions);
        var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.PostAsync("rest/api/3/issue/bulkfetch", content);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to connect to Jira API: {ex.Message}. " +
                "Please check 'Jira:BaseUrl' in appsettings.json and your network connection.", ex);
        }

        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
            response.StatusCode == System.Net.HttpStatusCode.Forbidden)
        {
            throw new InvalidOperationException(
                "Jira authentication failed. Please verify 'Jira:UserEmail' and 'Jira:ApiToken' " +
                "in appsettings.json are correct.");
        }

        if (!response.IsSuccessStatusCode)
        {
            var error = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException(
                $"Jira API returned {(int)response.StatusCode} for bulk fetch: {error}");
        }

        var responseJson = await response.Content.ReadAsStringAsync();
        var bulkResponse = JsonSerializer.Deserialize<JiraBulkFetchResponse>(responseJson, JsonOptions);

        if (bulkResponse?.Issues != null)
        {
            foreach (var issue in bulkResponse.Issues)
            {
                if (issue.Key == null) continue;

                var budgetValue = _budgetFieldId != null
                    ? ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(_budgetFieldId))
                    : null;
                var accountValue = _accountFieldId != null
                    ? ExtractStringField(issue.Fields?.CustomFields?.GetValueOrDefault(_accountFieldId))
                    : null;

                target[issue.Key] = new JiraIssue
                {
                    Key = issue.Key,
                    IssueType = issue.Fields?.Issuetype?.Name ?? string.Empty,
                    Summary = issue.Fields?.Summary ?? string.Empty,
                    Budget = budgetValue,
                    Account = accountValue,
                    FixVersions = issue.Fields?.FixVersions?.Select(v => v.Name ?? string.Empty)
                                      .Where(n => !string.IsNullOrEmpty(n)).ToList() ?? new(),
                    Estimate = issue.Fields?.TimeOriginalEstimate.HasValue == true
                        ? issue.Fields.TimeOriginalEstimate.Value / 3600.0
                        : null,
                    ParentKey = issue.Fields?.Parent?.Key
                };
            }
        }

        if (bulkResponse?.Errors != null)
        {
            foreach (var err in bulkResponse.Errors)
            {
                if (err.IssueKey != null)
                {
                    target[err.IssueKey] = null;
                    if (errorStatus != null)
                        errorStatus[err.IssueKey] = err.Status;
                }
            }
        }
    }

    // Internal response models
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
    }

    private sealed class JiraIssueType
    {
        public string? Name { get; set; }
    }

    private sealed class JiraFieldResponse
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
    }

    private sealed class JiraSearchResponse
    {
        public int Total { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("issues")]
        public List<JiraIssueResponse>? Issues { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("isLast")]
        public bool IsLast { get; set; }
    
  }

    private sealed class JiraBulkFetchRequest
    {
        [System.Text.Json.Serialization.JsonPropertyName("issueIdsOrKeys")]
        public List<string> IssueIdsOrKeys { get; set; } = new();

        [System.Text.Json.Serialization.JsonPropertyName("fields")]
        public List<string> Fields { get; set; } = new();
    }

    private sealed class JiraBulkFetchResponse
    {
        [System.Text.Json.Serialization.JsonPropertyName("issues")]
        public List<JiraIssueResponse>? Issues { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("errors")]
        public List<JiraBulkFetchError>? Errors { get; set; }
    }

    private sealed class JiraBulkFetchError
    {
        [System.Text.Json.Serialization.JsonPropertyName("issueKey")]
        public string? IssueKey { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("status")]
        public int Status { get; set; }
    }

    private sealed class JiraSearchRequest
    {
        [System.Text.Json.Serialization.JsonPropertyName("jql")]
        public string Jql { get; set; } = string.Empty;

        [System.Text.Json.Serialization.JsonPropertyName("fields")]
        public List<string> Fields { get; set; } = new();

        [System.Text.Json.Serialization.JsonPropertyName("nextPageToken")]
        public string? NextPageToken { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("maxResults")]
        public int MaxResults { get; set; }
    }

    public async Task<List<JiraIssue>> GetIssuesByFixVersionAsync(string version)
    {
        var results = new List<JiraIssue>();
        string? nextPageToken = null;
        bool isLastPage = false;
        const int maxResults = 100;

        do
        {
            var requestBody = new JiraSearchRequest
            {
                Jql = $"fixVersion = \"{version}\"",
                Fields = new List<string> { "summary", "issuetype", "timeoriginalestimate" },
                NextPageToken = nextPageToken,
                MaxResults = maxResults
            };
            var requestJson = JsonSerializer.Serialize(requestBody, JsonOptions);
            var content = new StringContent(requestJson, System.Text.Encoding.UTF8, "application/json");

            HttpResponseMessage response;
            try
            {
                response = await _httpClient.PostAsync("rest/api/3/search/jql", content);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to connect to Jira API: {ex.Message}. " +
                    "Please check 'Jira:BaseUrl' in appsettings.json and your network connection.", ex);
            }

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new InvalidOperationException(
                    $"Jira API returned {(int)response.StatusCode} when searching for fixVersion '{version}': {error}");
            }

            var json = await response.Content.ReadAsStringAsync();
            var searchResult = JsonSerializer.Deserialize<JiraSearchResponse>(json, JsonOptions);
            if (searchResult?.Issues == null) break;
            isLastPage = searchResult.IsLast;
            nextPageToken = searchResult.NextPageToken;
      
            foreach (var issue in searchResult.Issues)
            {
                if (issue.Key == null) continue;
                results.Add(new JiraIssue
                {
                    Key = issue.Key,
                    IssueType = issue.Fields?.Issuetype?.Name ?? string.Empty,
                    Summary = issue.Fields?.Summary ?? string.Empty,
                    Estimate = issue.Fields?.TimeOriginalEstimate.HasValue == true
                        ? issue.Fields.TimeOriginalEstimate.Value / 3600.0
                        : null
                });
            }

        }
        while (!isLastPage);

        return results;
    }

    private static void PrintProgressBar(int current, int total)
    {
        const int barWidth = 30;
        int filled = total > 0 ? (int)((double)current / total * barWidth) : barWidth;
        var bar = new string('█', filled) + new string('░', barWidth - filled);
        Console.WriteLine($"  Fetching Jira issues: [{bar}] {current}/{total}");
    }

    public void Dispose() => _httpClient.Dispose();
}
