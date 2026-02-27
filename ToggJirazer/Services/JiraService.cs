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

        var fieldsParam = "summary,issuetype,fixVersions,timeoriginalestimate";
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
                : null
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

        var fields = new List<string> { "summary", "issuetype", "fixVersions", "timeoriginalestimate" };
        if (_budgetFieldId != null) fields.Add(_budgetFieldId);
        if (_accountFieldId != null) fields.Add(_accountFieldId);

        var result = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);
        var errorStatus = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var keyList = issueKeys.ToList();
        int fetched = 0;

        PrintProgressBar(fetched, keyList.Count);

        const int batchSize = 100;
        for (int i = 0; i < keyList.Count; i += batchSize)
        {
            var batch = keyList.Skip(i).Take(batchSize).ToList();
            var requestBody = new JiraBulkFetchRequest { IssueIdsOrKeys = batch, Fields = fields };
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

                    result[issue.Key] = new JiraIssue
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
                            : null
                    };
                }
            }

            if (bulkResponse?.Errors != null)
            {
                foreach (var err in bulkResponse.Errors)
                {
                    if (err.IssueKey != null)
                    {
                        result[err.IssueKey] = null;
                        errorStatus[err.IssueKey] = err.Status;
                    }
                }
            }

            fetched += batch.Count;
            PrintProgressBar(fetched, keyList.Count);
        }

        Console.WriteLine();

        foreach (var kv in errorStatus)
            Console.WriteLine($"  Warning: Jira issue '{kv.Key}' could not be fetched (status {kv.Value}).");

        return result;
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

        [System.Text.Json.Serialization.JsonExtensionData]
        public Dictionary<string, object?>? CustomFields { get; set; }
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

        [System.Text.Json.Serialization.JsonPropertyName("startAt")]
        public int StartAt { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("maxResults")]
        public int MaxResults { get; set; }
    }

    public async Task<List<JiraIssue>> GetIssuesByFixVersionAsync(string version)
    {
        var results = new List<JiraIssue>();
        int startAt = 0;
        const int maxResults = 100;
        int total;

        do
        {
            var requestBody = new JiraSearchRequest
            {
                Jql = $"fixVersion = \"{version}\"",
                Fields = new List<string> { "summary", "issuetype", "timeoriginalestimate" },
                StartAt = startAt,
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

            total = searchResult.Total;

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

            startAt += searchResult.Issues.Count;
        }
        while (startAt < total);

        return results;
    }

    private static void PrintProgressBar(int current, int total)
    {
        const int barWidth = 30;
        int filled = total > 0 ? (int)((double)current / total * barWidth) : barWidth;
        var bar = new string('█', filled) + new string('░', barWidth - filled);
        Console.Write($"\r  Fetching Jira issues: [{bar}] {current}/{total}");
    }

    public void Dispose() => _httpClient.Dispose();
}
