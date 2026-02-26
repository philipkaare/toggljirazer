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

    public async Task<List<JiraIssue>> GetIssuesByFixVersionAsync(string version)
    {
        var results = new List<JiraIssue>();
        int startAt = 0;
        const int maxResults = 100;
        int total;

        do
        {
            var jql = Uri.EscapeDataString($"fixVersion = \"{version}\"");
            var url = $"rest/api/3/search?jql={jql}&fields=summary,issuetype,timeoriginalestimate&startAt={startAt}&maxResults={maxResults}";

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

    public void Dispose() => _httpClient.Dispose();
}
