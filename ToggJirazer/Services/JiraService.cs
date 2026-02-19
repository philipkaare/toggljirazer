using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

public class JiraService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly JiraConfig _config;

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

    public async Task<JiraIssue?> GetIssueAsync(string issueKey)
    {
        var url = $"rest/api/3/issue/{Uri.EscapeDataString(issueKey)}?fields=summary,issuetype,customfield_10016,customfield_10014";

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

        return new JiraIssue
        {
            Key = issue.Key ?? issueKey,
            IssueType = issue.Fields?.Issuetype?.Name ?? string.Empty,
            Summary = issue.Fields?.Summary ?? string.Empty,
            Budget = ExtractStringField(issue.Fields?.Customfield_10016),
            Account = ExtractStringField(issue.Fields?.Customfield_10014)
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
        public object? Customfield_10016 { get; set; }
        public object? Customfield_10014 { get; set; }
    }

    private sealed class JiraIssueType
    {
        public string? Name { get; set; }
    }

    public void Dispose() => _httpClient.Dispose();
}
