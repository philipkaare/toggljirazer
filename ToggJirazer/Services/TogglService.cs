using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

public class TogglService : IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly TogglConfig _config;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private static readonly string[] AttributeNames =
    [
        "client_id", "description", "duration", "project_id", "start", "stop",
        "tag_ids", "time_entry_id", "user_id", "user_timezone", "task_id",
        "billable_duration", "billable"
    ];

    public TogglService(TogglConfig config)
    {
        _config = config;
        _httpClient = new HttpClient();
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{config.ApiToken}:api_token"));
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "ToggJirazer/1.0");
    }

    public async Task<List<TogglTimeEntry>> GetDetailedReportAsync(DateTime startDate, DateTime endDate)
    {
        var allEntries = new List<TogglTimeEntry>();
        var url = $"https://track.toggl.com/analytics/api/organizations/{_config.OrganizationId}/query" +
                  "?response_format=json_row";
        int page = 1;
        const int perPage = 50;

        Console.WriteLine($"Fetching Toggl entries from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}...");

        var users = await GetWorkspaceUsersAsync();
        while (true)
        {
            Console.WriteLine($"  Fetching page {page}...");

            var requestBody = BuildRequest(page, perPage, startDate.ToString("yyyy-MM-dd"), endDate.ToString("yyyy-MM-dd"));
            var content = new StringContent(requestBody.ToJsonString(), Encoding.UTF8, "application/json");

            HttpResponseMessage response;
            try
            {
                response = await _httpClient.PostAsync(url, content);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to connect to Toggl API: {ex.Message}. " +
                    "Please check your network connection.", ex);
            }

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                if (response.StatusCode == System.Net.HttpStatusCode.Forbidden ||
                    response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    throw new InvalidOperationException(
                        "Toggl authentication failed (403/401). Please verify that 'Toggl:ApiToken', " +
                        "'Toggl:OrganizationId', and 'Toggl:WorkspaceId' in appsettings.json are correct.");
                }
                throw new InvalidOperationException(
                    $"Toggl API returned {(int)response.StatusCode}: {error}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var analyticsResponse = JsonSerializer.Deserialize<TogglAnalyticsResponse>(responseJson, JsonOptions);

            if (analyticsResponse?.Data == null || analyticsResponse.Data.Count == 0)
                break;

            var pageEntryCount = analyticsResponse.Data.Count;
            foreach (var row in analyticsResponse.Data)
            {
                var entry = MapAnalyticsEntry(row, users);
                if (entry != null)
                    allEntries.Add(entry);
            }

            Console.WriteLine($"  Retrieved {allEntries.Count} entries so far.");

            if (pageEntryCount < perPage)
                break;

            page++;
        }

        Console.WriteLine($"Total Toggl entries fetched: {allEntries.Count}");
        return allEntries;
    }

    private JsonObject BuildRequest(int page, int perPage, string startDateStr, string endDateStr)
    {
        var filters = new JsonArray();
        filters.Add(new JsonObject
        {
            ["property"] = "workspace_id",
            ["operator"] = "=",
            ["value"] = _config.WorkspaceId
        });

        if (_config.ProjectId != 0)
        {
            filters.Add(new JsonObject
            {
                ["operator"] = "and",
                ["conditions"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["property"] = "project_id",
                        ["operator"] = "in",
                        ["value"] = new JsonArray { _config.ProjectId }
                    }
                }
            });
        }

        var attributes = new JsonArray();
        foreach (var attr in AttributeNames)
            attributes.Add(new JsonObject { ["property"] = attr });

        return new JsonObject
        {
            ["pagination"] = new JsonObject { ["per_page"] = perPage, ["page"] = page },
            ["transformations"] = new JsonArray(),
            ["period"] = new JsonObject { ["from"] = startDateStr, ["to"] = endDateStr },
            ["filters"] = filters,
            ["attributes"] = attributes,
            ["ordinations"] = new JsonArray
            {
                new JsonObject { ["property"] = "start", ["direction"] = "asc" }
            }
        };
    }

    private async Task<Dictionary<long, string>> GetWorkspaceUsersAsync()
    {
        var url = $"https://api.track.toggl.com/api/v9/organizations/{_config.OrganizationId}" +
                  $"/workspaces/{_config.WorkspaceId}/workspace_users";

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.GetAsync(url);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to connect to Toggl API: {ex.Message}. " +
                "Please check your network connection.", ex);
        }

        if (!response.IsSuccessStatusCode)
        {
            var error = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException(
                $"Toggl API returned {(int)response.StatusCode} when fetching workspace users: {error}");
        }

        var json = await response.Content.ReadAsStringAsync();
        var workspaceUsers = JsonSerializer.Deserialize<List<WorkspaceUser>>(json, JsonOptions) ?? [];
        return workspaceUsers
            .Where(u => u.UserId != 0 && u.Name != null)
            .ToDictionary(u => u.UserId, u => u.Name!);
    }

    private static TogglTimeEntry? MapAnalyticsEntry(
        TogglAnalyticsEntry row,
        Dictionary<long, string> users)
    {
        if (row.TimeEntryId == 0)
            return null;

        users.TryGetValue(row.UserId, out var userName);

        DateTime start = default, stop = default;
        if (row.Start != null)
            DateTime.TryParse(row.Start, null, System.Globalization.DateTimeStyles.RoundtripKind, out start);
        if (row.Stop != null)
            DateTime.TryParse(row.Stop, null, System.Globalization.DateTimeStyles.RoundtripKind, out stop);

        return new TogglTimeEntry
        {
            Id = row.TimeEntryId,
            Description = row.Description ?? string.Empty,
            User = userName ?? row.UserId.ToString(),
            Email = string.Empty,
            Start = start,
            End = stop,
            Duration = row.Duration,
            Project = string.Empty,
            ProjectId = row.ProjectId ?? 0
        };
    }

    // Response models for the analytics API
    private sealed class WorkspaceUser
    {
        [JsonPropertyName("user_id")]
        public long UserId { get; set; }

        [JsonPropertyName("uid")]
        public long Uid { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("email")]
        public string? Email { get; set; }
    }

    private sealed class TogglAnalyticsEntry
    {
        [JsonPropertyName("time_entry_id")]
        public long TimeEntryId { get; set; }

        [JsonPropertyName("description")]
        public string? Description { get; set; }

        [JsonPropertyName("duration")]
        public long Duration { get; set; }

        [JsonPropertyName("project_id")]
        public long? ProjectId { get; set; }

        [JsonPropertyName("start")]
        public string? Start { get; set; }

        [JsonPropertyName("stop")]
        public string? Stop { get; set; }

        [JsonPropertyName("user_id")]
        public long UserId { get; set; }

        [JsonPropertyName("user_timezone")]
        public string? UserTimezone { get; set; }

        [JsonPropertyName("billable")]
        public bool Billable { get; set; }

        [JsonPropertyName("billable_duration")]
        public long BillableDuration { get; set; }

        [JsonPropertyName("client_id")]
        public long? ClientId { get; set; }

        [JsonPropertyName("tag_ids")]
        public List<long>? TagIds { get; set; }
    }

    private sealed class TogglAnalyticsResponse
    {
        [JsonPropertyName("data_json_row")]
        public List<TogglAnalyticsEntry>? Data { get; set; }

        [JsonPropertyName("pagination")]
        public TogglAnalyticsPagination? Pagination { get; set; }
    }

    private sealed class TogglAnalyticsPagination
    {
        [JsonPropertyName("total_count")]
        public int TotalCount { get; set; }

        [JsonPropertyName("per_page")]
        public int PerPage { get; set; }

        [JsonPropertyName("page")]
        public int Page { get; set; }
    }

    public async Task<List<TogglTimeEntry>> GetAllEntriesAsync()
    {
        return await GetDetailedReportAsync(new DateTime(2000, 1, 1), DateTime.Today);
    }

    public void Dispose() => _httpClient.Dispose();
}
