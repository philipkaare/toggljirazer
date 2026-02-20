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
                  "?response_format=json_row&include_dicts=true";
        int page = 1;
        const int perPage = 50;

        Console.WriteLine($"Fetching Toggl entries from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}...");

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

            // The API returns columns in the order of the requested attributes (AttributeNames),
            // so the fallback is always aligned with the request even if the schema is omitted.
            var schema = analyticsResponse.Schema ?? [.. AttributeNames];
            var schemaIndex = schema
                .Select((name, idx) => (name, idx))
                .ToDictionary(x => x.name, x => x.idx, StringComparer.OrdinalIgnoreCase);

            var users = BuildUserLookup(analyticsResponse.Dicts);

            var pageEntryCount = analyticsResponse.Data.Count;
            foreach (var row in analyticsResponse.Data)
            {
                var entry = MapAnalyticsRow(schemaIndex, row, users);
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

    private static Dictionary<string, string> BuildUserLookup(JsonElement? dicts)
    {
        var users = new Dictionary<string, string>();
        if (dicts == null || dicts.Value.ValueKind != JsonValueKind.Object)
            return users;

        if (!dicts.Value.TryGetProperty("users", out var usersElement) ||
            usersElement.ValueKind != JsonValueKind.Object)
            return users;

        foreach (var user in usersElement.EnumerateObject())
        {
            if (user.Value.TryGetProperty("name", out var nameProp))
                users[user.Name] = nameProp.GetString() ?? string.Empty;
        }

        return users;
    }

    private static TogglTimeEntry? MapAnalyticsRow(
        Dictionary<string, int> schemaIndex,
        List<JsonElement> row,
        Dictionary<string, string> users)
    {
        var id = GetLong(row, schemaIndex, "time_entry_id");
        if (id == 0)
            return null;

        var description = GetString(row, schemaIndex, "description") ?? string.Empty;
        var durationSeconds = GetLong(row, schemaIndex, "duration");
        var projectId = GetLong(row, schemaIndex, "project_id");
        var start = GetDateTime(row, schemaIndex, "start");
        var stop = GetDateTime(row, schemaIndex, "stop");
        var userId = GetLong(row, schemaIndex, "user_id").ToString();
        users.TryGetValue(userId, out var userName);

        return new TogglTimeEntry
        {
            Id = id,
            Description = description,
            User = userName ?? userId,
            Email = string.Empty,
            Start = start,
            End = stop,
            Duration = durationSeconds * 1000L,
            Project = string.Empty,
            ProjectId = projectId
        };
    }

    private static long GetLong(List<JsonElement> row, Dictionary<string, int> schemaIndex, string column)
    {
        if (!schemaIndex.TryGetValue(column, out var idx) || idx >= row.Count)
            return 0;
        var element = row[idx];
        if (element.ValueKind == JsonValueKind.Number)
        {
            if (element.TryGetInt64(out var longVal))
                return longVal;
            if (element.TryGetDouble(out var dblVal))
                return (long)Math.Round(dblVal);
        }
        return 0;
    }

    private static string? GetString(List<JsonElement> row, Dictionary<string, int> schemaIndex, string column)
    {
        if (!schemaIndex.TryGetValue(column, out var idx) || idx >= row.Count)
            return null;
        var element = row[idx];
        return element.ValueKind == JsonValueKind.String ? element.GetString() : null;
    }

    private static DateTime GetDateTime(List<JsonElement> row, Dictionary<string, int> schemaIndex, string column)
    {
        if (!schemaIndex.TryGetValue(column, out var idx) || idx >= row.Count)
            return default;
        var element = row[idx];
        if (element.ValueKind == JsonValueKind.String)
        {
            var s = element.GetString();
            if (s != null && DateTime.TryParse(s, null, System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
                return dt;
        }
        return default;
    }

    // Response models for the analytics API
    private sealed class TogglAnalyticsResponse
    {
        [JsonPropertyName("schema")]
        public List<string>? Schema { get; set; }

        [JsonPropertyName("data")]
        public List<List<JsonElement>>? Data { get; set; }

        [JsonPropertyName("dicts")]
        public JsonElement? Dicts { get; set; }

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

    public void Dispose() => _httpClient.Dispose();
}
