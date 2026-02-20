using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
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
        var url = $"https://api.track.toggl.com/reports/api/v3/workspace/{_config.WorkspaceId}/search/time_entries";
        int? firstRowNumber = null;
        int page = 1;

        Console.WriteLine($"Fetching Toggl entries from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}...");

        while (true)
        {
            Console.WriteLine($"  Fetching page {page}...");

            var requestBody = new TogglSearchRequest
            {
                StartDate = startDate.ToString("yyyy-MM-dd"),
                EndDate = endDate.ToString("yyyy-MM-dd"),
                ProjectIds = _config.ProjectId != 0 ? new List<long> { _config.ProjectId } : null,
                FirstRowNumber = firstRowNumber
            };

            var requestJson = JsonSerializer.Serialize(requestBody, JsonOptions);
            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");

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
                        "Toggl authentication failed (403/401). Please verify that 'Toggl:ApiToken' " +
                        "and 'Toggl:WorkspaceId' in appsettings.json are correct.");
                }
                throw new InvalidOperationException(
                    $"Toggl API returned {(int)response.StatusCode}: {error}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var rows = JsonSerializer.Deserialize<List<TogglSearchResponseRow>>(responseJson, JsonOptions);

            if (rows == null || rows.Count == 0)
                break;

            foreach (var row in rows)
                allEntries.AddRange(row.TimeEntries.Select(te => MapEntry(row, te)));

            Console.WriteLine($"  Retrieved {allEntries.Count} entries so far.");

            if (response.Headers.TryGetValues("X-Next-Row-Number", out var headerValues) &&
                int.TryParse(headerValues.FirstOrDefault(), out var nextRowNumber))
            {
                firstRowNumber = nextRowNumber;
                page++;
            }
            else
            {
                break;
            }
        }

        Console.WriteLine($"Total Toggl entries fetched: {allEntries.Count}");
        return allEntries;
    }

    private static TogglTimeEntry MapEntry(TogglSearchResponseRow row, TogglSearchTimeEntry entry) => new()
    {
        Id = entry.Id,
        Description = row.Description ?? string.Empty,
        User = row.Username ?? string.Empty,
        Email = string.Empty,           // v3 API does not expose user email
        Start = entry.Start,
        End = entry.Stop,
        Duration = entry.Seconds * 1000L, // v3 returns seconds; convert to ms for compatibility
        Project = string.Empty,         // v3 API returns project_id only, not project name
        ProjectId = row.ProjectId
    };

    // Internal request/response models for the v3 search API
    private sealed class TogglSearchRequest
    {
        [JsonPropertyName("start_date")]
        public string StartDate { get; set; } = string.Empty;

        [JsonPropertyName("end_date")]
        public string EndDate { get; set; } = string.Empty;

        [JsonPropertyName("project_ids")]
        public List<long>? ProjectIds { get; set; }

        [JsonPropertyName("first_row_number")]
        public int? FirstRowNumber { get; set; }
    }

    private sealed class TogglSearchResponseRow
    {
        [JsonPropertyName("user_id")]
        public long UserId { get; set; }

        [JsonPropertyName("username")]
        public string? Username { get; set; }

        [JsonPropertyName("project_id")]
        public long ProjectId { get; set; }

        [JsonPropertyName("description")]
        public string? Description { get; set; }

        [JsonPropertyName("time_entries")]
        public List<TogglSearchTimeEntry> TimeEntries { get; set; } = new();
    }

    private sealed class TogglSearchTimeEntry
    {
        [JsonPropertyName("id")]
        public long Id { get; set; }

        [JsonPropertyName("seconds")]
        public long Seconds { get; set; }

        [JsonPropertyName("start")]
        public DateTime Start { get; set; }

        [JsonPropertyName("stop")]
        public DateTime Stop { get; set; }
    }

    public void Dispose() => _httpClient.Dispose();
}
