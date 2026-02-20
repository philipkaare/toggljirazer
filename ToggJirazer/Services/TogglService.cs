using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
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
        int page = 1;

        Console.WriteLine($"Fetching Toggl entries from {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}...");

        while (true)
        {
            var url = $"https://api.track.toggl.com/reports/api/v2/details" +
                      $"?workspace_id={_config.WorkspaceId}" +
                      $"&project_ids={_config.ProjectId}" +
                      $"&since={startDate:yyyy-MM-dd}" +
                      $"&until={endDate:yyyy-MM-dd}" +
                      $"&page={page}" +
                      $"&user_agent=ToggJirazer";

            Console.WriteLine($"  Fetching page {page}...");

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

            var json = await response.Content.ReadAsStringAsync();
            var report = JsonSerializer.Deserialize<TogglDetailedReportResponse>(json, JsonOptions);

            if (report?.Data == null || report.Data.Count == 0)
                break;

            allEntries.AddRange(report.Data.Select(MapEntry));

            Console.WriteLine($"  Retrieved {allEntries.Count} of {report.TotalCount} entries.");

            if (allEntries.Count >= report.TotalCount)
                break;

            page++;
        }

        Console.WriteLine($"Total Toggl entries fetched: {allEntries.Count}");
        return allEntries;
    }

    private static TogglTimeEntry MapEntry(TogglDetailedReportEntry entry) => new()
    {
        Id = entry.Id,
        Description = entry.Description ?? string.Empty,
        User = entry.User ?? string.Empty,
        Email = entry.Email ?? string.Empty,
        Start = entry.Start,
        End = entry.End,
        Duration = entry.Dur,
        Project = entry.Project ?? string.Empty,
        ProjectId = entry.Pid
    };

    // Internal response models for JSON deserialization
    private sealed class TogglDetailedReportResponse
    {
        public List<TogglDetailedReportEntry> Data { get; set; } = new();
        public int TotalCount { get; set; }
        public int PerPage { get; set; }
    }

    private sealed class TogglDetailedReportEntry
    {
        public long Id { get; set; }
        public string? Description { get; set; }
        public string? User { get; set; }
        public string? Email { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public long Dur { get; set; }
        public string? Project { get; set; }
        public long Pid { get; set; }
    }

    public void Dispose() => _httpClient.Dispose();
}
