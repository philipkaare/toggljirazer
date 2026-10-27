using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ToggJirazer.Services;

public sealed record TogglAccount(string ApiToken, string FullName, string Email);
public sealed record TogglOrganization(long Id, string Name);
public sealed record TogglWorkspace(long Id, string Name, long OrganizationId);
public sealed record TogglProject(long Id, string Name);
public sealed class TogglSetupException(string message) : Exception(message);

/// <summary>
/// Toggl API v9 client for the setup wizard. Credentials live only in memory for the
/// duration of a request and never appear in exception messages.
/// </summary>
public sealed class TogglSetupService : IDisposable
{
    private const string BaseUrl = "https://api.track.toggl.com/api/v9";
    private const int PerPage = 200;
    // Guards against a Toggl that ignores "page" and keeps answering with the same full page.
    private const int MaxProjectPages = 50;
    private const string WrongPasswordMessage =
        "Wrong email or password. If you sign in to Toggl with Google, Apple or SSO, use an API token instead.";
    private const string TokenRejectedMessage = "That API token was not accepted by Toggl.";

    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    private readonly HttpClient _http;

    public TogglSetupService(HttpMessageHandler? handler = null)
    {
        _http = new HttpClient(handler ?? new HttpClientHandler());
        _http.DefaultRequestHeaders.Add("User-Agent", "ToggJirazer/1.0");
    }

    public async Task<TogglAccount> SignInAsync(string email, string password, CancellationToken ct = default)
    {
        var dto = await GetAsync<MeDto>($"{BaseUrl}/me", email, password, password, WrongPasswordMessage, ct);
        return ToAccount(dto);
    }

    public async Task<TogglAccount> ValidateTokenAsync(string apiToken, CancellationToken ct = default)
    {
        var dto = await GetAsync<MeDto>($"{BaseUrl}/me", apiToken, "api_token", apiToken, TokenRejectedMessage, ct);
        return ToAccount(dto);
    }

    public async Task<IReadOnlyList<TogglOrganization>> GetOrganizationsAsync(string apiToken, CancellationToken ct = default)
    {
        var list = await GetAsync<List<OrganizationDto>>($"{BaseUrl}/me/organizations", apiToken, "api_token", apiToken, TokenRejectedMessage, ct);
        return (list ?? []).Select(o => new TogglOrganization(o.Id, o.Name ?? "")).ToList();
    }

    public async Task<IReadOnlyList<TogglWorkspace>> GetWorkspacesAsync(string apiToken, long organizationId, CancellationToken ct = default)
    {
        var list = await GetAsync<List<WorkspaceDto>>($"{BaseUrl}/me/workspaces", apiToken, "api_token", apiToken, TokenRejectedMessage, ct);
        return (list ?? [])
            .Where(w => w.OrganizationId == organizationId)
            .Select(w => new TogglWorkspace(w.Id, w.Name ?? "", w.OrganizationId))
            .ToList();
    }

    public async Task<IReadOnlyList<TogglProject>> GetActiveProjectsAsync(string apiToken, long workspaceId, CancellationToken ct = default)
    {
        var all = new List<TogglProject>();
        long? previousFirstId = null;
        for (var page = 1; page <= MaxProjectPages; page++)
        {
            var url = $"{BaseUrl}/workspaces/{workspaceId}/projects?active=true&per_page={PerPage}&page={page}";
            var items = await GetAsync<List<ProjectDto>>(url, apiToken, "api_token", apiToken, TokenRejectedMessage, ct) ?? [];
            if (items.Count > 0 && items[0].Id == previousFirstId) break;
            all.AddRange(items.Select(p => new TogglProject(p.Id, p.Name ?? "")));
            if (items.Count < PerPage) break;
            previousFirstId = items[0].Id;
        }
        return all.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static TogglAccount ToAccount(MeDto? dto) =>
        dto is null || string.IsNullOrEmpty(dto.ApiToken)
            ? throw new TogglSetupException("Toggl returned an unexpected response.")
            : new TogglAccount(dto.ApiToken, dto.FullName ?? "", dto.Email ?? "");

    /// <param name="secret">The credential to keep out of messages: the password, or the API token itself.</param>
    private async Task<T?> GetAsync<T>(string url, string user, string password, string secret, string authFailureMessage,
        CancellationToken ct)
    {
        var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{user}:{password}"));
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);

        HttpResponseMessage response;
        try
        {
            response = await _http.SendAsync(request, ct);
        }
        catch (HttpRequestException ex)
        {
            throw new TogglSetupException($"Could not reach Toggl: {Scrub(ex.Message, secret, credentials)}");
        }

        using (response)
        {
            if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
                throw new TogglSetupException(authFailureMessage);

            var body = await response.Content.ReadAsStringAsync(ct);
            if (!response.IsSuccessStatusCode)
            {
                var detail = !string.IsNullOrWhiteSpace(body) && !ContainsSecret(body, secret, credentials)
                    ? body.Trim()
                    : response.ReasonPhrase ?? response.StatusCode.ToString();
                throw new TogglSetupException($"Toggl returned {(int)response.StatusCode}: {detail}");
            }

            try
            {
                return JsonSerializer.Deserialize<T>(body, JsonOptions);
            }
            catch (JsonException)
            {
                throw new TogglSetupException("Toggl returned an unexpected response.");
            }
        }
    }

    private static bool ContainsSecret(string text, string secret, string credentials) =>
        (secret.Length > 0 && text.Contains(secret, StringComparison.Ordinal)) ||
        text.Contains(credentials, StringComparison.Ordinal);

    private static string Scrub(string text, string secret, string credentials) =>
        ContainsSecret(text, secret, credentials) ? "network error" : text;

    public void Dispose() => _http.Dispose();

    private sealed class MeDto
    {
        [JsonPropertyName("api_token")] public string? ApiToken { get; set; }
        [JsonPropertyName("fullname")] public string? FullName { get; set; }
        [JsonPropertyName("email")] public string? Email { get; set; }
    }

    private sealed class OrganizationDto
    {
        [JsonPropertyName("id")] public long Id { get; set; }
        [JsonPropertyName("name")] public string? Name { get; set; }
    }

    private sealed class WorkspaceDto
    {
        [JsonPropertyName("id")] public long Id { get; set; }
        [JsonPropertyName("name")] public string? Name { get; set; }
        [JsonPropertyName("organization_id")] public long OrganizationId { get; set; }
    }

    private sealed class ProjectDto
    {
        [JsonPropertyName("id")] public long Id { get; set; }
        [JsonPropertyName("name")] public string? Name { get; set; }
    }
}
