using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using ToggJirazer.Models;

namespace ToggJirazer.Services;

/// <summary>
/// Reads and writes the application configuration to/from appsettings.json.
/// </summary>
public class SettingsService
{
    private readonly string _settingsPath;

    public SettingsService(string settingsPath)
    {
        _settingsPath = settingsPath;
    }

    public AppConfig Load()
    {
        var node = LoadJsonNode(_settingsPath) ?? new JsonObject();

        var devPath = Path.Combine(
            Path.GetDirectoryName(_settingsPath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(_settingsPath) + ".development" + Path.GetExtension(_settingsPath));

        var devNode = LoadJsonNode(devPath);
        if (devNode != null)
            MergeInto(node, devNode);

        try
        {
            return new AppConfig
            {
                Toggl = new TogglConfig
                {
                    ApiToken = node["Toggl"]?["ApiToken"]?.GetValue<string>() ?? string.Empty,
                    OrganizationId = node["Toggl"]?["OrganizationId"]?.GetValue<long>() ?? 0,
                    WorkspaceId = node["Toggl"]?["WorkspaceId"]?.GetValue<long>() ?? 0,
                    ProjectId = node["Toggl"]?["ProjectId"]?.GetValue<long>() ?? 0
                },
                Jira = new JiraConfig
                {
                    BaseUrl = node["Jira"]?["BaseUrl"]?.GetValue<string>() ?? string.Empty,
                    UserEmail = node["Jira"]?["UserEmail"]?.GetValue<string>() ?? string.Empty,
                    ApiToken = node["Jira"]?["ApiToken"]?.GetValue<string>() ?? string.Empty,
                    BudgetFieldName = node["Jira"]?["BudgetFieldName"]?.GetValue<string>() ?? "Budget",
                    AccountFieldName = node["Jira"]?["AccountFieldName"]?.GetValue<string>() ?? "Account"
                },
                Report = new ReportConfig
                {
                    StartDate = node["Report"]?["StartDate"]?.ToString().NullIfEmpty(),
                    EndDate = node["Report"]?["EndDate"]?.ToString().NullIfEmpty(),
                    OutputFile = node["Report"]?["OutputFile"]?.GetValue<string>() ?? "report.xlsx",
                    Format = node["Report"]?["Format"]?.GetValue<string>() ?? "xlsx"
                }
            };
        }
        catch
        {
            return new AppConfig();
        }
    }

    private static JsonObject? LoadJsonNode(string path)
    {
        if (!File.Exists(path)) return null;
        try
        {
            var json = File.ReadAllText(path);
            return JsonNode.Parse(json)?.AsObject();
        }
        catch { return null; }
    }

    private static void MergeInto(JsonObject target, JsonObject source)
    {
        foreach (var (key, value) in source)
        {
            if (value is JsonObject sourceObj && target[key] is JsonObject targetObj)
                MergeInto(targetObj, sourceObj);
            else
                target[key] = value?.DeepClone();
        }
    }

    public void Save(AppConfig config)
    {
        var node = new JsonObject
        {
            ["Toggl"] = new JsonObject
            {
                ["ApiToken"] = config.Toggl.ApiToken,
                ["OrganizationId"] = config.Toggl.OrganizationId,
                ["WorkspaceId"] = config.Toggl.WorkspaceId,
                ["ProjectId"] = config.Toggl.ProjectId
            },
            ["Jira"] = new JsonObject
            {
                ["BaseUrl"] = config.Jira.BaseUrl,
                ["UserEmail"] = config.Jira.UserEmail,
                ["ApiToken"] = config.Jira.ApiToken,
                ["BudgetFieldName"] = config.Jira.BudgetFieldName,
                ["AccountFieldName"] = config.Jira.AccountFieldName
            },
            ["Report"] = new JsonObject
            {
                ["StartDate"] = config.Report.StartDate ?? string.Empty,
                ["EndDate"] = config.Report.EndDate ?? string.Empty,
                ["OutputFile"] = config.Report.OutputFile,
                ["Format"] = config.Report.Format
            }
        };

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(_settingsPath, node.ToJsonString(options), System.Text.Encoding.UTF8);
    }
}

file static class StringExtensions
{
    internal static string? NullIfEmpty(this string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
