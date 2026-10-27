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

        var devNode = LoadJsonNode(DevelopmentPathFor(_settingsPath));
        if (devNode != null)
            MergeInto(node, devNode);

        try
        {
            var cloudId = node["Jira"]?["CloudId"]?.GetValue<string>() ?? string.Empty;
            var siteUrl = node["Jira"]?["SiteUrl"]?.GetValue<string>() ?? string.Empty;
            // Settings from before the Atlassian sign-in: keep the old base URL as the site to show.
            if (cloudId.Length == 0 && siteUrl.Length == 0)
                siteUrl = node["Jira"]?["BaseUrl"]?.GetValue<string>() ?? string.Empty;

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
                    CloudId = cloudId,
                    SiteUrl = siteUrl,
                    BudgetFieldName = node["Jira"]?["BudgetFieldName"]?.GetValue<string>() ?? "Budget",
                    AccountFieldName = node["Jira"]?["AccountFieldName"]?.GetValue<string>() ?? "Account"
                },
                Report = new ReportConfig
                {
                    StartDate = node["Report"]?["StartDate"]?.ToString().NullIfEmpty(),
                    EndDate = node["Report"]?["EndDate"]?.ToString().NullIfEmpty(),
                    OutputFile = node["Report"]?["OutputFile"]?.GetValue<string>() ?? "report.xlsx",
                    Format = node["Report"]?["Format"]?.GetValue<string>() ?? "xlsx"
                },
                Email = new EmailConfig
                {
                    SmtpHost = node["Email"]?["SmtpHost"]?.GetValue<string>() ?? string.Empty,
                    SmtpPort = node["Email"]?["SmtpPort"]?.GetValue<int>() ?? 587,
                    UseSsl = node["Email"]?["UseSsl"]?.GetValue<bool>() ?? false,
                    Username = node["Email"]?["Username"]?.GetValue<string>() ?? string.Empty,
                    Password = node["Email"]?["Password"]?.GetValue<string>() ?? string.Empty,
                    FromAddress = node["Email"]?["FromAddress"]?.GetValue<string>() ?? string.Empty,
                    Recipients = node["Email"]?["Recipients"]?.GetValue<string>() ?? string.Empty,
                    Subject = node["Email"]?["Subject"]?.GetValue<string>() ?? new EmailConfig().Subject,
                    Body = node["Email"]?["Body"]?.GetValue<string>() ?? new EmailConfig().Body
                }
            };
        }
        catch
        {
            return new AppConfig();
        }
    }

    /// <summary>The sibling development overrides file (appsettings.development.json).</summary>
    public static string DevelopmentPathFor(string settingsPath) => Path.Combine(
        Path.GetDirectoryName(settingsPath) ?? string.Empty,
        Path.GetFileNameWithoutExtension(settingsPath) + ".development" + Path.GetExtension(settingsPath));

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

    /// <summary>
    /// Keeps development overrides out of the saved file: every value that still equals its
    /// development override is written as the base file's value (or omitted if it has none).
    /// </summary>
    private static void RestoreDevelopmentOverrides(JsonObject target, JsonObject dev, JsonObject? existing)
    {
        foreach (var (key, devValue) in dev)
        {
            if (devValue is JsonObject devObj)
            {
                if (target[key] is JsonObject targetObj)
                    RestoreDevelopmentOverrides(targetObj, devObj, existing?[key] as JsonObject);
                continue;
            }

            if (!target.ContainsKey(key) || !JsonNode.DeepEquals(target[key], devValue))
                continue;

            if (existing != null && existing.TryGetPropertyValue(key, out var baseValue))
                target[key] = baseValue?.DeepClone();
            else
                target.Remove(key);
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
                ["CloudId"] = config.Jira.CloudId,
                ["SiteUrl"] = config.Jira.SiteUrl,
                ["BudgetFieldName"] = config.Jira.BudgetFieldName,
                ["AccountFieldName"] = config.Jira.AccountFieldName
            },
            ["Report"] = new JsonObject
            {
                ["StartDate"] = config.Report.StartDate ?? string.Empty,
                ["EndDate"] = config.Report.EndDate ?? string.Empty,
                ["OutputFile"] = config.Report.OutputFile,
                ["Format"] = config.Report.Format
            },
            ["Email"] = new JsonObject
            {
                ["SmtpHost"] = config.Email.SmtpHost,
                ["SmtpPort"] = config.Email.SmtpPort,
                ["UseSsl"] = config.Email.UseSsl,
                ["Username"] = config.Email.Username,
                ["Password"] = config.Email.Password,
                ["FromAddress"] = config.Email.FromAddress,
                ["Recipients"] = config.Email.Recipients,
                ["Subject"] = config.Email.Subject,
                ["Body"] = config.Email.Body
            }
        };

        var devNode = LoadJsonNode(DevelopmentPathFor(_settingsPath));
        if (devNode != null)
            RestoreDevelopmentOverrides(node, devNode, LoadJsonNode(_settingsPath));

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(_settingsPath, node.ToJsonString(options), System.Text.Encoding.UTF8);
    }
}

file static class StringExtensions
{
    internal static string? NullIfEmpty(this string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
