using ToggJirazer.Models;

namespace ToggJirazer.Cli;

/// <summary>
/// Validates an <see cref="AppConfig"/> before a report run.
/// Returns <c>null</c> when the config is complete, otherwise a message
/// naming the first missing required field.
/// </summary>
public static class ConfigValidator
{
    public static string? Validate(AppConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.Toggl.ApiToken))
            return "Toggl API Token is required.";
        if (config.Toggl.OrganizationId == 0)
            return "Toggl Organization ID is required.";
        if (config.Toggl.WorkspaceId == 0)
            return "Toggl Workspace ID is required.";
        if (config.Toggl.ProjectId == 0)
            return "Toggl Project ID is required.";
        if (string.IsNullOrWhiteSpace(config.Jira.CloudId))
            return "Pick your Atlassian site (Jira tab).";
        if (string.IsNullOrWhiteSpace(config.Report.OutputFile))
            return "Output File path is required.";
        return null;
    }
}
