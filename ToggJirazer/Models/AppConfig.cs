namespace ToggJirazer.Models;

public class AppConfig
{
    public TogglConfig Toggl { get; set; } = new();
    public JiraConfig Jira { get; set; } = new();
    public ReportConfig Report { get; set; } = new();
}

public class TogglConfig
{
    public string ApiToken { get; set; } = string.Empty;
    public long OrganizationId { get; set; }
    public long WorkspaceId { get; set; }
    public long ProjectId { get; set; }
}

public class JiraConfig
{
    public string BaseUrl { get; set; } = string.Empty;
    public string UserEmail { get; set; } = string.Empty;
    public string ApiToken { get; set; } = string.Empty;
}

public class ReportConfig
{
    public string? StartDate { get; set; }
    public string? EndDate { get; set; }
    public string OutputFile { get; set; } = "report.csv";
}
