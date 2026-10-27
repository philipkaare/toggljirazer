namespace ToggJirazer.Models;

public class AppConfig
{
    public TogglConfig Toggl { get; set; } = new();
    public JiraConfig Jira { get; set; } = new();
    public ReportConfig Report { get; set; } = new();
    public EmailConfig Email { get; set; } = new();
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
    public string CloudId { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string BaseUrl { get; set; } = string.Empty;
    public string UserEmail { get; set; } = string.Empty;
    public string ApiToken { get; set; } = string.Empty;
    public string BudgetFieldName { get; set; } = "Budget";
    public string AccountFieldName { get; set; } = "Account";
}

public class ReportConfig
{
    public string? StartDate { get; set; }
    public string? EndDate { get; set; }
    public string OutputFile { get; set; } = "report.csv";
    public string Format { get; set; } = "csv";
}

public class EmailConfig
{
    public string SmtpHost { get; set; } = string.Empty;
    public int SmtpPort { get; set; } = 587;

    /// <summary>True = implicit SSL (typically port 465); false = STARTTLS (typically port 587).</summary>
    public bool UseSsl { get; set; }

    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string FromAddress { get; set; } = string.Empty;

    /// <summary>Semicolon-separated list of recipient addresses.</summary>
    public string Recipients { get; set; } = string.Empty;

    /// <summary>Supports {month} (Danish month name) and {year} placeholders.</summary>
    public string Subject { get; set; } = "Timerapport {month} {year}";

    /// <summary>Supports {month} (Danish month name) and {year} placeholders.</summary>
    public string Body { get; set; } =
        "Hej,\n\nVedhæftet timerapport for {month} {year}.\n\nMed venlig hilsen";
}
