using ToggJirazer.Models;

namespace ToggJirazer.Setup;

public enum SetupStep { Welcome, TogglSignIn, TogglProject, JiraSignIn, JiraFields, ReportOutput, Done }

/// <summary>
/// UI-free state of the setup wizard: a working copy of the settings and the current step.
/// The window owns saving; this class never touches the settings file.
/// </summary>
public sealed class SetupWizardState
{
    internal const string TokenPlaceholder = "your-toggl-api-token";

    public SetupWizardState(AppConfig startingConfig)
    {
        Draft = Copy(startingConfig);
    }

    public AppConfig Draft { get; }
    public SetupStep Step { get; private set; } = SetupStep.Welcome;

    /// <summary>Set by the window from the Atlassian sign-in control.</summary>
    public bool JiraSignedIn { get; set; }

    public bool CanGoBack => Step != SetupStep.Welcome;

    public bool CanGoNext => Step switch
    {
        SetupStep.Welcome => true,
        SetupStep.TogglSignIn => HasTogglToken(Draft),
        SetupStep.TogglProject => Draft.Toggl.OrganizationId != 0
            && Draft.Toggl.WorkspaceId != 0
            && Draft.Toggl.ProjectId != 0,
        SetupStep.JiraSignIn => JiraSignedIn && Draft.Jira.CloudId.Length > 0,
        SetupStep.JiraFields => true,
        SetupStep.ReportOutput => Draft.Report.OutputFile.Length > 0
            && Draft.Report.Format is "xlsx" or "csv",
        _ => false
    };

    public void Next()
    {
        if (!CanGoNext)
            throw new InvalidOperationException($"Cannot go next from step {Step}.");
        Step++;
    }

    public void Back()
    {
        if (CanGoBack) Step--;
    }

    public static bool IsSetupIncomplete(AppConfig config, bool hasAtlassianToken) =>
        !HasTogglToken(config)
        || config.Toggl.OrganizationId == 0
        || config.Toggl.WorkspaceId == 0
        || config.Toggl.ProjectId == 0
        || config.Jira.CloudId.Length == 0
        || !hasAtlassianToken;

    private static bool HasTogglToken(AppConfig config) =>
        config.Toggl.ApiToken.Length > 0 && config.Toggl.ApiToken != TokenPlaceholder;

    private static AppConfig Copy(AppConfig c) => new()
    {
        Toggl = new TogglConfig
        {
            ApiToken = c.Toggl.ApiToken,
            OrganizationId = c.Toggl.OrganizationId,
            WorkspaceId = c.Toggl.WorkspaceId,
            ProjectId = c.Toggl.ProjectId
        },
        Jira = new JiraConfig
        {
            CloudId = c.Jira.CloudId,
            SiteUrl = c.Jira.SiteUrl,
            BudgetFieldName = c.Jira.BudgetFieldName,
            AccountFieldName = c.Jira.AccountFieldName
        },
        Report = new ReportConfig
        {
            StartDate = c.Report.StartDate,
            EndDate = c.Report.EndDate,
            OutputFile = c.Report.OutputFile,
            Format = c.Report.Format
        },
        Email = new EmailConfig
        {
            SmtpHost = c.Email.SmtpHost,
            SmtpPort = c.Email.SmtpPort,
            UseSsl = c.Email.UseSsl,
            Username = c.Email.Username,
            Password = c.Email.Password,
            FromAddress = c.Email.FromAddress,
            Recipients = c.Email.Recipients,
            Subject = c.Email.Subject,
            Body = c.Email.Body
        }
    };
}
