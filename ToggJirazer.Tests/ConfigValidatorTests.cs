using Xunit;
using ToggJirazer.Cli;
using ToggJirazer.Models;

namespace ToggJirazer.Tests;

public class ConfigValidatorTests
{
    private static AppConfig Valid() => new()
    {
        Toggl = new TogglConfig { ApiToken = "t", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
        Jira = new JiraConfig { BaseUrl = "https://x.atlassian.net", UserEmail = "a@b.c", ApiToken = "j" },
        Report = new ReportConfig { OutputFile = "r.xlsx", Format = "xlsx" }
    };

    [Fact]
    public void Valid_config_returns_null()
        => Assert.Null(ConfigValidator.Validate(Valid()));

    [Fact]
    public void Missing_toggl_token_is_reported()
    {
        var c = Valid();
        c.Toggl.ApiToken = "";
        Assert.Contains("Toggl API Token", ConfigValidator.Validate(c));
    }

    [Fact]
    public void Zero_workspace_id_is_reported()
    {
        var c = Valid();
        c.Toggl.WorkspaceId = 0;
        Assert.Contains("Workspace ID", ConfigValidator.Validate(c));
    }

    [Fact]
    public void Missing_jira_base_url_is_reported()
    {
        var c = Valid();
        c.Jira.BaseUrl = "";
        Assert.Contains("Jira Base URL", ConfigValidator.Validate(c));
    }

    [Fact]
    public void Missing_output_file_is_reported()
    {
        var c = Valid();
        c.Report.OutputFile = "";
        Assert.Contains("Output File", ConfigValidator.Validate(c));
    }
}
