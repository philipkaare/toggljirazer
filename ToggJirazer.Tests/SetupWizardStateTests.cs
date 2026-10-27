using System.IO;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Setup;
using Xunit;

namespace ToggJirazer.Tests;

public class SetupWizardStateTests
{
    private static AppConfig Complete() => new()
    {
        Toggl = new TogglConfig { ApiToken = "tok", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
        Jira = new JiraConfig { CloudId = "cloud", SiteUrl = "https://x.atlassian.net" }
    };

    [Fact]
    public void IsSetupIncomplete_complete_config_with_token_is_false() =>
        Assert.False(SetupWizardState.IsSetupIncomplete(Complete(), true));

    [Fact]
    public void IsSetupIncomplete_without_atlassian_token_is_true() =>
        Assert.True(SetupWizardState.IsSetupIncomplete(Complete(), false));

    [Theory]
    [InlineData("")]
    [InlineData("your-toggl-api-token")]
    public void IsSetupIncomplete_missing_or_placeholder_token_is_true(string token)
    {
        var c = Complete();
        c.Toggl.ApiToken = token;
        Assert.True(SetupWizardState.IsSetupIncomplete(c, true));
    }

    [Theory]
    [InlineData(0, 2, 3)]
    [InlineData(1, 0, 3)]
    [InlineData(1, 2, 0)]
    public void IsSetupIncomplete_any_zero_id_is_true(long org, long ws, long project)
    {
        var c = Complete();
        c.Toggl.OrganizationId = org;
        c.Toggl.WorkspaceId = ws;
        c.Toggl.ProjectId = project;
        Assert.True(SetupWizardState.IsSetupIncomplete(c, true));
    }

    [Fact]
    public void IsSetupIncomplete_empty_cloud_id_is_true()
    {
        var c = Complete();
        c.Jira.CloudId = "";
        Assert.True(SetupWizardState.IsSetupIncomplete(c, true));
    }

    [Fact]
    public void Next_disabled_until_toggl_token_present()
    {
        var s = new SetupWizardState(new AppConfig());
        Assert.True(s.CanGoNext);
        s.Next();
        Assert.Equal(SetupStep.TogglSignIn, s.Step);
        Assert.False(s.CanGoNext);
        Assert.Throws<InvalidOperationException>(() => s.Next());

        s.Draft.Toggl.ApiToken = "your-toggl-api-token";
        Assert.False(s.CanGoNext);

        s.Draft.Toggl.ApiToken = "tok";
        Assert.True(s.CanGoNext);
        s.Next();
        Assert.Equal(SetupStep.TogglProject, s.Step);
    }

    [Fact]
    public void Project_step_requires_all_three_ids()
    {
        var s = new SetupWizardState(new AppConfig { Toggl = new TogglConfig { ApiToken = "tok" } });
        s.Next();
        s.Next();
        Assert.Equal(SetupStep.TogglProject, s.Step);
        Assert.False(s.CanGoNext);
        s.Draft.Toggl.OrganizationId = 1;
        Assert.False(s.CanGoNext);
        s.Draft.Toggl.WorkspaceId = 2;
        Assert.False(s.CanGoNext);
        s.Draft.Toggl.ProjectId = 3;
        Assert.True(s.CanGoNext);
    }

    [Fact]
    public void Jira_step_requires_signed_in_and_site()
    {
        var s = new SetupWizardState(Complete());
        s.Next(); s.Next(); s.Next();
        Assert.Equal(SetupStep.JiraSignIn, s.Step);
        Assert.False(s.CanGoNext);
        s.JiraSignedIn = true;
        Assert.True(s.CanGoNext);
        s.Draft.Jira.CloudId = "";
        Assert.False(s.CanGoNext);
        s.Draft.Jira.CloudId = "cloud";
        s.JiraSignedIn = false;
        Assert.False(s.CanGoNext);
    }

    [Fact]
    public void Back_from_toggl_returns_to_welcome()
    {
        var s = new SetupWizardState(new AppConfig());
        Assert.False(s.CanGoBack);
        s.Next();
        Assert.True(s.CanGoBack);
        s.Back();
        Assert.Equal(SetupStep.Welcome, s.Step);
        Assert.False(s.CanGoBack);
    }

    [Fact]
    public void Starting_config_not_mutated()
    {
        var start = Complete();
        var s = new SetupWizardState(start);
        s.Draft.Toggl.ApiToken = "other";
        s.Draft.Toggl.ProjectId = 99;
        s.Draft.Jira.CloudId = "other";
        s.Draft.Report.Format = "xlsx";
        s.Draft.Email.SmtpHost = "smtp";
        Assert.Equal("tok", start.Toggl.ApiToken);
        Assert.Equal(3, start.Toggl.ProjectId);
        Assert.Equal("cloud", start.Jira.CloudId);
        Assert.Equal("csv", start.Report.Format);
        Assert.Equal("", start.Email.SmtpHost);
    }

    [Fact]
    public void Field_step_keeps_current_name()
    {
        var c = Complete();
        c.Jira.AccountFieldName = "Konto";
        var s = new SetupWizardState(c);
        s.JiraSignedIn = true;
        s.Next(); s.Next(); s.Next(); s.Next();
        Assert.Equal(SetupStep.JiraFields, s.Step);
        Assert.True(s.CanGoNext);
        s.Next();
        Assert.Equal(SetupStep.ReportOutput, s.Step);
        Assert.True(s.CanGoNext);
        s.Next();
        Assert.Equal(SetupStep.Done, s.Step);
        Assert.False(s.CanGoNext);
        Assert.Equal("Konto", s.Draft.Jira.AccountFieldName);
    }

    [Theory]
    [InlineData("", "csv", false)]
    [InlineData("r.txt", "pdf", false)]
    [InlineData("r.xlsx", "xlsx", true)]
    [InlineData("r.csv", "csv", true)]
    public void Report_step_requires_file_and_format(string file, string format, bool expected)
    {
        var c = Complete();
        c.Report.OutputFile = file;
        c.Report.Format = format;
        var s = new SetupWizardState(c);
        s.JiraSignedIn = true;
        for (var i = 0; i < 5; i++) s.Next();
        Assert.Equal(SetupStep.ReportOutput, s.Step);
        Assert.Equal(expected, s.CanGoNext);
    }

    [Fact]
    public void Cancel_leaves_settings_file_unchanged()
    {
        var dir = Path.Combine(Path.GetTempPath(), "toggjirazer-wiz-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "appsettings.json");
            var settings = new SettingsService(path);
            settings.Save(Complete());
            var before = File.ReadAllBytes(path);

            var s = new SetupWizardState(settings.Load());
            s.Draft.Toggl.ApiToken = "changed";
            s.Draft.Toggl.ProjectId = 42;
            s.Draft.Jira.AccountFieldName = "Other";
            s.Draft.Report.OutputFile = "elsewhere.xlsx";

            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
