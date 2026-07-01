using ToggJirazer.Models;
using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class SettingsServiceTests
{
    [Fact]
    public void Save_then_Load_round_trips_config()
    {
        var path = Path.Combine(Path.GetTempPath(), $"tj-{Guid.NewGuid():N}.json");
        try
        {
            var service = new SettingsService(path);
            var original = new AppConfig
            {
                Toggl = new TogglConfig { ApiToken = "tok", OrganizationId = 1, WorkspaceId = 2, ProjectId = 3 },
                Jira = new JiraConfig { BaseUrl = "https://x.atlassian.net", UserEmail = "a@b.c", ApiToken = "jt" },
                Report = new ReportConfig { StartDate = "2026-01-01", EndDate = "2026-06-30", OutputFile = "r.xlsx", Format = "xlsx" }
            };

            service.Save(original);
            var loaded = service.Load();

            Assert.Equal(original.Toggl.ApiToken, loaded.Toggl.ApiToken);
            Assert.Equal(original.Toggl.WorkspaceId, loaded.Toggl.WorkspaceId);
            Assert.Equal(original.Jira.BaseUrl, loaded.Jira.BaseUrl);
            Assert.Equal(original.Report.OutputFile, loaded.Report.OutputFile);
            Assert.Equal(original.Report.StartDate, loaded.Report.StartDate);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
