using Xunit;
using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Tests;

public class BudgetInheritanceTests
{
    private const string NoBudget = "(Intet budget)";

    private static Dictionary<string, string> Empty() =>
        new(StringComparer.OrdinalIgnoreCase);

    [Fact]
    public void Issue_without_own_budget_inherits_parent_budget()
    {
        var issue = new JiraIssue { Key = "FOND-1776", IssueType = "Defect", ParentKey = "FOND-2045" };
        var inherited = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-1776"] = "Fondsdrift"
        };

        var budget = ReportService.ResolveIssueBudget(issue, Empty(), inherited, NoBudget);

        Assert.Equal("Fondsdrift", budget);
    }

    [Fact]
    public void Fix_version_budget_wins_over_inherited_parent_budget()
    {
        var issue = new JiraIssue
        {
            Key = "FOND-1776",
            IssueType = "Defect",
            ParentKey = "FOND-2045",
            FixVersions = ["2026.4"]
        };
        var versionToBudget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["2026.4"] = "Leverance A"
        };
        var inherited = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-1776"] = "Fondsdrift"
        };

        var budget = ReportService.ResolveIssueBudget(issue, versionToBudget, inherited, NoBudget);

        Assert.Equal("Leverance A", budget);
    }

    [Fact]
    public void Own_budget_wins_over_inherited_parent_budget()
    {
        var issue = new JiraIssue
        {
            Key = "FOND-1776",
            IssueType = "Defect",
            ParentKey = "FOND-2045",
            Budget = "Eget budget"
        };
        var inherited = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-1776"] = "Fondsdrift"
        };

        var budget = ReportService.ResolveIssueBudget(issue, Empty(), inherited, NoBudget);

        Assert.Equal("Eget budget", budget);
    }

    [Fact]
    public void Issue_with_no_budget_anywhere_falls_back_to_missing_marker()
    {
        var issue = new JiraIssue { Key = "FOND-1776", IssueType = "Defect", ParentKey = "FOND-2045" };

        var budget = ReportService.ResolveIssueBudget(issue, Empty(), Empty(), NoBudget);

        Assert.Equal(NoBudget, budget);
    }

    [Fact]
    public void Nearest_ancestor_budget_is_found_on_direct_parent()
    {
        var child = new JiraIssue { Key = "FOND-1776", ParentKey = "FOND-2045" };
        var issues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-2045"] = new JiraIssue { Key = "FOND-2045", Budget = "Fondsdrift" }
        };

        Assert.Equal("Fondsdrift", JiraService.NearestAncestorBudget(child, issues));
    }

    [Fact]
    public void Nearest_ancestor_budget_skips_ancestors_without_budget()
    {
        var child = new JiraIssue { Key = "FOND-1776", ParentKey = "FOND-2045" };
        var issues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-2045"] = new JiraIssue { Key = "FOND-2045", Budget = "   ", ParentKey = "FOND-1000" },
            ["FOND-1000"] = new JiraIssue { Key = "FOND-1000", Budget = "Fondsdrift" }
        };

        Assert.Equal("Fondsdrift", JiraService.NearestAncestorBudget(child, issues));
    }

    [Fact]
    public void Nearest_ancestor_budget_prefers_closest_ancestor()
    {
        var child = new JiraIssue { Key = "FOND-1776", ParentKey = "FOND-2045" };
        var issues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-2045"] = new JiraIssue { Key = "FOND-2045", Budget = "Nærmest", ParentKey = "FOND-1000" },
            ["FOND-1000"] = new JiraIssue { Key = "FOND-1000", Budget = "Fjernest" }
        };

        Assert.Equal("Nærmest", JiraService.NearestAncestorBudget(child, issues));
    }

    [Fact]
    public void Circular_parent_references_return_no_budget_without_looping()
    {
        var child = new JiraIssue { Key = "FOND-1776", ParentKey = "FOND-2045" };
        var issues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase)
        {
            ["FOND-2045"] = new JiraIssue { Key = "FOND-2045", ParentKey = "FOND-1776" },
            ["FOND-1776"] = child
        };

        Assert.Null(JiraService.NearestAncestorBudget(child, issues));
    }

    [Fact]
    public void Unfetched_ancestor_yields_no_budget()
    {
        var child = new JiraIssue { Key = "FOND-1776", ParentKey = "FOND-2045" };
        var issues = new Dictionary<string, JiraIssue?>(StringComparer.OrdinalIgnoreCase);

        Assert.Null(JiraService.NearestAncestorBudget(child, issues));
    }
}
