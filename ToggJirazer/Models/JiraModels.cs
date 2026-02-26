namespace ToggJirazer.Models;

public class JiraIssue
{
    public string Key { get; set; } = string.Empty;
    public string IssueType { get; set; } = string.Empty;
    public string Summary { get; set; } = string.Empty;
    public string? Budget { get; set; }
    public string? Account { get; set; }
    public List<string> FixVersions { get; set; } = new();
    public double? Estimate { get; set; } // in hours
}
