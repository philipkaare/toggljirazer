namespace ToggJirazer.Models;

public class TogglTimeEntry
{
    public long Id { get; set; }
    public string Description { get; set; } = string.Empty;
    public string User { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public long Duration { get; set; }
    public string Project { get; set; } = string.Empty;
    public long ProjectId { get; set; }
}

public class TogglDetailedReport
{
    public List<TogglTimeEntry> Data { get; set; } = new();
    public int TotalCount { get; set; }
    public int PerPage { get; set; }
}
