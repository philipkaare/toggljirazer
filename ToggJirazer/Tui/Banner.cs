using System.Text;

namespace ToggJirazer.Tui;

/// <summary>Static ASCII masthead shown at the top of the TUI.</summary>
public static class Banner
{
    public static string Text(string version)
    {
        var sb = new StringBuilder();
        sb.AppendLine(@" _____                 _ _ _                     ");
        sb.AppendLine(@"|_   _|__  __ _ __ _   | (_) |_ __ _ _______ _ __ ");
        sb.AppendLine(@"  | |/ _ \/ _` / _` |  | | | __/ _` |_  / _ \ '__|");
        sb.AppendLine(@"  | | (_) \__, \__, |  | | | || (_| |/ /  __/ |   ");
        sb.AppendLine(@"  |_|\___/|___/|___/   |_|_|\__\__,_/___\___|_|   ");
        sb.AppendLine($"ToggJirazer v{version}  ·  Toggl × Jira time reports");
        return sb.ToString();
    }
}
