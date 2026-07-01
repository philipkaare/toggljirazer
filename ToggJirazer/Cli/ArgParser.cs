namespace ToggJirazer.Cli;

public enum LaunchMode { Tui, Headless, Help }

public static class ArgParser
{
    public static LaunchMode Resolve(string[] args, bool inputRedirected, bool outputRedirected)
    {
        foreach (var a in args)
        {
            if (a is "--help" or "-h") return LaunchMode.Help;
        }
        foreach (var a in args)
        {
            if (a is "--headless" or "--run") return LaunchMode.Headless;
        }
        if (inputRedirected || outputRedirected) return LaunchMode.Headless;
        return LaunchMode.Tui;
    }
}
