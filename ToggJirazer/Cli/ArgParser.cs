namespace ToggJirazer.Cli;

public enum LaunchMode { Tui, Headless, Help, Auto, Gui }

public static class ArgParser
{
    public static LaunchMode Resolve(string[] args, bool inputRedirected, bool outputRedirected, bool windowsGuiAvailable = false)
    {
        foreach (var a in args)
        {
            if (a is "--help" or "-h") return LaunchMode.Help;
        }
        if (args.Any(a => a.Equals("--auto", StringComparison.OrdinalIgnoreCase)))
            return LaunchMode.Auto;
        foreach (var a in args)
        {
            if (a is "--headless" or "--run") return LaunchMode.Headless;
        }
        if (inputRedirected || outputRedirected) return LaunchMode.Headless;
        if (args.Contains("--tui")) return LaunchMode.Tui;
        return windowsGuiAvailable ? LaunchMode.Gui : LaunchMode.Headless;
    }
}
