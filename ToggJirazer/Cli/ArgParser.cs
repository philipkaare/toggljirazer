namespace ToggJirazer.Cli;

public enum LaunchMode { Gui, Headless, Help, AutoRemoved }

public static class ArgParser
{
    public static LaunchMode Resolve(string[] args, bool inputRedirected, bool outputRedirected)
    {
        foreach (var a in args)
        {
            if (a is "--help" or "-h") return LaunchMode.Help;
        }
        if (args.Any(a => a.Equals("--auto", StringComparison.OrdinalIgnoreCase)))
            return LaunchMode.AutoRemoved;
        foreach (var a in args)
        {
            if (a is "--headless" or "--run") return LaunchMode.Headless;
        }
        if (inputRedirected || outputRedirected) return LaunchMode.Headless;
        return LaunchMode.Gui;
    }
}
