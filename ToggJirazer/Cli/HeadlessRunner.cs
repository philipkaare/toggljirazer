using System.IO;
using ToggJirazer.Models;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Cli;

/// <summary>Non-interactive report run: validate, check the Atlassian sign-in, run, return an exit code.</summary>
public static class HeadlessRunner
{
    public const string NotSignedInMessage =
        "Not signed in to Atlassian. Open ToggJirazer and use 'Sign in with Atlassian' on the Jira tab.";

    public static Task<int> RunAsync(AppConfig config, AtlassianMcp mcp) =>
        RunAsync(config, () => mcp.CheckSignInAsync(CancellationToken.None), mcp.Tools, Console.Error);

    /// <summary>Seam for tests: the sign-in check, Jira tools and error writer are injected.</summary>
    internal static async Task<int> RunAsync(AppConfig config, Func<Task<SignInStatus>> checkSignIn, IJiraTools jiraTools, TextWriter stderr)
    {
        var error = ConfigValidator.Validate(config);
        if (error != null)
        {
            stderr.WriteLine($"[ERROR] {error}");
            return 2;
        }

        try
        {
            // Headless runs never open a browser; the user signs in from the GUI.
            var signIn = await checkSignIn();
            if (!signIn.IsSignedIn)
            {
                stderr.WriteLine(NotSignedInMessage);
                // A tool error isn't necessarily a sign-in problem; show what Atlassian said.
                if (signIn.ToolError != null)
                    stderr.WriteLine($"Details: {signIn.ToolError}");
                return 2;
            }

            await ReportRunner.RunAsync(config, jiraTools, Console.WriteLine);
            return 0;
        }
        catch (Exception ex)
        {
            stderr.WriteLine($"[ERROR] {ex.Message}");
            return 1;
        }
    }
}
