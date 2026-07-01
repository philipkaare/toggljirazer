using ToggJirazer.Models;

namespace ToggJirazer.Cli;

/// <summary>Non-interactive report run: validate, run, return an exit code.</summary>
public static class HeadlessRunner
{
    public static async Task<int> RunAsync(AppConfig config)
    {
        var error = ConfigValidator.Validate(config);
        if (error != null)
        {
            Console.Error.WriteLine($"[ERROR] {error}");
            return 2;
        }

        try
        {
            await ReportRunner.RunAsync(config, Console.WriteLine);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[ERROR] {ex.Message}");
            return 1;
        }
    }
}
