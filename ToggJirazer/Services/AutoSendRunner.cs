using System.IO;

namespace ToggJirazer.Services;

/// <summary>
/// Headless flow behind the --auto flag: decide whether this month's report is due,
/// generate it, email it, and advance the send-state. Logs to logs\auto.log.
/// On failure the state is NOT advanced, so the next daily run retries.
/// </summary>
public static class AutoSendRunner
{
    public static async Task<int> RunAsync(string baseDirectory)
    {
        var logDir = Path.Combine(baseDirectory, "logs");
        Directory.CreateDirectory(logDir);
        using var logFile = new StreamWriter(Path.Combine(logDir, "auto.log"), append: true) { AutoFlush = true };
        void Log(string line) => logFile.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {line}");

        // Route the services' Console progress output into the log file.
        var originalOut = Console.Out;
        var originalError = Console.Error;
        using var outWriter = new LogTextWriter(Log);
        using var errorWriter = new LogTextWriter(line => Log($"[ERROR] {line}"));
        Console.SetOut(outWriter);
        Console.SetError(errorWriter);

        try
        {
            var today = DateTime.Today;
            var stateStore = new AutomationStateStore(baseDirectory);
            var state = stateStore.Load();

            // Already sent this month's report — exit silently so the daily run doesn't bloat the log.
            if (!AutomationLogic.ShouldSend(state.LastSentPeriod, today))
                return 0;

            Log($"Starting automatic report run (last sent: {state.LastSentPeriod ?? "never"}).");

            var config = new SettingsService(Path.Combine(baseDirectory, "appsettings.json")).Load();

            var error = ReportRunner.ValidateConfig(config) ?? EmailService.ValidateEmailConfig(config.Email);
            if (error == null && !string.Equals(config.Report.Format?.Trim(), "xlsx", StringComparison.OrdinalIgnoreCase))
                error = "Automatic email sending requires the xlsx report format.";
            if (error != null)
            {
                Log($"[ERROR] {error}");
                return 1;
            }

            var (startDate, endDate) = AutomationLogic.ComputeReportPeriod(today);
            Log($"Generating report for {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}.");

            await ReportRunner.RunAsync(config, startDate, endDate,
                cacheEnabled: true, cacheWritePermissionPrompt: _ => false);

            var subject = AutomationLogic.ResolvePlaceholders(config.Email.Subject, endDate);
            var body = AutomationLogic.ResolvePlaceholders(config.Email.Body, endDate);
            await EmailService.SendReportAsync(config.Email, subject, body, config.Report.OutputFile);

            state.LastSentPeriod = AutomationLogic.GetTargetPeriod(today);
            stateStore.Save(state);
            Log($"Report emailed to {string.Join(", ", EmailService.ParseRecipients(config.Email.Recipients))}; state advanced to {state.LastSentPeriod}.");
            return 0;
        }
        catch (Exception ex)
        {
            Log($"[ERROR] {ex}");
            return 1;
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }
    }
}
