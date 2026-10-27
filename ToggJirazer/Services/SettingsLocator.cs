using System.IO;

namespace ToggJirazer.Services;

/// <summary>
/// Locates the per-user appsettings.json and creates it on first launch from the
/// newest legacy settings file that holds real values (falling back to the template).
/// </summary>
public static class SettingsLocator
{
    private const string FileName = "appsettings.json";
    private const string PlaceholderToken = "your-toggl-api-token";
    private const string SettingsDirVariable = "TOGGJIRAZER_SETTINGS_DIR";

    /// <summary>
    /// %APPDATA%\ToggJirazer\appsettings.json on Windows,
    /// ~/Library/Application Support/ToggJirazer/appsettings.json on macOS.
    /// The TOGGJIRAZER_SETTINGS_DIR environment variable overrides the folder (a test seam,
    /// so the app can be launched against throwaway settings).
    /// </summary>
    public static string UserSettingsPath()
    {
        var overrideDir = Environment.GetEnvironmentVariable(SettingsDirVariable);
        if (!string.IsNullOrEmpty(overrideDir))
            return Path.Combine(overrideDir, FileName);

        var root = OperatingSystem.IsMacOS()
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Application Support")
            : Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        return Path.Combine(root, "ToggJirazer", FileName);
    }

    /// <summary>The Atlassian sign-in file, kept next to the settings file.</summary>
    public static string AtlassianAuthPath(string settingsPath) =>
        Path.Combine(Path.GetDirectoryName(settingsPath) ?? string.Empty, "atlassian-auth.json");

    /// <summary>
    /// Settings files used before settings moved per-user: the working directory, the
    /// executable directory, and sibling Windows build outputs (e.g. bin/Debug/net10.0-windows).
    /// </summary>
    public static IEnumerable<string> LegacyCandidates(string currentDirectory, string baseDirectory)
    {
        yield return Path.Combine(currentDirectory, FileName);
        yield return Path.Combine(baseDirectory, FileName);

        var buildConfigDir = Directory.GetParent(Path.TrimEndingDirectorySeparator(baseDirectory))?.FullName;
        if (buildConfigDir == null || !Directory.Exists(buildConfigDir)) yield break;
        foreach (var dir in Directory.EnumerateDirectories(buildConfigDir, "*-windows"))
            yield return Path.Combine(dir, FileName);
    }

    /// <summary>
    /// Creates <paramref name="userPath"/> when it does not exist yet. Returns the source
    /// file used, or null when the user file already existed.
    /// </summary>
    public static string? EnsureUserSettings(string userPath, IEnumerable<string> candidates,
        string templatePath, Action<string> log)
    {
        if (File.Exists(userPath)) return null;
        Directory.CreateDirectory(Path.GetDirectoryName(userPath)!);

        var source = candidates
            .Where(File.Exists)
            .Select(path => (Path: path, Config: new SettingsService(path).Load()))
            .Where(c => !string.IsNullOrWhiteSpace(c.Config.Toggl.ApiToken) && c.Config.Toggl.ApiToken != PlaceholderToken)
            .OrderByDescending(c => LatestWriteTimeUtc(c.Path))
            .FirstOrDefault();

        if (source.Path != null)
        {
            new SettingsService(userPath).Save(source.Config);
            log($"Settings copied from {source.Path} to {userPath}");
            return source.Path;
        }

        File.Copy(templatePath, userPath);
        log($"Settings copied from {templatePath} to {userPath}");
        return templatePath;
    }

    private static DateTime LatestWriteTimeUtc(string path)
    {
        var time = File.GetLastWriteTimeUtc(path);
        var devPath = SettingsService.DevelopmentPathFor(path);
        return File.Exists(devPath) ? new[] { time, File.GetLastWriteTimeUtc(devPath) }.Max() : time;
    }
}
