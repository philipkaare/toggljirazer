using ToggJirazer.Services;
using Xunit;

namespace ToggJirazer.Tests;

public class SettingsLocatorTests : IDisposable
{
    private const string TemplateJson = """{"Toggl":{"ApiToken":"your-toggl-api-token"}}""";
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"tj-loc-{Guid.NewGuid():N}");

    public SettingsLocatorTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    private string WriteFile(string relativePath, string json, DateTime? writeTimeUtc = null)
    {
        var path = Path.Combine(_root, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, json);
        if (writeTimeUtc != null) File.SetLastWriteTimeUtc(path, writeTimeUtc.Value);
        return path;
    }

    private static string Token(string token) => $$$"""{"Toggl":{"ApiToken":"{{{token}}}"}}""";

    [Fact]
    public void EnsureUserSettings_does_nothing_when_user_file_exists()
    {
        var userPath = WriteFile("user/appsettings.json", Token("mine"));
        var candidate = WriteFile("a/appsettings.json", Token("other"));
        var template = WriteFile("template.json", TemplateJson);

        var result = SettingsLocator.EnsureUserSettings(userPath, [candidate], template, _ => { });

        Assert.Null(result);
        Assert.Equal(Token("mine"), File.ReadAllText(userPath));
    }

    [Fact]
    public void EnsureUserSettings_copies_newest_qualifying_candidate()
    {
        var t = new DateTime(2026, 10, 1, 12, 0, 0, DateTimeKind.Utc);
        var a = WriteFile("a/appsettings.json", Token("old"), t);
        var b = WriteFile("b/appsettings.json", Token("new"), t.AddMinutes(1));
        var c = WriteFile("c/appsettings.json", TemplateJson, t.AddMinutes(2));
        var template = WriteFile("template.json", TemplateJson);
        var userPath = Path.Combine(_root, "user", "appsettings.json");

        var result = SettingsLocator.EnsureUserSettings(userPath, [a, b, c], template, _ => { });

        Assert.Equal(b, result);
        Assert.Equal("new", new SettingsService(userPath).Load().Toggl.ApiToken);
    }

    [Fact]
    public void EnsureUserSettings_copies_merged_dev_values_from_candidate()
    {
        var candidate = WriteFile("repo/appsettings.json", TemplateJson);
        WriteFile("repo/appsettings.development.json", """{"Toggl":{"ApiToken":"real","WorkspaceId":42}}""");
        var template = WriteFile("template.json", TemplateJson);
        var userPath = Path.Combine(_root, "user", "appsettings.json");

        SettingsLocator.EnsureUserSettings(userPath, [candidate], template, _ => { });

        var loaded = new SettingsService(userPath).Load();
        Assert.Equal("real", loaded.Toggl.ApiToken);
        Assert.Equal(42, loaded.Toggl.WorkspaceId);
    }

    [Fact]
    public void EnsureUserSettings_falls_back_to_template()
    {
        var candidate = WriteFile("a/appsettings.json", TemplateJson);
        var template = WriteFile("template.json", TemplateJson);
        var userPath = Path.Combine(_root, "user", "appsettings.json");

        var result = SettingsLocator.EnsureUserSettings(userPath, [candidate, Path.Combine(_root, "missing.json")], template, _ => { });

        Assert.Equal(template, result);
        Assert.Equal(TemplateJson, File.ReadAllText(userPath));
    }

    [Fact]
    public void EnsureUserSettings_creates_missing_directory()
    {
        var template = WriteFile("template.json", TemplateJson);
        var userPath = Path.Combine(_root, "deep", "new", "appsettings.json");

        SettingsLocator.EnsureUserSettings(userPath, [], template, _ => { });

        Assert.True(File.Exists(userPath));
    }

    [Fact]
    public void EnsureUserSettings_logs_source_and_destination()
    {
        var template = WriteFile("template.json", TemplateJson);
        var userPath = Path.Combine(_root, "user", "appsettings.json");
        var lines = new List<string>();

        SettingsLocator.EnsureUserSettings(userPath, [], template, lines.Add);

        Assert.Equal($"Settings copied from {template} to {userPath}", Assert.Single(lines));
    }

    [Fact]
    public void LegacyCandidates_includes_windows_build_output()
    {
        var windowsFile = WriteFile(Path.Combine("bin", "Debug", "net10.0-windows", "appsettings.json"), Token("x"));
        var baseDirectory = Path.Combine(_root, "bin", "Debug", "net10.0") + Path.DirectorySeparatorChar;
        Directory.CreateDirectory(baseDirectory);

        var candidates = SettingsLocator.LegacyCandidates(_root, baseDirectory).Select(Path.GetFullPath);

        Assert.Contains(Path.GetFullPath(windowsFile), candidates);
        Assert.Contains(Path.GetFullPath(Path.Combine(_root, "appsettings.json")), candidates);
        Assert.Contains(Path.GetFullPath(Path.Combine(baseDirectory, "appsettings.json")), candidates);
    }

    [Fact]
    public void UserSettingsPath_ends_in_ToggJirazer_folder()
    {
        var path = SettingsLocator.UserSettingsPath();
        Assert.Equal("appsettings.json", Path.GetFileName(path));
        Assert.Equal("ToggJirazer", Path.GetFileName(Path.GetDirectoryName(path)));
    }

    [Fact]
    public void Settings_dir_env_override()
    {
        var dir = Path.Combine(_root, "override");
        try
        {
            Environment.SetEnvironmentVariable("TOGGJIRAZER_SETTINGS_DIR", dir);
            Assert.Equal(Path.Combine(dir, "appsettings.json"), SettingsLocator.UserSettingsPath());

            Environment.SetEnvironmentVariable("TOGGJIRAZER_SETTINGS_DIR", "");
            Assert.Equal("ToggJirazer", Path.GetFileName(Path.GetDirectoryName(SettingsLocator.UserSettingsPath())));
        }
        finally
        {
            Environment.SetEnvironmentVariable("TOGGJIRAZER_SETTINGS_DIR", null);
        }
    }

    [Fact]
    public void AtlassianAuthPath_sits_next_to_settings()
    {
        var settings = Path.Combine(_root, "user", "appsettings.json");
        Assert.Equal(Path.Combine(_root, "user", "atlassian-auth.json"), SettingsLocator.AtlassianAuthPath(settings));
    }
}
