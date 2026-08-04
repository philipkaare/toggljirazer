using System.IO;
using System.Text.Json;

namespace ToggJirazer.Services;

/// <summary>State persisted between automatic runs (which period was last emailed).</summary>
public class AutomationState
{
    /// <summary>Last report period sent, as "yyyy-MM" (the previous month at send time).</summary>
    public string? LastSentPeriod { get; set; }
}

/// <summary>
/// Reads/writes automation-state.json. Kept separate from appsettings.json so the
/// settings UI can never overwrite the send-state.
/// </summary>
public class AutomationStateStore
{
    private readonly string _path;

    public AutomationStateStore(string directory)
        => _path = Path.Combine(directory, "automation-state.json");

    public AutomationState Load()
    {
        if (!File.Exists(_path)) return new AutomationState();
        try
        {
            return JsonSerializer.Deserialize<AutomationState>(File.ReadAllText(_path))
                   ?? new AutomationState();
        }
        catch
        {
            return new AutomationState();
        }
    }

    public void Save(AutomationState state)
    {
        var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_path, json);
    }
}
