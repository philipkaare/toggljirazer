using ToggJirazer.Models;
using ToggJirazer.Services;

namespace ToggJirazer.Setup.Steps;

/// <summary>What the wizard window shares with its step controls.</summary>
internal sealed class WizardContext(
    SetupWizardState state,
    Func<TogglSetupService> togglFactory,
    Action<string> openUrl,
    Action updateButtons)
{
    public SetupWizardState State { get; } = state;

    public AppConfig Draft => State.Draft;

    /// <summary>Creates a Toggl client for one request or check; the caller disposes it.</summary>
    public Func<TogglSetupService> TogglFactory { get; } = togglFactory;

    /// <summary>
    /// Bumped whenever the step changes or a list load starts; a load or check whose generation is stale
    /// is dropped. Shared by all steps so a step change invalidates every pending answer.
    /// </summary>
    public int LoadGeneration { get; private set; }

    public int NextLoadGeneration() => ++LoadGeneration;

    /// <summary>Opens a URL in the browser (the window's replaceable <c>OpenUrl</c>).</summary>
    public void OpenUrl(string url) => openUrl(url);

    /// <summary>Re-evaluates Back/Next/Cancel after a step changed the draft or its busy state.</summary>
    public void UpdateButtons() => updateButtons();
}
