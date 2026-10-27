using Avalonia.Controls;
using ToggJirazer.Services.Atlassian;

namespace ToggJirazer.Setup.Steps;

/// <summary>
/// Wizard step 4: Atlassian sign-in and site. Sign-in and sign-out take effect immediately (session state);
/// a browser opens only from the sign-in button.
/// </summary>
public partial class JiraSignInStep : UserControl
{
    private WizardContext _ctx = null!;

    public JiraSignInStep()
    {
        InitializeComponent();
    }

    internal void Attach(WizardContext context, AtlassianMcp? mcp)
    {
        _ctx = context;
        if (mcp != null) JiraSignIn.Initialize(mcp);
        JiraSignIn.SetSelection(_ctx.Draft.Jira.CloudId, _ctx.Draft.Jira.SiteUrl);
        JiraSignIn.SignInChanged += (_, _) => SyncFromControl();
        JiraSignIn.SiteComboBox.SelectionChanged += (_, _) => SyncFromControl();
        // While the Atlassian sign-in waits for the browser, the wizard can't be left or closed with its buttons.
        JiraSignIn.SigningInChanged += (_, _) => _ctx.UpdateButtons();
    }

    /// <summary>True while an interactive sign-in waits for the user to finish in the browser.</summary>
    internal bool IsSigningIn => JiraSignIn.IsSigningIn;

    /// <summary>Cancels a waiting Atlassian sign-in, e.g. when the wizard closes.</summary>
    internal void CancelSignIn() => JiraSignIn.CancelSignIn();

    internal async Task EnterAsync()
    {
        // Never prompts or opens a browser; a browser opens only from the control's sign-in button.
        await JiraSignIn.RefreshStatusAsync();
        SyncFromControl();
    }

    private void SyncFromControl()
    {
        _ctx.State.JiraSignedIn = JiraSignIn.IsSignedIn;
        _ctx.Draft.Jira.CloudId = JiraSignIn.CloudId;
        _ctx.Draft.Jira.SiteUrl = JiraSignIn.SiteUrl;
        _ctx.UpdateButtons();
    }
}
