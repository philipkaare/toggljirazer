namespace ToggJirazer.Tests.Fakes;

/// <summary>
/// Answers requests from a responder, but holds back those matching <see cref="Hold"/> until the test
/// calls <see cref="ReleaseAll"/>. One instance can back many handlers (one per TogglSetupService).
/// </summary>
public sealed class HeldHttpHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
{
    private readonly List<(HttpRequestMessage Request, TaskCompletionSource<HttpResponseMessage> Answer)> _pending = [];

    public Func<HttpRequestMessage, bool> Hold { get; set; } = _ => false;

    public Func<HttpRequestMessage, HttpResponseMessage> Responder { get; set; } = responder;

    public int PendingCount => _pending.Count;

    public HttpMessageHandler NewHandler() => new Handler(this);

    /// <summary>Answers every held request with the current <see cref="Responder"/>.</summary>
    public void ReleaseAll()
    {
        var pending = _pending.ToList();
        _pending.Clear();
        foreach (var (request, answer) in pending) answer.SetResult(Responder(request));
    }

    private sealed class Handler(HeldHttpHandler owner) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            if (!owner.Hold(request)) return Task.FromResult(owner.Responder(request));

            var answer = new TaskCompletionSource<HttpResponseMessage>(TaskCreationOptions.RunContinuationsAsynchronously);
            owner._pending.Add((request, answer));
            return answer.Task;
        }
    }
}
