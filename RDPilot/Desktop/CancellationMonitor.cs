namespace RDPilot.Desktop;

/// <summary>Owns the hotkey worker until it has stopped, then releases its cancellation source.</summary>
internal sealed class CancellationMonitor : IAsyncDisposable
{
    readonly CancellationTokenSource source = new();
    readonly Task worker;
    internal CancellationToken Token { get; }

    internal CancellationMonitor(RunContext run, Func<bool> isAbortPressed)
    {
        run.CancelRequested = false;
        Token = source.Token;
        worker = Task.Run(async () =>
        {
            try
            {
                while (!Token.IsCancellationRequested)
                {
                    if (isAbortPressed())
                    {
                        run.CancelRequested = true;
                        source.Cancel();
                        break;
                    }
                    await RuntimeClock.Delay(50, Token).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException) when (Token.IsCancellationRequested) { }
        });
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            await source.CancelAsync().ConfigureAwait(false);
            await worker.ConfigureAwait(false);
        }
        finally { source.Dispose(); }
    }
}
