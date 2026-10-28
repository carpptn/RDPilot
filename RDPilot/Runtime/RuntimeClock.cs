
namespace RDPilot.Runtime;

/// <summary>Time boundary shared by delays, elapsed-time budgets and persistence timestamps.</summary>
internal static class RuntimeClock
{
    internal static TimeProvider Provider { get; set; } = TimeProvider.System;
    internal static DateTime UtcNow => Provider.GetUtcNow().UtcDateTime;
    internal static Task Delay(int milliseconds, CancellationToken token = default) => Delay(TimeSpan.FromMilliseconds(milliseconds), token);
    internal static Task Delay(TimeSpan delay, CancellationToken token = default) => Task.Delay(delay, Provider, token);
    internal static void Sleep(int milliseconds) => Delay(milliseconds).GetAwaiter().GetResult();
    internal static ElapsedTimer StartTimer() => new(Provider);

    internal sealed class ElapsedTimer(TimeProvider provider)
    {
        readonly long started = provider.GetTimestamp();
        TimeSpan? stopped;
        internal TimeSpan Elapsed => stopped ?? provider.GetElapsedTime(started);
        internal long ElapsedMilliseconds => (long)Elapsed.TotalMilliseconds;
        internal void Stop() => stopped ??= Elapsed;
    }
}
