namespace RDPilot.Persistence;

/// <summary>Owns the local writer identity and conservatively detects live peer writers.</summary>
internal static class RecoveryWriterSession
{
    static readonly object Gate = new();
    static readonly string OwnerSuffix = $"-{Environment.ProcessId}-{Guid.NewGuid():N}";
    static string currentId = CreateId(RuntimeClock.UtcNow);
    internal static string CurrentId => IdFor(null);

    internal static string IdFor(DateTime? compactedThrough, IEnumerable<string>? knownWriters = null)
    {
        lock (Gate)
        {
            var cutoff = compactedThrough ?? DateTime.MinValue;
            var existing = knownWriters?.Where(id => id.EndsWith(OwnerSuffix, StringComparison.Ordinal) &&
                    RecoveryWriterCounters.WriterStartedUtc(id) > cutoff)
                .OrderByDescending(RecoveryWriterCounters.WriterStartedUtc).FirstOrDefault();
            if (existing is not null) return existing;
            // A first write to a different counter must not inherit an old process-wide timestamp:
            // peers may already have compacted that counter without ever seeing this writer.
            if (knownWriters is not null || RecoveryWriterCounters.WriterStartedUtc(currentId) <= cutoff)
            {
                if (cutoff >= DateTime.MaxValue.AddTicks(-1))
                    throw new InvalidOperationException("Recovery compaction cutoff leaves no room for a new writer.");
                var now = RuntimeClock.UtcNow;
                currentId = CreateId(now > cutoff ? now : cutoff.AddTicks(1));
            }
            return currentId;
        }
    }

    static string CreateId(DateTime started) => $"{started.Ticks:x16}{OwnerSuffix}";

    internal static bool IsActive(string writerId)
    {
        var parts = writerId.Split('-');
        if (parts.Length != 3 || !Guid.TryParseExact(parts[2], "N", out _) ||
            !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var pid) || pid <= 0)
            return false; // Legacy counters have no process owner.
        try
        {
            using var process = Process.GetProcessById(pid);
            // PID reuse may postpone compaction, but must never cause a live writer to lose data.
            return !process.HasExited;
        }
        catch (ArgumentException) { return false; } // Process no longer exists.
        catch (InvalidOperationException) { return true; }
        catch (System.ComponentModel.Win32Exception) { return true; } // Cannot prove it is inactive.
    }
}
