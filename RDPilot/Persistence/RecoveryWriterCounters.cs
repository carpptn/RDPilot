namespace RDPilot.Persistence;

/// <summary>Shared merge and compaction rules for lesson and calibration writer counters.</summary>
internal static class RecoveryWriterCounters
{
    // An absent cutoff must never erase a cutoff already observed in another snapshot.
    internal static DateTime? LatestCompactionCutoff(DateTime? left, DateTime? right) =>
        LatestCompactionCutoff(new[] { left, right });

    internal static DateTime? LatestCompactionCutoff(IEnumerable<DateTime?> cutoffs) => cutoffs.Max();

    internal static Dictionary<string, int> MergeWriterCounters(
        IReadOnlyDictionary<string, int> left,
        IReadOnlyDictionary<string, int> right)
    {
        var merged = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var key in left.Keys.Concat(right.Keys).Distinct(StringComparer.Ordinal))
        {
            left.TryGetValue(key, out var leftValue);
            right.TryGetValue(key, out var rightValue);
            merged[key] = Math.Max(leftValue, rightValue);
        }
        return merged;
    }

    internal static Dictionary<string, double> MergeWriterDoubleCounters(
        IReadOnlyDictionary<string, double> left,
        IReadOnlyDictionary<string, double> right)
    {
        var merged = new Dictionary<string, double>(
            StringComparer.Ordinal);
        foreach (var key in left.Keys
                     .Concat(right.Keys)
                     .Distinct(StringComparer.Ordinal))
        {
            left.TryGetValue(key, out var leftValue);
            right.TryGetValue(key, out var rightValue);
            merged[key] = Math.Max(leftValue, rightValue);
        }
        return merged;
    }

    internal static DateTime? WriterCompactionCutoff(IEnumerable<string> writerIds)
    {
        const int maximumComponents = 64;
        const int componentsToKeep = 32;
        var writers = writerIds.Distinct(StringComparer.Ordinal)
            .Select(id => (Id: id, StartedUtc: WriterStartedUtc(id))).ToArray();
        // The cutoff removes every older component, so filtering out live IDs alone is insufficient.
        var oldestActive = writers.Where(item => RecoveryWriterSession.IsActive(item.Id))
            .Select(item => item.StartedUtc).DefaultIfEmpty(DateTime.MaxValue).Min();
        var parsed = writers
            .Where(item => item.StartedUtc < oldestActive && item.StartedUtc <= RuntimeClock.UtcNow.AddDays(-7))
            .OrderBy(item => item.StartedUtc)
            .ToArray();
        var total = writers.Length;
        var removeCount = Math.Max(0, total - componentsToKeep);
        if (total <= maximumComponents || removeCount == 0 || parsed.Length == 0)
            return null;
        return parsed[Math.Min(removeCount, parsed.Length) - 1].StartedUtc;
    }

    internal static DateTime WriterStartedUtc(string writerId)
    {
        var separator = writerId.IndexOf('-');
        if (separator <= 0 ||
            !long.TryParse(
                writerId.AsSpan(0, separator),
                NumberStyles.HexNumber,
                CultureInfo.InvariantCulture,
                out var ticks) ||
            ticks <= DateTime.MinValue.Ticks ||
            ticks >= DateTime.MaxValue.Ticks)
        {
            return DateTime.MinValue;
        }
        return new DateTime(ticks, DateTimeKind.Utc);
    }

    internal static int RemoveWriterEntriesThrough(
        Dictionary<string, int> counters,
        DateTime cutoff)
    {
        var removed = 0;
        foreach (var key in counters.Keys
                     .Where(key => WriterStartedUtc(key) <= cutoff)
                     .ToArray())
        {
            removed += counters[key];
            counters.Remove(key);
        }
        return removed;
    }

    internal static double RemoveWriterEntriesThrough(
        Dictionary<string, double> counters,
        DateTime cutoff)
    {
        var removed = 0.0;
        foreach (var key in counters.Keys
                     .Where(key => WriterStartedUtc(key) <= cutoff)
                     .ToArray())
        {
            removed += counters[key];
            counters.Remove(key);
        }
        return removed;
    }

    internal static void RemoveCompactedWriterEntries(
        Dictionary<string, int> counters,
        DateTime? cutoff)
    {
        if (!cutoff.HasValue)
            return;
        foreach (var key in counters.Keys
                     .Where(key => WriterStartedUtc(key) <= cutoff.Value)
                     .ToArray())
        {
            counters.Remove(key);
        }
    }

    internal static void RemoveCompactedWriterEntries(
        Dictionary<string, double> counters,
        DateTime? cutoff)
    {
        if (!cutoff.HasValue)
            return;
        foreach (var key in counters.Keys
                     .Where(key =>
                         WriterStartedUtc(key) <= cutoff.Value)
                     .ToArray())
        {
            counters.Remove(key);
        }
    }

}
