namespace RDPilot.Persistence;

/// <summary>Accepts a recovery file only after structural validation, then falls back to its backup.</summary>
internal enum RecoveryStoreReadStatus { Missing, Loaded, Unreadable, UnsupportedVersion }
internal sealed record RecoveryStoreReadResult(RecoveryStoreReadStatus Status, RecoveryLessonStore? Store = null, bool UsedBackup = false, int? Version = null);

internal static class RecoveryStoreReader
{
    internal static RecoveryLessonStore? Read(string path, out bool usedBackup)
    {
        var result = Read(path);
        usedBackup = result.UsedBackup;
        return result.Store;
    }

    internal static RecoveryStoreReadResult Read(string path)
    {
        var found = false;
        foreach (var candidate in new[] { path, RecoveryStoragePaths.BackupFor(path) })
        {
            // Read directly: File.Exists also returns false for inaccessible files.
            // Such failures must not be mistaken for an absent, writable store.
            try
            {
                var json = File.ReadAllText(candidate);
                var version = RecoveryStoreFormat.ReadVersion(json);
                if (version > CurrentMemoryVersion)
                    return new(RecoveryStoreReadStatus.UnsupportedVersion, UsedBackup: candidate != path, Version: version);
                var store = JsonSerializer.Deserialize<RecoveryLessonStore>(json, PrettyJson)
                    ?? throw new InvalidDataException("Recovery store is null.");
                Validate(store);
                RecoveryStoreNormalizer.Normalize(store);
                return new(RecoveryStoreReadStatus.Loaded, store, candidate != path);
            }
            catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException) { }
            catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or JsonException or ArgumentException or OverflowException)
            {
                found = true;
                Console.WriteLine($"[memory] could not load {candidate}: {ex.Message}");
            }
        }
        return new(found ? RecoveryStoreReadStatus.Unreadable : RecoveryStoreReadStatus.Missing);
    }

    internal static void Validate(RecoveryLessonStore store)
    {
        // Missing/null collections are supported by migration; null entries are not.
        foreach (var lesson in store.Lessons ?? [])
        {
            if (lesson is null) throw new InvalidDataException("Recovery lessons contain a null entry.");
            if (string.IsNullOrWhiteSpace(lesson.Id)) throw new InvalidDataException("Recovery lesson has no ID.");
            if (lesson.StrategySteps?.Any(step => step is null) == true ||
                lesson.WinningActionTypes?.Any(type => type is null) == true)
                throw new InvalidDataException("Recovery strategy contains a null entry.");
            Counters(lesson.CompactedSuccessCount, lesson.SuccessCount, lesson.SuccessByWriter);
            Counters(lesson.CompactedFailureCount, lesson.FailureCount, lesson.FailureByWriter);
            Counters(lesson.CompactedSelectionCount, lesson.SelectionCount, lesson.SelectionByWriter);
            Counters(lesson.CompactedRewardObservationCount, lesson.RewardObservationCount, lesson.RewardObservationByWriter);
            if (!double.IsFinite(lesson.CumulativeReward) || !double.IsFinite(lesson.CompactedCumulativeReward) ||
                lesson.RewardByWriter?.Values.Any(value => !double.IsFinite(value)) == true ||
                !double.IsFinite(lesson.CompactedCumulativeReward + (lesson.RewardByWriter?.Values.Sum() ?? 0)))
                throw new InvalidDataException("Recovery rewards are not finite.");
        }
        if (store.Calibration is not null &&
            (store.Calibration.Values.Any(bucket => bucket is null) ||
             store.Calibration.Keys.Distinct(StringComparer.OrdinalIgnoreCase).Count() != store.Calibration.Count))
            throw new InvalidDataException("Recovery calibration contains null or duplicate buckets.");
        foreach (var bucket in store.Calibration?.Values ?? Enumerable.Empty<LoopCalibrationBucket>())
        {
            Counters(bucket.CompactedCandidateCount, bucket.CandidateCount, bucket.CandidateByWriter);
            Counters(bucket.CompactedConfirmedCount, bucket.ConfirmedCount, bucket.ConfirmedByWriter);
            Counters(bucket.CompactedRejectedCount, bucket.RejectedCount, bucket.RejectedByWriter);
            Counters(bucket.CompactedInconclusiveCount, bucket.InconclusiveCount, bucket.InconclusiveByWriter);
        }
    }

    static void Counters(int compacted, int total, Dictionary<string, int>? writers)
    {
        if (compacted < 0 || total < 0 || writers?.Values.Any(value => value < 0) == true ||
            compacted + (writers?.Values.Sum(value => (long)value) ?? 0) > int.MaxValue)
            throw new InvalidDataException("Recovery counter is negative or exceeds Int32 capacity.");
    }
}
