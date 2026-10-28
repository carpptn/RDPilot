namespace RDPilot.Persistence;

/// <summary>Migrates a validated candidate before the reader accepts it.</summary>
internal static class RecoveryStoreNormalizer
{
    internal static void Normalize(RecoveryLessonStore store)
    {
        var sourceVersion = store.Version <= 0 ? 1 : store.Version;
        store.Version = sourceVersion;
        store.Lessons ??= [];
        store.Calibration = store.Calibration is null
            ? new(StringComparer.OrdinalIgnoreCase)
            : new(store.Calibration, StringComparer.OrdinalIgnoreCase);

        foreach (var lesson in store.Lessons)
        {
            NormalizeLessonCounters(lesson);
            lesson.Status = string.IsNullOrWhiteSpace(lesson.Status) ? "active" : lesson.Status;
            lesson.GoalMode = string.IsNullOrWhiteSpace(lesson.GoalMode)
                ? "finite"
                : lesson.GoalMode;
            lesson.StrategySteps ??= [];
            lesson.WinningActionTypes ??= [];
            if (lesson.StrategySteps.Count == 0 && lesson.WinningActionTypes.Length > 0)
            {
                lesson.StrategySteps = lesson.WinningActionTypes
                    .Select(family => new RecoveryStrategyStep { ActionFamily = family })
                    .ToList();
            }
            lesson.StrategySignature = sourceVersion < 5 ||
                                       string.IsNullOrWhiteSpace(lesson.StrategySignature)
                ? StrategySignature(lesson.StrategySteps)
                : lesson.StrategySignature;
        }
        foreach (var bucket in store.Calibration.Values)
            NormalizeCalibrationBucket(bucket);
    }

}
