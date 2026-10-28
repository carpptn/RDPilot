using static RDPilot.Persistence.RecoveryWriterCounters;

namespace RDPilot.Persistence;

    /// <summary>
    /// Loads, migrates, merges, calibrates, and durably writes recovery memory.
    /// </summary>
    internal static partial class RecoveryMemoryService
    {
        internal static string EffectiveRecoveryMemoryPath() => RecoveryStoragePaths.ResolvePrimary(Options);

        internal static string EffectiveLoopReplayCorpusPath() => RecoveryStoragePaths.ResolveValidated(Options).Corpus;

        internal static CommandOutcome ExecuteRecoveryMemoryMaintenance(
            string command,
            string? exportPath)
        {
            if (!Options.RecoveryMemoryEnabled)
            {
                Console.Error.WriteLine("[memory] maintenance requires recovery memory to be enabled.");
                return CommandOutcome.Failed;
            }
            var paths = RecoveryStoragePaths.ResolveValidated(Options);
            var exportDestination = command.Equals("export", StringComparison.OrdinalIgnoreCase)
                ? paths.ValidateMemoryExportDestination(string.IsNullOrWhiteSpace(exportPath)
                    ? throw new InvalidOperationException("--memory-export requires a destination path.") : exportPath)
                : null;
            var lessons = LoadRecoveryLessons();
            if (RecoveryMemoryReadOnly)
            {
                Console.Error.WriteLine("[memory] maintenance failed: no supported, readable recovery store is available.");
                return CommandOutcome.Failed;
            }
            switch (command.ToLowerInvariant())
            {
                case "list":
                    PrintRecoveryMemoryReport(lessons);
                    return CommandOutcome.Success;
                case "prune":
                    var before = lessons.Count;
                    var pruneSaved = SaveRecoveryLessons(lessons);
                    Console.WriteLine(
                        pruneSaved
                            ? $"[memory] retention completed; entries before={before}; after={lessons.Count}"
                            : "[memory] retention was computed but could not be persisted.");
                    PrintRecoveryMemoryReport(lessons);
                    return pruneSaved ? CommandOutcome.Success : CommandOutcome.Failed;
                case "export":
                    if (!SaveRecoveryLessons(lessons))
                    {
                        Console.Error.WriteLine("[memory] export aborted because the current memory could not be persisted.");
                        return CommandOutcome.Failed;
                    }
                    var source = EffectiveRecoveryMemoryPath();
                    var destination = exportDestination!;
                    if (string.Equals(
                            Path.GetFullPath(source),
                            destination,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"[memory] export destination is the active memory file: {destination}");
                        return CommandOutcome.Success;
                    }
                    var directory = Path.GetDirectoryName(destination);
                    if (!string.IsNullOrWhiteSpace(directory))
                        Directory.CreateDirectory(directory);
                    File.Copy(source, destination, overwrite: true);
                    Console.WriteLine($"[memory] exported to {destination}");
                    return CommandOutcome.Success;
                default:
                    throw new InvalidOperationException($"Unknown memory command '{command}'.");
            }
        }

        static void PrintRecoveryMemoryReport(IReadOnlyCollection<RecoveryLesson> lessons)
        {
            var active = lessons.Count(IsLessonActive);
            var quarantined = lessons.Count - active;
            Console.WriteLine($"Recovery memory: active={active}; quarantined={quarantined}; path={EffectiveRecoveryMemoryPath()}");
            foreach (var lesson in lessons
                         .OrderByDescending(IsLessonActive)
                         .ThenByDescending(lesson => ContextualBanditScore(lesson, 1))
                         .ThenByDescending(lesson => lesson.UpdatedUtc))
            {
                Console.WriteLine(
                    $"- [{lesson.Status}] {lesson.Id} domain={lesson.GoalDomain}/{lesson.InteractionDomain} " +
                    $"loop={lesson.LoopKind}/{lesson.LoopTopology} " +
                    $"success={lesson.SuccessCount} failure={lesson.FailureCount} " +
                    $"reliability={RecoveryReliability(lesson):0.00}");
                Console.WriteLine($"  {TrimForMeta(lesson.WinningStrategy, 300)}");
                if (!string.IsNullOrWhiteSpace(lesson.LastFailureReason))
                    Console.WriteLine($"  last failure: {TrimForMeta(lesson.LastFailureReason, 200)}");
            }
            foreach (var (kind, bucket) in RecoveryCalibration.OrderBy(item => item.Key, StringComparer.Ordinal))
            {
                Console.WriteLine(
                    $"Calibration {kind}: candidates={bucket.CandidateCount}; " +
                    $"confirmed={bucket.ConfirmedCount}; rejected={bucket.RejectedCount}; " +
                    $"inconclusive={bucket.InconclusiveCount}; " +
                    $"threshold={CalibratedLoopThreshold(kind):0.00}");
            }
        }

        internal static List<RecoveryLesson> LoadRecoveryLessons()
        {
            if (!Options.RecoveryMemoryEnabled)
                return [];

            var path = RecoveryStoragePaths.ResolveValidated(Options).Primary;
            lock (RecoveryFileGate)
            {
                RecoveryMemoryReadOnly = false;
                RecoveryCalibration = new(StringComparer.OrdinalIgnoreCase);
                using var mutex = CreateRecoveryFileMutex(path);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                {
                    Console.WriteLine("[memory] shared file stayed locked; loading a read-only snapshot.");
                    var snapshot = ReadRecoveryStoreWithBackup(path);
                    if (snapshot is null)
                        return [];
                    if (snapshot.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] file version {snapshot.Version} is newer than supported version {CurrentMemoryVersion}; memory is read-only.");
                    }
                    MergePendingRecoveryStore(path, snapshot);
                    RecoveryCalibration = new(
                        snapshot.Calibration,
                        StringComparer.OrdinalIgnoreCase);
                    return snapshot.Lessons;
                }
                try
                {
                    var store = ReadRecoveryStoreWithBackup(path);
                    if (store is null)
                    {
                        if (RecoveryMemoryReadOnly) return [];
                        store = new RecoveryLessonStore();
                        if (!PendingRecoveryWrites.ContainsKey(path))
                        {
                            _ = SaveRecoveryLessonsWithoutLock(path, []);
                            return [];
                        }
                    }
                    if (store.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] file version {store.Version} is newer than supported version {CurrentMemoryVersion}; memory is read-only.");
                    }

                    MergePendingRecoveryStore(path, store);
                    RecoveryCalibration = new(store.Calibration, StringComparer.OrdinalIgnoreCase);
                    var loadArchive = ApplyLessonRetention(store.Lessons);
                    loadArchive.AddRange(
                        EnforceRecoveryMemoryFileSize(store.Lessons));
                    if (loadArchive.Count > 0 && !RecoveryMemoryReadOnly)
                    {
                        if (ArchiveRecoveryLessonsWithoutLock(
                                path,
                                loadArchive))
                        {
                            _ = SaveRecoveryLessonsWithoutLock(
                                path,
                                store.Lessons);
                        }
                        else
                        {
                            store.Lessons.AddRange(loadArchive);
                        }
                    }
                    if (LastRecoveryLoadUsedBackup && !RecoveryMemoryReadOnly)
                    {
                        if (SaveRecoveryLessonsWithoutLock(path, store.Lessons))
                            Console.WriteLine("[memory] restored the primary recovery file from its verified backup.");
                    }
                    var active = store.Lessons.Count(lesson => IsLessonActive(lesson));
                    var quarantined = store.Lessons.Count - active;
                    Console.WriteLine($"[memory] loaded {active} active and {quarantined} quarantined recovery lesson(s) from {path}");
                    return store.Lessons;
                }
                finally
                {
                    if (lockTaken)
                        mutex.ReleaseMutex();
                }
            }
        }

        static RecoveryLessonStore? ReadRecoveryStoreWithBackup(string path)
        {
            var result = RecoveryStoreReader.Read(path);
            LastRecoveryLoadUsedBackup = result.UsedBackup;
            if (result.Status is RecoveryStoreReadStatus.Unreadable or RecoveryStoreReadStatus.UnsupportedVersion)
            {
                RecoveryMemoryReadOnly = true;
                Console.WriteLine(result.Status == RecoveryStoreReadStatus.UnsupportedVersion
                    ? $"[memory] file version {result.Version} is newer than supported version {CurrentMemoryVersion}; preserving existing files and disabling writes until the next load."
                    : "[memory] no valid recovery file is available; preserving existing files and disabling writes until the next load.");
            }
            return result.Store;
        }

        internal static Mutex CreateRecoveryFileMutex(string path) => StorageFileLock.Create(path);

        internal static bool WaitForRecoveryMutex(Mutex mutex) => StorageFileLock.TryEnter(mutex);

        static bool SaveRecoveryLessons(List<RecoveryLesson> lessons)
        {
            if (RecoveryMemoryReadOnly)
                return false;

            var path = EffectiveRecoveryMemoryPath();
            lock (RecoveryFileGate)
            {
                using var mutex = CreateRecoveryFileMutex(path);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                {
                    Console.WriteLine("[memory] save skipped because the shared memory file stayed locked.");
                    MarkRecoveryMemoryDirty(lessons);
                    return false;
                }

                try
                {
                    var diskStore = ReadRecoveryStoreWithBackup(path) ?? new RecoveryLessonStore();
                    if (RecoveryMemoryReadOnly) return false;
                    if (diskStore.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] save skipped because file version {diskStore.Version} is newer than supported version {CurrentMemoryVersion}.");
                        return false;
                    }
                    MergePendingRecoveryStore(path, diskStore);
                    RecoveryCalibration = MergeCalibration(
                        diskStore.Calibration,
                        RecoveryCalibration);
                    var merged = MergeRecoveryLessons(diskStore.Lessons, lessons);
                    var archive = ApplyLessonRetention(merged);
                    archive.AddRange(
                        EnforceRecoveryMemoryFileSize(merged));
                    if (archive.Count > 0 &&
                        !ArchiveRecoveryLessonsWithoutLock(path, archive))
                    {
                        Console.WriteLine(
                            "[memory] save deferred because displaced lessons could not be archived.");
                        MarkRecoveryMemoryDirty(lessons);
                        return false;
                    }
                    if (!SaveRecoveryLessonsWithoutLock(path, merged))
                    {
                        MarkRecoveryMemoryDirty(lessons);
                        return false;
                    }
                    lessons.Clear();
                    lessons.AddRange(merged);
                    PendingRecoveryWrites.Remove(path);
                    return true;
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }

        static bool SaveRecoveryLessonsWithoutLock(
            string path,
            IReadOnlyCollection<RecoveryLesson> lessons)
        {
            try
            {
                var paths = RecoveryStoragePaths.ResolveValidated(Options, path);
                var store = new RecoveryLessonStore
                {
                    Version = CurrentMemoryVersion,
                    Lessons = lessons.ToList(),
                    Calibration = new(RecoveryCalibration, StringComparer.OrdinalIgnoreCase)
                };
                AtomicTextStore.Write(paths.Primary, JsonSerializer.Serialize(store, PrettyJson),
                    LastRecoveryLoadUsedBackup ? null : paths.Backup);
                LastRecoveryLoadUsedBackup = false;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[memory] could not save recovery lessons: {ex.Message}");
                return false;
            }
        }

        private static void MergePendingRecoveryStore(string path, RecoveryLessonStore store)
        {
            if (!PendingRecoveryWrites.TryGetValue(path, out var pending)) return;
            pending = JsonSerializer.Deserialize<RecoveryLessonStore>(JsonSerializer.Serialize(pending, PrettyJson), PrettyJson)!;
            store.Lessons = MergeRecoveryLessons(store.Lessons, pending.Lessons);
            store.Calibration = MergeCalibration(store.Calibration, pending.Calibration);
        }

        static void MarkRecoveryMemoryDirty(List<RecoveryLesson> lessons)
        {
            var path = EffectiveRecoveryMemoryPath();
            var snapshot = new RecoveryLessonStore { Lessons = lessons, Calibration = RecoveryCalibration };
            // Pending writes own a snapshot, independent of mutations in the next run.
            snapshot = JsonSerializer.Deserialize<RecoveryLessonStore>(JsonSerializer.Serialize(snapshot, PrettyJson), PrettyJson)!;
            MergePendingRecoveryStore(path, snapshot);
            PendingRecoveryWrites[path] = snapshot;
        }

        internal static bool FlushPendingRecoveryMemory()
        {
            lock (RecoveryFileGate)
            {
                var path = EffectiveRecoveryMemoryPath();
                if (!PendingRecoveryWrites.TryGetValue(path, out var pending)) return true;
                if (!Options.RecoveryMemoryEnabled) return false;
                Console.WriteLine("[memory] retrying pending durable recovery-memory write.");
                RecoveryCalibration = MergeCalibration(RecoveryCalibration, pending.Calibration);
                return SaveRecoveryLessons(pending.Lessons.ToList());
            }
        }

        static void PersistCalibrationSnapshot()
        {
            if (!Options.RecoveryMemoryEnabled || RecoveryMemoryReadOnly)
                return;

            var path = EffectiveRecoveryMemoryPath();
            lock (RecoveryFileGate)
            {
                using var mutex = CreateRecoveryFileMutex(path);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                {
                    MarkRecoveryMemoryDirty([]);
                    return;
                }
                try
                {
                    var store = ReadRecoveryStoreWithBackup(path) ?? new RecoveryLessonStore();
                    if (RecoveryMemoryReadOnly) return;
                    if (store.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] calibration save skipped because file version {store.Version} is newer than supported version {CurrentMemoryVersion}.");
                        return;
                    }
                    MergePendingRecoveryStore(path, store);
                    RecoveryCalibration = MergeCalibration(store.Calibration, RecoveryCalibration);
                    if (!SaveRecoveryLessonsWithoutLock(path, store.Lessons))
                        MarkRecoveryMemoryDirty(store.Lessons);
                    else
                        PendingRecoveryWrites.Remove(path);
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }

        static List<RecoveryLesson> MergeRecoveryLessons(
            IReadOnlyCollection<RecoveryLesson> disk,
            IReadOnlyCollection<RecoveryLesson> incoming)
        {
            var merged = disk
                .Concat(incoming)
                .GroupBy(lesson => lesson.Id, StringComparer.Ordinal)
                .Select(group =>
                {
                    var versions = group.ToArray();
                    foreach (var version in versions)
                        NormalizeLessonCounters(version);
                    var latest = versions
                        .OrderByDescending(lesson => lesson.UpdatedUtc)
                        .ThenByDescending(lesson => lesson.SuccessCount + lesson.FailureCount)
                        .First();
                    latest.SuccessByWriter = versions
                        .Select(version => version.SuccessByWriter)
                        .Aggregate(
                            new Dictionary<string, int>(StringComparer.Ordinal),
                            MergeWriterCounters);
                    latest.FailureByWriter = versions
                        .Select(version => version.FailureByWriter)
                        .Aggregate(
                            new Dictionary<string, int>(StringComparer.Ordinal),
                            MergeWriterCounters);
                    latest.SelectionByWriter = versions
                        .Select(version => version.SelectionByWriter)
                        .Aggregate(
                            new Dictionary<string, int>(StringComparer.Ordinal),
                            MergeWriterCounters);
                    latest.RewardByWriter = versions
                        .Select(version => version.RewardByWriter)
                        .Aggregate(
                            new Dictionary<string, double>(StringComparer.Ordinal),
                            MergeWriterDoubleCounters);
                    latest.RewardObservationByWriter = versions
                        .Select(version => version.RewardObservationByWriter)
                        .Aggregate(
                            new Dictionary<string, int>(StringComparer.Ordinal),
                            MergeWriterCounters);
                    latest.CountersCompactedBeforeUtc = LatestCompactionCutoff(versions
                        .Select(version => version.CountersCompactedBeforeUtc));
                    latest.CompactedSuccessCount = versions.Max(version => version.CompactedSuccessCount);
                    latest.CompactedFailureCount = versions.Max(version => version.CompactedFailureCount);
                    latest.CompactedSelectionCount = versions.Max(version => version.CompactedSelectionCount);
                    latest.CompactedCumulativeReward = versions.Max(version => version.CompactedCumulativeReward);
                    latest.CompactedRewardObservationCount = versions.Max(version => version.CompactedRewardObservationCount);
                    RemoveCompactedWriterEntries(
                        latest.SuccessByWriter,
                        latest.CountersCompactedBeforeUtc);
                    RemoveCompactedWriterEntries(
                        latest.FailureByWriter,
                        latest.CountersCompactedBeforeUtc);
                    RemoveCompactedWriterEntries(
                        latest.SelectionByWriter,
                        latest.CountersCompactedBeforeUtc);
                    RemoveCompactedWriterEntries(
                        latest.RewardByWriter,
                        latest.CountersCompactedBeforeUtc);
                    RemoveCompactedWriterEntries(
                        latest.RewardObservationByWriter,
                        latest.CountersCompactedBeforeUtc);
                    NormalizeLessonCounters(latest);
                    CompactLessonWriterCounters(latest);
                    return latest;
                })
                .ToList();
            return merged;
        }

        internal static void NormalizeLessonCounters(RecoveryLesson lesson)
        {
            lesson.SuccessByWriter ??= [];
            lesson.FailureByWriter ??= [];
            lesson.SelectionByWriter ??= [];
            lesson.RewardByWriter ??= [];
            lesson.RewardObservationByWriter ??= [];
            if (lesson.SuccessByWriter.Count == 0 &&
                lesson.SuccessCount > lesson.CompactedSuccessCount)
            {
                lesson.SuccessByWriter["legacy"] =
                    lesson.SuccessCount - lesson.CompactedSuccessCount;
            }
            if (lesson.FailureByWriter.Count == 0 &&
                lesson.FailureCount > lesson.CompactedFailureCount)
            {
                lesson.FailureByWriter["legacy"] =
                    lesson.FailureCount - lesson.CompactedFailureCount;
            }
            if (lesson.SelectionByWriter.Count == 0 &&
                lesson.SelectionCount > lesson.CompactedSelectionCount)
            {
                lesson.SelectionByWriter["legacy"] =
                    lesson.SelectionCount - lesson.CompactedSelectionCount;
            }
            if (lesson.RewardByWriter.Count == 0 &&
                lesson.CumulativeReward >
                lesson.CompactedCumulativeReward)
            {
                lesson.RewardByWriter["legacy"] =
                    lesson.CumulativeReward -
                    lesson.CompactedCumulativeReward;
            }
            if (lesson.RewardObservationByWriter.Count == 0 &&
                lesson.RewardObservationCount >
                lesson.CompactedRewardObservationCount)
            {
                lesson.RewardObservationByWriter["legacy"] =
                    lesson.RewardObservationCount -
                    lesson.CompactedRewardObservationCount;
            }
            RemoveCompactedWriterEntries(
                lesson.SuccessByWriter,
                lesson.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                lesson.FailureByWriter,
                lesson.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                lesson.SelectionByWriter,
                lesson.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                lesson.RewardByWriter,
                lesson.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                lesson.RewardObservationByWriter,
                lesson.CountersCompactedBeforeUtc);
            lesson.SuccessCount =
                lesson.CompactedSuccessCount + lesson.SuccessByWriter.Values.Sum();
            lesson.FailureCount =
                lesson.CompactedFailureCount + lesson.FailureByWriter.Values.Sum();
            lesson.SelectionCount =
                lesson.CompactedSelectionCount +
                lesson.SelectionByWriter.Values.Sum();
            lesson.CumulativeReward =
                lesson.CompactedCumulativeReward +
                lesson.RewardByWriter.Values.Sum();
            lesson.RewardObservationCount =
                lesson.CompactedRewardObservationCount +
                lesson.RewardObservationByWriter.Values.Sum();
        }

        static void RecordLessonSuccess(RecoveryLesson lesson, double reward)
        {
            NormalizeLessonCounters(lesson);
            IncrementWriterCounter(lesson.SuccessByWriter, lesson.CountersCompactedBeforeUtc);
            IncrementWriterCounter(lesson.RewardObservationByWriter, lesson.CountersCompactedBeforeUtc);
            IncrementWriterReward(
                lesson.RewardByWriter,
                Math.Clamp(reward, 0, 1), lesson.CountersCompactedBeforeUtc);
            NormalizeLessonCounters(lesson);
        }

        static void RecordLessonFailure(RecoveryLesson lesson)
        {
            NormalizeLessonCounters(lesson);
            IncrementWriterCounter(lesson.FailureByWriter, lesson.CountersCompactedBeforeUtc);
            IncrementWriterCounter(lesson.RewardObservationByWriter, lesson.CountersCompactedBeforeUtc);
            NormalizeLessonCounters(lesson);
        }

        static void RecordLessonSelection(RecoveryLesson lesson)
        {
            NormalizeLessonCounters(lesson);
            IncrementWriterCounter(lesson.SelectionByWriter, lesson.CountersCompactedBeforeUtc);
            NormalizeLessonCounters(lesson);
        }

        static void CompactLessonWriterCounters(RecoveryLesson lesson)
        {
            var cutoff = WriterCompactionCutoff(
                lesson.SuccessByWriter.Keys
                    .Concat(lesson.FailureByWriter.Keys)
                    .Concat(lesson.SelectionByWriter.Keys)
                    .Concat(lesson.RewardByWriter.Keys)
                    .Concat(lesson.RewardObservationByWriter.Keys));
            if (!cutoff.HasValue)
                return;

            lesson.CompactedSuccessCount += RemoveWriterEntriesThrough(
                lesson.SuccessByWriter,
                cutoff.Value);
            lesson.CompactedFailureCount += RemoveWriterEntriesThrough(
                lesson.FailureByWriter,
                cutoff.Value);
            lesson.CompactedSelectionCount += RemoveWriterEntriesThrough(
                lesson.SelectionByWriter,
                cutoff.Value);
            lesson.CompactedCumulativeReward +=
                RemoveWriterEntriesThrough(
                    lesson.RewardByWriter,
                    cutoff.Value);
            lesson.CompactedRewardObservationCount +=
                RemoveWriterEntriesThrough(
                    lesson.RewardObservationByWriter,
                    cutoff.Value);
            lesson.CountersCompactedBeforeUtc = LatestCompactionCutoff(
                lesson.CountersCompactedBeforeUtc, cutoff);
            NormalizeLessonCounters(lesson);
        }

        static bool IsLessonActive(RecoveryLesson lesson) =>
            string.Equals(lesson.Status, "active", StringComparison.OrdinalIgnoreCase);

        static List<RecoveryLesson> ApplyLessonRetention(
            List<RecoveryLesson> lessons)
        {
            var archived = new List<RecoveryLesson>();
            var staleCutoff = RuntimeClock.UtcNow.AddDays(-365);
            foreach (var lesson in lessons.Where(IsLessonActive))
            {
                if (lesson.UpdatedUtc < staleCutoff &&
                    lesson.SuccessCount <= 1 &&
                    RecoveryReliability(lesson) < 0.55)
                {
                    lesson.Status = "quarantined";
                    lesson.QuarantinedUtc = RuntimeClock.UtcNow;
                    lesson.LastFailureReason = "stale low-confidence strategy";
                    lesson.UpdatedUtc = RuntimeClock.UtcNow;
                }
            }

            var activeKeep = SelectActiveLessonIdsForRetention(
                lessons.Where(IsLessonActive));
            foreach (var lesson in lessons.Where(lesson =>
                         IsLessonActive(lesson) && !activeKeep.Contains(lesson.Id)))
            {
                lesson.Status = "quarantined";
                lesson.QuarantinedUtc = RuntimeClock.UtcNow;
                lesson.LastFailureReason = "superseded by higher-value strategies at the active-memory limit";
                lesson.UpdatedUtc = RuntimeClock.UtcNow;
            }

            var quarantineKeep = lessons
                .Where(lesson => !IsLessonActive(lesson))
                .OrderByDescending(lesson => lesson.UpdatedUtc)
                .Take(Math.Max(1, Options.RecoveryMemoryMaxQuarantinedLessons))
                .Select(lesson => lesson.Id)
                .ToHashSet(StringComparer.Ordinal);
            var quarantineOverflow = lessons
                .Where(lesson =>
                    !IsLessonActive(lesson) &&
                    !quarantineKeep.Contains(lesson.Id))
                .ToList();
            foreach (var lesson in quarantineOverflow)
            {
                lesson.Status = "archived";
                lesson.LastFailureReason =
                    string.IsNullOrWhiteSpace(lesson.LastFailureReason)
                        ? "moved out of bounded quarantine retention"
                        : lesson.LastFailureReason;
                archived.Add(lesson);
            }
            lessons.RemoveAll(lesson =>
                quarantineOverflow.Contains(lesson));
            return archived;
        }

        internal static HashSet<string> SelectActiveLessonIdsForRetention(
            IEnumerable<RecoveryLesson> activeLessons)
        {
            var globalLimit = Math.Max(1, Options.RecoveryMemoryMaxLessons);
            var reservePerContext = Math.Max(
                0,
                Options.RecoveryMemoryReservedLessonsPerContext);
            var softMaxPerContext = Math.Max(
                Math.Max(1, reservePerContext),
                Options.RecoveryMemorySoftMaxLessonsPerContext);
            var ranked = activeLessons
                .OrderByDescending(lesson =>
                    ContextualBanditScore(lesson, 1))
                .ThenByDescending(lesson => lesson.UpdatedUtc)
                .ToList();
            var groups = ranked
                .GroupBy(RecoveryRetentionContextKey)
                .Select(group => group.ToList())
                .OrderByDescending(group =>
                    ContextualBanditScore(group[0], 1))
                .ToList();
            var keep = new HashSet<string>(StringComparer.Ordinal);

            // Reserve a small number of strong lessons for every observed
            // application/domain context before global competition.
            for (var round = 0;
                 round < reservePerContext && keep.Count < globalLimit;
                 round++)
            {
                foreach (var group in groups)
                {
                    if (round < group.Count)
                        keep.Add(group[round].Id);
                    if (keep.Count >= globalLimit)
                        break;
                }
            }

            var contextCounts = ranked
                .Where(lesson => keep.Contains(lesson.Id))
                .GroupBy(RecoveryRetentionContextKey)
                .ToDictionary(
                    group => group.Key,
                    group => group.Count(),
                    StringComparer.Ordinal);
            foreach (var lesson in ranked)
            {
                if (keep.Count >= globalLimit)
                    break;
                if (keep.Contains(lesson.Id))
                    continue;
                var context = RecoveryRetentionContextKey(lesson);
                contextCounts.TryGetValue(context, out var contextCount);
                if (contextCount >= softMaxPerContext)
                    continue;
                keep.Add(lesson.Id);
                contextCounts[context] = contextCount + 1;
            }

            // The per-context ceiling is soft: unused global capacity remains
            // available when there are not enough diverse contexts.
            foreach (var lesson in ranked)
            {
                if (keep.Count >= globalLimit)
                    break;
                keep.Add(lesson.Id);
            }
            return keep;
        }

        static string RecoveryRetentionContextKey(
            RecoveryLesson lesson) =>
            $"{NormalizeText(lesson.ActiveProcess)}|" +
            $"{NormalizeText(lesson.GoalDomain)}";

        static List<RecoveryLesson> EnforceRecoveryMemoryFileSize(
            List<RecoveryLesson> lessons)
        {
            var archived = new List<RecoveryLesson>();
            var maxBytes = Math.Max(1024, Options.RecoveryMemoryMaxFileBytes);
            if (RecoveryStoreSerializedSize(lessons) <= maxBytes)
                return archived;

            var removalOrder = lessons
                .OrderBy(IsLessonActive)
                .ThenBy(lesson =>
                    IsLessonActive(lesson)
                        ? ContextualBanditScore(lesson, 1)
                        : 0)
                .ThenBy(lesson => lesson.UpdatedUtc)
                .ToList();
            foreach (var lesson in removalOrder)
            {
                if (lessons.Count <= 1 ||
                    RecoveryStoreSerializedSize(lessons) <= maxBytes)
                {
                    break;
                }
                lessons.Remove(lesson);
                lesson.Status = "archived";
                lesson.QuarantinedUtc ??= RuntimeClock.UtcNow;
                lesson.LastFailureReason =
                    "moved to archive to enforce the primary memory file-size limit";
                lesson.UpdatedUtc = RuntimeClock.UtcNow;
                archived.Add(lesson);
            }
            return archived;
        }

        static int RecoveryStoreSerializedSize(
            IReadOnlyCollection<RecoveryLesson> lessons)
        {
            var store = new RecoveryLessonStore
            {
                Version = CurrentMemoryVersion,
                Lessons = lessons.ToList(),
                Calibration = new(
                    RecoveryCalibration,
                    StringComparer.OrdinalIgnoreCase)
            };
            return JsonSerializer.SerializeToUtf8Bytes(
                store,
                PrettyJson).Length;
        }

        internal static string EffectiveRecoveryMemoryArchivePath() => RecoveryStoragePaths.Resolve(Options).Archive;

        internal static string ResolveApplicationPath(string path) => RecoveryStoragePaths.ResolvePath(path);

        static bool ArchiveRecoveryLessonsWithoutLock(string primaryPath, IReadOnlyCollection<RecoveryLesson> lessons) =>
            RecoveryArchiveStore.Append(primaryPath, lessons);

        static LoopCalibrationBucket CalibrationBucket(string loopKind)
        {
            if (!RecoveryCalibration.TryGetValue(loopKind, out var bucket))
            {
                bucket = new LoopCalibrationBucket();
                RecoveryCalibration[loopKind] = bucket;
            }
            return bucket;
        }

        static void RegisterCalibrationCandidate(string loopKind)
        {
            var bucket = CalibrationBucket(loopKind);
            IncrementWriterCounter(bucket.CandidateByWriter, bucket.CountersCompactedBeforeUtc);
            NormalizeCalibrationBucket(bucket);
            bucket.UpdatedUtc = RuntimeClock.UtcNow;
        }

        static void RegisterCalibrationOutcome(string loopKind, bool confirmed)
        {
            var bucket = CalibrationBucket(loopKind);
            if (confirmed) IncrementWriterCounter(bucket.ConfirmedByWriter, bucket.CountersCompactedBeforeUtc);
            else IncrementWriterCounter(bucket.RejectedByWriter, bucket.CountersCompactedBeforeUtc);
            NormalizeCalibrationBucket(bucket);
            bucket.UpdatedUtc = RuntimeClock.UtcNow;
        }

        static void RegisterCalibrationInconclusive(string loopKind)
        {
            if (string.IsNullOrWhiteSpace(loopKind))
                return;
            var bucket = CalibrationBucket(loopKind);
            IncrementWriterCounter(bucket.InconclusiveByWriter, bucket.CountersCompactedBeforeUtc);
            NormalizeCalibrationBucket(bucket);
            bucket.UpdatedUtc = RuntimeClock.UtcNow;
        }

        static void IncrementWriterCounter(Dictionary<string, int> counters, DateTime? cutoff)
        {
            var writerId = RecoveryWriterSession.IdFor(cutoff, counters.Keys);
            counters.TryGetValue(writerId, out var value);
            counters[writerId] = value + 1;
        }

        static void IncrementWriterReward(
            Dictionary<string, double> rewards,
            double value, DateTime? cutoff)
        {
            var writerId = RecoveryWriterSession.IdFor(cutoff, rewards.Keys);
            rewards.TryGetValue(writerId, out var existing);
            rewards[writerId] = existing + value;
        }

        internal static void NormalizeCalibrationBucket(LoopCalibrationBucket bucket)
        {
            bucket.CandidateByWriter ??= [];
            bucket.ConfirmedByWriter ??= [];
            bucket.RejectedByWriter ??= [];
            bucket.InconclusiveByWriter ??= [];
            if (bucket.CandidateByWriter.Count == 0 &&
                bucket.CandidateCount > bucket.CompactedCandidateCount)
            {
                bucket.CandidateByWriter["legacy"] =
                    bucket.CandidateCount - bucket.CompactedCandidateCount;
            }
            if (bucket.ConfirmedByWriter.Count == 0 &&
                bucket.ConfirmedCount > bucket.CompactedConfirmedCount)
            {
                bucket.ConfirmedByWriter["legacy"] =
                    bucket.ConfirmedCount - bucket.CompactedConfirmedCount;
            }
            if (bucket.RejectedByWriter.Count == 0 &&
                bucket.RejectedCount > bucket.CompactedRejectedCount)
            {
                bucket.RejectedByWriter["legacy"] =
                    bucket.RejectedCount - bucket.CompactedRejectedCount;
            }
            if (bucket.InconclusiveByWriter.Count == 0 &&
                bucket.InconclusiveCount >
                bucket.CompactedInconclusiveCount)
            {
                bucket.InconclusiveByWriter["legacy"] =
                    bucket.InconclusiveCount -
                    bucket.CompactedInconclusiveCount;
            }
            RemoveCompactedWriterEntries(
                bucket.CandidateByWriter,
                bucket.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                bucket.ConfirmedByWriter,
                bucket.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                bucket.RejectedByWriter,
                bucket.CountersCompactedBeforeUtc);
            RemoveCompactedWriterEntries(
                bucket.InconclusiveByWriter,
                bucket.CountersCompactedBeforeUtc);
            bucket.CandidateCount =
                bucket.CompactedCandidateCount + bucket.CandidateByWriter.Values.Sum();
            bucket.ConfirmedCount =
                bucket.CompactedConfirmedCount + bucket.ConfirmedByWriter.Values.Sum();
            bucket.RejectedCount =
                bucket.CompactedRejectedCount + bucket.RejectedByWriter.Values.Sum();
            bucket.InconclusiveCount =
                bucket.CompactedInconclusiveCount +
                bucket.InconclusiveByWriter.Values.Sum();
        }

        static Dictionary<string, LoopCalibrationBucket> MergeCalibration(
            IReadOnlyDictionary<string, LoopCalibrationBucket> left,
            IReadOnlyDictionary<string, LoopCalibrationBucket> right)
        {
            var result = new Dictionary<string, LoopCalibrationBucket>(StringComparer.OrdinalIgnoreCase);
            foreach (var kind in left.Keys.Concat(right.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                left.TryGetValue(kind, out var leftBucket);
                right.TryGetValue(kind, out var rightBucket);
                leftBucket ??= new LoopCalibrationBucket();
                rightBucket ??= new LoopCalibrationBucket();
                NormalizeCalibrationBucket(leftBucket);
                NormalizeCalibrationBucket(rightBucket);
                var merged = new LoopCalibrationBucket
                {
                    CandidateByWriter = MergeWriterCounters(leftBucket.CandidateByWriter, rightBucket.CandidateByWriter),
                    ConfirmedByWriter = MergeWriterCounters(leftBucket.ConfirmedByWriter, rightBucket.ConfirmedByWriter),
                    RejectedByWriter = MergeWriterCounters(leftBucket.RejectedByWriter, rightBucket.RejectedByWriter),
                    InconclusiveByWriter = MergeWriterCounters(
                        leftBucket.InconclusiveByWriter,
                        rightBucket.InconclusiveByWriter),
                    CompactedCandidateCount = Math.Max(
                        leftBucket.CompactedCandidateCount,
                        rightBucket.CompactedCandidateCount),
                    CompactedConfirmedCount = Math.Max(
                        leftBucket.CompactedConfirmedCount,
                        rightBucket.CompactedConfirmedCount),
                    CompactedRejectedCount = Math.Max(
                        leftBucket.CompactedRejectedCount,
                        rightBucket.CompactedRejectedCount),
                    CompactedInconclusiveCount = Math.Max(
                        leftBucket.CompactedInconclusiveCount,
                        rightBucket.CompactedInconclusiveCount),
                    CountersCompactedBeforeUtc = LatestCompactionCutoff(
                        leftBucket.CountersCompactedBeforeUtc,
                        rightBucket.CountersCompactedBeforeUtc),
                    UpdatedUtc = leftBucket.UpdatedUtc >= rightBucket.UpdatedUtc
                        ? leftBucket.UpdatedUtc
                        : rightBucket.UpdatedUtc
                };
                RemoveCompactedWriterEntries(
                    merged.CandidateByWriter,
                    merged.CountersCompactedBeforeUtc);
                RemoveCompactedWriterEntries(
                    merged.ConfirmedByWriter,
                    merged.CountersCompactedBeforeUtc);
                RemoveCompactedWriterEntries(
                    merged.RejectedByWriter,
                    merged.CountersCompactedBeforeUtc);
                RemoveCompactedWriterEntries(
                    merged.InconclusiveByWriter,
                    merged.CountersCompactedBeforeUtc);
                NormalizeCalibrationBucket(merged);
                CompactCalibrationWriterCounters(merged);
                result[kind] = merged;
            }
            return result;
        }

        static void CompactCalibrationWriterCounters(LoopCalibrationBucket bucket)
        {
            var cutoff = WriterCompactionCutoff(
                bucket.CandidateByWriter.Keys
                    .Concat(bucket.ConfirmedByWriter.Keys)
                    .Concat(bucket.RejectedByWriter.Keys)
                    .Concat(bucket.InconclusiveByWriter.Keys));
            if (!cutoff.HasValue)
                return;

            bucket.CompactedCandidateCount += RemoveWriterEntriesThrough(
                bucket.CandidateByWriter,
                cutoff.Value);
            bucket.CompactedConfirmedCount += RemoveWriterEntriesThrough(
                bucket.ConfirmedByWriter,
                cutoff.Value);
            bucket.CompactedRejectedCount += RemoveWriterEntriesThrough(
                bucket.RejectedByWriter,
                cutoff.Value);
            bucket.CompactedInconclusiveCount +=
                RemoveWriterEntriesThrough(
                    bucket.InconclusiveByWriter,
                    cutoff.Value);
            bucket.CountersCompactedBeforeUtc = LatestCompactionCutoff(
                bucket.CountersCompactedBeforeUtc, cutoff);
            NormalizeCalibrationBucket(bucket);
        }

        static double CalibratedLoopThreshold(
            string calibrationKey,
            string? fallbackKey = null)
        {
            if (!RecoveryCalibration.TryGetValue(calibrationKey, out var bucket) &&
                (string.IsNullOrWhiteSpace(fallbackKey) ||
                 !RecoveryCalibration.TryGetValue(fallbackKey, out bucket)))
            {
                return Options.ProactiveLoopConfidenceThreshold;
            }
            var labeled = bucket.ConfirmedCount + bucket.RejectedCount;
            if (labeled < 10)
                return Options.ProactiveLoopConfidenceThreshold;

            var precision = bucket.ConfirmedCount / (double)labeled;
            var adjustment = precision switch
            {
                < 0.60 => 0.10,
                < 0.75 => 0.05,
                > 0.92 => -0.04,
                _ => 0
            };
            return Math.Clamp(Options.ProactiveLoopConfidenceThreshold + adjustment, 0.55, 0.95);
        }

    }
