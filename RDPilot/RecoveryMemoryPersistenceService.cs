internal static partial class RDPilotApplication
{
    /// <summary>
    /// Loads, migrates, merges, calibrates, and durably writes recovery memory.
    /// </summary>
    internal static partial class RecoveryMemoryService
    {
        internal static string EffectiveRecoveryMemoryPath()
        {
            if (!string.IsNullOrWhiteSpace(RecoveryMemoryPath))
                return Path.GetFullPath(Environment.ExpandEnvironmentVariables(RecoveryMemoryPath.Trim()));

            return Path.Combine(AppContext.BaseDirectory, "memory", "recovery-memory.json");
        }

        internal static string EffectiveLoopReplayCorpusPath()
        {
            if (!string.IsNullOrWhiteSpace(LoopReplayCorpusPath))
            {
                return Path.GetFullPath(
                    Environment.ExpandEnvironmentVariables(
                        LoopReplayCorpusPath.Trim()));
            }

            var directory = Path.GetDirectoryName(EffectiveRecoveryMemoryPath())
                            ?? AppContext.BaseDirectory;
            return Path.Combine(directory, "loop-replay-corpus.json");
        }

        internal static void ExecuteRecoveryMemoryMaintenance(
            string command,
            string? exportPath)
        {
            var lessons = LoadRecoveryLessons();
            switch (command.ToLowerInvariant())
            {
                case "list":
                    PrintRecoveryMemoryReport(lessons);
                    break;
                case "prune":
                    var before = lessons.Count;
                    var pruneSaved = SaveRecoveryLessons(lessons);
                    Console.WriteLine(
                        pruneSaved
                            ? $"[memory] retention completed; entries before={before}; after={lessons.Count}"
                            : "[memory] retention was computed but could not be persisted.");
                    PrintRecoveryMemoryReport(lessons);
                    break;
                case "export":
                    if (string.IsNullOrWhiteSpace(exportPath))
                        throw new InvalidOperationException("--memory-export requires a destination path.");
                    _ = SaveRecoveryLessons(lessons);
                    var source = EffectiveRecoveryMemoryPath();
                    var destination = Path.GetFullPath(
                        Environment.ExpandEnvironmentVariables(exportPath));
                    if (string.Equals(
                            Path.GetFullPath(source),
                            destination,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"[memory] export destination is the active memory file: {destination}");
                        break;
                    }
                    var directory = Path.GetDirectoryName(destination);
                    if (!string.IsNullOrWhiteSpace(directory))
                        Directory.CreateDirectory(directory);
                    File.Copy(source, destination, overwrite: true);
                    Console.WriteLine($"[memory] exported to {destination}");
                    break;
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
            if (!RecoveryMemoryEnabled)
                return [];

            var path = EffectiveRecoveryMemoryPath();
            lock (RecoveryFileGate)
            {
                RecoveryMemoryReadOnly = false;
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
                    NormalizeRecoveryStore(snapshot);
                    RecoveryCalibration = new(
                        snapshot.Calibration,
                        StringComparer.OrdinalIgnoreCase);
                    return snapshot.Lessons;
                }
                try
                {
                    if (!File.Exists(path))
                    {
                        RecoveryCalibration = new(StringComparer.OrdinalIgnoreCase);
                        _ = SaveRecoveryLessonsWithoutLock(path, []);
                        return [];
                    }

                    var store = ReadRecoveryStoreWithBackup(path);
                    if (store is null)
                        return [];
                    if (store.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] file version {store.Version} is newer than supported version {CurrentMemoryVersion}; memory is read-only.");
                    }

                    NormalizeRecoveryStore(store);
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
                        try
                        {
                            File.Copy(path + ".bak", path, overwrite: true);
                            Console.WriteLine("[memory] restored the primary recovery file from its backup.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[memory] could not restore primary file from backup: {ex.Message}");
                        }
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
            LastRecoveryLoadUsedBackup = false;
            foreach (var candidate in new[] { path, path + ".bak" })
            {
                if (!File.Exists(candidate))
                    continue;
                try
                {
                    var store = JsonSerializer.Deserialize<RecoveryLessonStore>(
                        File.ReadAllText(candidate),
                        PrettyJson);
                    if (store is null)
                        continue;
                    LastRecoveryLoadUsedBackup = !string.Equals(
                        candidate,
                        path,
                        StringComparison.OrdinalIgnoreCase);
                    return store;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[memory] could not load {candidate}: {ex.Message}");
                }
            }

            return null;
        }

        static void NormalizeRecoveryStore(RecoveryLessonStore store)
        {
            var sourceVersion = store.Version <= 0 ? 1 : store.Version;
            store.Version = Math.Min(store.Version <= 0 ? 1 : store.Version, CurrentMemoryVersion);
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

        internal static Mutex CreateRecoveryFileMutex(string path)
        {
            var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
                Path.GetFullPath(path).ToUpperInvariant())));
            return new Mutex(false, $@"Local\RDPilotRecoveryMemory_{hash[..24]}");
        }

        internal static bool WaitForRecoveryMutex(Mutex mutex)
        {
            try { return mutex.WaitOne(TimeSpan.FromSeconds(5)); }
            catch (AbandonedMutexException) { return true; }
        }

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
                    if (diskStore.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] save skipped because file version {diskStore.Version} is newer than supported version {CurrentMemoryVersion}.");
                        return false;
                    }
                    NormalizeRecoveryStore(diskStore);
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
                    RecoveryMemoryDirty = false;
                    PendingRecoveryLessons = null;
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
            string? tempPath = null;
            try
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);

                tempPath = $"{path}.{Guid.NewGuid():N}.tmp";
                var store = new RecoveryLessonStore
                {
                    Version = CurrentMemoryVersion,
                    Lessons = lessons.ToList(),
                    Calibration = new(RecoveryCalibration, StringComparer.OrdinalIgnoreCase)
                };
                File.WriteAllText(tempPath, JsonSerializer.Serialize(store, PrettyJson), Encoding.UTF8);
                if (File.Exists(path))
                {
                    try
                    {
                        File.Replace(tempPath, path, path + ".bak", ignoreMetadataErrors: true);
                    }
                    catch (PlatformNotSupportedException)
                    {
                        File.Copy(path, path + ".bak", overwrite: true);
                        File.Move(tempPath, path, overwrite: true);
                    }
                    catch (IOException)
                    {
                        File.Copy(path, path + ".bak", overwrite: true);
                        File.Move(tempPath, path, overwrite: true);
                    }
                }
                else
                {
                    File.Move(tempPath, path);
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[memory] could not save recovery lessons: {ex.Message}");
                return false;
            }
            finally
            {
                if (tempPath != null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); }
                    catch { }
                }
            }
        }

        static void MarkRecoveryMemoryDirty(List<RecoveryLesson> lessons)
        {
            RecoveryMemoryDirty = true;
            PendingRecoveryLessons = lessons;
        }

        internal static bool FlushPendingRecoveryMemory()
        {
            if (!RecoveryMemoryDirty)
                return true;
            if (PendingRecoveryLessons is null)
            {
                Console.WriteLine("[memory] retrying pending calibration write.");
                PersistCalibrationSnapshot();
                return !RecoveryMemoryDirty;
            }

            Console.WriteLine("[memory] retrying pending durable recovery-memory write.");
            return SaveRecoveryLessons(PendingRecoveryLessons);
        }

        static void PersistCalibrationSnapshot()
        {
            if (!RecoveryMemoryEnabled || RecoveryMemoryReadOnly)
                return;

            var path = EffectiveRecoveryMemoryPath();
            lock (RecoveryFileGate)
            {
                using var mutex = CreateRecoveryFileMutex(path);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                    return;
                try
                {
                    var store = ReadRecoveryStoreWithBackup(path) ?? new RecoveryLessonStore();
                    if (store.Version > CurrentMemoryVersion)
                    {
                        RecoveryMemoryReadOnly = true;
                        Console.WriteLine($"[memory] calibration save skipped because file version {store.Version} is newer than supported version {CurrentMemoryVersion}.");
                        return;
                    }
                    NormalizeRecoveryStore(store);
                    RecoveryCalibration = MergeCalibration(store.Calibration, RecoveryCalibration);
                    if (!SaveRecoveryLessonsWithoutLock(path, store.Lessons))
                        RecoveryMemoryDirty = true;
                    else if (PendingRecoveryLessons is null)
                        RecoveryMemoryDirty = false;
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
                    latest.CountersCompactedBeforeUtc = versions
                        .Select(version => version.CountersCompactedBeforeUtc)
                        .Where(value => value.HasValue)
                        .Max();
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

        static void NormalizeLessonCounters(RecoveryLesson lesson)
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
            IncrementWriterCounter(lesson.SuccessByWriter);
            IncrementWriterCounter(lesson.RewardObservationByWriter);
            IncrementWriterReward(
                lesson.RewardByWriter,
                Math.Clamp(reward, 0, 1));
            NormalizeLessonCounters(lesson);
        }

        static void RecordLessonFailure(RecoveryLesson lesson)
        {
            NormalizeLessonCounters(lesson);
            IncrementWriterCounter(lesson.FailureByWriter);
            IncrementWriterCounter(lesson.RewardObservationByWriter);
            NormalizeLessonCounters(lesson);
        }

        static void RecordLessonSelection(RecoveryLesson lesson)
        {
            NormalizeLessonCounters(lesson);
            IncrementWriterCounter(lesson.SelectionByWriter);
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
            lesson.CountersCompactedBeforeUtc =
                lesson.CountersCompactedBeforeUtc is DateTime existing &&
                existing > cutoff.Value
                    ? existing
                    : cutoff.Value;
            NormalizeLessonCounters(lesson);
        }

        static bool IsLessonActive(RecoveryLesson lesson) =>
            string.Equals(lesson.Status, "active", StringComparison.OrdinalIgnoreCase);

        static List<RecoveryLesson> ApplyLessonRetention(
            List<RecoveryLesson> lessons)
        {
            var archived = new List<RecoveryLesson>();
            var staleCutoff = DateTime.UtcNow.AddDays(-365);
            foreach (var lesson in lessons.Where(IsLessonActive))
            {
                if (lesson.UpdatedUtc < staleCutoff &&
                    lesson.SuccessCount <= 1 &&
                    RecoveryReliability(lesson) < 0.55)
                {
                    lesson.Status = "quarantined";
                    lesson.QuarantinedUtc = DateTime.UtcNow;
                    lesson.LastFailureReason = "stale low-confidence strategy";
                    lesson.UpdatedUtc = DateTime.UtcNow;
                }
            }

            var activeKeep = SelectActiveLessonIdsForRetention(
                lessons.Where(IsLessonActive));
            foreach (var lesson in lessons.Where(lesson =>
                         IsLessonActive(lesson) && !activeKeep.Contains(lesson.Id)))
            {
                lesson.Status = "quarantined";
                lesson.QuarantinedUtc = DateTime.UtcNow;
                lesson.LastFailureReason = "superseded by higher-value strategies at the active-memory limit";
                lesson.UpdatedUtc = DateTime.UtcNow;
            }

            var quarantineKeep = lessons
                .Where(lesson => !IsLessonActive(lesson))
                .OrderByDescending(lesson => lesson.UpdatedUtc)
                .Take(Math.Max(1, RecoveryMemoryMaxQuarantinedLessons))
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
            var globalLimit = Math.Max(1, RecoveryMemoryMaxLessons);
            var reservePerContext = Math.Max(
                0,
                RecoveryMemoryReservedLessonsPerContext);
            var softMaxPerContext = Math.Max(
                Math.Max(1, reservePerContext),
                RecoveryMemorySoftMaxLessonsPerContext);
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
            var maxBytes = Math.Max(1024, RecoveryMemoryMaxFileBytes);
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
                lesson.QuarantinedUtc ??= DateTime.UtcNow;
                lesson.LastFailureReason =
                    "moved to archive to enforce the primary memory file-size limit";
                lesson.UpdatedUtc = DateTime.UtcNow;
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

        internal static string EffectiveRecoveryMemoryArchivePath()
        {
            if (!string.IsNullOrWhiteSpace(RecoveryMemoryArchivePath))
            {
                return Path.GetFullPath(
                    Environment.ExpandEnvironmentVariables(
                        RecoveryMemoryArchivePath.Trim()));
            }

            var primary = EffectiveRecoveryMemoryPath();
            var directory = Path.GetDirectoryName(primary)
                            ?? AppContext.BaseDirectory;
            var name = Path.GetFileNameWithoutExtension(primary);
            return Path.Combine(
                directory,
                $"{name}-archive.json");
        }

        static bool ArchiveRecoveryLessonsWithoutLock(
            string primaryPath,
            IReadOnlyCollection<RecoveryLesson> lessons)
        {
            if (lessons.Count == 0)
                return true;

            var archivePath = EffectiveRecoveryMemoryArchivePath();
            string? tempPath = null;
            try
            {
                if (string.Equals(
                        Path.GetFullPath(primaryPath),
                        Path.GetFullPath(archivePath),
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Recovery memory archive path must differ from the primary memory path.");
                }
                var directory = Path.GetDirectoryName(archivePath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);

                var existing = new RecoveryLessonArchiveStore();
                if (File.Exists(archivePath))
                {
                    existing =
                        JsonSerializer.Deserialize<RecoveryLessonArchiveStore>(
                            File.ReadAllText(archivePath),
                            PrettyJson)
                        ?? new RecoveryLessonArchiveStore();
                }
                existing.Lessons ??= [];
                var merged = existing.Lessons
                    .Concat(lessons)
                    .GroupBy(lesson => lesson.Id, StringComparer.Ordinal)
                    .Select(group => group
                        .OrderByDescending(lesson => lesson.UpdatedUtc)
                        .First())
                    .OrderByDescending(lesson => lesson.UpdatedUtc)
                    .ToList();
                var store = new RecoveryLessonArchiveStore
                {
                    UpdatedUtc = DateTime.UtcNow,
                    Lessons = merged
                };
                var payload = JsonSerializer.SerializeToUtf8Bytes(
                    store,
                    PrettyJson);
                if (File.Exists(archivePath) &&
                    payload.LongLength >
                    Math.Max(1024, RecoveryMemoryArchiveMaxBytes))
                {
                    RotateRecoveryArchive(archivePath);
                    store.Lessons = lessons
                        .GroupBy(
                            lesson => lesson.Id,
                            StringComparer.Ordinal)
                        .Select(group => group
                            .OrderByDescending(lesson => lesson.UpdatedUtc)
                            .First())
                        .ToList();
                    payload = JsonSerializer.SerializeToUtf8Bytes(
                        store,
                        PrettyJson);
                }

                tempPath =
                    $"{archivePath}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
                File.WriteAllBytes(tempPath, payload);
                File.Move(tempPath, archivePath, overwrite: true);
                Console.WriteLine(
                    $"[memory] archived {lessons.Count} displaced lesson(s) to {archivePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[memory] could not archive displaced lessons from {primaryPath}: {ex.Message}");
                return false;
            }
            finally
            {
                if (tempPath is not null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); }
                    catch { }
                }
            }
        }

        static void RotateRecoveryArchive(string archivePath)
        {
            var retained = Math.Clamp(
                RecoveryMemoryArchiveRetainedFiles,
                1,
                20);
            var oldest = $"{archivePath}.{retained}";
            if (File.Exists(oldest))
                File.Delete(oldest);
            for (var index = retained - 1; index >= 1; index--)
            {
                var source = $"{archivePath}.{index}";
                if (File.Exists(source))
                {
                    File.Move(
                        source,
                        $"{archivePath}.{index + 1}",
                        overwrite: true);
                }
            }
            File.Move(
                archivePath,
                $"{archivePath}.1",
                overwrite: true);
        }

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
            IncrementWriterCounter(bucket.CandidateByWriter);
            NormalizeCalibrationBucket(bucket);
            bucket.UpdatedUtc = DateTime.UtcNow;
        }

        static void RegisterCalibrationOutcome(string loopKind, bool confirmed)
        {
            var bucket = CalibrationBucket(loopKind);
            if (confirmed) IncrementWriterCounter(bucket.ConfirmedByWriter);
            else IncrementWriterCounter(bucket.RejectedByWriter);
            NormalizeCalibrationBucket(bucket);
            bucket.UpdatedUtc = DateTime.UtcNow;
        }

        static void RegisterCalibrationInconclusive(string loopKind)
        {
            if (string.IsNullOrWhiteSpace(loopKind))
                return;
            var bucket = CalibrationBucket(loopKind);
            IncrementWriterCounter(bucket.InconclusiveByWriter);
            NormalizeCalibrationBucket(bucket);
            bucket.UpdatedUtc = DateTime.UtcNow;
        }

        static void IncrementWriterCounter(Dictionary<string, int> counters)
        {
            counters.TryGetValue(CalibrationWriterId, out var value);
            counters[CalibrationWriterId] = value + 1;
        }

        static void IncrementWriterReward(
            Dictionary<string, double> rewards,
            double value)
        {
            rewards.TryGetValue(CalibrationWriterId, out var existing);
            rewards[CalibrationWriterId] = existing + value;
        }

        static void NormalizeCalibrationBucket(LoopCalibrationBucket bucket)
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
                    CountersCompactedBeforeUtc =
                        leftBucket.CountersCompactedBeforeUtc >=
                        rightBucket.CountersCompactedBeforeUtc
                            ? leftBucket.CountersCompactedBeforeUtc
                            : rightBucket.CountersCompactedBeforeUtc,
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

        static Dictionary<string, int> MergeWriterCounters(
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

        static Dictionary<string, double> MergeWriterDoubleCounters(
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
            bucket.CountersCompactedBeforeUtc =
                bucket.CountersCompactedBeforeUtc is DateTime existing &&
                existing > cutoff.Value
                    ? existing
                    : cutoff.Value;
            NormalizeCalibrationBucket(bucket);
        }

        static DateTime? WriterCompactionCutoff(IEnumerable<string> writerIds)
        {
            const int maximumComponents = 64;
            const int componentsToKeep = 32;
            var parsed = writerIds
                .Distinct(StringComparer.Ordinal)
                .Select(id => (Id: id, StartedUtc: WriterStartedUtc(id)))
                .Where(item => item.StartedUtc <= DateTime.UtcNow.AddDays(-7))
                .OrderBy(item => item.StartedUtc)
                .ToArray();
            var total = writerIds.Distinct(StringComparer.Ordinal).Count();
            var removeCount = Math.Max(0, total - componentsToKeep);
            if (total <= maximumComponents || removeCount == 0 || parsed.Length == 0)
                return null;
            return parsed[Math.Min(removeCount, parsed.Length) - 1].StartedUtc;
        }

        static DateTime WriterStartedUtc(string writerId)
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

        static int RemoveWriterEntriesThrough(
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

        static double RemoveWriterEntriesThrough(
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

        static void RemoveCompactedWriterEntries(
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

        static void RemoveCompactedWriterEntries(
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

        static double CalibratedLoopThreshold(
            string calibrationKey,
            string? fallbackKey = null)
        {
            if (!RecoveryCalibration.TryGetValue(calibrationKey, out var bucket) &&
                (string.IsNullOrWhiteSpace(fallbackKey) ||
                 !RecoveryCalibration.TryGetValue(fallbackKey, out bucket)))
            {
                return ProactiveLoopConfidenceThreshold;
            }
            var labeled = bucket.ConfirmedCount + bucket.RejectedCount;
            if (labeled < 10)
                return ProactiveLoopConfidenceThreshold;

            var precision = bucket.ConfirmedCount / (double)labeled;
            var adjustment = precision switch
            {
                < 0.60 => 0.10,
                < 0.75 => 0.05,
                > 0.92 => -0.04,
                _ => 0
            };
            return Math.Clamp(ProactiveLoopConfidenceThreshold + adjustment, 0.55, 0.95);
        }

        static void AppendLoopTelemetry(
            string eventName,
            int step,
            string loopKind,
            string loopTopology,
            string interactionDomain,
            LoopDetectionAssessment assessment,
            bool? confirmed,
            object? details = null)
        {
            if (!RecoveryMemoryEnabled)
                return;
            try
            {
                var memoryPath = EffectiveRecoveryMemoryPath();
                var directory = Path.GetDirectoryName(memoryPath);
                if (string.IsNullOrWhiteSpace(directory))
                    return;
                Directory.CreateDirectory(directory);
                var telemetryPath = Path.Combine(directory, "loop-telemetry.jsonl");
                var payload = JsonSerializer.Serialize(new
                {
                    timestampUtc = DateTime.UtcNow,
                    @event = eventName,
                    runId = assessment.RunId,
                    step,
                    loopKind,
                    loopTopology,
                    interactionDomain,
                    assessment.Confidence,
                    assessment.DecisionThreshold,
                    assessment.CycleLength,
                    assessment.MatchingPriorStates,
                    assessment.GraphCycle,
                    assessment.RepeatedActionCycle,
                    assessment.ConsistentReturnPeriod,
                    assessment.SemanticCycle,
                    assessment.IsProductiveCycle,
                    assessment.CycleDisposition,
                    assessment.IndependentlyConfirmed,
                    assessment.CalibrationKey,
                    assessment.Evidence,
                    confirmed,
                    details
                });
                AppendTelemetryPayload(memoryPath, telemetryPath, payload);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[memory] could not append loop telemetry: {ex.Message}");
            }
        }

        internal static void AppendLoopReplayObservation(
            string runId,
            int step,
            int screenWidth,
            int screenHeight,
            byte[] screenFingerprint,
            byte[] activeWindowFingerprint,
            UiPromptContext context,
            ResolvedActionSnapshot? previousAction,
            double lastDelta,
            LoopDetectionAssessment assessment,
            string goalMode,
            bool recurringWorkflowIntent)
        {
            if (!RecoveryMemoryEnabled)
                return;
            try
            {
                var memoryPath = EffectiveRecoveryMemoryPath();
                var directory = Path.GetDirectoryName(memoryPath);
                if (string.IsNullOrWhiteSpace(directory))
                    return;
                Directory.CreateDirectory(directory);
                var telemetryPath = Path.Combine(directory, "loop-telemetry.jsonl");
                var replayFrame = new LoopReplayFrame
                {
                    ScreenFingerprintBase64 =
                        Convert.ToBase64String(screenFingerprint),
                    ActiveWindowFingerprintBase64 =
                        Convert.ToBase64String(activeWindowFingerprint),
                    ActiveProcess = NormalizeText(context.ActiveProcessName),
                    WindowTitle = ReplaySafeTokens(context.ActiveWindowTitle),
                    FocusSummary = ReplaySafeTokens(context.FocusedUiaSummary),
                    PreviousAction = ReplaySafeAction(previousAction),
                    LastDelta = double.IsFinite(lastDelta)
                        ? lastDelta
                        : null
                };
                var payload = JsonSerializer.Serialize(new
                {
                    timestampUtc = DateTime.UtcNow,
                    @event = "observation",
                    runId,
                    step,
                    screenWidth,
                    screenHeight,
                    assessment.Confidence,
                    assessment.IndependentlyConfirmed,
                    goalMode,
                    recurringWorkflowIntent,
                    replayFrame
                });
                AppendTelemetryPayload(memoryPath, telemetryPath, payload);
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[memory] could not append replay observation: {ex.Message}");
            }
        }

        static void AppendTelemetryPayload(
            string memoryPath,
            string telemetryPath,
            string payload)
        {
            lock (RecoveryFileGate)
            {
                using var mutex = CreateRecoveryFileMutex(memoryPath);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                    return;
                try
                {
                    RotateTelemetryIfNeeded(
                        telemetryPath,
                        Encoding.UTF8.GetByteCount(payload) +
                        Encoding.UTF8.GetByteCount(Environment.NewLine));
                    File.AppendAllText(
                        telemetryPath,
                        payload + Environment.NewLine,
                        Encoding.UTF8);
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }

        static string ReplaySafeTokens(string? value)
        {
            var tokens = Regex.Matches(
                    NormalizeText(value),
                    @"[\p{L}\p{Nd}]{2,}")
                .Cast<Match>()
                .Select(match => match.Value)
                .Distinct(StringComparer.Ordinal)
                .Take(24)
                .Select(token =>
                {
                    var hash = SHA256.HashData(Encoding.UTF8.GetBytes(token));
                    return Convert.ToHexString(hash.AsSpan(0, 5))
                        .ToLowerInvariant();
                });
            return string.Join(' ', tokens);
        }

        static ActionDto? ReplaySafeAction(
            ResolvedActionSnapshot? snapshot)
        {
            if (snapshot is null)
                return null;

            var source = snapshot.Action;
            return new ActionDto
            {
                Type = source.Type,
                X = source.X,
                Y = source.Y,
                XPx = source.XPx,
                YPx = source.YPx,
                ToX = source.ToX,
                ToY = source.ToY,
                ToXPx = source.ToXPx,
                ToYPx = source.ToYPx,
                Button = source.Button,
                Keys = source.Keys?
                    .Select(ReplaySafeKey)
                    .ToArray(),
                Text = source.Text is null ? null : "<redacted>",
                UiaIndex = source.UiaIndex,
                ScrollDy = source.ScrollDy,
                BBox = source.BBox,
                ToBBox = source.ToBBox,
                Crop = source.Crop,
                DragDurationMs = source.DragDurationMs,
                WaitSeconds = source.WaitSeconds,
                Note = ReplaySafeTokens(snapshot.SemanticTokens)
            };
        }

        static string ReplaySafeKey(string? key)
        {
            var normalized = NormalizeText(key);
            return Regex.IsMatch(
                normalized,
                @"^(ctrl|control|alt|shift|win|enter|escape|esc|tab|space|backspace|delete|del|insert|ins|home|end|pageup|pagedown|pgup|pgdn|up|down|left|right|arrowup|arrowdown|arrowleft|arrowright|f(?:[1-9]|1[0-2]))$",
                RegexOptions.CultureInvariant)
                ? normalized
                : "<text-key>";
        }

        internal static void TryAutoExportLoopReplayCorpus()
        {
            try
            {
                ExportLoopTelemetryToReplayCorpus(
                    EffectiveLoopReplayCorpusPath(),
                    quiet: true);
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[memory] could not auto-export loop replay corpus: {ex.Message}");
            }
        }

        internal static void ExportLoopTelemetryToReplayCorpus(
            string destinationPath,
            bool quiet = false)
        {
            var destination = Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(destinationPath));
            lock (RecoveryFileGate)
            {
                using var mutex = CreateRecoveryFileMutex(destination);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                {
                    throw new IOException(
                        "Could not acquire the loop replay corpus lock.");
                }
                try
                {
                    ExportLoopTelemetryToReplayCorpusCore(
                        destination,
                        quiet);
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }

        static void ExportLoopTelemetryToReplayCorpusCore(
            string destination,
            bool quiet)
        {
            var existing = File.Exists(destination)
                ? JsonSerializer.Deserialize<LoopReplayCorpus>(
                      File.ReadAllText(destination),
                      PrettyJson)
                : null;
            var corpus = BuildLoopReplayCorpus(
                ReadLoopTelemetryLines(),
                existing);
            if (corpus.Cases.Count == 0)
            {
                if (!quiet)
                    Console.WriteLine(
                        "[loop-replay] no labelled telemetry sequences are available yet.");
                return;
            }

            var directory = Path.GetDirectoryName(destination);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);
            var temporary = destination + $".tmp.{Environment.ProcessId}.{Guid.NewGuid():N}";
            try
            {
                File.WriteAllText(
                    temporary,
                    JsonSerializer.Serialize(corpus, PrettyJson),
                    Encoding.UTF8);
                File.Move(temporary, destination, overwrite: true);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }

            if (!quiet)
            {
                Console.WriteLine(
                    $"[loop-replay] exported {corpus.Cases.Count} case(s) to {destination}");
            }
        }

        internal static LoopReplayCorpus BuildLoopReplayCorpus(
            IEnumerable<string> telemetryLines,
            LoopReplayCorpus? existing = null)
        {
            var envelopes = new List<LoopTelemetryReplayEnvelope>();
            foreach (var line in telemetryLines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;
                try
                {
                    var envelope =
                        JsonSerializer.Deserialize<LoopTelemetryReplayEnvelope>(
                            line,
                            PrettyJson);
                    if (envelope is not null &&
                        !string.IsNullOrWhiteSpace(envelope.RunId))
                    {
                        envelopes.Add(envelope);
                    }
                }
                catch (JsonException)
                {
                    // A partially written or older line must not prevent export
                    // of the remaining telemetry.
                }
            }

            var existingCases = existing?.Cases ?? [];
            var manualCases = existingCases
                .Where(replayCase =>
                    string.IsNullOrWhiteSpace(replayCase.LabelSource) ||
                    !replayCase.LabelSource.StartsWith(
                        "telemetry:",
                        StringComparison.Ordinal))
                .ToList();
            var previousGeneratedCases = existingCases
                .Where(replayCase =>
                    replayCase.LabelSource?.StartsWith(
                        "telemetry:",
                        StringComparison.Ordinal) == true)
                .ToList();
            var generatedCases = new List<LoopReplayCase>();
            foreach (var run in envelopes.GroupBy(item => item.RunId))
            {
                var observations = run
                    .Where(item =>
                        item.Event == "observation" &&
                        item.ReplayFrame is not null)
                    .GroupBy(item => item.Step)
                    .Select(group => group
                        .OrderByDescending(item => item.TimestampUtc)
                        .First())
                    .OrderBy(item => item.Step)
                    .ToList();
                if (observations.Count < 2)
                    continue;

                var firstPositiveStep = run
                    .Where(item =>
                        item.IndependentlyConfirmed ||
                        item.Confirmed == true)
                    .Select(item => item.Step)
                    .Where(step => step > 0)
                    .DefaultIfEmpty(0)
                    .Min();
                var expectedLoop = firstPositiveStep > 0;
                var quietNegative =
                    observations.Count >= 6 &&
                    observations.All(item => item.Confidence < 0.5);
                if (!expectedLoop && !quietNegative)
                    continue;

                var endStep = expectedLoop
                    ? firstPositiveStep
                    : observations[^1].Step;
                var selected = observations
                    .Where(item => item.Step <= endStep)
                    .TakeLast(24)
                    .ToList();
                if (selected.Count < 2)
                    continue;

                var capturedUtc = selected
                    .Select(item => item.TimestampUtc)
                    .Where(value => value != default)
                    .DefaultIfEmpty(DateTime.UtcNow)
                    .Max();
                generatedCases.Add(new LoopReplayCase
                {
                    Name =
                        $"telemetry:{run.Key}:{(expectedLoop ? "positive" : "negative")}",
                    ExpectedLoop = expectedLoop,
                    LabelSource = expectedLoop
                        ? "telemetry:independent_confirmation"
                        : "telemetry:quiet_run",
                    CapturedUtc = capturedUtc,
                    GoalMode = string.IsNullOrWhiteSpace(
                        selected[^1].GoalMode)
                        ? "finite"
                        : selected[^1].GoalMode,
                    RecurringWorkflowIntent =
                        selected[^1].RecurringWorkflowIntent,
                    ScreenWidth = Math.Max(1, selected[^1].ScreenWidth),
                    ScreenHeight = Math.Max(1, selected[^1].ScreenHeight),
                    Frames = selected
                        .Select(item => item.ReplayFrame!)
                        .ToList()
                });
            }

            return new LoopReplayCorpus
            {
                Cases = manualCases
                    .Concat(previousGeneratedCases
                        .Concat(generatedCases)
                        .GroupBy(
                            replayCase =>
                            {
                                var separator =
                                    replayCase.Name.LastIndexOf(':');
                                return separator > 0
                                    ? replayCase.Name[..separator]
                                    : replayCase.Name;
                            },
                            StringComparer.Ordinal)
                        .Select(group => group
                            .OrderByDescending(item => item.ExpectedLoop)
                            .ThenByDescending(item => item.CapturedUtc)
                            .First())
                        .OrderByDescending(item => item.CapturedUtc)
                        .Take(200))
                    .ToList()
            };
        }

        static IEnumerable<string> ReadLoopTelemetryLines()
        {
            var memoryPath = EffectiveRecoveryMemoryPath();
            var directory = Path.GetDirectoryName(memoryPath);
            if (string.IsNullOrWhiteSpace(directory))
                return [];
            var telemetryPath = Path.Combine(directory, "loop-telemetry.jsonl");
            var paths = new List<string>();
            for (var index = RecoveryTelemetryRetainedFiles;
                 index >= 1;
                 index--)
            {
                var rotated = $"{telemetryPath}.{index}";
                if (File.Exists(rotated))
                    paths.Add(rotated);
            }
            if (File.Exists(telemetryPath))
                paths.Add(telemetryPath);

            var lines = new List<string>();
            lock (RecoveryFileGate)
            {
                using var mutex = CreateRecoveryFileMutex(memoryPath);
                var lockTaken = WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                    return [];
                try
                {
                    foreach (var path in paths)
                        lines.AddRange(File.ReadLines(path));
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
            return lines;
        }

        static void RotateTelemetryIfNeeded(string telemetryPath, int incomingBytes)
        {
            if (!File.Exists(telemetryPath) ||
                new FileInfo(telemetryPath).Length + incomingBytes <=
                Math.Max(65536, RecoveryTelemetryMaxBytes))
            {
                return;
            }

            var retained = Math.Clamp(RecoveryTelemetryRetainedFiles, 1, 20);
            var oldest = $"{telemetryPath}.{retained}";
            if (File.Exists(oldest))
                File.Delete(oldest);
            for (var index = retained - 1; index >= 1; index--)
            {
                var source = $"{telemetryPath}.{index}";
                if (!File.Exists(source))
                    continue;
                File.Move(source, $"{telemetryPath}.{index + 1}", overwrite: true);
            }
            File.Move(telemetryPath, $"{telemetryPath}.1", overwrite: true);
        }
    }
}
