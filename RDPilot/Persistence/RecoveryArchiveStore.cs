namespace RDPilot.Persistence;

/// <summary>Serializes archive merges and rotation across all memory writers sharing the archive.</summary>
internal static class RecoveryArchiveStore
{
    internal static bool Append(
        string primaryPath,
        IReadOnlyCollection<RecoveryLesson> lessons)
    {
        if (lessons.Count == 0)
            return true;

        try
        {
            var paths = RecoveryStoragePaths.ResolveValidated(Options, primaryPath);
            using var archiveLock = StorageFileLock.Acquire(paths.Archive);
            var archivePath = paths.Archive;
            var directory = Path.GetDirectoryName(archivePath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            var existing = RecoveryArchiveReader.ReadExisting(archivePath);
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
                UpdatedUtc = RuntimeClock.UtcNow,
                Lessons = merged
            };
            var payload = JsonSerializer.SerializeToUtf8Bytes(
                store,
                PrettyJson);
            var rotate = File.Exists(archivePath) &&
                payload.LongLength >
                Math.Max(1024, Options.RecoveryMemoryArchiveMaxBytes);
            if (rotate)
            {
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

            RotatingFileStore.Write(archivePath, payload, rotate
                ? Enumerable.Range(1, paths.RetainedArchives).Select(paths.ArchiveGeneration).ToArray()
                : []);
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
    }
}
