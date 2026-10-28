namespace RDPilot.Persistence;

/// <summary>Defines every recovery storage file and validates destinations before mutation.</summary>
internal sealed class RecoveryStoragePaths
{
    internal string Primary { get; }
    internal string Backup => BackupFor(Primary);
    internal string Archive { get; }
    internal int RetainedArchives { get; }
    internal string Telemetry { get; }
    internal string Corpus { get; }
    internal int RetainedTelemetry { get; }

    RecoveryStoragePaths(string primary, string archive, string corpus, AppOptions options)
    {
        Primary = primary;
        Archive = archive;
        RetainedArchives = Math.Clamp(options.RecoveryMemoryArchiveRetainedFiles, 1, 20);
        RetainedTelemetry = Math.Clamp(options.RecoveryTelemetryRetainedFiles, 1, 20);
        Telemetry = Path.Combine(Path.GetDirectoryName(primary)!, "loop-telemetry.jsonl");
        Corpus = corpus;
    }

    internal static string ResolvePrimary(AppOptions options) =>
        string.IsNullOrWhiteSpace(options.RecoveryMemoryPath)
            ? Path.Combine(AppContext.BaseDirectory, "memory", "recovery-memory.json")
            : ResolvePath(options.RecoveryMemoryPath);

    internal static RecoveryStoragePaths Resolve(AppOptions options, string? primaryPath = null)
    {
        var primary = primaryPath is null ? ResolvePrimary(options) : ResolvePath(primaryPath);
        var archive = string.IsNullOrWhiteSpace(options.RecoveryMemoryArchivePath)
            ? Path.Combine(Path.GetDirectoryName(primary) ?? AppContext.BaseDirectory,
                $"{Path.GetFileNameWithoutExtension(primary)}-archive.json")
            : ResolvePath(options.RecoveryMemoryArchivePath);
        var corpus = string.IsNullOrWhiteSpace(options.LoopReplayCorpusPath)
            ? Path.Combine(Path.GetDirectoryName(primary)!, "loop-replay-corpus.json")
            : ResolvePath(options.LoopReplayCorpusPath);
        return new(primary, archive, corpus, options);
    }

    internal static RecoveryStoragePaths ResolveValidated(AppOptions options, string? primaryPath = null)
    {
        var paths = Resolve(options, primaryPath);
        paths.Validate();
        return paths;
    }

    internal static string ResolvePath(string path, string? baseDirectory = null)
    {
        var expanded = Environment.ExpandEnvironmentVariables(path.Trim());
        RejectDevicePath(expanded);
        var resolved = Path.IsPathFullyQualified(expanded)
            ? Path.GetFullPath(expanded)
            : Path.GetFullPath(expanded, baseDirectory ?? AppContext.BaseDirectory);
        RejectDevicePath(resolved);
        return resolved;
    }

    static void RejectDevicePath(string path)
    {
        // Device namespaces have different normalization rules; removing their prefix
        // can change the target and undermine primary/backup/archive collision checks.
        var separators = path.Replace('/', '\\');
        if (separators.StartsWith(@"\\?\", StringComparison.Ordinal) ||
            separators.StartsWith(@"\\.\", StringComparison.Ordinal) ||
            separators.StartsWith(@"\??\", StringComparison.Ordinal))
            throw new ArgumentException(
                "Windows device paths are not supported for storage. Use a regular drive or UNC path.", nameof(path));
    }

    internal static string BackupFor(string primary) => primary + ".bak";

    internal string ArchiveGeneration(int generation)
    {
        if (generation < 1 || generation > RetainedArchives)
            throw new ArgumentOutOfRangeException(nameof(generation));
        return $"{Archive}.{generation}";
    }

    internal string TelemetryGeneration(int generation)
    {
        if (generation < 1 || generation > RetainedTelemetry)
            throw new ArgumentOutOfRangeException(nameof(generation));
        return $"{Telemetry}.{generation}";
    }

    IEnumerable<(string Role, string Path)> Files()
    {
        yield return ("memory", Primary);
        yield return ("backup", Backup);
        yield return ("archive", Archive);
        for (var i = 1; i <= RetainedArchives; i++)
            yield return ($"archive generation {i}", ArchiveGeneration(i));
        yield return ("telemetry", Telemetry);
        for (var i = 1; i <= RetainedTelemetry; i++)
            yield return ($"telemetry generation {i}", TelemetryGeneration(i));
        yield return ("replay corpus", Corpus);
        yield return ("prompt history", PromptHistoryService.EffectivePath());
    }

    internal void Validate()
    {
        var seen = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (role, path) in Files())
        {
            if (seen.TryGetValue(path, out var previous))
                throw new InvalidOperationException(
                    $"Recovery storage path collision: {role} '{path}' conflicts with {previous}. Choose separate storage paths.");
            seen.Add(path, role);
        }
    }

    internal string ValidateCorpusDestination(string destination) =>
        ValidateDestination(destination, "replay corpus");

    internal string ValidateMemoryExportDestination(string destination) =>
        ValidateDestination(destination, "memory");

    string ValidateDestination(string destination, string allowedRole)
    {
        Validate();
        // Explicit export paths have always been relative to the working directory.
        var path = ResolvePath(destination, Environment.CurrentDirectory);
        foreach (var (role, reserved) in Files())
            if (role != allowedRole && StringComparer.OrdinalIgnoreCase.Equals(path, reserved))
                throw new InvalidOperationException(
                    $"Recovery storage path collision: export destination '{path}' conflicts with {role}.");
        return path;
    }
}
