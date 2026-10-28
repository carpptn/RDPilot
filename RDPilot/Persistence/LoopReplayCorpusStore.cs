namespace RDPilot.Persistence;

/// <summary>Validates corpus documents and owns atomic import/export under the destination lock.</summary>
internal static class LoopReplayCorpusStore
{
    internal static LoopReplayCorpus Read(string path)
    {
        var json = File.ReadAllText(path);
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("Replay corpus must be an object with a cases array.");
        var properties = root.EnumerateObject().ToArray();
        var cases = properties.Where(p => p.Name.Equals("cases", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (cases.Length != 1 || cases[0].Value.ValueKind != JsonValueKind.Array ||
            properties.Any(p => p.Name.Equals("lessons", StringComparison.OrdinalIgnoreCase) ||
                                p.Name.Equals("prompts", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("The existing document is not a replay corpus; refusing to replace it.");
        var corpus = JsonSerializer.Deserialize<LoopReplayCorpus>(json, PrettyJson)
            ?? throw new InvalidDataException("Replay corpus is null.");
        foreach (var replayCase in corpus.Cases)
            ReplayCaseValidator.Validate(replayCase);
        return corpus;
    }

    static LoopReplayCorpus ReadExisting(string path)
    {
        // File.Exists hides access errors; only a genuinely missing file may start an empty corpus.
        try { return Read(path); }
        catch (FileNotFoundException) { return new(); }
        catch (DirectoryNotFoundException) { return new(); }
    }

    internal static void Import(string sourcePath)
    {
        var paths = RecoveryStoragePaths.ResolveValidated(Options);
        var destination = paths.ValidateCorpusDestination(paths.Corpus);
        var source = Path.GetFullPath(Environment.ExpandEnvironmentVariables(sourcePath));
        var imported = Read(source);
        if (imported.Cases.Count == 0)
            throw new InvalidOperationException("Independent loop replay corpus does not contain any cases.");
        foreach (var replayCase in imported.Cases)
        {
            if (string.IsNullOrWhiteSpace(replayCase.Name))
                throw new InvalidOperationException("Every imported replay case requires a stable name.");
            if (!replayCase.HasIndependentLabel)
                throw new InvalidOperationException(
                    $"Replay case '{replayCase.Name}' requires a non-telemetry labelSource such as 'manual:reviewed'.");
        }

        using var lease = StorageFileLock.Acquire(destination);
        var existing = ReadExisting(destination);
        var importedNames = imported.Cases.Select(item => item.Name).ToHashSet(StringComparer.Ordinal);
        var output = new LoopReplayCorpus
        {
            Cases = existing.Cases.Where(item => !importedNames.Contains(item.Name)).Concat(imported.Cases).ToList()
        };
        AtomicTextStore.Write(destination, JsonSerializer.Serialize(output, PrettyJson));
        Console.WriteLine($"[loop-replay] imported {imported.Cases.Count} independently labelled case(s) into {destination}");
    }

    internal static void Export(string destinationPath, bool quiet = false)
    {
        var paths = RecoveryStoragePaths.ResolveValidated(Options);
        var destination = paths.ValidateCorpusDestination(destinationPath);
        // Take a telemetry snapshot first, so corpus and telemetry locks are never nested.
        var telemetry = LoopTelemetryStore.Read(paths);
        using var lease = StorageFileLock.Acquire(destination);
        var existing = ReadExisting(destination);
        var corpus = BuildLoopReplayCorpus(telemetry, existing);
        if (corpus.Cases.Count == 0)
        {
            if (!quiet) Console.WriteLine("[loop-replay] no labelled telemetry sequences are available yet.");
            return;
        }
        AtomicTextStore.Write(destination, JsonSerializer.Serialize(corpus, PrettyJson));
        if (!quiet) Console.WriteLine($"[loop-replay] exported {corpus.Cases.Count} case(s) to {destination}");
    }
}
