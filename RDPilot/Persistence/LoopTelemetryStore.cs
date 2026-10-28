namespace RDPilot.Persistence;

/// <summary>Owns telemetry append, rotation and consistent snapshots under the telemetry file lock.</summary>
internal static class LoopTelemetryStore
{
    internal static void Append(RecoveryStoragePaths paths, string payload)
    {
        paths.Validate();
        using var lease = StorageFileLock.Acquire(paths.Telemetry);
        Directory.CreateDirectory(Path.GetDirectoryName(paths.Telemetry)!);
        var line = payload + Environment.NewLine;
        var incoming = Encoding.UTF8.GetBytes(line);
        if (File.Exists(paths.Telemetry) &&
            new FileInfo(paths.Telemetry).Length + incoming.LongLength > Math.Max(65536, Options.RecoveryTelemetryMaxBytes))
        {
            // Include the new record in the staged file, so a failed append cannot leave a completed rotation.
            var contents = Encoding.UTF8.GetPreamble().Concat(incoming).ToArray();
            RotatingFileStore.Write(paths.Telemetry, contents,
                Enumerable.Range(1, paths.RetainedTelemetry).Select(paths.TelemetryGeneration).ToArray());
        }
        else
            File.AppendAllText(paths.Telemetry, line, Encoding.UTF8);
    }

    internal static string[] Read(RecoveryStoragePaths paths)
    {
        paths.Validate();
        using var lease = StorageFileLock.Acquire(paths.Telemetry);
        // Enumerate generations only after acquiring the same lock used by rotation.
        var files = Enumerable.Range(1, paths.RetainedTelemetry).Reverse()
            .Select(paths.TelemetryGeneration).Append(paths.Telemetry);
        return files.Where(File.Exists).SelectMany(File.ReadLines).ToArray();
    }
}