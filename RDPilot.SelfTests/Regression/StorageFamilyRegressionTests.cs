using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class StorageFamilyRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
        new[] { "memory-telemetry", "memory-telemetry-generation", "archive-telemetry",
            "archive-telemetry-generation", "corpus-memory", "corpus-backup", "corpus-archive",
            "corpus-archive-generation", "corpus-telemetry", "corpus-telemetry-generation" }
        .Select(kind => ($"storage rejects {kind} collision before writes", (Action)(() => Collision(kind))))
        .Concat(new[] { "memory", "backup", "archive", "archive-generation", "telemetry", "telemetry-generation" }
            .Select(kind => ($"replay export preserves reserved {kind} file", (Action)(() => ExportCollision(kind)))))
        .Concat(new[] { "memory", "empty-object", "null-cases", "duplicate-cases", "null-case", "invalid-frame" }
            .SelectMany(kind => new[] { false, true }.Select(export =>
                ($"replay {(export ? "export" : "import")} preserves invalid {kind} destination",
                (Action)(() => ForeignCorpus(kind, export))))))
        .Concat(new (string, Action)[]
        {
            ("CLI prune rejects telemetry collision without mutation", () => CliCollision(false)),
            ("CLI import rejects memory collision even with recovery disabled", () => CliCollision(true)),
            ("CLI replay export rejects backup collision", CliExportCollision),
            ("direct telemetry append preserves conflicting memory", DirectTelemetryCollision),
            ("memory export cannot replace an archive", MemoryExportCollision),
            ("ordinary replay export preserves manual cases", ValidExport),
            ("export paths retain working-directory resolution", RelativeExport),
            ("export destinations reject device namespace variants", DeviceExport),
            ("storage paths cannot overwrite prompt history", HistoryCollision),
            ("shared archives serialize writers using different memory paths", () => SharedWriters("archive")),
            ("shared archive rotation retains both writers", () => SharedWriters("rotation")),
            ("shared telemetry serializes writers using different memory paths", () => SharedWriters("telemetry"))
        }).ToArray();

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class Files : IDisposable
    {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string Memory => Path.Combine(Root, "memory.json");
        internal string Archive => Path.Combine(Root, "archive.json");
        internal string Corpus => Path.Combine(Root, "corpus.json");
        internal string Telemetry => Path.Combine(Root, "loop-telemetry.jsonl");
        internal string Source => Path.Combine(Root, "source.json");
        internal Files()
        {
            Directory.CreateDirectory(Root);
            Write(Memory, new RecoveryLessonStore { Version = 7, Lessons = [new() { Id = "keep", Status = "active" }] });
            File.Copy(Memory, Memory + ".bak");
            Write(Archive, new RecoveryLessonArchiveStore());
            File.Copy(Archive, Archive + ".1");
            Text(Telemetry, "");
            Text(Telemetry + ".1", "");
            Write(Corpus, new LoopReplayCorpus { Cases = [Case("existing")] });
            Write(Source, new LoopReplayCorpus { Cases = [Case("new")] });
        }
        internal void Configure()
        {
            A.Options.RecoveryMemoryEnabled = true;
            A.Options.RecoveryMemoryPath = Memory;
            A.Options.RecoveryMemoryArchivePath = Archive;
            A.Options.LoopReplayCorpusPath = Corpus;
            A.Options.RecoveryMemoryArchiveRetainedFiles = 2;
            A.Options.RecoveryTelemetryRetainedFiles = 2;
        }
        internal Dictionary<string, byte[]> Snapshot() => Directory.GetFiles(Root)
            .ToDictionary(path => path, File.ReadAllBytes, StringComparer.OrdinalIgnoreCase);
        internal void Unchanged(Dictionary<string, byte[]> before)
        {
            var after = Snapshot();
            Check(before.Count == after.Count && before.All(p => after.TryGetValue(p.Key, out var bytes) && bytes.SequenceEqual(p.Value)),
                "rejected operation changed storage files");
        }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }

    static void Text(string path, string value) => File.WriteAllText(path,
        value.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
    static void Write(string path, object value) => Text(path, JsonSerializer.Serialize(value, A.PrettyJson));
    static LoopReplayCase Case(string name) => new()
    {
        Name = name, LabelSource = "manual:reviewed",
        Frames = [new() { ScreenFingerprintBase64 = "AQID" }, new() { ScreenFingerprintBase64 = "AQID" }]
    };

    static void Reject(Action action)
    {
        try { action(); }
        catch (Exception ex) when (ex is InvalidOperationException or InvalidDataException or JsonException) { return; }
        throw new InvalidOperationException("unsafe storage operation succeeded");
    }

    static string Reserved(Files f, string kind) => kind switch
    {
        "memory" => f.Memory, "backup" => f.Memory + ".bak", "archive" => f.Archive,
        "archive-generation" => f.Archive + ".1", "telemetry" => f.Telemetry,
        _ => f.Telemetry + ".1"
    };

    static void Collision(string kind)
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        var dash = kind.IndexOf('-');
        var owner = kind[..dash];
        var target = Reserved(f, kind[(dash + 1)..]);
        switch (owner)
        {
            case "memory": A.Options.RecoveryMemoryPath = target; break;
            case "archive": A.Options.RecoveryMemoryArchivePath = target; break;
            default: A.Options.LoopReplayCorpusPath = target; break;
        }
        var before = f.Snapshot();
        Reject(() => RecoveryStoragePaths.ResolveValidated(A.Options));
        if (owner == "corpus") Reject(() => LoopReplayService.ImportIndependentLoopReplayCorpus(f.Source));
        else Reject(() => RecoveryMemoryService.LoadRecoveryLessons());
        f.Unchanged(before);
    }

    static void ExportCollision(string kind)
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        var before = f.Snapshot();
        Reject(() => RecoveryMemoryService.ExportLoopTelemetryToReplayCorpus(Reserved(f, kind)));
        f.Unchanged(before);
    }

    static void ForeignCorpus(string kind, bool export)
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        Text(f.Corpus, kind switch
        {
            "memory" => File.ReadAllText(f.Memory),
            "empty-object" => "{}",
            "null-cases" => "{\"cases\":null}",
            "duplicate-cases" => "{\"cases\":[],\"Cases\":[]}",
            "null-case" => "{\"cases\":[null]}",
            _ => "{\"cases\":[{\"name\":\"bad\",\"frames\":[{},{}]}]}"
        });
        var before = f.Snapshot();
        Reject(() =>
        {
            if (export) RecoveryMemoryService.ExportLoopTelemetryToReplayCorpus(f.Corpus);
            else LoopReplayService.ImportIndependentLoopReplayCorpus(f.Source);
        });
        f.Unchanged(before);
    }

    static void DirectTelemetryCollision()
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        File.Copy(f.Memory, f.Telemetry, overwrite: true);
        A.Options.RecoveryMemoryPath = f.Telemetry;
        var before = f.Snapshot();
        RecoveryMemoryService.AppendLoopReplayObservation("run", 1, 320, 200, [1, 2, 3], [1, 2, 3],
            new("", "", null, null, null, null), null, 0, new(false, 0, 0, 0, false, false, ""), null, "finite", false);
        f.Unchanged(before);
    }

    static void MemoryExportCollision()
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        var before = f.Snapshot();
        Reject(() => RecoveryMemoryService.ExecuteRecoveryMemoryMaintenance("export", f.Archive));
        f.Unchanged(before);
    }

    static void ValidExport()
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        RecoveryMemoryService.ExportLoopTelemetryToReplayCorpus(f.Corpus);
        Check(LoopReplayCorpusStore.Read(f.Corpus).Cases.Single().Name == "existing", "export lost a reviewed case");
    }

    static void RelativeExport()
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        var paths = RecoveryStoragePaths.ResolveValidated(A.Options);
        Check(paths.ValidateCorpusDestination("relative-export.json") ==
            Path.Combine(Environment.CurrentDirectory, "relative-export.json"), "export base directory changed");
        var destination = Path.Combine(f.Root, "new-export.json");
        RecoveryMemoryService.ExportLoopTelemetryToReplayCorpus(destination);
        Check(!File.Exists(destination), "empty telemetry unexpectedly created a corpus");
    }

    static void DeviceExport()
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        var before = f.Snapshot();
        foreach (var prefix in new[] { @"\\?\", @"\\.\", @"\??\", "//?/", "//./", "/??/" })
        {
            var rejected = false;
            try { RecoveryMemoryService.ExportLoopTelemetryToReplayCorpus(prefix + f.Corpus); }
            catch (ArgumentException) { rejected = true; }
            Check(rejected, "device export path accepted");
        }
        f.Unchanged(before);
    }

    static void HistoryCollision()
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        A.Options.LoopReplayCorpusPath = PromptHistoryService.EffectivePath();
        Reject(() => RecoveryStoragePaths.ResolveValidated(A.Options));
    }

    static Process Start(params string[] args)
    {
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        foreach (var arg in args) start.ArgumentList.Add(arg);
        return Process.Start(start)!;
    }

    static int Complete(Process process)
    {
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        try
        {
            Check(process.WaitForExit(20000), "storage worker timed out");
            Task.WhenAll(output, error).GetAwaiter().GetResult();
            if (process.ExitCode != 0) Console.WriteLine(output.Result + error.Result);
            return process.ExitCode;
        }
        finally { if (!process.HasExited) { process.Kill(); process.WaitForExit(); } }
    }

    static string[] CliArgs(Files f) => ["--command-main-probe", "--recovery-memory-path", f.Memory,
        "--recovery-archive-path", f.Archive, "--loop-replay-corpus", f.Corpus];

    static void CliCollision(bool import)
    {
        using var f = new Files();
        var args = CliArgs(f).Concat(import
            ? new[] { "--no-recovery-memory", "--loop-replay-corpus", f.Memory, "--loop-replay-import", f.Source }
            : new[] { "--recovery-memory", "--recovery-memory-path", f.Telemetry, "--memory-prune" }).ToArray();
        var before = f.Snapshot();
        using var child = Start(args);
        Check(Complete(child) == 2, "CLI accepted storage collision");
        f.Unchanged(before);
    }

    static void CliExportCollision()
    {
        using var f = new Files();
        var before = f.Snapshot();
        using var child = Start(CliArgs(f).Concat(new[] { "--loop-replay-export", f.Memory + ".bak" }).ToArray());
        Check(Complete(child) == 2, "CLI export accepted reserved destination");
        f.Unchanged(before);
    }

    internal static int Worker(string[] args)
    {
        using var scope = A.BeginRun();
        var root = args[1];
        var writer = args[2];
        var mode = args[3];
        A.Options.RecoveryMemoryEnabled = true;
        A.Options.RecoveryMemoryPath = Path.Combine(root, writer + ".json");
        A.Options.RecoveryMemoryArchivePath = Path.Combine(root, "archive.json");
        A.Options.LoopReplayCorpusPath = Path.Combine(root, "corpus.json");
        A.Options.RecoveryMemoryMaxQuarantinedLessons = 1;
        A.Options.RecoveryMemoryArchiveRetainedFiles = 3;
        A.Options.RecoveryMemoryArchiveMaxBytes = mode == "rotation" ? 1024 : 64 * 1024 * 1024;
        var paths = RecoveryStoragePaths.ResolveValidated(A.Options);
        Text(Path.Combine(root, writer + ".ready"), "");
        if (mode == "telemetry")
        {
            for (var i = 0; i < 50; i++)
                LoopTelemetryStore.Append(paths, JsonSerializer.Serialize(new { writer, i }));
            return 0;
        }
        return (int)RecoveryMemoryService.ExecuteRecoveryMemoryMaintenance("prune", null);
    }

    static void SharedWriters(string mode)
    {
        using var scope = A.BeginRun();
        using var f = new Files();
        f.Configure();
        foreach (var writer in new[] { "alpha", "beta" })
            Write(Path.Combine(f.Root, writer + ".json"), new RecoveryLessonStore
            {
                Version = 7,
                Lessons = Enumerable.Range(0, 3).Select(i => new RecoveryLesson
                {
                    Id = $"{writer}-{i}", Status = "quarantined", StrategySignature = $"{writer}-{i}",
                    UpdatedUtc = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddDays(i)
                }).ToList()
            });
        Write(f.Archive, new RecoveryLessonArchiveStore
        {
            Lessons = [new() { Id = "seed", WinningStrategy = new string('x', mode == "rotation" ? 4000 : 10) }]
        });
        var lockedPath = mode == "telemetry" ? f.Telemetry : f.Archive;
        var original = File.ReadAllBytes(lockedPath);
        Process? alpha = null, beta = null;
        try
        {
            using (StorageFileLock.Acquire(lockedPath))
            {
                alpha = Start("--storage-family-worker", f.Root, "alpha", mode);
                beta = Start("--storage-family-worker", f.Root, "beta", mode);
                var ready = SpinWait.SpinUntil(() => File.Exists(Path.Combine(f.Root, "alpha.ready")) &&
                    File.Exists(Path.Combine(f.Root, "beta.ready")), TimeSpan.FromSeconds(10));
                Check(ready, "workers did not start");
                Thread.Sleep(250);
                Check(!alpha.HasExited && !beta.HasExited, "writer bypassed the shared artifact lock");
                Check(File.ReadAllBytes(lockedPath).SequenceEqual(original), "writer modified a locked artifact");
            }
            Check(Complete(alpha!) == 0 && Complete(beta!) == 0, "shared storage worker failed");
            if (mode == "telemetry")
            {
                var lines = File.ReadAllLines(f.Telemetry).Where(line => !string.IsNullOrWhiteSpace(line)).ToArray();
                Check(lines.Length == 100 && lines.Distinct().Count() == 100, "concurrent telemetry lost records");
                foreach (var line in lines) using (JsonDocument.Parse(line)) { }
            }
            else
            {
                var files = mode == "rotation"
                    ? new[] { f.Archive, f.Archive + ".1", f.Archive + ".2", f.Archive + ".3" }
                    : new[] { f.Archive };
                var ids = files.Where(File.Exists).SelectMany(path =>
                    JsonSerializer.Deserialize<RecoveryLessonArchiveStore>(File.ReadAllText(path), A.PrettyJson)!.Lessons)
                    .Select(lesson => lesson.Id).OrderBy(id => id).ToArray();
                Check(ids.SequenceEqual(new[] { "alpha-0", "alpha-1", "beta-0", "beta-1", "seed" }),
                    "concurrent archive merge or rotation lost lessons");
            }
        }
        finally
        {
            foreach (var process in new[] { alpha, beta }.OfType<Process>())
            {
                if (!process.HasExited) { process.Kill(); process.WaitForExit(); }
                process.Dispose();
            }
        }
    }
}
