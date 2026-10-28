using System.IO;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class ArchiveRotationRegressionTests
{
    internal static (string Name, Action Run)[] Cases => InvalidArchives.Select((json, i) =>
        ($"archive prune preserves invalid document {i + 1}", (Action)(() => InvalidArchive(json))))
        .Concat(Enumerable.Range(1, 4).Select(step =>
            ($"rotation restores every file after move failure {step}", (Action)(() => MoveFailure(step)))))
        .Concat(new[] { false, true }.SelectMany(telemetry => Enumerable.Range(0, 3).Select(generation =>
            ($"{(telemetry ? "telemetry" : "archive")} preserves files with directory at generation {generation}",
                (Action)(() => DirectoryCollision(telemetry, generation))))))
        .Concat(new (string, Action)[]
        {
            ("archive missing file and supported version merge correctly", ValidArchive),
            ("archive successful rotation retains generations", () => SuccessfulRotation(false)),
            ("telemetry successful rotation retains generations and new record", () => SuccessfulRotation(true)),
            ("archive rotation restores files when Windows denies moving active file", () => LockedActive(false)),
            ("telemetry rotation restores files when Windows denies moving active file", () => LockedActive(true)),
            ("rotation staging failure preserves all files", StageFailure),
            ("nonrotating replacement failure preserves active file", ReplacementFailure),
            ("rotation handles missing intermediate generations", SparseRotation),
            ("rotation rollback failure preserves recoverable originals", RollbackFailure),
            ("rotation cleanup failure leaves committed data and backup", CleanupFailure)
        }).ToArray();

    static readonly string[] InvalidArchives =
    [
        "null", "[]", "{}", "{\"cases\":[{\"name\":\"keep\"}]}",
        "{\"version\":999,\"lessons\":[{\"id\":\"old\",\"futureData\":\"keep\"}]}",
        "{\"version\":7,\"lessons\":[],\"calibration\":{}}",
        "{\"lessons\":[]}", "{\"version\":1,\"lessons\":null}",
        "{\"version\":1,\"lessons\":[null]}", "{\"version\":1,\"lessons\":[42]}",
        "{\"version\":1,\"lessons\":[],\"Lessons\":[]}",
        "{\"version\":999,\"Version\":1,\"lessons\":[]}",
        "{\"version\":\"1\",\"lessons\":[]}", "{\"version\":1,\"lessons\":[],\"prompts\":[]}",
        "{\"version\":1,\"lessons\":[{\"updatedUtc\":\"bad-date\"}]}", "{broken"
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class Fixture : IDisposable
    {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string Memory => Path.Combine(Root, "memory.json");
        internal string Archive => Path.Combine(Root, "archive.json");
        internal string Telemetry => Path.Combine(Root, "loop-telemetry.jsonl");
        internal Fixture()
        {
            Directory.CreateDirectory(Root);
            A.Options.RecoveryMemoryEnabled = true;
            A.Options.RecoveryMemoryPath = Memory;
            A.Options.RecoveryMemoryArchivePath = Archive;
            A.Options.LoopReplayCorpusPath = Path.Combine(Root, "corpus.json");
            A.Options.RecoveryMemoryArchiveRetainedFiles = 2;
            A.Options.RecoveryTelemetryRetainedFiles = 2;
            A.Options.RecoveryMemoryArchiveMaxBytes = 1024;
            A.Options.RecoveryTelemetryMaxBytes = 65536;
            A.Options.RecoveryMemoryMaxQuarantinedLessons = 1;
        }
        internal void Seed(string path)
        {
            Text(path, "active"); Text(path + ".1", "previous"); Text(path + ".2", "oldest");
        }
        internal Dictionary<string, byte[]> Snapshot() => Directory.GetFiles(Root)
            .ToDictionary(path => path, File.ReadAllBytes);
        internal void Unchanged(Dictionary<string, byte[]> before)
        {
            var after = Snapshot();
            Check(before.Count == after.Count && before.All(p => after.TryGetValue(p.Key, out var bytes) && bytes.SequenceEqual(p.Value)),
                "failed operation changed retained bytes or left temporary files");
        }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }

    static void Text(string path, string value) => File.WriteAllText(path,
        value.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
    static void Json(string path, object value) => Text(path, JsonSerializer.Serialize(value, A.PrettyJson));
    static string[] Generations(string path) => [path + ".1", path + ".2"];
    static byte[] NewContents => Encoding.UTF8.GetBytes("replacement");
    static void ExpectFailure(Action action)
    {
        try { action(); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return; }
        throw new InvalidOperationException("expected I/O failure");
    }

    static void InvalidArchive(string json)
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        Json(f.Memory, new RecoveryLessonStore { Lessons = Enumerable.Range(0, 3).Select(i => new RecoveryLesson
        {
            Id = $"lesson-{i}", Status = "quarantined", StrategySignature = $"unique-{i}",
            UpdatedUtc = DateTime.UtcNow.AddDays(i - 3)
        }).ToList() });
        File.Copy(f.Memory, f.Memory + ".bak");
        Text(f.Archive, json); Text(f.Archive + ".1", "older"); Text(f.Archive + ".2", "oldest");
        var before = f.Snapshot();
        var result = RecoveryMemoryService.ExecuteRecoveryMemoryMaintenance("prune", null);
        Check(result == CommandOutcome.Failed, "prune reported success for an invalid archive");
        f.Unchanged(before);
    }

    static void ValidArchive()
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        A.Options.RecoveryMemoryArchiveMaxBytes = int.MaxValue;
        Check(RecoveryArchiveStore.Append(f.Memory, [new() { Id = "first" }]), "missing archive was not created");
        Check(RecoveryArchiveStore.Append(f.Memory, [new() { Id = "second" }]), "supported archive was rejected");
        Check(RecoveryArchiveReader.ReadExisting(f.Archive).Lessons.Select(l => l.Id).ToHashSet().SetEquals(["first", "second"]),
            "merge lost an existing lesson");
    }

    static void SeedLarge(Fixture f, bool telemetry)
    {
        if (telemetry)
            Text(f.Telemetry, new string('x', 66000));
        else
            Json(f.Archive, new RecoveryLessonArchiveStore { Lessons = [new() { Id = "active", WinningStrategy = new string('x', 4000) }] });
        var path = telemetry ? f.Telemetry : f.Archive;
        Text(path + ".1", "previous"); Text(path + ".2", "oldest");
    }

    static void DirectoryCollision(bool telemetry, int generation)
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        SeedLarge(f, telemetry);
        var path = telemetry ? f.Telemetry : f.Archive;
        var blocked = generation == 0 ? path : path + $".{generation}";
        File.Delete(blocked); Directory.CreateDirectory(blocked);
        var before = f.Snapshot();
        if (telemetry)
            ExpectFailure(() => LoopTelemetryStore.Append(RecoveryStoragePaths.ResolveValidated(A.Options), "new"));
        else
            Check(!RecoveryArchiveStore.Append(f.Memory, [new() { Id = "new" }]), "blocked archive succeeded");
        Check(Directory.Exists(blocked), "blocking directory was changed");
        f.Unchanged(before);
    }

    static void LockedActive(bool telemetry)
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        SeedLarge(f, telemetry);
        var path = telemetry ? f.Telemetry : f.Archive;
        var before = f.Snapshot();
        using (var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            if (telemetry)
                ExpectFailure(() => LoopTelemetryStore.Append(RecoveryStoragePaths.ResolveValidated(A.Options), "new"));
            else
                Check(!RecoveryArchiveStore.Append(f.Memory, [new() { Id = "new" }]), "locked archive rotation succeeded");
        }
        f.Unchanged(before);
    }
    static void SuccessfulRotation(bool telemetry)
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        SeedLarge(f, telemetry);
        var path = telemetry ? f.Telemetry : f.Archive;
        var active = File.ReadAllBytes(path);
        var previous = File.ReadAllBytes(path + ".1");
        if (telemetry)
        {
            LoopTelemetryStore.Append(RecoveryStoragePaths.ResolveValidated(A.Options), "new");
            Check(File.ReadAllText(path) == "new" + Environment.NewLine, "new telemetry record was not committed");
        }
        else
        {
            Check(RecoveryArchiveStore.Append(f.Memory, [new() { Id = "new" }]), "archive rotation failed");
            Check(RecoveryArchiveReader.ReadExisting(path).Lessons.Single().Id == "new", "new archive contents are wrong");
        }
        Check(File.ReadAllBytes(path + ".1").SequenceEqual(active) && File.ReadAllBytes(path + ".2").SequenceEqual(previous),
            "rotation did not preserve retained generations");
        Check(Directory.GetFiles(f.Root).Length == 3, "successful rotation left staging files");
    }

    sealed class FaultFiles : RotatingFileStore.IFileOperations
    {
        internal HashSet<int> FailMoves { get; } = [];
        internal bool FailWrite, FailReplace, FailDelete;
        int moves;
        public bool Exists(string path) => File.Exists(path);
        public bool IsDirectory(string path) => Directory.Exists(path);
        public void Write(string path, byte[] contents)
        {
            File.WriteAllBytes(path, FailWrite ? contents.Take(1).ToArray() : contents);
            if (FailWrite) throw new IOException("injected partial write failure");
        }
        public void Move(string source, string destination)
        {
            if (FailMoves.Contains(++moves)) throw new IOException("injected move failure");
            File.Move(source, destination);
        }
        public void Replace(string source, string destination)
        {
            if (FailReplace) throw new IOException("injected replacement failure");
            File.Move(source, destination, overwrite: true);
        }
        public void Delete(string path)
        {
            if (FailDelete) throw new IOException("injected cleanup failure");
            File.Delete(path);
        }
    }

    static void MoveFailure(int step) => PreservesOnFailure(new FaultFiles { FailMoves = { step } });
    static void StageFailure() => PreservesOnFailure(new FaultFiles { FailWrite = true });
    static void ReplacementFailure() => PreservesOnFailure(new FaultFiles { FailReplace = true }, rotate: false);
    static void PreservesOnFailure(FaultFiles files, bool rotate = true)
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        f.Seed(f.Archive);
        var before = f.Snapshot();
        ExpectFailure(() => RotatingFileStore.Write(f.Archive, NewContents, rotate ? Generations(f.Archive) : [], files));
        f.Unchanged(before);
    }

    static void SparseRotation()
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        f.Seed(f.Archive); File.Delete(f.Archive + ".1");
        RotatingFileStore.Write(f.Archive, NewContents, Generations(f.Archive));
        Check(File.ReadAllText(f.Archive) == "replacement" && File.ReadAllText(f.Archive + ".1") == "active" &&
            !File.Exists(f.Archive + ".2") && Directory.GetFiles(f.Root).Length == 2, "sparse rotation retained incorrect generations");
    }

    static void RollbackFailure()
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        f.Seed(f.Archive);
        try
        {
            RotatingFileStore.Write(f.Archive, NewContents, Generations(f.Archive), new FaultFiles { FailMoves = { 3, 4 } });
            throw new InvalidOperationException("expected incomplete rollback");
        }
        catch (AggregateException ex)
        {
            Check(ex.Message.Contains(f.Archive) && ex.Message.Contains(".rollback"), "recovery paths were not reported");
        }
        var values = Directory.GetFiles(f.Root).Select(File.ReadAllText).ToHashSet();
        Check(values.SetEquals(["active", "previous", "oldest"]), "rollback failure deleted recoverable original data");
    }

    static void CleanupFailure()
    {
        using var scope = A.BeginRun();
        using var f = new Fixture();
        f.Seed(f.Archive);
        RotatingFileStore.Write(f.Archive, NewContents, Generations(f.Archive), new FaultFiles { FailDelete = true });
        Check(File.ReadAllText(f.Archive) == "replacement" && File.ReadAllText(f.Archive + ".1") == "active" &&
            File.ReadAllText(f.Archive + ".2") == "previous", "cleanup failure lost committed data");
        Check(File.ReadAllText(Directory.GetFiles(f.Root, "*.rollback").Single()) == "oldest", "cleanup failure discarded backup");
    }
}