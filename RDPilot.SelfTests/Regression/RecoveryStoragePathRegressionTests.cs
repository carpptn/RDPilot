using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class RecoveryStoragePathRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
        new[] { "primary", "backup", "case", "dot", "environment", "rotation-1", "rotation-2", "rotation-case" }
        .Select(kind => ($"memory CLI rejects {kind} path collision without writing", (Action)(() => InvalidCli(kind))))
        .Concat(new[] { "extended", "device", "nt", "extended-slashes", "device-slashes", "nt-slashes", "mixed", "environment" }
            .SelectMany(kind => new[] { "backup", "primary", "rotation" }
                .Select(target => ($"memory CLI rejects {kind} device path for {target} without writing",
                    (Action)(() => DeviceCli(kind, target))))))
        .Concat(new (string, Action)[]
        {
            ("device storage paths do not create missing files", () => DeviceCli("extended", "backup", seed: false)),
            ("device storage paths do not restore a backup", DeviceBackupOnly),
            ("memory load rejects device paths without CLI normalization", () => DirectDevicePath("load")),
            ("memory writer rejects device paths without a prior load", () => DirectDevicePath("save")),
            ("archive writer rejects device paths before rotation", () => DirectDevicePath("archive")),
            ("device UNC and volume paths are rejected without filesystem access", DeviceNamespaceResolution),
            ("ordinary UNC and long paths remain supported", OrdinaryPathResolution),
            ("invalid storage paths do not create missing files", MissingFiles),
            ("invalid storage paths do not restore a backup", BackupOnly),
            ("memory load rejects collisions without CLI normalization", DirectLoad),
            ("memory writer rejects collisions without a prior load", () => DirectWrite(false)),
            ("archive writer rejects collisions before rotation", () => DirectWrite(true)),
            ("valid memory prune preserves every displaced lesson", ValidPrune),
            ("final CLI override can correct a conflicting config file", ConfigOverride),
            ("archive rotation preserves memory and backup files", ValidRotation),
            ("storage paths preserve application-relative defaults and rotation bounds", PathResolution)
        }).ToArray();

    static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
    }

    sealed class Files : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string Primary { get; }
        internal string Archive { get; }
        internal string Config => Path.Combine(Root, "config.json");
        internal Files(string kind = "valid", bool seed = true)
        {
            Directory.CreateDirectory(Root);
            var archive = Path.Combine(Root, "archive.json");
            Primary = kind.StartsWith("rotation", StringComparison.Ordinal)
                ? archive + (kind == "rotation-2" ? ".2" : ".1") : Path.Combine(Root, "memory.json");
            Archive = kind switch
            {
                "primary" => Primary,
                "backup" => Primary + ".bak",
                "case" => (Primary + ".bak").ToUpperInvariant(),
                "dot" => Path.Combine(Root, ".", "unused", "..", "memory.json.bak"),
                "environment" => "%RDPILOT_TEST_STORAGE_ROOT%/memory.json.bak",
                "rotation-case" => archive.ToUpperInvariant(),
                _ => archive
            };
            if (!seed) return;
            Write(Primary, JsonSerializer.Serialize(new RecoveryLessonStore
            {
                Version = 7,
                Lessons = Enumerable.Range(0, 3).Select(index => new RecoveryLesson
                {
                    Id = $"lesson-{index}", Status = "quarantined", StrategySignature = $"strategy-{index}",
                    UpdatedUtc = DateTime.UtcNow.AddDays(index - 3)
                }).ToList()
            }, A.PrettyJson));
            Write(Primary + ".bak", "previous-backup-must-survive");
            foreach (var candidate in new[] { archive, archive + ".1", archive + ".2" })
                if (!File.Exists(candidate)) Write(candidate, JsonSerializer.Serialize(new RecoveryLessonArchiveStore(), A.PrettyJson));
        }
        internal static void Write(string path, string text) => File.WriteAllText(path,
            text.Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n", new UTF8Encoding(true));
        internal Dictionary<string, byte[]> Snapshot() => Directory.GetFiles(Root, "*", SearchOption.AllDirectories)
            .ToDictionary(path => path, File.ReadAllBytes, StringComparer.OrdinalIgnoreCase);
        internal void Unchanged(Dictionary<string, byte[]> before)
        {
            var after = Snapshot();
            Check(before.Count == after.Count && before.All(item => after.TryGetValue(item.Key, out var bytes) && bytes.SequenceEqual(item.Value)),
                "rejected configuration changed, deleted, or created a file");
        }
        public void Dispose() => Directory.Delete(Root, true);
    }

    static (int Exit, string Output, string Error) Cli(Files files, bool configOverride = false, string? primary = null, string? archive = null)
    {
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        start.Environment["RDPILOT_TEST_STORAGE_ROOT"] = files.Root;
        start.Environment["RDPILOT_TEST_DEVICE_ROOT"] = @"\\?\" + files.Root;
        if (configOverride) start.Environment["RDPILOT_CONFIG"] = files.Config;
        foreach (var arg in new[] { "--command-main-probe", "--recovery-memory", "--recovery-memory-path", primary ?? files.Primary,
            "--recovery-archive-path", archive ?? files.Archive, "--recovery-archive-retained-files", "2",
            "--recovery-max-quarantined", "1", "--memory-prune" }) start.ArgumentList.Add(arg);
        using var child = Process.Start(start)!;
        try
        {
            var output = child.StandardOutput.ReadToEndAsync();
            var error = child.StandardError.ReadToEndAsync();
            Check(child.WaitForExit(10000), "memory CLI did not exit");
            Task.WhenAll(output, error).GetAwaiter().GetResult();
            return (child.ExitCode, output.Result, error.Result);
        }
        finally { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } }
    }

    static string DevicePath(string path, string kind) => kind switch
    {
        "extended" => @"\\?\" + path,
        "device" => @"\\.\" + path,
        "nt" => @"\??\" + path,
        "extended-slashes" => "//?/" + path.Replace('\\', '/'),
        "device-slashes" => "//./" + path.Replace('\\', '/'),
        "nt-slashes" => "/??/" + path.Replace('\\', '/'),
        "mixed" => @"\/?\" + path,
        "environment" => @"%RDPILOT_TEST_DEVICE_ROOT%\" + Path.GetFileName(path),
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };

    static void DeviceCli(string kind, string target, bool seed = true)
    {
        using var files = new Files(target == "rotation" ? "rotation-1" : "backup", seed);
        if (seed) File.Copy(files.Primary, files.Primary + ".bak", overwrite: true);
        var before = files.Snapshot();
        var result = target == "primary"
            ? Cli(files, primary: DevicePath(files.Primary, kind))
            : Cli(files, archive: DevicePath(files.Archive, kind));
        Check(result.Exit == 2 && result.Error.Contains("device paths", StringComparison.OrdinalIgnoreCase),
            $"unexpected device path CLI result: {result}");
        files.Unchanged(before);
    }

    static void DeviceBackupOnly()
    {
        using var files = new Files("backup");
        File.Copy(files.Primary, files.Primary + ".bak", overwrite: true);
        File.Delete(files.Primary);
        var before = files.Snapshot();
        Check(Cli(files, archive: DevicePath(files.Archive, "extended")).Exit == 2, "device path restored a backup");
        files.Unchanged(before);
    }

    static void DirectDevicePath(string operation)
    {
        using var scope = A.BeginRun();
        using var files = new Files(operation == "archive" ? "rotation-1" : "backup");
        Configure(files);
        A.Options.RecoveryMemoryArchivePath = DevicePath(files.Archive, "extended");
        if (operation == "archive")
            Files.Write(files.Archive, JsonSerializer.Serialize(new RecoveryLessonArchiveStore
            {
                Lessons = [new() { Id = "old", WinningStrategy = new string('x', 2000) }]
            }, A.PrettyJson));
        var before = files.Snapshot();
        if (operation == "load")
        {
            var rejected = false;
            try { _ = RecoveryMemoryService.LoadRecoveryLessons(); }
            catch (ArgumentException ex) { rejected = ex.Message.Contains("device paths", StringComparison.OrdinalIgnoreCase); }
            Check(rejected, "load accepted a device path");
        }
        else
        {
            A.Options.RecoveryMemoryArchiveMaxBytes = 1024; // The seeded archive exceeds this limit.
            Check(!WriteWithoutLoad(files, operation == "archive", new[] { new RecoveryLesson { Id = "new" } }),
                "writer accepted a device path");
        }
        files.Unchanged(before);
    }

    static void DeviceNamespaceResolution()
    {
        foreach (var path in new[] { @"\\?\UNC\server\share\memory.json", @"\\.\UNC\server\share\memory.json",
            @"\??\UNC\server\share\memory.json", @"\\?\Volume{11111111-1111-1111-1111-111111111111}\memory.json",
            @"\\.\GLOBALROOT\Device\HarddiskVolume1\memory.json" })
        {
            var rejected = false;
            try { _ = RecoveryStoragePaths.ResolvePath(path); }
            catch (ArgumentException ex) { rejected = ex.Message.Contains("device paths", StringComparison.OrdinalIgnoreCase); }
            Check(rejected, $"device namespace accepted: {path}");
        }
    }

    static void OrdinaryPathResolution()
    {
        const string unc = @"\\server\share\memory.json";
        Check(RecoveryStoragePaths.ResolvePath(unc) == unc, "ordinary UNC path changed");
        var longPath = Path.Combine(Path.GetTempPath(), string.Join("\\", Enumerable.Repeat(new string('x', 80), 5)), "memory.json");
        Check(RecoveryStoragePaths.ResolvePath(longPath) == longPath, "ordinary long path changed");
    }

    static void InvalidCli(string kind)
    {
        using var files = new Files(kind);
        var before = files.Snapshot();
        var result = Cli(files);
        Check(result.Exit == 2 && result.Error.Contains("path collision"), $"unexpected CLI result: {result}");
        files.Unchanged(before);
    }

    static void MissingFiles()
    {
        using var files = new Files("backup", seed: false);
        var before = files.Snapshot();
        Check(Cli(files).Exit == 2, "missing-file collision accepted");
        files.Unchanged(before);
    }

    static void BackupOnly()
    {
        using var files = new Files("backup");
        File.Copy(files.Primary, files.Primary + ".bak", overwrite: true);
        File.Delete(files.Primary);
        var before = files.Snapshot();
        Check(Cli(files).Exit == 2, "backup-only collision accepted");
        files.Unchanged(before);
    }

    static void Configure(Files files)
    {
        A.Options.RecoveryMemoryEnabled = true;
        A.Options.RecoveryMemoryPath = files.Primary;
        A.Options.RecoveryMemoryArchivePath = files.Archive;
        A.Options.RecoveryMemoryArchiveRetainedFiles = 2;
    }

    static void DirectLoad()
    {
        using var scope = A.BeginRun();
        using var files = new Files("backup");
        Configure(files);
        var before = files.Snapshot();
        var rejected = false;
        try { _ = RecoveryMemoryService.LoadRecoveryLessons(); }
        catch (InvalidOperationException ex) { rejected = ex.Message.Contains("path collision"); }
        Check(rejected, "load bypassed collision validation");
        files.Unchanged(before);
    }

    static bool WriteWithoutLoad(Files files, bool archive, IReadOnlyCollection<RecoveryLesson> lessons) =>
        (bool)typeof(RecoveryMemoryService).GetMethod(archive ? "ArchiveRecoveryLessonsWithoutLock" : "SaveRecoveryLessonsWithoutLock",
            BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, new object[] { files.Primary, lessons })!;

    static void DirectWrite(bool archive)
    {
        using var scope = A.BeginRun();
        using var files = new Files(archive ? "rotation-1" : "backup");
        Configure(files);
        var before = files.Snapshot();
        Check(!WriteWithoutLoad(files, archive, new[] { new RecoveryLesson { Id = "new-lesson" } }), "writer accepted a collision");
        files.Unchanged(before);
    }

    static string[] Ids(string path) => JsonSerializer.Deserialize<RecoveryLessonArchiveStore>(File.ReadAllText(path), A.PrettyJson)!
        .Lessons.Select(lesson => lesson.Id).OrderBy(id => id).ToArray();

    static void AssertRetention(Files files, (int Exit, string Output, string Error) result)
    {
        Check(result.Exit == 0, $"valid prune failed: {result}");
        Check(Ids(files.Primary).SequenceEqual(new[] { "lesson-2" }), "unexpected retained lesson");
        Check(Ids(files.Archive).SequenceEqual(new[] { "lesson-0", "lesson-1" }), "prune lost an archived lesson");
    }

    static void ValidPrune()
    {
        using var files = new Files();
        AssertRetention(files, Cli(files));
    }

    static void ConfigOverride()
    {
        using var files = new Files();
        Files.Write(files.Config, JsonSerializer.Serialize(new { recoveryMemoryArchivePath = files.Primary + ".bak" }));
        AssertRetention(files, Cli(files, configOverride: true));
    }

    static void ValidRotation()
    {
        using var scope = A.BeginRun();
        using var files = new Files();
        Configure(files);
        A.Options.RecoveryMemoryArchiveMaxBytes = 1024;
        var primary = File.ReadAllBytes(files.Primary);
        var backup = File.ReadAllBytes(files.Primary + ".bak");
        void Seed(string path, string id, string text = "") => Files.Write(path,
            JsonSerializer.Serialize(new RecoveryLessonArchiveStore { Lessons = [new() { Id = id, WinningStrategy = text }] }, A.PrettyJson));
        Seed(files.Archive, "old", new string('x', 2000));
        Seed(files.Archive + ".1", "older");
        Seed(files.Archive + ".2", "oldest");
        Check(WriteWithoutLoad(files, archive: true, new[] { new RecoveryLesson { Id = "new" } }), "valid rotation failed");
        Check(Ids(files.Archive).SequenceEqual(new[] { "new" }) &&
            Ids(files.Archive + ".1").SequenceEqual(new[] { "old" }) &&
            Ids(files.Archive + ".2").SequenceEqual(new[] { "older" }), "rotation moved the wrong generations");
        Check(primary.SequenceEqual(File.ReadAllBytes(files.Primary)) && backup.SequenceEqual(File.ReadAllBytes(files.Primary + ".bak")),
            "rotation touched primary or backup data");
    }

    static void PathResolution()
    {
        var options = new AppOptions { RecoveryMemoryPath = "memory/custom.json", RecoveryMemoryArchivePath = null,
            RecoveryMemoryArchiveRetainedFiles = 2 };
        var paths = RecoveryStoragePaths.ResolveValidated(options);
        Check(paths.Primary == Path.Combine(AppContext.BaseDirectory, "memory", "custom.json") &&
            paths.Backup == paths.Primary + ".bak" && paths.Archive == Path.Combine(AppContext.BaseDirectory, "memory", "custom-archive.json"),
            "default path resolution changed");
        options.RecoveryMemoryPath = "memory/archive.json.3";
        options.RecoveryMemoryArchivePath = "memory/archive.json";
        _ = RecoveryStoragePaths.ResolveValidated(options); // Generation 3 is unused with retention 2.
        options.RecoveryMemoryArchiveRetainedFiles = 3;
        var rejected = false;
        try { _ = RecoveryStoragePaths.ResolveValidated(options); }
        catch (InvalidOperationException) { rejected = true; }
        Check(rejected, "increased retention did not validate the newly used generation");
    }
}
