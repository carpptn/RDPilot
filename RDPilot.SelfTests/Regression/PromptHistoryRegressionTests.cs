using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;

namespace RDPilot.SelfTests.Regression;

internal static class PromptHistoryRegressionTests
{
    internal static (string Name, Action Run)[] Cases => InvalidDocuments.Select((json, index) =>
        ($"history preserves invalid document {index + 1} during save and migration", (Action)(() => InvalidDocument(json))))
        .Concat(new (string, Action)[]
        {
            ("history stale sessions preserve other writers and recency", InterleavedSessions),
            ("history stale sessions do not resurrect evicted entries", Retention),
            ("history retries failed prompts without resaving stale snapshots", PendingRetry),
            ("history retries after an invalid document is repaired", RepairedDocument),
            ("history preserves unreadable Windows file", UnreadableFile),
            ("history preserves a directory destination", DirectoryDestination),
            ("history migration preserves an existing destination", ExistingDestination),
            ("history migration preserves source when write fails", MigrationWriteFailure),
            ("history migration rejects identical paths", SamePathMigration),
            ("history normalizes supported documents and ignores empty input", SupportedDocument),
            ("history writers serialize across processes", SharedWriters),
            ("history migration locks the legacy source", () => LockedMigration(true)),
            ("history migration locks the destination", () => LockedMigration(false))
        }).ToArray();

    static readonly string[] InvalidDocuments =
    [
        "null", "[]", "{}", "{broken", "{\"version\":999,\"prompts\":[\"old\"],\"futureData\":\"keep\"}",
        "{\"prompts\":[\"old\"]}", "{\"version\":1,\"prompts\":null}",
        "{\"version\":1,\"prompts\":[null]}", "{\"version\":1,\"prompts\":[42]}",
        "{\"version\":999,\"Version\":1,\"prompts\":[]}",
        "{\"version\":1,\"prompts\":[],\"Prompts\":[\"keep\"]}",
        "{\"version\":\"1\",\"prompts\":[]}", "{\"version\":1,\"lessons\":[]}",
        "{\"version\":1,\"prompts\":[],\"cases\":[]}",
        "{\"version\":1,\"prompts\":[\"recoverable\"]"
    ];

    static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
    }

    sealed class Fixture : IDisposable
    {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string History => Path.Combine(Root, "history.json");
        internal string Legacy => Path.Combine(Root, "legacy", "history.json");
        internal Fixture() => Directory.CreateDirectory(Root);
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }

    static void Text(string path, string text)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, text.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
    }
    static void Seed(string path, params string[] prompts) =>
        Check(PromptHistoryService.Save(prompts, path), "could not seed history");
    static void Equal(string path, params string[] expected) =>
        Check(PromptHistoryService.Load(path).SequenceEqual(expected), "history content or recency was incorrect");

    static void InvalidDocument(string json)
    {
        using var f = new Fixture();
        Text(f.History, json);
        var original = File.ReadAllBytes(f.History);
        var entries = PromptHistoryService.Load(f.History);
        Check(entries.Count == 0, "unsupported history was exposed as supported");
        Check(!PromptHistoryService.Save(["new"], f.History), "save accepted invalid history");
        PromptHistoryService.Remember(entries, "new", f.History);
        Check(entries.SequenceEqual(["new"]), "failed persistence made console history unusable");
        Check(File.ReadAllBytes(f.History).SequenceEqual(original), "invalid history was overwritten");
        Text(f.Legacy, json);
        var destination = Path.Combine(f.Root, "migrated.json");
        Check(!PromptHistoryService.TryMigrateLegacyHistory(f.Legacy, destination) && !File.Exists(destination),
            "invalid legacy document was migrated");
        Check(File.ReadAllBytes(f.Legacy).SequenceEqual(original), "invalid legacy source was changed");
        File.Delete(f.Legacy);
        Seed(f.Legacy, "legacy");
        Check(!PromptHistoryService.TryMigrateLegacyHistory(f.Legacy, f.History), "migration overwrote invalid destination");
        Check(File.Exists(f.Legacy) && File.ReadAllBytes(f.History).SequenceEqual(original), "migration lost source or destination");
    }

    static void InterleavedSessions()
    {
        using var f = new Fixture();
        Seed(f.History, "seed");
        var a = PromptHistoryService.Load(f.History);
        var b = PromptHistoryService.Load(f.History);
        PromptHistoryService.Remember(a, "alpha", f.History);
        PromptHistoryService.Remember(b, "beta", f.History);
        Equal(f.History, "seed", "alpha", "beta");
        Check(b.SequenceEqual(["seed", "alpha", "beta"]), "local navigation did not refresh after merge");
        PromptHistoryService.Remember(a, "seed", f.History);
        Equal(f.History, "alpha", "beta", "seed");
    }

    static void Retention()
    {
        using var f = new Fixture();
        Seed(f.History, Enumerable.Range(0, 500).Select(i => $"old-{i}").ToArray());
        var stale = PromptHistoryService.Load(f.History);
        Seed(f.History, Enumerable.Range(0, 500).Select(i => $"new-{i}").ToArray());
        PromptHistoryService.Remember(stale, "latest", f.History);
        var expected = Enumerable.Range(1, 499).Select(i => $"new-{i}").Append("latest").ToArray();
        Equal(f.History, expected);
    }

    sealed class FailingWrites(AtomicTextStore.IFileOperations inner) : AtomicTextStore.IFileOperations
    {
        public bool Exists(string path) => inner.Exists(path);
        public void Write(string path, string text) { inner.Write(path, "partial"); throw new IOException("injected history write failure"); }
        public void Replace(string source, string destination, string? backup) => inner.Replace(source, destination, backup);
        public void Move(string source, string destination) => inner.Move(source, destination);
        public void Delete(string path) => inner.Delete(path);
    }

    static void PendingRetry()
    {
        using var f = new Fixture();
        Seed(f.History, "seed");
        var entries = PromptHistoryService.Load(f.History);
        var original = File.ReadAllBytes(f.History);
        var files = AtomicTextStore.Files;
        try
        {
            AtomicTextStore.Files = new FailingWrites(files);
            PromptHistoryService.Remember(entries, "pending", f.History);
        }
        finally { AtomicTextStore.Files = files; }
        Check(File.ReadAllBytes(f.History).SequenceEqual(original), "failed write modified history");
        Check(Directory.GetFiles(f.Root).Length == 1, "failed write left partial output");
        Seed(f.History, "other-writer");
        PromptHistoryService.Remember(entries, "next", f.History);
        Equal(f.History, "seed", "other-writer", "pending", "next");
    }

    static void RepairedDocument()
    {
        using var f = new Fixture();
        Text(f.History, "{broken");
        var entries = PromptHistoryService.Load(f.History);
        PromptHistoryService.Remember(entries, "pending", f.History);
        Text(f.History, "{\"version\":1,\"prompts\":[\"recovered\"]}");
        PromptHistoryService.Remember(entries, "next", f.History);
        Equal(f.History, "recovered", "pending", "next");
    }

    static void UnreadableFile()
    {
        using var f = new Fixture();
        Seed(f.History, "seed");
        var before = File.ReadAllBytes(f.History);
        using (var held = new FileStream(f.History, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            var entries = PromptHistoryService.Load(f.History);
            PromptHistoryService.Remember(entries, "pending", f.History);
            Check(!PromptHistoryService.Save(["new"], f.History), "save replaced unreadable history");
        }
        Check(File.ReadAllBytes(f.History).SequenceEqual(before), "unreadable file was changed");
    }

    static void DirectoryDestination()
    {
        using var f = new Fixture();
        Directory.CreateDirectory(f.History);
        Seed(f.Legacy, "legacy");
        Check(!PromptHistoryService.Save(["new"], f.History), "directory accepted as a file");
        Check(!PromptHistoryService.TryMigrateLegacyHistory(f.Legacy, f.History), "migration accepted directory");
        Check(Directory.Exists(f.History) && File.Exists(f.Legacy), "directory collision removed existing data");
    }

    static void ExistingDestination()
    {
        using var f = new Fixture();
        Seed(f.History, "destination"); Seed(f.Legacy, "legacy");
        Check(!PromptHistoryService.TryMigrateLegacyHistory(f.Legacy, f.History), "existing destination was overwritten");
        Equal(f.History, "destination"); Equal(f.Legacy, "legacy");
    }

    static void MigrationWriteFailure()
    {
        using var f = new Fixture();
        Seed(f.Legacy, "legacy");
        var before = File.ReadAllBytes(f.Legacy);
        var files = AtomicTextStore.Files;
        try
        {
            AtomicTextStore.Files = new FailingWrites(files);
            Check(!PromptHistoryService.TryMigrateLegacyHistory(f.Legacy, f.History), "failed migration reported success");
        }
        finally { AtomicTextStore.Files = files; }
        Check(!File.Exists(f.History) && File.ReadAllBytes(f.Legacy).SequenceEqual(before), "failed migration lost source data");
        Check(Directory.GetFiles(f.Root, "*", SearchOption.AllDirectories).Length == 1, "failed migration left temporary files");
    }

    static void SamePathMigration()
    {
        using var f = new Fixture();
        Seed(f.History, "seed");
        Check(!PromptHistoryService.TryMigrateLegacyHistory(f.History, Path.Combine(f.Root, ".", "history.json")),
            "same path migration was accepted");
        Equal(f.History, "seed");
    }

    static void SupportedDocument()
    {
        using var f = new Fixture();
        Text(f.History, "{\"Version\":1,\"Prompts\":[\" first \",\"line\\nbreak\",\"first\"]}");
        Equal(f.History, "line break", "first");
        var entries = PromptHistoryService.Load(f.History);
        var before = File.ReadAllBytes(f.History);
        PromptHistoryService.Remember(entries, " \r\n ", f.History);
        Check(File.ReadAllBytes(f.History).SequenceEqual(before), "empty prompt caused a write");
        PromptHistoryService.Remember(entries, " next ", f.History);
        Equal(f.History, "line break", "first", "next");
        var bytes = File.ReadAllBytes(f.History);
        Check(bytes.Take(3).SequenceEqual(new byte[] { 239, 187, 191 }) &&
            !File.ReadAllText(f.History).Replace("\r\n", "").Contains('\n'), "history encoding or line endings changed");
    }

    static Process Start(params string[] args)
    {
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        start.ArgumentList.Add("--prompt-history-worker");
        foreach (var arg in args) start.ArgumentList.Add(arg);
        return Process.Start(start)!;
    }
    static void Ready(Func<bool> predicate) => Check(SpinWait.SpinUntil(predicate, TimeSpan.FromSeconds(10)), "history worker barrier timed out");
    static void Complete(Process process)
    {
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        Check(process.WaitForExit(20000), "history worker timed out");
        Task.WhenAll(output, error).GetAwaiter().GetResult();
        Check(process.ExitCode == 0, "history worker failed: " + output.Result + error.Result);
    }
    static void Stop(Process? process)
    {
        if (process is null) return;
        if (!process.HasExited) { process.Kill(); process.WaitForExit(); }
        process.Dispose();
    }

    internal static int Worker(string[] args)
    {
        var root = args[1];
        var mode = args[2];
        var path = Path.Combine(root, "history.json");
        if (mode == "migrate")
        {
            Text(Path.Combine(root, "migrate.ready"), "");
            return PromptHistoryService.TryMigrateLegacyHistory(Path.Combine(root, "legacy", "history.json"), path) ? 0 : 1;
        }
        var entries = PromptHistoryService.Load(path);
        Text(Path.Combine(root, mode + ".ready"), "");
        Ready(() => File.Exists(Path.Combine(root, "go")));
        Text(Path.Combine(root, mode + ".attempt"), "");
        for (var index = 0; index < 25; index++) PromptHistoryService.Remember(entries, $"{mode}-{index}", path);
        return 0;
    }

    static void SharedWriters()
    {
        using var f = new Fixture();
        Seed(f.History, "seed");
        Process? a = null, b = null;
        try
        {
            a = Start(f.Root, "alpha"); b = Start(f.Root, "beta");
            Ready(() => File.Exists(Path.Combine(f.Root, "alpha.ready")) && File.Exists(Path.Combine(f.Root, "beta.ready")));
            var before = File.ReadAllBytes(f.History);
            using (StorageFileLock.Acquire(f.History))
            {
                Text(Path.Combine(f.Root, "go"), "");
                Ready(() => File.Exists(Path.Combine(f.Root, "alpha.attempt")) && File.Exists(Path.Combine(f.Root, "beta.attempt")));
                Thread.Sleep(200);
                Check(!a.HasExited && !b.HasExited && File.ReadAllBytes(f.History).SequenceEqual(before), "writer bypassed history lock");
            }
            Complete(a); Complete(b);
            var expected = Enumerable.Range(0, 25).SelectMany(i => new[] { $"alpha-{i}", $"beta-{i}" }).Append("seed").ToHashSet();
            var actual = PromptHistoryService.Load(f.History);
            Check(actual.Count == 51 && actual.ToHashSet().SetEquals(expected), "concurrent writers lost history entries");
        }
        finally { Stop(a); Stop(b); }
    }

    static void LockedMigration(bool source)
    {
        using var f = new Fixture();
        Seed(f.Legacy, "legacy");
        Process? child = null;
        try
        {
            using (StorageFileLock.Acquire(source ? f.Legacy : f.History))
            {
                child = Start(f.Root, "migrate");
                Ready(() => File.Exists(Path.Combine(f.Root, "migrate.ready")));
                Thread.Sleep(200);
                Check(!child.HasExited && !File.Exists(f.History) && File.Exists(f.Legacy), "migration bypassed file lock");
            }
            Complete(child);
            Equal(f.History, "legacy");
            Check(!File.Exists(f.Legacy), "successful migration retained the legacy file");
        }
        finally { Stop(child); }
    }
}