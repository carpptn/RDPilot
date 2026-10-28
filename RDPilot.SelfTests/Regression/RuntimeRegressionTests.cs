using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;

namespace RDPilot.SelfTests.Regression;

internal static class RuntimeRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("run scope isolates settings and rejects overlapping runs", RunScope),
        ("hotkey monitoring stops before its source is disposed", () => MonitorLifetime().GetAwaiter().GetResult()),
        ("hotkey cancellation belongs to its captured run", () => MonitorCancellation().GetAwaiter().GetResult()),
        ("atomic replacement failure preserves primary and backup", AtomicFailure),
        ("request logs hydrate to the exact original request", RequestRoundTrip),
        ("console writers are restored when a run fails", ConsoleLifetime),
        ("single key partial failure releases accepted down event", SingleKeyFailure),
        ("failed emergency releases are retried by run cleanup", RetryRelease),
        ("cancellation blocks wheel input but permits releases", CancelledInput)
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    static void RunScope()
    {
        var previous = A.CurrentRun;
        var options = A.Options;
        var originalModel = options.Model;
        var scope = A.BeginRun();
        try
        {
            Check(!ReferenceEquals(previous, A.CurrentRun), "run context was reused");
            Check(!ReferenceEquals(options, A.Options), "options were not snapshotted");
            A.Options.Model = "test-only";
            A.CurrentRun.CancelRequested = true;
            var rejected = Task.Run(() =>
            {
                try { using var unexpected = A.BeginRun(); return false; }
                catch (InvalidOperationException) { return true; }
            }).GetAwaiter().GetResult();
            Check(rejected, "two runs can race on console and desktop state");
            Check(options.Model == originalModel, "run modified configured options");
        }
        finally { scope.Dispose(); }
        Check(ReferenceEquals(previous, A.CurrentRun), "previous context was not restored");
        using var next = A.BeginRun();
        var nextContext = A.CurrentRun;
        scope.Dispose();
        Check(ReferenceEquals(nextContext, A.CurrentRun), "double dispose corrupted a later run");
        Check(!A.CurrentRun.CancelRequested, "transient state leaked to the next run");
    }

    static async Task MonitorLifetime()
    {
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var polls = 0;
        var monitor = new CancellationMonitor(new RunContext(), () =>
        {
            Interlocked.Increment(ref polls);
            entered.TrySetResult();
            return false;
        });
        await entered.Task.WaitAsync(TimeSpan.FromSeconds(3));
        await monitor.DisposeAsync();
        var afterDispose = polls;
        await Task.Delay(80);
        Check(afterDispose == polls && monitor.Token.IsCancellationRequested, "monitor is still polling after disposal");
    }

    static async Task MonitorCancellation()
    {
        var owner = new RunContext();
        await using var monitor = new CancellationMonitor(owner, () => true);
        try { await Task.Delay(Timeout.Infinite, monitor.Token).WaitAsync(TimeSpan.FromSeconds(3)); }
        catch (OperationCanceledException) { }
        Check(owner.CancelRequested, "abort did not reach the owner");
        Check(!A.CurrentRun.CancelRequested, "abort affected a different run");
    }

    sealed class FailingReplace : AtomicTextStore.IFileOperations
    {
        public bool Exists(string path) => File.Exists(path);
        public void Write(string path, string text) => File.WriteAllText(path, text);
        public void Replace(string source, string destination, string? backup) => throw new IOException("test replacement failure");
        public void Move(string source, string destination) => throw new InvalidOperationException("unsafe fallback move");
        public void Delete(string path) => File.Delete(path);
    }

    static void AtomicFailure()
    {
        var directory = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "state.json");
        var original = AtomicTextStore.Files;
        try
        {
            File.WriteAllText(path, "primary");
            File.WriteAllText(path + ".bak", "backup");
            AtomicTextStore.Files = new FailingReplace();
            var failed = false;
            try { AtomicTextStore.Write(path, "replacement", path + ".bak"); }
            catch (IOException) { failed = true; }
            Check(failed, "I/O failure was hidden");
            Check(File.ReadAllText(path) == "primary" && File.ReadAllText(path + ".bak") == "backup", "replacement destroyed recoverable data");
            Check(Directory.GetFiles(directory, "*.tmp").Length == 0, "failed write leaked its temporary file");
            AtomicTextStore.Files = original;
            AtomicTextStore.Write(path, "new\ncontent", path + ".bak");
            Check(File.ReadAllText(path) == "new\r\ncontent" && File.ReadAllText(path + ".bak") == "primary", "successful atomic rotation is invalid");
            Check(File.ReadAllBytes(path).Take(3).SequenceEqual(new byte[] { 239, 187, 191 }), "persisted text has no UTF-8 BOM");
        }
        finally
        {
            AtomicTextStore.Files = original;
            Directory.Delete(directory, true);
        }
    }

    static void RequestRoundTrip()
    {
        using var scope = A.BeginRun();
        A.Options.LogRequests = true;
        A.Options.LogScreens = true;
        var id = "test_" + Guid.NewGuid().ToString("N");
        var path = Path.Combine(Path.GetTempPath(), id + ".json");
        using var bitmap = new System.Drawing.Bitmap(32, 24);
        var data = ScreenshotService.EncodeBitmapToDataUrl(bitmap, 0, "png", 80);
        var request = PromptAndRequestFactory.BuildQARequestBody("test-model", "rules", "question", data, 32, 24, 0, 0, 0, 0, null, data, new("window", "app", null, null, null, null));
        var expected = JsonSerializer.SerializeToNode(request)!;
        try
        {
            ArtifactStorageService.SaveRequestLog(path, request);
            var text = File.ReadAllText(path);
            Check(!text.Contains("base64,"), "request log contains full images");
            var logged = JsonNode.Parse(text)!;
            var stats = ArtifactAnalysisService.HydrateReplayImages(logged);
            Check(stats.Images > 0 && JsonNode.DeepEquals(expected, logged), "replay changed request data, schema, or image bytes");
            Check(JsonNode.DeepEquals(expected, JsonSerializer.SerializeToNode(request)), "logging mutated the outgoing request");
            A.Options.LogScreens = false;
            ArtifactStorageService.SaveRequestLog(path, request);
            Check(!File.ReadAllText(path).Contains("base64,"), "disabled screenshots leaked pixels");
        }
        finally
        {
            File.Delete(path);
            var imageDirectory = Path.Combine(AppContext.BaseDirectory, "screens");
            if (Directory.Exists(imageDirectory))
                foreach (var image in Directory.GetFiles(imageDirectory, id + "_input_*")) File.Delete(image);
        }
    }

    static void ConsoleLifetime()
    {
        var output = Console.Out;
        var error = Console.Error;
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".log");
        try
        {
            try
            {
                using var log = new RunLogSession(path);
                Console.WriteLine("test log line");
                throw new IOException("test failure");
            }
            catch (IOException) { }
            Check(ReferenceEquals(output, Console.Out) && ReferenceEquals(error, Console.Error), "console still references a disposed log");
            Check(File.ReadAllText(path).Contains("test log line"), "run log is empty");
        }
        finally { File.Delete(path); }
    }

    sealed class InputBackend(params uint[] counts) : Input.IInputBackend
    {
        public bool SetCursorPosition(int x, int y) => throw new InvalidOperationException("Unexpected cursor movement in test.");
        public void SetClipboardText(string text) => throw new InvalidOperationException("Unexpected clipboard write in test.");
        readonly Queue<uint> counts = new(counts);
        internal List<Input.INPUT[]> Calls { get; } = [];
        public int LastError => 5;
        public uint Send(Input.INPUT[] inputs)
        {
            Calls.Add(inputs);
            return counts.Count == 0 ? (uint)inputs.Length : counts.Dequeue();
        }
    }

    static void SingleKeyFailure()
    {
        using var scope = A.BeginRun();
        A.Options.SendInputMaxRetries = 0;
        var previous = Input.InputBackend;
        var fake = new InputBackend(1, 1);
        try
        {
            Input.InputBackend = fake;
            var failed = false;
            try { Input.PressKey("tab"); }
            catch (InvalidOperationException) { failed = true; }
            Check(failed && fake.Calls.Count == 2, "partial key press was hidden");
            Check(fake.Calls[1].Length == 1 && fake.Calls[1][0].U.ki.dwFlags == 2, "accepted key-down was not released");
        }
        finally { Input.InputBackend = previous; }
    }

    static void RetryRelease()
    {
        using var scope = A.BeginRun();
        A.Options.SendInputMaxRetries = 0;
        var previous = Input.InputBackend;
        var fake = new InputBackend(3, 0, 1);
        try
        {
            Input.InputBackend = fake;
            try { Input.PressChord(["ctrl"], "x"); }
            catch (InvalidOperationException) { }
            Check(fake.Calls.Count == 2, "emergency release was not attempted");
            A.CurrentRun.CancelRequested = true;
            Input.ReleaseAllHeldInputs();
            Check(fake.Calls.Count == 3 && fake.Calls[2].Length == 1 &&
                fake.Calls[2][0].U.ki.wVk == 0x11 && fake.Calls[2][0].U.ki.dwFlags == 2,
                "cleanup did not retry the unaccepted Ctrl release");
        }
        finally { Input.InputBackend = previous; }
    }

    static void CancelledInput()
    {
        using var scope = A.BeginRun();
        A.CurrentRun.CancelRequested = true;
        var previous = Input.InputBackend;
        var fake = new InputBackend();
        try
        {
            Input.InputBackend = fake;
            var cancelled = false;
            try { Input.MouseScroll(120); }
            catch (OperationCanceledException) { cancelled = true; }
            Check(cancelled && fake.Calls.Count == 0, "wheel input was sent after cancellation");
            Input.KeyUp("ctrl");
            Check(fake.Calls.Count == 1, "cancellation prevented modifier release");
        }
        finally { Input.InputBackend = previous; }
    }
}
