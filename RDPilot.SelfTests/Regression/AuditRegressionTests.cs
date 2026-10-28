using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;

namespace RDPilot.SelfTests.Regression;

internal static class AuditRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("backup recovery preserves valid data when retention trims it", RecoveryRetention),
        ("Terra requests include configured reasoning and omit temperature", TerraRequest),
        ("helper image pixels and metadata share one source", HelperImage),
        ("fast URL completion rejects multi-step tasks", FastUrl),
        ("partial input sends only the remaining suffix", PartialInput),
        ("failed partial chords release modifiers without replay", FailedChord),
        ("transport timeout covers stalled SSE bodies", () => Timeout().GetAwaiter().GetResult()),
        ("transport distinguishes user cancellation", () => Cancellation().GetAwaiter().GetResult())
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class Settings : IDisposable
    {
        readonly Dictionary<TestState.Slot, object?> originals = new();
        internal void Set(string name, object? value)
        {
            var field = TestState.Field(typeof(A), name, BindingFlags.Static | BindingFlags.NonPublic)!;
            originals.TryAdd(field, field.GetValue(null));
            field.SetValue(null, value);
        }
        public void Dispose()
        {
            foreach (var (field, value) in originals) field.SetValue(null, value);
        }
    }

    static void RecoveryRetention()
    {
        using var settings = new Settings();
        var root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, "memory.json");
        settings.Set("RecoveryMemoryEnabled", true);
        settings.Set("RecoveryMemoryPath", path);
        settings.Set("RecoveryMemoryArchivePath", Path.Combine(root, "archive.json"));
        settings.Set("RecoveryMemoryMaxQuarantinedLessons", 1);
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        var store = new RecoveryLessonStore { Lessons =
        [
            new() { Id = "keep", Status = "quarantined", UpdatedUtc = DateTime.UtcNow },
            new() { Id = "archive", Status = "quarantined", UpdatedUtc = DateTime.UtcNow.AddDays(-1) }
        ] };
        try
        {
            File.WriteAllText(path, "{broken-json");
            File.WriteAllText(path + ".bak", JsonSerializer.Serialize(store, options));
            var loaded = RDPilot.Persistence.RecoveryMemoryService.LoadRecoveryLessons();
            var primary = JsonSerializer.Deserialize<RecoveryLessonStore>(File.ReadAllText(path), options)!;
            var backup = JsonSerializer.Deserialize<RecoveryLessonStore>(File.ReadAllText(path + ".bak"), options)!;
            Check(loaded.Count == 1 && primary.Lessons.Single().Id == "keep", "recovered primary is not the retained valid store");
            Check(backup.Lessons.Count == 2, "original valid backup was overwritten");
            Check(RDPilot.Persistence.RecoveryMemoryService.LoadRecoveryLessons().Single().Id == "keep", "restarting loses recovered data");
        }
        finally { Directory.Delete(root, true); }
    }

    static void TerraRequest()
    {
        using var settings = new Settings();
        settings.Set("ReasoningEffort", "high");
        var request = RDPilot.OpenAI.PromptAndRequestFactory.BuildQARequestBody("gpt-6.1-terra", "rules", "question", "data:image/png;base64,test", 1280,720,0,0,0,0,null,null,new("window","app",null,null,null,null));
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(request));
        Check(doc.RootElement.GetProperty("reasoning").GetProperty("effort").GetString() == "high", "configured effort is missing");
        Check(!doc.RootElement.TryGetProperty("temperature",out _), "temperature is unsupported");
    }

    static void HelperImage()
    {
        using var bitmap = new Bitmap(1280,720);
        var original = RDPilot.Observation.ScreenshotService.EncodeBitmapToDataUrl(bitmap,0,"png",80);
        var image = RDPilot.Observation.ScreenshotService.PrepareHelperImage(original,1024);
        using var stream = new MemoryStream(Convert.FromBase64String(image.DataUrl.Split(',')[1]));
        using var actual = Image.FromStream(stream);
        Check(actual.Width == 1024 && actual.Height == 576, "helper limit was not applied");
        Check(actual.Width == image.Width && actual.Height == image.Height, "metadata differs from encoded image");
        var unchanged = RDPilot.Observation.ScreenshotService.PrepareHelperImage(original,0);
        Check(unchanged.DataUrl == original && unchanged.Width == 1280, "unlimited helper unexpectedly changed image");
    }

    static void FastUrl()
    {
        foreach (var goal in new[] { "Open https://example.com and fill in the form", "Otwórz example.com i zaloguj się", "Copy the URL https://example.com", "Read report.txt and summarize it" })
            Check(!RDPilot.Control.ControlGoalPolicy.TryExtractUrl(goal,out _), "multi-step goal incorrectly completed: " + goal);
        Check(RDPilot.Control.ControlGoalPolicy.TryExtractUrl("open https://example.com/path?q=test",out var url) && url == "https://example.com/path?q=test", "single URL action no longer works");
    }

    sealed class FakeInput(params uint[] results) : Input.IInputBackend
    {
        public bool SetCursorPosition(int x, int y) => throw new InvalidOperationException("Unexpected cursor movement in test.");
        public void SetClipboardText(string text) => throw new InvalidOperationException("Unexpected clipboard write in test.");
        readonly Queue<uint> results = new(results);
        internal List<Input.INPUT[]> Calls { get; } = [];
        public int LastError => 5;
        public uint Send(Input.INPUT[] inputs)
        {
            Calls.Add(inputs);
            return results.Count > 0 ? results.Dequeue() : (uint)inputs.Length;
        }
    }

    static void PartialInput()
    {
        using var settings = new Settings();
        settings.Set("SendInputMaxRetries",2);
        settings.Set("SendInputRetryDelayMs",0);
        var original = Input.InputBackend;
        var fake = new FakeInput(3,1);
        try
        {
            Input.InputBackend = fake;
            Input.PressChord(["ctrl"],"x");
            Check(fake.Calls.Count == 2 && fake.Calls[0].Length == 4 && fake.Calls[1].Length == 1, "accepted input prefix was replayed");
            Check(fake.Calls[1][0].U.ki.wVk == 0x11 && fake.Calls[1][0].U.ki.dwFlags == 2, "wrong remaining input");
        }
        finally { Input.InputBackend = original; }
    }

    static void FailedChord()
    {
        using var settings = new Settings();
        settings.Set("SendInputMaxRetries",1);
        settings.Set("SendInputRetryDelayMs",0);
        var original = Input.InputBackend;
        var fake = new FakeInput(3,0,1);
        try
        {
            Input.InputBackend = fake;
            var failed = false;
            try { Input.PressChord(["ctrl"],"x"); }
            catch (InvalidOperationException) { failed = true; }
            Check(failed && fake.Calls.Count == 3, "failure must be reported without replaying chord");
            Check(fake.Calls[2].Length == 1 && fake.Calls[2][0].U.ki.wVk == 0x11 && fake.Calls[2][0].U.ki.dwFlags == 2, "Ctrl was not released");
        }
        finally { Input.InputBackend = original; }
    }

    sealed class StalledHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(new StalledStream())
                { Headers = { ContentType = new("text/event-stream") } }
            });
    }
    sealed class StalledStream : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => 0; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer,int offset,int count) => throw new NotSupportedException();
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer,CancellationToken token = default)
        { await Task.Delay(System.Threading.Timeout.Infinite,token); return 0; }
        public override Task<int> ReadAsync(byte[] buffer,int offset,int count,CancellationToken token) => ReadAsync(buffer.AsMemory(offset,count),token).AsTask();
        public override long Seek(long offset,SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer,int offset,int count) => throw new NotSupportedException();
    }

    static async Task Timeout()
    {
        using var settings = new Settings();
        settings.Set("OpenAiTimeoutSeconds",1);
        settings.Set("OpenAiMaxRetries",0);
        var original = RDPilot.OpenAI.OpenAiTransport.HttpClient;
        using var client = new HttpClient(new StalledHandler());
        using var watchdog = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        try
        {
            RDPilot.OpenAI.OpenAiTransport.HttpClient = client;
            var result = await RDPilot.OpenAI.OpenAiTransport.SendOpenAIControlRequestAsync("test-only",new { model = "test" },watchdog.Token);
            Check(!result.ok && !watchdog.IsCancellationRequested && result.raw != "cancelled", "stream did not observe its own timeout");
        }
        finally { RDPilot.OpenAI.OpenAiTransport.HttpClient = original; }
    }

    static async Task Cancellation()
    {
        using var settings = new Settings();
        settings.Set("OpenAiTimeoutSeconds",10);
        settings.Set("OpenAiMaxRetries",2);
        var original = RDPilot.OpenAI.OpenAiTransport.HttpClient;
        using var client = new HttpClient(new StalledHandler());
        using var cancellation = new CancellationTokenSource(50);
        try
        {
            RDPilot.OpenAI.OpenAiTransport.HttpClient = client;
            var result = await RDPilot.OpenAI.OpenAiTransport.SendOpenAIControlRequestAsync("test-only",new { model = "test" },cancellation.Token);
            Check(!result.ok && result.raw == "cancelled", "user cancellation was lost");
        }
        finally { RDPilot.OpenAI.OpenAiTransport.HttpClient = original; }
    }
}
