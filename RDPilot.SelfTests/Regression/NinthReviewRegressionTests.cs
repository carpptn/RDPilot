using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;

namespace RDPilot.SelfTests.Regression;

internal static class NinthReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("punctuation shortcuts use physical keys from the layout", Punctuation),
        ("plus shortcut adds Shift and releases modifiers in reverse order", Plus),
        ("layout-required modifiers are not duplicated", LayoutModifiers),
        ("unmappable shortcut rejects the entire sequence before input", Unmappable),
        ("ordinary Unicode input remains text", UnicodeText),
        ("live old writer continues accumulating after compaction", ActiveWriter),
        ("compaction stops before the oldest live writer", ActiveCutoff),
        ("peer writer is protected until its process exits", PeerWriter),
        ("writer identity rotates past a previously compacted cutoff", RotateWriter),
        ("first write to a new counter gets a fresh timestamp", NewCounter),
        ("response replay reports missing malformed and non-action files", ReplayFailures),
        ("response replay accepts a valid action payload", ReplaySuccess),
        ("response replay CLI propagates failure exit code", () => ReplayCli(false)),
        ("response replay CLI preserves success exit code", () => ReplayCli(true)),
        ("interactive Q&A reports an empty answer as failure", () => Question(CommandOutcome.Failed)),
        ("interactive Q&A reports successful answer", () => Question(CommandOutcome.Success)),
        ("interactive Q&A reports cancellation even with a late answer", () => Question(CommandOutcome.Cancelled)),
        ("Q&A presentation uses the same status in both modes", OutcomePresentation)
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
    static object? Invoke(string method, params object?[] args) => typeof(RecoveryMemoryService)
        .GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, args);

    sealed class Mapping(Func<char, KeyboardInputEncoder.CharacterKey?> mapper) : IDisposable
    {
        readonly Func<char, KeyboardInputEncoder.CharacterKey?> previous = KeyboardInputEncoder.CharacterMapper;
        internal void Apply() => KeyboardInputEncoder.CharacterMapper = mapper;
        public void Dispose() => KeyboardInputEncoder.CharacterMapper = previous;
    }
    static void Punctuation()
    {
        using var f = new Fixture();
        foreach (var key in new ushort[] { 0xBD, 0xBF })
        {
            using var mapping = new Mapping(_ => new(key, 0)); mapping.Apply();
            f.Input.Events.Clear();
            Input.ExecuteAction(new() { Type = "keys", Keys = ["ctrl+-"] });
            Check(f.Input.Events.Select(e => e.U.ki.wVk).SequenceEqual(new ushort[] { 0x11, key, key, 0x11 }), "layout mapping was ignored");
            Check(f.Input.Events.All(e => (e.U.ki.dwFlags & 4) == 0), "shortcut used VK_PACKET");
        }
    }
    static void Plus()
    {
        using var f = new Fixture();
        using var mapping = new Mapping(ch => ch == '+' ? new(0xBB, 1) : null); mapping.Apply();
        foreach (var keys in new[] { new[] { "ctrl++" }, new[] { "ctrl", "+" }, new[] { "ctrl+shift++" } })
        {
            f.Input.Events.Clear();
            Input.ExecuteAction(new() { Type = "keys", Keys = keys });
            Check(f.Input.Events.Select(e => e.U.ki.wVk).SequenceEqual(new ushort[] { 0x11, 0x10, 0xBB, 0xBB, 0x10, 0x11 }), "incorrect plus shortcut");
            Check(f.Input.Events.Select(e => e.U.ki.dwFlags).SequenceEqual(new uint[] { 0, 0, 0, 2, 2, 2 }), "modifiers were not released");
        }
    }
    static void LayoutModifiers()
    {
        using var f = new Fixture();
        using var mapping = new Mapping(_ => new(0x51, 6)); mapping.Apply();
        Input.ExecuteAction(new() { Type = "keys", Keys = ["ctrl+@"] });
        Check(f.Input.Events.Select(e => e.U.ki.wVk).SequenceEqual(new ushort[] { 0x11, 0x12, 0x51, 0x51, 0x12, 0x11 }), "Ctrl was duplicated or layout modifiers lost");
    }
    static void Unmappable()
    {
        using var f = new Fixture();
        using var mapping = new Mapping(_ => null); mapping.Apply();
        var action = new ActionDto { Type = "keys", Keys = ["tab", "ctrl+ż"] };
        Check(!Input.CaptureResolvedAction(action, null).IsValid, "unmappable shortcut passed validation");
        try { Input.ExecuteAction(action); throw new Exception("unmappable shortcut executed"); }
        catch (ArgumentException) { }
        Check(f.Input.Events.Count == 0, "earlier key was sent before rejecting shortcut");
    }
    static void UnicodeText()
    {
        using var f = new Fixture();
        using var mapping = new Mapping(_ => throw new Exception("text invoked physical mapper")); mapping.Apply();
        Input.ExecuteAction(new() { Type = "type_text", Text = "ż+" });
        Check(f.Input.Events.Count == 4 && f.Input.Events.All(e => (e.U.ki.dwFlags & 4) != 0), "Unicode text behavior changed");
    }

    sealed class Future(DateTime now) : TimeProvider { public override DateTimeOffset GetUtcNow() => new(now); }
    static LoopCalibrationBucket Components(string active, int before, int after)
    {
        var started = RecoveryWriterCounters.WriterStartedUtc(active);
        var result = new LoopCalibrationBucket { CandidateByWriter = new() { [active] = 1 } };
        foreach (var offset in Enumerable.Range(-before, before + after + 1).Where(i => i != 0))
            result.CandidateByWriter[$"{started.AddMinutes(offset).Ticks:x16}-legacy-{offset}"] = 1;
        RecoveryMemoryService.NormalizeCalibrationBucket(result);
        return result;
    }
    static void ActiveWriter()
    {
        var writer = RecoveryWriterSession.CurrentId;
        var clock = RuntimeClock.Provider;
        try
        {
            RuntimeClock.Provider = new Future(RecoveryWriterCounters.WriterStartedUtc(writer).AddDays(8));
            var bucket = Components(writer, 0, 64);
            Invoke("CompactCalibrationWriterCounters", bucket);
            Invoke("IncrementWriterCounter", bucket.CandidateByWriter, bucket.CountersCompactedBeforeUtc);
            RecoveryMemoryService.NormalizeCalibrationBucket(bucket);
            Check(bucket.CandidateCount == 66 && bucket.CandidateByWriter[writer] == 2, "live writer lost its next event");
            var lesson = new RecoveryLesson { SuccessByWriter = new(bucket.CandidateByWriter) };
            RecoveryMemoryService.NormalizeLessonCounters(lesson);
            Invoke("CompactLessonWriterCounters", lesson);
            Invoke("RecordLessonSuccess", lesson, 0.5);
            Check(lesson.SuccessCount == 67 && lesson.CumulativeReward == 0.5 && lesson.RewardObservationCount == 1, "live lesson writer lost counters");
        }
        finally { RuntimeClock.Provider = clock; }
    }
    static void ActiveCutoff()
    {
        var writer = RecoveryWriterSession.CurrentId;
        var started = RecoveryWriterCounters.WriterStartedUtc(writer);
        var clock = RuntimeClock.Provider;
        try
        {
            RuntimeClock.Provider = new Future(started.AddDays(8));
            var bucket = Components(writer, 10, 60);
            Invoke("CompactCalibrationWriterCounters", bucket);
            Check(bucket.CandidateCount == 71 && bucket.CompactedCandidateCount == 10 &&
                bucket.CountersCompactedBeforeUtc < started && bucket.CandidateByWriter.ContainsKey(writer), "cutoff crossed an active writer");
        }
        finally { RuntimeClock.Provider = clock; }
    }
    static Process StartProbe(params string[] args)
    {
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        {
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true,
            RedirectStandardOutput = true, RedirectStandardError = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        foreach (var arg in args) start.ArgumentList.Add(arg);
        return Process.Start(start)!;
    }
    static void PeerWriter()
    {
        using var peer = StartProbe("--recovery-writer-probe");
        var clock = RuntimeClock.Provider;
        try
        {
            var writer = peer.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult()!;
            Check(RecoveryWriterSession.IsActive(writer), "peer process was not recognized");
            RuntimeClock.Provider = new Future(RecoveryWriterCounters.WriterStartedUtc(writer).AddDays(8));
            var bucket = Components(writer, 0, 64);
            Invoke("CompactCalibrationWriterCounters", bucket);
            Check(bucket.CandidateByWriter.ContainsKey(writer), "active peer was compacted");
            peer.StandardInput.WriteLine();
            Check(peer.WaitForExit(10000), "peer did not exit");
            Check(!RecoveryWriterSession.IsActive(writer), "exited peer still appears active");
            Invoke("CompactCalibrationWriterCounters", bucket);
            Check(bucket.CandidateCount == 65 && bucket.CandidateByWriter.Count == 32 && !bucket.CandidateByWriter.ContainsKey(writer), "inactive peer was not compacted safely");
        }
        finally
        {
            RuntimeClock.Provider = clock;
            if (!peer.HasExited) { peer.Kill(); peer.WaitForExit(); }
        }
    }
    static void RotateWriter()
    {
        var field = typeof(RecoveryWriterSession).GetField("currentId", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = RecoveryWriterSession.CurrentId;
        try
        {
            var cutoff = RecoveryWriterCounters.WriterStartedUtc(original).AddDays(1);
            var lesson = new RecoveryLesson { SuccessCount = 10, CompactedSuccessCount = 10, CountersCompactedBeforeUtc = cutoff };
            Invoke("RecordLessonSuccess", lesson, 0.75);
            Check(lesson.SuccessCount == 11 && lesson.CumulativeReward == 0.75 && lesson.RewardObservationCount == 1, "rotated writer lost events");
            Check(RecoveryWriterCounters.WriterStartedUtc(RecoveryWriterSession.CurrentId) > cutoff, "writer did not rotate beyond cutoff");
        }
        finally { field.SetValue(null, original); }
    }

    static void NewCounter()
    {
        var field = typeof(RecoveryWriterSession).GetField("currentId", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = RecoveryWriterSession.CurrentId;
        var started = RecoveryWriterCounters.WriterStartedUtc(original);
        var clock = RuntimeClock.Provider;
        try
        {
            RuntimeClock.Provider = new Future(started.AddDays(20));
            var bucket = new LoopCalibrationBucket();
            Invoke("IncrementWriterCounter", bucket.CandidateByWriter, null);
            var writer = bucket.CandidateByWriter.Keys.Single();
            Check(RecoveryWriterCounters.WriterStartedUtc(writer) == started.AddDays(20), "new counter inherited stale writer timestamp");
            // Another process compacted this bucket after our earlier load, before our first write.
            bucket.CountersCompactedBeforeUtc = started.AddDays(10);
            RecoveryMemoryService.NormalizeCalibrationBucket(bucket);
            Check(bucket.CandidateCount == 1, "first write was mistaken for compacted history");
            Invoke("IncrementWriterCounter", bucket.CandidateByWriter, bucket.CountersCompactedBeforeUtc);
            Check(bucket.CandidateByWriter.Count == 1 && bucket.CandidateByWriter[writer] == 2, "existing component was not reused");
        }
        finally { RuntimeClock.Provider = clock; field.SetValue(null, original); }
    }

    sealed class ReplayFile : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string FilePath => Path.Combine(Root, "response.json");
        internal ReplayFile() => Directory.CreateDirectory(Root);
        internal void Write(string text) => File.WriteAllText(FilePath, text.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
        public void Dispose() => Directory.Delete(Root, true);
    }
    static string ValidReplay => JsonSerializer.Serialize(new { status = "completed", output_text = "{\"actions\":[{\"type\":\"wait\",\"wait_seconds\":1}]}" });
    static void ReplayFailures()
    {
        using var file = new ReplayFile();
        Check(ArtifactAnalysisService.ReplayResponse(file.FilePath) == CommandOutcome.Failed, "missing response succeeded");
        foreach (var text in new[] { "", "not json", "{", "{}", "{\"output_text\":\"{}\"}" })
        {
            file.Write(text);
            Check(ArtifactAnalysisService.ReplayResponse(file.FilePath) == CommandOutcome.Failed, "bad response succeeded");
        }
    }
    static void ReplaySuccess()
    {
        using var file = new ReplayFile(); file.Write(ValidReplay);
        Check(ArtifactAnalysisService.ReplayResponse(file.FilePath) == CommandOutcome.Success, "valid response failed");
    }
    static void ReplayCli(bool valid)
    {
        using var file = new ReplayFile(); if (valid) file.Write(ValidReplay);
        using var child = StartProbe("--command-main-probe", "--replay-response", file.FilePath);
        try
        {
            var stdout = child.StandardOutput.ReadToEndAsync(); var stderr = child.StandardError.ReadToEndAsync();
            Check(child.WaitForExit(10000), "CLI did not exit");
            Task.WhenAll(stdout, stderr).GetAwaiter().GetResult();
            Check(child.ExitCode == (valid ? 0 : 1), $"wrong CLI exit code: {child.ExitCode}; {stderr.Result}");
        }
        finally { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } }
    }

    sealed class Capture : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multiple) => new(0, 0, 320, 200);
        public void CopyPixels(Graphics graphics, Rectangle region) => graphics.Clear(Color.Navy);
    }
    sealed class Response(CommandOutcome outcome) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            if (outcome == CommandOutcome.Cancelled) DesktopConsoleService.CancelRequested = true;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(new
            {
                status = "completed", output_text = outcome == CommandOutcome.Failed ? "{}" : "{\"answer_text\":\"A screen.\"}"
            }), Encoding.UTF8, "application/json") });
        }
    }
    static void Question(CommandOutcome expected)
    {
        var options = typeof(AppOptions).GetFields(BindingFlags.Instance | BindingFlags.NonPublic)
            .Where(f => !f.IsLiteral && !f.IsInitOnly).Select(f => (Field: f, Value: f.GetValue(A.Options))).ToArray();
        var capture = DesktopCapture.Backend; var window = ScreenshotService.ActiveWindowRectangleProvider;
        var context = UiAutomationService.PromptContextProvider; var http = OpenAiTransport.HttpClient;
        var output = Console.Out; var error = Console.Error;
        using var handler = new Response(expected); using var client = new HttpClient(handler); using var transcript = new StringWriter();
        try
        {
            A.Options.BatchMode = true; A.Options.LogScreens = A.Options.LogRequests = false;
            A.Options.IncludeFocusUia = A.Options.IncludeUiaTargets = false; A.Options.MaxArtifactsPerDir = 0;
            DesktopCapture.Backend = new Capture();
            ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(0, 0, 320, 200);
            UiAutomationService.PromptContextProvider = (_, _, _) => new("Synthetic", "synthetic", null, null, null, null);
            OpenAiTransport.HttpClient = client; Console.SetOut(transcript); Console.SetError(transcript);
            var actual = A.RunInteractiveQuestionAsync("synthetic-key", "/ask What is displayed?").GetAwaiter().GetResult();
            Check(actual == expected, "wrong interactive result");
            var text = transcript.ToString();
            Check(text.Contains(expected switch { CommandOutcome.Success => "Answer completed", CommandOutcome.Cancelled => "Answer cancelled", _ => "Answer failed" }), "wrong completion message");
            Check(expected == CommandOutcome.Success || !text.Contains("Answer completed"), "failure reported as success");
        }
        finally
        {
            Console.SetOut(output); Console.SetError(error); OpenAiTransport.HttpClient = http;
            DesktopCapture.Backend = capture; ScreenshotService.ActiveWindowRectangleProvider = window; UiAutomationService.PromptContextProvider = context;
            foreach (var option in options) option.Field.SetValue(A.Options, option.Value);
        }
    }
    static void OutcomePresentation()
    {
        foreach (var outcome in Enum.GetValues<CommandOutcome>())
        {
            var stdout = Console.Out; var stderr = Console.Error; using var transcript = new StringWriter();
            try
            {
                Console.SetOut(transcript); Console.SetError(transcript);
                CommandPresentation.WriteQuestionResult(outcome, batch: true);
                Check(transcript.ToString().Contains(outcome switch { CommandOutcome.Success => "completed", CommandOutcome.Cancelled => "cancelled", _ => "failed" }), "batch presentation disagrees with outcome");
                Check((int)outcome == (outcome == CommandOutcome.Success ? 0 : outcome == CommandOutcome.Cancelled ? 130 : 1), "wrong outcome exit code");
            }
            finally { Console.SetOut(stdout); Console.SetError(stderr); }
        }
    }
}
