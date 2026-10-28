using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;

namespace RDPilot.SelfTests.Regression;

internal static class ThirdReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("pending recovery writes survive failed cleanup and reload", PendingWrites),
        ("malformed response shapes are rejected without escaping the transport", MalformedResponses),
        ("JSON action parsing respects braces and escaped quotes inside text", JsonStrings),
        ("incomplete helper payloads are retried before acceptance", () => HelperIncomplete(false)),
        ("incomplete helper payloads are rejected when retry budget expires", () => HelperIncomplete(true)),
        ("SSE logical server failures retry using the common failure policy", () => LogicalFailure(true, "server_error", 2)),
        ("JSON logical rate limits retry using the common failure policy", () => LogicalFailure(false, "rate_limit_exceeded", 2)),
        ("permanent SSE logical failures do not retry", () => LogicalFailure(true, "invalid_prompt", 1)),
        ("context fallback restores the image and standalone input", ContextFallback),
        ("JSON configuration rejects wrong types, invalid ranges and missing files", Configuration),
        ("JSON nullable settings reset earlier configuration layers", NullableConfiguration),
        ("replay is independent of live desktop geometry and UIA", ReplayGeometry),
        ("replay validates redacted, malformed, missing and valid images", ReplayImages),
        ("failed cursor movement blocks a mouse click", CursorFailure),
        ("clipboard failure blocks paste keystrokes", ClipboardFailure)
    ];

    static void Check(bool value, string reason)
    {
        if (!value) throw new InvalidOperationException(reason);
    }

    static string TempDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    static void Write(string path, string value) => File.WriteAllText(path,
        value.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));

    sealed class FailingFiles(AtomicTextStore.IFileOperations inner) : AtomicTextStore.IFileOperations
    {
        public bool Exists(string path) => inner.Exists(path);
        public void Write(string path, string value) => throw new IOException("Synthetic disk failure");
        public void Replace(string source, string destination, string? backup) => inner.Replace(source, destination, backup);
        public void Move(string source, string destination) => inner.Move(source, destination);
        public void Delete(string path) => inner.Delete(path);
    }

    static void PendingWrites()
    {
        using var run = A.BeginRun();
        var root = TempDirectory();
        var files = AtomicTextStore.Files;
        var save = typeof(RecoveryMemoryService).GetMethod("SaveRecoveryLessons", BindingFlags.Static | BindingFlags.NonPublic)!;
        try
        {
            A.Options.RecoveryMemoryEnabled = true;
            var path = Path.Combine(root, "first.json");
            A.Options.RecoveryMemoryPath = path;
            A.Options.RecoveryMemoryArchivePath = Path.Combine(root, "archive.json");
            var lessons = RecoveryMemoryService.LoadRecoveryLessons();
            AtomicTextStore.Files = new FailingFiles(files);
            lessons.Add(new() { Id = "valuable", UpdatedUtc = DateTime.UtcNow, WinningStrategy = "preserve original" });
            Check(!(bool)save.Invoke(null, [lessons])!, "failed write reported success");
            Check(!RecoveryMemoryService.FlushPendingRecoveryMemory(), "failed cleanup reported success");
            lessons[0].WinningStrategy = "mutated after save";
            AtomicTextStore.Files = files;
            A.Options.RecoveryMemoryPath = Path.Combine(root, "other.json");
            Check(RecoveryMemoryService.LoadRecoveryLessons().Count == 0, "pending writes leaked to another file");
            A.Options.RecoveryMemoryPath = path;
            lessons = RecoveryMemoryService.LoadRecoveryLessons();
            Check(lessons.Single().WinningStrategy == "preserve original", "pending write did not preserve its snapshot");
            lessons[0].WinningStrategy = "mutated after reload";
            lessons = RecoveryMemoryService.LoadRecoveryLessons();
            Check(lessons.Single().WinningStrategy == "preserve original", "loaded lessons shared pending snapshot references");
            lessons.Add(new() { Id = "next-run", UpdatedUtc = DateTime.UtcNow });
            Check((bool)save.Invoke(null, [lessons])!, "next write failed");
            Check(RecoveryMemoryService.FlushPendingRecoveryMemory(), "pending writes did not clear");
            Check(RecoveryMemoryService.LoadRecoveryLessons().Select(l => l.Id).Order().SequenceEqual(new[] { "next-run", "valuable" }), "pending lesson disappeared");
        }
        finally { AtomicTextStore.Files = files; Directory.Delete(root, true); }
    }

    static void MalformedResponses()
    {
        using var run = A.BeginRun();
        A.Options.OpenAiMaxRetries = 0;
        var old = OpenAiTransport.HttpClient;
        try
        {
            foreach (var raw in new[] { "[]", "null", "42", "{}", "{\"status\":\"in_progress\"}", "{\"status\":\"completed\",\"error\":\"failure\"}" })
            {
                using var handler = new Responses(_ => new(HttpStatusCode.OK) { Content = new StringContent(raw) });
                using var client = new HttpClient(handler);
                OpenAiTransport.HttpClient = client;
                var result = OpenAiTransport.SendOpenAIRequestAsync("synthetic-key", new { input = "test" }).GetAwaiter().GetResult();
                Check(!result.ok && !A.CurrentRun.LastOpenAiFailureWasRetriable, "invalid response escaped validation: " + raw);
            }
            string? id = null;
            Check(!OpenAiResponsesService.TryHandleControlStreamEvent("[]", false, ref id, out _, out _, out _), "invalid SSE object accepted");
        }
        finally { OpenAiTransport.HttpClient = old; }
    }

    static void JsonStrings()
    {
        foreach (var text in new[] { "}", "{", "function x() {", "quote \" then }", "backslash \\ and {\n}" })
        {
            var json = JsonSerializer.Serialize(new { actions = new[] { new ActionDto { Type = "type_text", Text = text } } });
            Check(OpenAiResponsesService.TryParseControlActionJson(json, out var actions) && actions.Single().Text == text, "text was rejected or changed: " + text);
        }
    }

    sealed class Responses(Func<int, HttpResponseMessage> response) : HttpMessageHandler
    {
        internal List<string> Bodies = [];
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            Bodies.Add(await request.Content!.ReadAsStringAsync(token));
            return response(Bodies.Count);
        }
    }

    static HttpResponseMessage JsonResponse(object response, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(response), Encoding.UTF8, "application/json") };

    static object CompleteAction() => new { id = "complete", status = "completed", output_text = "{\"actions\":[{\"type\":\"wait\",\"wait_seconds\":0}]}" };

    static void HelperIncomplete(bool exhaust)
    {
        using var run = A.BeginRun();
        A.Options.IncompleteMaxOutputRetries = 0;
        A.Options.LogRequests = false;
        var old = OpenAiTransport.HttpClient;
        using var handler = new Responses(call => JsonResponse(new
        {
            status = exhaust || call == 1 ? "incomplete" : "completed",
            incomplete_details = new { reason = "max_output_tokens" },
            output_text = "{\"verdict\":\"yes\",\"reason\":\"test\"}"
        }));
        using var client = new HttpClient(handler);
        try
        {
            OpenAiTransport.HttpClient = client;
            var result = OpenAiResponsesService.CallOpenAIParsedAsync<VerifyDto>("synthetic-key",
                new { model = "synthetic", input = "test", max_output_tokens = 1000, reasoning = new { effort = "medium" } }).GetAwaiter().GetResult();
            Check(handler.Bodies.Count == 2, "incomplete output was accepted before the bounded effort retry");
            Check(exhaust ? result.parsed is null : result.parsed?.Verdict == "yes", "incorrect helper acceptance");
        }
        finally { OpenAiTransport.HttpClient = old; }
    }

    static void LogicalFailure(bool stream, string code, int expectedCalls)
    {
        using var run = A.BeginRun();
        A.Options.OpenAiMaxRetries = 1;
        A.Options.UsePreviousResponseState = true;
        A.Options.LogRequests = false;
        var old = OpenAiTransport.HttpClient;
        using var handler = new Responses(call =>
        {
            if (call > 1) return JsonResponse(CompleteAction());
            var failed = new { id = "failed", status = "failed", error = new { code, message = "synthetic failure" } };
            return stream ? new(HttpStatusCode.OK) { Content = new StringContent("data: " + JsonSerializer.Serialize(new { type = "response.failed", response = failed }) + "\n\n", Encoding.UTF8, "text/event-stream") } : JsonResponse(failed);
        });
        using var client = new HttpClient(handler);
        try
        {
            OpenAiTransport.HttpClient = client;
            var result = OpenAiResponsesService.CallOpenAIAsync("synthetic-key", new { input = "test" }).GetAwaiter().GetResult();
            Check(handler.Bodies.Count == expectedCalls, "logical failure retry count is incorrect");
            Check((result.parsed is not null) == (expectedCalls > 1), "logical failure acceptance is incorrect");
            Check(A.CurrentRun.LastOpenAiResponseId == (expectedCalls > 1 ? "complete" : null), "failed response ID leaked");
        }
        finally { OpenAiTransport.HttpClient = old; }
    }

    static void ContextFallback()
    {
        using var run = A.BeginRun();
        A.Options.LogRequests = false;
        var old = OpenAiTransport.HttpClient;
        using var handler = new Responses(call => call == 1
            ? JsonResponse(new { error = new { message = "previous_response_id not found", param = "previous_response_id" } }, HttpStatusCode.NotFound)
            : JsonResponse(CompleteAction()));
        using var client = new HttpClient(handler);
        var standalone = new { input = new[] { new { role = "user", content = new object[] { new { type = "input_text", text = "history and full UIA targets" }, new { type = "input_image", image_url = "data:image/png;base64,synthetic" } } } }, max_output_tokens = 100 };
        try
        {
            OpenAiTransport.HttpClient = client;
            var request = new { input = new[] { new { type = "input_text", text = "omitted image" } }, previous_response_id = "expired", max_output_tokens = 2000 };
            var result = OpenAiResponsesService.CallOpenAIAsync("synthetic-key", request, standaloneRequest: standalone).GetAwaiter().GetResult();
            var retry = JsonNode.Parse(handler.Bodies[1])!;
            Check(result.parsed is not null && result.contextFallbackUsed, "context fallback failed");
            Check(retry["previous_response_id"] is null && retry["max_output_tokens"]!.GetValue<int>() == 2000, "fallback lost retry settings or retained response ID");
            Check(handler.Bodies[1].Contains("input_image") && handler.Bodies[1].Contains("history and full UIA targets") && !handler.Bodies[1].Contains("omitted image"), "standalone context was not restored");
            Check(!OpenAiResponsesService.TryBuildStandaloneRetry(request, null, out _), "fallback accepted a request with no image");
        }
        finally { OpenAiTransport.HttpClient = old; }
    }

    static void Configuration()
    {
        using var run = A.BeginRun();
        var root = TempDirectory();
        try
        {
            var file = Path.Combine(root, "config.json");
            A.Options.MaxSteps = 12;
            A.Options.MouseEnabled = true;
            Write(file, "{\"maxSteps\":\"2\",\"mouseEnabled\":\"false\",\"openAiMaxRetries\":11,\"skipVerifyConfidenceThreshold\":-1}");
            ConfigurationService.ApplyConfigFile(file);
            Check(A.CurrentRun.ConfigurationErrors.Count == 4, "invalid JSON scalars were silently accepted");
            Write(file, "{\"maxSteps\":-1}");
            ConfigurationService.ApplyConfigFile(file);
            ConfigurationService.ApplyConfigFile(file + ".missing");
            Check(A.CurrentRun.ConfigurationErrors.Count == 6 && A.Options.MaxSteps == 12 && A.Options.MouseEnabled, "invalid limits or missing file changed safe settings");
            Write(file, "{\"maxSteps\":0,\"mouseEnabled\":false,\"openAiMaxRetries\":10}");
            ConfigurationService.ApplyConfigFile(file);
            Check(A.Options.MaxSteps == 0 && !A.Options.MouseEnabled && A.Options.OpenAiMaxRetries == 10, "valid boundary values were rejected");
        }
        finally { Directory.Delete(root, true); }
    }

    static void NullableConfiguration()
    {
        using var run = A.BeginRun();
        var root = TempDirectory();
        try
        {
            A.Options.QaModel = "old";
            A.Options.QaReasoningEffort = "high";
            A.Options.QaReasoningEffortExplicit = true;
            var file = Path.Combine(root, "config.json");
            Write(file, "{\"qaModel\":null,\"qaReasoningEffort\":null}");
            ConfigurationService.ApplyConfigFile(file);
            Check(A.Options.QaModel is null && A.Options.QaReasoningEffort is null && !A.Options.QaReasoningEffortExplicit && A.CurrentRun.ConfigurationErrors.Count == 0, "null did not clear inherited configuration");
        }
        finally { Directory.Delete(root, true); }
    }

    sealed class NoDesktop : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multiple) => throw new InvalidOperationException("Replay consulted the live desktop");
        public void CopyPixels(Graphics target, Rectangle source) => throw new InvalidOperationException("Replay captured the desktop");
    }

    static void ReplayGeometry()
    {
        using var run = A.BeginRun();
        var old = DesktopCapture.Backend;
        var map = ScreenCoordinateMapper.Create(-1920, -200, 3000, 1000, 1500, 500);
        A.CurrentRun.CurrentScreenMap = map;
        A.CurrentRun.CurrentUiaTargets = [];
        try
        {
            DesktopCapture.Backend = new NoDesktop();
            var fingerprint = Convert.ToBase64String(new byte[96 * 54]);
            var frame = new LoopReplayFrame { ScreenFingerprintBase64 = fingerprint, PreviousAction = new() { Type = "click_uia", UiaIndex = 0 } };
            _ = LoopReplayService.ReplayCase(new() { Frames = [new() { ScreenFingerprintBase64 = fingerprint }, frame], ScreenWidth = 320, ScreenHeight = 200 });
            Check(ReplayActionResolver.Resolve(frame, 320, 200).ScreenPoint is null, "legacy UIA target geometry was fabricated");
            var snapshot = new ResolvedActionSnapshot(frame.PreviousAction, "private description", "click_uia:0", new Point(-123, 42));
            frame.PreviousActionGeometry = ReplayActionResolver.Record(snapshot);
            var serialized = JsonSerializer.Serialize(frame);
            var restored = JsonSerializer.Deserialize<LoopReplayFrame>(serialized)!;
            Check(ReplayActionResolver.Resolve(restored, 320, 200).ScreenPoint == snapshot.ScreenPoint, "recorded geometry did not round-trip");
            Check(!serialized.Contains("private description") && A.CurrentRun.CurrentScreenMap == map, "replay exposed text or mutated the live map");
            frame.PreviousActionGeometry = null;
            frame.PreviousAction = new() { Type = "click", X = 1, Y = 1 };
            Check(ReplayActionResolver.Resolve(frame, 320, 200).ScreenPoint == new Point(319, 199), "normalized coordinates used live dimensions");
        }
        finally { DesktopCapture.Backend = old; }
    }

    static void ReplayImages()
    {
        using var run = A.BeginRun();
        foreach (var reference in new[] { "(not saved; base64 omitted)", "", "data:image/png;base64,bad!", "data:image/png;base64,dGVzdA==", "file://Z:/missing-image.png" })
        {
            var rejected = false;
            try { ArtifactAnalysisService.HydrateReplayImages(new JsonObject { ["image_url"] = reference }); }
            catch (Exception ex) when (ex is ArgumentException or FormatException or InvalidOperationException or IOException) { rejected = true; }
            Check(rejected, "invalid replay image accepted: " + reference);
        }
        var missingRejected = false;
        try { ArtifactAnalysisService.HydrateReplayImages(new JsonObject { ["type"] = "input_image" }); }
        catch (InvalidOperationException) { missingRejected = true; }
        Check(missingRejected, "image content without a reference was accepted");
        using var bitmap = new Bitmap(4, 4);
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        var url = "data:image/png;base64," + Convert.ToBase64String(stream.ToArray());
        var stats = ArtifactAnalysisService.HydrateReplayImages(new JsonObject { ["image_url"] = url });
        Check(stats.Images == 1 && stats.ImageBytes == stream.Length, "inline image was not validated/counted");
    }

    sealed class FailedInput : Input.IInputBackend
    {
        internal int Sends;
        internal int CursorCalls;
        internal int ClipboardCalls;
        public int LastError => 5;
        public bool SetCursorPosition(int x, int y) { CursorCalls++; return false; }
        public void SetClipboardText(string text) { ClipboardCalls++; throw new InvalidOperationException("Synthetic clipboard failure"); }
        public uint Send(Input.INPUT[] inputs) { Sends++; return (uint)inputs.Length; }
    }

    static void CursorFailure() => InputFailure(new() { Type = "click_uia", UiaIndex = 0 });
    static void ClipboardFailure() => InputFailure(new() { Type = "paste_text", Text = "test" });
    static void InputFailure(ActionDto action)
    {
        using var run = A.BeginRun();
        var old = Input.InputBackend;
        var backend = new FailedInput();
        try
        {
            Input.InputBackend = backend;
            A.Options.MouseEnabled = true;
            A.Options.AllowHighLevelActions = true;
            A.Options.ForceRealUiOnly = false;
            A.CurrentRun.CurrentUiaTargets = [new(0, "synthetic", "button", new Rectangle(10, 10, 20, 20))];
            var failed = false;
            try { Input.ExecuteAction(action); }
            catch (InvalidOperationException) { failed = true; }
            Check(failed && backend.Sends == 0, "input continued after cursor/clipboard failure");
            Check(action.Type == "click_uia" ? backend.CursorCalls == 1 : backend.ClipboardCalls == 1, "test did not reach the failing input operation");
        }
        finally { Input.InputBackend = old; }
    }
}
