using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;

namespace RDPilot.SelfTests.Regression;

internal static class SeventhReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("future recovery payload bypasses current structural validation", () => FutureStore("{\"version\":999,\"lessons\":[null]}")),
        ("future recovery payload bypasses current deserialization", () => FutureStore("{\"Version\":999,\"lessons\":{\"newLayout\":true}}")),
        ("duplicate version cannot hide a future recovery format", () => FutureStore("{\"version\":999,\"Version\":7,\"lessons\":[]}")),
        ("future backup remains protected when primary is corrupt", FutureBackup),
        ("ordinary batches preserve repeated text before Enter", () => RepeatedText(false)),
        ("ordinary batches preserve identical values in different fields", () => RepeatedText(true)),
        ("invalid batch element stops the remaining sequence", InvalidBatchPrefix),
        ("batch focus change is not attributed to failed text input", () => BatchContext(false)),
        ("batch context still allows a delayed final effect", () => BatchContext(true)),
        ("standalone input uses fresh semantic context", StandaloneContext),
        ("observed inputs retain their own semantic baselines", ObservedContexts),
        ("redacted monitoring telemetry preserves productive cycle verdict", () => MonitoringReplay(true)),
        ("redacted telemetry still detects harmful finite cycles", () => MonitoringReplay(false)),
        ("inspection round trips retain semantics after redaction", InspectionReplay),
        ("replay retains raster and placement classification", DomainReplay),
        ("single turn input detects equal-luminance color changes", () => ColorTurn(false)),
        ("observed turn input detects equal-luminance color changes", () => ColorTurn(true)),
        ("color turn identity distinguishes states and recognizes returns", ColorIdentity),
        ("color analysis retains grayscale fallback and pixel masking", ColorFallback)
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class Memory : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string FilePath => Path.Combine(Root, "memory.json");
        internal Memory()
        {
            Directory.CreateDirectory(Root);
            A.Options.RecoveryMemoryEnabled = true;
            A.Options.RecoveryMemoryPath = FilePath;
            // Reset persistence state inherited from an earlier read-only fixture.
            RecoveryMemoryService.LoadRecoveryLessons();
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
    static void Write(string path, string text) => File.WriteAllText(path,
        text.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
    static object? Invoke(string name, params object[] args) => typeof(RecoveryMemoryService)
        .GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, args);

    static void AssertProtected(string primary, string backup)
    {
        var primaryBytes = File.ReadAllBytes(primary);
        var backupBytes = File.ReadAllBytes(backup);
        RecoveryMemoryService.LoadRecoveryLessons();
        Invoke("PersistCalibrationSnapshot");
        Check(!(bool)Invoke("SaveRecoveryLessons", new List<RecoveryLesson>())!, "unsupported format remained writable");
        Check(primaryBytes.SequenceEqual(File.ReadAllBytes(primary)) && backupBytes.SequenceEqual(File.ReadAllBytes(backup)), "future store or backup was overwritten");
    }
    static void FutureStore(string json)
    {
        using var scope = A.BeginRun();
        using var memory = new Memory();
        Write(memory.FilePath, json);
        Write(memory.FilePath + ".bak", "{\"version\":7,\"lessons\":[]}");
        var result = RecoveryStoreReader.Read(memory.FilePath);
        Check(result.Status == RecoveryStoreReadStatus.UnsupportedVersion && result.Version == 999 && !result.UsedBackup, "reader fell through to an older backup");
        AssertProtected(memory.FilePath, memory.FilePath + ".bak");
    }
    static void FutureBackup()
    {
        using var scope = A.BeginRun();
        using var memory = new Memory();
        Write(memory.FilePath, "{broken");
        Write(memory.FilePath + ".bak", "{\"version\":999,\"lessons\":{\"future\":true}}");
        var result = RecoveryStoreReader.Read(memory.FilePath);
        Check(result.Status == RecoveryStoreReadStatus.UnsupportedVersion && result.UsedBackup, "future backup misclassified");
        AssertProtected(memory.FilePath, memory.FilePath + ".bak");
    }

    static void RepeatedText(bool fields)
    {
        using var f = new Fixture();
        ActionDto first = new() { Type = "type_text", Text = "ha" };
        var candidates = new List<ActionDto> { first };
        if (fields) candidates.Add(new() { Type = "keys", Keys = ["tab"] });
        candidates.Add(new() { Type = "type_text", Text = "ha" });
        candidates.Add(new() { Type = "keys", Keys = ["enter"] });
        var turn = f.Turn(first);
        turn.Decision = BatchPlanner.CreateDecision(candidates);
        f.Execute(turn);
        var typed = new string(f.Input.Events.Where(e => e.type == 1 && (e.U.ki.dwFlags & 6) == 4).Select(e => (char)e.U.ki.wScan).ToArray());
        Check(!turn.actionExecutionFailed && typed == "haha", "duplicate text was lost");
        Check(f.Run.recentExecutedActions.Count == candidates.Count && f.Run.prevAction!.Action.Keys![0] == "enter", "batch order changed");
    }
    static void InvalidBatchPrefix()
    {
        using var f = new Fixture();
        var result = BatchPlanner.CreateDecision([
            new() { Type = "type_text", Text = "first" }, new() { Type = "keys", Keys = ["invalid+chord"] },
            new() { Type = "keys", Keys = ["enter"] }]);
        Check(result.FollowUps.Count == 0, "invalid element skipped before a later commit");
    }

    static readonly UiPromptContext Document = new("Browser", "browser", "document", null, null, null);
    static readonly UiPromptContext Address = Document with { FocusedUiaSummary = "address bar" };
    static void BatchContext(bool delayed)
    {
        using var f = new Fixture();
        UiAutomationService.PromptContextProvider = (_, _, _) => f.Input.Sends == 0 ? Document : Address;
        var turn = f.Turn(new() { Type = "keys", Keys = ["ctrl+l"] }, new() { Type = "type_text", Text = "missing" });
        turn.promptContext = Document;
        f.Execute(turn);
        Check(f.Run.prevObservationContext == Address, "final baseline kept pre-batch focus");
        if (delayed) f.DelayedShift = 40;
        var result = f.Run.observationSession.Assess(f.Run.prevObservationFrame!, ScreenshotService.CaptureObservationFrameProbe(true),
            f.Run.prevAction, f.Run.prevObservationContext, Address, "finite");
        Check(result.ActionOutcome == (delayed ? ActionOutcomeState.Confirmed : ActionOutcomeState.NoEffect), "focus change attributed to text input");
    }
    static void StandaloneContext()
    {
        using var f = new Fixture();
        UiAutomationService.PromptContextProvider = (_, _, _) => Address;
        var turn = f.Turn(new() { Type = "type_text", Text = "missing" });
        turn.promptContext = Document;
        f.Execute(turn);
        Check(f.Run.prevObservationContext == Address, "standalone baseline used stale planning context");
    }
    static void ObservedContexts()
    {
        using var f = new Fixture { Moving = true };
        UiAutomationService.PromptContextProvider = (_, _, _) => f.Input.Sends == 0 ? Document : Address;
        var region = new Rectangle(0, 0, 320, 200);
        var actions = new[] { "right", "down" }.Select(key => Input.CaptureResolvedAction(new() { Type = "keys", Keys = [key] }, null) with { ObservationRegion = region }).ToArray();
        var result = ControlBatchService.ExecuteObservedTurnActionSequenceAsync(actions, ScreenshotService.CaptureObservationFrameProbe(true),
            region, f.Run.turnBasedTransitions, 0, f.Run.Cancellation).GetAwaiter().GetResult();
        Check(result.Executions.Count == 2 && result.Executions[0].BeforeContext == Document && result.Executions[1].BeforeContext == Address,
            "observed sequence lost input-specific contexts");
    }

    static void MonitoringReplay(bool recurring)
    {
        using var f = new Fixture();
        using var memory = new Memory();
        var liveDetected = false;
        var graph = new LoopStateGraph();
        var mode = recurring ? "continuous" : "finite";
        for (var i = 0; i < 8; i++)
        {
            var fp = Enumerable.Repeat((byte)(i % 2 * 150), 64).ToArray();
            ResolvedActionSnapshot? previous = null;
            if (i > 0)
            {
                previous = Input.CaptureResolvedAction(new() { Type = "keys", Keys = [i % 2 == 0 ? "left" : "right"], Note = "check status PRIVATE-CUSTOMER-NOTE" }, null);
                f.Run.recentExecutedActions.Enqueue(previous);
            }
            var live = RecoveryMemoryService.AssessVisualStateCycle(graph, fp, fp, f.Context, i + 1, f.Run.recentExecutedActions, previous, .5,
                recordLearning: false, goalMode: mode, recurringWorkflowIntent: recurring);
            liveDetected |= live.IsLoop;
            RecoveryMemoryService.AppendLoopReplayObservation("monitoring", i + 1, 320, 200, fp, fp, f.Context, previous, .5, live, null,
                mode, recurring, 0, f.Run.recentExecutedActions);
        }
        var telemetry = File.ReadAllText(Path.Combine(memory.Root, "loop-telemetry.jsonl"));
        Check(!telemetry.Contains("PRIVATE-CUSTOMER", StringComparison.OrdinalIgnoreCase), "private note leaked into semantic metadata");
        var frames = telemetry.Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries)
            .Select(line => JsonSerializer.Deserialize<LoopTelemetryReplayEnvelope>(line, A.PrettyJson)!.ReplayFrame!).ToList();
        var replay = LoopReplayService.ReplayCase(new() { Name = "monitoring", ScreenWidth = 320, ScreenHeight = 200, GoalMode = mode, RecurringWorkflowIntent = recurring, Frames = frames });
        Check(replay.Detected == liveDetected && liveDetected == !recurring, "redaction changed productive-cycle verdict");
    }
    static ResolvedActionSnapshot Redact(ResolvedActionSnapshot action)
    {
        var geometry = JsonSerializer.Deserialize<ReplayActionGeometry>(JsonSerializer.Serialize(ReplayActionResolver.Record(action), A.PrettyJson), A.PrettyJson);
        var dto = new ActionDto { Type = action.Action.Type, Keys = action.Action.Keys is null ? null : ["<text-key>"], Note = "opaque-hash" };
        return ReplayActionResolver.Resolve(dto, geometry, 320, 200);
    }
    static void InspectionReplay()
    {
        using var f = new Fixture();
        var live = new[] {
            Input.CaptureResolvedAction(new() { Type = "click", XPx = 20, YPx = 20, Note = "open help" }, null),
            Input.CaptureResolvedAction(new() { Type = "wait", WaitSeconds = 1 }, null),
            Input.CaptureResolvedAction(new() { Type = "click", XPx = 40, YPx = 20, Note = "close details" }, null)
        };
        var replay = live.Select(Redact).ToArray();
        Check(RecoveryMemoryService.IsCompletedInspectionRoundTrip(live, live[^1]) && RecoveryMemoryService.IsCompletedInspectionRoundTrip(replay, replay[^1]), "inspection semantics lost");
    }
    static void DomainReplay()
    {
        using var f = new Fixture();
        var context = new UiPromptContext("PRIVATE Canvas", "chrome", "document", null, null, null);
        var recordedContext = new UiPromptContext("opaque", "chrome", "opaque", null, null, null)
        { SemanticFeatures = JsonSerializer.Deserialize<LoopContextFeatures>(JsonSerializer.Serialize(LoopSemanticClassifier.Context(context), A.PrettyJson), A.PrettyJson) };
        foreach (var note in new[] { "inspect", "place tile" })
        {
            var action = Input.CaptureResolvedAction(new() { Type = "click", XPx = 50, YPx = 50, Note = note }, null);
            var redacted = Redact(action);
            var fp = new byte[64];
            var live = RecoveryMemoryService.AssessVisualStateCycle(new(), fp, fp, context, 1, [action], action, .5, recordLearning: false);
            var replay = RecoveryMemoryService.AssessVisualStateCycle(new(), fp, fp, recordedContext, 1, [redacted], redacted, .5, recordLearning: false);
            Check(live.InteractionDomain == replay.InteractionDomain && live.InteractionDomain == (note == "inspect" ? "raster-canvas" : "placement"), "replay domain differs");
        }
    }

    static ScreenObservationFrame ColorFrame(bool green)
    {
        using var bitmap = new Bitmap(100, 100);
        using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(green ? Color.FromArgb(0, 130, 0) : Color.Red);
        var detail = ImageComparisonService.BuildDetailFingerprints(bitmap, 100, 100);
        var global = ImageComparisonService.BuildImageFingerprint(bitmap);
        return new(global, global, detail.Grayscale, 100, 100, new(0, 0, 100, 100)) { DetailColorFingerprint = detail.Color };
    }
    static void ColorTurn(bool batch)
    {
        using var f = new Fixture(turnBased: true);
        var before = ColorFrame(false); var after = ColorFrame(true); var region = before.ScreenBounds;
        Check(before.DetailFingerprint.SequenceEqual(after.DetailFingerprint), "colors are not equal-luminance");
        var action = Input.CaptureResolvedAction(new() { Type = "click", XPx = 50, YPx = 50 }, null) with { ObservationRegion = region };
        var observation = f.Run.observationSession.Assess(before, after, action, f.Context, f.Context, "finite");
        var tracker = f.Run.turnBasedTransitions;
        tracker.ObserveState(before, region);
        var original = tracker.CurrentStateId;
        if (batch)
        {
            tracker.BeginBatch(before, region);
            var result = tracker.RecordBatchActionStep(before, after, region, action, true);
            Check(!result.NoEffect, "batch rejected a persistent color change");
        }
        else Check(!tracker.RecordTransition(before, after, region, action, observation), "tracker overrode confirmed color change");
        Check(observation.ActionOutcome == ActionOutcomeState.Confirmed && tracker.CurrentStateId != original, "color state was merged with the old state");
    }
    static void ColorIdentity()
    {
        using var f = new Fixture(turnBased: true);
        var red = ColorFrame(false); var green = ColorFrame(true); var region = red.ScreenBounds;
        var action = Input.CaptureResolvedAction(new() { Type = "click", XPx = 50, YPx = 50 }, null) with { ObservationRegion = region };
        var tracker = f.Run.turnBasedTransitions;
        tracker.ObserveState(red, region); var redId = tracker.CurrentStateId;
        tracker.RecordTransition(red, green, region, action, f.Run.observationSession.Assess(red, green, action, f.Context, f.Context, "finite"));
        Check(tracker.CurrentStateId != redId, "distinct colors merged");
        tracker.RecordTransition(green, red, region, action, f.Run.observationSession.Assess(green, red, action, f.Context, f.Context, "finite"));
        Check(tracker.CurrentStateId == redId, "return to known color state missed");
    }
    static void ColorFallback()
    {
        var frame = ColorFrame(false);
        var fallback = frame with { DetailColorFingerprint = [] };
        var first = TurnVisualAnalysis.Extract(fallback, frame.ScreenBounds);
        var second = TurnVisualAnalysis.Extract(fallback, frame.ScreenBounds);
        var mask = new bool[TurnVisualAnalysis.Side * TurnVisualAnalysis.Side];
        Check(TurnVisualAnalysis.Compare(first, second, mask) == (0.0, 0.0), "grayscale fallback is unstable");
        second[3 * 100] = 240;
        Check(TurnVisualAnalysis.Compare(first, second, mask).ChangedRatio > 0, "channel change missed");
        mask[100] = true;
        Check(TurnVisualAnalysis.Compare(first, second, mask) == (0.0, 0.0), "mask treated channels as spatial pixels");
    }
}