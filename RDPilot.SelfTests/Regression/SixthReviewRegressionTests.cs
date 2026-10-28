using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;

namespace RDPilot.SelfTests.Regression;

internal static class SixthReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("ordinary batch measures only the final input effect", () => OrdinaryBatch(false)),
        ("ordinary batch still detects a delayed text effect", () => OrdinaryBatch(true)),
        ("corrupt recovery remains unchanged after calibration and lesson saves", CorruptWrites),
        ("missing recovery is distinguishable and can be initialized", MissingStore),
        ("locked recovery is unreadable rather than missing", LockedStore),
        ("overflowing lesson writer counters use the backup", () => NumericBackup("lesson")),
        ("overflowing compacted counters use the backup", () => NumericBackup("compacted")),
        ("overflowing calibration counters use the backup", () => NumericBackup("calibration")),
        ("negative counters use the backup", () => NumericBackup("negative")),
        ("non-finite accumulated rewards use the backup", () => NumericBackup("reward")),
        ("future recovery versions remain read-only after normalization", FutureStore),
        ("601-point gestures respect interpolation bounds", () => Gesture(601)),
        ("700-point gestures respect interpolation bounds", () => Gesture(700)),
        ("1024-point gestures respect interpolation bounds", () => Gesture(1024)),
        ("non-ASCII held keys are rejected before any input", UnicodeHold),
        ("physical held-key aliases normalize and release once", HeldAliases),
        ("batched execution history survives telemetry and replay", () => ReplayHistory(false, false)),
        ("partial execution history survives absent previous action", () => ReplayHistory(true, false)),
        ("execution history is cleared at external state epochs", () => ReplayHistory(false, true)),
        ("version two replay rejects missing execution history", () => InvalidReplay("missing")),
        ("version two replay rejects invalid previous action index", () => InvalidReplay("index")),
        ("version two replay rejects null execution records", () => InvalidReplay("null")),
        ("turn state pruning also removes navigation positions and edges", StateRetention),
        ("turn state reset clears all associated navigation data", StateReset)
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    static void OrdinaryBatch(bool delayed)
    {
        using var f = new Fixture { Moving = true };
        var turn = f.Turn(new() { Type = "keys", Keys = ["ctrl", "a"] }, new() { Type = "type_text", Text = "missing" });
        f.Execute(turn);
        Check(!turn.actionExecutionFailed && f.Run.recentExecutedActions.Count == 2, "successful prefix accounting changed");
        if (delayed) f.DelayedShift = 40;
        var latest = ScreenshotService.CaptureObservationFrameProbe(true);
        var final = f.Run.prevAction!;
        Check(final.Action.Type == "type_text", "wrong action carried forward");
        Check(ImageComparisonService.ComputeImageDelta(turn.observationFrame.GlobalFingerprint, latest.GlobalFingerprint) > .005, "fixture did not change after the first input");
        var assessment = f.Run.observationSession.Assess(f.Run.prevObservationFrame!, latest, final, f.Context, f.Context, "finite");
        Check(assessment.ActionOutcome == (delayed ? ActionOutcomeState.Confirmed : ActionOutcomeState.NoEffect) && assessment.GoalProgress == (delayed ? GoalProgressState.Progress : GoalProgressState.NoProgress),
            "first input effect attributed to final input");
    }

    sealed class Memory : IDisposable
    {
        readonly IDisposable scope = A.BeginRun();
        internal readonly string Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string Path => System.IO.Path.Combine(Root, "memory.json");
        internal Memory()
        {
            Directory.CreateDirectory(Root);
            A.Options.RecoveryMemoryEnabled = true;
            A.Options.RecoveryMemoryPath = Path;
        }
        public void Dispose() { scope.Dispose(); Directory.Delete(Root, true); }
    }
    static void Write(string path, string contents) => File.WriteAllText(path,
        contents.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
    static object? Invoke(string method, params object[] args) => typeof(RecoveryMemoryService)
        .GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, args);

    static void CorruptWrites()
    {
        using var m = new Memory();
        const string original = "{\"lessons\":[{\"id\":\"salvage\",\"winningStrategy\":\"preserve\"},null]}";
        Write(m.Path, original); Write(m.Path + ".bak", original);
        var primary = File.ReadAllBytes(m.Path); var backup = File.ReadAllBytes(m.Path + ".bak");
        Check(RecoveryMemoryService.LoadRecoveryLessons().Count == 0, "corrupt store accepted");
        for (var i = 0; i < 2; i++) Invoke("PersistCalibrationSnapshot");
        Check(Invoke("SaveRecoveryLessons", new List<RecoveryLesson> { new() { Id = "new" } }) is false, "lesson save should be blocked");
        Check(File.ReadAllBytes(m.Path).SequenceEqual(primary) && File.ReadAllBytes(m.Path + ".bak").SequenceEqual(backup), "original data overwritten");
        // Explicitly loading a repaired store re-enables persistence.
        Write(m.Path, "{\"lessons\":[]}");
        RecoveryMemoryService.LoadRecoveryLessons();
        Check(Invoke("SaveRecoveryLessons", new List<RecoveryLesson> { new() { Id = "new" } }) is true, "repaired file remains locked read-only");
    }

    static void MissingStore()
    {
        using var m = new Memory();
        Check(RecoveryStoreReader.Read(m.Path).Status == RecoveryStoreReadStatus.Missing, "wrong missing status");
        RecoveryMemoryService.LoadRecoveryLessons();
        Check(RecoveryStoreReader.Read(m.Path).Status == RecoveryStoreReadStatus.Loaded, "new store was not created");
    }
    static void LockedStore()
    {
        using var m = new Memory();
        Write(m.Path, "{\"lessons\":[]}");
        using var locked = new FileStream(m.Path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Check(RecoveryStoreReader.Read(m.Path).Status == RecoveryStoreReadStatus.Unreadable, "locked file treated as absent");
    }
    static void NumericBackup(string kind)
    {
        using var m = new Memory();
        var corrupt = kind switch
        {
            "lesson" => "{\"lessons\":[{\"id\":\"bad\",\"successByWriter\":{\"a\":2147483647,\"b\":1}}]}",
            "compacted" => "{\"lessons\":[{\"id\":\"bad\",\"compactedFailureCount\":2147483647,\"failureByWriter\":{\"a\":1}}]}",
            "calibration" => "{\"calibration\":{\"cycle\":{\"confirmedByWriter\":{\"a\":2147483647,\"b\":1}}}}",
            "negative" => "{\"lessons\":[{\"id\":\"bad\",\"selectionCount\":-1}]}",
            _ => "{\"lessons\":[{\"id\":\"bad\",\"rewardByWriter\":{\"a\":1e308,\"b\":1e308}}]}"
        };
        Write(m.Path, corrupt); Write(m.Path + ".bak", "{\"version\":1,\"lessons\":[{\"id\":\"saved\",\"winningActionTypes\":[\"keys\"]}]}");
        var backup = File.ReadAllBytes(m.Path + ".bak");
        var candidate = RecoveryStoreReader.Read(m.Path);
        Check(candidate.Status == RecoveryStoreReadStatus.Loaded && candidate.UsedBackup && candidate.Store!.Lessons[0].StrategySteps.Count == 1, "backup not normalized and selected");
        var lessons = RecoveryMemoryService.LoadRecoveryLessons();
        Check(lessons.Single().Id == "saved" && File.ReadAllBytes(m.Path + ".bak").SequenceEqual(backup), "backup restore damaged data");
    }
    static void FutureStore()
    {
        using var m = new Memory();
        Write(m.Path, "{\"version\":999,\"lessons\":[],\"future\":\"preserve\"}");
        var original = File.ReadAllBytes(m.Path);
        var result = RecoveryStoreReader.Read(m.Path);
        Check(result.Status == RecoveryStoreReadStatus.UnsupportedVersion && result.Version == 999 && result.Store is null, "future version was interpreted as the current schema");
        RecoveryMemoryService.LoadRecoveryLessons();
        Invoke("PersistCalibrationSnapshot");
        Check(Invoke("SaveRecoveryLessons", new List<RecoveryLesson>()) is false && File.ReadAllBytes(m.Path).SequenceEqual(original), "future schema overwritten");
    }

    static void Gesture(int count)
    {
        using var f = new Fixture();
        A.Options.MaxGesturePathPoints = 1024;
        var action = new ActionDto { Type = "drag_path", DurationMs = 1000,
            Path = Enumerable.Range(0, count).Select(i => new GesturePointDto { XPx = 10 + i % 250, YPx = 10 + i / 250 }).ToArray() };
        Check(Input.CaptureResolvedAction(action, null).IsValid, "valid long path rejected");
        var path = Input.ResolveGesturePath(action);
        var plan = Input.BuildMouseDragPlan(path, 1000);
        Check(plan.Moves.Length >= path.Count - 1 && plan.Moves.Last() == path.Last(), "gesture lost segments or endpoint");
        Check(Math.Abs(Input.GestureMoveTargetElapsedMs(plan, plan.Moves.Length - 1) - 1000) < .01, "duration changed");
    }
    static void UnicodeHold()
    {
        using var f = new Fixture();
        foreach (var key in new[] { "Ą", "é", "Ж", "١" })
        {
            var action = new ActionDto { Type = "hold_keys", Keys = ["ctrl", key], DurationMs = 100 };
            Check(!Input.CaptureResolvedAction(action, null).IsValid, "Unicode hold was accepted");
            var rejected = false;
            try { Input.ExecuteAction(action); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && f.Input.Events.Count == 0, "unsupported hold emitted native input");
        }
    }
    static void HeldAliases()
    {
        using var f = new Fixture();
        var action = new ActionDto { Type = "hold_keys", Keys = ["control", "ctrl", " "], DurationMs = 100 };
        Check(Input.CaptureResolvedAction(action, null).IsValid, "physical space alias rejected");
        Check(Input.CaptureResolvedAction(action, null).IneffectiveSignature == Input.CaptureResolvedAction(new() { Type = "hold_keys", Keys = ["ctrl", "space"], DurationMs = 100 }, null).IneffectiveSignature, "held-key identity disagrees with execution");
        Input.ExecuteAction(action);
        Check(f.Input.Events.Select(item => item.U.ki.wVk).SequenceEqual(new ushort[] { 0x11, 0x20, 0x20, 0x11 }), "aliases produced duplicate modifier presses");
    }

    static void ReplayHistory(bool failedPrefix, bool epochs)
    {
        using var f = new Fixture();
        var root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            A.Options.RecoveryMemoryEnabled = true; A.Options.RecoveryMemoryPath = Path.Combine(root, "memory.json");
            var right = Input.CaptureResolvedAction(new() { Type = "keys", Keys = ["right"] }, null);
            var down = Input.CaptureResolvedAction(new() { Type = "keys", Keys = ["down"] }, null);
            var liveDetected = false;
            for (var i = 0; i < 8; i++)
            {
                if (epochs && i == 2) f.Run.BeginExternalStateEpoch();
                var fp = Enumerable.Repeat((byte)(i == 2 ? 0 : i * 30), 64).ToArray();
                if (i > 0)
                {
                    ControlActionOutcome.RecordExecuted(f.Run, i + ".1", right, true);
                    ControlActionOutcome.RecordExecuted(f.Run, i + ".2", down, true);
                }
                var previous = i == 0 || failedPrefix && i == 2 ? null : down;
                var live = RecoveryMemoryService.AssessVisualStateCycle(f.Run.loopStateGraph, fp, fp, f.Context, i + 1,
                    f.Run.recentExecutedActions, previous, .5, recordLearning: false);
                liveDetected |= live.IsLoop;
                RecoveryMemoryService.AppendLoopReplayObservation("review", i + 1, 320, 200, fp, fp, f.Context, previous, .5, live, null,
                    "finite", false, f.Run.stateEpoch, f.Run.recentExecutedActions);
            }
            var frames = File.ReadAllLines(Path.Combine(root, "loop-telemetry.jsonl"))
                .Select(line => JsonSerializer.Deserialize<LoopTelemetryReplayEnvelope>(line, A.PrettyJson)!.ReplayFrame!).ToList();
            Check(frames.All(frame => frame.ReplayStateVersion == 2 && frame.RecentActions is not null), "missing execution history");
            Check(frames[2].RecentActions!.Count == (epochs ? 2 : 4), "prefix lost or duplicated");
            var replay = new LoopReplayCase { Name = "batch-prefix", Frames = frames, ScreenWidth = 320, ScreenHeight = 200 };
            var actual = LoopReplayService.ReplayCase(replay);
            Check(actual.Detected == liveDetected && liveDetected == !epochs, "live/replay verdict differs");
            if (!epochs) Check(actual.Step == 3, "replay detected the cycle at a different step");
        }
        finally { Directory.Delete(root, true); }
    }
    static void InvalidReplay(string kind)
    {
        var frame = new LoopReplayFrame { ReplayStateVersion = 2, ScreenFingerprintBase64 = Convert.ToBase64String(new byte[64]), RecentActions = [] };
        if (kind == "missing") frame.RecentActions = null;
        else if (kind == "null") frame.RecentActions!.Add(null!);
        else { frame.PreviousAction = new() { Type = "wait" }; frame.PreviousActionIndex = 0; }
        var rejected = false;
        try { LoopReplayService.ReplayCase(new() { Frames = [frame, frame] }); } catch (InvalidOperationException) { rejected = true; }
        Check(rejected, "malformed history accepted");
    }
    static void StateRetention()
    {
        var store = new TurnStateStore();
        store.Add(new("S0", [0]), null); store.Anchor("S0");
        for (var i = 1; i < 1800; i++)
        {
            store.Add(new($"S{i}", [(byte)i]), $"S{i - 1}");
            store.RecordTransition($"S{i - 1}", "ArrowRight", $"S{i}", "changed", $"S{i}");
        }
        var ids = store.States.Select(state => state.Id).ToHashSet();
        Check(store.States.Count == TurnStateStore.StateLimit && store.Positions.Count <= TurnStateStore.StateLimit, "navigation state grew beyond the prototype bound");
        Check(store.Edges.Count <= TurnStateStore.EdgeLimit && store.Edges.All(edge => ids.Contains(edge.FromState) && ids.Contains(edge.ToState)), "dangling or unbounded navigation edges");
        Check(store.Positions.Keys.All(ids.Contains) && store.Positions["S1799"].X == 1799, "current pose lost during retention");
    }
    static void StateReset()
    {
        var store = new TurnStateStore();
        store.Add(new("one", [0]), null); store.Add(new("two", [1]), "one");
        store.RecordTransition("one", "ArrowRight", "two", "changed", "two");
        store.Clear();
        Check(store.States.Count == 0 && store.Positions.Count == 0 && store.Edges.Count == 0, "reset left stale navigation data");
    }
}
