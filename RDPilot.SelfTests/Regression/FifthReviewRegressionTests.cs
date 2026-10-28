using System.Drawing;
using System.IO;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;
using Inputs = RDPilot.SelfTests.Support.SyntheticControlFixture.Inputs;

namespace RDPilot.SelfTests.Regression;

internal static class FifthReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("key aliases produce identical native chords through ExecuteAction", KeyAliases),
        ("literal space survives the production key sequence parser", Spaces),
        ("key validation rejects the entire malformed sequence before sending input", InvalidSequence),
        ("key signatures distinguish sequences from chords and preserve Unicode", KeyIdentity),
        ("text cooldown blocks a batch follow-up", () => BatchPolicy("text")),
        ("repeat cooldown blocks a batch follow-up", () => BatchPolicy("repeat")),
        ("expired text cooldown permits a batch follow-up", () => BatchPolicy("expired")),
        ("paste shortcut aliases obey the shared text cooldown", PastePolicy),
        ("observed batch obeys cooldown before its second input", ObservedCooldown),
        ("observed batch keeps the final input baseline and no-effect verdict", () => ObservedOutcome(false)),
        ("delayed reaction supersedes a previous batch no-effect verdict", () => ObservedOutcome(true)),
        ("mouse buttons remain distinct in exact and spatial cooldowns", MouseIdentity),
        ("null recovery lessons fall back to the backup", () => Memory("{\"lessons\":[null]}", true)),
        ("null recovery strategy steps fall back to the backup", () => Memory("{\"lessons\":[{\"id\":\"bad\",\"strategySteps\":[null]}]}", true)),
        ("null calibration buckets fall back to the backup", () => Memory("{\"calibration\":{\"cycle\":null}}", true)),
        ("duplicate calibration names fall back to the backup", () => Memory("{\"calibration\":{\"cycle\":{},\"CYCLE\":{}}}", true)),
        ("invalid recovery without a backup returns an empty store safely", () => Memory("{\"lessons\":[null]}", false)),
        ("legacy recovery collections still migrate", LegacyMemory),
        ("replay preserves external state boundaries through serialization", () => ReplayEpochs(true)),
        ("legacy replay retains uninterrupted loop detection", () => ReplayEpochs(false)),
        ("replay rejects unsupported state metadata", ReplayVersion),
        ("telemetry export preserves epoch boundaries", TelemetryEpochs)
    ];

    static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
    }

    static (ushort Vk, ushort Scan, uint Flags)[] Events(Inputs input) =>
        input.Events.Select(e => (e.U.ki.wVk, e.U.ki.wScan, e.U.ki.dwFlags)).ToArray();

    static void KeyAliases()
    {
        using var f = new Fixture();
        foreach (var pair in new[] { ("ctrl", "control"), ("alt", "option"), ("win", "windows"), ("win", "super"), ("win", "meta"), ("win", "cmd") })
        {
            f.Input.Events.Clear();
            Input.ExecuteAction(new() { Type = "keys", Keys = [pair.Item1, "a"] });
            var expected = Events(f.Input);
            foreach (var keys in new[] { new[] { pair.Item2, "a" }, new[] { pair.Item2 + "+a" } })
            {
                f.Input.Events.Clear();
                var action = new ActionDto { Type = "keys", Keys = keys };
                Check(Input.CaptureResolvedAction(action, null).IsValid, "alias rejected");
                Input.ExecuteAction(action);
                Check(Events(f.Input).SequenceEqual(expected), "different native events for " + pair.Item2);
                Check(expected.Length == 4 && expected[0].Flags == 0 && expected[3].Flags == 2, "modifier was not held across main key");
            }
        }
    }

    static void Spaces()
    {
        using var f = new Fixture();
        foreach (var key in new[] { " ", "space", "spacebar" })
        {
            f.Input.Events.Clear();
            Input.ExecuteAction(new() { Type = "keys", Keys = [key, "tab", key] });
            Check(Events(f.Input).SequenceEqual(new (ushort, ushort, uint)[] { (32, 0, 0), (32, 0, 2), (9, 0, 0), (9, 0, 2), (32, 0, 0), (32, 0, 2) }), "space alias lost in sequence");
        }
    }

    static void InvalidSequence()
    {
        using var f = new Fixture();
        var action = new ActionDto { Type = "keys", Keys = ["a", "unsupported-key"] };
        Check(!Input.CaptureResolvedAction(action, null).IsValid, "invalid key accepted");
        try { Input.ExecuteAction(action); throw new Exception("invalid sequence executed"); }
        catch (ArgumentException) { }
        Check(f.Input.Events.Count == 0, "first key escaped before full validation");
    }

    static void KeyIdentity()
    {
        using var f = new Fixture();
        string Sig(params string[] keys) => Input.ActionSignature(new() { Type = "keys", Keys = keys });
        Check(Sig("control", "a") == Sig("ctrl+a"), "chord aliases have different identities");
        Check(Sig(" ") == Sig("spacebar"), "space aliases have different identities");
        Check(Sig("a", "b") != Sig("alt+b"), "sequence/chord identity collision");
        Check(BatchPlanner.IsTextEntryPrelude(new() { Type = "keys", Keys = ["control", "a"] }), "batch planner uses a different grammar");
        Input.ExecuteAction(new() { Type = "keys", Keys = ["Ą"] });
        Check(f.Input.Events[0].U.ki.wScan == 'Ą', "Unicode character case changed");
    }

    static void BatchPolicy(string kind)
    {
        using var f = new Fixture();
        var text = new ActionDto { Type = "type_text", Text = "BLOCKED" };
        f.Run.textInputCooldownUntilStep = kind == "expired" ? 0 : kind == "text" ? 10 : 0;
        if (kind == "repeat") ControlLoopGuards.RegisterActionCooldown(Input.CaptureResolvedAction(text, null), 10, f.Run.actionCooldownUntilStep, f.Run.spatialActionCooldowns);
        var turn = f.Turn(new() { Type = "keys", Keys = ["control", "a"] }, text);
        Check(turn.decision.FollowUps.Count == 1, "fixture did not queue text");
        f.Execute(turn);
        var textSent = f.Input.Events.Any(e => (e.U.ki.dwFlags & 4) != 0);
        Check(textSent == (kind == "expired"), "follow-up bypassed cooldown or valid text was blocked");
        Check(f.Run.recentExecutedActions.Count == (kind == "expired" ? 2 : 1), "prefix accounting incorrect");
        Check(turn.actionExecutionFailed == (kind != "expired"), "incorrect batch outcome");
    }

    static void PastePolicy()
    {
        using var f = new Fixture();
        f.Run.textInputCooldownUntilStep = 10;
        foreach (var keys in new[] { new[] { "control", "v" }, new[] { "control+shift+v" }, new[] { "shift+ins" } })
        {
            var turn = f.Turn(new() { Type = "keys", Keys = keys });
            f.Execute(turn);
            Check(turn.actionExecutionFailed && f.Input.Events.Count == 0, "paste alias bypassed cooldown");
        }
    }

    static void ObservedCooldown()
    {
        using var f = new Fixture(turnBased: true) { Moving = true };
        A.Options.ObservationMode = "turn_based_interaction";
        f.Run.turnBasedInteractionRect = new(0, 0, 320, 200);
        var second = new ActionDto { Type = "keys", Keys = ["down"] };
        ControlLoopGuards.RegisterActionCooldown(Input.CaptureResolvedAction(second, null), 10, f.Run.actionCooldownUntilStep, f.Run.spatialActionCooldowns);
        var turn = f.Turn(new() { Type = "keys", Keys = ["right"], PlannedInputs = ["right", "down"], PlanConfidence = .99 }, second, 4);
        f.Execute(turn);
        Check(f.Input.Sends == 1 && f.Run.recentExecutedActions.Count == 1 && turn.actionExecutionFailed, "observed follow-up bypassed cooldown");
    }

    static void ObservedOutcome(bool delayed)
    {
        using var f = new Fixture(turnBased: true) { Moving = true };
        A.Options.ObservationMode = "turn_based_interaction";
        f.Run.turnBasedInteractionRect = new(0, 0, 320, 200);
        var turn = f.Turn(new() { Type = "keys", Keys = ["right"], PlannedInputs = ["right", "down"], PlanConfidence = .99 }, new() { Type = "keys", Keys = ["down"] }, 4);
        f.Execute(turn);
        var result = f.Run.prevObservedAction!;
        Check(result is not null && result.Action.Action.Keys!.Single() == "down" && result.Assessment.NoEffect, "last input verdict lost");
        Check(ReferenceEquals(f.Run.prevObservationFrame, result!.Before), "whole-batch baseline retained");
        Check(ImageComparisonService.ComputeImageDelta(turn.observationFrame.GlobalFingerprint, result.Before.GlobalFingerprint) > .005, "fixture has no first-input effect");
        if (delayed) f.DelayedShift = 40;
        var latest = ScreenshotService.CaptureObservationFrameProbe(true);
        var assessment = f.Run.observationSession.Assess(f.Run.prevObservationFrame!, latest, f.Run.prevAction, f.Context, f.Context, "finite");
        assessment = result.Reconcile(assessment, latest);
        Check(assessment.ActionOutcome == (delayed ? ActionOutcomeState.Confirmed : ActionOutcomeState.NoEffect), "wrong final input effect");
        if (!delayed) Check(assessment.GoalProgress == GoalProgressState.NoProgress, "no-effect batch did not contribute stagnation");
    }

    static void MouseIdentity()
    {
        using var f = new Fixture();
        var left = Input.CaptureResolvedAction(new() { Type = "click", XPx = 100, YPx = 100 }, null);
        var right = Input.CaptureResolvedAction(new() { Type = "click", XPx = 100, YPx = 100, Button = "right" }, null);
        Check(left.IneffectiveSignature != right.IneffectiveSignature, "mouse buttons collapsed");
        foreach (var spatial in new[] { false, true })
        {
            f.Run.actionCooldownUntilStep.Clear(); f.Run.spatialActionCooldowns.Clear();
            ControlLoopGuards.RegisterActionCooldown(left, 10, f.Run.actionCooldownUntilStep, f.Run.spatialActionCooldowns, spatial);
            Check(!ControlLoopGuards.IsActionOnCooldown(right, 2, f.Run.actionCooldownUntilStep, f.Run.spatialActionCooldowns, out _), "right click blocked by left click");
            Check(ControlLoopGuards.IsActionOnCooldown(left, 2, f.Run.actionCooldownUntilStep, f.Run.spatialActionCooldowns, out _), "left click cooldown disappeared");
        }
        var history = new List<ResolvedActionSnapshot> { left };
        Check(ControlLoopGuards.UpdateRepeatDetection(right, true, 0, null, history).RepeatCount == 0, "different mouse button counted as repeat");
        Check(!RecoveryMemoryService.HasEmergingLoopPattern(new[] { left, right }, right), "different buttons form a false repeat");
        var legacy = ReplayActionResolver.Resolve(new() { PreviousAction = right.Action }, 320, 200);
        var recorded = ReplayActionResolver.Resolve(new() { PreviousAction = right.Action, PreviousActionGeometry = ReplayActionResolver.Record(right) }, 320, 200);
        Check(legacy.IneffectiveSignature == recorded.IneffectiveSignature, "replay button identity differs");
    }

    static string TempRoot()
    {
        var root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }
    static void Write(string path, string content) => File.WriteAllText(path, content.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));

    static void Memory(string corrupt, bool backup)
    {
        using var scope = A.BeginRun();
        var root = TempRoot();
        try
        {
            A.Options.RecoveryMemoryEnabled = true;
            var path = A.Options.RecoveryMemoryPath = Path.Combine(root, "memory.json");
            Write(path, corrupt);
            if (backup) Write(path + ".bak", "{\"version\":7,\"lessons\":[{\"id\":\"saved\",\"winningStrategy\":\"recover me\"}]}");
            var backupBytes = backup ? File.ReadAllBytes(path + ".bak") : [];
            var lessons = RecoveryMemoryService.LoadRecoveryLessons();
            Check(lessons.Count == (backup ? 1 : 0), "wrong recovery load result");
            if (backup)
            {
                Check(lessons.Single().Id == "saved" && lessons.Single().WinningStrategy == "recover me", "backup lesson lost");
                Check(RecoveryStoreReader.Read(path, out var usedBackup) is not null && !usedBackup, "primary not restored");
                Check(File.ReadAllBytes(path + ".bak").SequenceEqual(backupBytes), "verified backup overwritten by corrupt primary");
            }
            else Check(File.ReadAllText(path) == corrupt, "unrecoverable file silently overwritten");
        }
        finally { Directory.Delete(root, true); }
    }

    static void LegacyMemory()
    {
        using var scope = A.BeginRun();
        var root = TempRoot();
        try
        {
            A.Options.RecoveryMemoryEnabled = true;
            A.Options.RecoveryMemoryPath = Path.Combine(root, "memory.json");
            Write(A.Options.RecoveryMemoryPath, "{\"version\":1,\"lessons\":null,\"calibration\":null}");
            Check(RecoveryMemoryService.LoadRecoveryLessons().Count == 0, "legacy null collections rejected");
        }
        finally { Directory.Delete(root, true); }
    }

    static LoopReplayCase EpochCase(bool epochs)
    {
        var replay = new LoopReplayCase { Name = "epochs", ScreenWidth = 320, ScreenHeight = 200 };
        var run = new ControlRunState("synthetic", "Move", AppContext.BaseDirectory, AppContext.BaseDirectory);
        for (var i = 0; i < 12; i++)
        {
            if (i > 0 && i % 4 == 0) run.BeginExternalStateEpoch();
            var fp = Enumerable.Repeat((byte)(i % 2 == 0 ? 0 : 255), 96 * 54).ToArray();
            ActionDto? action = i % 4 == 0 ? null : new() { Type = "keys", Keys = [i % 2 == 0 ? "left" : "right"] };
            var snapshot = action is null ? null : Input.CaptureResolvedAction(action, null);
            if (snapshot is not null) run.recentExecutedActions.Enqueue(snapshot);
            var live = RecoveryMemoryService.AssessVisualStateCycle(run.loopStateGraph, fp, fp, new("Game", "game", null, null, null, null), i + 1, run.recentExecutedActions, snapshot, .5, recordLearning: false, goalProgress: GoalProgressState.Neutral);
            Check(!live.IsLoop, "live epoch fixture unexpectedly loops");
            replay.Frames.Add(new() { ScreenFingerprintBase64 = Convert.ToBase64String(fp), ActiveWindowFingerprintBase64 = Convert.ToBase64String(fp), PreviousAction = action, PreviousActionGeometry = ReplayActionResolver.Record(snapshot), ActiveProcess = "game", WindowTitle = "Game", GoalProgress = "Neutral", LastDelta = .5, ReplayStateVersion = epochs ? 1 : 0, StateEpoch = epochs ? run.stateEpoch : null });
        }
        return replay;
    }

    static void ReplayEpochs(bool epochs)
    {
        using var f = new Fixture();
        var replay = JsonSerializer.Deserialize<LoopReplayCase>(JsonSerializer.Serialize(EpochCase(epochs)))!;
        Check(LoopReplayService.ReplayCase(replay).Detected == !epochs, "replay epoch compatibility mismatch");
    }

    static void ReplayVersion()
    {
        using var f = new Fixture();
        var replay = EpochCase(true);
        replay.Frames[1].ReplayStateVersion = 99;
        try { LoopReplayService.ReplayCase(replay); throw new Exception("future state version accepted"); }
        catch (InvalidOperationException) { }
    }

    static void TelemetryEpochs()
    {
        using var f = new Fixture();
        var root = TempRoot();
        try
        {
            A.Options.RecoveryMemoryEnabled = true;
            A.Options.RecoveryMemoryPath = Path.Combine(root, "memory.json");
            var fp = new byte[96 * 54];
            var assessment = new LoopDetectionAssessment(false, 0, 0, 0, false, false, "fixture");
            for (var i = 0; i < 8; i++)
                RecoveryMemoryService.AppendLoopReplayObservation("test-run", i + 1, 320, 200, fp, fp, f.Context, null, double.NaN, assessment, null, "finite", false, i / 4);
            var corpus = RecoveryMemoryService.BuildLoopReplayCorpus(File.ReadAllLines(Path.Combine(root, "loop-telemetry.jsonl")));
            var frames = corpus.Cases.Single().Frames;
            Check(frames.Count == 8 && frames.All(frame => frame.ReplayStateVersion == 1), "telemetry omitted replay version");
            Check(frames.Take(4).All(frame => frame.StateEpoch == 0) && frames.Skip(4).All(frame => frame.StateEpoch == 1), "export lost epoch boundary");
        }
        finally { Directory.Delete(root, true); }
    }
}
