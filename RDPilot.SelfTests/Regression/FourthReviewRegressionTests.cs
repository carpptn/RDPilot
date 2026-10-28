using System.Drawing;
using System.Runtime.InteropServices;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;

namespace RDPilot.SelfTests.Regression;

internal static class FourthReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("clipboard uses a window owner and transfers Unicode memory ownership", () => Clipboard("")),
        ("clipboard owner creation failure does not allocate memory", () => Clipboard("owner")),
        ("clipboard allocation failure destroys its owner", () => Clipboard("allocate")),
        ("clipboard lock failure releases memory and owner", () => Clipboard("lock")),
        ("clipboard open failure releases memory and owner", () => Clipboard("open")),
        ("clipboard empty failure closes and releases resources", () => Clipboard("empty")),
        ("clipboard set failure closes and releases resources", () => Clipboard("set")),
        ("clipboard retries busy handles with the same non-null owner", () => Clipboard("busy")),
        ("clipboard cancellation closes and releases resources", () => Clipboard("cancel")),
        ("all space aliases generate physical key events", SpaceAliases),
        ("auto verification honors confidence thresholds and boundaries", Verification),
        ("English and Polish high impact goals always verify in auto mode", HighImpact),
        ("legacy replay action identity ignores incidental decision metadata", ReplayIdentity),
        ("legacy and recorded replay share canonical action signatures", ReplayCanonicalSignatures),
        ("replay confirmed progress matches live loop detection", () => ReplayProgress("Progress", false)),
        ("legacy replay without progress retains loop detection", () => ReplayProgress(null, true)),
        ("invalid replay progress falls back to unknown", () => ReplayProgress("unexpected", true)),
        ("observed batches record every successful action exactly once", () => Observed("success", false)),
        ("observed batches preserve the prefix when the next input fails", () => Observed("input", false)),
        ("observed batches preserve the prefix when cancelled", () => Observed("cancel", false)),
        ("observed batches retain input when its observation fails", () => Observed("capture", false)),
        ("directional batches preserve the prefix on input failure", () => Observed("input", true)),
        ("directional batches preserve the prefix on cancellation", () => Observed("cancel", true))
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class ClipboardNative(string failure) : WindowsClipboard.INative, IDisposable
    {
        internal readonly List<string> Events = [];
        internal int Opens;
        internal string? Text;
        nint allocated;
        public int LastError => 5;
        public nint CreateOwner() { Events.Add("owner"); return failure == "owner" ? 0 : 123; }
        public void DestroyOwner(nint owner) { Check(owner == 123, "wrong owner destroyed"); Events.Add("destroy"); }
        public nint Allocate(int bytes) { Events.Add("allocate"); return allocated = failure == "allocate" ? 0 : Marshal.AllocHGlobal(bytes); }
        public nint Lock(nint memory) { Events.Add("lock"); return failure == "lock" ? 0 : memory; }
        public void Unlock(nint memory) => Events.Add("unlock");
        public void Free(nint memory) { Check(memory == allocated, "wrong memory freed"); Events.Add("free"); Marshal.FreeHGlobal(allocated); allocated = 0; }
        public bool Open(nint owner)
        {
            Check(owner == 123, "clipboard opened without its window owner");
            Events.Add("open"); Opens++;
            if (failure == "cancel") A.CurrentRun.CancelRequested = true;
            return failure != "open" && (failure != "busy" || Opens == 3);
        }
        public bool Empty() { Events.Add("empty"); return failure != "empty"; }
        public bool SetUnicodeText(nint memory) { Events.Add("set"); Text = Marshal.PtrToStringUni(memory); return failure != "set"; }
        public void Close() => Events.Add("close");
        public void Dispose() { if (allocated != 0) Marshal.FreeHGlobal(allocated); }
    }

    static void Clipboard(string failure)
    {
        using var scope = A.BeginRun();
        using var native = new ClipboardNative(failure);
        Exception? error = null;
        try { WindowsClipboard.SetText("Zażółć 🐈", native); }
        catch (Exception ex) { error = ex; }
        var success = failure is "" or "busy";
        Check(success == (error is null), "unexpected clipboard outcome");
        if (success)
        {
            Check(native.Text == "Zażółć 🐈", "Unicode text changed");
            Check(!native.Events.Contains("free"), "transferred clipboard memory was freed");
            Check(native.Events.IndexOf("open") < native.Events.IndexOf("empty") && native.Events.IndexOf("empty") < native.Events.IndexOf("set"), "incorrect clipboard call order");
        }
        Check(native.Events.Contains("destroy") == (failure != "owner"), "owner lifetime mismatch");
        Check(native.Events.Contains("free") == (!success && failure is not "owner" and not "allocate"), "memory cleanup mismatch");
        Check(native.Events.Contains("close") == (success || failure is "empty" or "set" or "cancel"), "open clipboard not closed exactly when needed");
        if (failure == "open") Check(native.Opens == 6, "retry limit changed");
        if (failure == "busy") Check(native.Opens == 3, "busy clipboard not retried");
        if (failure == "cancel") Check(error is OperationCanceledException && !native.Events.Contains("empty"), "cancelled clipboard was modified");
    }

    sealed class Inputs(string failure = "") : Input.IInputBackend
    {
        internal List<Input.INPUT> Events = [];
        internal int Sends;
        public uint Send(Input.INPUT[] inputs)
        {
            Sends++;
            if (Sends == 2 && failure == "input") throw new InvalidOperationException("synthetic input failure");
            if (Sends == 2 && failure == "cancel")
            {
                A.CurrentRun.CancelRequested = true;
                throw new OperationCanceledException();
            }
            Events.AddRange(inputs);
            return (uint)inputs.Length;
        }
        public int LastError => 5;
        public bool SetCursorPosition(int x, int y) => throw new InvalidOperationException("Native input prohibited");
        public void SetClipboardText(string text) => throw new InvalidOperationException("Native clipboard prohibited");
    }

    sealed class Capture(Func<bool>? fail = null) : DesktopCapture.IBackend
    {
        int count;
        public Rectangle Bounds(bool multiple) => new(0, 0, 320, 200);
        public void CopyPixels(Graphics graphics, Rectangle region)
        {
            if (fail?.Invoke() == true) throw new InvalidOperationException("synthetic capture failure");
            graphics.Clear(Color.Navy);
            graphics.FillRectangle(Brushes.White, 20 + (++count % 10) * 5, 20, 40, 40);
        }
    }

    static void SpaceAliases()
    {
        using var scope = A.BeginRun();
        var original = Input.InputBackend;
        var input = new Inputs();
        try
        {
            Input.InputBackend = input;
            foreach (var key in new[] { "space", "spacebar", " " })
            {
                input.Events.Clear();
                Input.PressKey(key);
                Check(input.Events.Count == 2 && input.Events.All(e => e.U.ki.wVk == 32 && e.U.ki.wScan == 0), "space alias used Unicode input: " + key);
                Check(input.Events[0].U.ki.dwFlags == 0 && input.Events[1].U.ki.dwFlags == 2, "space did not produce down/up events");
            }
        }
        finally { Input.InputBackend = original; }
    }

    static void Verification()
    {
        using var scope = A.BeginRun();
        A.Options.VerifyMode = "auto";
        A.Options.VerifyEarlySteps = 2;
        A.Options.VerifyLowConfidenceThreshold = .75;
        A.Options.SkipVerifyConfidenceThreshold = .92;
        foreach (var confidence in new[] { .74, .75, .8, .919, .92, 1.0 })
            Check(ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 11, new() { Type = "done", Confidence = confidence }) == (confidence < .92), "wrong confidence boundary: " + confidence);
        A.Options.SkipVerifyConfidenceThreshold = 1;
        Check(ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 11, new() { Type = "done", Confidence = .99 }), "skip threshold 1 bypassed");
        foreach (double? confidence in new double?[] { null, double.NaN, double.PositiveInfinity, -1, 2 })
            Check(ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 11, new() { Type = "done", Confidence = confidence }), "invalid confidence skipped verification");
        Check(ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 2, new() { Type = "done", Confidence = 1 }), "early completion skipped verification");
        A.Options.VerifyMode = "always";
        Check(ControlGoalPolicy.ShouldVerifyGoal("draw a cat", 11, new() { Type = "done", Confidence = 1 }), "always mode skipped");
        A.Options.VerifyMode = "off";
        Check(!ControlGoalPolicy.ShouldVerifyGoal("delete all photos", 1), "explicit off ignored");
    }

    static void HighImpact()
    {
        using var scope = A.BeginRun();
        A.Options.VerifyMode = "auto";
        foreach (var goal in new[] { "delete all photos", "remove the file", "erase documents", "save changes", "send the message", "submit the form", "publish the article", "buy the product", "run script", "execute script", "usuń zdjęcia", "zapisz zmiany", "wyślij wiadomość", "opublikuj artykuł", "kup produkt", "uruchom skrypt" })
            Check(ControlGoalPolicy.IsHighImpactGoal(goal) && ControlGoalPolicy.ShouldVerifyGoal(goal, 100, new() { Type = "done", Confidence = 1 }), "high impact goal skipped: " + goal);
        Check(!ControlGoalPolicy.IsHighImpactGoal("draw a cat"), "ordinary goal incorrectly classified");
    }

    static void ReplayIdentity()
    {
        using var scope = A.BeginRun();
        var a = new ActionDto { Type = "keys", Keys = ["right"], Confidence = .8, Note = "Move right" };
        var b = new ActionDto { Type = "keys", Keys = ["right"], Confidence = .99, Note = "Continue to the right", PlannedInputs = ["right", "down"] };
        string Signature(ActionDto action) => ReplayActionResolver.Resolve(new() { PreviousAction = action }, 320, 200).IneffectiveSignature;
        Check(Signature(a) == Signature(b), "decision metadata changed legacy identity");
        Check(Signature(a) != Signature(new() { Type = "keys", Keys = ["left"] }), "different inputs collapsed");
        Check(ControlLoopGuards.RepeatedStringCycleLength(new[] { Signature(a), Signature(b) }) == 1, "legacy repeated input was not detected");
        var sensitive = Signature(new() { Type = "type_text", Text = "private-secret", Note = "username field" });
        Check(!sensitive.Contains("private-secret") && sensitive != Signature(new() { Type = "type_text", Text = "private-secret", Note = "search field" }), "sensitive text context lost");
        Check(sensitive != Signature(new() { Type = "type_text", Text = "different-secret", Note = "username field" }), "different text collapsed");
    }

    static void ReplayCanonicalSignatures()
    {
        using var scope = A.BeginRun();
        var old = DesktopCapture.Backend;
        try
        {
            DesktopCapture.Backend = new Capture();
            A.CurrentRun.CurrentScreenMap = ScreenCoordinateMapper.Create(320, 200, 320, 200);
            A.Options.FocusCropSize = 64;
            ActionDto[] actions =
            [
                new() { Type = "keys", Keys = ["right"] },
                new() { Type = "hold_keys", Keys = ["left"], DurationMs = 250 },
                new() { Type = "click", X = .25, Y = .5 },
                new() { Type = "move", XPx = 210, YPx = 150 },
                new() { Type = "drag_drop", XPx = 20, YPx = 20, ToXPx = 220, ToYPx = 120 },
                new() { Type = "type_text", Text = "sensitive", Note = "target field" },
                new() { Type = "aim", BBox = new() { Left = 20, Top = 40, Right = 80, Bottom = 100 } },
                new() { Type = "request_crop", Crop = new() { Left = 160, Top = 40, Right = 220, Bottom = 100 } },
                new() { Type = "point", XPx = 319, YPx = 199 }
            ];
            foreach (var action in actions)
            {
                var live = Input.CaptureResolvedAction(action, null);
                Check(live.IsValid, "invalid signature fixture: " + action.Type);
                var legacy = ReplayActionResolver.Resolve(new() { PreviousAction = action }, 320, 200);
                var recorded = ReplayActionResolver.Resolve(new() { PreviousAction = action, PreviousActionGeometry = ReplayActionResolver.Record(live) }, 320, 200);
                Check(legacy.IneffectiveSignature == recorded.IneffectiveSignature, "signature mismatch: " + action.Type);
            }
            var left = ReplayActionResolver.Resolve(new() { PreviousAction = new() { Type = "aim", XPx = 20, YPx = 100 } }, 320, 200);
            var right = ReplayActionResolver.Resolve(new() { PreviousAction = new() { Type = "aim", XPx = 280, YPx = 100 } }, 320, 200);
            Check(left.IneffectiveSignature != right.IneffectiveSignature, "distinct observation targets collapsed");
        }
        finally { DesktopCapture.Backend = old; }
    }

    static void ReplayProgress(string? progress, bool expectedLoop)
    {
        using var scope = A.BeginRun();
        var replay = new LoopReplayCase { Name = "progress-regression", ScreenWidth = 320, ScreenHeight = 200 };
        var graph = new LoopStateGraph();
        var recent = new Queue<ResolvedActionSnapshot>();
        var liveDetected = false;
        for (var i = 0; i < 12; i++)
        {
            var fingerprint = Enumerable.Repeat((byte)(i % 2 == 0 ? 0 : 255), 96 * 54).ToArray();
            ActionDto? action = i == 0 ? null : new() { Type = "keys", Keys = [i % 2 == 0 ? "left" : "right"] };
            var snapshot = action is null ? null : Input.CaptureResolvedAction(action, null);
            if (snapshot is not null) recent.Enqueue(snapshot);
            var assessment = RecoveryMemoryService.AssessVisualStateCycle(graph, fingerprint, fingerprint, new("test", "test", null, null, null, null), i + 1, recent, snapshot, .5, recordLearning: false, goalProgress: progress == "Progress" ? GoalProgressState.Progress : GoalProgressState.Unknown);
            liveDetected |= assessment.IsLoop;
            replay.Frames.Add(new() { ScreenFingerprintBase64 = Convert.ToBase64String(fingerprint), ActiveWindowFingerprintBase64 = Convert.ToBase64String(fingerprint), PreviousAction = action, PreviousActionGeometry = ReplayActionResolver.Record(snapshot), WindowTitle = "test", ActiveProcess = "test", LastDelta = .5, GoalProgress = progress ?? "", ObservationPolicyVersion = 1 });
        }
        var result = LoopReplayService.ReplayCase(replay);
        Check(result.Detected == liveDetected && result.Detected == expectedLoop, "replay disagrees with live progress handling");
    }

    static void Observed(string failure, bool directional)
    {
        using var scope = A.BeginRun();
        A.Options.ObservationMode = "turn_based_interaction";
        A.Options.ScreenPollingEnabled = false;
        A.Options.ScreenPollTimeoutMs = 0;
        A.Options.UiSettleDelayMs = 0;
        A.Options.MaxActionFailuresBeforeAbort = 3;
        var oldInput = Input.InputBackend;
        var oldCapture = DesktopCapture.Backend;
        var input = new Inputs(failure);
        var run = new ControlRunState("synthetic-key", "Move the game character", AppContext.BaseDirectory, AppContext.BaseDirectory);
        run.cancelCts = new CancellationMonitor(A.CurrentRun, () => false);
        try
        {
            Input.InputBackend = input;
            DesktopCapture.Backend = new Capture(() => failure == "capture" && input.Sends > 0);
            run.turnBasedInteractionRect = new Rectangle(0, 0, 320, 200);
            var first = new ActionDto { Type = "keys", Keys = ["right"], PlannedInputs = ["right", "down"], PlanConfidence = .99 };
            var second = new ActionDto { Type = "keys", Keys = ["down"] };
            var turn = new ControlStepState(1) { Decision = new ControlDecision(first, new[] { second }), promptContext = new("Game", "game", null, null, null, null), observationFrame = ScreenshotService.CaptureObservationFrameProbe(true) };
            StepFlow flow;
            if (directional)
            {
                var result = ControlBatchService.ExecuteTurnBasedDirectionalSequenceAsync(["right", "down"], turn.observationFrame, run.turnBasedInteractionRect.Value, run.turnBasedTransitions, IntPtr.Zero, run.Cancellation).GetAwaiter().GetResult();
                Check(result.ExecutedActions.Count == 1 && result.FailedAction?.Action.Keys?.Single() == "down", "directional prefix or failed input lost");
                flow = ControlActionOutcome.ApplyObservedBatch(run, turn, result);
            }
            else flow = ControlExecutionStage.RunAsync(run, turn).GetAwaiter().GetResult();
            if (flow == StepFlow.Next) ControlCompletionStage.RunAsync(run, turn).GetAwaiter().GetResult();
            var expected = failure == "success" ? 2 : 1;
            Check(run.recentExecutedActions.Count == expected && turn.actionExecuted, "executed prefix missing or duplicated");
            Check(run.recentExecutedActions.First().Action.Keys?.Single() == "right", "wrong first executed input");
            Check(turn.turnBatchObserved && turn.executionHistoryRecorded, "observed batch was not accounted for");
            var history = run.historyBuffer.ToString();
            Check(history.Split("batch keys [right]").Length == 2, "first input not recorded exactly once");
            if (failure == "success") Check(!turn.actionExecutionFailed && run.consecutiveActionFailures == 0, "successful sequence marked failed");
            else if (failure == "cancel") Check(flow == StepFlow.Stop && run.runResult?.Outcome == ControlRunOutcome.Cancelled, "cancellation outcome lost");
            else Check(turn.actionExecutionFailed && run.consecutiveActionFailures == 1 && run.prevAction is null && history.Contains(failure == "input" ? "keys [down] failed" : "synthetic capture failure"), "failure accounting or failed action incorrect");
        }
        finally
        {
            run.cancelCts.DisposeAsync().AsTask().GetAwaiter().GetResult();
            Input.InputBackend = oldInput;
            DesktopCapture.Backend = oldCapture;
        }
    }
}
