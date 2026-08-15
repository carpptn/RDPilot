using System.Drawing;
using System.Reflection;

typeof(RDPilotApplication)
    .GetField("RecoveryMemoryEnabled", BindingFlags.Static | BindingFlags.NonPublic)!
    .SetValue(null, false);

var tests = new (string Name, Action Run)[]
{
    ("malformed pointer actions are captured without throwing", TestMalformedActionCapture),
    ("invalid bounding boxes are rejected", TestInvalidBoundingBox),
    ("out-of-range coordinates are rejected", TestCoordinateBounds),
    ("coordinate mapping remains inclusive", TestCoordinateMapping),
    ("virtual desktop mapping preserves a negative origin", TestVirtualDesktopCoordinateMapping),
    ("multi-monitor control remains explicit opt-in", TestMultiMonitorOptIn),
    ("continuous goals cannot emit done", TestContinuousActionSchema),
    ("prompt history persists, deduplicates, and navigates", TestPromptHistory),
    ("gesture actions are exposed by the control schema", TestGestureActionSchema),
    ("action policies distinguish terminal and realtime input", TestActionPolicyResolution),
    ("drag paths are validated and mapped", TestDragPathMapping),
    ("drag paths inject a complete absolute mouse-move plan", TestDragPathInputPlan),
    ("control responses expose an explicit bounded action sequence", TestControlActionBatchSchema),
    ("control streaming accepts the first complete action sequence", TestControlActionStreaming),
    ("control requests use persisted reasoning and high-threshold compaction", TestControlContextRequestOptions),
    ("control context chains restart safely and remain isolated", TestControlContextChainLifecycle),
    ("invalid response state falls back without losing request context", TestControlContextFallback),
    ("safe candidate batches stop at observation barriers", TestSafeCandidateBatchBarriers),
    ("inspection budget blocks crop loops until interaction", TestObservationActionGuard),
    ("bounded key holds reject unsafe input", TestHoldKeysValidation),
    ("adaptive observation detects a local path change", TestAdaptiveLocalObservation),
    ("turn-based observation detects small moves and records transitions", TestTurnBasedLocalTransitions),
    ("adaptive observation detects color-only local changes", TestAdaptiveColorObservation),
    ("adaptive observation detects focused text edits", TestFocusedTextObservation),
    ("ambient realtime motion is not goal progress", TestRealtimeAmbientMotion),
    ("path gestures are not classified as placement loops", TestPathGestureLoopClassification),
    ("rejected proposal cycles include multi-step patterns", TestRejectedProposalCycles),
    ("rejected proposal history resets only after observed progress", TestRejectedProposalResetPolicy),
    ("successful actions do not seed ineffective-repeat history", TestRepeatAccounting),
    ("done is excluded from recovery action history", TestDoneIsNotLearned),
    ("recovery episode action history stays bounded", TestRecoveryActionHistoryBound),
    ("spatial cycles cross grid boundaries", TestSpatialCycleDetection),
    ("state graph detects a wider repeated cycle", TestStateGraphCycle),
    ("graph calibration requires an independent recurrence", TestIndependentGraphConfirmation),
    ("goal-aligned recurring workflows are not harmful loops", TestProductiveContinuousCycle),
    ("stale graph candidates expire", TestGraphCandidateExpiry),
    ("expired graph candidates are recorded as inconclusive", TestGraphCandidateInconclusiveExpiry),
    ("runtime semantic state history stays bounded", TestRuntimeStateBounds),
    ("bandit score rewards reliable strategies", TestBanditRanking),
    ("goal context prevents unrelated lesson reuse", TestGoalAwareSimilarity),
    ("finite and continuous goals are classified generally", TestGoalModes),
    ("explicit goal mode overrides heuristic classification", TestGoalModeOverride),
    ("zero max steps enables an unlimited run", TestUnlimitedStepConfiguration),
    ("continuous wait is healthy idle rather than stagnation", TestContinuousIdle),
    ("goal progress requires an independent positive verdict", TestProgressVerdict),
    ("auto verification checks early and low-confidence completion", TestAutoVerificationPolicy),
    ("recovery validation requires post-candidate mutation progress", TestRecoveryValidationProgress),
    ("strategy attribution requires semantic target evidence", TestStrictStrategyAttribution),
    ("explicit strategy identity is attributed deterministically", TestExplicitStrategyAttribution),
    ("sensitive action signatures distinguish different inputs", TestSensitiveActionSignatures),
    ("semantic strategies do not merge different targets", TestSemanticStrategyIdentity),
    ("custom profile restores code defaults", TestCustomProfileReset),
    ("adaptive effort never lowers max", TestMaxReasoningEffortDoesNotDowngrade),
    ("profiles preserve stronger configured effort and budgets", TestProfilesPreserveStrongEffort),
    ("output retries follow the effort fallback ladder", TestOutputRetriesFollowReasoningFallback),
    ("partial token telemetry suppresses cache warnings", TestPartialTokenTelemetrySuppressesCacheWarning),
    ("stale weak lessons enter quarantine", TestQuarantine),
    ("retention preserves diverse application contexts", TestContextDiverseRetention),
    ("overflow lessons are durably archived", TestRecoveryArchive),
    ("primary memory file-size limit displaces low-value lessons", TestRecoveryFileSizeLimit),
    ("writer counters are compacted without losing totals", TestWriterCounterCompaction),
    ("concurrent bandit statistics merge without loss", TestBanditCounterMerge),
    ("concurrent inconclusive calibration statistics merge without loss", TestInconclusiveCalibrationMerge),
    ("telemetry builds labelled replay cases", TestTelemetryReplayCorpus),
    ("independent replay cases can be imported safely", TestIndependentReplayImport),
    ("recovery JSON survives corruption through backup", TestRecoveryPersistence)
};

var failures = new List<string>();
foreach (var test in tests)
{
    try
    {
        test.Run();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception ex)
    {
        failures.Add($"{test.Name}: {RootException(ex).Message}");
        Console.WriteLine($"FAIL {test.Name}: {RootException(ex).Message}");
    }
}

Console.WriteLine($"{tests.Length - failures.Count}/{tests.Length} self-tests passed.");
if (failures.Count > 0)
{
    foreach (var failure in failures)
        Console.Error.WriteLine(failure);
    return 1;
}
return 0;

static void TestMalformedActionCapture()
{
    foreach (var type in new[] { "click", "move" })
    {
        var snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
            new ActionDto { Type = type },
            null);
        Assert(!snapshot.IsValid, $"{type} should be invalid");
        Assert(snapshot.ValidationError?.Contains("requires", StringComparison.OrdinalIgnoreCase) == true,
            $"{type} should explain missing coordinates");
    }
}

static void TestInvalidBoundingBox()
{
    var action = new ActionDto
    {
        Type = "click",
        BBox = new BBox { Left = 100, Top = 100, Right = 50, Bottom = 50 }
    };
    var snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(action, null);
    Assert(!snapshot.IsValid, "inverted bbox must be invalid");

    var noOpDrag = new ActionDto
    {
        Type = "drag_drop",
        XPx = 100,
        YPx = 100,
        ToXPx = 101,
        ToYPx = 101
    };
    var dragSnapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(noOpDrag, null);
    Assert(!dragSnapshot.IsValid, "effectively zero-distance drag must be invalid");
}

static void TestCoordinateBounds()
{
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(1920, 1080, 1280, 720);
    var pixel = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto { Type = "click", XPx = 1280, YPx = 10 },
        null);
    Assert(!pixel.IsValid, "pixel outside the screenshot was accepted");

    var normalized = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto { Type = "move", X = double.NaN, Y = 0.5 },
        null);
    Assert(!normalized.IsValid, "non-finite normalized coordinate was accepted");
}

static void TestCoordinateMapping()
{
    var mapper = ScreenCoordinateMapper.Create(1920, 1080, 1280, 720);
    Assert(mapper.ImageToScreenPoint(0, 0) == (0, 0), "origin changed");
    Assert(mapper.ImageToScreenPoint(1279, 719) == (1919, 1079), "bottom-right is not inclusive");
    Assert(mapper.ScreenToImagePoint(1919, 1079) == (1279, 719), "round trip failed");
}

static void TestVirtualDesktopCoordinateMapping()
{
    var mapper = ScreenCoordinateMapper.Create(
        -1920,
        -200,
        4480,
        1640,
        2240,
        820);
    Assert(
        mapper.ImageToScreenPoint(0, 0) == (-1920, -200),
        "virtual desktop origin was lost");
    Assert(
        mapper.ImageToScreenPoint(2239, 819) == (2559, 1439),
        "virtual desktop bottom-right was not mapped inclusively");
    Assert(
        mapper.ScreenToImagePoint(-1920, -200) == (0, 0),
        "negative screen origin did not round-trip");
}

static void TestMultiMonitorOptIn()
{
    var root = typeof(RDPilotApplication);
    var field = root.GetField(
        "MultiMonitorEnabled",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var original = field.GetValue(null);
    try
    {
        field.SetValue(null, false);
        _ = RDPilotApplication.ConfigurationService.ApplyCliArgs(
            ["--multi-monitor"]);
        Assert(
            (bool)field.GetValue(null)!,
            "--multi-monitor did not enable the virtual desktop");
        _ = RDPilotApplication.ConfigurationService.ApplyCliArgs(
            ["--primary-monitor-only"]);
        Assert(
            !(bool)field.GetValue(null)!,
            "--primary-monitor-only did not restore the safe default");
    }
    finally
    {
        field.SetValue(null, original);
    }
}

static void TestContinuousActionSchema()
{
    var continuous =
        RDPilotApplication.PromptAndRequestFactory.ControlActionTypes(
            "continuous");
    var finite =
        RDPilotApplication.PromptAndRequestFactory.ControlActionTypes(
            "finite");
    Assert(
        !continuous.Contains("done", StringComparer.OrdinalIgnoreCase),
        "continuous action schema still exposes done");
    Assert(
        finite.Contains("done", StringComparer.OrdinalIgnoreCase),
        "finite action schema lost done");
}

static void TestGestureActionSchema()
{
    var types = RDPilotApplication.PromptAndRequestFactory.ControlActionTypes("finite");
    Assert(types.Contains("drag_path", StringComparer.OrdinalIgnoreCase), "drag_path is missing from the action schema");
    Assert(types.Contains("hold_keys", StringComparer.OrdinalIgnoreCase), "hold_keys is missing from the action schema");
    Assert(RDPilotApplication.ControlLoopService.IsSafeBatchedAction(
        new ActionDto { Type = "drag_path" }), "bounded draw batches are missing from the batch action gate");
    Assert(!RDPilotApplication.ControlLoopService.IsSafeBatchedAction(
        new ActionDto { Type = "hold_keys" }), "hold_keys must not be batchable");
}

static void TestDragPathMapping()
{
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(1920, 1080, 960, 540);
    var action = new ActionDto
    {
        Type = "drag_path",
        GestureKind = "draw",
        DurationMs = 600,
        Path =
        [
            new GesturePointDto { XPx = 100, YPx = 100 },
            new GesturePointDto { XPx = 400, YPx = 200 },
            new GesturePointDto { XPx = 800, YPx = 400 }
        ]
    };
    var snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(action, null);
    Assert(snapshot.IsValid, snapshot.ValidationError ?? "valid path was rejected");
    Assert(snapshot.ScreenPath.Count == 3, "resolved path lost points");
    Assert(snapshot.ScreenPath[0] == new Point(200, 200), "path start was not mapped to the desktop");
    Assert(snapshot.ObservationRegion is not null, "path did not capture an observation footprint");

    action.Path[1].XPx = 960;
    snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(action, null);
    Assert(!snapshot.IsValid, "out-of-range path point was accepted");
}

static void TestDragPathInputPlan()
{
    Point[] path =
    [
        new Point(10, 20),
        new Point(40, 20),
        new Point(40, 50),
        new Point(10, 20)
    ];
    var plan = RDPilotApplication.DesktopInputService.BuildMouseDragPlan(path, 600);
    Assert(plan.Moves.Length >= path.Length - 1, "drag plan lost movement events");
    Assert(plan.EffectiveDurationMs == 600, "drag plan changed the requested duration");
    Assert(plan.InitialHoldMs is >= 20 and <= 80, "drag plan produced an invalid initial hold");
    Assert(plan.Moves[^1] == path[^1], "drag plan did not finish at the final path point");
    foreach (var waypoint in path.Skip(1))
        Assert(plan.Moves.Contains(waypoint), $"drag plan skipped waypoint {waypoint}");

    const uint expectedFlags = 0x0001 | 0x2000 | 0x4000 | 0x8000;
    var topLeft = RDPilotApplication.DesktopInputService.BuildAbsoluteMouseMoveData(
        -1920,
        0,
        -1920,
        0,
        3840,
        1080);
    var bottomRight = RDPilotApplication.DesktopInputService.BuildAbsoluteMouseMoveData(
        1919,
        1079,
        -1920,
        0,
        3840,
        1080);
    Assert(topLeft.Dx == 0 && topLeft.Dy == 0, "virtual-desktop origin was normalized incorrectly");
    Assert(bottomRight.Dx == 65535 && bottomRight.Dy == 65535, "virtual-desktop edge was normalized incorrectly");
    Assert(topLeft.Flags == expectedFlags, "absolute drag movement is missing required SendInput flags");
    var firstDeadline = RDPilotApplication.DesktopInputService.GestureMoveTargetElapsedMs(plan, 0);
    var finalDeadline = RDPilotApplication.DesktopInputService.GestureMoveTargetElapsedMs(plan, plan.Moves.Length - 1);
    Assert(firstDeadline >= plan.InitialHoldMs, "first drag movement precedes the initial button hold");
    Assert(Math.Abs(finalDeadline - plan.EffectiveDurationMs) < 0.001,
        "drag schedule accumulates per-move sleep instead of ending at the absolute deadline");
}

static void TestSafeCandidateBatchBarriers()
{
    List<ActionDto> launchChain =
    [
        new ActionDto { Type = "keys", Keys = ["win+r"] },
        new ActionDto { Type = "type_text", Text = "mspaint" },
        new ActionDto { Type = "keys", Keys = ["enter"] },
        new ActionDto { Type = "wait", WaitSeconds = 2 },
        new ActionDto { Type = "keys", Keys = ["escape"] }
    ];
    var safe = RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(launchChain);
    Assert(safe.Select(action => action.Type).SequenceEqual(["type_text", "keys", "wait"]),
        "safe launch chain did not stop at its wait barrier");

    var afterCommit = RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
    [
        new ActionDto { Type = "keys", Keys = ["enter"] },
        new ActionDto { Type = "keys", Keys = ["escape"] }
    ]);
    Assert(afterCommit.Length == 0, "state-dependent key was queued after a committing key");

    List<ActionDto> formChain =
    [
        new ActionDto { Type = "keys", Keys = ["ctrl", "a"] },
        new ActionDto { Type = "type_text", Text = "1100" },
        new ActionDto { Type = "keys", Keys = ["tab"] },
        new ActionDto { Type = "type_text", Text = "540" },
        new ActionDto { Type = "keys", Keys = ["enter"] }
    ];
    var safeForm = RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(formChain);
    Assert(safeForm.Select(action => action.Type).SequenceEqual(["type_text", "keys", "type_text", "keys"]),
        "deterministic focused form-edit sequence was split by an observation barrier");

    List<ActionDto> observedTurnClicks =
    [
        new ActionDto
        {
            Type = "click",
            XPx = 100,
            YPx = 100,
            Note = "Move up through the corridor and then right toward the goal.",
            PlannedInputs = ["ArrowUp", "ArrowRight", "ArrowRight"],
            PlanConfidence = 0.8
        },
        new ActionDto
        {
            Type = "click",
            XPx = 120,
            YPx = 100,
            Note = "Continue toward the upper terminal."
        },
        new ActionDto
        {
            Type = "click",
            XPx = 120,
            YPx = 100,
            Note = "Continue toward the upper terminal."
        }
    ];
    var observedTurnFollowUps =
        RDPilotApplication.OpenAiResponsesService.SafeObservedTurnFollowUps(
            observedTurnClicks,
            12);
    Assert(observedTurnFollowUps.Length == 2 &&
           observedTurnClicks[0].ResolvedTurnInputLabel == "ArrowUp" &&
           observedTurnFollowUps.All(action =>
               action.ResolvedTurnInputLabel == "ArrowRight"),
        "descriptive notes overrode the positional planned-input mapping");

    var drawOne = new ActionDto
    {
        Type = "drag_path",
        GestureKind = "draw",
        DurationMs = 300,
        Path = [new GesturePointDto { XPx = 10, YPx = 10 }, new GesturePointDto { XPx = 20, YPx = 20 }]
    };
    var drawTwo = new ActionDto
    {
        Type = "drag_path",
        GestureKind = "draw",
        DurationMs = 300,
        Path = [new GesturePointDto { XPx = 30, YPx = 30 }, new GesturePointDto { XPx = 40, YPx = 40 }]
    };
    Assert(RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps([drawOne, drawTwo]).Length == 0,
        "draw strokes were batched without a stable-canvas context");
    var drawBatch = RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
        [drawOne, drawTwo],
        allowDrawGestureBatch: true);
    Assert(drawBatch.Length == 1 && drawBatch[0] == drawTwo,
        "independent draw strokes on a stable canvas were not batched");

    var pan = new ActionDto
    {
        Type = "drag_path",
        GestureKind = "pan",
        DurationMs = 300,
        Path = [new GesturePointDto { XPx = 50, YPx = 50 }, new GesturePointDto { XPx = 60, YPx = 60 }]
    };
    Assert(RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
        [drawOne, pan],
        allowDrawGestureBatch: true).Length == 0,
        "state-changing pan gesture was accepted into a draw batch");
    Assert(RDPilotApplication.ControlLoopService.AllowsStableCanvasDrawBatch(
        new UiPromptContext("Paint", "mspaint", "canvas", null, null, null),
        "static_ui"),
        "recognized drawing application did not enable stable-canvas draw batching");
    Assert(!RDPilotApplication.ControlLoopService.AllowsStableCanvasDrawBatch(
        new UiPromptContext("Game", "game", "canvas", null, null, null),
        "realtime_interaction"),
        "realtime canvas incorrectly enabled draw batching");

    var unsafeGesture = RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
    [
        new ActionDto { Type = "type_text", Text = "x" },
        new ActionDto { Type = "drag_path", GestureKind = "draw" }
    ]);
    Assert(unsafeGesture.Length == 0, "gesture was queued without a fresh observation");

    var invalidFollowUp = RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
    [
        new ActionDto { Type = "keys", Keys = ["win"] },
        new ActionDto { Type = "type_text" }
    ]);
    Assert(invalidFollowUp.Length == 0, "invalid text input was queued after a valid prelude");

    Assert(
        RDPilotApplication.ControlLoopService.BatchTransitionDelayMs(
            launchChain[0],
            launchChain[1]) >= 200,
        "text-entry prelude does not give the target time to receive focus");
    Assert(
        RDPilotApplication.ControlLoopService.BatchTransitionDelayMs(
            launchChain[1],
            launchChain[2]) >= 400,
        "commit key does not give dynamic suggestions time to become ready");
    Assert(
        RDPilotApplication.ControlLoopService.BatchTransitionDelayMs(
            launchChain[2],
            launchChain[3]) == 0,
        "a terminal wait received a redundant transition delay");
    Assert(RDPilotApplication.ControlLoopService.IsAdaptiveBatchWaitTrigger(launchChain[2]),
        "a committing key does not enable adaptive terminal waiting");
    Assert(!RDPilotApplication.ControlLoopService.IsAdaptiveBatchWaitTrigger(launchChain[1]),
        "plain text entry incorrectly enables adaptive terminal waiting");
    Assert(RDPilotApplication.ControlLoopService.IsNavigationBatchPrelude(launchChain[0]) &&
           RDPilotApplication.ControlLoopService.IsNavigationBatchPrelude(
               new ActionDto { Type = "keys", Keys = ["ctrl+l"] }) &&
           !RDPilotApplication.ControlLoopService.IsNavigationBatchPrelude(
               new ActionDto { Type = "keys", Keys = ["ctrl+a"] }),
        "navigation batch preludes are not distinguished from ordinary shortcuts");
    Assert(RDPilotApplication.ControlLoopService.AdaptiveBatchMinimumObservationMs(
               timeoutMs: 6000,
               navigationLikeBatch: true) == 6000 &&
           RDPilotApplication.ControlLoopService.AdaptiveBatchMinimumObservationMs(
               timeoutMs: 6000,
               navigationLikeBatch: false) == 900,
        "navigation waits can still settle on a short-lived static loading screen");
    var delayedTurnReactionPolicy =
        RDPilotApplication.ControlLoopService.ResolveTurnNoEffectPolicy(
            immediateReactionObserved: false);
    Assert(delayedTurnReactionPolicy.ExtendObservation &&
           !delayedTurnReactionPolicy.ReplayInput,
        "a delayed turn reaction still permits duplicate input injection");

    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(1000, 800, 1000, 800);
    var firstTurnClick = new ActionDto
    {
        Type = "click",
        X = 0.2,
        Y = 0.8,
        Note = "move up",
        PlannedInputs = ["ArrowUp", "ArrowUp", "ArrowRight"],
        PlanConfidence = 0.60
    };
    var repeatedTurnClick = new ActionDto
    {
        Type = "click",
        X = 0.2,
        Y = 0.8,
        Note = "move up"
    };
    var mixedTurnKey = new ActionDto
    {
        Type = "keys",
        Keys = ["ArrowRight"]
    };
    Assert(RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
            [firstTurnClick, repeatedTurnClick, mixedTurnKey]).Length == 0,
        "turn clicks bypassed the explicit observed-batch gate");
    var observedTurnBatch =
        RDPilotApplication.OpenAiResponsesService.SafeBatchFollowUps(
            [firstTurnClick, repeatedTurnClick, mixedTurnKey],
            observedTurnBatchLimit: 4);
    Assert(observedTurnBatch.SequenceEqual([repeatedTurnClick, mixedTurnKey]),
        "a reversible planned click/key route was not accepted with per-input observation");
}

static void TestObservationActionGuard()
{
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(1920, 1080, 640, 360);
    var guard = new RDPilotApplication.ControlLoopService.ObservationActionGuardState();
    ResolvedActionSnapshot Snapshot(ActionDto action) =>
        RDPilotApplication.DesktopInputService.CaptureResolvedAction(action, null);

    var boardCrop = Snapshot(new ActionDto
    {
        Type = "request_crop",
        Crop = new BBox { Left = 230, Top = 90, Right = 405, Bottom = 220 }
    });
    var controlsCrop = Snapshot(new ActionDto
    {
        Type = "request_crop",
        Crop = new BBox { Left = 230, Top = 215, Right = 410, Bottom = 310 }
    });
    var thirdCrop = Snapshot(new ActionDto
    {
        Type = "request_crop",
        Crop = new BBox { Left = 30, Top = 40, Right = 150, Bottom = 150 }
    });

    Assert(!guard.TryGetBlockReason(boardCrop, 2, out _), "first inspection was blocked");
    guard.RecordExecuted(boardCrop);
    Assert(guard.TryGetBlockReason(boardCrop, 2, out var duplicateReason) &&
           duplicateReason.Contains("already inspected", StringComparison.Ordinal),
        "repeated inspection of the same region was not blocked");
    Assert(guard.TryGetBlockReason(
               controlsCrop,
               2,
               singleInspectionBeforeInteraction: true,
               out var nestedTurnCropReason) &&
           nestedTurnCropReason.Contains("already supplied", StringComparison.Ordinal),
        "a second nested crop was allowed in the same turn-based state");
    Assert(!guard.TryGetBlockReason(controlsCrop, 2, out _), "second distinct inspection was blocked");
    guard.RecordExecuted(controlsCrop);
    Assert(guard.TryGetBlockReason(thirdCrop, 2, out var budgetReason) &&
           budgetReason.Contains("limit of 2", StringComparison.Ordinal),
        "inspection budget did not block an A-B-A style observation loop");

    var aim = Snapshot(new ActionDto
    {
        Type = "aim",
        BBox = new BBox { Left = 250, Top = 230, Right = 300, Bottom = 280 }
    });
    Assert(!guard.TryGetBlockReason(aim, 2, out _), "precision AIM was blocked after inspection");
    guard.RecordExecuted(aim);
    Assert(guard.TryGetBlockReason(aim, 2, out var aimReason) &&
           aimReason.Contains("already active", StringComparison.Ordinal),
        "repeated AIM was not blocked");

    guard.RecordExecuted(Snapshot(new ActionDto { Type = "wait", WaitSeconds = 1 }));
    Assert(guard.RequiresInteraction(2), "wait incorrectly reset the inspection budget");
    guard.RecordExecuted(Snapshot(new ActionDto { Type = "keys", Keys = ["left"] }));
    Assert(!guard.RequiresInteraction(2) && !guard.AimIssuedSinceInteraction,
        "state-changing input did not reset the inspection budget");
    Assert(!guard.TryGetBlockReason(thirdCrop, 2, out _),
        "new inspection remained blocked after interaction");
}

static void TestControlActionBatchSchema()
{
    var schemaJson = System.Text.Json.JsonSerializer.Serialize(
        RDPilotApplication.PromptAndRequestFactory.ControlActionBatchSchema(1280, 720));
    using var schema = System.Text.Json.JsonDocument.Parse(schemaJson);
    var actions = schema.RootElement.GetProperty("properties").GetProperty("actions");
    Assert(actions.GetProperty("minItems").GetInt32() == 1, "action sequence may be empty");
    Assert(actions.GetProperty("maxItems").GetInt32() >= 1, "action sequence has an invalid bound");
    var variants = actions.GetProperty("items").GetProperty("anyOf");
    Assert(variants.GetArrayLength() > 1, "control actions are not represented by compact variants");
    var keysVariant = variants.EnumerateArray().First(variant =>
        variant.GetProperty("properties").GetProperty("type").GetProperty("enum")[0].GetString() == "keys");
    var keysProperties = keysVariant.GetProperty("properties");
    Assert(keysProperties.TryGetProperty("keys", out _), "keys variant lost its action payload");
    Assert(keysProperties.GetProperty("keys").GetProperty("maxItems").GetInt32() == 32,
        "directional key actions do not expose the configured extended horizon");
    Assert(!keysProperties.TryGetProperty("x", out _) &&
           !keysProperties.TryGetProperty("text", out _) &&
           !keysProperties.TryGetProperty("path", out _),
        "keys variant still forces unrelated action fields");
    Assert(
        keysVariant.GetProperty("required").GetArrayLength() == keysProperties.EnumerateObject().Count(),
        "strict action variant has optional fields");
    var envelopeProperties = schema.RootElement.GetProperty("properties");
    Assert(envelopeProperties.TryGetProperty("world_state_summary", out _) &&
           envelopeProperties.TryGetProperty("mechanics_hypothesis", out _) &&
           envelopeProperties.TryGetProperty("salient_change_observation", out _) &&
           envelopeProperties.TryGetProperty("short_term_plan", out _) &&
           envelopeProperties.TryGetProperty("plan_status", out _) &&
           envelopeProperties.TryGetProperty("plan_revision_reason", out _) &&
           envelopeProperties.TryGetProperty("planned_inputs", out _) &&
           envelopeProperties.TryGetProperty("plan_waypoint", out _) &&
           envelopeProperties.TryGetProperty("plan_state_id", out _) &&
           envelopeProperties.TryGetProperty("plan_confidence", out _),
        "ActionBatch envelope lost stateful working-memory fields");
    Assert(envelopeProperties.GetProperty("planned_inputs").GetProperty("maxItems").GetInt32() == 32,
        "structured turn plans do not preserve the configured extended horizon");
    Assert(schema.RootElement.GetProperty("required").GetArrayLength() == 15,
        "ActionBatch envelope does not carry shared sequence metadata");

    var rules = RDPilotApplication.PromptAndRequestFactory.BuildSystemRules();
    Assert(rules.Contains("actions[0]", StringComparison.Ordinal),
        "control prompt does not identify the first executable action");
    Assert(!rules.Contains("Return EXACTLY ONE action per round", StringComparison.Ordinal),
        "legacy single-action instruction still disables batching");
    Assert(rules.Contains("short_term_plan", StringComparison.Ordinal) &&
           rules.Contains("planned_inputs", StringComparison.Ordinal) &&
           rules.Contains("aggressive batching is the default", StringComparison.Ordinal) &&
           rules.Contains("count the route cell by cell", StringComparison.Ordinal),
        "control prompt does not connect a conditional plan with bounded execution");

    var temporalRequest = RDPilotApplication.PromptAndRequestFactory.BuildRequestBody(
        "gpt-5.6-luna",
        rules,
        "solve the visible board",
        "",
        "data:image/jpeg;base64,current-screen",
        1280,
        720,
        0,
        0,
        0,
        0,
        "data:image/jpeg;base64,current-focus",
        new Rectangle(100, 100, 600, 400),
        null,
        null,
        new UiPromptContext("Board", "browser", null, null, null, null),
        false,
        null,
        false,
        "finite",
        "data:image/jpeg;base64,previous-focus",
        "data:image/jpeg;base64,reference-focus",
        [new TurnChangeImagePair(
            "data:image/jpeg;base64,change-before",
            "data:image/jpeg;base64,change-after",
            null,
            null,
            1)],
        10000);
    var temporalRequestJson = System.Text.Json.JsonSerializer.Serialize(temporalRequest);
    Assert(temporalRequestJson.Contains("PREVIOUS_TURN_STATE_IMAGE", StringComparison.Ordinal) &&
           temporalRequestJson.Contains("TURN_REFERENCE_IMAGE", StringComparison.Ordinal) &&
           temporalRequestJson.Contains("current visual epoch only", StringComparison.Ordinal) &&
           temporalRequestJson.Contains("CURRENT_FOCUS_IMAGE", StringComparison.Ordinal) &&
           temporalRequestJson.Contains("SALIENT_CHANGE_REGION_1_BEFORE", StringComparison.Ordinal) &&
           temporalRequestJson.Contains("SALIENT_CHANGE_REGION_1_AFTER", StringComparison.Ordinal),
        "control request did not label temporal turn images");
    Assert(temporalRequestJson.Contains("\"max_output_tokens\":10000", StringComparison.Ordinal),
        "salient control request lost its dynamic output budget");

    var envelope = System.Text.Json.JsonSerializer.Serialize(new ActionBatchDto
    {
        Actions =
        [
            new ActionDto { Type = "keys", Keys = ["win"] },
            new ActionDto { Type = "type_text", Text = "Paint" }
        ],
        PlannedInputs = ["ArrowUp", "ArrowLeft"],
        PlanWaypoint = "visible opening",
        PlanStateId = "S3",
        PlanConfidence = 0.9
    });
    var responseJson = System.Text.Json.JsonSerializer.Serialize(new
    {
        output = new[]
        {
            new
            {
                content = new[]
                {
                    new { type = "output_text", text = envelope }
                }
            }
        }
    });
    using var response = System.Text.Json.JsonDocument.Parse(responseJson);
    Assert(
        RDPilotApplication.OpenAiResponsesService.TryParseControlActionSequence(
            response.RootElement,
            out var parsedActions,
            out var payloadCount,
            out var legacyPayload),
        "explicit action sequence could not be parsed from a Responses payload");
    Assert(parsedActions.Count == 2 && payloadCount == 1 && !legacyPayload &&
           parsedActions[0].PlannedInputs?.SequenceEqual(
               ["ArrowUp", "ArrowLeft"]) == true &&
           parsedActions[0].PlanStateId == "S3" &&
           parsedActions[0].PlanConfidence == 0.9,
        "explicit action sequence lost actions or was treated as legacy output");
}

static void TestControlActionStreaming()
{
    var actionJson = "{\"actions\":[{\"type\":\"keys\",\"keys\":[\"win\"]}],\"confidence\":0.99,\"note\":\"open Start\",\"recovery_strategy_id\":null,\"recovery_strategy_step\":null}";
    var eventJson = System.Text.Json.JsonSerializer.Serialize(new
    {
        type = "response.output_text.done",
        text = actionJson
    });
    string? responseId = "resp_test";
    Assert(
        RDPilotApplication.OpenAiResponsesService.TryHandleControlStreamEvent(
            eventJson,
            allowEarlyAccept: true,
            ref responseId,
            out var raw,
            out var actions,
            out var completed),
        "complete streamed ActionBatch was not accepted");
    Assert(actions is { Count: 1 } && actions[0].Type == "keys" &&
           actions[0].Confidence == 0.99 && actions[0].Note == "open Start",
        "streamed ActionBatch lost its first action");
    Assert(!completed && raw.Contains("rdpilot_stream_early_accept", StringComparison.Ordinal),
        "early stream acceptance did not preserve a replayable response");

    Assert(
        !RDPilotApplication.OpenAiResponsesService.TryHandleControlStreamEvent(
            eventJson,
            allowEarlyAccept: false,
            ref responseId,
            out _,
            out _,
            out _),
        "stream was accepted early while previous_response_id state required completion");
}

static void TestControlContextRequestOptions()
{
    var root = typeof(RDPilotApplication);
    var stateField = root.GetField("UsePreviousResponseState", BindingFlags.Static | BindingFlags.NonPublic)!;
    var contextField = root.GetField("ControlReasoningContext", BindingFlags.Static | BindingFlags.NonPublic)!;
    var compactionField = root.GetField("ControlContextCompactionEnabled", BindingFlags.Static | BindingFlags.NonPublic)!;
    var thresholdField = root.GetField("ControlContextCompactThreshold", BindingFlags.Static | BindingFlags.NonPublic)!;
    var originals = new[]
    {
        stateField.GetValue(null),
        contextField.GetValue(null),
        compactionField.GetValue(null),
        thresholdField.GetValue(null)
    };

    try
    {
        stateField.SetValue(null, true);
        contextField.SetValue(null, "all_turns");
        compactionField.SetValue(null, true);
        thresholdField.SetValue(null, 700_000);

        var control = new Dictionary<string, object>();
        RDPilotApplication.PromptAndRequestFactory.AddReasoningOptions(
            control,
            "gpt-5.6-luna",
            "max",
            6000,
            "control");
        RDPilotApplication.PromptAndRequestFactory.AddControlContextManagement(
            control,
            enableContextCompaction: true);
        var controlJson = System.Text.Json.JsonSerializer.Serialize(control);
        Assert(controlJson.Contains("\"context\":\"all_turns\"", StringComparison.Ordinal),
            "control request did not preserve reasoning across turns");
        Assert(controlJson.Contains("\"context_management\":[{\"type\":\"compaction\",\"compact_threshold\":700000}]", StringComparison.Ordinal),
            "control request did not use the documented server-side compaction shape");

        var verifier = new Dictionary<string, object>();
        RDPilotApplication.PromptAndRequestFactory.AddReasoningOptions(
            verifier,
            "gpt-5.6-luna",
            "max",
            6000,
            "verify");
        var verifierJson = System.Text.Json.JsonSerializer.Serialize(verifier);
        Assert(verifierJson.Contains("\"context\":\"current_turn\"", StringComparison.Ordinal),
            "independent verifier inherited the control reasoning context");
        Assert(!verifierJson.Contains("previous_response_id", StringComparison.Ordinal),
            "independent verifier inherited the control response chain");
    }
    finally
    {
        stateField.SetValue(null, originals[0]);
        contextField.SetValue(null, originals[1]);
        compactionField.SetValue(null, originals[2]);
        thresholdField.SetValue(null, originals[3]);
    }
}

static void TestControlContextChainLifecycle()
{
    var chain = new RDPilotApplication.ControlLoopService.ControlContextChain(
        "1234567890abcdef",
        enabled: true,
        compactionEnabled: true,
        fallbackLimit: 3);
    Assert(chain.PreviousResponseIdForRequest is null, "new task inherited an old response id");

    chain.RecordResult("resp_first", false, false, false);
    Assert(chain.PreviousResponseIdForRequest == "resp_first" && chain.TurnCount == 1,
        "completed response did not advance the control chain");

    chain.RecordResult("resp_restarted", true, false, true);
    Assert(chain.PreviousResponseIdForRequest == "resp_restarted" && chain.RestartCount == 1,
        "checkpoint fallback did not establish a replacement chain");

    chain.RecordResult("resp_no_compaction", false, true, false);
    Assert(chain.Enabled && !chain.CompactionEnabled,
        "compaction rejection disabled the whole control chain");

    chain.RecordResult("resp_restart_2", true, false, false);
    chain.RecordResult("resp_restart_3", true, false, false);
    Assert(!chain.Enabled && chain.PreviousResponseIdForRequest is null,
        "repeated response-state failures did not fall back to explicit application history");
}

static void TestControlContextFallback()
{
    var body = new Dictionary<string, object>
    {
        ["model"] = "gpt-5.6-luna",
        ["previous_response_id"] = "resp_old",
        ["context_management"] = new object[]
        {
            new { type = "compaction", compact_threshold = 700_000 }
        },
        ["reasoning"] = new { effort = "max", context = "all_turns" },
        ["input"] = new object[] { new { role = "user", content = "checkpoint" } }
    };

    Assert(
        RDPilotApplication.OpenAiResponsesService.IsPreviousResponseStateFailure(
            400,
            "Invalid previous_response_id: response not found"),
        "previous-response API error was not classified as a chain failure");
    Assert(
        RDPilotApplication.OpenAiResponsesService.TryBuildRequestWithoutProperty(
            body,
            "previous_response_id",
            out var fallbackBody),
        "could not construct a checkpoint retry body");
    var fallbackJson = System.Text.Json.JsonSerializer.Serialize(fallbackBody);
    Assert(!fallbackJson.Contains("previous_response_id", StringComparison.Ordinal) &&
           fallbackJson.Contains("context_management", StringComparison.Ordinal) &&
           fallbackJson.Contains("checkpoint", StringComparison.Ordinal),
        "checkpoint retry discarded current request state or retained the invalid response id");

    var completed = "{\"id\":\"resp_ok\",\"status\":\"completed\",\"output\":[{\"type\":\"compaction\"}]}";
    var incomplete = "{\"id\":\"resp_bad\",\"status\":\"incomplete\",\"output\":[]}";
    Assert(RDPilotApplication.OpenAiResponsesService.TryGetCompletedResponseId(completed) == "resp_ok",
        "completed response id was not accepted");
    Assert(RDPilotApplication.OpenAiResponsesService.TryGetCompletedResponseId(incomplete) is null,
        "incomplete response id was accepted into the chain");
    using var completedDocument = System.Text.Json.JsonDocument.Parse(completed);
    Assert(RDPilotApplication.OpenAiResponsesService.ContainsCompactionItem(completedDocument.RootElement),
        "server compaction output was not detected");
}

static void TestHoldKeysValidation()
{
    var valid = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto
        {
            Type = "hold_keys",
            Keys = ["w", "d"],
            DurationMs = 500
        },
        null);
    Assert(valid.IsValid, valid.ValidationError ?? "valid key hold was rejected");

    var unbounded = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto
        {
            Type = "hold_keys",
            Keys = ["w"],
            DurationMs = 60_000
        },
        null);
    Assert(!unbounded.IsValid, "unbounded key hold was accepted");

    var chord = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto
        {
            Type = "hold_keys",
            Keys = ["shift+w"],
            DurationMs = 500
        },
        null);
    Assert(!chord.IsValid, "chord string was accepted as an individual held key");

    var systemKey = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto
        {
            Type = "hold_keys",
            Keys = ["win"],
            DurationMs = 500
        },
        null);
    Assert(!systemKey.IsValid, "unsafe system key was accepted for holding");

    Assert(RDPilotApplication.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(
            ["win+down"],
            ownConsoleForeground: true) &&
           RDPilotApplication.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(
            ["win", "ArrowDown"],
            ownConsoleForeground: true),
        "Win+Down aimed at RDPilot's own console was not routed through its window handle");
    Assert(!RDPilotApplication.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(
            ["win+down"],
            ownConsoleForeground: false) &&
           !RDPilotApplication.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(
            ["win+left"],
            ownConsoleForeground: true),
        "the own-console safeguard blocked a legitimate system window shortcut");
}

static void TestActionPolicyResolution()
{
    RDPilotApplication.ConfigurationService.ApplyObservationMode("auto", "self-test");
    var session = new RDPilotApplication.AdaptiveObservationSession();
    var terminalKeys = new ResolvedActionSnapshot(
        new ActionDto { Type = "keys", Keys = ["enter"] },
        "keys [enter]",
        "keys:enter",
        null);
    Assert(
        session.ResolveActionPolicy(
            terminalKeys,
            new UiPromptContext("Terminal", "windowsterminal", "console", null, null, null)) == "event_driven",
        "terminal Enter did not select event-driven observation");

    var hold = new ResolvedActionSnapshot(
        new ActionDto { Type = "hold_keys", Keys = ["w"], DurationMs = 250 },
        "hold_keys [w] 250ms",
        "hold_keys:w:250",
        null);
    Assert(
        session.ResolveActionPolicy(
            hold,
            new UiPromptContext("Game", "game", "canvas", null, null, null)) == "realtime_interaction",
        "bounded key hold did not select realtime observation");

    var puzzleInspection = new ResolvedActionSnapshot(
        new ActionDto { Type = "request_crop", Note = "inspect the level board" },
        "request_crop",
        "request_crop:board",
        null);
    Assert(
        session.ResolveActionPolicy(
            puzzleInspection,
            new UiPromptContext("ARC-AGI-3 Task #ls20", "msedge", "level board controls", null, null, null),
            "graj w grę zagadkę i przechodź kolejne poziomy") == "turn_based_interaction",
        "turn-based puzzle did not select discrete interaction observation");
    var promptSession = new RDPilotApplication.AdaptiveObservationSession();
    promptSession.PrepareForPrompt(
        new UiPromptContext("ARC-AGI-3 Task #ls20", "msedge", "level board controls", null, null, null),
        "graj w grę zagadkę i przechodź kolejne poziomy");
    Assert(promptSession.EffectiveProfile == "turn_based_interaction",
        "visible puzzle context did not enable batching before the first board action");
    Assert(
        !RDPilotApplication.AdaptiveObservationSession.IsTurnBasedInteractionContext(
            "play a realtime racing game",
            "Racing Game",
            "animated canvas",
            "steer the car"),
        "realtime game was misclassified as turn-based");
}

static void TestAdaptiveLocalObservation()
{
    RDPilotApplication.ConfigurationService.ApplyObservationMode("auto", "self-test");
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
    var action = new ActionDto
    {
        Type = "drag_path",
        GestureKind = "draw",
        DurationMs = 500,
        Path =
        [
            new GesturePointDto { XPx = 10, YPx = 50 },
            new GesturePointDto { XPx = 90, YPx = 50 }
        ]
    };
    var snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(action, null);
    var beforeDetail = Enumerable.Repeat((byte)255, 10_000).ToArray();
    var afterDetail = beforeDetail.ToArray();
    for (var x = 10; x <= 90; x++)
        afterDetail[50 * 100 + x] = 0;
    var stable = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
    var before = new ScreenObservationFrame(
        stable,
        stable,
        beforeDetail,
        100,
        100,
        new Rectangle(0, 0, 100, 100));
    var after = new ScreenObservationFrame(
        stable.ToArray(),
        stable.ToArray(),
        afterDetail,
        100,
        100,
        new Rectangle(0, 0, 100, 100));
    var session = new RDPilotApplication.AdaptiveObservationSession();
    var assessment = session.Assess(
        before,
        after,
        snapshot,
        new UiPromptContext("Paint", "mspaint", "canvas", null, null, null),
        new UiPromptContext("Paint", "mspaint", "canvas", null, null, null),
        "finite");
    Assert(assessment.Profile == "local_editing", "draw path did not select local_editing");
    Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed, "thin local path change was not confirmed");
    Assert(assessment.GoalProgress == GoalProgressState.Progress, "confirmed path was not treated as progress");

    var click = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto { Type = "click", XPx = 50, YPx = 50 },
        null);
    var coarseOnly = new ScreenObservationFrame(
        stable,
        stable,
        [],
        0,
        0,
        new Rectangle(0, 0, 100, 100));
    var clickAssessment = session.Assess(
        coarseOnly,
        coarseOnly,
        click,
        new UiPromptContext("Paint", "mspaint", "canvas", null, null, null),
        new UiPromptContext("Paint", "mspaint", "canvas", null, null, null),
        "finite");
    Assert(clickAssessment.ActionPolicy == "static_ui", "a discrete click inherited the prior drawing policy");
}

static void TestTurnBasedLocalTransitions()
{
    RDPilotApplication.ConfigurationService.ApplyObservationMode("auto", "self-test");
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
    ScreenObservationFrame SolidFrame(byte value) => new(
        Enumerable.Repeat(value, 64).ToArray(),
        Enumerable.Repeat(value, 64).ToArray(),
        [],
        0,
        0,
        new Rectangle(0, 0, 100, 100));
    var promptLoading = SolidFrame(0);
    var settledStart = SolidFrame(2);
    Assert(!RDPilotApplication.ControlLoopService.ShouldCheckPreRegionTurnActionFreshness(
            "turn_based_interaction",
            null,
            new ActionDto { Type = "keys", Keys = ["Space"] }),
        "a reversible semantic key was discarded solely because loading completed");
    Assert(RDPilotApplication.ControlLoopService.ShouldCheckPreRegionTurnActionFreshness(
            "turn_based_interaction",
            null,
            new ActionDto { Type = "click", X = 0.5, Y = 0.5 }),
        "a coordinate-dependent click bypasses stale-screen protection");
    Assert(RDPilotApplication.ControlLoopService.ShouldDiscardPreRegionTurnAction(
            promptLoading,
            settledStart,
            settledStart,
            out var stalePromptDelta,
            out var staleStabilityDelta) &&
           stalePromptDelta > 0.0035 &&
           staleStabilityDelta == 0,
        "a settled Loading-to-Start transition did not invalidate the stale action");
    Assert(!RDPilotApplication.ControlLoopService.ShouldDiscardPreRegionTurnAction(
            promptLoading,
            settledStart,
            SolidFrame(4),
            out _,
            out _),
        "ongoing animation was mistaken for a settled external state transition");
    Assert(RDPilotApplication.ControlLoopService.IsTurnBasedStateInspection(
            new ActionDto { Type = "request_crop", Note = "przybliżam planszę poziomu" }),
        "Polish board crop was not recognized as a persistent turn-based region");
    Assert(RDPilotApplication.ControlLoopService.IsTurnBasedStateInspection(
            new ActionDto
            {
                Type = "request_crop",
                Note = "Inspect the level playfield before testing one control."
            }),
        "playfield crop was rejected because its note mentioned a later control test");
    Assert(!RDPilotApplication.ControlLoopService.IsTurnBasedStateInspection(
            new ActionDto { Type = "request_crop", Note = "inspect control buttons" }),
        "controls crop incorrectly replaced the persistent board region");
    Assert(!RDPilotApplication.ControlLoopService.IsTurnBasedStateInspection(
            new ActionDto { Type = "request_crop", Note = "przybliż sterowanie pod planszą" }),
        "Polish controls crop incorrectly replaced the persistent board region");
    Assert(RDPilotApplication.ControlLoopService.ShouldEstablishTurnBasedInteractionRegion(
            new ActionDto
            {
                Type = "request_crop",
                Note = "Inspekcja głównego pola gry przed pierwszym ruchem"
            },
            regionRequired: true),
        "required primary region still depended on a recognized language form");
    Assert(RDPilotApplication.ControlLoopService.ShouldEstablishTurnBasedInteractionRegion(
            new ActionDto
            {
                Type = "request_crop",
                Note = "Examine the requested primary interaction surface"
            },
            regionRequired: true),
        "required primary region still depended on known board vocabulary");
    Assert(!RDPilotApplication.ControlLoopService.ShouldEstablishTurnBasedInteractionRegion(
            new ActionDto
            {
                Type = "request_crop",
                Note = "inspect the help and control panel"
            },
            regionRequired: true),
        "an explicitly auxiliary crop became the required primary region");
    var nestedBoardCrop = new ActionDto
    {
        Type = "request_crop",
        Note = "Powiększam centralne rozgałęzienie planszy."
    };
    Assert(RDPilotApplication.ControlLoopService.CanReplaceTurnBasedInteractionRegion(
            hasExistingRegion: true,
            existingRegionIsAutomatic: true,
            automaticRegionWasRefined: false,
            nestedBoardCrop,
            regionRequired: false) &&
           !RDPilotApplication.ControlLoopService.CanReplaceTurnBasedInteractionRegion(
               hasExistingRegion: true,
               existingRegionIsAutomatic: true,
               automaticRegionWasRefined: true,
               nestedBoardCrop,
               regionRequired: false) &&
           !RDPilotApplication.ControlLoopService.CanReplaceTurnBasedInteractionRegion(
               hasExistingRegion: true,
               existingRegionIsAutomatic: false,
               automaticRegionWasRefined: false,
               nestedBoardCrop,
               regionRequired: false),
        "an inspection crop could replace an already established/refined primary board region");
    var region = new Rectangle(10, 10, 80, 80);
    ResolvedActionSnapshot Snapshot(string key) =>
        RDPilotApplication.ControlLoopService.AttachTurnBasedObservationRegion(
            RDPilotApplication.DesktopInputService.CaptureResolvedAction(
                new ActionDto
                {
                    Type = "keys",
                    Keys = [key],
                    Note = $"move the level block with {key}"
                },
                null),
            region,
            "turn_based_interaction");
    ScreenObservationFrame Frame(int blockX, int blockY, int timerEnd)
    {
        var detail = Enumerable.Repeat((byte)230, 10_000).ToArray();
        for (var y = blockY; y < blockY + 12; y++)
        for (var x = blockX; x < blockX + 12; x++)
            detail[y * 100 + x] = 20;
        for (var x = 15; x < timerEnd; x++)
            detail[84 * 100 + x] = 80;
        var global = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
        return new ScreenObservationFrame(
            global,
            global,
            detail,
            100,
            100,
            new Rectangle(0, 0, 100, 100));
    }
    ObservationAssessment TurnAssessment(
        ActionOutcomeState outcome,
        VisualChangeState visual) =>
        new(
            "turn_based_interaction",
            visual,
            outcome,
            GoalProgressState.Neutral,
            0.9,
            outcome == ActionOutcomeState.NoEffect ? 0 : 0.02,
            0,
            0,
            outcome == ActionOutcomeState.NoEffect ? 0 : 0.02,
            outcome == ActionOutcomeState.NoEffect ? 0 : 0.02,
            0.004,
            "turn-based test")
        {
            ActionPolicy = "turn_based_interaction"
        };

    var snapshot = Snapshot("ArrowLeft");
    Assert(snapshot.ObservationRegion == region,
        "turn-based input did not retain the board observation region");
    Assert(RDPilotApplication.ControlLoopService.TryGetTurnBasedDirectionalSequenceLength(
            new ActionDto { Type = "keys", Keys = ["ArrowLeft", "ArrowUp"] },
            out var sequenceLength) && sequenceLength == 2,
        "ordered directional key sequence was not recognized");
    Assert(RDPilotApplication.ControlLoopService.CanExpandTurnKeyActionFromPlan(
            new ActionDto { Type = "keys", Keys = ["ArrowLeft"] },
            queuedFollowUpCount: 0) &&
           !RDPilotApplication.ControlLoopService.CanExpandTurnKeyActionFromPlan(
            new ActionDto { Type = "keys", Keys = ["ArrowLeft"] },
            queuedFollowUpCount: 13),
        "the first key of an already queued route could be expanded into a duplicate full route");
    Assert(RDPilotApplication.OpenAiResponsesService.TryNormalizeDirectionalLabel(
            "Continue upward through the corridor.",
            out var upwardLabel) &&
           upwardLabel == "ArrowUp" &&
           RDPilotApplication.OpenAiResponsesService.TryNormalizeDirectionalLabel(
            "Test the downward branch.",
            out var downwardLabel) &&
           downwardLabel == "ArrowDown" &&
           RDPilotApplication.OpenAiResponsesService.TryNormalizeDirectionalLabel(
            "Move leftwards, then inspect.",
            out var leftwardLabel) &&
           leftwardLabel == "ArrowLeft" &&
           RDPilotApplication.OpenAiResponsesService.TryNormalizeDirectionalLabel(
            "Proceed rightward.",
            out var rightwardLabel) &&
           rightwardLabel == "ArrowRight",
        "natural English directional notes were not normalized to authoritative turn inputs");
    Assert(RDPilotApplication.OpenAiResponsesService.TryNormalizeUnambiguousDirectionalLabel(
               "Po zablokowanym lewo testuję ruch w górę.",
               out _) == false &&
           RDPilotApplication.OpenAiResponsesService.TryNormalizeUnambiguousDirectionalLabel(
               "Testuję ruch w górę.",
               out var unambiguousUp) &&
           unambiguousUp == "ArrowUp",
        "a note mentioning a previous direction overrode the direction being executed");
    Assert(RDPilotApplication.OpenAiResponsesService.HasImmediateDirectionalReversal(
               ["ArrowUp", "ArrowLeft", "ArrowRight", "ArrowDown"]) &&
           !RDPilotApplication.OpenAiResponsesService.HasImmediateDirectionalReversal(
               ["ArrowUp", "ArrowUp", "ArrowLeft", "ArrowLeft"]),
        "an incoherent route with an immediate reversal was accepted for batching");
    Assert(!RDPilotApplication.ControlLoopService.RequiresPrimaryTurnBasedRegion(
            new ActionDto { Type = "click", XPx = 50, YPx = 50 }),
        "an unambiguous visible control click was blocked before the primary region existed");
    Assert(!RDPilotApplication.ControlLoopService.RequiresPrimaryTurnBasedRegion(
            new ActionDto { Type = "keys", Keys = ["Space"] }),
        "a conventional start/confirm key was blocked before the primary region existed");
    Assert(RDPilotApplication.ControlLoopService.RequiresPrimaryTurnBasedRegion(
            new ActionDto { Type = "keys", Keys = ["ArrowLeft"] }),
        "a subtle directional input was allowed without a primary observation region");

    var before = Frame(56, 42, 70);
    var immediateBaseline = Frame(56, 42, 60);
    var after = Frame(44, 42, 58);
    var inferredObservationRegion =
        RDPilotApplication.ControlLoopService.InferTurnBasedObservationRegion(
            before,
            after,
            new Rectangle(0, 0, 100, 100));
    Assert(inferredObservationRegion is Rectangle inferred &&
           inferred.Width < 80 &&
           inferred.Contains(50, 48) &&
           inferred.Contains(62, 48),
        "automatic turn observation did not isolate the changed playfield while preserving the full model view");
    var context = new UiPromptContext(
        "ARC-AGI-3 Task #ls20",
        "msedge",
        "level board controls",
        null,
        null,
        null);
    var assessment = new RDPilotApplication.AdaptiveObservationSession().Assess(
        immediateBaseline,
        after,
        snapshot,
        context,
        context,
        "finite",
        "graj w grę zagadkę i przechodź kolejne poziomy");
    Assert(assessment.ActionPolicy == "turn_based_interaction",
        "turn-based input did not use turn-based observation");
    Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed &&
           assessment.GoalProgress == GoalProgressState.Neutral,
        "turn-based movement was confused with verified goal progress");

    var tracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    tracker.ObserveState(before, region);
    var ambientBaseline = tracker.PrepareActionBaseline(immediateBaseline, region);
    Assert(!ambientBaseline.ExternalStateChange,
        "ambient timer movement was mistaken for external interaction");
    tracker.RecordTransition(immediateBaseline, after, region, snapshot, assessment);
    Assert(!tracker.RequiresReanalysis,
        "a timer boundary moving just beyond its learned position created a false distant event");

    var secondBaseline = Frame(44, 42, 50);
    var thirdState = Frame(44, 20, 50);
    tracker.PrepareActionBaseline(secondBaseline, region);
    tracker.RecordTransition(
        secondBaseline,
        thirdState,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));

    var thirdBaseline = Frame(44, 20, 40);
    var returnedState = Frame(56, 42, 40);
    tracker.PrepareActionBaseline(thirdBaseline, region);
    tracker.RecordTransition(
        thirdBaseline,
        returnedState,
        region,
        Snapshot("ArrowDown"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    tracker.RecordTransition(
        returnedState,
        returnedState,
        region,
        Snapshot("ArrowRight"),
        TurnAssessment(ActionOutcomeState.NoEffect, VisualChangeState.Stable));

    var summary = tracker.BuildPromptSummary();
    Assert(summary.Contains("S1 --ArrowLeft--> S2 [changed]", StringComparison.Ordinal),
        "turn transition ledger did not distinguish the moved state");
    Assert(summary.Contains("S2 --ArrowUp--> S3 [changed]", StringComparison.Ordinal),
        "turn transition ledger collapsed a distinct state");
    Assert(summary.Contains("S3 --ArrowDown--> S1 [changed]", StringComparison.Ordinal),
        "turn transition ledger did not recognize a return to an earlier state");
    Assert(summary.Contains("S1 --ArrowRight--> S1 [no_effect]", StringComparison.Ordinal),
        "turn transition ledger did not retain a legal blocked input");
    Assert(summary.Contains("TURN_TOPOLOGY", StringComparison.Ordinal) &&
           summary.Contains("ArrowRight=blocked", StringComparison.Ordinal),
        "turn transition ledger did not expose its directed topology and blocked edges");
    Assert(tracker.CanUseExecutionBatch &&
           summary.Contains("TURN_PHASE: execution_ready", StringComparison.Ordinal) &&
           summary.Contains("TURN_HYPOTHESIS_CONTRADICTION", StringComparison.Ordinal),
        "a blocked transition did not preserve aggressive replanning while falsifying the active hypothesis");

    var replayAssessment = tracker.RecordBatchStep(
        returnedState,
        after,
        region,
        "ArrowLeft");
    Assert(replayAssessment.KnownTransition && replayAssessment.ContinueBatch,
        "a known ordinary transition could not continue an observed key sequence");

    var transientTracker =
        new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    transientTracker.ObserveState(returnedState, region);
    var transientOnlyAssessment = transientTracker.RecordBatchStep(
        returnedState,
        returnedState,
        region,
        "ArrowRight",
        actionReactionObserved: true);
    Assert(transientOnlyAssessment.NoEffect &&
           !transientOnlyAssessment.ContinueBatch &&
           transientOnlyAssessment.Summary.Contains(
               "transient_reaction=true",
               StringComparison.Ordinal),
        "a transient key animation overrode the unchanged settled turn state");

    var auxiliaryOnlyTracker =
        new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    var auxiliaryOnlyBefore = Frame(56, 42, 70);
    var auxiliaryOnlyAfter = Frame(56, 42, 55);
    auxiliaryOnlyTracker.ObserveState(auxiliaryOnlyBefore, region);
    var auxiliaryOnlyAssessment = auxiliaryOnlyTracker.RecordBatchStep(
        auxiliaryOnlyBefore,
        auxiliaryOnlyAfter,
        region,
        "ArrowUp",
        actionReactionObserved: true);
    Assert(auxiliaryOnlyAssessment.NoEffect &&
           !auxiliaryOnlyAssessment.ContinueBatch &&
           auxiliaryOnlyAssessment.Summary.Contains(
               "movement_evidence=false",
               StringComparison.Ordinal) &&
           auxiliaryOnlyAssessment.Summary.Contains(
               "auxiliary_only=true",
               StringComparison.Ordinal),
        "a small auxiliary HUD change was mistaken for movement of the controlled object");

    var hudCausalTracker =
        new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    var hudStart = Frame(56, 42, 70);
    var hudOnly = Frame(56, 42, 55);
    var movedAfterHud = Frame(44, 42, 40);
    hudCausalTracker.ObserveState(hudStart, region);
    var blockedByHud = hudCausalTracker.RecordTransition(
        hudStart,
        hudOnly,
        region,
        Snapshot("ArrowLeft"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    var movedAfterBlockedHud = hudCausalTracker.RecordTransition(
        hudOnly,
        movedAfterHud,
        region,
        Snapshot("ArrowLeft"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    var hudCausalSummary = hudCausalTracker.BuildPromptSummary();
    Assert(blockedByHud &&
           !movedAfterBlockedHud &&
           !hudCausalTracker.RequiresReanalysis &&
           hudCausalSummary.Contains(
               "S1 --ArrowLeft--> S1 [no_effect]",
               StringComparison.Ordinal) &&
           hudCausalSummary.Contains(
               "S1 --ArrowLeft--> S2 [changed]",
               StringComparison.Ordinal),
        "a recurring peripheral HUD decrement advanced the state or created a false causal event");

    var tinyRegion = new Rectangle(0, 0, 64, 64);
    ScreenObservationFrame TinyFrame(int blockX)
    {
        var detail = Enumerable.Repeat((byte)230, 64 * 64).ToArray();
        for (var y = 30; y < 32; y++)
        for (var x = blockX; x < blockX + 2; x++)
            detail[y * 64 + x] = 20;
        var global = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
        return new ScreenObservationFrame(
            global,
            global,
            detail,
            64,
            64,
            tinyRegion);
    }
    var tinyBefore = TinyFrame(24);
    var tinyAfter = TinyFrame(27);
    var tinyTracker =
        new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    tinyTracker.ObserveState(tinyBefore, tinyRegion);
    tinyTracker.RecordTransition(
        tinyBefore,
        tinyAfter,
        tinyRegion,
        Snapshot("ArrowRight"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(tinyTracker.BuildPromptSummary().Contains(
               "S1 --ArrowRight--> S2 [changed]",
               StringComparison.Ordinal),
        "a small spatially coherent directional move was aliased to its origin state");

    var corridorTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    var corridorState1 = Frame(44, 60, 70);
    var corridorState2 = Frame(44, 52, 70);
    var corridorState3 = Frame(44, 44, 70);
    var corridorState4 = Frame(44, 36, 70);
    var corridorState5 = Frame(44, 28, 70);
    corridorTracker.ObserveState(corridorState1, region);
    corridorTracker.RecordTransition(
        corridorState1,
        corridorState2,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    corridorTracker.RecordTransition(
        corridorState2,
        corridorState3,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(corridorTracker.CanUseExecutionBatch &&
           corridorTracker.BuildPromptSummary().Contains(
               "Aggressive batching is available immediately",
               StringComparison.Ordinal),
        "two predictable moves in one direction did not enable a bounded batch");
    corridorTracker.RecordTransition(
        corridorState3,
        corridorState4,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(corridorTracker.CanUseExecutionBatch,
        "an additional predictable move unexpectedly disabled batching");
    var corridorBatchStep = corridorTracker.RecordBatchStep(
        corridorState4,
        corridorState5,
        region,
        "ArrowUp");
    Assert(!corridorBatchStep.KnownTransition && corridorBatchStep.ContinueBatch,
        "a learned directional pattern stopped at every new ordinary state");
    Assert(corridorTracker.BuildPromptSummary().Contains(
            "S4 --ArrowUp--> S5 [changed]",
            StringComparison.Ordinal),
        "successive small moves oscillated between similar historical state ids");
    var corridorBlockedStep = corridorTracker.RecordBatchStep(
        corridorState5,
        corridorState5,
        region,
        "ArrowUp");
    Assert(corridorBlockedStep.NoEffect &&
           !corridorBlockedStep.ContinueBatch &&
           !corridorTracker.CanUseExecutionBatch &&
           corridorTracker.BuildPromptSummary().Contains(
               "TURN_HYPOTHESIS_CONTRADICTION",
               StringComparison.Ordinal),
        "a confirmed blocked step did not immediately interrupt the directional batch");

    var plannedRoute = new RDPilotApplication.ControlLoopService.ShortTermPlanTracker();
    plannedRoute.Update(new ActionDto
    {
        ShortTermPlan = "Move upward twice, then turn right toward the next waypoint.",
        PlanStatus = "active",
        PlannedInputs = ["ArrowUp", "ArrowUp", "ArrowRight", "ArrowRight"],
        PlanWaypoint = "the next visible corridor opening",
        PlanStateId = "S1",
        PlanConfidence = 0.60
    }, "S1");
    var plannedTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker(
        plannedRoute);
    plannedTracker.ObserveState(corridorState1, region);
    Assert(plannedTracker.CanProposeExecutionBatch &&
           plannedTracker.CanUseExecutionBatch &&
           plannedTracker.AdvertisedMaxExecutionBatchLength == 32 &&
           plannedTracker.MaxExecutionBatchLength == 32,
        "a reversible structured route with 0.60 confidence was not executable immediately after establishing the board state");
    Assert(plannedRoute.TryExpandDirectionalSequence(
               ["ArrowUp"],
               "S1",
               plannedTracker.MaxExecutionBatchLength,
               out var immediateExpandedRoute) &&
           immediateExpandedRoute.SequenceEqual(
               ["ArrowUp", "ArrowUp", "ArrowRight", "ArrowRight"]),
        "the runtime did not expand the first conservative move into the initial aggressive route");
    plannedTracker.RecordTransition(
        corridorState1,
        corridorState2,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    plannedRoute.Update(new ActionDto
    {
        ShortTermPlan = "Continue the same route toward the opening.",
        PlanStatus = "active",
        PlannedInputs = ["ArrowUp", "ArrowUp", "ArrowRight", "ArrowRight"],
        PlanWaypoint = "the next visible corridor opening",
        PlanStateId = "S2",
        PlanConfidence = 0.9
    }, "S2");
    Assert(plannedRoute.BuildPromptSummary().Contains(
               "CURRENT_PLAN_INDEX: 1/4",
               StringComparison.Ordinal),
        "retaining the same structured route reset application-owned progress");
    Assert(plannedTracker.CanUseExecutionBatch &&
           plannedTracker.MaxExecutionBatchLength == 32,
        "one confirmed high-confidence structured move lost the progressive execution batch");
    Assert(plannedRoute.TryExpandDirectionalSequence(
               ["ArrowUp"],
               "S2",
               plannedTracker.MaxExecutionBatchLength,
               out var expandedRoute) &&
           expandedRoute.SequenceEqual(["ArrowUp", "ArrowRight", "ArrowRight"]),
        "a short model batch was not expanded from the remaining structured route");
    plannedTracker.RecordTransition(
        corridorState2,
        corridorState3,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(plannedTracker.CanUseExecutionBatch &&
           plannedTracker.MaxExecutionBatchLength == 32,
        "two confirmed plan-backed moves did not retain progressive execution");
    var plannedTurn = plannedTracker.RecordBatchStep(
        corridorState3,
        corridorState4,
        region,
        "ArrowRight");
    Assert(plannedTurn.ContinueBatch &&
           plannedTurn.Summary.Contains("plan_backed=true", StringComparison.Ordinal),
        "an active plan could not safely continue through a new ordinary direction");
    var blockedPlannedTurn = plannedTracker.RecordBatchStep(
        corridorState4,
        corridorState4,
        region,
        "ArrowRight");
    Assert(blockedPlannedTurn.NoEffect &&
           !blockedPlannedTurn.ContinueBatch &&
           !plannedRoute.HasActivePlan &&
           plannedRoute.Status == "invalidated" &&
           plannedRoute.BuildPromptSummary().Contains(
               "SHORT_TERM_PLAN_REVISION_REASON",
               StringComparison.Ordinal),
        "a blocked planned input did not invalidate the persisted short-term plan");

    var longRoute = new RDPilotApplication.ControlLoopService.ShortTermPlanTracker();
    var longRouteTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker(
        longRoute);
    longRouteTracker.ObserveState(corridorState1, region);
    longRouteTracker.RecordTransition(
        corridorState1,
        corridorState2,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    longRouteTracker.RecordTransition(
        corridorState2,
        corridorState3,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    longRoute.Update(new ActionDto
    {
        ShortTermPlan = "Follow the visible route through its turn to the terminal target.",
        PlanStatus = "active",
        PlannedInputs =
        [
            "ArrowRight", "ArrowRight", "ArrowRight",
            "ArrowUp", "ArrowUp", "ArrowUp"
        ],
        PlanWaypoint = "the visible terminal target",
        PlanStateId = "S3",
        PlanConfidence = 0.95
    }, "S3");
    Assert(longRouteTracker.CanUseExecutionBatch &&
           longRouteTracker.MaxExecutionBatchLength == 32 &&
           longRouteTracker.PreferredAdvertisedExecutionBatchMinimum == 6,
        "a high-confidence semantic route did not enable an immediate extended batch");
    Assert(longRoute.TryExpandDirectionalSequence(
               ["ArrowRight"],
               "S3",
               longRouteTracker.MaxExecutionBatchLength,
               out var expandedLongRoute) &&
           expandedLongRoute.SequenceEqual(
               [
                   "ArrowRight", "ArrowRight", "ArrowRight",
                   "ArrowUp", "ArrowUp", "ArrowUp"
               ]),
        "the runtime did not expand a short proposal across a visible predictable turn");

    var twentyStepRoute =
        new RDPilotApplication.ControlLoopService.ShortTermPlanTracker();
    var twentyInputs = Enumerable.Repeat("ArrowRight", 20).ToArray();
    twentyStepRoute.Update(new ActionDto
    {
        ShortTermPlan = "Follow the long visible straight corridor.",
        PlanStatus = "active",
        PlannedInputs = twentyInputs,
        PlanWaypoint = "the far end of the corridor",
        PlanStateId = "S1",
        PlanConfidence = 0.95
    }, "S1");
    Assert(twentyStepRoute.BuildPromptSummary().Contains(
               "CURRENT_PLAN_INDEX: 0/20",
               StringComparison.Ordinal) &&
           twentyStepRoute.TryExpandDirectionalSequence(
               ["ArrowRight"],
               "S1",
               32,
               out var expandedTwentyStepRoute) &&
           expandedTwentyStepRoute.Length == 20,
        "the application silently truncated a valid structured route after twelve inputs");

    var extendedRoute = new RDPilotApplication.ControlLoopService.ShortTermPlanTracker();
    var extendedTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker(
        extendedRoute);
    extendedTracker.ObserveState(corridorState1, region);
    extendedTracker.RecordTransition(
        corridorState1,
        corridorState2,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    extendedTracker.RecordTransition(
        corridorState2,
        corridorState3,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    extendedTracker.RecordTransition(
        corridorState3,
        corridorState4,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    extendedTracker.RecordTransition(
        corridorState4,
        corridorState5,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(extendedTracker.CanProposeExecutionBatch &&
           extendedTracker.AdvertisedMaxExecutionBatchLength == 32,
        "four confirmed directional transitions did not advertise the extended input horizon");
    extendedRoute.Update(new ActionDto
    {
        ShortTermPlan = "Follow the long visible route to its terminal target.",
        PlanStatus = "active",
        PlannedInputs =
        [
            "ArrowRight", "ArrowRight", "ArrowRight", "ArrowRight",
            "ArrowDown", "ArrowDown", "ArrowDown", "ArrowDown",
            "ArrowLeft", "ArrowLeft", "ArrowLeft", "ArrowLeft"
        ],
        PlanWaypoint = "the visible terminal target",
        PlanStateId = "S5",
        PlanConfidence = 0.95
    }, "S5");
    Assert(extendedTracker.CanUseExecutionBatch &&
           extendedTracker.MaxExecutionBatchLength == 32,
        "a mature high-confidence route did not enable the extended input batch");

    var staleRoute = new RDPilotApplication.ControlLoopService.ShortTermPlanTracker();
    staleRoute.Update(new ActionDto
    {
        ShortTermPlan = "Follow a stale route.",
        PlanStatus = "active",
        PlannedInputs = ["ArrowUp", "ArrowLeft"],
        PlanWaypoint = "old waypoint",
        PlanStateId = "S9",
        PlanConfidence = 0.95
    }, "S1");
    Assert(!staleRoute.HasActivePlan && staleRoute.Status == "invalidated",
        "a structured route bound to a stale state was accepted");

    var completedRoute = new RDPilotApplication.ControlLoopService.ShortTermPlanTracker();
    completedRoute.Update(new ActionDto
    {
        ShortTermPlan = "Reach the next corner.",
        PlanStatus = "active",
        PlannedInputs = ["ArrowUp", "ArrowRight"],
        PlanWaypoint = "next corner",
        PlanStateId = "S1",
        PlanConfidence = 0.9
    }, "S1");
    completedRoute.RecordDirectionalResult("ArrowUp", "S1", "S2", changed: true);
    completedRoute.RecordDirectionalResult("ArrowRight", "S2", "S3", changed: true);
    Assert(!completedRoute.HasActivePlan &&
           completedRoute.Status == "completed" &&
           completedRoute.BuildPromptSummary().Contains(
               "CURRENT_PLAN_INDEX: 2/2",
               StringComparison.Ordinal),
        "a fully executed structured route did not complete at its waypoint");

    tracker.UpdateWorkingMemory(new ActionDto
    {
        WorldStateSummary = "A movable block is navigating a bounded board.",
        MechanicsHypothesis = "Special board elements may change persistent world state.",
        SalientChangeObservation = "Touching the white marker changed a distant reference glyph."
    });
    summary = tracker.BuildPromptSummary();
    Assert(summary.Contains("TURN_WORLD_STATE_MEMORY", StringComparison.Ordinal) &&
           summary.Contains("TURN_MECHANICS_HYPOTHESIS", StringComparison.Ordinal),
        "turn-based working memory was not returned to the next model request");
    tracker.UpdateWorkingMemory(new ActionDto
    {
        WorldStateSummary = "A movable block is navigating a bounded board.",
        MechanicsHypothesis = "The visible marker may instead be the controlled object.",
        SalientChangeObservation = "Entering the blue terminal changed the level."
    });
    summary = tracker.BuildPromptSummary();
    Assert(summary.Contains("TURN_PRIOR_MECHANICS_HYPOTHESES", StringComparison.Ordinal) &&
           summary.Contains("Special board elements", StringComparison.Ordinal),
        "a revised mechanics claim erased the prior unverified hypothesis");
    tracker.BeginExternalStateEpoch();
    summary = tracker.BuildPromptSummary();
    Assert(!summary.Contains("TURN_MECHANICS_HYPOTHESIS", StringComparison.Ordinal) &&
           !summary.Contains("TURN_PRIOR_MECHANICS_HYPOTHESES", StringComparison.Ordinal) &&
           summary.Contains("TURN_TRANSIENT_CONTEXT_WARNING", StringComparison.Ordinal) &&
           summary.Contains("TURN_PRIOR_EPOCH_OBSERVED_EVIDENCE", StringComparison.Ordinal) &&
           !summary.Contains("TURN_PRIOR_EPOCH_HYPOTHESES", StringComparison.Ordinal) &&
           !summary.Contains("Special board elements", StringComparison.Ordinal) &&
           !summary.Contains("visible marker", StringComparison.Ordinal) &&
           !summary.Contains("white marker changed", StringComparison.Ordinal) &&
           summary.Contains("blue terminal changed", StringComparison.Ordinal),
        "an external state boundary retained stale topology/hypotheses or discarded the latest causal fact");
    tracker.UpdateWorkingMemory(new ActionDto
    {
        WorldStateSummary = "The stable next board is now visible."
    });
    Assert(!tracker.BuildPromptSummary().Contains(
            "TURN_TRANSIENT_CONTEXT_WARNING",
            StringComparison.Ordinal),
        "the transient-context warning survived a committed stable-state update");

    using var referenceCrop = new Bitmap(48, 36);
    using (var graphics = Graphics.FromImage(referenceCrop))
        graphics.Clear(Color.FromArgb(40, 80, 120));
    using var matchingNativeCrop = new Bitmap(96, 72);
    using (var graphics = Graphics.FromImage(matchingNativeCrop))
        graphics.Clear(Color.FromArgb(40, 80, 120));
    using var changedNativeCrop = new Bitmap(96, 72);
    using (var graphics = Graphics.FromImage(changedNativeCrop))
        graphics.Clear(Color.FromArgb(220, 30, 20));
    Assert(RDPilotApplication.ScreenshotService.AreTurnRegionCapturesConsistent(
               referenceCrop,
               matchingNativeCrop) &&
           !RDPilotApplication.ScreenshotService.AreTurnRegionCapturesConsistent(
               referenceCrop,
               changedNativeCrop),
        "native board crops were not checked against the primary screenshot before use");
    Assert(RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker
               .IsBoundedAttemptResetEvidence(
                   returnedToAttemptOrigin: true,
                   directionalInputs: 50,
                   broadChange: true,
                   distantChange: true,
                   novelRegionCount: 3,
                   auxiliaryRegionCount: 1) &&
           !RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker
               .IsBoundedAttemptResetEvidence(
                   returnedToAttemptOrigin: true,
                   directionalInputs: 20,
                   broadChange: true,
                   distantChange: false,
                   novelRegionCount: 1,
                   auxiliaryRegionCount: 2),
        "a closed route was confused with broad bounded-attempt reset evidence");

    ScreenObservationFrame WithRemoteChange(
        ScreenObservationFrame source,
        int left,
        int top)
    {
        var detail = source.DetailFingerprint.ToArray();
        for (var y = top; y < top + 10; y++)
        for (var x = left; x < left + 10; x++)
            detail[y * 100 + x] = 20;
        return source with { DetailFingerprint = detail };
    }

    var eventTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    var eventBefore = Frame(56, 42, 60);
    var eventAfter = WithRemoteChange(Frame(44, 42, 60), 18, 72);
    eventTracker.ObserveState(eventBefore, region);
    eventTracker.RecordTransition(
        eventBefore,
        eventAfter,
        region,
        Snapshot("ArrowLeft"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    var eventSummary = eventTracker.BuildPromptSummary();
    Assert(!eventTracker.RequiresReanalysis &&
           eventTracker.CanProposeExecutionBatch &&
           eventSummary.Contains("TURN_REANALYSIS_REQUIRED: false", StringComparison.Ordinal) &&
           eventSummary.Contains("distant=true", StringComparison.Ordinal),
        "a small distant change interrupted aggressive execution instead of remaining visible evidence");
    Assert(eventTracker.HasRequiredSalientObservation(new ActionDto
           {
               Type = "keys",
               Keys = ["ArrowUp"]
           }) &&
           eventTracker.HasRequiredSalientObservation(new ActionDto
           {
               Type = "keys",
               Keys = ["ArrowUp"],
               SalientChangeObservation = "A separate persistent board element changed shape."
           }),
        "a non-broad distant change still imposed a salient-observation barrier");
    eventTracker.AcknowledgeReanalysis();
    Assert(!eventTracker.RequiresReanalysis,
        "the reanalysis barrier did not clear after a fresh one-step decision");

    var auxiliaryTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    var auxiliaryState1 = Frame(56, 42, 70);
    var auxiliaryState2 = Frame(44, 42, 58);
    var auxiliaryState3 = Frame(32, 42, 50);
    auxiliaryTracker.ObserveState(auxiliaryState1, region);
    auxiliaryTracker.RecordTransition(
        auxiliaryState1,
        auxiliaryState2,
        region,
        Snapshot("ArrowLeft"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    auxiliaryTracker.AcknowledgeReanalysis();
    auxiliaryTracker.RecordTransition(
        auxiliaryState2,
        auxiliaryState3,
        region,
        Snapshot("ArrowLeft"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    var auxiliarySummary = auxiliaryTracker.BuildPromptSummary();
    Assert(!auxiliaryTracker.RequiresReanalysis &&
           auxiliarySummary.Contains("TURN_AUXILIARY_CHANGES", StringComparison.Ordinal) &&
           auxiliarySummary.Contains("predictable_regions=", StringComparison.Ordinal),
        "a recurring action-correlated auxiliary change kept forcing reanalysis");

    var auxiliaryReturn = Frame(44, 42, 42);
    auxiliaryTracker.RecordTransition(
        auxiliaryState3,
        auxiliaryReturn,
        region,
        Snapshot("ArrowRight"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    auxiliarySummary = auxiliaryTracker.BuildPromptSummary();
    Assert(auxiliarySummary.Contains(
               "S3 --ArrowRight--> S2 [changed] [returned_to_known_state]",
               StringComparison.Ordinal) &&
           auxiliarySummary.Contains("TURN_STATE_RETURN", StringComparison.Ordinal) &&
           auxiliarySummary.Contains("TURN_AUXILIARY_SIGNAL", StringComparison.Ordinal) &&
           auxiliarySummary.Contains(
               "TURN_NAVIGATION_POSE: state=S2; relative=(-1,0)",
               StringComparison.Ordinal) &&
           auxiliarySummary.Contains(
               "(-2,0) S3: ArrowRight->(-1,0) S2",
               StringComparison.Ordinal),
        "a changed peripheral indicator prevented recognition of a return to the same logical board state");

    var novelWorldState = WithRemoteChange(Frame(20, 42, 42), 70, 18);
    auxiliaryTracker.RecordTransition(
        auxiliaryReturn,
        novelWorldState,
        region,
        Snapshot("ArrowLeft"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(auxiliaryTracker.RequiresReanalysis &&
           auxiliaryTracker.SalientChangeRegions.Count > 0 &&
           auxiliaryTracker.BuildPromptSummary().Contains(
               "causal_distant=true",
               StringComparison.Ordinal),
        "a novel local-to-distant causal change did not interrupt execution for reanalysis");

    var broadWorldState = novelWorldState with
    {
        DetailFingerprint = Enumerable.Repeat((byte)15, 10_000).ToArray()
    };
    auxiliaryTracker.RecordTransition(
        novelWorldState,
        broadWorldState,
        region,
        Snapshot("ArrowUp"),
        TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
    Assert(auxiliaryTracker.RequiresReanalysis &&
           auxiliaryTracker.SalientChangeRegions.Count > 0,
        "a broad screen transition did not stop aggressive execution");

    var externalTracker = new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    externalTracker.ObserveState(before, region);
    var externalBaseline = externalTracker.PrepareActionBaseline(after, region);
    Assert(externalBaseline.ExternalStateChange && externalTracker.RequiresReanalysis,
        "a state-changing input during model latency was learned as ambient motion");

    ScreenObservationFrame WithLowIntensityAnimation(
        ScreenObservationFrame source,
        params Rectangle[] animatedRegions)
    {
        var detail = source.DetailFingerprint.ToArray();
        foreach (var animatedRegion in animatedRegions)
        for (var y = animatedRegion.Top; y < animatedRegion.Bottom; y++)
        for (var x = animatedRegion.Left; x < animatedRegion.Right; x++)
            detail[y * 100 + x] = (byte)Math.Max(0, detail[y * 100 + x] - 18);
        return source with { DetailFingerprint = detail };
    }

    var animatedStartTracker =
        new RDPilotApplication.ControlLoopService.TurnBasedTransitionTracker();
    var animatedStart = Frame(56, 42, 70);
    var animatedStartDrift = WithLowIntensityAnimation(
        animatedStart,
        new Rectangle(38, 30, 14, 8),
        new Rectangle(50, 38, 14, 8),
        new Rectangle(36, 44, 12, 8));
    animatedStartTracker.ObserveState(animatedStart, region);
    var animatedStartBaseline = animatedStartTracker.PrepareActionBaseline(
        animatedStartDrift,
        region);
    Assert(!animatedStartBaseline.ExternalStateChange &&
           !animatedStartTracker.RequiresReanalysis &&
           animatedStartBaseline.Summary.Contains(
               "cold_start_ambient_calibration",
               StringComparison.Ordinal),
        "low-intensity startup animation invalidated a still-relevant action");
    var repeatedAnimatedStartBaseline = animatedStartTracker.PrepareActionBaseline(
        animatedStart,
        region);
    Assert(!repeatedAnimatedStartBaseline.ExternalStateChange &&
           !animatedStartTracker.RequiresReanalysis,
        "learned startup animation repeatedly invalidated a still-relevant action");

    Assert(RDPilotApplication.ControlLoopService.IsOverlappingTurnInspection(
            new Rectangle(100, 100, 400, 400),
            new Rectangle(50, 50, 500, 500)),
        "a containing turn inspection crop was not treated as transient");
    Assert(!RDPilotApplication.ControlLoopService.IsOverlappingTurnInspection(
            new Rectangle(100, 100, 200, 200),
            new Rectangle(350, 100, 200, 200)),
        "a separate board region was incorrectly treated as the same interaction crop");
}

static void TestAdaptiveColorObservation()
{
    RDPilotApplication.ConfigurationService.ApplyObservationMode("auto", "self-test");
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
    var snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto
        {
            Type = "drag_path",
            GestureKind = "draw",
            DurationMs = 500,
            Path =
            [
                new GesturePointDto { XPx = 10, YPx = 50 },
                new GesturePointDto { XPx = 90, YPx = 50 }
            ]
        },
        null);
    var grayscale = Enumerable.Repeat((byte)128, 10_000).ToArray();
    var beforeColor = Enumerable.Repeat((byte)128, 30_000).ToArray();
    var afterColor = beforeColor.ToArray();
    for (var x = 10; x <= 90; x++)
    {
        var index = (50 * 100 + x) * 3;
        afterColor[index] = 255;
        afterColor[index + 1] = 32;
        afterColor[index + 2] = 128;
    }
    var global = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
    var before = new ScreenObservationFrame(
        global,
        global,
        grayscale,
        100,
        100,
        new Rectangle(0, 0, 100, 100))
    {
        DetailColorFingerprint = beforeColor
    };
    var after = new ScreenObservationFrame(
        global.ToArray(),
        global.ToArray(),
        grayscale.ToArray(),
        100,
        100,
        new Rectangle(0, 0, 100, 100))
    {
        DetailColorFingerprint = afterColor
    };
    var context = new UiPromptContext("Paint", "mspaint", "canvas", null, null, null);
    var assessment = new RDPilotApplication.AdaptiveObservationSession().Assess(
        before,
        after,
        snapshot,
        context,
        context,
        "finite");
    Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed, "color-only path change was missed");
}

static void TestFocusedTextObservation()
{
    RDPilotApplication.ConfigurationService.ApplyObservationMode("auto", "self-test");
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
    var snapshot = RDPilotApplication.ControlLoopService.AttachFocusedTextObservationRegion(
        RDPilotApplication.DesktopInputService.CaptureResolvedAction(
            new ActionDto { Type = "type_text", Text = "1100" },
            null),
        new Rectangle(35, 40, 30, 12));
    Assert(snapshot.ObservationRegion is not null, "focused text action did not receive a local observation region");
    Assert(RDPilotApplication.ControlLoopService.ShouldDeferFocusedTextStagnation(snapshot, true, 0),
        "first visually ambiguous focused text edit does not receive a grace attempt");
    Assert(!RDPilotApplication.ControlLoopService.ShouldDeferFocusedTextStagnation(snapshot, true, 1),
        "repeated ineffective text edits receive unlimited stagnation grace");

    var beforeDetail = Enumerable.Repeat((byte)230, 10_000).ToArray();
    var afterDetail = beforeDetail.ToArray();
    for (var y = 44; y < 49; y++)
    for (var x = 42; x < 58; x++)
        afterDetail[y * 100 + x] = 20;
    var stable = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
    var before = new ScreenObservationFrame(
        stable,
        stable,
        beforeDetail,
        100,
        100,
        new Rectangle(0, 0, 100, 100));
    var after = new ScreenObservationFrame(
        stable.ToArray(),
        stable.ToArray(),
        afterDetail,
        100,
        100,
        new Rectangle(0, 0, 100, 100));
    var context = new UiPromptContext(
        "Properties",
        "mspaint",
        "textbox edit",
        null,
        null,
        null);
    var assessment = new RDPilotApplication.AdaptiveObservationSession().Assess(
        before,
        after,
        snapshot,
        context,
        context,
        "finite");
    Assert(assessment.ActionPolicy == "local_editing", "focused text edit did not select local observation");
    Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed &&
           assessment.GoalProgress == GoalProgressState.Progress,
        "small focused text change was treated as stagnation");
}

static void TestRealtimeAmbientMotion()
{
    RDPilotApplication.ConfigurationService.ApplyObservationMode("auto", "self-test");
    var zero = new byte[96 * 54];
    var four = Enumerable.Repeat((byte)4, 96 * 54).ToArray();
    var eight = Enumerable.Repeat((byte)8, 96 * 54).ToArray();
    var twelve = Enumerable.Repeat((byte)12, 96 * 54).ToArray();
    var detail = new byte[100];
    ScreenObservationFrame Frame(byte[] fingerprint) => new(
        fingerprint,
        fingerprint,
        detail,
        10,
        10,
        new Rectangle(0, 0, 100, 100));
    var action = new ResolvedActionSnapshot(
        new ActionDto { Type = "keys", Keys = ["space"] },
        "keys [space]",
        "keys:space",
        null);
    var context = new UiPromptContext("Game", "game", "animated canvas", null, null, null);
    var transientSession = new RDPilotApplication.AdaptiveObservationSession();
    transientSession.RecordAmbientMotion(Frame(zero), Frame(four));
    var transientAssessment = transientSession.Assess(
        Frame(four),
        Frame(eight),
        action,
        context,
        context,
        "finite");
    Assert(transientAssessment.Profile != "realtime_interaction",
        "a single transient animation immediately selected realtime observation");

    var session = new RDPilotApplication.AdaptiveObservationSession();
    session.RecordAmbientMotion(Frame(zero), Frame(four));
    session.RecordAmbientMotion(Frame(four), Frame(eight));
    var assessment = session.Assess(
        Frame(eight),
        Frame(twelve),
        action,
        context,
        context,
        "finite");
    Assert(assessment.Profile == "realtime_interaction", "ambient motion did not select realtime observation");
    Assert(assessment.GoalProgress != GoalProgressState.Progress, "ambient animation was treated as goal progress");

    var shellSession = new RDPilotApplication.AdaptiveObservationSession();
    shellSession.RecordAmbientMotion(Frame(zero), Frame(four));
    shellSession.RecordAmbientMotion(Frame(four), Frame(eight));
    var shellAction = new ResolvedActionSnapshot(
        new ActionDto { Type = "keys", Keys = ["enter"], Note = "open selected application" },
        "keys [enter]",
        "keys:enter",
        null);
    var shellContext = new UiPromptContext(
        "Search",
        "explorer",
        "search results",
        null,
        null,
        null);
    var shellAssessment = shellSession.Assess(
        Frame(eight),
        Frame(twelve),
        shellAction,
        shellContext,
        shellContext,
        "finite");
    Assert(shellAssessment.Profile != "realtime_interaction",
        "animated shell UI incorrectly selected realtime observation");
    Assert(shellAssessment.ActionPolicy == "static_ui",
        "discrete shell action lost settle observation while the UI was animated");
}

static void TestPathGestureLoopClassification()
{
    RDPilotApplication.DesktopInputService.SetCurrentScreenMap(800, 600, 800, 600);
    var snapshot = RDPilotApplication.DesktopInputService.CaptureResolvedAction(
        new ActionDto
        {
            Type = "drag_path",
            GestureKind = "draw",
            DurationMs = 500,
            Note = "draw a cat outline on canvas",
            Path =
            [
                new GesturePointDto { XPx = 100, YPx = 100 },
                new GesturePointDto { XPx = 200, YPx = 200 }
            ]
        },
        null);
    var kind = RDPilotApplication.RecoveryMemoryService.ClassifyLoopKind(
        [snapshot],
        snapshot,
        new UiPromptContext("Paint", "mspaint", "canvas drawing surface", null, null, null));
    Assert(
        kind == RDPilotApplication.RecoveryMemoryService.RasterCanvasPointerLoop,
        $"draw path was classified as {kind}");
}

static void TestRejectedProposalCycles()
{
    Assert(
        RDPilotApplication.ControlLoopService.RepeatedStringCycleLength(
            ["done", "done"]) == 1,
        "direct rejected-proposal repeat was missed");
    Assert(
        RDPilotApplication.ControlLoopService.RepeatedStringCycleLength(
            ["click:a", "keys:tab", "click:a", "keys:tab"]) == 2,
        "multi-step rejected-proposal cycle was missed");
    Assert(
        RDPilotApplication.ControlLoopService.RepeatedStringCycleLength(
            ["click:a", "keys:tab", "scroll:down"]) == 0,
        "non-cyclic proposal history was marked as a loop");
}

static void TestRejectedProposalResetPolicy()
{
    var observation = Snapshot("request_crop", new Point(100, 100));
    var mutation = Snapshot("click", new Point(100, 100));
    Assert(
        !RDPilotApplication.ControlLoopService
            .ShouldResetRejectedProposalLoop(
                observation,
                noChange: false,
                expectedContinuousIdle: false),
        "an observation-only action cleared rejected proposal history");
    Assert(
        !RDPilotApplication.ControlLoopService
            .ShouldResetRejectedProposalLoop(
                mutation,
                noChange: true,
                expectedContinuousIdle: false),
        "an ineffective action cleared rejected proposal history");
    Assert(
        RDPilotApplication.ControlLoopService
            .ShouldResetRejectedProposalLoop(
                mutation,
                noChange: false,
                expectedContinuousIdle: false),
        "observed progress did not clear rejected proposal history");
}

static void TestRepeatAccounting()
{
    var points = new List<ResolvedActionSnapshot>();
    var successful = Snapshot("click", new Point(100, 100));
    var state = RDPilotApplication.ControlLoopService.UpdateRepeatDetection(
        successful,
        noChange: false,
        repeatCount: 0,
        lastSignature: null,
        points);
    Assert(points.Count == 0 && state.RepeatCount == 0, "successful click polluted ineffective history");

    state = RDPilotApplication.ControlLoopService.UpdateRepeatDetection(
        successful,
        noChange: true,
        state.RepeatCount,
        state.LastSignature,
        points);
    Assert(state.RepeatCount == 0, "first ineffective click was called a repeat");

    var nearby = Snapshot("click", new Point(102, 101));
    state = RDPilotApplication.ControlLoopService.UpdateRepeatDetection(
        nearby,
        noChange: true,
        state.RepeatCount,
        state.LastSignature,
        points);
    Assert(state.RepeatCount == 1, "second nearby ineffective click was not counted");

    ObservationAssessment NoEffectAssessment(string policy) => new(
        policy,
        VisualChangeState.Stable,
        ActionOutcomeState.NoEffect,
        GoalProgressState.NoProgress,
        0.82,
        0,
        0,
        0,
        double.NaN,
        0,
        0.004,
        "no effect test")
    {
        ActionPolicy = policy
    };
    var staticNoEffect = NoEffectAssessment("static_ui");
    var turnNoEffect = NoEffectAssessment("turn_based_interaction");

    Assert(
        RDPilotApplication.ControlLoopService.ShouldRegisterImmediateNoEffectCooldown(
            new ActionDto { Type = "keys", Keys = ["SPACE"] },
            staticNoEffect,
            repeatCount: 0),
        "the first unambiguous ineffective key action was not eligible for cooldown");
    Assert(
        !RDPilotApplication.ControlLoopService.ShouldRegisterImmediateNoEffectCooldown(
            new ActionDto { Type = "click", X = 0.5, Y = 0.5 },
            turnNoEffect,
            repeatCount: 0) &&
        RDPilotApplication.ControlLoopService.ShouldRegisterImmediateNoEffectCooldown(
            new ActionDto { Type = "click", X = 0.5, Y = 0.5 },
            turnNoEffect,
            repeatCount: 1),
        "turn input cooldown does not require a repeated confirmed failure");
    Assert(
        !RDPilotApplication.ControlLoopService.ShouldRegisterImmediateNoEffectCooldown(
            new ActionDto { Type = "wait", WaitSeconds = 1 },
            staticNoEffect,
            repeatCount: 0) &&
        !RDPilotApplication.ControlLoopService.ShouldRegisterImmediateNoEffectCooldown(
            new ActionDto { Type = "type_text", Text = "test" },
            staticNoEffect,
            repeatCount: 0),
        "wait or focused text handling was incorrectly replaced by the immediate action cooldown");

    var exactCooldowns = new Dictionary<string, int>();
    var spatialCooldowns = new List<SpatialActionCooldown>();
    RDPilotApplication.ControlLoopService.RegisterActionCooldown(
        successful,
        untilStep: 5,
        exactCooldowns,
        spatialCooldowns,
        clusterSpatially: false);
    Assert(
        RDPilotApplication.ControlLoopService.IsActionOnCooldown(
            successful,
            step: 2,
            exactCooldowns,
            spatialCooldowns,
            out _),
        "the exact first ineffective click was not blocked");
    Assert(
        !RDPilotApplication.ControlLoopService.IsActionOnCooldown(
            nearby,
            step: 2,
            exactCooldowns,
            spatialCooldowns,
            out _),
        "an immediate exact cooldown incorrectly blocked a nearby alternative target");
}

static void TestDoneIsNotLearned()
{
    var queue = new Queue<ResolvedActionSnapshot>();
    var episode = new RecoveryEpisodeState();
    var done = new ResolvedActionSnapshot(
        new ActionDto { Type = "done" },
        "done",
        "done",
        null);
    RDPilotApplication.RecoveryMemoryService.RecordRecoveryAction(
        episode,
        queue,
        done,
        Array.Empty<RecoveryLesson>());
    Assert(queue.Count == 0, "done entered recent action history");
    Assert(episode.RecoveryActions.Count == 0, "done entered the learned recovery strategy");
}

static void TestRecoveryActionHistoryBound()
{
    var root = typeof(RDPilotApplication);
    var limitField = root.GetField(
        "RuntimeRecoveryActionLimit",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var original = limitField.GetValue(null);
    try
    {
        limitField.SetValue(null, 8);
        var episode = new RecoveryEpisodeState();
        var queue = new Queue<ResolvedActionSnapshot>();
        for (var index = 0; index < 24; index++)
        {
            RDPilotApplication.RecoveryMemoryService.RecordRecoveryAction(
                episode,
                queue,
                Snapshot("click", new Point(index * 10, 10)),
                Array.Empty<RecoveryLesson>());
        }
        Assert(
            episode.RecoveryActions.Count == 8,
            "recovery action history exceeded its runtime limit");
    }
    finally
    {
        limitField.SetValue(null, original);
    }
}

static void TestSpatialCycleDetection()
{
    var actions = new List<ResolvedActionSnapshot>
    {
        Snapshot("click", new Point(95, 100)),
        Snapshot("click", new Point(400, 100)),
        Snapshot("click", new Point(97, 101)),
        Snapshot("click", new Point(401, 102))
    };
    Assert(
        RDPilotApplication.RecoveryMemoryService.HasEmergingLoopPattern(actions, null),
        "nearby points on opposite grid sides should form the same spatial cycle");
}

static void TestStateGraphCycle()
{
    var graph = new LoopStateGraph();
    var context = new UiPromptContext("\"Puzzle\"", "game", "canvas board", null, "0,0 800x600", null);
    var actions = new Queue<ResolvedActionSnapshot>();
    var states = new[]
    {
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10), Fingerprint(100), Fingerprint(200), Fingerprint(10)
    };
    LoopDetectionAssessment? assessment = null;
    for (var i = 0; i < states.Length; i++)
    {
        var previous = i == 0 ? null : Snapshot("click", new Point(100 + i * 10, 100));
        if (previous is not null)
        {
            actions.Enqueue(previous);
            while (actions.Count > 12) actions.Dequeue();
        }
        assessment = RDPilotApplication.RecoveryMemoryService.AssessVisualStateCycle(
            graph,
            states[i],
            states[i],
            context,
            i + 1,
            actions,
            previous,
            i == 0 ? double.NaN : 0.1,
            recordLearning: false);
    }
    Assert(assessment?.GraphCycle == true, "graph cycle was not found");
    Assert(assessment!.IsLoop, "repeated A-B-C-A graph cycle did not cross the confidence threshold");
    Assert(assessment.CycleLength == 3, "unexpected graph cycle length");
}

static void TestIndependentGraphConfirmation()
{
    var graph = new LoopStateGraph();
    var context = new UiPromptContext("\"Puzzle\"", "game", "canvas board", null, null, null);
    var actions = new Queue<ResolvedActionSnapshot>();
    var states = new[]
    {
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10), Fingerprint(100), Fingerprint(200), Fingerprint(10)
    };
    LoopDetectionAssessment? firstDetected = null;
    LoopDetectionAssessment? independentlyConfirmed = null;
    var trace = new List<string>();
    for (var index = 0; index < states.Length; index++)
    {
        var previous = index == 0
            ? null
            : Snapshot("click", new Point(100 + index * 20, 100));
        if (previous is not null)
            actions.Enqueue(previous);
        var assessment = RDPilotApplication.RecoveryMemoryService.AssessVisualStateCycle(
            graph,
            states[index],
            states[index],
            context,
            index + 1,
            actions,
            previous,
            index == 0 ? double.NaN : 0.1,
            recordLearning: false);
        trace.Add(
            $"{index + 1}: loop={assessment.IsLoop}, independent={assessment.IndependentlyConfirmed}, " +
            $"confidence={assessment.Confidence:0.00}, topology={assessment.LoopTopology}, " +
            $"key={assessment.CalibrationKey}");
        if (assessment.IsLoop && firstDetected is null)
            firstDetected = assessment;
        if (assessment.IndependentlyConfirmed)
            independentlyConfirmed = assessment;
    }
    Assert(firstDetected is not null, "loop was never detected");
    Assert(!firstDetected!.IndependentlyConfirmed, "first threshold crossing was used as its own label");
    Assert(
        independentlyConfirmed is not null,
        $"later recurrence did not independently confirm the candidate ({string.Join("; ", trace)})");
}

static void TestProductiveContinuousCycle()
{
    var graph = new LoopStateGraph();
    var context = new UiPromptContext(
        "Operations dashboard",
        "msedge",
        "status monitor",
        null,
        null,
        null);
    var actions = new Queue<ResolvedActionSnapshot>();
    var states = new[]
    {
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10), Fingerprint(100), Fingerprint(200),
        Fingerprint(10)
    };
    LoopDetectionAssessment? assessment = null;
    for (var index = 0; index < states.Length; index++)
    {
        ResolvedActionSnapshot? previous = null;
        if (index > 0)
        {
            var type = index % 2 == 0 ? "request_crop" : "wait";
            previous = Snapshot(
                type,
                new Point(100 + index, 100));
            previous = previous with
            {
                SemanticTokens = "monitor status refresh"
            };
            actions.Enqueue(previous);
        }
        assessment =
            RDPilotApplication.RecoveryMemoryService
                .AssessVisualStateCycle(
                    graph,
                    states[index],
                    states[index],
                    context,
                    index + 1,
                    actions,
                    previous,
                    index == 0 ? double.NaN : 0.1,
                    recordLearning: false,
                    goalMode: "continuous",
                    recurringWorkflowIntent: true);
    }

    Assert(
        assessment?.IsProductiveCycle == true,
        "goal-aligned recurring workflow was not classified as productive");
    Assert(
        assessment!.IsLoop == false,
        "productive recurring workflow was treated as a harmful loop");
    Assert(
        assessment.CycleDisposition == "productive",
        "productive recurring workflow has the wrong disposition");

    var inspectionActions = new Queue<ResolvedActionSnapshot>();
    var openHelp = Snapshot("click", new Point(400, 400)) with
    {
        SemanticTokens = "open help instructions"
    };
    var inspectHelp = Snapshot("request_crop", new Point(420, 420)) with
    {
        SemanticTokens = "inspect instruction details"
    };
    var returnFromHelp = Snapshot("click", new Point(100, 100)) with
    {
        SemanticTokens = "return back to prior workspace"
    };
    inspectionActions.Enqueue(openHelp);
    inspectionActions.Enqueue(inspectHelp);
    inspectionActions.Enqueue(returnFromHelp);
    Assert(
        RDPilotApplication.RecoveryMemoryService.IsCompletedInspectionRoundTrip(
            inspectionActions,
            returnFromHelp),
        "a completed help inspection was not recognized as a productive round-trip");
}

static void TestGraphCandidateExpiry()
{
    var root = typeof(RDPilotApplication);
    var ttlField = root.GetField(
        "GraphCandidateTtlSteps",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var original = ttlField.GetValue(null);
    try
    {
        ttlField.SetValue(null, 2);
        var graph = new LoopStateGraph
        {
            PendingCandidateStep = 1,
            PendingCandidateCalibrationKey = "old-context",
            PendingCandidateWasActionable = true
        };
        _ = RDPilotApplication.RecoveryMemoryService.AssessVisualStateCycle(
            graph,
            Fingerprint(10),
            Fingerprint(10),
            new UiPromptContext(
                "Window",
                "test",
                "focus",
                null,
                null,
                null),
            5,
            Array.Empty<ResolvedActionSnapshot>(),
            null,
            double.NaN,
            recordLearning: false);
        Assert(
            graph.PendingCandidateStep is null,
            "stale graph candidate did not expire");
    }
    finally
    {
        ttlField.SetValue(null, original);
    }
}

static void TestGraphCandidateInconclusiveExpiry()
{
    var key = $"expiry-test-{Guid.NewGuid():N}";
    var memoryType =
        typeof(RDPilotApplication.RecoveryMemoryService);
    var calibrationField = memoryType.GetField(
        "RecoveryCalibration",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var calibration =
        (System.Collections.IDictionary)calibrationField.GetValue(null)!;
    var graph = new LoopStateGraph
    {
        PendingCandidateStep = 1,
        PendingCandidateConfidence = 0.65,
        PendingCandidateCalibrationKey = key,
        PendingCandidateTopology = "state-return",
        PendingCandidateDomain = "generic-ui",
        PendingCandidateWasActionable = true
    };
    try
    {
        _ = RDPilotApplication.RecoveryMemoryService
            .AssessVisualStateCycle(
                graph,
                Fingerprint(10),
                Fingerprint(10),
                new UiPromptContext(
                    "Window",
                    "test",
                    "focus",
                    null,
                    null,
                    null),
                100,
                Array.Empty<ResolvedActionSnapshot>(),
                null,
                double.NaN,
                recordLearning: true);
        var bucket = (LoopCalibrationBucket)calibration[key]!;
        Assert(
            bucket.InconclusiveCount == 1,
            "expired candidate was not counted as inconclusive");
        Assert(
            bucket.RejectedCount == 0,
            "expired candidate was incorrectly counted as rejected");
    }
    finally
    {
        calibration.Remove(key);
    }
}

static void TestRuntimeStateBounds()
{
    var root = typeof(RDPilotApplication);
    var limitField = root.GetField(
        "RuntimeSemanticStateLimit",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var edgeLimitField = root.GetField(
        "RuntimeGraphEdgeLimit",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var original = limitField.GetValue(null);
    var originalEdgeLimit = edgeLimitField.GetValue(null);
    try
    {
        limitField.SetValue(null, 32);
        edgeLimitField.SetValue(null, 32);
        var graph = new LoopStateGraph();
        for (var step = 1; step <= 80; step++)
        {
            _ = RDPilotApplication.RecoveryMemoryService.AssessVisualStateCycle(
                graph,
                Fingerprint((byte)step),
                Fingerprint((byte)step),
                new UiPromptContext(
                    $"Window {step}",
                    "test",
                    $"focus {step}",
                    null,
                    null,
                    null),
                step,
                Array.Empty<ResolvedActionSnapshot>(),
                null,
                0.1,
                recordLearning: false);
        }
        Assert(
            graph.SemanticVisitSteps.Count <= 32,
            "semantic state history exceeded its runtime limit");
        for (var index = 0; index < 100; index++)
        {
            graph.Edges.Add(new LoopStateEdge
            {
                FromNodeId = index % 5,
                ToNodeId = (index + 1) % 5,
                ActionKey = $"action-{index}",
                LastSeenStep = index,
                TraversalCount = 1
            });
        }
        InvokePrivate(
            typeof(RDPilotApplication.RecoveryMemoryService),
            "PruneGraphEdges",
            graph);
        Assert(
            graph.Edges.Count <= 32,
            "graph edges exceeded their runtime limit");
    }
    finally
    {
        limitField.SetValue(null, original);
        edgeLimitField.SetValue(null, originalEdgeLimit);
    }
}

static void TestBanditRanking()
{
    var reliable = new RecoveryLesson
    {
        SuccessCount = 20,
        FailureCount = 1,
        UpdatedUtc = DateTime.UtcNow
    };
    var unreliable = new RecoveryLesson
    {
        SuccessCount = 1,
        FailureCount = 10,
        UpdatedUtc = DateTime.UtcNow
    };
    var reliableScore = RDPilotApplication.RecoveryMemoryService.ContextualBanditScore(reliable, 0.8);
    var unreliableScore = RDPilotApplication.RecoveryMemoryService.ContextualBanditScore(unreliable, 0.8);
    Assert(reliableScore > unreliableScore, "reliability did not affect selection score");
}

static void TestGoalAwareSimilarity()
{
    var fingerprint = Fingerprint(40);
    var lesson = new RecoveryLesson
    {
        ActiveProcess = "game",
        GoalTokens = "blue tile left slot",
        WindowTitleTokens = "puzzle",
        FocusTokens = "canvas board",
        TriggerActionFamily = "click",
        LoopKind = RDPilotApplication.RecoveryMemoryService.PointerRegionLoop,
        LoopTopology = "direct-repeat",
        InteractionDomain = "raster-canvas",
        ScreenFingerprintBase64 = Convert.ToBase64String(fingerprint),
        ActiveWindowFingerprintBase64 = Convert.ToBase64String(fingerprint)
    };
    var context = new UiPromptContext("\"Puzzle\"", "game", "canvas board", null, null, null);
    var score = RDPilotApplication.RecoveryMemoryService.LessonSimilarity(
        lesson,
        context,
        fingerprint,
        fingerprint,
        "click",
        lesson.LoopKind,
        lesson.LoopTopology,
        lesson.InteractionDomain,
        "delete red card",
        "game",
        "delete",
        "finite");
    Assert(score == 0, "unrelated goal should reject the lesson");
}

static void TestPromptHistory()
{
    var root = Path.Combine(
        Path.GetTempPath(),
        "RDPilotSelfTests",
        Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(root);
    var path = Path.Combine(root, "prompt-history.json");
    try
    {
        var entries = RDPilotApplication.PromptHistoryService.NormalizeEntries(
            ["first", "second", "first", " third "]);
        Assert(entries.SequenceEqual(["second", "first", "third"]),
            "history did not keep unique prompts in most-recent order");
        RDPilotApplication.PromptHistoryService.Save(entries, path);
        var loaded = RDPilotApplication.PromptHistoryService.Load(path);
        Assert(loaded.SequenceEqual(entries),
            "prompt history did not survive persistence");
        RDPilotApplication.PromptHistoryService.Remember(
            loaded,
            "second",
            path);
        Assert(loaded.SequenceEqual(["first", "third", "second"]),
            "reissued prompt was duplicated instead of moved to newest");

        var navigation =
            new RDPilotApplication.PromptHistoryService.NavigationState(loaded);
        Assert(navigation.Up("draft") == "second" &&
               navigation.Up("second") == "third" &&
               navigation.Down("third") == "second" &&
               navigation.Down("second") == "draft",
            "Up/Down navigation did not preserve newest-first traversal and draft input");
    }
    finally
    {
        Directory.Delete(root, recursive: true);
    }
}

static void TestGoalModes()
{
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode("Monitoruj kolejkę nowych zadań") == "continuous",
        "general monitoring goal was not continuous");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode("Utrzymuj proces aktywny do odwołania") == "continuous",
        "user-terminated maintenance goal was not continuous");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode("Obserwuj logi przez 10 minut") == "finite",
        "time-bounded observation goal was not finite");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode("Graj w grę") == "continuous",
        "open-ended game goal was not continuous");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode(
            "Otwórz zadanie, samodzielnie rozwiąż pierwszy poziom i doprowadź go do stanu ukończenia. Nie przerywaj po częściowym postępie.") == "finite",
        "explicitly bounded game goal was overridden by open-ended wording");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode(
            "Sprawdzaj okresowo stan kolejki") == "continuous",
        "general recurring workflow was not continuous");
    Assert(
        RDPilotApplication.RecoveryMemoryService.HasRecurringWorkflowIntent(
            "Sprawdzaj okresowo stan kolejki"),
        "recurring workflow intent was not recognized");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ClassifyGoalMode(
            "Powtarzaj sprawdzenie 5 razy") == "finite",
        "count-bounded recurring workflow was not finite");
}

static void TestGoalModeOverride()
{
    Assert(
        RDPilotApplication.RecoveryMemoryService.ResolveGoalMode(
            "Monitoruj kolejkę",
            "finite") == "finite",
        "finite override did not win over the heuristic");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ResolveGoalMode(
            "Zamknij Notatnik",
            "continuous") == "continuous",
        "continuous override did not win over the heuristic");
    Assert(
        RDPilotApplication.RecoveryMemoryService.ResolveGoalMode(
            "Monitoruj kolejkę",
            "auto") == "continuous",
        "auto mode did not delegate to classification");
}

static void TestUnlimitedStepConfiguration()
{
    var root = typeof(RDPilotApplication);
    var maxStepsField = root.GetField(
        "MaxSteps",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var original = maxStepsField.GetValue(null);
    try
    {
        _ = RDPilotApplication.ConfigurationService.ApplyCliArgs(
            ["--max-steps", "0"]);
        RDPilotApplication.ConfigurationService.NormalizeConfig();
        Assert(
            (int)maxStepsField.GetValue(null)! == 0,
            "maxSteps=0 was normalized back to a finite limit");
    }
    finally
    {
        maxStepsField.SetValue(null, original);
    }
}

static void TestContinuousIdle()
{
    var wait = new ActionDto { Type = "wait", WaitSeconds = 10 };
    Assert(
        RDPilotApplication.ControlLoopService.IsExpectedContinuousIdle(
            "continuous",
            wait,
            noChange: true),
        "unchanged continuous wait was treated as stagnation");
    Assert(
        !RDPilotApplication.ControlLoopService.IsExpectedContinuousIdle(
            "finite",
            wait,
            noChange: true),
        "finite wait incorrectly bypassed stagnation");
    Assert(
        !RDPilotApplication.ControlLoopService.IsExpectedContinuousIdle(
            "continuous",
            new ActionDto { Type = "click", XPx = 10, YPx = 10 },
            noChange: true),
        "ineffective continuous click incorrectly bypassed stagnation");
    Assert(
        !RDPilotApplication.ControlLoopService.IsExpectedContinuousIdle(
            "continuous",
            wait,
            noChange: false),
        "a changed screen was labelled idle");
}

static void TestProgressVerdict()
{
    Assert(
        RDPilotApplication.RecoveryMemoryService.IsVerifiedRecoveryProgress(
            new RecoveryProgressDto { Verdict = "yes", Confidence = 0.9 }),
        "strong positive verifier verdict was rejected");
    Assert(
        !RDPilotApplication.RecoveryMemoryService.IsVerifiedRecoveryProgress(
            new RecoveryProgressDto { Verdict = "yes", Confidence = 0.51 }),
        "weak verifier verdict was accepted");
    Assert(
        !RDPilotApplication.RecoveryMemoryService.IsVerifiedRecoveryProgress(
            new RecoveryProgressDto { Verdict = "uncertain", Confidence = 0.99 }),
        "uncertain verifier verdict was accepted");
}

static void TestAutoVerificationPolicy()
{
    var root = typeof(RDPilotApplication);
    var modeField = root.GetField("VerifyMode", BindingFlags.Static | BindingFlags.NonPublic)!;
    var thresholdField = root.GetField("SkipVerifyConfidenceThreshold", BindingFlags.Static | BindingFlags.NonPublic)!;
    var earlyField = root.GetField("VerifyEarlySteps", BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalMode = modeField.GetValue(null);
    var originalThreshold = thresholdField.GetValue(null);
    var originalEarly = earlyField.GetValue(null);
    try
    {
        modeField.SetValue(null, "auto");
        thresholdField.SetValue(null, 0.92);
        earlyField.SetValue(null, 2);
        Assert(!RDPilotApplication.ControlLoopService.ShouldVerifyGoal(
                "draw a cat",
                11,
                new ActionDto { Type = "done", Confidence = 0.99 }),
            "auto mode did not skip a strong late completion");
        Assert(RDPilotApplication.ControlLoopService.VerificationSkipReason(
                11,
                new ActionDto { Type = "done", Confidence = 0.99 })
                .Contains("high-confidence completion", StringComparison.Ordinal),
            "verification skip log does not explain the high-confidence threshold");
        Assert(RDPilotApplication.ControlLoopService.ShouldVerifyGoal(
                "draw a cat",
                11,
                new ActionDto { Type = "done", Confidence = 0.7 }),
            "auto mode skipped a low-confidence completion");
        Assert(RDPilotApplication.ControlLoopService.ShouldVerifyGoal(
                "draw a cat",
                1,
                new ActionDto { Type = "done", Confidence = 0.99 }),
            "auto mode skipped verification during early steps");
    }
    finally
    {
        modeField.SetValue(null, originalMode);
        thresholdField.SetValue(null, originalThreshold);
        earlyField.SetValue(null, originalEarly);
    }
}

static void TestRecoveryValidationProgress()
{
    var observationOnly = new ResolvedActionSnapshot(
        new ActionDto { Type = "aim", BBox = new BBox { Left = 1, Top = 1, Right = 10, Bottom = 10 } },
        "aim",
        "aim",
        null);
    var mutation = new ResolvedActionSnapshot(
        new ActionDto { Type = "drag_path", GestureKind = "draw" },
        "drag_path",
        "drag_path",
        null);
    var progress = new ObservationAssessment(
        "local_editing",
        VisualChangeState.Changed,
        ActionOutcomeState.Confirmed,
        GoalProgressState.Progress,
        0.9,
        0.02,
        0.01,
        0.01,
        0.02,
        0.1,
        0.005,
        "confirmed drawing progress");
    var noProgress = progress with
    {
        ActionOutcome = ActionOutcomeState.NoEffect,
        GoalProgress = GoalProgressState.NoProgress
    };

    Assert(
        !RDPilotApplication.RecoveryMemoryService.IsRecoveryValidationProgress(
            observationOnly,
            progress,
            meaningfulStateTransition: true),
        "observation-only action incorrectly confirmed a recovery lesson");
    Assert(
        !RDPilotApplication.RecoveryMemoryService.IsRecoveryValidationProgress(
            mutation,
            noProgress,
            meaningfulStateTransition: true),
        "visual noise overrode an explicit no-progress observation");
    Assert(
        RDPilotApplication.RecoveryMemoryService.IsRecoveryValidationProgress(
            mutation,
            progress,
            meaningfulStateTransition: false),
        "confirmed mutation progress did not validate recovery");
}

static void TestStrictStrategyAttribution()
{
    var expected = new RecoveryStrategyStep
    {
        ActionFamily = "click",
        ParameterSignature = "left",
        TargetTokens = "blue tile"
    };
    var withoutTarget = new ResolvedActionSnapshot(
        new ActionDto { Type = "click", XPx = 10, YPx = 10 },
        "click",
        "click",
        new Point(10, 10));
    var matches = (bool)InvokePrivate(
        typeof(RDPilotApplication.RecoveryMemoryService),
        "StrategyStepMatches",
        expected,
        withoutTarget,
        false)!;
    Assert(!matches, "strategy matched even though actual target semantics were missing");
}

static void TestExplicitStrategyAttribution()
{
    var lesson = SemanticLesson("blue tile", "click:left:blue tile");
    var episode = new RecoveryEpisodeState();
    episode.SuggestedLessonIds.Add(lesson.Id);
    var queue = new Queue<ResolvedActionSnapshot>();
    var action = new ResolvedActionSnapshot(
        new ActionDto
        {
            Type = "click",
            XPx = 10,
            YPx = 10,
            Note = "blue tile",
            RecoveryStrategyId = lesson.Id,
            RecoveryStrategyStep = 1
        },
        "click blue tile",
        "click:10,10",
        new Point(10, 10))
    {
        SemanticTokens = "blue tile"
    };
    RDPilotApplication.RecoveryMemoryService.RecordRecoveryAction(
        episode,
        queue,
        action,
        [lesson]);
    Assert(episode.AppliedLessonId == lesson.Id, "explicit strategy id was not attributed");
}

static void TestSensitiveActionSignatures()
{
    var first = RDPilotApplication.DesktopInputService.ActionSignature(
        new ActionDto { Type = "run_command", Command = "echo first" });
    var second = RDPilotApplication.DesktopInputService.ActionSignature(
        new ActionDto { Type = "run_command", Command = "echo second" });
    Assert(first != second, "different commands share an ineffective-action signature");

    var textA = RDPilotApplication.DesktopInputService.ActionSignature(
        new ActionDto { Type = "type_text", Text = "alpha", Note = "search field" });
    var textB = RDPilotApplication.DesktopInputService.ActionSignature(
        new ActionDto { Type = "type_text", Text = "beta", Note = "search field" });
    Assert(textA != textB, "different text inputs share an ineffective-action signature");
}

static void TestSemanticStrategyIdentity()
{
    var memoryType = typeof(RDPilotApplication.RecoveryMemoryService);
    var readOnlyField = memoryType.GetField("RecoveryMemoryReadOnly", BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalReadOnly = readOnlyField.GetValue(null);
    readOnlyField.SetValue(null, true);
    try
    {
        var lessons = new List<RecoveryLesson>();
        var blue = SemanticLesson("blue tile", "click:left:blue tile");
        var red = SemanticLesson("red tile", "click:left:red tile");
        InvokePrivate(memoryType, "StoreConfirmedLesson", blue, lessons);
        InvokePrivate(memoryType, "StoreConfirmedLesson", red, lessons);
        Assert(lessons.Count == 2, "different semantic targets were merged");

        blue.Status = "quarantined";
        blue.QuarantinedUtc = DateTime.UtcNow;
        var recoveredBlue = SemanticLesson("blue tile", "click:left:blue tile");
        InvokePrivate(memoryType, "StoreConfirmedLesson", recoveredBlue, lessons);
        Assert(blue.Status == "active", "confirmed success did not revive quarantined strategy");
    }
    finally
    {
        readOnlyField.SetValue(null, originalReadOnly);
    }
}

static void TestCustomProfileReset()
{
    RDPilotApplication.ConfigurationService.ApplyProfile("fast");
    RDPilotApplication.ConfigurationService.ApplyProfile("custom");
    var root = typeof(RDPilotApplication);
    Assert(GetStatic<string>(root, "RunProfile") == "custom", "profile name did not reset");
    Assert(GetStatic<string?>(root, "ReasoningEffort") == "max", "reasoning did not reset to the max code default");
    Assert(GetStatic<string?>(root, "VerifyReasoningEffort") == "max", "verifier reasoning did not reset to the max code default");
    Assert(GetStatic<int>(root, "MaxStagnationStepsBeforeAbort") == 20, "stagnation default did not reset");
    Assert(GetStatic<int>(root, "RecoveryMemoryMaxLessons") == 500, "recovery-memory active default changed");
    Assert(GetStatic<int>(root, "RecoveryMemoryMaxQuarantinedLessons") == 500, "recovery-memory quarantine default changed");
}

static void TestMaxReasoningEffortDoesNotDowngrade()
{
    var root = typeof(RDPilotApplication);
    var effortField = root.GetField("ReasoningEffort", BindingFlags.Static | BindingFlags.NonPublic)!;
    var adaptiveField = root.GetField("AdaptiveReasoningEffort", BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalEffort = effortField.GetValue(null);
    var originalAdaptive = adaptiveField.GetValue(null);
    try
    {
        effortField.SetValue(null, "max");
        adaptiveField.SetValue(null, true);
        var effective = (string?)InvokePrivate(
            typeof(RDPilotApplication.ControlLoopService),
            "EffectiveReasoningEffort",
            8,
            4,
            4);
        Assert(effective == "max", "adaptive effort downgraded max after stagnation");
    }
    finally
    {
        effortField.SetValue(null, originalEffort);
        adaptiveField.SetValue(null, originalAdaptive);
    }
}

static void TestProfilesPreserveStrongEffort()
{
    var root = typeof(RDPilotApplication);
    var effortField = root.GetField("ReasoningEffort", BindingFlags.Static | BindingFlags.NonPublic)!;
    var explicitField = root.GetField("ReasoningEffortExplicit", BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalEffort = effortField.GetValue(null);
    var originalExplicit = explicitField.GetValue(null);
    try
    {
        explicitField.SetValue(null, false);
        RDPilotApplication.ConfigurationService.ApplyProfile("custom");
        RDPilotApplication.ConfigurationService.ApplyProfile("fast");
        Assert(GetStatic<string?>(root, "ReasoningEffort") == "max", "fast profile lowered max reasoning effort");
        Assert(GetStatic<int>(root, "MaxOutputTokens") >= 10000, "fast profile lowered the max reasoning budget");
        Assert((string?)InvokePrivate(typeof(RDPilotApplication.ConfigurationService), "EffectiveQaReasoningEffort") == "max", "fast profile lowered QA reasoning effort");
        Assert(GetStatic<int>(root, "QaMaxOutputTokens") >= 4000, "fast profile lowered the QA reasoning budget");
        Assert((string?)InvokePrivate(typeof(RDPilotApplication.ConfigurationService), "EffectiveVerifyReasoningEffort") == "max", "fast profile lowered verifier reasoning effort");
        Assert(GetStatic<int>(root, "VerifyMaxOutputTokens") >= 6000, "fast profile lowered the verifier reasoning budget");
        Assert(GetStatic<int>(root, "TurnReanalysisMaxOutputTokens") >= 10000,
            "fast profile lowered the salient reanalysis budget");

        explicitField.SetValue(null, true);
        effortField.SetValue(null, "low");
        RDPilotApplication.ConfigurationService.ApplyProfile("quality");
        Assert(GetStatic<string?>(root, "ReasoningEffort") == "low", "explicit low effort was overwritten by a profile");
    }
    finally
    {
        effortField.SetValue(null, originalEffort);
        explicitField.SetValue(null, originalExplicit);
        RDPilotApplication.ConfigurationService.ApplyProfile("custom");
    }
}

static void TestOutputRetriesFollowReasoningFallback()
{
    object ladderBody = new Dictionary<string, object?>
    {
        ["model"] = "gpt-5.6-luna",
        ["max_output_tokens"] = 6000,
        ["previous_response_id"] = "resp_previous",
        ["reasoning"] = new Dictionary<string, object?>
        {
            ["effort"] = "max",
            ["context"] = "all_turns"
        }
    };
    var expectedEfforts = new[] { "max", "xhigh", "high", "medium", "low" };
    for (var index = 1; index < expectedEfforts.Length; index++)
    {
        var canRetry = RDPilotApplication.OpenAiResponsesService.TryBuildMaxOutputRetryBody(
            ladderBody,
            out var nextBody,
            out var nextMaxTokens,
            out var nextEffort,
            out _);
        Assert(canRetry, $"retry stopped before reaching {expectedEfforts[index]}");
        Assert(nextMaxTokens == 6000,
            $"retry enlarged the token budget instead of lowering effort at {expectedEfforts[index]}");
        Assert(nextEffort == expectedEfforts[index], $"retry did not step down to {expectedEfforts[index]}");
        var nextReasoning = (System.Text.Json.Nodes.JsonObject)((System.Text.Json.Nodes.JsonObject)nextBody)["reasoning"]!;
        Assert(nextReasoning["effort"]?.GetValue<string>() == expectedEfforts[index], $"retry body did not contain {expectedEfforts[index]}");
        Assert(nextReasoning["context"]?.GetValue<string>() == "all_turns",
            "effort fallback lost persisted reasoning context");
        Assert(((System.Text.Json.Nodes.JsonObject)nextBody)["previous_response_id"]?.GetValue<string>() == "resp_previous",
            "effort fallback lost the previous response id");
        ladderBody = nextBody;
    }

    var exhausted = RDPilotApplication.OpenAiResponsesService.TryBuildMaxOutputRetryBody(
        ladderBody,
        out _,
        out var exhaustedMaxTokens,
        out var exhaustedEffort,
        out _);
    Assert(!exhausted, "retry continued below the low effort floor");
    Assert(exhaustedMaxTokens == 6000, "exhausted retry changed the final token cap");
    Assert(exhaustedEffort == "low", "exhausted retry lost the low effort marker");

    var bodyWithoutEffort = new Dictionary<string, object?>
    {
        ["model"] = "model-without-effort",
        ["max_output_tokens"] = 1000
    };
    Assert(RDPilotApplication.OpenAiResponsesService.TryBuildMaxOutputRetryBody(
            bodyWithoutEffort,
            out _,
            out var expandedTokens,
            out _,
            out _) &&
           expandedTokens > 1000,
        "a request without a reasoning ladder cannot expand its output budget");
}

static void TestPartialTokenTelemetrySuppressesCacheWarning()
{
    Assert(
        !RDPilotApplication.RunMetrics.ShouldWarnPromptCache(true, 7, 1101, 0, 6),
        "partial token telemetry produced a false cache warning");
    Assert(
        RDPilotApplication.RunMetrics.ShouldWarnPromptCache(true, 7, 1101, 0, 0),
        "complete token telemetry did not report a missing cache hit");
}

static void TestQuarantine()
{
    var lesson = new RecoveryLesson
    {
        UpdatedUtc = DateTime.UtcNow.AddDays(-500),
        SuccessCount = 1,
        FailureCount = 5
    };
    var lessons = new List<RecoveryLesson> { lesson };
    InvokePrivate(
        typeof(RDPilotApplication.RecoveryMemoryService),
        "ApplyLessonRetention",
        lessons);
    Assert(lesson.Status == "quarantined", "stale weak lesson was not quarantined");
}

static void TestContextDiverseRetention()
{
    var root = typeof(RDPilotApplication);
    var maxField = root.GetField(
        "RecoveryMemoryMaxLessons",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var reserveField = root.GetField(
        "RecoveryMemoryReservedLessonsPerContext",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var softMaxField = root.GetField(
        "RecoveryMemorySoftMaxLessonsPerContext",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalMax = maxField.GetValue(null);
    var originalReserve = reserveField.GetValue(null);
    var originalSoftMax = softMaxField.GetValue(null);
    try
    {
        maxField.SetValue(null, 6);
        reserveField.SetValue(null, 2);
        softMaxField.SetValue(null, 3);
        var lessons = new List<RecoveryLesson>();
        for (var index = 0; index < 10; index++)
        {
            lessons.Add(new RecoveryLesson
            {
                Id = $"dominant-{index}",
                ActiveProcess = "dominant-app",
                GoalDomain = "general",
                SuccessCount = 20,
                UpdatedUtc = DateTime.UtcNow.AddMinutes(-index)
            });
        }
        for (var index = 0; index < 2; index++)
        {
            lessons.Add(new RecoveryLesson
            {
                Id = $"rare-b-{index}",
                ActiveProcess = "rare-app-b",
                GoalDomain = "monitoring",
                SuccessCount = 1,
                UpdatedUtc = DateTime.UtcNow.AddHours(-1 - index)
            });
            lessons.Add(new RecoveryLesson
            {
                Id = $"rare-c-{index}",
                ActiveProcess = "rare-app-c",
                GoalDomain = "development",
                SuccessCount = 1,
                UpdatedUtc = DateTime.UtcNow.AddHours(-2 - index)
            });
        }

        var keep =
            RDPilotApplication.RecoveryMemoryService
                .SelectActiveLessonIdsForRetention(lessons);
        Assert(keep.Count == 6, "global active-memory limit was not respected");
        Assert(
            keep.Count(id => id.StartsWith("rare-b-", StringComparison.Ordinal)) == 2,
            "rare application B lost its reserved lessons");
        Assert(
            keep.Count(id => id.StartsWith("rare-c-", StringComparison.Ordinal)) == 2,
            "rare application C lost its reserved lessons");
    }
    finally
    {
        maxField.SetValue(null, originalMax);
        reserveField.SetValue(null, originalReserve);
        softMaxField.SetValue(null, originalSoftMax);
    }
}

static void TestRecoveryArchive()
{
    var root = typeof(RDPilotApplication);
    var memoryType = typeof(RDPilotApplication.RecoveryMemoryService);
    var pathField = root.GetField(
        "RecoveryMemoryPath",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var archivePathField = root.GetField(
        "RecoveryMemoryArchivePath",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var maxField = root.GetField(
        "RecoveryMemoryMaxLessons",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var quarantineField = root.GetField(
        "RecoveryMemoryMaxQuarantinedLessons",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var readOnlyField = memoryType.GetField(
        "RecoveryMemoryReadOnly",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var originals = new[]
    {
        pathField.GetValue(null),
        archivePathField.GetValue(null),
        maxField.GetValue(null),
        quarantineField.GetValue(null),
        readOnlyField.GetValue(null)
    };
    var tempRoot = Path.Combine(
        Path.GetTempPath(),
        "RDPilotSelfTests",
        Guid.NewGuid().ToString("N"));
    var primaryPath = Path.Combine(
        tempRoot,
        "memory",
        "recovery-memory.json");
    try
    {
        pathField.SetValue(null, primaryPath);
        archivePathField.SetValue(null, null);
        maxField.SetValue(null, 1);
        quarantineField.SetValue(null, 1);
        readOnlyField.SetValue(null, false);
        var lessons = Enumerable.Range(0, 4)
            .Select(index => new RecoveryLesson
            {
                Id = $"archive-{index}",
                ActiveProcess = $"app-{index}",
                GoalDomain = "general",
                SuccessCount = index + 1,
                UpdatedUtc = DateTime.UtcNow.AddMinutes(-index)
            })
            .ToList();

        var saved = (bool)InvokePrivate(
            memoryType,
            "SaveRecoveryLessons",
            lessons)!;
        Assert(saved, "bounded memory could not be saved");
        Assert(
            lessons.Count == 2 &&
            lessons.Count(item => item.Status == "active") == 1,
            "primary memory did not retain one active and one quarantined lesson");
        var archivePath = Path.Combine(
            tempRoot,
            "memory",
            "recovery-memory-archive.json");
        Assert(File.Exists(archivePath), "archive file was not created");
        var archive =
            System.Text.Json.JsonSerializer.Deserialize<RecoveryLessonArchiveStore>(
                File.ReadAllText(archivePath),
                new System.Text.Json.JsonSerializerOptions(
                    System.Text.Json.JsonSerializerDefaults.Web));
        Assert(
            archive?.Lessons.Count == 2,
            "displaced lessons were not preserved in the archive");
    }
    finally
    {
        pathField.SetValue(null, originals[0]);
        archivePathField.SetValue(null, originals[1]);
        maxField.SetValue(null, originals[2]);
        quarantineField.SetValue(null, originals[3]);
        readOnlyField.SetValue(null, originals[4]);
        var expectedParent = Path.GetFullPath(
            Path.Combine(Path.GetTempPath(), "RDPilotSelfTests"));
        var resolved = Path.GetFullPath(tempRoot);
        if (resolved.StartsWith(
                expectedParent,
                StringComparison.OrdinalIgnoreCase) &&
            Directory.Exists(resolved))
        {
            Directory.Delete(resolved, recursive: true);
        }
    }
}

static void TestRecoveryFileSizeLimit()
{
    var root = typeof(RDPilotApplication);
    var memoryType = typeof(RDPilotApplication.RecoveryMemoryService);
    var maxBytesField = root.GetField(
        "RecoveryMemoryMaxFileBytes",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalMaxBytes = maxBytesField.GetValue(null);
    try
    {
        maxBytesField.SetValue(null, 12_000L);
        var lessons = Enumerable.Range(0, 8)
            .Select(index => new RecoveryLesson
            {
                Id = $"sized-{index}",
                ActiveProcess = "large-app",
                GoalDomain = "general",
                WinningStrategy = new string((char)('a' + index), 4_000),
                SuccessCount = index + 1,
                UpdatedUtc = DateTime.UtcNow.AddMinutes(-index)
            })
            .ToList();
        var archived = (List<RecoveryLesson>)InvokePrivate(
            memoryType,
            "EnforceRecoveryMemoryFileSize",
            lessons)!;
        var remainingBytes = (int)InvokePrivate(
            memoryType,
            "RecoveryStoreSerializedSize",
            lessons)!;

        Assert(archived.Count > 0, "oversized memory did not displace any lessons");
        Assert(lessons.Count < 8, "primary memory lesson count did not shrink");
        Assert(
            remainingBytes <= 12_000 || lessons.Count == 1,
            "primary memory stayed oversized despite removable lessons");
        Assert(
            archived.All(item => item.Status == "archived"),
            "displaced size-limit lessons were not marked archived");
    }
    finally
    {
        maxBytesField.SetValue(null, originalMaxBytes);
    }
}

static void TestWriterCounterCompaction()
{
    var lesson = new RecoveryLesson();
    var started = DateTime.UtcNow.AddDays(-30);
    for (var index = 0; index < 70; index++)
    {
        var writer = $"{started.AddMinutes(index).Ticks:x16}-{index}-test";
        lesson.SuccessByWriter[writer] = 1;
    }
    InvokePrivate(
        typeof(RDPilotApplication.RecoveryMemoryService),
        "NormalizeLessonCounters",
        lesson);
    InvokePrivate(
        typeof(RDPilotApplication.RecoveryMemoryService),
        "CompactLessonWriterCounters",
        lesson);
    Assert(lesson.SuccessCount == 70, "counter compaction changed the success total");
    Assert(lesson.SuccessByWriter.Count <= 32, "writer components were not bounded");
    Assert(lesson.CompactedSuccessCount > 0, "old writer components were not checkpointed");
}

static void TestBanditCounterMerge()
{
    var id = Guid.NewGuid().ToString("N");
    var writerA = $"{DateTime.UtcNow.Ticks:x16}-1-a";
    var writerB = $"{DateTime.UtcNow.AddTicks(1).Ticks:x16}-2-b";
    var left = new RecoveryLesson
    {
        Id = id,
        UpdatedUtc = DateTime.UtcNow.AddSeconds(-1),
        SelectionByWriter = new() { [writerA] = 1 },
        RewardByWriter = new() { [writerA] = 0.8 },
        RewardObservationByWriter = new() { [writerA] = 1 }
    };
    var right = new RecoveryLesson
    {
        Id = id,
        UpdatedUtc = DateTime.UtcNow,
        SelectionByWriter = new() { [writerB] = 2 },
        RewardByWriter = new() { [writerB] = 0.5 },
        RewardObservationByWriter = new() { [writerB] = 2 }
    };
    var merged = (List<RecoveryLesson>)InvokePrivate(
        typeof(RDPilotApplication.RecoveryMemoryService),
        "MergeRecoveryLessons",
        new List<RecoveryLesson> { left },
        new List<RecoveryLesson> { right })!;
    var lesson = merged.Single();
    Assert(
        lesson.SelectionCount == 3,
        "concurrent selection counts were not merged");
    Assert(
        Math.Abs(lesson.CumulativeReward - 1.3) < 0.0001,
        "concurrent cumulative rewards were not merged");
    Assert(
        lesson.RewardObservationCount == 3,
        "concurrent reward observation counts were not merged");
}

static void TestInconclusiveCalibrationMerge()
{
    var writerA = $"{DateTime.UtcNow.Ticks:x16}-1-a";
    var writerB =
        $"{DateTime.UtcNow.AddTicks(1).Ticks:x16}-2-b";
    var left = new Dictionary<string, LoopCalibrationBucket>(
        StringComparer.OrdinalIgnoreCase)
    {
        ["test"] = new LoopCalibrationBucket
        {
            InconclusiveByWriter = new() { [writerA] = 1 }
        }
    };
    var right = new Dictionary<string, LoopCalibrationBucket>(
        StringComparer.OrdinalIgnoreCase)
    {
        ["test"] = new LoopCalibrationBucket
        {
            InconclusiveByWriter = new() { [writerB] = 2 }
        }
    };
    var merged =
        (Dictionary<string, LoopCalibrationBucket>)InvokePrivate(
            typeof(RDPilotApplication.RecoveryMemoryService),
            "MergeCalibration",
            left,
            right)!;
    Assert(
        merged["test"].InconclusiveCount == 3,
        "concurrent inconclusive calibration counts were not merged");
    Assert(
        merged["test"].ConfirmedCount == 0 &&
        merged["test"].RejectedCount == 0,
        "inconclusive calibration changed labelled precision counters");
}

static void TestTelemetryReplayCorpus()
{
    var lines = new List<string>();
    for (var step = 1; step <= 6; step++)
    {
        lines.Add(TelemetryObservation(
            "negative-run",
            step,
            Fingerprint((byte)(10 + step)),
            confidence: 0.1,
            independentlyConfirmed: false,
            goalMode: "continuous",
            recurringWorkflowIntent: true));
    }
    byte[] cycleStates = [10, 100, 200];
    for (var step = 1; step <= 10; step++)
    {
        var state = cycleStates[(step - 1) % cycleStates.Length];
        lines.Add(TelemetryObservation(
            "positive-run",
            step,
            Fingerprint(state),
            confidence: step == 10 ? 0.95 : 0.4,
            independentlyConfirmed: step == 10));
    }
    lines.Add("{truncated");

    var existing = new LoopReplayCorpus
    {
        Cases =
        [
            new LoopReplayCase
            {
                Name = "manual-case",
                LabelSource = "manual",
                ExpectedLoop = false,
                Frames =
                [
                    new LoopReplayFrame
                    {
                        ScreenFingerprintBase64 =
                            Convert.ToBase64String(Fingerprint(1))
                    },
                    new LoopReplayFrame
                    {
                        ScreenFingerprintBase64 =
                            Convert.ToBase64String(Fingerprint(2))
                    }
                ]
            }
        ]
    };
    var corpus =
        RDPilotApplication.RecoveryMemoryService.BuildLoopReplayCorpus(
            lines,
            existing);

    Assert(
        corpus.Cases.Any(item =>
            item.Name == "manual-case" &&
            item.LabelSource == "manual"),
        "manual replay case was not preserved");
    Assert(
        corpus.Cases.Single(item => item.Name == "manual-case")
            .HasIndependentLabel,
        "manual replay label was not treated as independent");
    Assert(
        corpus.Cases.Any(item =>
            item.Name == "telemetry:positive-run:positive" &&
            item.ExpectedLoop &&
            item.Frames.Count == 10),
        "independently confirmed telemetry was not exported as positive");
    Assert(
        corpus.Cases.Any(item =>
            item.Name == "telemetry:negative-run:negative" &&
            !item.ExpectedLoop &&
            item.Frames.Count == 6),
        "quiet telemetry was not exported as negative");
    var positive = corpus.Cases.Single(item =>
        item.Name == "telemetry:positive-run:positive");
    var negative = corpus.Cases.Single(item =>
        item.Name == "telemetry:negative-run:negative");
    Assert(
        !positive.HasIndependentLabel &&
        !negative.HasIndependentLabel,
        "telemetry-derived replay labels were treated as independent accuracy labels");
    Assert(
        negative.GoalMode == "continuous" &&
        negative.RecurringWorkflowIntent,
        "goal cycle context was not preserved in replay telemetry");
    var replayState0 =
        Convert.FromBase64String(
            positive.Frames[0].ScreenFingerprintBase64);
    var replayState3 =
        Convert.FromBase64String(
            positive.Frames[3].ScreenFingerprintBase64);
    Assert(
        replayState0.SequenceEqual(replayState3),
        $"positive replay lost a repeated screen state ({string.Join(',', positive.Frames.Select(item => Convert.FromBase64String(item.ScreenFingerprintBase64)[0]))})");
    Assert(
        positive.Frames.Select(item => item.ActiveProcess)
            .Distinct(StringComparer.Ordinal)
            .Count() == 1,
        "positive replay lost stable process context");
    var positiveResult =
        RDPilotApplication.LoopReplayService.ReplayCase(positive);
    Assert(
        positiveResult.Detected,
        $"generated positive case was not replayable by the detector (step={positiveResult.Step}, confidence={positiveResult.Confidence:0.00})");
    Assert(
        !RDPilotApplication.LoopReplayService.ReplayCase(negative).Detected,
        "generated negative case produced a loop during replay");
}

static void TestIndependentReplayImport()
{
    var root = typeof(RDPilotApplication);
    var corpusPathField = root.GetField(
        "LoopReplayCorpusPath",
        BindingFlags.Static | BindingFlags.NonPublic)!;
    var original = corpusPathField.GetValue(null);
    var tempRoot = Path.Combine(
        Path.GetTempPath(),
        "RDPilotSelfTests",
        Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(tempRoot);
    try
    {
        var source = Path.Combine(tempRoot, "reviewed.json");
        var destination = Path.Combine(tempRoot, "merged.json");
        var reviewed = new LoopReplayCorpus
        {
            Cases =
            [
                new LoopReplayCase
                {
                    Name = "reviewed-negative",
                    LabelSource = "manual:reviewed",
                    ExpectedLoop = false,
                    Frames =
                    [
                        new LoopReplayFrame
                        {
                            ScreenFingerprintBase64 =
                                Convert.ToBase64String(Fingerprint(10))
                        },
                        new LoopReplayFrame
                        {
                            ScreenFingerprintBase64 =
                                Convert.ToBase64String(Fingerprint(20))
                        }
                    ]
                }
            ]
        };
        File.WriteAllText(
            source,
            System.Text.Json.JsonSerializer.Serialize(reviewed));
        corpusPathField.SetValue(null, destination);
        RDPilotApplication.LoopReplayService
            .ImportIndependentLoopReplayCorpus(source);
        var merged =
            System.Text.Json.JsonSerializer.Deserialize<LoopReplayCorpus>(
                File.ReadAllText(destination),
                new System.Text.Json.JsonSerializerOptions(
                    System.Text.Json.JsonSerializerDefaults.Web));
        Assert(
            merged?.Cases.Single().HasIndependentLabel == true,
            "reviewed replay case was not imported as an independent label");
    }
    finally
    {
        corpusPathField.SetValue(null, original);
        if (Directory.Exists(tempRoot))
            Directory.Delete(tempRoot, recursive: true);
    }
}

static void TestRecoveryPersistence()
{
    var root = typeof(RDPilotApplication);
    var pathField = root.GetField("RecoveryMemoryPath", BindingFlags.Static | BindingFlags.NonPublic)!;
    var enabledField = root.GetField("RecoveryMemoryEnabled", BindingFlags.Static | BindingFlags.NonPublic)!;
    var memoryType = typeof(RDPilotApplication.RecoveryMemoryService);
    var readOnlyField = memoryType.GetField("RecoveryMemoryReadOnly", BindingFlags.Static | BindingFlags.NonPublic)!;
    var originalPath = pathField.GetValue(null);
    var originalEnabled = enabledField.GetValue(null);
    var originalReadOnly = readOnlyField.GetValue(null);
    var tempRoot = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
    var path = Path.Combine(tempRoot, "memory", "recovery-memory.json");
    Directory.CreateDirectory(tempRoot);
    try
    {
        pathField.SetValue(null, path);
        enabledField.SetValue(null, true);
        readOnlyField.SetValue(null, false);
        var lessons = RDPilotApplication.RecoveryMemoryService.LoadRecoveryLessons();
        var lesson = new RecoveryLesson
        {
            ActiveProcess = "test",
            GoalTokens = "test goal",
            LoopKind = "test-loop",
            LoopTopology = "direct-repeat",
            InteractionDomain = "generic-ui",
            StrategySignature = "keys:enter",
            UpdatedUtc = DateTime.UtcNow
        };
        lessons.Add(lesson);
        InvokePrivate(memoryType, "SaveRecoveryLessons", lessons);
        var concurrentLesson = new RecoveryLesson
        {
            ActiveProcess = "other",
            GoalTokens = "other goal",
            LoopKind = "other-loop",
            LoopTopology = "stagnation",
            InteractionDomain = "generic-ui",
            StrategySignature = "keys:escape",
            UpdatedUtc = DateTime.UtcNow
        };
        var staleWriterView = new List<RecoveryLesson> { concurrentLesson };
        InvokePrivate(memoryType, "SaveRecoveryLessons", staleWriterView);
        var merged = RDPilotApplication.RecoveryMemoryService.LoadRecoveryLessons();
        Assert(
            merged.Any(item => item.Id == lesson.Id) &&
            merged.Any(item => item.Id == concurrentLesson.Id),
            "merge-on-write lost an entry from another writer");
        var exportPath = Path.Combine(tempRoot, "exported-memory.json");
        RDPilotApplication.RecoveryMemoryService.ExecuteRecoveryMemoryMaintenance(
            "export",
            exportPath);
        Assert(File.Exists(exportPath), "memory export was not created");
        lesson.SuccessCount++;
        lesson.UpdatedUtc = DateTime.UtcNow.AddSeconds(1);
        InvokePrivate(memoryType, "SaveRecoveryLessons", lessons);
        Assert(File.Exists(path + ".bak"), "backup was not created");

        File.WriteAllText(path, "{broken-json");
        var restored = RDPilotApplication.RecoveryMemoryService.LoadRecoveryLessons();
        Assert(restored.Any(item => item.Id == lesson.Id), "backup did not restore the lesson");

        File.WriteAllText(path, """{"version":999,"lessons":[],"calibration":{}}""");
        _ = RDPilotApplication.RecoveryMemoryService.LoadRecoveryLessons();
        InvokePrivate(memoryType, "SaveRecoveryLessons", new List<RecoveryLesson>());
        Assert(
            File.ReadAllText(path).Contains("\"version\":999", StringComparison.Ordinal),
            "a newer memory schema was downgraded");
    }
    finally
    {
        pathField.SetValue(null, originalPath);
        enabledField.SetValue(null, originalEnabled);
        readOnlyField.SetValue(null, originalReadOnly);
        var expectedParent = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "RDPilotSelfTests"));
        var resolved = Path.GetFullPath(tempRoot);
        if (resolved.StartsWith(expectedParent, StringComparison.OrdinalIgnoreCase) &&
            Directory.Exists(resolved))
        {
            Directory.Delete(resolved, recursive: true);
        }
    }
}

static ResolvedActionSnapshot Snapshot(string type, Point point) =>
    new(
        new ActionDto { Type = type, XPx = point.X, YPx = point.Y, Note = "test target" },
        $"{type} {point}",
        $"{type}:{point.X},{point.Y}",
        point)
    {
        SemanticTokens = "test target"
    };

static RecoveryLesson SemanticLesson(string target, string signature) =>
    new()
    {
        ActiveProcess = "game",
        GoalTokens = "place tile slot",
        GoalDomain = "game",
        GoalIntentTokens = "move",
        WindowTitleTokens = "puzzle",
        FocusTokens = "canvas board",
        TriggerActionFamily = "click",
        LoopKind = RDPilotApplication.RecoveryMemoryService.PointerRegionLoop,
        LoopTopology = "direct-repeat",
        InteractionDomain = "raster-canvas",
        StrategySignature = signature,
        StrategySteps =
        [
            new RecoveryStrategyStep
            {
                ActionFamily = "click",
                TargetTokens = target,
                ParameterSignature = "left"
            }
        ],
        WinningActionTypes = ["click"],
        WinningStrategy = $"click {target}",
        ExpectedOutcomeTokens = "canvas board",
        UpdatedUtc = DateTime.UtcNow
    };

static byte[] Fingerprint(byte value) =>
    Enumerable.Repeat(value, 96 * 54).ToArray();

static string TelemetryObservation(
    string runId,
    int step,
    byte[] fingerprint,
    double confidence,
    bool independentlyConfirmed,
    string goalMode = "finite",
    bool recurringWorkflowIntent = false) =>
    System.Text.Json.JsonSerializer.Serialize(new
    {
        timestampUtc = DateTime.UtcNow.AddSeconds(step),
        @event = "observation",
        runId,
        step,
        screenWidth = 1280,
        screenHeight = 720,
        confidence,
        independentlyConfirmed,
        goalMode,
        recurringWorkflowIntent,
        replayFrame = new LoopReplayFrame
        {
            ScreenFingerprintBase64 =
                Convert.ToBase64String(fingerprint),
            ActiveWindowFingerprintBase64 =
                Convert.ToBase64String(fingerprint),
            ActiveProcess = "test",
            WindowTitle = "hashed title",
            FocusSummary = "hashed focus",
            LastDelta = step == 1 ? null : 0.1
        }
    });

static T GetStatic<T>(Type type, string name) =>
    (T)type.GetField(name, BindingFlags.Static | BindingFlags.NonPublic)!.GetValue(null)!;

static object? InvokePrivate(Type type, string name, params object?[] args)
{
    var method = type.GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic)
                 ?? throw new MissingMethodException(type.FullName, name);
    return method.Invoke(null, args);
}

static Exception RootException(Exception ex)
{
    while (ex.InnerException is not null)
        ex = ex.InnerException;
    return ex;
}

static void Assert(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}
