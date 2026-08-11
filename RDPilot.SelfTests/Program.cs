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
    ("strategy attribution requires semantic target evidence", TestStrictStrategyAttribution),
    ("explicit strategy identity is attributed deterministically", TestExplicitStrategyAttribution),
    ("sensitive action signatures distinguish different inputs", TestSensitiveActionSignatures),
    ("semantic strategies do not merge different targets", TestSemanticStrategyIdentity),
    ("custom profile restores code defaults", TestCustomProfileReset),
    ("adaptive effort never lowers max", TestMaxReasoningEffortDoesNotDowngrade),
    ("profiles preserve stronger configured effort and budgets", TestProfilesPreserveStrongEffort),
    ("output retries follow the effort fallback ladder", TestOutputRetriesFollowReasoningFallback),
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
        Assert(GetStatic<int>(root, "MaxOutputTokens") >= 4000, "fast profile lowered the max reasoning budget");
        Assert((string?)InvokePrivate(typeof(RDPilotApplication.ConfigurationService), "EffectiveQaReasoningEffort") == "max", "fast profile lowered QA reasoning effort");
        Assert(GetStatic<int>(root, "QaMaxOutputTokens") >= 2500, "fast profile lowered the QA reasoning budget");
        Assert((string?)InvokePrivate(typeof(RDPilotApplication.ConfigurationService), "EffectiveVerifyReasoningEffort") == "max", "fast profile lowered verifier reasoning effort");
        Assert(GetStatic<int>(root, "VerifyMaxOutputTokens") >= 1500, "fast profile lowered the verifier reasoning budget");

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
    foreach (var effort in new[] { "low", "medium", "high", "xhigh", "max" })
    {
        var body = new Dictionary<string, object?>
        {
            ["model"] = "gpt-5.6-luna",
            ["max_output_tokens"] = 4000,
            ["reasoning"] = new Dictionary<string, object?> { ["effort"] = effort }
        };
        var retryBody = (System.Text.Json.Nodes.JsonObject)RDPilotApplication.OpenAiResponsesService.BuildMaxOutputRetryBody(
            body,
            out var retryMaxTokens,
            out var retryEffort);
        var retryReasoning = (System.Text.Json.Nodes.JsonObject)retryBody["reasoning"]!;

        Assert(retryEffort == effort, $"retry changed reasoning effort from {effort} to {retryEffort}");
        Assert(retryReasoning["effort"]?.GetValue<string>() == effort, $"retry body changed reasoning effort from {effort}");
        Assert(retryMaxTokens > 4000, "retry did not increase the output budget");
    }

    object ladderBody = new Dictionary<string, object?>
    {
        ["model"] = "gpt-5.6-luna",
        ["max_output_tokens"] = 4000,
        ["reasoning"] = new Dictionary<string, object?> { ["effort"] = "max" }
    };
    var expectedEfforts = new[] { "max", "xhigh", "high", "medium", "low" };
    var firstRetry = RDPilotApplication.OpenAiResponsesService.TryBuildMaxOutputRetryBody(
        ladderBody,
        out var firstRetryBody,
        out var firstRetryMaxTokens,
        out var firstRetryEffort,
        out _);
    Assert(firstRetry, "retry did not expand the output budget before lowering effort");
    Assert(firstRetryMaxTokens == 8000, "first retry did not reach the configured token cap");
    Assert(firstRetryEffort == expectedEfforts[0], "first retry lowered effort before exhausting the token cap");
    ladderBody = firstRetryBody;

    for (var index = 1; index < expectedEfforts.Length; index++)
    {
        var canRetry = RDPilotApplication.OpenAiResponsesService.TryBuildMaxOutputRetryBody(
            ladderBody,
            out var nextBody,
            out var nextMaxTokens,
            out var nextEffort,
            out _);
        Assert(canRetry, $"retry stopped before reaching {expectedEfforts[index]}");
        Assert(nextMaxTokens == 8000, $"retry changed the cap unexpectedly at {expectedEfforts[index]}");
        Assert(nextEffort == expectedEfforts[index], $"retry did not step down to {expectedEfforts[index]}");
        var nextReasoning = (System.Text.Json.Nodes.JsonObject)((System.Text.Json.Nodes.JsonObject)nextBody)["reasoning"]!;
        Assert(nextReasoning["effort"]?.GetValue<string>() == expectedEfforts[index], $"retry body did not contain {expectedEfforts[index]}");
        ladderBody = nextBody;
    }

    var exhausted = RDPilotApplication.OpenAiResponsesService.TryBuildMaxOutputRetryBody(
        ladderBody,
        out _,
        out var exhaustedMaxTokens,
        out var exhaustedEffort,
        out _);
    Assert(!exhausted, "retry continued below the low effort floor");
    Assert(exhaustedMaxTokens == 8000, "exhausted retry changed the final token cap");
    Assert(exhaustedEffort == "low", "exhausted retry lost the low effort marker");
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
