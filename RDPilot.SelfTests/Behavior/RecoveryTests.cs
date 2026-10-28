using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static void TestMalformedActionCapture()
    {
        foreach (var type in new[]
        {
            "click",
            "move"
        }

        )
        {
            var snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = type }, null);
            Assert(!snapshot.IsValid, $"{type} should be invalid");
            Assert(snapshot.ValidationError?.Contains("requires", StringComparison.OrdinalIgnoreCase) == true, $"{type} should explain missing coordinates");
        }
    }

    static void TestActionPolicyResolution()
    {
        RDPilot.Configuration.ConfigurationService.ApplyObservationMode("auto", "self-test");
        var session = new RDPilot.Observation.AdaptiveObservationSession();
        var terminalKeys = new ResolvedActionSnapshot(new ActionDto { Type = "keys", Keys = ["enter"] }, "keys [enter]", "keys:enter", null);
        Assert(session.ResolveActionPolicy(terminalKeys, new UiPromptContext("Terminal", "windowsterminal", "console", null, null, null)) == "event_driven", "terminal Enter did not select event-driven observation");
        var hold = new ResolvedActionSnapshot(new ActionDto { Type = "hold_keys", Keys = ["w"], DurationMs = 250 }, "hold_keys [w] 250ms", "hold_keys:w:250", null);
        Assert(session.ResolveActionPolicy(hold, new UiPromptContext("Game", "game", "canvas", null, null, null)) == "realtime_interaction", "bounded key hold did not select realtime observation");
        var puzzleInspection = new ResolvedActionSnapshot(new ActionDto { Type = "request_crop", Note = "inspect the level board" }, "request_crop", "request_crop:board", null);
        Assert(session.ResolveActionPolicy(puzzleInspection, new UiPromptContext("ARC-AGI-3 Task #ls20", "msedge", "level board controls", null, null, null), "graj w grę zagadkę i przechodź kolejne poziomy") == "turn_based_interaction", "turn-based puzzle did not select discrete interaction observation");
        var promptSession = new RDPilot.Observation.AdaptiveObservationSession();
        promptSession.PrepareForPrompt(new UiPromptContext("ARC-AGI-3 Task #ls20", "msedge", "level board controls", null, null, null), "graj w grę zagadkę i przechodź kolejne poziomy");
        Assert(promptSession.EffectiveProfile == "turn_based_interaction", "visible puzzle context did not enable batching before the first board action");
        Assert(!RDPilot.Observation.AdaptiveObservationSession.IsTurnBasedInteractionContext("play a realtime racing game", "Racing Game", "animated canvas", "steer the car"), "realtime game was misclassified as turn-based");
    }

    static void TestRejectedProposalCycles()
    {
        Assert(RDPilot.Control.ControlLoopGuards.RepeatedStringCycleLength(["done", "done"]) == 1, "direct rejected-proposal repeat was missed");
        Assert(RDPilot.Control.ControlLoopGuards.RepeatedStringCycleLength(["click:a", "keys:tab", "click:a", "keys:tab"]) == 2, "multi-step rejected-proposal cycle was missed");
        Assert(RDPilot.Control.ControlLoopGuards.RepeatedStringCycleLength(["click:a", "keys:tab", "scroll:down"]) == 0, "non-cyclic proposal history was marked as a loop");
    }

    static void TestRejectedProposalResetPolicy()
    {
        var observation = Snapshot("request_crop", new Point(100, 100));
        var mutation = Snapshot("click", new Point(100, 100));
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldResetRejectedProposalLoop(observation, noChange: false, expectedContinuousIdle: false), "an observation-only action cleared rejected proposal history");
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldResetRejectedProposalLoop(mutation, noChange: true, expectedContinuousIdle: false), "an ineffective action cleared rejected proposal history");
        Assert(RDPilot.Control.ControlActionPolicy.ShouldResetRejectedProposalLoop(mutation, noChange: false, expectedContinuousIdle: false), "observed progress did not clear rejected proposal history");
    }

    static void TestRepeatAccounting()
    {
        var points = new List<ResolvedActionSnapshot>();
        var successful = Snapshot("click", new Point(100, 100));
        var state = RDPilot.Control.ControlLoopGuards.UpdateRepeatDetection(successful, noChange: false, repeatCount: 0, lastSignature: null, points);
        Assert(points.Count == 0 && state.RepeatCount == 0, "successful click polluted ineffective history");
        state = RDPilot.Control.ControlLoopGuards.UpdateRepeatDetection(successful, noChange: true, state.RepeatCount, state.LastSignature, points);
        Assert(state.RepeatCount == 0, "first ineffective click was called a repeat");
        var nearby = Snapshot("click", new Point(102, 101));
        state = RDPilot.Control.ControlLoopGuards.UpdateRepeatDetection(nearby, noChange: true, state.RepeatCount, state.LastSignature, points);
        Assert(state.RepeatCount == 1, "second nearby ineffective click was not counted");
        ObservationAssessment NoEffectAssessment(string policy) => new(policy, VisualChangeState.Stable, ActionOutcomeState.NoEffect, GoalProgressState.NoProgress, 0.82, 0, 0, 0, double.NaN, 0, 0.004, "no effect test")
        {
            ActionPolicy = policy
        };
        var staticNoEffect = NoEffectAssessment("static_ui");
        var turnNoEffect = NoEffectAssessment("turn_based_interaction");
        Assert(RDPilot.Control.ControlActionPolicy.ShouldRegisterImmediateNoEffectCooldown(new ActionDto { Type = "keys", Keys = ["SPACE"] }, staticNoEffect, repeatCount: 0), "the first unambiguous ineffective key action was not eligible for cooldown");
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldRegisterImmediateNoEffectCooldown(new ActionDto { Type = "click", X = 0.5, Y = 0.5 }, turnNoEffect, repeatCount: 0) && RDPilot.Control.ControlActionPolicy.ShouldRegisterImmediateNoEffectCooldown(new ActionDto { Type = "click", X = 0.5, Y = 0.5 }, turnNoEffect, repeatCount: 1), "turn input cooldown does not require a repeated confirmed failure");
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldRegisterImmediateNoEffectCooldown(new ActionDto { Type = "wait", WaitSeconds = 1 }, staticNoEffect, repeatCount: 0) && !RDPilot.Control.ControlActionPolicy.ShouldRegisterImmediateNoEffectCooldown(new ActionDto { Type = "type_text", Text = "test" }, staticNoEffect, repeatCount: 0), "wait or focused text handling was incorrectly replaced by the immediate action cooldown");
        var exactCooldowns = new Dictionary<string, int>();
        var spatialCooldowns = new List<SpatialActionCooldown>();
        RDPilot.Control.ControlLoopGuards.RegisterActionCooldown(successful, untilStep: 5, exactCooldowns, spatialCooldowns, clusterSpatially: false);
        Assert(RDPilot.Control.ControlLoopGuards.IsActionOnCooldown(successful, step: 2, exactCooldowns, spatialCooldowns, out _), "the exact first ineffective click was not blocked");
        Assert(!RDPilot.Control.ControlLoopGuards.IsActionOnCooldown(nearby, step: 2, exactCooldowns, spatialCooldowns, out _), "an immediate exact cooldown incorrectly blocked a nearby alternative target");
    }

    static void TestDoneIsNotLearned()
    {
        var queue = new Queue<ResolvedActionSnapshot>();
        var episode = new RecoveryEpisodeState();
        var done = new ResolvedActionSnapshot(new ActionDto { Type = "done" }, "done", "done", null);
        RDPilot.Persistence.RecoveryMemoryService.RecordRecoveryAction(episode, queue, done, Array.Empty<RecoveryLesson>());
        Assert(queue.Count == 0, "done entered recent action history");
        Assert(episode.RecoveryActions.Count == 0, "done entered the learned recovery strategy");
    }

    static void TestRecoveryActionHistoryBound()
    {
        var root = typeof(RDPilotApplication);
        var limitField = TestState.Field(root, "RuntimeRecoveryActionLimit", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = limitField.GetValue(null);
        try
        {
            limitField.SetValue(null, 8);
            var episode = new RecoveryEpisodeState();
            var queue = new Queue<ResolvedActionSnapshot>();
            for (var index = 0; index < 24; index++)
            {
                RDPilot.Persistence.RecoveryMemoryService.RecordRecoveryAction(episode, queue, Snapshot("click", new Point(index * 10, 10)), Array.Empty<RecoveryLesson>());
            }

            Assert(episode.RecoveryActions.Count == 8, "recovery action history exceeded its runtime limit");
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
        Assert(RDPilot.Persistence.RecoveryMemoryService.HasEmergingLoopPattern(actions, null), "nearby points on opposite grid sides should form the same spatial cycle");
    }

    static void TestStateGraphCycle()
    {
        var graph = new LoopStateGraph();
        var context = new UiPromptContext("\"Puzzle\"", "game", "canvas board", null, "0,0 800x600", null);
        var actions = new Queue<ResolvedActionSnapshot>();
        var states = new[]
        {
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10)
        };
        LoopDetectionAssessment? assessment = null;
        for (var i = 0; i < states.Length; i++)
        {
            var previous = i == 0 ? null : Snapshot("click", new Point(100 + i * 10, 100));
            if (previous is not null)
            {
                actions.Enqueue(previous);
                while (actions.Count > 12)
                    actions.Dequeue();
            }

            assessment = RDPilot.Persistence.RecoveryMemoryService.AssessVisualStateCycle(graph, states[i], states[i], context, i + 1, actions, previous, i == 0 ? double.NaN : 0.1, recordLearning: false);
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
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10),
            Fingerprint(100),
            Fingerprint(200),
            Fingerprint(10)
        };
        LoopDetectionAssessment? firstDetected = null;
        LoopDetectionAssessment? independentlyConfirmed = null;
        var trace = new List<string>();
        for (var index = 0; index < states.Length; index++)
        {
            var previous = index == 0 ? null : Snapshot("click", new Point(100 + index * 20, 100));
            if (previous is not null)
                actions.Enqueue(previous);
            var assessment = RDPilot.Persistence.RecoveryMemoryService.AssessVisualStateCycle(graph, states[index], states[index], context, index + 1, actions, previous, index == 0 ? double.NaN : 0.1, recordLearning: false);
            trace.Add($"{index + 1}: loop={assessment.IsLoop}, independent={assessment.IndependentlyConfirmed}, " + $"confidence={assessment.Confidence:0.00}, topology={assessment.LoopTopology}, " + $"key={assessment.CalibrationKey}");
            if (assessment.IsLoop && firstDetected is null)
                firstDetected = assessment;
            if (assessment.IndependentlyConfirmed)
                independentlyConfirmed = assessment;
        }

        Assert(firstDetected is not null, "loop was never detected");
        Assert(!firstDetected!.IndependentlyConfirmed, "first threshold crossing was used as its own label");
        Assert(independentlyConfirmed is not null, $"later recurrence did not independently confirm the candidate ({string.Join("; ", trace)})");
    }

    static void TestGraphCandidateExpiry()
    {
        var root = typeof(RDPilotApplication);
        var ttlField = TestState.Field(root, "GraphCandidateTtlSteps", BindingFlags.Static | BindingFlags.NonPublic)!;
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
            _ = RDPilot.Persistence.RecoveryMemoryService.AssessVisualStateCycle(graph, Fingerprint(10), Fingerprint(10), new UiPromptContext("Window", "test", "focus", null, null, null), 5, Array.Empty<ResolvedActionSnapshot>(), null, double.NaN, recordLearning: false);
            Assert(graph.PendingCandidateStep is null, "stale graph candidate did not expire");
        }
        finally
        {
            ttlField.SetValue(null, original);
        }
    }

    static void TestGraphCandidateInconclusiveExpiry()
    {
        var key = $"expiry-test-{Guid.NewGuid():N}";
        var memoryType = typeof(RDPilot.Persistence.RecoveryMemoryService);
        var calibrationField = TestState.Field(memoryType, "RecoveryCalibration", BindingFlags.Static | BindingFlags.NonPublic)!;
        var calibration = (System.Collections.IDictionary)calibrationField.GetValue(null)!;
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
            _ = RDPilot.Persistence.RecoveryMemoryService.AssessVisualStateCycle(graph, Fingerprint(10), Fingerprint(10), new UiPromptContext("Window", "test", "focus", null, null, null), 100, Array.Empty<ResolvedActionSnapshot>(), null, double.NaN, recordLearning: true);
            var bucket = (LoopCalibrationBucket)calibration[key]!;
            Assert(bucket.InconclusiveCount == 1, "expired candidate was not counted as inconclusive");
            Assert(bucket.RejectedCount == 0, "expired candidate was incorrectly counted as rejected");
        }
        finally
        {
            calibration.Remove(key);
        }
    }

    static void TestRuntimeStateBounds()
    {
        var root = typeof(RDPilotApplication);
        var limitField = TestState.Field(root, "RuntimeSemanticStateLimit", BindingFlags.Static | BindingFlags.NonPublic)!;
        var edgeLimitField = TestState.Field(root, "RuntimeGraphEdgeLimit", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = limitField.GetValue(null);
        var originalEdgeLimit = edgeLimitField.GetValue(null);
        try
        {
            limitField.SetValue(null, 32);
            edgeLimitField.SetValue(null, 32);
            var graph = new LoopStateGraph();
            for (var step = 1; step <= 80; step++)
            {
                _ = RDPilot.Persistence.RecoveryMemoryService.AssessVisualStateCycle(graph, Fingerprint((byte)step), Fingerprint((byte)step), new UiPromptContext($"Window {step}", "test", $"focus {step}", null, null, null), step, Array.Empty<ResolvedActionSnapshot>(), null, 0.1, recordLearning: false);
            }

            Assert(graph.SemanticVisitSteps.Count <= 32, "semantic state history exceeded its runtime limit");
            for (var index = 0; index < 100; index++)
            {
                graph.Edges.Add(new LoopStateEdge { FromNodeId = index % 5, ToNodeId = (index + 1) % 5, ActionKey = $"action-{index}", LastSeenStep = index, TraversalCount = 1 });
            }

            InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "PruneGraphEdges", graph);
            Assert(graph.Edges.Count <= 32, "graph edges exceeded their runtime limit");
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
        var reliableScore = RDPilot.Persistence.RecoveryMemoryService.ContextualBanditScore(reliable, 0.8);
        var unreliableScore = RDPilot.Persistence.RecoveryMemoryService.ContextualBanditScore(unreliable, 0.8);
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
            LoopKind = RDPilot.Persistence.RecoveryMemoryService.PointerRegionLoop,
            LoopTopology = "direct-repeat",
            InteractionDomain = "raster-canvas",
            ScreenFingerprintBase64 = Convert.ToBase64String(fingerprint),
            ActiveWindowFingerprintBase64 = Convert.ToBase64String(fingerprint)
        };
        var context = new UiPromptContext("\"Puzzle\"", "game", "canvas board", null, null, null);
        var score = RDPilot.Persistence.RecoveryMemoryService.LessonSimilarity(lesson, context, fingerprint, fingerprint, "click", lesson.LoopKind, lesson.LoopTopology, lesson.InteractionDomain, "delete red card", "game", "delete", "finite");
        Assert(score == 0, "unrelated goal should reject the lesson");
    }

    static void TestPromptHistory()
    {
        var root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, "prompt-history.json");
        try
        {
            Assert(RDPilot.Persistence.PromptHistoryService.EffectivePath() == Path.Combine(AppContext.BaseDirectory, "memory", "prompt-history.json"), "default prompt history is not stored beside the executable artifacts");
            var entries = RDPilot.Persistence.PromptHistoryService.NormalizeEntries(["first", "second", "first", " third "]);
            Assert(entries.SequenceEqual(["second", "first", "third"]), "history did not keep unique prompts in most-recent order");
            RDPilot.Persistence.PromptHistoryService.Save(entries, path);
            var loaded = RDPilot.Persistence.PromptHistoryService.Load(path);
            Assert(loaded.SequenceEqual(entries), "prompt history did not survive persistence");
            RDPilot.Persistence.PromptHistoryService.Remember(loaded, "second", path);
            Assert(loaded.SequenceEqual(["first", "third", "second"]), "reissued prompt was duplicated instead of moved to newest");
            var navigation = new RDPilot.Persistence.PromptHistoryService.NavigationState(loaded);
            Assert(navigation.Up("draft") == "second" && navigation.Up("second") == "third" && navigation.Down("third") == "second" && navigation.Down("second") == "draft", "Up/Down navigation did not preserve newest-first traversal and draft input");
            var legacyPath = Path.Combine(root, "legacy", "RDPilot", "prompt-history.json");
            var migratedPath = Path.Combine(root, "new", "memory", "prompt-history.json");
            Assert(RDPilot.Persistence.PromptHistoryService.Save(entries, legacyPath) && RDPilot.Persistence.PromptHistoryService.TryMigrateLegacyHistory(legacyPath, migratedPath) && !File.Exists(legacyPath) && RDPilot.Persistence.PromptHistoryService.Load(migratedPath).SequenceEqual(entries), "legacy prompt history was not safely moved to the artifact directory");
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    static void TestProgressVerdict()
    {
        Assert(RDPilot.Persistence.RecoveryMemoryService.IsVerifiedRecoveryProgress(new RecoveryProgressDto { Verdict = "yes", Confidence = 0.9 }), "strong positive verifier verdict was rejected");
        Assert(!RDPilot.Persistence.RecoveryMemoryService.IsVerifiedRecoveryProgress(new RecoveryProgressDto { Verdict = "yes", Confidence = 0.51 }), "weak verifier verdict was accepted");
        Assert(!RDPilot.Persistence.RecoveryMemoryService.IsVerifiedRecoveryProgress(new RecoveryProgressDto { Verdict = "uncertain", Confidence = 0.99 }), "uncertain verifier verdict was accepted");
    }

    static void TestRecoveryValidationProgress()
    {
        var observationOnly = new ResolvedActionSnapshot(new ActionDto { Type = "aim", BBox = new BBox { Left = 1, Top = 1, Right = 10, Bottom = 10 } }, "aim", "aim", null);
        var mutation = new ResolvedActionSnapshot(new ActionDto { Type = "drag_path", GestureKind = "draw" }, "drag_path", "drag_path", null);
        var progress = new ObservationAssessment("local_editing", VisualChangeState.Changed, ActionOutcomeState.Confirmed, GoalProgressState.Progress, 0.9, 0.02, 0.01, 0.01, 0.02, 0.1, 0.005, "confirmed drawing progress");
        var noProgress = progress with
        {
            ActionOutcome = ActionOutcomeState.NoEffect,
            GoalProgress = GoalProgressState.NoProgress
        };
        Assert(!RDPilot.Persistence.RecoveryMemoryService.IsRecoveryValidationProgress(observationOnly, progress, meaningfulStateTransition: true), "observation-only action incorrectly confirmed a recovery lesson");
        Assert(!RDPilot.Persistence.RecoveryMemoryService.IsRecoveryValidationProgress(mutation, noProgress, meaningfulStateTransition: true), "visual noise overrode an explicit no-progress observation");
        Assert(RDPilot.Persistence.RecoveryMemoryService.IsRecoveryValidationProgress(mutation, progress, meaningfulStateTransition: false), "confirmed mutation progress did not validate recovery");
    }

    static void TestStrictStrategyAttribution()
    {
        var expected = new RecoveryStrategyStep
        {
            ActionFamily = "click",
            ParameterSignature = "left",
            TargetTokens = "blue tile"
        };
        var withoutTarget = new ResolvedActionSnapshot(new ActionDto { Type = "click", XPx = 10, YPx = 10 }, "click", "click", new Point(10, 10));
        var matches = (bool)InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "StrategyStepMatches", expected, withoutTarget, false)!;
        Assert(!matches, "strategy matched even though actual target semantics were missing");
    }

    static void TestExplicitStrategyAttribution()
    {
        var lesson = SemanticLesson("blue tile", "click:left:blue tile");
        var episode = new RecoveryEpisodeState();
        episode.SuggestedLessonIds.Add(lesson.Id);
        var queue = new Queue<ResolvedActionSnapshot>();
        var action = new ResolvedActionSnapshot(new ActionDto { Type = "click", XPx = 10, YPx = 10, Note = "blue tile", RecoveryStrategyId = lesson.Id, RecoveryStrategyStep = 1 }, "click blue tile", "click:10,10", new Point(10, 10))
        {
            SemanticTokens = "blue tile"
        };
        RDPilot.Persistence.RecoveryMemoryService.RecordRecoveryAction(episode, queue, action, [lesson]);
        Assert(episode.AppliedLessonId == lesson.Id, "explicit strategy id was not attributed");
    }

    static void TestSensitiveActionSignatures()
    {
        var first = RDPilot.Desktop.DesktopInputService.ActionSignature(new ActionDto { Type = "run_command", Command = "echo first" });
        var second = RDPilot.Desktop.DesktopInputService.ActionSignature(new ActionDto { Type = "run_command", Command = "echo second" });
        Assert(first != second, "different commands share an ineffective-action signature");
        var textA = RDPilot.Desktop.DesktopInputService.ActionSignature(new ActionDto { Type = "type_text", Text = "alpha", Note = "search field" });
        var textB = RDPilot.Desktop.DesktopInputService.ActionSignature(new ActionDto { Type = "type_text", Text = "beta", Note = "search field" });
        Assert(textA != textB, "different text inputs share an ineffective-action signature");
    }

    static void TestSemanticStrategyIdentity()
    {
        var memoryType = typeof(RDPilot.Persistence.RecoveryMemoryService);
        var readOnlyField = TestState.Field(memoryType, "RecoveryMemoryReadOnly", BindingFlags.Static | BindingFlags.NonPublic)!;
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

    static void TestPartialTokenTelemetrySuppressesCacheWarning()
    {
        Assert(!RDPilot.Runtime.RunMetrics.ShouldWarnPromptCache(true, 7, 1101, 0, 6), "partial token telemetry produced a false cache warning");
        Assert(RDPilot.Runtime.RunMetrics.ShouldWarnPromptCache(true, 7, 1101, 0, 0), "complete token telemetry did not report a missing cache hit");
    }

    static void TestQuarantine()
    {
        var lesson = new RecoveryLesson
        {
            UpdatedUtc = DateTime.UtcNow.AddDays(-500),
            SuccessCount = 1,
            FailureCount = 5
        };
        var lessons = new List<RecoveryLesson>
        {
            lesson
        };
        InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "ApplyLessonRetention", lessons);
        Assert(lesson.Status == "quarantined", "stale weak lesson was not quarantined");
    }

    static void TestContextDiverseRetention()
    {
        var root = typeof(RDPilotApplication);
        var maxField = TestState.Field(root, "RecoveryMemoryMaxLessons", BindingFlags.Static | BindingFlags.NonPublic)!;
        var reserveField = TestState.Field(root, "RecoveryMemoryReservedLessonsPerContext", BindingFlags.Static | BindingFlags.NonPublic)!;
        var softMaxField = TestState.Field(root, "RecoveryMemorySoftMaxLessonsPerContext", BindingFlags.Static | BindingFlags.NonPublic)!;
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
                lessons.Add(new RecoveryLesson { Id = $"dominant-{index}", ActiveProcess = "dominant-app", GoalDomain = "general", SuccessCount = 20, UpdatedUtc = DateTime.UtcNow.AddMinutes(-index) });
            }

            for (var index = 0; index < 2; index++)
            {
                lessons.Add(new RecoveryLesson { Id = $"rare-b-{index}", ActiveProcess = "rare-app-b", GoalDomain = "monitoring", SuccessCount = 1, UpdatedUtc = DateTime.UtcNow.AddHours(-1 - index) });
                lessons.Add(new RecoveryLesson { Id = $"rare-c-{index}", ActiveProcess = "rare-app-c", GoalDomain = "development", SuccessCount = 1, UpdatedUtc = DateTime.UtcNow.AddHours(-2 - index) });
            }

            var keep = RDPilot.Persistence.RecoveryMemoryService.SelectActiveLessonIdsForRetention(lessons);
            Assert(keep.Count == 6, "global active-memory limit was not respected");
            Assert(keep.Count(id => id.StartsWith("rare-b-", StringComparison.Ordinal)) == 2, "rare application B lost its reserved lessons");
            Assert(keep.Count(id => id.StartsWith("rare-c-", StringComparison.Ordinal)) == 2, "rare application C lost its reserved lessons");
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
        var memoryType = typeof(RDPilot.Persistence.RecoveryMemoryService);
        var pathField = TestState.Field(root, "RecoveryMemoryPath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var archivePathField = TestState.Field(root, "RecoveryMemoryArchivePath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var maxField = TestState.Field(root, "RecoveryMemoryMaxLessons", BindingFlags.Static | BindingFlags.NonPublic)!;
        var quarantineField = TestState.Field(root, "RecoveryMemoryMaxQuarantinedLessons", BindingFlags.Static | BindingFlags.NonPublic)!;
        var readOnlyField = TestState.Field(memoryType, "RecoveryMemoryReadOnly", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originals = new[]
        {
            pathField.GetValue(null),
            archivePathField.GetValue(null),
            maxField.GetValue(null),
            quarantineField.GetValue(null),
            readOnlyField.GetValue(null)
        };
        var tempRoot = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        var primaryPath = Path.Combine(tempRoot, "memory", "recovery-memory.json");
        try
        {
            pathField.SetValue(null, primaryPath);
            archivePathField.SetValue(null, null);
            maxField.SetValue(null, 1);
            quarantineField.SetValue(null, 1);
            readOnlyField.SetValue(null, false);
            var lessons = Enumerable.Range(0, 4).Select(index => new RecoveryLesson { Id = $"archive-{index}", ActiveProcess = $"app-{index}", GoalDomain = "general", SuccessCount = index + 1, UpdatedUtc = DateTime.UtcNow.AddMinutes(-index) }).ToList();
            var saved = (bool)InvokePrivate(memoryType, "SaveRecoveryLessons", lessons)!;
            Assert(saved, "bounded memory could not be saved");
            Assert(lessons.Count == 2 && lessons.Count(item => item.Status == "active") == 1, "primary memory did not retain one active and one quarantined lesson");
            var archivePath = Path.Combine(tempRoot, "memory", "recovery-memory-archive.json");
            Assert(File.Exists(archivePath), "archive file was not created");
            var archive = System.Text.Json.JsonSerializer.Deserialize<RecoveryLessonArchiveStore>(File.ReadAllText(archivePath), new System.Text.Json.JsonSerializerOptions(System.Text.Json.JsonSerializerDefaults.Web));
            Assert(archive?.Lessons.Count == 2, "displaced lessons were not preserved in the archive");
        }
        finally
        {
            pathField.SetValue(null, originals[0]);
            archivePathField.SetValue(null, originals[1]);
            maxField.SetValue(null, originals[2]);
            quarantineField.SetValue(null, originals[3]);
            readOnlyField.SetValue(null, originals[4]);
            var expectedParent = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "RDPilotSelfTests"));
            var resolved = Path.GetFullPath(tempRoot);
            if (resolved.StartsWith(expectedParent, StringComparison.OrdinalIgnoreCase) && Directory.Exists(resolved))
            {
                Directory.Delete(resolved, recursive: true);
            }
        }
    }

    static void TestRecoveryFileSizeLimit()
    {
        var root = typeof(RDPilotApplication);
        var memoryType = typeof(RDPilot.Persistence.RecoveryMemoryService);
        var maxBytesField = TestState.Field(root, "RecoveryMemoryMaxFileBytes", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originalMaxBytes = maxBytesField.GetValue(null);
        try
        {
            maxBytesField.SetValue(null, 12_000L);
            var lessons = Enumerable.Range(0, 8).Select(index => new RecoveryLesson { Id = $"sized-{index}", ActiveProcess = "large-app", GoalDomain = "general", WinningStrategy = new string ((char)('a' + index), 4_000), SuccessCount = index + 1, UpdatedUtc = DateTime.UtcNow.AddMinutes(-index) }).ToList();
            var archived = (List<RecoveryLesson>)InvokePrivate(memoryType, "EnforceRecoveryMemoryFileSize", lessons)!;
            var remainingBytes = (int)InvokePrivate(memoryType, "RecoveryStoreSerializedSize", lessons)!;
            Assert(archived.Count > 0, "oversized memory did not displace any lessons");
            Assert(lessons.Count < 8, "primary memory lesson count did not shrink");
            Assert(remainingBytes <= 12_000 || lessons.Count == 1, "primary memory stayed oversized despite removable lessons");
            Assert(archived.All(item => item.Status == "archived"), "displaced size-limit lessons were not marked archived");
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

        InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "NormalizeLessonCounters", lesson);
        InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "CompactLessonWriterCounters", lesson);
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
            SelectionByWriter = new()
            {
                [writerA] = 1
            },
            RewardByWriter = new()
            {
                [writerA] = 0.8
            },
            RewardObservationByWriter = new()
            {
                [writerA] = 1
            }
        };
        var right = new RecoveryLesson
        {
            Id = id,
            UpdatedUtc = DateTime.UtcNow,
            SelectionByWriter = new()
            {
                [writerB] = 2
            },
            RewardByWriter = new()
            {
                [writerB] = 0.5
            },
            RewardObservationByWriter = new()
            {
                [writerB] = 2
            }
        };
        var merged = (List<RecoveryLesson>)InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "MergeRecoveryLessons", new List<RecoveryLesson> { left }, new List<RecoveryLesson> { right })!;
        var lesson = merged.Single();
        Assert(lesson.SelectionCount == 3, "concurrent selection counts were not merged");
        Assert(Math.Abs(lesson.CumulativeReward - 1.3) < 0.0001, "concurrent cumulative rewards were not merged");
        Assert(lesson.RewardObservationCount == 3, "concurrent reward observation counts were not merged");
    }

    static void TestInconclusiveCalibrationMerge()
    {
        var writerA = $"{DateTime.UtcNow.Ticks:x16}-1-a";
        var writerB = $"{DateTime.UtcNow.AddTicks(1).Ticks:x16}-2-b";
        var left = new Dictionary<string, LoopCalibrationBucket>(StringComparer.OrdinalIgnoreCase)
        {
            ["test"] = new LoopCalibrationBucket
            {
                InconclusiveByWriter = new()
                {
                    [writerA] = 1
                }
            }
        };
        var right = new Dictionary<string, LoopCalibrationBucket>(StringComparer.OrdinalIgnoreCase)
        {
            ["test"] = new LoopCalibrationBucket
            {
                InconclusiveByWriter = new()
                {
                    [writerB] = 2
                }
            }
        };
        var merged = (Dictionary<string, LoopCalibrationBucket>)InvokePrivate(typeof(RDPilot.Persistence.RecoveryMemoryService), "MergeCalibration", left, right)!;
        Assert(merged["test"].InconclusiveCount == 3, "concurrent inconclusive calibration counts were not merged");
        Assert(merged["test"].ConfirmedCount == 0 && merged["test"].RejectedCount == 0, "inconclusive calibration changed labelled precision counters");
    }

    static void TestTelemetryReplayCorpus()
    {
        var lines = new List<string>();
        for (var step = 1; step <= 6; step++)
        {
            lines.Add(TelemetryObservation("negative-run", step, Fingerprint((byte)(10 + step)), confidence: 0.1, independentlyConfirmed: false, goalMode: "continuous", recurringWorkflowIntent: true));
        }

        byte[] cycleStates = [10, 100, 200];
        for (var step = 1; step <= 10; step++)
        {
            var state = cycleStates[(step - 1) % cycleStates.Length];
            lines.Add(TelemetryObservation("positive-run", step, Fingerprint(state), confidence: step == 10 ? 0.95 : 0.4, independentlyConfirmed: step == 10));
        }

        lines.Add("{truncated");
        var existing = new LoopReplayCorpus
        {
            Cases = [new LoopReplayCase
            {
                Name = "manual-case",
                LabelSource = "manual",
                ExpectedLoop = false,
                Frames = [new LoopReplayFrame
                {
                    ScreenFingerprintBase64 = Convert.ToBase64String(Fingerprint(1))
                }, new LoopReplayFrame
                {
                    ScreenFingerprintBase64 = Convert.ToBase64String(Fingerprint(2))
                }

                ]
            }

            ]
        };
        var corpus = RDPilot.Persistence.RecoveryMemoryService.BuildLoopReplayCorpus(lines, existing);
        Assert(corpus.Cases.Any(item => item.Name == "manual-case" && item.LabelSource == "manual"), "manual replay case was not preserved");
        Assert(corpus.Cases.Single(item => item.Name == "manual-case").HasIndependentLabel, "manual replay label was not treated as independent");
        Assert(corpus.Cases.Any(item => item.Name == "telemetry:positive-run:positive" && item.ExpectedLoop && item.Frames.Count == 10), "independently confirmed telemetry was not exported as positive");
        Assert(corpus.Cases.Any(item => item.Name == "telemetry:negative-run:negative" && !item.ExpectedLoop && item.Frames.Count == 6), "quiet telemetry was not exported as negative");
        var positive = corpus.Cases.Single(item => item.Name == "telemetry:positive-run:positive");
        var negative = corpus.Cases.Single(item => item.Name == "telemetry:negative-run:negative");
        Assert(!positive.HasIndependentLabel && !negative.HasIndependentLabel, "telemetry-derived replay labels were treated as independent accuracy labels");
        Assert(negative.GoalMode == "continuous" && negative.RecurringWorkflowIntent, "goal cycle context was not preserved in replay telemetry");
        var replayState0 = Convert.FromBase64String(positive.Frames[0].ScreenFingerprintBase64);
        var replayState3 = Convert.FromBase64String(positive.Frames[3].ScreenFingerprintBase64);
        Assert(replayState0.SequenceEqual(replayState3), $"positive replay lost a repeated screen state ({string.Join(',', positive.Frames.Select(item => Convert.FromBase64String(item.ScreenFingerprintBase64)[0]))})");
        Assert(positive.Frames.Select(item => item.ActiveProcess).Distinct(StringComparer.Ordinal).Count() == 1, "positive replay lost stable process context");
        var positiveResult = RDPilot.Diagnostics.LoopReplayService.ReplayCase(positive);
        Assert(positiveResult.Detected, $"generated positive case was not replayable by the detector (step={positiveResult.Step}, confidence={positiveResult.Confidence:0.00})");
        Assert(!RDPilot.Diagnostics.LoopReplayService.ReplayCase(negative).Detected, "generated negative case produced a loop during replay");
    }

    static void TestIndependentReplayImport()
    {
        var root = typeof(RDPilotApplication);
        var corpusPathField = TestState.Field(root, "LoopReplayCorpusPath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = corpusPathField.GetValue(null);
        var tempRoot = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        try
        {
            var source = Path.Combine(tempRoot, "reviewed.json");
            var destination = Path.Combine(tempRoot, "merged.json");
            var reviewed = new LoopReplayCorpus
            {
                Cases = [new LoopReplayCase
                {
                    Name = "reviewed-negative",
                    LabelSource = "manual:reviewed",
                    ExpectedLoop = false,
                    Frames = [new LoopReplayFrame
                    {
                        ScreenFingerprintBase64 = Convert.ToBase64String(Fingerprint(10))
                    }, new LoopReplayFrame
                    {
                        ScreenFingerprintBase64 = Convert.ToBase64String(Fingerprint(20))
                    }

                    ]
                }

                ]
            };
            File.WriteAllText(source, System.Text.Json.JsonSerializer.Serialize(reviewed));
            corpusPathField.SetValue(null, destination);
            RDPilot.Diagnostics.LoopReplayService.ImportIndependentLoopReplayCorpus(source);
            var merged = System.Text.Json.JsonSerializer.Deserialize<LoopReplayCorpus>(File.ReadAllText(destination), new System.Text.Json.JsonSerializerOptions(System.Text.Json.JsonSerializerDefaults.Web));
            Assert(merged?.Cases.Single().HasIndependentLabel == true, "reviewed replay case was not imported as an independent label");
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
        var pathField = TestState.Field(root, "RecoveryMemoryPath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var enabledField = TestState.Field(root, "RecoveryMemoryEnabled", BindingFlags.Static | BindingFlags.NonPublic)!;
        var memoryType = typeof(RDPilot.Persistence.RecoveryMemoryService);
        var readOnlyField = TestState.Field(memoryType, "RecoveryMemoryReadOnly", BindingFlags.Static | BindingFlags.NonPublic)!;
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
            var lessons = RDPilot.Persistence.RecoveryMemoryService.LoadRecoveryLessons();
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
            var staleWriterView = new List<RecoveryLesson>
            {
                concurrentLesson
            };
            InvokePrivate(memoryType, "SaveRecoveryLessons", staleWriterView);
            var merged = RDPilot.Persistence.RecoveryMemoryService.LoadRecoveryLessons();
            Assert(merged.Any(item => item.Id == lesson.Id) && merged.Any(item => item.Id == concurrentLesson.Id), "merge-on-write lost an entry from another writer");
            var exportPath = Path.Combine(tempRoot, "exported-memory.json");
            RDPilot.Persistence.RecoveryMemoryService.ExecuteRecoveryMemoryMaintenance("export", exportPath);
            Assert(File.Exists(exportPath), "memory export was not created");
            lesson.SuccessCount++;
            lesson.UpdatedUtc = DateTime.UtcNow.AddSeconds(1);
            InvokePrivate(memoryType, "SaveRecoveryLessons", lessons);
            Assert(File.Exists(path + ".bak"), "backup was not created");
            File.WriteAllText(path, "{broken-json");
            var restored = RDPilot.Persistence.RecoveryMemoryService.LoadRecoveryLessons();
            Assert(restored.Any(item => item.Id == lesson.Id), "backup did not restore the lesson");
            File.WriteAllText(path, """{"version":999,"lessons":[],"calibration":{}}""");
            _ = RDPilot.Persistence.RecoveryMemoryService.LoadRecoveryLessons();
            InvokePrivate(memoryType, "SaveRecoveryLessons", new List<RecoveryLesson>());
            Assert(File.ReadAllText(path).Contains("\"version\":999", StringComparison.Ordinal), "a newer memory schema was downgraded");
        }
        finally
        {
            pathField.SetValue(null, originalPath);
            enabledField.SetValue(null, originalEnabled);
            readOnlyField.SetValue(null, originalReadOnly);
            var expectedParent = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "RDPilotSelfTests"));
            var resolved = Path.GetFullPath(tempRoot);
            if (resolved.StartsWith(expectedParent, StringComparison.OrdinalIgnoreCase) && Directory.Exists(resolved))
            {
                Directory.Delete(resolved, recursive: true);
            }
        }
    }
}
