using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static void TestTurnBasedLocalTransitions()
    {
        RDPilot.Configuration.ConfigurationService.ApplyObservationMode("auto", "self-test");
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
        ScreenObservationFrame SolidFrame(byte value) => new(Enumerable.Repeat(value, 64).ToArray(), Enumerable.Repeat(value, 64).ToArray(), [], 0, 0, new Rectangle(0, 0, 100, 100));
        var promptLoading = SolidFrame(0);
        var settledStart = SolidFrame(2);
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldCheckPreRegionTurnActionFreshness("turn_based_interaction", null, new ActionDto { Type = "keys", Keys = ["Space"] }), "a reversible semantic key was discarded solely because loading completed");
        Assert(RDPilot.Control.ControlActionPolicy.ShouldCheckPreRegionTurnActionFreshness("turn_based_interaction", null, new ActionDto { Type = "click", X = 0.5, Y = 0.5 }), "a coordinate-dependent click bypasses stale-screen protection");
        Assert(RDPilot.Control.ControlActionPolicy.ShouldDiscardPreRegionTurnAction(promptLoading, settledStart, settledStart, out var stalePromptDelta, out var staleStabilityDelta) && stalePromptDelta > 0.0035 && staleStabilityDelta == 0, "a settled Loading-to-Start transition did not invalidate the stale action");
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldDiscardPreRegionTurnAction(promptLoading, settledStart, SolidFrame(4), out _, out _), "ongoing animation was mistaken for a settled external state transition");
        Assert(RDPilot.Observation.ControlObservationService.IsTurnBasedStateInspection(new ActionDto { Type = "request_crop", Note = "przybliżam planszę poziomu" }), "Polish board crop was not recognized as a persistent turn-based region");
        Assert(RDPilot.Observation.ControlObservationService.IsTurnBasedStateInspection(new ActionDto { Type = "request_crop", Note = "Inspect the level playfield before testing one control." }), "playfield crop was rejected because its note mentioned a later control test");
        Assert(!RDPilot.Observation.ControlObservationService.IsTurnBasedStateInspection(new ActionDto { Type = "request_crop", Note = "inspect control buttons" }), "controls crop incorrectly replaced the persistent board region");
        Assert(!RDPilot.Observation.ControlObservationService.IsTurnBasedStateInspection(new ActionDto { Type = "request_crop", Note = "przybliż sterowanie pod planszą" }), "Polish controls crop incorrectly replaced the persistent board region");
        Assert(RDPilot.Observation.ControlObservationService.ShouldEstablishTurnBasedInteractionRegion(new ActionDto { Type = "request_crop", Note = "Inspekcja głównego pola gry przed pierwszym ruchem" }, regionRequired: true), "required primary region still depended on a recognized language form");
        Assert(RDPilot.Observation.ControlObservationService.ShouldEstablishTurnBasedInteractionRegion(new ActionDto { Type = "request_crop", Note = "Examine the requested primary interaction surface" }, regionRequired: true), "required primary region still depended on known board vocabulary");
        Assert(!RDPilot.Observation.ControlObservationService.ShouldEstablishTurnBasedInteractionRegion(new ActionDto { Type = "request_crop", Note = "inspect the help and control panel" }, regionRequired: true), "an explicitly auxiliary crop became the required primary region");
        var nestedBoardCrop = new ActionDto
        {
            Type = "request_crop",
            Note = "Powiększam centralne rozgałęzienie planszy."
        };
        Assert(RDPilot.Observation.ControlObservationService.CanReplaceTurnBasedInteractionRegion(hasExistingRegion: true, existingRegionIsAutomatic: true, automaticRegionWasRefined: false, nestedBoardCrop, regionRequired: false) && !RDPilot.Observation.ControlObservationService.CanReplaceTurnBasedInteractionRegion(hasExistingRegion: true, existingRegionIsAutomatic: true, automaticRegionWasRefined: true, nestedBoardCrop, regionRequired: false) && !RDPilot.Observation.ControlObservationService.CanReplaceTurnBasedInteractionRegion(hasExistingRegion: true, existingRegionIsAutomatic: false, automaticRegionWasRefined: false, nestedBoardCrop, regionRequired: false), "an inspection crop could replace an already established/refined primary board region");
        var region = new Rectangle(10, 10, 80, 80);
        ResolvedActionSnapshot Snapshot(string key) => RDPilot.Observation.ControlObservationService.AttachTurnBasedObservationRegion(RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "keys", Keys = [key], Note = $"move the level block with {key}" }, null), region, "turn_based_interaction");
        ScreenObservationFrame Frame(int blockX, int blockY, int timerEnd)
        {
            var detail = Enumerable.Repeat((byte)230, 10_000).ToArray();
            for (var y = blockY; y < blockY + 12; y++)
                for (var x = blockX; x < blockX + 12; x++)
                    detail[y * 100 + x] = 20;
            for (var x = 15; x < timerEnd; x++)
                detail[84 * 100 + x] = 80;
            var global = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
            return new ScreenObservationFrame(global, global, detail, 100, 100, new Rectangle(0, 0, 100, 100));
        }

        ObservationAssessment TurnAssessment(ActionOutcomeState outcome, VisualChangeState visual) => new("turn_based_interaction", visual, outcome, GoalProgressState.Neutral, 0.9, outcome == ActionOutcomeState.NoEffect ? 0 : 0.02, 0, 0, outcome == ActionOutcomeState.NoEffect ? 0 : 0.02, outcome == ActionOutcomeState.NoEffect ? 0 : 0.02, 0.004, "turn-based test")
        {
            ActionPolicy = "turn_based_interaction"
        };
        var snapshot = Snapshot("ArrowLeft");
        Assert(snapshot.ObservationRegion == region, "turn-based input did not retain the board observation region");
        Assert(RDPilot.Control.ControlActionPolicy.TryGetTurnBasedDirectionalSequenceLength(new ActionDto { Type = "keys", Keys = ["ArrowLeft", "ArrowUp"] }, out var sequenceLength) && sequenceLength == 2, "ordered directional key sequence was not recognized");
        Assert(RDPilot.Control.ControlActionPolicy.CanExpandTurnKeyActionFromPlan(new ActionDto { Type = "keys", Keys = ["ArrowLeft"] }, queuedFollowUpCount: 0) && !RDPilot.Control.ControlActionPolicy.CanExpandTurnKeyActionFromPlan(new ActionDto { Type = "keys", Keys = ["ArrowLeft"] }, queuedFollowUpCount: 13), "the first key of an already queued route could be expanded into a duplicate full route");
        Assert(RDPilot.Control.BatchPlanner.TryNormalizeDirectionalLabel("Continue upward through the corridor.", out var upwardLabel) && upwardLabel == "ArrowUp" && RDPilot.Control.BatchPlanner.TryNormalizeDirectionalLabel("Test the downward branch.", out var downwardLabel) && downwardLabel == "ArrowDown" && RDPilot.Control.BatchPlanner.TryNormalizeDirectionalLabel("Move leftwards, then inspect.", out var leftwardLabel) && leftwardLabel == "ArrowLeft" && RDPilot.Control.BatchPlanner.TryNormalizeDirectionalLabel("Proceed rightward.", out var rightwardLabel) && rightwardLabel == "ArrowRight", "natural English directional notes were not normalized to authoritative turn inputs");
        Assert(RDPilot.Control.BatchPlanner.TryNormalizeUnambiguousDirectionalLabel("Po zablokowanym lewo testuję ruch w górę.", out _) == false && RDPilot.Control.BatchPlanner.TryNormalizeUnambiguousDirectionalLabel("Testuję ruch w górę.", out var unambiguousUp) && unambiguousUp == "ArrowUp", "a note mentioning a previous direction overrode the direction being executed");
        Assert(RDPilot.Control.BatchPlanner.HasImmediateDirectionalReversal(["ArrowUp", "ArrowLeft", "ArrowRight", "ArrowDown"]) && !RDPilot.Control.BatchPlanner.HasImmediateDirectionalReversal(["ArrowUp", "ArrowUp", "ArrowLeft", "ArrowLeft"]), "an incoherent route with an immediate reversal was accepted for batching");
        Assert(!RDPilot.Control.ControlActionPolicy.RequiresPrimaryTurnBasedRegion(new ActionDto { Type = "click", XPx = 50, YPx = 50 }), "an unambiguous visible control click was blocked before the primary region existed");
        Assert(!RDPilot.Control.ControlActionPolicy.RequiresPrimaryTurnBasedRegion(new ActionDto { Type = "keys", Keys = ["Space"] }), "a conventional start/confirm key was blocked before the primary region existed");
        Assert(RDPilot.Control.ControlActionPolicy.RequiresPrimaryTurnBasedRegion(new ActionDto { Type = "keys", Keys = ["ArrowLeft"] }), "a subtle directional input was allowed without a primary observation region");
        var before = Frame(56, 42, 70);
        var immediateBaseline = Frame(56, 42, 60);
        var after = Frame(44, 42, 58);
        var inferredObservationRegion = RDPilot.Observation.ControlObservationService.InferTurnBasedObservationRegion(before, after, new Rectangle(0, 0, 100, 100));
        Assert(inferredObservationRegion is Rectangle inferred && inferred.Width < 80 && inferred.Contains(50, 48) && inferred.Contains(62, 48), "automatic turn observation did not isolate the changed playfield while preserving the full model view");
        var context = new UiPromptContext("ARC-AGI-3 Task #ls20", "msedge", "level board controls", null, null, null);
        var assessment = new RDPilot.Observation.AdaptiveObservationSession().Assess(immediateBaseline, after, snapshot, context, context, "finite", "graj w grę zagadkę i przechodź kolejne poziomy");
        Assert(assessment.ActionPolicy == "turn_based_interaction", "turn-based input did not use turn-based observation");
        Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed && assessment.GoalProgress == GoalProgressState.Neutral, "turn-based movement was confused with verified goal progress");
        var tracker = new RDPilot.Control.TurnBasedTransitionTracker();
        tracker.ObserveState(before, region);
        var ambientBaseline = tracker.PrepareActionBaseline(immediateBaseline, region);
        Assert(!ambientBaseline.ExternalStateChange, "ambient timer movement was mistaken for external interaction");
        tracker.RecordTransition(immediateBaseline, after, region, snapshot, assessment);
        Assert(!tracker.RequiresReanalysis, "a timer boundary moving just beyond its learned position created a false distant event");
        var secondBaseline = Frame(44, 42, 50);
        var thirdState = Frame(44, 20, 50);
        tracker.PrepareActionBaseline(secondBaseline, region);
        tracker.RecordTransition(secondBaseline, thirdState, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        var thirdBaseline = Frame(44, 20, 40);
        var returnedState = Frame(56, 42, 40);
        tracker.PrepareActionBaseline(thirdBaseline, region);
        tracker.RecordTransition(thirdBaseline, returnedState, region, Snapshot("ArrowDown"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        tracker.RecordTransition(returnedState, returnedState, region, Snapshot("ArrowRight"), TurnAssessment(ActionOutcomeState.NoEffect, VisualChangeState.Stable));
        var summary = tracker.BuildPromptSummary();
        Assert(summary.Contains("S1 --ArrowLeft--> S2 [changed]", StringComparison.Ordinal), "turn transition ledger did not distinguish the moved state");
        Assert(summary.Contains("S2 --ArrowUp--> S3 [changed]", StringComparison.Ordinal), "turn transition ledger collapsed a distinct state");
        Assert(summary.Contains("S3 --ArrowDown--> S1 [changed]", StringComparison.Ordinal), "turn transition ledger did not recognize a return to an earlier state");
        Assert(summary.Contains("S1 --ArrowRight--> S1 [no_effect]", StringComparison.Ordinal), "turn transition ledger did not retain a legal blocked input");
        Assert(summary.Contains("TURN_TOPOLOGY", StringComparison.Ordinal) && summary.Contains("ArrowRight=blocked", StringComparison.Ordinal), "turn transition ledger did not expose its directed topology and blocked edges");
        Assert(tracker.CanUseExecutionBatch && summary.Contains("TURN_PHASE: execution_ready", StringComparison.Ordinal) && summary.Contains("TURN_HYPOTHESIS_CONTRADICTION", StringComparison.Ordinal), "a blocked transition did not preserve aggressive replanning while falsifying the active hypothesis");
        var replayAssessment = tracker.RecordBatchStep(returnedState, after, region, "ArrowLeft");
        Assert(replayAssessment.KnownTransition && replayAssessment.ContinueBatch, "a known ordinary transition could not continue an observed key sequence");
        var transientTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        transientTracker.ObserveState(returnedState, region);
        var transientOnlyAssessment = transientTracker.RecordBatchStep(returnedState, returnedState, region, "ArrowRight", actionReactionObserved: true);
        Assert(transientOnlyAssessment.NoEffect && !transientOnlyAssessment.ContinueBatch && transientOnlyAssessment.Summary.Contains("transient_reaction=true", StringComparison.Ordinal), "a transient key animation overrode the unchanged settled turn state");
        var auxiliaryOnlyTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        var auxiliaryOnlyBefore = Frame(56, 42, 70);
        var auxiliaryOnlyAfter = Frame(56, 42, 55);
        auxiliaryOnlyTracker.ObserveState(auxiliaryOnlyBefore, region);
        var auxiliaryOnlyAssessment = auxiliaryOnlyTracker.RecordBatchStep(auxiliaryOnlyBefore, auxiliaryOnlyAfter, region, "ArrowUp", actionReactionObserved: true);
        Assert(auxiliaryOnlyAssessment.NoEffect && !auxiliaryOnlyAssessment.ContinueBatch && auxiliaryOnlyAssessment.Summary.Contains("movement_evidence=false", StringComparison.Ordinal) && auxiliaryOnlyAssessment.Summary.Contains("auxiliary_only=true", StringComparison.Ordinal), "a small auxiliary HUD change was mistaken for movement of the controlled object");
        var hudCausalTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        var hudStart = Frame(56, 42, 70);
        var hudOnly = Frame(56, 42, 55);
        var movedAfterHud = Frame(44, 42, 40);
        hudCausalTracker.ObserveState(hudStart, region);
        var blockedByHud = hudCausalTracker.RecordTransition(hudStart, hudOnly, region, Snapshot("ArrowLeft"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        var movedAfterBlockedHud = hudCausalTracker.RecordTransition(hudOnly, movedAfterHud, region, Snapshot("ArrowLeft"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        var hudCausalSummary = hudCausalTracker.BuildPromptSummary();
        Assert(blockedByHud && !movedAfterBlockedHud && !hudCausalTracker.RequiresReanalysis && hudCausalSummary.Contains("S1 --ArrowLeft--> S1 [no_effect]", StringComparison.Ordinal) && hudCausalSummary.Contains("S1 --ArrowLeft--> S2 [changed]", StringComparison.Ordinal), "a recurring peripheral HUD decrement advanced the state or created a false causal event");
        var tinyRegion = new Rectangle(0, 0, 64, 64);
        ScreenObservationFrame TinyFrame(int blockX)
        {
            var detail = Enumerable.Repeat((byte)230, 64 * 64).ToArray();
            for (var y = 30; y < 32; y++)
                for (var x = blockX; x < blockX + 2; x++)
                    detail[y * 64 + x] = 20;
            var global = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
            return new ScreenObservationFrame(global, global, detail, 64, 64, tinyRegion);
        }

        var tinyBefore = TinyFrame(24);
        var tinyAfter = TinyFrame(27);
        var tinyTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        tinyTracker.ObserveState(tinyBefore, tinyRegion);
        tinyTracker.RecordTransition(tinyBefore, tinyAfter, tinyRegion, Snapshot("ArrowRight"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(tinyTracker.BuildPromptSummary().Contains("S1 --ArrowRight--> S2 [changed]", StringComparison.Ordinal), "a small spatially coherent directional move was aliased to its origin state");
        var corridorTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        var corridorState1 = Frame(44, 60, 70);
        var corridorState2 = Frame(44, 52, 70);
        var corridorState3 = Frame(44, 44, 70);
        var corridorState4 = Frame(44, 36, 70);
        var corridorState5 = Frame(44, 28, 70);
        corridorTracker.ObserveState(corridorState1, region);
        corridorTracker.RecordTransition(corridorState1, corridorState2, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        corridorTracker.RecordTransition(corridorState2, corridorState3, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(corridorTracker.CanUseExecutionBatch && corridorTracker.BuildPromptSummary().Contains("Aggressive batching is available immediately", StringComparison.Ordinal), "two predictable moves in one direction did not enable a bounded batch");
        corridorTracker.RecordTransition(corridorState3, corridorState4, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(corridorTracker.CanUseExecutionBatch, "an additional predictable move unexpectedly disabled batching");
        var corridorBatchStep = corridorTracker.RecordBatchStep(corridorState4, corridorState5, region, "ArrowUp");
        Assert(!corridorBatchStep.KnownTransition && corridorBatchStep.ContinueBatch, "a learned directional pattern stopped at every new ordinary state");
        Assert(corridorTracker.BuildPromptSummary().Contains("S4 --ArrowUp--> S5 [changed]", StringComparison.Ordinal), "successive small moves oscillated between similar historical state ids");
        var corridorBlockedStep = corridorTracker.RecordBatchStep(corridorState5, corridorState5, region, "ArrowUp");
        Assert(corridorBlockedStep.NoEffect && !corridorBlockedStep.ContinueBatch && !corridorTracker.CanUseExecutionBatch && corridorTracker.BuildPromptSummary().Contains("TURN_HYPOTHESIS_CONTRADICTION", StringComparison.Ordinal), "a confirmed blocked step did not immediately interrupt the directional batch");
        var plannedRoute = new RDPilot.Control.ShortTermPlanTracker();
        plannedRoute.Update(new ActionDto { ShortTermPlan = "Move upward twice, then turn right toward the next waypoint.", PlanStatus = "active", PlannedInputs = ["ArrowUp", "ArrowUp", "ArrowRight", "ArrowRight"], PlanWaypoint = "the next visible corridor opening", PlanStateId = "S1", PlanConfidence = 0.60 }, "S1");
        var plannedTracker = new RDPilot.Control.TurnBasedTransitionTracker(plannedRoute);
        plannedTracker.ObserveState(corridorState1, region);
        Assert(plannedTracker.CanProposeExecutionBatch && plannedTracker.CanUseExecutionBatch && plannedTracker.AdvertisedMaxExecutionBatchLength == 32 && plannedTracker.MaxExecutionBatchLength == 32, "a reversible structured route with 0.60 confidence was not executable immediately after establishing the board state");
        Assert(plannedRoute.TryExpandDirectionalSequence(["ArrowUp"], "S1", plannedTracker.MaxExecutionBatchLength, out var immediateExpandedRoute) && immediateExpandedRoute.SequenceEqual(["ArrowUp", "ArrowUp", "ArrowRight", "ArrowRight"]), "the runtime did not expand the first conservative move into the initial aggressive route");
        plannedTracker.RecordTransition(corridorState1, corridorState2, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        plannedRoute.Update(new ActionDto { ShortTermPlan = "Continue the same route toward the opening.", PlanStatus = "active", PlannedInputs = ["ArrowUp", "ArrowUp", "ArrowRight", "ArrowRight"], PlanWaypoint = "the next visible corridor opening", PlanStateId = "S2", PlanConfidence = 0.9 }, "S2");
        Assert(plannedRoute.BuildPromptSummary().Contains("CURRENT_PLAN_INDEX: 1/4", StringComparison.Ordinal), "retaining the same structured route reset application-owned progress");
        Assert(plannedTracker.CanUseExecutionBatch && plannedTracker.MaxExecutionBatchLength == 32, "one confirmed high-confidence structured move lost the progressive execution batch");
        Assert(plannedRoute.TryExpandDirectionalSequence(["ArrowUp"], "S2", plannedTracker.MaxExecutionBatchLength, out var expandedRoute) && expandedRoute.SequenceEqual(["ArrowUp", "ArrowRight", "ArrowRight"]), "a short model batch was not expanded from the remaining structured route");
        plannedTracker.RecordTransition(corridorState2, corridorState3, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(plannedTracker.CanUseExecutionBatch && plannedTracker.MaxExecutionBatchLength == 32, "two confirmed plan-backed moves did not retain progressive execution");
        var plannedTurn = plannedTracker.RecordBatchStep(corridorState3, corridorState4, region, "ArrowRight");
        Assert(plannedTurn.ContinueBatch && plannedTurn.Summary.Contains("plan_backed=true", StringComparison.Ordinal), "an active plan could not safely continue through a new ordinary direction");
        var blockedPlannedTurn = plannedTracker.RecordBatchStep(corridorState4, corridorState4, region, "ArrowRight");
        Assert(blockedPlannedTurn.NoEffect && !blockedPlannedTurn.ContinueBatch && !plannedRoute.HasActivePlan && plannedRoute.Status == "invalidated" && plannedRoute.BuildPromptSummary().Contains("SHORT_TERM_PLAN_REVISION_REASON", StringComparison.Ordinal), "a blocked planned input did not invalidate the persisted short-term plan");
        var longRoute = new RDPilot.Control.ShortTermPlanTracker();
        var longRouteTracker = new RDPilot.Control.TurnBasedTransitionTracker(longRoute);
        longRouteTracker.ObserveState(corridorState1, region);
        longRouteTracker.RecordTransition(corridorState1, corridorState2, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        longRouteTracker.RecordTransition(corridorState2, corridorState3, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        longRoute.Update(new ActionDto { ShortTermPlan = "Follow the visible route through its turn to the terminal target.", PlanStatus = "active", PlannedInputs = ["ArrowRight", "ArrowRight", "ArrowRight", "ArrowUp", "ArrowUp", "ArrowUp"], PlanWaypoint = "the visible terminal target", PlanStateId = "S3", PlanConfidence = 0.95 }, "S3");
        Assert(longRouteTracker.CanUseExecutionBatch && longRouteTracker.MaxExecutionBatchLength == 32 && longRouteTracker.PreferredAdvertisedExecutionBatchMinimum == 6, "a high-confidence semantic route did not enable an immediate extended batch");
        Assert(longRoute.TryExpandDirectionalSequence(["ArrowRight"], "S3", longRouteTracker.MaxExecutionBatchLength, out var expandedLongRoute) && expandedLongRoute.SequenceEqual(["ArrowRight", "ArrowRight", "ArrowRight", "ArrowUp", "ArrowUp", "ArrowUp"]), "the runtime did not expand a short proposal across a visible predictable turn");
        var twentyStepRoute = new RDPilot.Control.ShortTermPlanTracker();
        var twentyInputs = Enumerable.Repeat("ArrowRight", 20).ToArray();
        twentyStepRoute.Update(new ActionDto { ShortTermPlan = "Follow the long visible straight corridor.", PlanStatus = "active", PlannedInputs = twentyInputs, PlanWaypoint = "the far end of the corridor", PlanStateId = "S1", PlanConfidence = 0.95 }, "S1");
        Assert(twentyStepRoute.BuildPromptSummary().Contains("CURRENT_PLAN_INDEX: 0/20", StringComparison.Ordinal) && twentyStepRoute.TryExpandDirectionalSequence(["ArrowRight"], "S1", 32, out var expandedTwentyStepRoute) && expandedTwentyStepRoute.Length == 20, "the application silently truncated a valid structured route after twelve inputs");
        var extendedRoute = new RDPilot.Control.ShortTermPlanTracker();
        var extendedTracker = new RDPilot.Control.TurnBasedTransitionTracker(extendedRoute);
        extendedTracker.ObserveState(corridorState1, region);
        extendedTracker.RecordTransition(corridorState1, corridorState2, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        extendedTracker.RecordTransition(corridorState2, corridorState3, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        extendedTracker.RecordTransition(corridorState3, corridorState4, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        extendedTracker.RecordTransition(corridorState4, corridorState5, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(extendedTracker.CanProposeExecutionBatch && extendedTracker.AdvertisedMaxExecutionBatchLength == 32, "four confirmed directional transitions did not advertise the extended input horizon");
        extendedRoute.Update(new ActionDto { ShortTermPlan = "Follow the long visible route to its terminal target.", PlanStatus = "active", PlannedInputs = ["ArrowRight", "ArrowRight", "ArrowRight", "ArrowRight", "ArrowDown", "ArrowDown", "ArrowDown", "ArrowDown", "ArrowLeft", "ArrowLeft", "ArrowLeft", "ArrowLeft"], PlanWaypoint = "the visible terminal target", PlanStateId = "S5", PlanConfidence = 0.95 }, "S5");
        Assert(extendedTracker.CanUseExecutionBatch && extendedTracker.MaxExecutionBatchLength == 32, "a mature high-confidence route did not enable the extended input batch");
        var staleRoute = new RDPilot.Control.ShortTermPlanTracker();
        staleRoute.Update(new ActionDto { ShortTermPlan = "Follow a stale route.", PlanStatus = "active", PlannedInputs = ["ArrowUp", "ArrowLeft"], PlanWaypoint = "old waypoint", PlanStateId = "S9", PlanConfidence = 0.95 }, "S1");
        Assert(!staleRoute.HasActivePlan && staleRoute.Status == "invalidated", "a structured route bound to a stale state was accepted");
        var completedRoute = new RDPilot.Control.ShortTermPlanTracker();
        completedRoute.Update(new ActionDto { ShortTermPlan = "Reach the next corner.", PlanStatus = "active", PlannedInputs = ["ArrowUp", "ArrowRight"], PlanWaypoint = "next corner", PlanStateId = "S1", PlanConfidence = 0.9 }, "S1");
        completedRoute.RecordDirectionalResult("ArrowUp", "S1", "S2", changed: true);
        completedRoute.RecordDirectionalResult("ArrowRight", "S2", "S3", changed: true);
        Assert(!completedRoute.HasActivePlan && completedRoute.Status == "completed" && completedRoute.BuildPromptSummary().Contains("CURRENT_PLAN_INDEX: 2/2", StringComparison.Ordinal), "a fully executed structured route did not complete at its waypoint");
        tracker.UpdateWorkingMemory(new ActionDto { WorldStateSummary = "A movable block is navigating a bounded board.", MechanicsHypothesis = "Special board elements may change persistent world state.", SalientChangeObservation = "Touching the white marker changed a distant reference glyph." });
        summary = tracker.BuildPromptSummary();
        Assert(summary.Contains("TURN_WORLD_STATE_MEMORY", StringComparison.Ordinal) && summary.Contains("TURN_MECHANICS_HYPOTHESIS", StringComparison.Ordinal), "turn-based working memory was not returned to the next model request");
        tracker.UpdateWorkingMemory(new ActionDto { WorldStateSummary = "A movable block is navigating a bounded board.", MechanicsHypothesis = "The visible marker may instead be the controlled object.", SalientChangeObservation = "Entering the blue terminal changed the level." });
        summary = tracker.BuildPromptSummary();
        Assert(summary.Contains("TURN_PRIOR_MECHANICS_HYPOTHESES", StringComparison.Ordinal) && summary.Contains("Special board elements", StringComparison.Ordinal), "a revised mechanics claim erased the prior unverified hypothesis");
        tracker.BeginExternalStateEpoch();
        summary = tracker.BuildPromptSummary();
        Assert(!summary.Contains("TURN_MECHANICS_HYPOTHESIS", StringComparison.Ordinal) && !summary.Contains("TURN_PRIOR_MECHANICS_HYPOTHESES", StringComparison.Ordinal) && summary.Contains("TURN_TRANSIENT_CONTEXT_WARNING", StringComparison.Ordinal) && summary.Contains("TURN_PRIOR_EPOCH_OBSERVED_EVIDENCE", StringComparison.Ordinal) && !summary.Contains("TURN_PRIOR_EPOCH_HYPOTHESES", StringComparison.Ordinal) && !summary.Contains("Special board elements", StringComparison.Ordinal) && !summary.Contains("visible marker", StringComparison.Ordinal) && !summary.Contains("white marker changed", StringComparison.Ordinal) && summary.Contains("blue terminal changed", StringComparison.Ordinal), "an external state boundary retained stale topology/hypotheses or discarded the latest causal fact");
        tracker.UpdateWorkingMemory(new ActionDto { WorldStateSummary = "The stable next board is now visible." });
        Assert(!tracker.BuildPromptSummary().Contains("TURN_TRANSIENT_CONTEXT_WARNING", StringComparison.Ordinal), "the transient-context warning survived a committed stable-state update");
        using var referenceCrop = new Bitmap(48, 36);
        using (var graphics = Graphics.FromImage(referenceCrop))
            graphics.Clear(Color.FromArgb(40, 80, 120));
        using var matchingNativeCrop = new Bitmap(96, 72);
        using (var graphics = Graphics.FromImage(matchingNativeCrop))
            graphics.Clear(Color.FromArgb(40, 80, 120));
        using var changedNativeCrop = new Bitmap(96, 72);
        using (var graphics = Graphics.FromImage(changedNativeCrop))
            graphics.Clear(Color.FromArgb(220, 30, 20));
        Assert(RDPilot.Observation.ScreenshotService.AreTurnRegionCapturesConsistent(referenceCrop, matchingNativeCrop) && !RDPilot.Observation.ScreenshotService.AreTurnRegionCapturesConsistent(referenceCrop, changedNativeCrop), "native board crops were not checked against the primary screenshot before use");
        Assert(RDPilot.Control.TurnBasedTransitionTracker.IsBoundedAttemptResetEvidence(returnedToAttemptOrigin: true, directionalInputs: 50, broadChange: true, distantChange: true, novelRegionCount: 3, auxiliaryRegionCount: 1) && !RDPilot.Control.TurnBasedTransitionTracker.IsBoundedAttemptResetEvidence(returnedToAttemptOrigin: true, directionalInputs: 20, broadChange: true, distantChange: false, novelRegionCount: 1, auxiliaryRegionCount: 2), "a closed route was confused with broad bounded-attempt reset evidence");
        ScreenObservationFrame WithRemoteChange(ScreenObservationFrame source, int left, int top)
        {
            var detail = source.DetailFingerprint.ToArray();
            for (var y = top; y < top + 10; y++)
                for (var x = left; x < left + 10; x++)
                    detail[y * 100 + x] = 20;
            return source with
            {
                DetailFingerprint = detail
            };
        }

        var eventTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        var eventBefore = Frame(56, 42, 60);
        var eventAfter = WithRemoteChange(Frame(44, 42, 60), 18, 72);
        eventTracker.ObserveState(eventBefore, region);
        eventTracker.RecordTransition(eventBefore, eventAfter, region, Snapshot("ArrowLeft"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        var eventSummary = eventTracker.BuildPromptSummary();
        Assert(!eventTracker.RequiresReanalysis && eventTracker.CanProposeExecutionBatch && eventSummary.Contains("TURN_REANALYSIS_REQUIRED: false", StringComparison.Ordinal) && eventSummary.Contains("distant=true", StringComparison.Ordinal), "a small distant change interrupted aggressive execution instead of remaining visible evidence");
        Assert(eventTracker.HasRequiredSalientObservation(new ActionDto { Type = "keys", Keys = ["ArrowUp"] }) && eventTracker.HasRequiredSalientObservation(new ActionDto { Type = "keys", Keys = ["ArrowUp"], SalientChangeObservation = "A separate persistent board element changed shape." }), "a non-broad distant change still imposed a salient-observation barrier");
        eventTracker.AcknowledgeReanalysis();
        Assert(!eventTracker.RequiresReanalysis, "the reanalysis barrier did not clear after a fresh one-step decision");
        var auxiliaryTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        var auxiliaryState1 = Frame(56, 42, 70);
        var auxiliaryState2 = Frame(44, 42, 58);
        var auxiliaryState3 = Frame(32, 42, 50);
        auxiliaryTracker.ObserveState(auxiliaryState1, region);
        auxiliaryTracker.RecordTransition(auxiliaryState1, auxiliaryState2, region, Snapshot("ArrowLeft"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        auxiliaryTracker.AcknowledgeReanalysis();
        auxiliaryTracker.RecordTransition(auxiliaryState2, auxiliaryState3, region, Snapshot("ArrowLeft"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        var auxiliarySummary = auxiliaryTracker.BuildPromptSummary();
        Assert(!auxiliaryTracker.RequiresReanalysis && auxiliarySummary.Contains("TURN_AUXILIARY_CHANGES", StringComparison.Ordinal) && auxiliarySummary.Contains("predictable_regions=", StringComparison.Ordinal), "a recurring action-correlated auxiliary change kept forcing reanalysis");
        var auxiliaryReturn = Frame(44, 42, 42);
        auxiliaryTracker.RecordTransition(auxiliaryState3, auxiliaryReturn, region, Snapshot("ArrowRight"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        auxiliarySummary = auxiliaryTracker.BuildPromptSummary();
        Assert(auxiliarySummary.Contains("S3 --ArrowRight--> S2 [changed] [returned_to_known_state]", StringComparison.Ordinal) && auxiliarySummary.Contains("TURN_STATE_RETURN", StringComparison.Ordinal) && auxiliarySummary.Contains("TURN_AUXILIARY_SIGNAL", StringComparison.Ordinal) && auxiliarySummary.Contains("TURN_NAVIGATION_POSE: state=S2; relative=(-1,0)", StringComparison.Ordinal) && auxiliarySummary.Contains("(-2,0) S3: ArrowRight->(-1,0) S2", StringComparison.Ordinal), "a changed peripheral indicator prevented recognition of a return to the same logical board state");
        var novelWorldState = WithRemoteChange(Frame(20, 42, 42), 70, 18);
        auxiliaryTracker.RecordTransition(auxiliaryReturn, novelWorldState, region, Snapshot("ArrowLeft"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(auxiliaryTracker.RequiresReanalysis && auxiliaryTracker.SalientChangeRegions.Count > 0 && auxiliaryTracker.BuildPromptSummary().Contains("causal_distant=true", StringComparison.Ordinal), "a novel local-to-distant causal change did not interrupt execution for reanalysis");
        var broadWorldState = novelWorldState with
        {
            DetailFingerprint = Enumerable.Repeat((byte)15, 10_000).ToArray()
        };
        auxiliaryTracker.RecordTransition(novelWorldState, broadWorldState, region, Snapshot("ArrowUp"), TurnAssessment(ActionOutcomeState.Confirmed, VisualChangeState.Changed));
        Assert(auxiliaryTracker.RequiresReanalysis && auxiliaryTracker.SalientChangeRegions.Count > 0, "a broad screen transition did not stop aggressive execution");
        var externalTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        externalTracker.ObserveState(before, region);
        var externalBaseline = externalTracker.PrepareActionBaseline(after, region);
        Assert(externalBaseline.ExternalStateChange && externalTracker.RequiresReanalysis, "a state-changing input during model latency was learned as ambient motion");
        ScreenObservationFrame WithLowIntensityAnimation(ScreenObservationFrame source, params Rectangle[] animatedRegions)
        {
            var detail = source.DetailFingerprint.ToArray();
            foreach (var animatedRegion in animatedRegions)
                for (var y = animatedRegion.Top; y < animatedRegion.Bottom; y++)
                    for (var x = animatedRegion.Left; x < animatedRegion.Right; x++)
                        detail[y * 100 + x] = (byte)Math.Max(0, detail[y * 100 + x] - 18);
            return source with
            {
                DetailFingerprint = detail
            };
        }

        var animatedStartTracker = new RDPilot.Control.TurnBasedTransitionTracker();
        var animatedStart = Frame(56, 42, 70);
        var animatedStartDrift = WithLowIntensityAnimation(animatedStart, new Rectangle(38, 30, 14, 8), new Rectangle(50, 38, 14, 8), new Rectangle(36, 44, 12, 8));
        animatedStartTracker.ObserveState(animatedStart, region);
        var animatedStartBaseline = animatedStartTracker.PrepareActionBaseline(animatedStartDrift, region);
        Assert(!animatedStartBaseline.ExternalStateChange && !animatedStartTracker.RequiresReanalysis && animatedStartBaseline.Summary.Contains("cold_start_ambient_calibration", StringComparison.Ordinal), "low-intensity startup animation invalidated a still-relevant action");
        var repeatedAnimatedStartBaseline = animatedStartTracker.PrepareActionBaseline(animatedStart, region);
        Assert(!repeatedAnimatedStartBaseline.ExternalStateChange && !animatedStartTracker.RequiresReanalysis, "learned startup animation repeatedly invalidated a still-relevant action");
        Assert(RDPilot.Observation.ControlObservationService.IsOverlappingTurnInspection(new Rectangle(100, 100, 400, 400), new Rectangle(50, 50, 500, 500)), "a containing turn inspection crop was not treated as transient");
        Assert(!RDPilot.Observation.ControlObservationService.IsOverlappingTurnInspection(new Rectangle(100, 100, 200, 200), new Rectangle(350, 100, 200, 200)), "a separate board region was incorrectly treated as the same interaction crop");
    }
}
