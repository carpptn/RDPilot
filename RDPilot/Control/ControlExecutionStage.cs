namespace RDPilot.Control;

/// <summary>Validates the proposed action and executes it under desktop safety guards.</summary>
internal static class ControlExecutionStage
{
    internal static async Task<StepFlow> RunAsync(ControlRunState run, ControlStepState turn)
    {
        if (CancelRequested || run.Cancellation.IsCancellationRequested)
            return ControlActionOutcome.Cancel(run, turn);

        turn.currentAction = new ResolvedActionSnapshot(turn.action, turn.action.Type, turn.action.Type, null);
        try
        {
            ActionDto modelResponseAction = turn.action;
            turn.currentAction = AttachFocusedTextObservationRegion(CaptureResolvedAction(turn.action, run.lastAimRect), turn.focusUiaRect);
            turn.currentActionPolicy = run.observationSession.ResolveActionPolicy(turn.currentAction, turn.promptContext, run.goal);
            turn.currentAction = AttachTurnBasedObservationRegion(
                turn.currentAction,
                ResolveTurnBasedObservationRegion(run.turnBasedInteractionRect, turn.currentActionPolicy),
                turn.currentActionPolicy);
            bool isTurnBasedAction = string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal);
            turn.plannedInputsForExecution ??= run.shortTermPlan.PrepareInputsForProposal(
                turn.action, isTurnBasedAction ? run.turnBasedTransitions.CurrentStateId : null);
            if (isTurnBasedAction &&
                turn.decision.FollowUps.Count == 0 &&
                turn.action.Type is "click" or "double_click" &&
                turn.plannedInputsForExecution is { Count: > 1 } observedPlannedInputs &&
                BatchPlanner.TryBindObservedTurnActionToPlannedInput(turn.action, observedPlannedInputs[0]) &&
                run.shortTermPlan.TryExpandDirectionalSequence(
                [observedPlannedInputs[0]],
                run.turnBasedTransitions.CurrentStateId,
                run.turnBasedTransitions.MaxExecutionBatchLength,
                out var expandedObservedInputs) &&
                run.turnBasedTransitions.TryBuildObservedDirectionalFollowUps(turn.action, expandedObservedInputs.Skip(1).ToArray(), out var inferredObservedFollowUps))
            {
                foreach (var followUp in inferredObservedFollowUps)
                    turn.decision.FollowUps.Enqueue(followUp);
                Console.WriteLine($"[turn-plan] expanded one observed click into {expandedObservedInputs.Length} route inputs using learned control positions.");
            }

            if (isTurnBasedAction &&
                CanExpandTurnKeyActionFromPlan(turn.action, turn.decision.FollowUps.Count) &&
                turn.action.Keys is { Length: > 0 } proposedKeys &&
                TryGetTurnBasedDirectionalSequenceLength(turn.action, out _) &&
                run.turnBasedTransitions.CanUseExecutionBatch &&
                run.shortTermPlan.TryExpandDirectionalSequence(
                proposedKeys,
                run.turnBasedTransitions.CurrentStateId,
                run.turnBasedTransitions.MaxExecutionBatchLength,
                out var expandedKeys))
            {
                Console.WriteLine($"[turn-plan] expanded model batch {proposedKeys.Length}->{expandedKeys.Length} from the high-confidence structured route.");
                turn.action.Keys = expandedKeys;
                turn.currentAction = AttachTurnBasedObservationRegion(
                    AttachFocusedTextObservationRegion(CaptureResolvedAction(turn.action, run.lastAimRect), turn.focusUiaRect),
                    ResolveTurnBasedObservationRegion(run.turnBasedInteractionRect, turn.currentActionPolicy),
                    turn.currentActionPolicy);
            }

            Console.WriteLine($"[{turn.step}] {turn.currentAction.Description}");
            if (turn.action.Confidence is double confidence)
                Console.WriteLine($"     confidence: {confidence:0.##}");
            if (!string.IsNullOrWhiteSpace(turn.action.Note))
                Console.WriteLine($"     note: {turn.action.Note}");
            run.nextFocusRect = null; // reset – set by aim/point/request_crop
            turn.actionExecutionFailed = false;
            turn.actionExecuted = false;
            turn.actionWasLocallyRejected = false;
            turn.turnBatchObserved = false;
            turn.turnReanalysisWasRequired = run.turnBasedTransitions.RequiresReanalysis;
            turn.actionBaselineFrame = null;
            turn.actionBaselineContext = null;
            turn.actionBaselineLocalFingerprint = null;
            if (!turn.currentAction.IsValid)
                throw new InvalidOperationException(turn.currentAction.ValidationError);
            if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                IsStateChangingInteractionAction(turn.currentAction.Action) &&
                run.turnBasedInteractionWindow != IntPtr.Zero &&
                GetForegroundWindow() != run.turnBasedInteractionWindow &&
                !TryRestoreForegroundWindow(run.turnBasedInteractionWindow))
            {
                throw new InvalidOperationException("the remembered turn-based interaction window could not regain focus");
            }

            if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                TryGetTurnBasedDirectionalSequenceLength(turn.currentAction.Action, out var directionalSequenceLength) &&
                directionalSequenceLength > 1 &&
                (!run.turnBasedTransitions.CanUseExecutionBatch ||
                directionalSequenceLength > run.turnBasedTransitions.MaxExecutionBatchLength))
            {
                throw new InvalidOperationException(run.turnBasedTransitions.CanUseExecutionBatch

                    ? $"turn-based execution batches currently accept at most {run.turnBasedTransitions.MaxExecutionBatchLength} directional inputs"
                    : "multiple directional inputs are blocked until turn-based exploration is complete");
            }

            if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                TryGetTurnBasedDirectionalSequenceLength(turn.currentAction.Action, out _) &&
                turn.currentAction.Action.Keys is { Length: > 0 } routeKeys &&
                !run.shortTermPlan.ProposedSequenceMatches(routeKeys, run.turnBasedTransitions.CurrentStateId, out var routeMismatchReason))
            {
                run.shortTermPlan.Invalidate(routeMismatchReason);
                throw new InvalidOperationException(routeMismatchReason);
            }

            if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                run.turnBasedTransitions.RequiresReanalysis &&
                IsStateChangingInteractionAction(turn.currentAction.Action) &&
                !run.turnBasedTransitions.HasRequiredSalientObservation(turn.action))
            {
                const string missingObservationReason = "a salient world change must be described in salient_change_observation before another state-changing input";
                Console.WriteLine($"[guard] turn reanalysis action blocked: {missingObservationReason}");
                AddHistory(
                    run.historyBuffer,
                    $"[{turn.step}] IGNORED (salient_change_not_observed): {turn.currentAction.Description}");
                run.lastExecutorFailure = missingObservationReason;
                turn.decision.FollowUps.Clear();
                run.RegisterRejectedProposal(turn.currentAction);
                run.prevAction = null;
                run.prevShotFingerprint = null;
                run.prevActiveWindowFingerprint = null;
                run.prevObservationFrame = null;
                run.prevObservationContext = null;
                return ControlActionOutcome.Fail(run, turn, run.lastExecutorFailure!) == StepFlow.Stop
                    ? StepFlow.Stop : StepFlow.Retry;
            }

            var blockReason = ActionExecutionPolicy.BlockReason(run, turn.step, turn.currentAction, turn.currentActionPolicy);
            if (blockReason is not null)
            {
                Console.WriteLine($"[guard] action blocked: {blockReason}");
                AddHistory(run.historyBuffer, $"[{turn.step}] IGNORED: {turn.currentAction.Description}");
                run.RegisterRejectedProposal(turn.currentAction);
                run.prevAction = null;
                run.prevShotFingerprint = null;
                run.prevActiveWindowFingerprint = null;
                run.prevObservationFrame = null;
                run.prevObservationContext = null;
                return ControlActionOutcome.Fail(run, turn, blockReason) == StepFlow.Stop
                    ? StepFlow.Stop : StepFlow.Retry;
            }

            if (ActionNeedsEffectObservation(turn.action) &&
                (!IsMouseAction(turn.action) || Options.MouseEnabled))
            {
                turn.actionBaselineFrame = CaptureObservationFrameProbe(turn.currentAction.ObservationRegion is not null ||
                    turn.action.Type is "drag_drop" or "drag_path" or "type_text" or "paste_text");
                if (turn.currentAction.ObservationRegion is Rectangle localRegion)
                {
                    turn.actionBaselineLocalFingerprint = CaptureRegionFingerprintProbe(localRegion);
                }

                if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                    run.turnBasedInteractionRect is Rectangle persistentTurnRegion)
                {
                    var baselineAssessment = run.turnBasedTransitions.PrepareActionBaseline(turn.actionBaselineFrame, persistentTurnRegion);
                    if (baselineAssessment.ExternalStateChange)
                    {
                        Console.WriteLine($"[turn] external state change detected while awaiting the model; discarding stale action; {baselineAssessment.Summary}");
                        AddHistory(
                            run.historyBuffer,
                            $"[{turn.step}] stale_action_discarded: external state changed while model was planning");
                        turn.decision.FollowUps.Clear();
                        run.BeginExternalStateEpoch();
                        run.lastExecutorFailure = "The interaction state changed externally while the model was planning. Reinspect the fresh state before acting.";
                        run.prevAction = null;
                        run.prevShotFingerprint = null;
                        run.prevActiveWindowFingerprint = null;
                        run.prevObservationFrame = null;
                        run.prevObservationContext = null;
                        return StepFlow.Retry;
                    }
                }
                else if (ShouldCheckPreRegionTurnActionFreshness(turn.currentActionPolicy, run.turnBasedInteractionRect, turn.currentAction.Action))
                {
                    await RuntimeClock.Delay(120, run.Cancellation);
                    var confirmationFrame = CaptureObservationFrameProbe(includeDetail: false);
                    if (ShouldDiscardPreRegionTurnAction(
                        turn.observationFrame,
                        turn.actionBaselineFrame,
                        confirmationFrame,
                        out var promptDelta,
                        out var stabilityDelta))
                    {
                        Console.WriteLine($"[turn] pre-region state changed while awaiting the model; discarding stale action; prompt_delta={promptDelta:0.####}; stability_delta={stabilityDelta:0.####}");
                        AddHistory(
                            run.historyBuffer,
                            $"[{turn.step}] stale_action_discarded: pre-region screen changed while model was planning");
                        turn.decision.FollowUps.Clear();
                        run.lastExecutorFailure = "The visible turn-based state changed while the model was planning. Reinspect the fresh screen and act on its current controls.";
                        run.prevAction = null;
                        run.prevShotFingerprint = null;
                        run.prevActiveWindowFingerprint = null;
                        run.prevObservationFrame = null;
                        run.prevObservationContext = null;
                        return StepFlow.Retry;
                    }

                    turn.actionBaselineFrame = confirmationFrame;
                }

                run.observationSession.RecordAmbientMotion(turn.observationFrame, turn.actionBaselineFrame);
            }

            turn.actionBaselineContext = ActionObservationCapture.Context(turn.actionBaselineFrame ?? turn.observationFrame);

            if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal))
            {
                run.turnBasedTransitions.UpdateWorkingMemory(modelResponseAction);
            }

            if (turn.action.Type == "aim")
            {
                var rect = ResolveAimRect(turn.action);
                if (rect is null)
                    throw new InvalidOperationException("aim without parameters (bbox/crop/x/y/x_px/y_px).");
                run.SetAim(rect.Value); // show crop/overlay on next screenshot
                turn.actionExecuted = true;
            }
            else if (turn.action.Type == "point")
            {
                // Visual pointer only – does NOT set AIM (clicks still blocked without 'aim')
                var rect = ResolveCropRect(turn.action);
                if (rect is null)
                    throw new InvalidOperationException("point without parameters.");
                run.nextFocusRect = rect.Value;
                turn.actionExecuted = true;
            }
            else if (turn.action.Type == "request_crop")
            {
                var rect = ResolveCropRect(turn.action);
                if (rect is null)
                    throw new InvalidOperationException("request_crop without parameters.");
                run.nextFocusRect = rect.Value;
                if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal))
                {
                    var clamped = ClampRect(rect.Value);
                    var canReplaceTurnRegion = CanReplaceTurnBasedInteractionRegion(
                        run.turnBasedInteractionRect is not null,
                        run.turnBasedInteractionRegionIsAutomatic,
                        run.turnBasedAutomaticRegionRefined,
                        turn.action,
                        regionRequired: false);
                    if (run.turnBasedInteractionRect is Rectangle existingTurnRegion && !canReplaceTurnRegion)
                    {
                        var inspectionKind = IsOverlappingTurnInspection(existingTurnRegion, clamped) ? "overlapping" : "auxiliary";
                        Console.WriteLine($"[turn] transient {inspectionKind} inspection crop=({clamped.Left},{clamped.Top})–({clamped.Right},{clamped.Bottom}); preserving primary interaction region, transitions, and visual memory.");
                    }
                    else if (canReplaceTurnRegion)
                    {
                        run.turnBasedInteractionRect = clamped;
                        run.turnBasedInteractionRegionIsAutomatic = false;
                        run.turnBasedAutomaticRegionRefined = false;
                        run.turnBasedInteractionWindow = GetForegroundWindow();
                        run.turnBasedInteractionContext = turn.currentTurnBasedContext ?? TurnInteractionContext.Capture(turn.promptContext);
                        run.previousTurnFocusDataUrl = null;
                        run.previousTurnFocusPath = null;
                        run.turnReferenceFocusDataUrl = null;
                        run.turnReferenceFocusPath = null;
                        run.activeTurnChangeImages = [];
                        run.turnBasedTransitions.Reset();
                        Console.WriteLine($"[turn] persistent interaction region set to ({clamped.Left},{clamped.Top})–({clamped.Right},{clamped.Bottom}).");
                    }
                }

                turn.actionExecuted = true;
            }
            else if (turn.action.Type == "wait")
            {
                int secs = EffectiveWaitSeconds(turn.action, out var requestedSecs);
                if (secs < requestedSecs)
                    Console.WriteLine($"[wait] Requested {requestedSecs}s capped to {secs}s.");
                Console.WriteLine($"[wait] Sleeping {secs} s (long-running operation on screen)...");
                await RuntimeClock.Delay(secs * 1000, run.Cancellation);
                turn.actionExecuted = true;
            }
            else if (turn.action.Type == "done")
            {
                return await GoalCompletionService.RunAsync(run, turn);
            }
            else if (turn.action.Type is "drag_drop" or "drag_path")
            {
                if (turn.action.Type == "drag_drop" &&
                    (!HasExplicitPoint(turn.action) || !HasExplicitDropPoint(turn.action)))
                    throw new InvalidOperationException("drag_drop requires an explicit source and destination.");
                if (turn.action.Type == "drag_path" &&
                    (turn.action.Path is null || turn.action.Path.Length < 2))
                    throw new InvalidOperationException("drag_path requires at least two path points.");
                var source = turn.action.Type == "drag_path"

                    ? ResolveGesturePath(turn.action)[0]

                    : new Point(ResolvePoint(turn.action).X, ResolvePoint(turn.action).Y);
                if (run.lastAimRect is null && !Options.DirectClickWithoutAim)
                {
                    Console.WriteLine($"[guard] {turn.action.Type} blocked: no active source AIM. Return 'aim' for the source first.");
                    AddHistory(run.historyBuffer, $"[{turn.step}] IGNORED (gesture_without_aim)");
                    run.lastExecutorFailure = $"{turn.action.Type} was blocked because its source had no active AIM";
                    turn.actionExecutionFailed = true;
                    turn.actionWasLocallyRejected = true;
                }
                else if (run.lastAimRect is Rectangle dragAim && !dragAim.Contains(source.X, source.Y))
                {
                    Console.WriteLine($"[guard] {turn.action.Type} source outside active AIM → ignoring. Set AIM around the gesture start.");
                    AddHistory(run.historyBuffer, $"[{turn.step}] IGNORED (gesture_source_outside_aim)");
                    run.lastExecutorFailure = $"{turn.action.Type} was blocked because its source was outside the active AIM";
                    turn.actionExecutionFailed = true;
                    turn.actionWasLocallyRejected = true;
                }
                else
                {
                    ExecuteAction(turn.action);
                    run.InvalidateAim();
                    turn.actionExecuted = true;
                }
            }
            else if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                turn.actionBaselineFrame is not null &&
                turn.currentAction.ObservationRegion is Rectangle observedTurnRegion &&
                TryTakeQueuedObservedTurnActions(turn.decision.FollowUps, turn.currentAction, observedTurnRegion, out var observedTurnActions, turn.plannedInputsForExecution))
            {
                var observedBatchResult = await ExecuteObservedTurnActionSequenceAsync(
                    observedTurnActions,
                    turn.actionBaselineFrame,
                    observedTurnRegion,
                    run.turnBasedTransitions,
                    run.turnBasedInteractionWindow,
                    run.Cancellation,
                    snapshot => ActionExecutionPolicy.EnsureAllowed(run, turn.step, snapshot, turn.currentActionPolicy));
                return ControlActionOutcome.ApplyObservedBatch(run, turn, observedBatchResult);
            }
            else if (turn.action.Type is "click" or "double_click")
            {
                var hasExplicitClickPoint = HasExplicitPoint(turn.action);
                var explicitClickOutsideAim = false;
                if (hasExplicitClickPoint && run.lastAimRect is Rectangle activeAim)
                {
                    var explicitPoint = ResolveClickPoint(turn.action, activeAim, logAdjustment: false);
                    explicitClickOutsideAim = !activeAim.Contains(explicitPoint.X, explicitPoint.Y);
                }

                if (Options.DirectClickWithoutAim &&
                    hasExplicitClickPoint &&
                    (run.lastAimRect is null ||
                    explicitClickOutsideAim))
                {
                    if (explicitClickOutsideAim)
                    {
                        Console.WriteLine("[aim] superseded by a new explicit click target after the model revised its intended interaction.");
                        run.InvalidateAim();
                    }

                    Console.WriteLine("[guard] direct explicit click allowed by profile.");
                    ExecuteAction(turn.action);
                    turn.actionExecuted = true;
                }
                else if (run.lastAimRect is null)
                {
                    Console.WriteLine("[guard] click blocked: no active AIM. Return 'aim' first.");
                    AddHistory(run.historyBuffer, $"[{turn.step}] IGNORED (click_without_aim)");
                    run.lastExecutorFailure = "click was blocked because there was no active AIM";
                    turn.actionExecutionFailed = true;
                    turn.actionWasLocallyRejected = true;
                }
                else
                {
                    // Clicks must include explicit coordinates (don't default to AIM center)
                    if (!HasExplicitPoint(turn.action))
                    {
                        Console.WriteLine("[guard] click/double_click requires explicit coordinates (x/y or x_px/y_px) – provide an exact point within AIM.");
                        AddHistory(run.historyBuffer, $"[{turn.step}] IGNORED (click_missing_coords)");
                        run.lastExecutorFailure = "click was blocked because explicit coordinates were missing";
                        turn.actionExecutionFailed = true;
                        turn.actionWasLocallyRejected = true;
                    }
                    else
                    {
                        var (xClick, yClick) = ResolveClickPoint(turn.action, run.lastAimRect, logAdjustment: false);
                        if (!run.lastAimRect.Value.Contains(xClick, yClick))
                        {
                            Console.WriteLine("[guard] click outside active AIM → ignoring. Set a proper 'aim' first.");
                            AddHistory(run.historyBuffer, $"[{turn.step}] IGNORED (click_outside_aim)");
                            run.lastExecutorFailure = "click was blocked because its point was outside the active AIM";
                            turn.actionExecutionFailed = true;
                            turn.actionWasLocallyRejected = true;
                        }
                        else
                        {
                            ExecuteAction(turn.action, run.lastAimRect);
                            turn.actionExecuted = true;
                        }
                    }
                }
            }
            else if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                TryGetTurnBasedDirectionalSequenceLength(turn.action, out var turnSequenceLength) &&
                turnSequenceLength > 1)
            {
                if (turn.actionBaselineFrame is null ||
                    turn.currentAction.ObservationRegion is not Rectangle turnBatchRegion)
                {
                    throw new InvalidOperationException("turn-based sequence requires a current interaction baseline");
                }

                var turnBatchResult = await ExecuteTurnBasedDirectionalSequenceAsync(
                    turn.action.Keys!,
                    turn.actionBaselineFrame,
                    turnBatchRegion,
                    run.turnBasedTransitions,
                    run.turnBasedInteractionWindow,
                    run.Cancellation,
                    snapshot => ActionExecutionPolicy.EnsureAllowed(run, turn.step, snapshot, turn.currentActionPolicy));
                return ControlActionOutcome.ApplyObservedBatch(run, turn, turnBatchResult);
            }
            else
            {
                // Move/scroll/keys/type_text – execute normally
                ExecuteAction(turn.action);
                turn.actionExecuted = true;
            }
        }
        catch (Exception ex)
        {
            if (ex is OperationCanceledException && (CancelRequested || run.Cancellation.IsCancellationRequested))
            {
                return ControlActionOutcome.Cancel(run, turn);
            }

            Console.WriteLine($"Action execution error: {ex.Message}");
            return ControlActionOutcome.Fail(run, turn, $"{turn.currentAction.Description} failed: {ex.Message}");
        }

        return StepFlow.Next;
    }
}
