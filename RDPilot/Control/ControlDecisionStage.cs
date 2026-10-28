namespace RDPilot.Control;

/// <summary>Builds the decision request, applies loop limits and handles model retries.</summary>
internal static class ControlDecisionStage
{
    internal static async Task<StepFlow> RunAsync(ControlRunState run, ControlStepState turn)
    {
        if (Options.MaxStagnationStepsBeforeAbort > 0 &&
            run.stagnationSteps >= Options.MaxStagnationStepsBeforeAbort)
        {
            Console.WriteLine($"[guard] stopping: no visible progress for {run.stagnationSteps} consecutive step(s). Use --max-stagnation 0 to disable.");
            run.runResult = new ControlRunResult(
                ControlRunOutcome.GuardStopped,
                turn.step,
                $"no visible progress for {run.stagnationSteps} consecutive steps");
            return StepFlow.Stop;
        }

        if (Options.MaxRepeatedActionBeforeAbort > 0 &&
            run.repeatCount >= Options.MaxRepeatedActionBeforeAbort)
        {
            Console.WriteLine($"[guard] stopping: repeated ineffective action {run.repeatCount} time(s). Use --max-repeated-actions 0 to disable.");
            run.runResult = new ControlRunResult(ControlRunOutcome.GuardStopped, turn.step, $"ineffective action repeated {run.repeatCount} times");
            return StepFlow.Stop;
        }

        if (Options.MaxRejectedProposalRepeatsBeforeAbort > 0 &&
            run.rejectedProposalCycleCount >= Options.MaxRejectedProposalRepeatsBeforeAbort)
        {
            Console.WriteLine($"[guard] stopping: rejected model proposal cycle repeated {run.rejectedProposalCycleCount} time(s); cycle_length={run.rejectedProposalCycleLength}. Use --max-rejected-proposals 0 to disable.");
            run.runResult = new ControlRunResult(
                ControlRunOutcome.GuardStopped,
                turn.step,
                $"rejected model proposal cycle repeated {run.rejectedProposalCycleCount} times");
            return StepFlow.Stop;
        }

        string recoveryMemoryPrompt = BuildRecoveryMemoryPrompt(
            run.recoveryLessons,
            turn.promptContext,
            turn.shotFingerprint,
            turn.activeWindowFingerprint,
            run.prevAction,
            run.recentExecutedActions,
            run.recoveryEpisode,
            run.stagnationSteps,
            run.repeatCount,
            run.goal);
        bool reuseUiaTargets = Options.ReuseUiaTargetsWhenScreenUnchanged &&
            run.promptVisualState.CanReuseTargets(turn, CurrentRun.CurrentUiaTargets);
        PrepareUiaTargetsForPrompt(reuseUiaTargets, turn.screenW, turn.screenH);
        run.promptVisualState.RecordTargets(turn, CurrentRun.CurrentUiaTargets);
        CurrentRun.RequestReasoningEffortOverride = EffectiveReasoningEffort(run.stagnationSteps, run.repeatCount, run.rejectedProposalCycleCount);
        if (!string.IsNullOrWhiteSpace(run.lastExecutorFailure))
            run.shortTermPlan.Invalidate($"executor rejected or failed the planned action: {TrimForMeta(run.lastExecutorFailure, 160)}");
        int controlMaxOutputTokens = EffectiveControlMaxOutputTokens(run.turnBasedTransitions.RequiresReanalysis);
        if (run.turnBasedTransitions.RequiresReanalysis &&
            controlMaxOutputTokens > Options.MaxOutputTokens)
        {
            Console.WriteLine($"[openai] salient turn reanalysis budget elevated: {Options.MaxOutputTokens}->{controlMaxOutputTokens} tokens.");
        }

        // inject observation metrics into the prompt
        StringBuilder metaSb = new StringBuilder().AppendLine($"LAST_STEP_DELTA: {(double.IsNaN(run.lastDelta) ? "N/A" : run.lastDelta.ToString("0.####"))} (observation_threshold={run.lastObservation?.ChangeThreshold ?? Options.NoChangeThreshold:0.####})")
            .AppendLine($"LAST_GLOBAL_DELTA: {(double.IsNaN(run.lastGlobalDelta) ? "N/A" : run.lastGlobalDelta.ToString("0.####"))}; LAST_ACTIVE_WINDOW_DELTA: {(double.IsNaN(run.lastActiveWindowDelta) ? "N/A" : run.lastActiveWindowDelta.ToString("0.####"))}")
            .AppendLine($"LAST_LOCAL_DELTA: {(run.lastObservation is not null && double.IsFinite(run.lastObservation.LocalDelta) ? run.lastObservation.LocalDelta.ToString("0.####") : "N/A")}; LAST_LOCAL_CHANGED_RATIO: {(run.lastObservation is null ? "N/A" : run.lastObservation.LocalChangedRatio.ToString("0.####"))}")
            .AppendLine($"OBSERVATION_PROFILE: {run.lastObservation?.Profile ?? run.observationSession.EffectiveProfile}; ACTION_OBSERVATION_POLICY: {run.lastObservation?.ActionPolicy ?? "N/A"}; VISUAL_CHANGE: {run.lastObservation?.VisualChange.ToString() ?? "N/A"}; ACTION_OUTCOME: {run.lastObservation?.ActionOutcome.ToString() ?? "N/A"}; GOAL_PROGRESS: {run.lastObservation?.GoalProgress.ToString() ?? "N/A"}; OBSERVATION_CONFIDENCE: {(run.lastObservation is null ? "N/A" : run.lastObservation.Confidence.ToString("0.00"))}")
            .AppendLine($"STAGNATION_STEPS: {run.stagnationSteps}")
            .AppendLine($"REPEAT_COUNT: {run.repeatCount}")
            .AppendLine($"REJECTED_PROPOSAL_CYCLE_COUNT: {run.rejectedProposalCycleCount}")
            .AppendLine($"CONTINUOUS_IDLE_STEPS: {run.continuousIdleSteps}")
            .AppendLine($"AMBIGUOUS_OBSERVATION_STEPS: {run.ambiguousObservationSteps}")
            .AppendLine($"INSPECTION_ACTIONS: {run.observationActionGuard.ConsecutiveInspectionActions}/{(Options.MaxConsecutiveInspectionActions > 0 ? Options.MaxConsecutiveInspectionActions.ToString() : "unlimited")}; AIM_SINCE_INTERACTION: {run.observationActionGuard.AimIssuedSinceInteraction.ToString().ToLowerInvariant()}")
            .AppendLine($"GOAL_MODE: {run.goalMode}")
            .AppendLine($"REQUEST_REASONING_EFFORT: {CurrentRun.RequestReasoningEffortOverride ?? Options.ReasoningEffort ?? "default"}")
            .AppendLine($"REQUEST_MAX_OUTPUT_TOKENS: {controlMaxOutputTokens}")
            .AppendLine($"LAST_ACTION: {(run.prevAction == null ? "N/A" : run.prevAction.Description)}")
            .AppendLine($"AIM_ACTIVE: {(run.lastAimRect is null ? "false" : $"true {FormatImageRect(CurrentRun.CurrentScreenMap.ScreenToImageRect(run.lastAimRect.Value))}")}")
            .AppendLine("CONTEXT_CHECKPOINT_VERSION: 1; the current screen and runtime state below are authoritative when they conflict with earlier model reasoning.");
        string recentActionCheckpoint = BuildRecentActionCheckpoint(run.recentExecutedActions);
        if (!string.IsNullOrWhiteSpace(recentActionCheckpoint))
            metaSb.AppendLine(recentActionCheckpoint);
        string shortTermPlanSummary = run.shortTermPlan.BuildPromptSummary();
        if (!string.IsNullOrWhiteSpace(shortTermPlanSummary))
            metaSb.AppendLine(shortTermPlanSummary);
        if (run.repeatCount > 0 || run.stagnationSteps > 0)
            metaSb.AppendLine("STRATEGY_HINT: The previous action did not visibly advance the screen. Do not repeat it; choose a different UI route or ask for a crop if the target is ambiguous.");
        if (run.prevAction is not null &&
            run.lastObservation is not null &&
            ShouldRegisterImmediateNoEffectCooldown(run.prevAction.Action, run.lastObservation, run.repeatCount))
        {
            metaSb.AppendLine("NO_EFFECT_ACTION_BLOCKED: LAST_ACTION is unavailable until the observed state changes. Do not propose it again; choose a different input modality, target, or diagnostic action.");
        }

        if (turn.expectedContinuousIdle)
            metaSb.AppendLine("CONTINUOUS_IDLE: The previous wait left the screen unchanged, which is valid for this open-ended goal. Reassess whether the requested state is still healthy or whether a new event is present; wait again only when continued observation is goal-aligned.");
        if (run.ambiguousObservationSteps > 0)
            metaSb.AppendLine("AMBIGUOUS_ACTION_OUTCOME: Visual activity was observed, but it could not be attributed confidently to the previous action. Reassess the state and expected effect; do not blindly repeat the same input.");
        if (run.observationActionGuard.RequiresInteraction(Options.MaxConsecutiveInspectionActions))
            metaSb.AppendLine("INTERACTION_REQUIRED: The inspection budget is exhausted. request_crop and point are mechanically blocked until RDPilot executes a state-changing interaction. Use visible controls or one safe, reversible input and then observe its effect; use aim only when precise pointer targeting is required.");
        else if (string.Equals(run.observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal) &&
            run.observationActionGuard.ConsecutiveInspectionActions >= 1)
            metaSb.AppendLine("TURN_INSPECTION_COMPLETE: a detailed crop has already been supplied for the current state. Another nested request_crop is blocked until a state-changing input is attempted.");
        if (string.Equals(run.observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal))
        {
            metaSb.AppendLine(run.turnBasedTransitions.CanProposeExecutionBatch

                ? $"TURN_BASED_INTERACTION: aggressive execution_ready batching is the default from the first visible route. Commit to the strongest visible control hypothesis and test it with a reversible route; semantic uncertainty is not a reason to inspect HELP, narrate, or delay. Before the first directional input, request_crop is justified only when the board is physically unreadable at the supplied resolution, and HELP/instructions are justified only when no reversible control input is visible. Preliminary single-move calibration is forbidden when at least two reversible moves form a coherent route. Bind the route to TURN_STATE with planned_inputs, plan_waypoint, and plan_confidence. When fixed visible directional controls such as a D-pad exist and keyboard focus has not been confirmed by a successful move, prefer an ordered click sequence over keyboard keys. If a keyboard direction had no effect, preserve the logical route and immediately remap its longest valid prefix to those visible controls; do not test only one button. Send up to {run.turnBasedTransitions.AdvertisedMaxExecutionBatchLength} observed click/key inputs and include every visible unconditional turn before the next semantic uncertainty; do not pad a route merely to reach the cap. A small or distant recurring auxiliary UI change is recorded but does not stop execution. A blocked input, unavailable observation, broad screen/state transition, or novel local-to-distant causal change interrupts the batch. Do not use hold_keys for discrete movement."
                : "TURN_BASED_INTERACTION: The interaction is in an exploration or reanalysis phase. Promptly perform exactly one maximally informative, reversible input, then inspect the resulting frame. Do not spend this turn narrating or solving a full route. Compare labeled temporal images and the reported change regions before acting. Prefer a visible directional/control button or a single Arrow/WASD/Space/F key when the interface suggests those controls. Keep planned_inputs null unless the route is already mechanically established. Do not use hold_keys or batch speculative moves.");
            metaSb.AppendLine("TURN_GRID_ROUTE_HINT: On a visible grid, board, or maze, trace and count the route cell by cell from the controlled object's current footprint to an evidence-backed semantic waypoint before emitting planned_inputs. Every planned input must correspond to exactly one adjacent traversable cell, every intermediate footprint must remain on the connected visible path, and the number and order of repeated directions must match the counted grid intervals. Do not estimate distance from a qualitative relation such as 'near', 'below-right', or apparent screen proximity, guess through visible barriers, select an unverified marker only because it is salient, or repeat a closed route that returned to a known state. Aggressive batching means executing the longest route that passes this geometric check decisively.");
            if (run.turnBasedTransitions.CanProposeExecutionBatch)
            {
                metaSb.AppendLine("TURN_ROUTE_DECISION_MODE: commit_fast. Do not exhaustively solve hidden mechanics before acting. Choose the strongest visible unconditional route, emit its structured key sequence immediately, and let the per-input observation barrier expose mistaken assumptions.");
            }

            var transitionSummary = run.turnBasedTransitions.BuildPromptSummary();
            if (!string.IsNullOrWhiteSpace(transitionSummary))
                metaSb.AppendLine(transitionSummary);
        }

        if (run.recoveryEpisode != null &&
            run.stagnationSteps < Options.RecoveryMemoryTriggerSteps &&
            run.repeatCount < 1)
        {
            metaSb.AppendLine("PROACTIVE_LOOP_SUSPECTED: recent actions indicate an emerging loop. Change strategy now, before a safety guard is reached.");
        }

        if (turn.proactiveVisualCycle)
            metaSb.AppendLine($"MULTI_STEP_LOOP_DETECTED: confidence={turn.loopAssessment.Confidence:0.00}; the UI returned to a prior visual state after {turn.visualCycleLength} steps with corroborating evidence. Do not repeat the intervening action sequence.");
        else if (turn.loopAssessment.IsProductiveCycle)
            metaSb.AppendLine($"PRODUCTIVE_CYCLE_OBSERVED: confidence={turn.loopAssessment.Confidence:0.00}; this state return matches an intentional recurring workflow and is not currently treated as a harmful loop. Continue only while the cycle remains goal-aligned and produces useful checks, maintenance, or event handling.");
        else if (turn.loopAssessment.Confidence >= 0.5)
            metaSb.AppendLine($"LOOP_CANDIDATE: confidence={turn.loopAssessment.Confidence:0.00}, below calibrated threshold={turn.loopAssessment.DecisionThreshold:0.00}. This is not yet a confirmed loop; seek another recurrence signal before changing a valid strategy.");
        if (run.planningLoopDetected)
            metaSb.AppendLine($"PLANNING_LOOP_DETECTED: rejected proposal cycle length={run.rejectedProposalCycleLength}, repeated={run.rejectedProposalCycleCount}. Do not propose the blocked action sequence again; choose a materially different, currently executable route.");
        if (!string.IsNullOrWhiteSpace(recoveryMemoryPrompt))
            metaSb.AppendLine(recoveryMemoryPrompt);
        if (!string.IsNullOrWhiteSpace(run.lastVerifierRejection))
            metaSb.AppendLine($"LAST_VERIFY_REJECTION: {TrimForMeta(run.lastVerifierRejection, 240)}");
        if (!string.IsNullOrWhiteSpace(run.lastExecutorFailure))
            metaSb.AppendLine($"LAST_EXECUTOR_FAILURE: {TrimForMeta(run.lastExecutorFailure, 240)}");
        if (run.lastPrecisionHintExpiresAfterStep > 0 &&
            turn.step > run.lastPrecisionHintExpiresAfterStep)
        {
            run.lastPrecisionHint = null;
            run.lastPrecisionHintExpiresAfterStep = 0;
        }

        if (!string.IsNullOrWhiteSpace(run.lastPrecisionHint))
            metaSb.AppendLine($"PRECISION_HINT: {TrimForMeta(run.lastPrecisionHint, 420)}");
        if (run.lastTextInputHintExpiresAfterStep > 0 &&
            turn.step > run.lastTextInputHintExpiresAfterStep)
        {
            run.lastTextInputHint = null;
            run.lastTextInputHintExpiresAfterStep = 0;
        }

        if (!string.IsNullOrWhiteSpace(run.lastTextInputHint))
            metaSb.AppendLine($"TEXT_INPUT_HINT: {TrimForMeta(run.lastTextInputHint, 420)}");
        if (run.textInputCooldownUntilStep >= turn.step)
            metaSb.AppendLine($"TEXT_INPUT_COOLDOWN: active until step {run.textInputCooldownUntilStep}; do not use paste_text, type_text, or paste shortcuts until focus/editability visibly changes.");
        bool omitFullScreenImage = Options.OmitUnchangedScreenImageWithState &&
            run.promptVisualState.CanOmitImage(turn);
        if (omitFullScreenImage)
            metaSb.AppendLine("SCREEN_IMAGE: omitted because screen fingerprint is unchanged; use previous_response_id state plus current metadata.");
        List<TurnChangeImagePair> turnChangeImages = new List<TurnChangeImagePair>();
        if (turn.hasPersistentTurnFocus &&
            run.turnBasedTransitions.RequiresReanalysis &&
            run.previousTurnFocusDataUrl is not null &&
            turn.currentTurnEvidenceDataUrl is not null)
        {
            var regionIndex = 1;
            foreach (var changeRegion in run.turnBasedTransitions.SalientChangeRegions)
            {
                var pair = BuildTurnChangeImagePair(
                    run.previousTurnFocusDataUrl,
                    turn.currentTurnEvidenceDataUrl,
                    changeRegion,
                    run.screensDir,
                    run.commandId,
                    turn.step,
                    regionIndex++);
                if (pair is not null)
                    turnChangeImages.Add(pair);
            }

            if (turnChangeImages.Count > 0)
            {
                run.activeTurnChangeImages = turnChangeImages;
                Console.WriteLine($"[turn-event] attached {turnChangeImages.Count} focused before/after change-region pair(s).");
            }
        }

        IReadOnlyList<TurnChangeImagePair> requestTurnChangeImages = run.turnBasedTransitions.RequiresReanalysis ||
            run.turnBasedTransitions.HasActiveCausalEvent

            ? run.activeTurnChangeImages

            : [];
        bool includeTurnTemporalImages = turn.hasTurnVisualContext &&
            (run.turnBasedTransitions.RequiresReanalysis ||
            run.turnBasedTransitions.HasActiveCausalEvent ||
            !run.turnBasedTransitions.CanUseExecutionBatch ||
            !run.shortTermPlan.HasActivePlan);
        if (turn.hasTurnVisualContext && !includeTurnTemporalImages)
        {
            metaSb.AppendLine("TURN_TEMPORAL_IMAGES: omitted during predictable plan execution; use CURRENT_FOCUS_IMAGE and the transition ledger. Full before/reference evidence will return after an invalidating or salient change.");
        }

        string? focusDataUrlForRequest = turn.focusUrl ?? turn.currentTurnEvidenceDataUrl;
        Rectangle? focusRectForRequest = turn.appliedFocusRectForPrompt ?? (turn.focusUrl is null &&
            run.turnBasedInteractionRect is Rectangle evidenceRegion

            ? ScreenRectToImage(evidenceRegion)

            : null);
        string? focusPathForRequest = turn.currentTurnFocusPath ?? turn.currentTurnEvidencePath;
        string? previousTurnImageForRequest = includeTurnTemporalImages ? run.previousTurnFocusDataUrl : null;
        string? referenceTurnImageForRequest = includeTurnTemporalImages ? run.turnReferenceFocusDataUrl : null;
        bool previousTurnImageAttached = previousTurnImageForRequest is not null &&
            !string.Equals(previousTurnImageForRequest, focusDataUrlForRequest, StringComparison.Ordinal);
        bool referenceTurnImageAttached = referenceTurnImageForRequest is not null &&
            !string.Equals(referenceTurnImageForRequest, focusDataUrlForRequest, StringComparison.Ordinal) &&
            !string.Equals(referenceTurnImageForRequest, previousTurnImageForRequest, StringComparison.Ordinal);
        metaSb.AppendLine($"TURN_TEMPORAL_EVIDENCE: current={(focusDataUrlForRequest is null ? "absent" : "attached")}; previous={(previousTurnImageAttached ? "attached" : "absent")}; reference={(referenceTurnImageAttached ? "attached" : "absent")}; change_pairs={requestTurnChangeImages.Count}.");
        object BuildBody(bool standalone) => BuildRequestBody(
            Options.Model,
            run.systemRules,
            run.goal,
            (standalone ? TailHistory(run.historyBuffer, Options.HistoryTailChars, Options.HistoryTailLines) : turn.historyTail) + "\n" +
                (standalone ? metaSb.ToString().Replace("SCREEN_IMAGE: omitted because screen fingerprint is unchanged; use previous_response_id state plus current metadata.", "SCREEN_IMAGE: attached to restore independent context.", StringComparison.Ordinal) : metaSb.ToString()),
            turn.dataUrl,
            turn.imageW,
            turn.imageH,
            turn.cx,
            turn.cy,
            turn.cnx,
            turn.cny,
            focusDataUrlForRequest,
            focusRectForRequest,
            turn.focusUiaRectForPrompt,
            turn.focusUiaDataUrl,
            turn.promptContext,
            !standalone && reuseUiaTargets,
            standalone ? null : turn.previousResponseIdForRequest,
            !standalone && omitFullScreenImage,
            run.goalMode,
            previousTurnImageForRequest,
            referenceTurnImageForRequest,
            requestTurnChangeImages,
            controlMaxOutputTokens,
            run.controlContextChain.CompactionEnabled,
            run.controlContextChain.Enabled);
        var reqBody = BuildBody(standalone: false);
        var fallbackBody = turn.previousResponseIdForRequest is null ? null : BuildBody(standalone: true);
        SaveRequestLog(Path.Combine(run.requestsDir, $"{run.commandId}_{turn.step}_request.json"), reqBody);
        if (turn.hasPersistentTurnFocus)
        {
            run.previousTurnFocusDataUrl = turn.currentTurnEvidenceDataUrl;
            run.previousTurnFocusPath = turn.currentTurnEvidencePath;
        }

        run.controlContextChain.LogRequest(turn.step);
        var (candidates, raw, completedResponseId, contextFallbackUsed, compactionFallbackUsed, compactionOccurred) = await CallOpenAIAsync(
            run.apiKey, reqBody, run.Cancellation, fallbackBody);
        run.controlContextChain.RecordResult(completedResponseId, contextFallbackUsed, compactionFallbackUsed, compactionOccurred);
        run.promptVisualState.RecordResponse(turn, completedResponseId, !omitFullScreenImage || contextFallbackUsed);
        CurrentRun.RequestReasoningEffortOverride = null;
        SaveRaw(Path.Combine(run.requestsDir, $"{run.commandId}_{turn.step}_response.json"), raw);
        if (candidates is null)
        {
            if (CancelRequested)
            {
                Console.WriteLine("Aborted (hotkey).");
                run.runResult = new ControlRunResult(ControlRunOutcome.Cancelled, turn.step, "cancelled with the emergency hotkey");
                return StepFlow.Stop;
            }

            run.consecutiveModelFailures++;
            if (CurrentRun.LastOpenAiFailureWasRetriable &&
                (Options.MaxModelFailuresBeforeAbort == 0 ||
                 run.consecutiveModelFailures < Options.MaxModelFailuresBeforeAbort))
            {
                Console.WriteLine($"[openai] transient {CurrentRun.LastOpenAiFailureKind}; keeping goal alive ({run.consecutiveModelFailures}/{(Options.MaxModelFailuresBeforeAbort == 0 ? "unlimited" : Options.MaxModelFailuresBeforeAbort.ToString())}).");
                AddHistory(run.historyBuffer, $"[{turn.step}] model_failure_retry: {CurrentRun.LastOpenAiFailureKind}");
                run.prevAction = null;
                run.prevShotFingerprint = null;
                run.prevActiveWindowFingerprint = null;
                run.prevObservationFrame = null;
                run.prevObservationContext = null;
                await RuntimeClock.Delay(750, run.Cancellation);
                return StepFlow.Retry;
            }

            Console.WriteLine("Could not parse action. Aborting this goal.");
            run.runResult = new ControlRunResult(ControlRunOutcome.Failed, turn.step, "the model response did not contain a valid action");
            return StepFlow.Stop;
        }

        turn.plannedInputsForExecution = run.shortTermPlan.PrepareInputsForProposal(candidates[0],
            run.observationSession.EffectiveProfile == "turn_based_interaction" ? run.turnBasedTransitions.CurrentStateId : null);
        turn.Decision = BatchPlanner.CreateDecision(candidates,
            AllowsStableCanvasDrawBatch(turn.promptContext, run.observationSession.EffectiveProfile),
            run.turnBasedInteractionRect is not null && run.turnBasedTransitions.CanProposeExecutionBatch
                ? Math.Max(0, run.turnBasedTransitions.AdvertisedMaxExecutionBatchLength - 1) : 0,
            turn.plannedInputsForExecution);
        run.consecutiveModelFailures = 0;
        return StepFlow.Next;
    }
}
