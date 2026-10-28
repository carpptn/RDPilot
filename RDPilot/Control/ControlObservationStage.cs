namespace RDPilot.Control;

/// <summary>Captures the current screen and updates observation and recovery state.</summary>
internal static class ControlObservationStage
{
    internal static async Task<StepFlow> RunAsync(ControlRunState run, ControlStepState turn)
    {
        int progressH;
        int progressW;
        string progressImage;
        int screenCy;
        int screenCx;
        run.currentStep = turn.step;
        if (CancelRequested)
        {
            Console.WriteLine("Aborted (hotkey).");
            run.runResult = new ControlRunResult(ControlRunOutcome.Cancelled, turn.step, "cancelled with the emergency hotkey");
            return StepFlow.Stop;
        }

        turn.expectedContinuousIdle = false;
        if (run.turnBasedInteractionWindow != IntPtr.Zero &&
            string.Equals(run.observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal) &&
            GetForegroundWindow() != run.turnBasedInteractionWindow &&
            TryRestoreForegroundWindow(run.turnBasedInteractionWindow))
        {
            Console.WriteLine("[turn-focus] restored the remembered interaction window before observation.");
        }

        // Discard stale crops before capture, even if the window title did not change.
        if (run.turnBasedInteractionContext is { GeometryIsCurrent: false })
            run.ResetTurnInteraction();

        // screenshot at the beginning of a step (state after previous action)
        Rectangle? persistentTurnBasedRect = string.Equals(run.observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal)

            ? !run.turnBasedInteractionRegionIsAutomatic || run.turnBasedAutomaticRegionRefined

            ? run.turnBasedInteractionRect

            : null

            : null;
        Rectangle? requestedFocusRect = run.nextFocusRect ?? persistentTurnBasedRect;
        bool includeObservationDetail = run.prevAction?.ObservationRegion is not null ||
            run.prevAction?.Action.Type is "drag_drop" or "drag_path" or "type_text" or "paste_text" ||
            persistentTurnBasedRect is not null;
        var capture = ScreenshotToDataUrl(run.screensDir, run.commandId, turn.step, requestedFocusRect, includeObservationDetail);
        turn.dataUrl = capture.DataUrl;
        turn.savedPath = capture.SavedPath;
        turn.screenW = capture.ScreenWidth;
        turn.screenH = capture.ScreenHeight;
        turn.imageW = capture.ImageWidth;
        turn.imageH = capture.ImageHeight;
        turn.focusUrl = capture.FocusDataUrl;
        turn.focusUiaRect = capture.FocusUiaRect;
        turn.focusUiaDataUrl = capture.FocusUiaDataUrl;
        turn.shotFingerprint = capture.DeltaFingerprint;
        turn.activeWindowFingerprint = capture.ActiveWindowFingerprint;
        turn.observationFrame = capture.ObservationFrame;
        SetCurrentScreenMap(turn.screenW, turn.screenH, turn.imageW, turn.imageH);
        Console.WriteLine($"[shot] {ShotLabel(turn.savedPath, run.commandId, turn.step)}");
        if (CurrentRun.CurrentScreenMap.RequiresMapping)
            Console.WriteLine($"[coords] model image {turn.imageW}x{turn.imageH} -> desktop ({CurrentRun.CurrentScreenMap.ScreenX},{CurrentRun.CurrentScreenMap.ScreenY}) {turn.screenW}x{turn.screenH}");
        if (capture.FocusRect is Rectangle rCrop)
        {
            var cropPath = ScreenLogPath(run.screensDir, $"{run.commandId}_{turn.step}_crop");
            Console.WriteLine($"[crop] {Path.GetFileName(cropPath)}  bbox=({rCrop.Left},{rCrop.Top})–({rCrop.Right},{rCrop.Bottom})");
            if (Options.DebugImages && Options.LogScreens && turn.savedPath != null)
            {
                SaveAimOverlay(
                    turn.savedPath,
                    rCrop,
                    Path.Combine(run.screensDir, $"{run.commandId}_{turn.step}_aim_overlay.png"));
                Console.WriteLine($"[aim-overlay] {Path.GetFileName(Path.Combine(run.screensDir, $"{run.commandId}_{turn.step}_aim_overlay.png"))}");
            }
        }

        if (turn.focusUiaRect is Rectangle fr)
            Console.WriteLine($"[focus_uia] bbox=({fr.Left},{fr.Top})–({fr.Right},{fr.Bottom})");
        turn.promptContext = CaptureUiPromptContext(capture.FocusUiaSummary, turn.screenW, turn.screenH);
        run.observationSession.PrepareForPrompt(turn.promptContext, run.goal);
        turn.currentTurnBasedContext = TurnInteractionContext.Capture(turn.promptContext);
        if (run.turnBasedInteractionRect is not null &&
            run.turnBasedInteractionContext != turn.currentTurnBasedContext)
        {
            run.ResetTurnInteraction();
            turn.focusUrl = null;
            capture = capture with { FocusRect = null, FocusDataUrl = null };
        }

        if (run.turnBasedInteractionRect is null &&
            string.Equals(run.observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal) &&
            ResolveTurnBasedObservationRegion(null, "turn_based_interaction") is Rectangle automaticTurnRegion)
        {
            run.turnBasedInteractionRect = automaticTurnRegion;
            run.turnBasedInteractionRegionIsAutomatic = true;
            run.turnBasedAutomaticRegionRefined = false;
            run.turnBasedInteractionWindow = GetForegroundWindow();
            run.turnBasedInteractionContext = turn.currentTurnBasedContext;
            var turnStateFrame = turn.observationFrame.DetailWidth > 0

                ? turn.observationFrame

                : CaptureObservationFrameProbe(includeDetail: true);
            run.turnBasedTransitions.Reset();
            run.turnBasedTransitions.ObserveState(turnStateFrame, automaticTurnRegion);
            Console.WriteLine($"[turn] automatic interaction region established at ({automaticTurnRegion.Left},{automaticTurnRegion.Top})–({automaticTurnRegion.Right},{automaticTurnRegion.Bottom}); aggressive batching is available from the first planned route.");
        }

        // — Visual delta vs previous screenshot (effect of last action)
        if (run.prevShotFingerprint != null && run.prevObservationFrame != null)
        {
            run.lastObservation = run.observationSession.Assess(
                run.prevObservationFrame,
                turn.observationFrame,
                run.prevAction,
                run.prevObservationContext,
                turn.promptContext,
                run.goalMode,
                run.goal);
            if (run.turnBasedInteractionRect is Rectangle turnRegion &&
                string.Equals(run.lastObservation.ActionPolicy, "turn_based_interaction", StringComparison.Ordinal))
            {
                if (run.turnBasedInteractionRegionIsAutomatic &&
                    !run.turnBasedAutomaticRegionRefined &&
                    run.prevAction is not null &&
                    IsStateChangingInteractionAction(run.prevAction.Action) &&
                    run.lastObservation.VisualChange == VisualChangeState.Changed &&
                    InferTurnBasedObservationRegion(run.prevObservationFrame, turn.observationFrame, turnRegion) is Rectangle refinedTurnRegion)
                {
                    run.turnBasedInteractionRect = refinedTurnRegion;
                    run.turnBasedAutomaticRegionRefined = true;
                    turnRegion = refinedTurnRegion;
                    run.previousTurnFocusDataUrl = null;
                    run.previousTurnFocusPath = null;
                    run.turnReferenceFocusDataUrl = null;
                    run.turnReferenceFocusPath = null;
                    run.activeTurnChangeImages = [];
                    run.turnBasedTransitions.Reset();
                    run.turnBasedTransitions.ObserveState(run.prevObservationFrame, refinedTurnRegion);
                    Console.WriteLine($"[turn] automatic observation region refined to ({refinedTurnRegion.Left},{refinedTurnRegion.Top})–({refinedTurnRegion.Right},{refinedTurnRegion.Bottom}); the model still receives the full screen.");
                }

                if (run.prevAction is not null &&
                    !run.prevAction.Action.TurnSequenceObserved &&
                    IsStateChangingInteractionAction(run.prevAction.Action))
                {
                    if (run.turnBasedTransitions.RecordTransition(run.prevObservationFrame, turn.observationFrame, turnRegion, run.prevAction, run.lastObservation))
                    {
                        run.lastObservation = run.lastObservation with
                        {
                            ActionOutcome = ActionOutcomeState.NoEffect,
                            GoalProgress = GoalProgressState.NoProgress
                        };
                    }
                }
                else
                {
                    var passiveBaseline = run.turnBasedTransitions.PrepareActionBaseline(turn.observationFrame, turnRegion);
                    if (passiveBaseline.ExternalStateChange)
                    {
                        Console.WriteLine($"[turn] passive observation crossed an external state boundary; {passiveBaseline.Summary}");
                        run.BeginExternalStateEpoch(run.lastObservation);
                    }
                }
            }

            if (run.prevObservedAction is { } observed && ReferenceEquals(observed.Action, run.prevAction))
                run.lastObservation = observed.Reconcile(run.lastObservation, turn.observationFrame);

            if (run.lastObservation.ActionOutcome == ActionOutcomeState.NoEffect)
                run.shortTermPlan.Invalidate("the latest planned action had no visible effect");
            else if (run.lastObservation.ActionOutcome == ActionOutcomeState.UnexpectedChange)
                run.shortTermPlan.Invalidate("the latest action produced an unexpected visual change");
            if (run.prevAction?.Action.Type is "drag_path" or "drag_drop" or "hold_keys" ||
                string.Equals(run.lastObservation.ActionPolicy, "turn_based_interaction", StringComparison.Ordinal))
            {
                Console.WriteLine($"[observation] result profile={run.lastObservation.Profile}; "+ $"policy={run.lastObservation.ActionPolicy}; visual={run.lastObservation.VisualChange}; "+ $"outcome={run.lastObservation.ActionOutcome}; progress={run.lastObservation.GoalProgress}; "+ $"local_delta={(double.IsFinite(run.lastObservation.LocalDelta) ? run.lastObservation.LocalDelta.ToString("0.####") : "n/a")}; "+ $"local_ratio={run.lastObservation.LocalChangedRatio:0.####}; "+ $"threshold={run.lastObservation.ChangeThreshold:0.####}; confidence={run.lastObservation.Confidence:0.00}");
            }

            run.lastGlobalDelta = run.lastObservation.GlobalDelta;
            run.lastActiveWindowDelta = run.lastObservation.ActiveWindowDelta;
            run.lastDelta = run.lastObservation.GoalProgress switch
            {
                GoalProgressState.Progress => Math.Max(run.lastObservation.EffectiveDelta, Options.NoChangeThreshold * 1.2),
                GoalProgressState.NoProgress => Math.Min(run.lastObservation.EffectiveDelta, Options.NoChangeThreshold * 0.5),
                _ => run.lastObservation.EffectiveDelta
            };
            bool noChange = run.lastObservation.GoalProgress == GoalProgressState.NoProgress;
            bool confirmedProgress = run.lastObservation.GoalProgress == GoalProgressState.Progress;
            bool productiveTurnObservation = string.Equals(run.lastObservation.ActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                run.lastObservation.ActionOutcome == ActionOutcomeState.Confirmed;
            bool confirmedStateChange = run.lastObservation.ActionOutcome == ActionOutcomeState.Confirmed &&
                run.lastObservation.VisualChange == VisualChangeState.Changed;
            bool ambiguousProgress = run.lastObservation.GoalProgress == GoalProgressState.Unknown;
            var observedNoChange = noChange;
            var textEditGrace = ShouldDeferFocusedTextStagnation(run.prevAction, noChange, run.textInputNoChangeAttempts);
            if (textEditGrace)
            {
                Console.WriteLine("[observation] focused text edit produced no confirmed pixels; deferring stagnation for one attempt.");
                run.lastObservation = run.lastObservation with
                {
                    ActionOutcome = ActionOutcomeState.Ambiguous,
                    GoalProgress = GoalProgressState.Unknown,
                    Evidence = run.lastObservation.Evidence + "; focused_text_grace=true"
                };
                noChange = false;
                ambiguousProgress = true;
            }

            bool previousWasObservationOnly = run.prevAction != null && IsLocalObservationAction(run.prevAction.Action);
            turn.expectedContinuousIdle = IsExpectedContinuousIdle(run.goalMode, run.prevAction?.Action, noChange);
            if (ShouldResetRejectedProposalLoop(run.prevAction, !(confirmedProgress || productiveTurnObservation), turn.expectedContinuousIdle))
            {
                run.ClearRejectedProposalLoop();
            }

            if (turn.expectedContinuousIdle)
            {
                run.continuousIdleSteps++;
                run.ambiguousObservationSteps = 0;
                run.stagnationSteps = 0;
                run.repeatCount = 0;
                run.lastSig = null;
                run.recentIneffectiveSpatialActions.Clear();
            }
            else if (!previousWasObservationOnly)
            {
                run.continuousIdleSteps = 0;
                if (noChange)
                {
                    run.ambiguousObservationSteps = 0;
                    run.stagnationSteps++;
                }
                else if (confirmedProgress || productiveTurnObservation)
                {
                    run.ambiguousObservationSteps = 0;
                    run.stagnationSteps = 0;
                }
                else if (ambiguousProgress)
                {
                    run.ambiguousObservationSteps++;
                    if (run.ambiguousObservationSteps >= 3)
                        run.stagnationSteps++;
                }
                else
                {
                    run.ambiguousObservationSteps = 0;
                }
            }
            else
            {
                run.continuousIdleSteps = 0;
                if (confirmedProgress || productiveTurnObservation)
                    run.stagnationSteps = 0;
            }

            if (confirmedStateChange)
            {
                run.actionCooldownUntilStep.Clear();
                run.spatialActionCooldowns.Clear();
            }

            if (run.prevAction != null && !turn.expectedContinuousIdle)
            {
                (run.repeatCount, run.lastSig) = UpdateRepeatDetection(run.prevAction, noChange, run.repeatCount, run.lastSig, run.recentIneffectiveSpatialActions);
                if (ShouldRegisterImmediateNoEffectCooldown(run.prevAction.Action, run.lastObservation, run.repeatCount))
                {
                    RegisterActionCooldown(
                        run.prevAction,
                        turn.step + Options.ActionRepeatCooldownSteps,
                        run.actionCooldownUntilStep,
                        run.spatialActionCooldowns,
                        clusterSpatially: run.repeatCount > 0);
                }
            }

            // Expire AIM after a large visual change
            var broadDelta = Math.Max(run.lastGlobalDelta, run.lastActiveWindowDelta);
            if (run.lastAimRect is not null && broadDelta > Options.AimExpireDelta)
            {
                Console.WriteLine($"[aim] expired (delta={broadDelta:0.###} > {Options.AimExpireDelta:0.###})");
                run.InvalidateAim();
            }

            if (noChange && run.prevAction != null && IsPointClickAction(run.prevAction.Action))
            {
                run.lastPrecisionHint = BuildPrecisionHint(run.prevAction, run.lastDelta, "Previous click produced little or no visible progress.");
                run.lastPrecisionHintExpiresAfterStep = turn.step + 1;
            }

            if (noChange && run.prevAction?.Action.Type is "drag_drop" or "drag_path")
            {
                run.lastPrecisionHint = run.prevAction.Action.Type == "drag_path"

                    ? "Previous drag_path did not produce the expected local effect. Do not replay the same path; verify the active tool/surface or change the gesture semantics."

                    : "Previous drag_drop did not visibly move the source or change the destination. Do not retry nearby coordinates; verify that the source is draggable, identify the semantic destination, or use a different interaction route.";
                run.lastPrecisionHintExpiresAfterStep = turn.step + 2;
            }

            if (observedNoChange && run.prevAction != null &&
                IsTextInputAttemptAction(run.prevAction.Action))
            {
                run.textInputNoChangeAttempts++;
                run.lastTextInputHint = BuildTextInputHint(run.prevAction.Action, run.lastDelta);
                run.lastTextInputHintExpiresAfterStep = turn.step + 4;
                if (run.textInputNoChangeAttempts >= 2)
                    run.textInputCooldownUntilStep = Math.Max(run.textInputCooldownUntilStep, turn.step + 4);
            }
            else if (confirmedProgress || productiveTurnObservation)
            {
                run.lastPrecisionHint = null;
                run.lastPrecisionHintExpiresAfterStep = 0;
                run.lastTextInputHint = null;
                run.lastTextInputHintExpiresAfterStep = 0;
                run.textInputNoChangeAttempts = 0;
                run.textInputCooldownUntilStep = 0;
            }
        }

        if (run.prevShotFingerprint is null || run.prevObservationFrame is null)
        {
            run.lastObservation = null;
            run.lastDelta = run.lastGlobalDelta = run.lastActiveWindowDelta = double.NaN;
        }

        turn.previousResponseIdForRequest = run.controlContextChain.PreviousResponseIdForRequest;
        turn.historyTail = turn.previousResponseIdForRequest != null || Options.HistoryTailChars <= 0

            ? ""

            : TailHistory(run.historyBuffer, Options.HistoryTailChars, Options.HistoryTailLines);
        (screenCx, screenCy, _, _) = GetCursorPositionInPrimary();
        (turn.cx, turn.cy, turn.cnx, turn.cny) = CursorToImageCoordinates(screenCx, screenCy);
        turn.appliedFocusRectForPrompt = ScreenRectToImage(capture.FocusRect);
        turn.focusUiaRectForPrompt = ScreenRectToImage(turn.focusUiaRect);
        turn.currentTurnFocusPath = capture.FocusRect is not null && turn.focusUrl != null && Options.LogScreens

            ? ScreenLogPath(run.screensDir, $"{run.commandId}_{turn.step}_crop")

            : null;
        turn.currentTurnEvidenceDataUrl = capture.FocusRect is Rectangle activeFocusRegion &&
            run.turnBasedInteractionRect is Rectangle activeTurnRegion &&
            activeFocusRegion == activeTurnRegion &&
            turn.focusUrl != null

            ? turn.focusUrl

            : null;
        turn.currentTurnEvidencePath = turn.currentTurnEvidenceDataUrl is not null ? turn.currentTurnFocusPath : null;
        if (turn.currentTurnEvidenceDataUrl is null &&
            run.turnBasedInteractionRegionIsAutomatic &&
            run.turnBasedAutomaticRegionRefined &&
            run.turnBasedInteractionRect is Rectangle automaticEvidenceRegion &&
            TryBuildTurnRegionEvidenceImage(
            turn.dataUrl,
            automaticEvidenceRegion,
            turn.observationFrame.ScreenBounds,
            run.screensDir,
            run.commandId,
            turn.step,
            out var automaticEvidenceDataUrl,
            out var automaticEvidencePath))
        {
            turn.currentTurnEvidenceDataUrl = automaticEvidenceDataUrl;
            turn.currentTurnEvidencePath = automaticEvidencePath;
        }

        turn.hasPersistentTurnFocus = turn.currentTurnEvidenceDataUrl is not null;
        turn.hasTurnVisualContext = string.Equals(run.observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal) &&
            run.turnBasedInteractionRect is not null &&
            turn.currentTurnEvidenceDataUrl is not null;
        if (turn.hasPersistentTurnFocus && run.turnReferenceFocusDataUrl is null)
        {
            run.turnReferenceFocusDataUrl = turn.currentTurnEvidenceDataUrl;
            run.turnReferenceFocusPath = turn.currentTurnEvidencePath;
            Console.WriteLine("[turn-memory] captured initial interaction reference image.");
        }

        turn.loopAssessment = AssessVisualStateCycle(
            run.loopStateGraph,
            turn.shotFingerprint,
            turn.activeWindowFingerprint,
            turn.promptContext,
            turn.step,
            run.recentExecutedActions,
            run.prevAction,
            run.lastDelta,
            goalMode: run.goalMode,
            recurringWorkflowIntent: run.recurringWorkflowIntent,
            goalProgress: run.lastObservation?.GoalProgress ?? GoalProgressState.Unknown);
        turn.proactiveVisualCycle = turn.loopAssessment.IsLoop;
        turn.visualCycleLength = turn.loopAssessment.CycleLength;
        if (turn.proactiveVisualCycle)
            Console.WriteLine($"[loop] proactive visual cycle detected; confidence={turn.loopAssessment.Confidence:0.00}; returned after {turn.visualCycleLength} step(s); {turn.loopAssessment.Evidence}");
        else if (turn.loopAssessment.IsProductiveCycle)
            Console.WriteLine($"[loop] productive recurring workflow observed; confidence={turn.loopAssessment.Confidence:0.00}; returned after {turn.visualCycleLength} step(s); no recovery intervention");
        else if (turn.loopAssessment.Confidence >= 0.5)
            Console.WriteLine($"[loop] possible visual recurrence not yet confirmed; confidence={turn.loopAssessment.Confidence:0.00}; threshold={turn.loopAssessment.DecisionThreshold:0.00}; {turn.loopAssessment.Evidence}");
        AppendLoopReplayObservation(
            run.commandId,
            turn.step,
            turn.screenW,
            turn.screenH,
            turn.shotFingerprint,
            turn.activeWindowFingerprint,
            turn.promptContext,
            run.prevAction,
            run.lastDelta,
            turn.loopAssessment,
            run.lastObservation,
            run.goalMode,
            run.recurringWorkflowIntent,
            run.stateEpoch,
            run.recentExecutedActions);
        run.recoveryEpisode = await UpdateRecoveryEpisodeAsync(
            run.recoveryEpisode,
            turn.step,
            run.stagnationSteps,
            run.repeatCount,
            run.lastDelta,
            run.lastObservation,
            turn.shotFingerprint,
            turn.activeWindowFingerprint,
            run.prevAction,
            turn.promptContext,
            run.recentExecutedActions,
            run.recoveryLessons,
            turn.loopAssessment,
            run.planningLoopDetected,
            run.rejectedProposalCycleLength,
            run.previousRejectedAction,
            run.recentRejectedActions,
            run.goal,
            run.goalMode,
            turn.dataUrl,
            turn.savedPath,
            async episode =>
        {
            var progress = PrepareHelperImage(turn.dataUrl, Options.VerifyScreenshotMaxWidth);
            (progressImage, progressW, progressH) = (progress.DataUrl, progress.Width, progress.Height);
            return await VerifyRecoveryProgressAsync(
                    run.apiKey,
                    run.goal,
                    run.goalMode,
                    episode,
                    progressImage,
                    turn.savedPath,
                    turn.promptContext,
                    progressW,
                    progressH,
                    run.requestsDir,
                    run.commandId,
                    turn.step,
                    run.Cancellation);
        });
        return StepFlow.Next;
    }
}
