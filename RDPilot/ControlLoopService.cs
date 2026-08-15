internal static partial class RDPilotApplication
{
    /// <summary>
    /// Runs the goal-driven desktop control loop and its safety guards.
    /// </summary>
    internal static class ControlLoopService
    {
            // === Control loop ===
            internal static async Task<ControlRunResult> RunOnce(
                string apiKey,
                string goal)
            {
                var commandId = Guid.NewGuid().ToString("N");
                var screensDir = EnsureScreensDir();
                var requestsDir = EnsureRequestsDir();
                var logDir = EnsureLogDir();

                var prevOut = Console.Out;
                var prevErr = Console.Error;
                var logPath = Path.Combine(logDir, $"{commandId}.log");
                using var logFile = new StreamWriter(logPath, append: false, Encoding.UTF8) { AutoFlush = true };
                var synchronizedLog = TextWriter.Synchronized(logFile);
                using var outputTee = new TeeTextWriter(prevOut, synchronizedLog);
                using var errorTee = new TeeTextWriter(prevErr, synchronizedLog);
                Console.SetOut(outputTee);
                Console.SetError(errorTee);
        
                CancellationTokenSource? cancelCts = null;
                ControlContextChain? controlContextChain = null;
                var consoleHidden = false;
                var currentStep = 0;
                var runResult = new ControlRunResult(
                    ControlRunOutcome.StepLimitReached,
                    MaxSteps,
                    MaxSteps > 0
                        ? $"configured step limit {MaxSteps} was reached"
                        : "continuous run ended unexpectedly");
                try
                {
                    ResetRunMetrics();
                    PendingSafeActions.Clear();
                    Console.WriteLine($"Command ID: {commandId}");
                    Console.WriteLine($"Goal: {goal}");
                    var goalMode = ResolveGoalMode(goal, GoalMode);
                    var recurringWorkflowIntent =
                        string.Equals(
                            goalMode,
                            "continuous",
                            StringComparison.Ordinal) &&
                        HasRecurringWorkflowIntent(goal);
                    Console.WriteLine($"Goal mode: {goalMode}");
                    if (recurringWorkflowIntent)
                        Console.WriteLine("Goal cycle policy: recurring workflow may contain productive state returns.");
                    Console.WriteLine("Loop start: one action -> screenshot -> next decision.");
                    Console.WriteLine("Emergency abort: Ctrl+Alt+Q\n");
        
                    if (!BatchMode &&
                        (AutoHideConsoleDuringRun || MinimizeConsoleDuringRun))
                        consoleHidden = ConcealConsoleWindow();
        
                    if (AllowHighLevelActions && TryExecuteFastGoal(goal))
                    {
                        Console.WriteLine("Finished (local fast path).");
                        return new ControlRunResult(
                            ControlRunOutcome.Completed,
                            0,
                            "completed through the local fast path");
                    }
        
                    controlContextChain = new ControlContextChain(
                        commandId,
                        UsePreviousResponseState,
                        ControlContextCompactionEnabled,
                        ControlContextFallbackLimit);
                    controlContextChain.LogStart(Model);
                    var systemRules = BuildSystemRules(); // stable per run; dynamic action availability is enforced by schema
                    var historyBuffer = new StringBuilder();
                    var recoveryLessons = LoadRecoveryLessons();
                    var recentExecutedActions = new Queue<ResolvedActionSnapshot>();
                    var loopStateGraph = new LoopStateGraph { RunId = commandId };
                    var observationSession = new AdaptiveObservationSession();
                    var observationActionGuard = new ObservationActionGuardState();
                    var shortTermPlan = new ShortTermPlanTracker();
                    var turnBasedTransitions = new TurnBasedTransitionTracker(shortTermPlan);
                    observationSession.LogInitialProfile();
                    RecoveryEpisodeState? recoveryEpisode = null;
                    cancelCts = StartCancelHotkeyListener();
        
                    Rectangle? nextFocusRect = null; // crop/overlay after 'aim'/'point'/'request_crop'
                    Rectangle? lastAimRect = null;   // active AIM – required before clicks
                    Rectangle? turnBasedInteractionRect = null;
                    var turnBasedInteractionRegionIsAutomatic = false;
                    var turnBasedAutomaticRegionRefined = false;
                    var turnBasedInteractionWindow = IntPtr.Zero;
                    string? turnBasedInteractionContext = null;
                    string? previousTurnFocusDataUrl = null;
                    string? previousTurnFocusPath = null;
                    string? turnReferenceFocusDataUrl = null;
                    string? turnReferenceFocusPath = null;
                    IReadOnlyList<TurnChangeImagePair> activeTurnChangeImages = [];
        
                    // change / strategy metrics
                    byte[]? prevShotFingerprint = null;
                    byte[]? prevActiveWindowFingerprint = null;
                    ScreenObservationFrame? prevObservationFrame = null;
                    UiPromptContext? prevObservationContext = null;
                    ObservationAssessment? lastObservation = null;
                    ResolvedActionSnapshot? prevAction = null;
                    string? lastSig = null;
                    var recentIneffectiveSpatialActions = new List<ResolvedActionSnapshot>();
                    int stagnationSteps = 0;
                    int repeatCount = 0;
                    int continuousIdleSteps = 0;
                    int ambiguousObservationSteps = 0;
                    double lastDelta = double.NaN;
                    double lastGlobalDelta = double.NaN;
                    double lastActiveWindowDelta = double.NaN;
                    string? lastVerifierRejection = null;
                    string? lastExecutorFailure = null;
                    string? lastPrecisionHint = null;
                    int lastPrecisionHintExpiresAfterStep = 0;
                    string? lastTextInputHint = null;
                    int lastTextInputHintExpiresAfterStep = 0;
                    int textInputNoChangeAttempts = 0;
                    int textInputCooldownUntilStep = 0;
                    int consecutiveModelFailures = 0;
                    int consecutiveActionFailures = 0;
                    var actionCooldownUntilStep = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    var spatialActionCooldowns = new List<SpatialActionCooldown>();
                    var recentRejectedActions = new Queue<ResolvedActionSnapshot>();
                    var recentRejectedProposalSignatures = new Queue<string>();
                    var rejectedProposalCycleCount = 0;
                    var rejectedProposalCycleLength = 0;
                    var planningLoopDetected = false;
                    ResolvedActionSnapshot? previousRejectedAction = null;

                    void RegisterRejectedProposal(
                        ResolvedActionSnapshot rejected)
                    {
                        previousRejectedAction = rejected;
                        recentRejectedActions.Enqueue(rejected);
                        while (recentRejectedActions.Count > 24)
                            recentRejectedActions.Dequeue();

                        recentRejectedProposalSignatures.Enqueue(
                            rejected.IneffectiveSignature);
                        while (recentRejectedProposalSignatures.Count > 32)
                            recentRejectedProposalSignatures.Dequeue();

                        var cycleLength = RepeatedStringCycleLength(
                            recentRejectedProposalSignatures.ToArray());
                        if (cycleLength <= 0)
                        {
                            rejectedProposalCycleCount = 0;
                            rejectedProposalCycleLength = 0;
                            planningLoopDetected = false;
                            return;
                        }

                        rejectedProposalCycleCount =
                            rejectedProposalCycleLength == cycleLength
                                ? rejectedProposalCycleCount + 1
                                : 1;
                        rejectedProposalCycleLength = cycleLength;
                        planningLoopDetected = true;
                        Console.WriteLine(
                            $"[loop] rejected-proposal cycle detected; length={cycleLength}; recurrence={rejectedProposalCycleCount}; action={rejected.Description}");
                    }

                    void ClearRejectedProposalLoop()
                    {
                        recentRejectedActions.Clear();
                        recentRejectedProposalSignatures.Clear();
                        rejectedProposalCycleCount = 0;
                        rejectedProposalCycleLength = 0;
                        planningLoopDetected = false;
                        previousRejectedAction = null;
                    }

                    void BeginExternalStateEpoch(
                        ObservationAssessment? boundaryObservation = null)
                    {
                        loopStateGraph = new LoopStateGraph { RunId = commandId };
                        recoveryEpisode = null;
                        recentExecutedActions.Clear();
                        recentIneffectiveSpatialActions.Clear();
                        ClearRejectedProposalLoop();
                        stagnationSteps = 0;
                        repeatCount = 0;
                        continuousIdleSteps = 0;
                        ambiguousObservationSteps = 0;
                        consecutiveActionFailures = 0;
                        lastSig = null;
                        lastObservation = boundaryObservation;
                        lastDelta = double.NaN;
                        lastGlobalDelta = double.NaN;
                        lastActiveWindowDelta = double.NaN;
                        actionCooldownUntilStep.Clear();
                        spatialActionCooldowns.Clear();
                        textInputNoChangeAttempts = 0;
                        textInputCooldownUntilStep = 0;
                        previousTurnFocusDataUrl = null;
                        previousTurnFocusPath = null;
                        turnReferenceFocusDataUrl = null;
                        turnReferenceFocusPath = null;
                        activeTurnChangeImages = [];
                        turnBasedTransitions.BeginExternalStateEpoch();
                        Console.WriteLine(
                            "[loop] external state boundary started a fresh visual and loop-detection epoch; local topology and temporal images were cleared while prior task-mechanics evidence was retained.");
                    }
        
                    for (int step = 1;
                         MaxSteps == 0 || step <= MaxSteps;
                         step++)
                    {
                        currentStep = step;
                        if (CancelRequested)
                        {
                            Console.WriteLine("Aborted (hotkey).");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.Cancelled,
                                step,
                                "cancelled with the emergency hotkey");
                            break;
                        }
                        var expectedContinuousIdle = false;

                        if (turnBasedInteractionWindow != IntPtr.Zero &&
                            string.Equals(
                                observationSession.EffectiveProfile,
                                "turn_based_interaction",
                                StringComparison.Ordinal) &&
                            GetForegroundWindow() != turnBasedInteractionWindow &&
                            TryRestoreForegroundWindow(turnBasedInteractionWindow))
                        {
                            Console.WriteLine("[turn-focus] restored the remembered interaction window before observation.");
                        }
        
                        // screenshot at the beginning of a step (state after previous action)
                        var persistentTurnBasedRect = string.Equals(
                                observationSession.EffectiveProfile,
                                "turn_based_interaction",
                                StringComparison.Ordinal)
                            ? !turnBasedInteractionRegionIsAutomatic ||
                              turnBasedAutomaticRegionRefined
                                ? turnBasedInteractionRect
                                : null
                            : null;
                        var requestedFocusRect = nextFocusRect ?? persistentTurnBasedRect;
                        var includeObservationDetail =
                            prevAction?.ObservationRegion is not null ||
                            prevAction?.Action.Type is "drag_drop" or "drag_path" or "type_text" or "paste_text" ||
                            persistentTurnBasedRect is not null;
                        var (dataUrl, savedPath, screenW, screenH, imageW, imageH, focusUrl, appliedFocusRect, focusUiaRect, focusUiaSummary, focusUiaDataUrl, focusUiaPath, shotFingerprint, activeWindowFingerprint, observationFrame) =
                            ScreenshotToDataUrl(
                                screensDir,
                                commandId,
                                step,
                                requestedFocusRect,
                                includeObservationDetail);
                        SetCurrentScreenMap(screenW, screenH, imageW, imageH);
                        Console.WriteLine($"[shot] {ShotLabel(savedPath, commandId, step)}");
                        if (CurrentScreenMap.RequiresMapping)
                            Console.WriteLine($"[coords] model image {imageW}x{imageH} -> desktop ({CurrentScreenMap.ScreenX},{CurrentScreenMap.ScreenY}) {screenW}x{screenH}");
        
                        if (appliedFocusRect is Rectangle rCrop)
                        {
                            var cropPath = ScreenLogPath(screensDir, $"{commandId}_{step}_crop");
                            Console.WriteLine($"[crop] {Path.GetFileName(cropPath)}  bbox=({rCrop.Left},{rCrop.Top})–({rCrop.Right},{rCrop.Bottom})");
                            if (DebugImages && LogScreens && savedPath != null)
                            {
                                SaveAimOverlay(savedPath, rCrop, Path.Combine(screensDir, $"{commandId}_{step}_aim_overlay.png"));
                                Console.WriteLine($"[aim-overlay] {Path.GetFileName(Path.Combine(screensDir, $"{commandId}_{step}_aim_overlay.png"))}");
                            }
                        }
                        if (focusUiaRect is Rectangle fr)
                            Console.WriteLine($"[focus_uia] bbox=({fr.Left},{fr.Top})–({fr.Right},{fr.Bottom})");

                        var promptContext = CaptureUiPromptContext(
                            focusUiaSummary,
                            screenW,
                            screenH);
                        observationSession.PrepareForPrompt(promptContext, goal);
                        var currentTurnBasedContext = TurnBasedContextKey(promptContext);
                        if (turnBasedInteractionRect is not null &&
                            !string.Equals(
                                turnBasedInteractionContext,
                                currentTurnBasedContext,
                                StringComparison.Ordinal))
                        {
                            Console.WriteLine("[turn] cleared persistent interaction region because the foreground context changed.");
                            turnBasedInteractionRect = null;
                            turnBasedInteractionRegionIsAutomatic = false;
                            turnBasedAutomaticRegionRefined = false;
                            turnBasedInteractionWindow = IntPtr.Zero;
                            turnBasedInteractionContext = null;
                            previousTurnFocusDataUrl = null;
                            previousTurnFocusPath = null;
                            turnReferenceFocusDataUrl = null;
                            turnReferenceFocusPath = null;
                            activeTurnChangeImages = [];
                            turnBasedTransitions.Reset();
                        }

                        if (turnBasedInteractionRect is null &&
                            string.Equals(
                                observationSession.EffectiveProfile,
                                "turn_based_interaction",
                                StringComparison.Ordinal) &&
                            ResolveTurnBasedObservationRegion(
                                null,
                                "turn_based_interaction") is Rectangle automaticTurnRegion)
                        {
                            turnBasedInteractionRect = automaticTurnRegion;
                            turnBasedInteractionRegionIsAutomatic = true;
                            turnBasedAutomaticRegionRefined = false;
                            turnBasedInteractionWindow = GetForegroundWindow();
                            turnBasedInteractionContext = currentTurnBasedContext;
                            var turnStateFrame = observationFrame.DetailWidth > 0
                                ? observationFrame
                                : CaptureObservationFrameProbe(includeDetail: true);
                            turnBasedTransitions.Reset();
                            turnBasedTransitions.ObserveState(
                                turnStateFrame,
                                automaticTurnRegion);
                            Console.WriteLine(
                                $"[turn] automatic interaction region established at ({automaticTurnRegion.Left},{automaticTurnRegion.Top})–({automaticTurnRegion.Right},{automaticTurnRegion.Bottom}); aggressive batching is available from the first planned route.");
                        }
        
                        // — Visual delta vs previous screenshot (effect of last action)
                        if (prevShotFingerprint != null && prevObservationFrame != null)
                        {
                            lastObservation = observationSession.Assess(
                                prevObservationFrame,
                                observationFrame,
                                prevAction,
                                prevObservationContext,
                                promptContext,
                                goalMode,
                                goal);
                            if (turnBasedInteractionRect is Rectangle turnRegion &&
                                string.Equals(
                                    lastObservation.ActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal))
                            {
                                if (turnBasedInteractionRegionIsAutomatic &&
                                    !turnBasedAutomaticRegionRefined &&
                                    prevAction is not null &&
                                    IsStateChangingInteractionAction(prevAction.Action) &&
                                    lastObservation.VisualChange == VisualChangeState.Changed &&
                                    InferTurnBasedObservationRegion(
                                        prevObservationFrame,
                                        observationFrame,
                                        turnRegion) is Rectangle refinedTurnRegion)
                                {
                                    turnBasedInteractionRect = refinedTurnRegion;
                                    turnBasedAutomaticRegionRefined = true;
                                    turnRegion = refinedTurnRegion;
                                    previousTurnFocusDataUrl = null;
                                    previousTurnFocusPath = null;
                                    turnReferenceFocusDataUrl = null;
                                    turnReferenceFocusPath = null;
                                    activeTurnChangeImages = [];
                                    turnBasedTransitions.Reset();
                                    turnBasedTransitions.ObserveState(
                                        prevObservationFrame,
                                        refinedTurnRegion);
                                    Console.WriteLine(
                                        $"[turn] automatic observation region refined to ({refinedTurnRegion.Left},{refinedTurnRegion.Top})–({refinedTurnRegion.Right},{refinedTurnRegion.Bottom}); the model still receives the full screen.");
                                }
                                if (prevAction is not null &&
                                    !prevAction.Action.TurnSequenceObserved &&
                                    IsStateChangingInteractionAction(prevAction.Action))
                                {
                                    if (turnBasedTransitions.RecordTransition(
                                            prevObservationFrame,
                                            observationFrame,
                                            turnRegion,
                                            prevAction,
                                            lastObservation))
                                    {
                                        lastObservation = lastObservation with
                                        {
                                            ActionOutcome = ActionOutcomeState.NoEffect,
                                            GoalProgress = GoalProgressState.NoProgress
                                        };
                                    }
                                }
                                else
                                {
                                    var passiveBaseline =
                                        turnBasedTransitions.PrepareActionBaseline(
                                        observationFrame,
                                        turnRegion);
                                    if (passiveBaseline.ExternalStateChange)
                                    {
                                        Console.WriteLine(
                                            $"[turn] passive observation crossed an external state boundary; {passiveBaseline.Summary}");
                                        BeginExternalStateEpoch(lastObservation);
                                    }
                                }
                            }
                            if (lastObservation.ActionOutcome == ActionOutcomeState.NoEffect)
                                shortTermPlan.Invalidate("the latest planned action had no visible effect");
                            else if (lastObservation.ActionOutcome == ActionOutcomeState.UnexpectedChange)
                                shortTermPlan.Invalidate("the latest action produced an unexpected visual change");
                            if (prevAction?.Action.Type is "drag_path" or "drag_drop" or "hold_keys" ||
                                string.Equals(
                                    lastObservation.ActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal))
                            {
                                Console.WriteLine(
                                    $"[observation] result profile={lastObservation.Profile}; " +
                                    $"policy={lastObservation.ActionPolicy}; visual={lastObservation.VisualChange}; " +
                                    $"outcome={lastObservation.ActionOutcome}; progress={lastObservation.GoalProgress}; " +
                                    $"local_delta={(double.IsFinite(lastObservation.LocalDelta) ? lastObservation.LocalDelta.ToString("0.####") : "n/a")}; " +
                                    $"local_ratio={lastObservation.LocalChangedRatio:0.####}; " +
                                    $"threshold={lastObservation.ChangeThreshold:0.####}; confidence={lastObservation.Confidence:0.00}");
                            }
                            lastGlobalDelta = lastObservation.GlobalDelta;
                            lastActiveWindowDelta = lastObservation.ActiveWindowDelta;
                            lastDelta = lastObservation.GoalProgress switch
                            {
                                GoalProgressState.Progress => Math.Max(
                                    lastObservation.EffectiveDelta,
                                    NoChangeThreshold * 1.2),
                                GoalProgressState.NoProgress => Math.Min(
                                    lastObservation.EffectiveDelta,
                                    NoChangeThreshold * 0.5),
                                _ => lastObservation.EffectiveDelta
                            };
                            bool noChange =
                                lastObservation.GoalProgress == GoalProgressState.NoProgress;
                            bool confirmedProgress =
                                lastObservation.GoalProgress == GoalProgressState.Progress;
                            bool productiveTurnObservation =
                                string.Equals(
                                    lastObservation.ActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal) &&
                                lastObservation.ActionOutcome == ActionOutcomeState.Confirmed;
                            bool confirmedStateChange =
                                lastObservation.ActionOutcome == ActionOutcomeState.Confirmed &&
                                lastObservation.VisualChange == VisualChangeState.Changed;
                            bool ambiguousProgress =
                                lastObservation.GoalProgress == GoalProgressState.Unknown;
                            var observedNoChange = noChange;
                            var textEditGrace = ShouldDeferFocusedTextStagnation(
                                prevAction,
                                noChange,
                                textInputNoChangeAttempts);
                            if (textEditGrace)
                            {
                                Console.WriteLine("[observation] focused text edit produced no confirmed pixels; deferring stagnation for one attempt.");
                                lastObservation = lastObservation with
                                {
                                    ActionOutcome = ActionOutcomeState.Ambiguous,
                                    GoalProgress = GoalProgressState.Unknown,
                                    Evidence = lastObservation.Evidence + "; focused_text_grace=true"
                                };
                                noChange = false;
                                ambiguousProgress = true;
                            }
                            bool previousWasObservationOnly = prevAction != null && IsLocalObservationAction(prevAction.Action);
                            expectedContinuousIdle = IsExpectedContinuousIdle(
                                goalMode,
                                prevAction?.Action,
                                noChange);

                            if (ShouldResetRejectedProposalLoop(
                                    prevAction,
                                    !(confirmedProgress || productiveTurnObservation),
                                    expectedContinuousIdle))
                            {
                                ClearRejectedProposalLoop();
                            }

                            if (expectedContinuousIdle)
                            {
                                continuousIdleSteps++;
                                ambiguousObservationSteps = 0;
                                stagnationSteps = 0;
                                repeatCount = 0;
                                lastSig = null;
                                recentIneffectiveSpatialActions.Clear();
                            }
                            else if (!previousWasObservationOnly)
                            {
                                continuousIdleSteps = 0;
                                if (noChange)
                                {
                                    ambiguousObservationSteps = 0;
                                    stagnationSteps++;
                                }
                                else if (confirmedProgress || productiveTurnObservation)
                                {
                                    ambiguousObservationSteps = 0;
                                    stagnationSteps = 0;
                                }
                                else if (ambiguousProgress)
                                {
                                    ambiguousObservationSteps++;
                                    if (ambiguousObservationSteps >= 3)
                                        stagnationSteps++;
                                }
                                else
                                {
                                    ambiguousObservationSteps = 0;
                                }
                            }
                            else
                            {
                                continuousIdleSteps = 0;
                                if (confirmedProgress || productiveTurnObservation)
                                    stagnationSteps = 0;
                            }

                            if (confirmedStateChange)
                            {
                                actionCooldownUntilStep.Clear();
                                spatialActionCooldowns.Clear();
                            }
        
                            if (prevAction != null && !expectedContinuousIdle)
                            {
                                (repeatCount, lastSig) = UpdateRepeatDetection(
                                    prevAction,
                                    noChange,
                                    repeatCount,
                                    lastSig,
                                    recentIneffectiveSpatialActions);

                                if (ShouldRegisterImmediateNoEffectCooldown(
                                        prevAction.Action,
                                        lastObservation,
                                        repeatCount))
                                {
                                    RegisterActionCooldown(
                                        prevAction,
                                        step + ActionRepeatCooldownSteps,
                                        actionCooldownUntilStep,
                                        spatialActionCooldowns,
                                        clusterSpatially: repeatCount > 0);
                                }
                            }
        
                            // Expire AIM after a large visual change
                            var broadDelta = Math.Max(lastGlobalDelta, lastActiveWindowDelta);
                            if (lastAimRect is not null && broadDelta > AimExpireDelta)
                            {
                                Console.WriteLine($"[aim] expired (delta={broadDelta:0.###} > {AimExpireDelta:0.###})");
                                lastAimRect = null;
                            }
        
                            if (noChange && prevAction != null && IsPointClickAction(prevAction.Action))
                            {
                                lastPrecisionHint = BuildPrecisionHint(prevAction, lastDelta, "Previous click produced little or no visible progress.");
                                lastPrecisionHintExpiresAfterStep = step + 1;
                            }
                            if (noChange && prevAction?.Action.Type is "drag_drop" or "drag_path")
                            {
                                lastPrecisionHint = prevAction.Action.Type == "drag_path"
                                    ? "Previous drag_path did not produce the expected local effect. Do not replay the same path; verify the active tool/surface or change the gesture semantics."
                                    : "Previous drag_drop did not visibly move the source or change the destination. Do not retry nearby coordinates; verify that the source is draggable, identify the semantic destination, or use a different interaction route.";
                                lastPrecisionHintExpiresAfterStep = step + 2;
                            }
                            if (observedNoChange && prevAction != null && IsTextInputAttemptAction(prevAction.Action))
                            {
                                textInputNoChangeAttempts++;
                                lastTextInputHint = BuildTextInputHint(prevAction.Action, lastDelta);
                                lastTextInputHintExpiresAfterStep = step + 4;
                                if (textInputNoChangeAttempts >= 2)
                                    textInputCooldownUntilStep = Math.Max(textInputCooldownUntilStep, step + 4);
                            }
                            else if (confirmedProgress || productiveTurnObservation)
                            {
                                lastPrecisionHint = null;
                                lastPrecisionHintExpiresAfterStep = 0;
                                lastTextInputHint = null;
                                lastTextInputHintExpiresAfterStep = 0;
                                textInputNoChangeAttempts = 0;
                                textInputCooldownUntilStep = 0;
                            }
                        }
        
                        var previousResponseIdForRequest = controlContextChain.PreviousResponseIdForRequest;
                        var historyTail = previousResponseIdForRequest != null || HistoryTailChars <= 0
                            ? ""
                            : TailHistory(historyBuffer, HistoryTailChars, HistoryTailLines);
                        var (screenCx, screenCy, _, _) = GetCursorPositionInPrimary();
                        var (cx, cy, cnx, cny) = CursorToImageCoordinates(screenCx, screenCy);
                        var appliedFocusRectForPrompt = ScreenRectToImage(appliedFocusRect);
                        var focusUiaRectForPrompt = ScreenRectToImage(focusUiaRect);
                        var currentTurnFocusPath =
                            appliedFocusRect is not null &&
                            focusUrl != null &&
                            LogScreens
                                ? ScreenLogPath(screensDir, $"{commandId}_{step}_crop")
                                : null;
                        var currentTurnEvidenceDataUrl =
                            appliedFocusRect is Rectangle activeFocusRegion &&
                            turnBasedInteractionRect is Rectangle activeTurnRegion &&
                            activeFocusRegion == activeTurnRegion &&
                            focusUrl != null
                                ? focusUrl
                                : null;
                        var currentTurnEvidencePath = currentTurnEvidenceDataUrl is not null
                            ? currentTurnFocusPath
                            : null;
                        if (currentTurnEvidenceDataUrl is null &&
                            turnBasedInteractionRegionIsAutomatic &&
                            turnBasedAutomaticRegionRefined &&
                            turnBasedInteractionRect is Rectangle automaticEvidenceRegion &&
                            TryBuildTurnRegionEvidenceImage(
                                dataUrl,
                                automaticEvidenceRegion,
                                observationFrame.ScreenBounds,
                                screensDir,
                                commandId,
                                step,
                                out var automaticEvidenceDataUrl,
                                out var automaticEvidencePath))
                        {
                            currentTurnEvidenceDataUrl = automaticEvidenceDataUrl;
                            currentTurnEvidencePath = automaticEvidencePath;
                        }
                        var hasPersistentTurnFocus =
                            currentTurnEvidenceDataUrl is not null;
                        var hasTurnVisualContext =
                            string.Equals(
                                observationSession.EffectiveProfile,
                                "turn_based_interaction",
                                StringComparison.Ordinal) &&
                            turnBasedInteractionRect is not null &&
                            currentTurnEvidenceDataUrl is not null;
                        if (hasPersistentTurnFocus && turnReferenceFocusDataUrl is null)
                        {
                            turnReferenceFocusDataUrl = currentTurnEvidenceDataUrl;
                            turnReferenceFocusPath = currentTurnEvidencePath;
                            Console.WriteLine("[turn-memory] captured initial interaction reference image.");
                        }
                        var loopAssessment = AssessVisualStateCycle(
                            loopStateGraph,
                            shotFingerprint,
                            activeWindowFingerprint,
                            promptContext,
                            step,
                            recentExecutedActions,
                            prevAction,
                            lastDelta,
                            goalMode: goalMode,
                            recurringWorkflowIntent:
                                recurringWorkflowIntent,
                            goalProgress:
                                lastObservation?.GoalProgress ?? GoalProgressState.Unknown);
                        var proactiveVisualCycle = loopAssessment.IsLoop;
                        var visualCycleLength = loopAssessment.CycleLength;
                        if (proactiveVisualCycle)
                            Console.WriteLine($"[loop] proactive visual cycle detected; confidence={loopAssessment.Confidence:0.00}; returned after {visualCycleLength} step(s); {loopAssessment.Evidence}");
                        else if (loopAssessment.IsProductiveCycle)
                            Console.WriteLine($"[loop] productive recurring workflow observed; confidence={loopAssessment.Confidence:0.00}; returned after {visualCycleLength} step(s); no recovery intervention");
                        else if (loopAssessment.Confidence >= 0.5)
                            Console.WriteLine($"[loop] possible visual recurrence not yet confirmed; confidence={loopAssessment.Confidence:0.00}; threshold={loopAssessment.DecisionThreshold:0.00}; {loopAssessment.Evidence}");
                        AppendLoopReplayObservation(
                            commandId,
                            step,
                            screenW,
                            screenH,
                            shotFingerprint,
                            activeWindowFingerprint,
                            promptContext,
                            prevAction,
                            lastDelta,
                            loopAssessment,
                            lastObservation,
                            goalMode,
                            recurringWorkflowIntent);

                        recoveryEpisode = await UpdateRecoveryEpisodeAsync(
                            recoveryEpisode,
                            step,
                            stagnationSteps,
                            repeatCount,
                            lastDelta,
                            lastObservation,
                            shotFingerprint,
                            activeWindowFingerprint,
                            prevAction,
                            promptContext,
                            recentExecutedActions,
                            recoveryLessons,
                            loopAssessment,
                            planningLoopDetected,
                            rejectedProposalCycleLength,
                            previousRejectedAction,
                            recentRejectedActions,
                            goal,
                            goalMode,
                            dataUrl,
                            savedPath,
                            async episode =>
                            {
                                var progressImage = DownscaleDataUrlForHelperCall(
                                    dataUrl,
                                    savedPath,
                                    VerifyScreenshotMaxWidth);
                                var (progressW, progressH) = HelperImageSize(
                                    savedPath,
                                    imageW,
                                    imageH,
                                    VerifyScreenshotMaxWidth);
                                return await VerifyRecoveryProgressAsync(
                                    apiKey,
                                    goal,
                                    goalMode,
                                    episode,
                                    progressImage,
                                    savedPath,
                                    promptContext,
                                    progressW,
                                    progressH,
                                    requestsDir,
                                    commandId,
                                    step,
                                    cancelCts.Token);
                            });

                        if (MaxStagnationStepsBeforeAbort > 0 && stagnationSteps >= MaxStagnationStepsBeforeAbort)
                        {
                            Console.WriteLine($"[guard] stopping: no visible progress for {stagnationSteps} consecutive step(s). Use --max-stagnation 0 to disable.");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.GuardStopped,
                                step,
                                $"no visible progress for {stagnationSteps} consecutive steps");
                            break;
                        }

                        if (MaxRepeatedActionBeforeAbort > 0 && repeatCount >= MaxRepeatedActionBeforeAbort)
                        {
                            Console.WriteLine($"[guard] stopping: repeated ineffective action {repeatCount} time(s). Use --max-repeated-actions 0 to disable.");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.GuardStopped,
                                step,
                                $"ineffective action repeated {repeatCount} times");
                            break;
                        }
                        if (MaxRejectedProposalRepeatsBeforeAbort > 0 &&
                            rejectedProposalCycleCount >=
                            MaxRejectedProposalRepeatsBeforeAbort)
                        {
                            Console.WriteLine(
                                $"[guard] stopping: rejected model proposal cycle repeated {rejectedProposalCycleCount} time(s); cycle_length={rejectedProposalCycleLength}. Use --max-rejected-proposals 0 to disable.");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.GuardStopped,
                                step,
                                $"rejected model proposal cycle repeated {rejectedProposalCycleCount} times");
                            break;
                        }
                        var recoveryMemoryPrompt = BuildRecoveryMemoryPrompt(
                            recoveryLessons,
                            promptContext,
                            shotFingerprint,
                            activeWindowFingerprint,
                            prevAction,
                            recentExecutedActions,
                            recoveryEpisode,
                            stagnationSteps,
                            repeatCount,
                            goal);
        
                        var reuseUiaTargets = ReuseUiaTargetsWhenScreenUnchanged &&
                                              lastObservation?.VisualChange == VisualChangeState.Stable &&
                                              lastObservation.SemanticStateChanged == false &&
                                              CurrentUiaTargets.Count > 0;
                        PrepareUiaTargetsForPrompt(reuseUiaTargets, screenW, screenH);
        
                        RequestReasoningEffortOverride = EffectiveReasoningEffort(
                            stagnationSteps,
                            repeatCount,
                            rejectedProposalCycleCount);
                        if (!string.IsNullOrWhiteSpace(lastExecutorFailure))
                            shortTermPlan.Invalidate($"executor rejected or failed the planned action: {TrimForMeta(lastExecutorFailure, 160)}");
                        var controlMaxOutputTokens =
                            EffectiveControlMaxOutputTokens(
                                turnBasedTransitions.RequiresReanalysis);
                        if (turnBasedTransitions.RequiresReanalysis &&
                            controlMaxOutputTokens > MaxOutputTokens)
                        {
                            Console.WriteLine(
                                $"[openai] salient turn reanalysis budget elevated: {MaxOutputTokens}->{controlMaxOutputTokens} tokens.");
                        }
        
                        // inject observation metrics into the prompt
                        var metaSb = new StringBuilder()
                            .AppendLine($"LAST_STEP_DELTA: {(double.IsNaN(lastDelta) ? "N/A" : lastDelta.ToString("0.####"))} (observation_threshold={lastObservation?.ChangeThreshold ?? NoChangeThreshold:0.####})")
                            .AppendLine($"LAST_GLOBAL_DELTA: {(double.IsNaN(lastGlobalDelta) ? "N/A" : lastGlobalDelta.ToString("0.####"))}; LAST_ACTIVE_WINDOW_DELTA: {(double.IsNaN(lastActiveWindowDelta) ? "N/A" : lastActiveWindowDelta.ToString("0.####"))}")
                            .AppendLine($"LAST_LOCAL_DELTA: {(lastObservation is not null && double.IsFinite(lastObservation.LocalDelta) ? lastObservation.LocalDelta.ToString("0.####") : "N/A")}; LAST_LOCAL_CHANGED_RATIO: {(lastObservation is null ? "N/A" : lastObservation.LocalChangedRatio.ToString("0.####"))}")
                            .AppendLine($"OBSERVATION_PROFILE: {lastObservation?.Profile ?? observationSession.EffectiveProfile}; ACTION_OBSERVATION_POLICY: {lastObservation?.ActionPolicy ?? "N/A"}; VISUAL_CHANGE: {lastObservation?.VisualChange.ToString() ?? "N/A"}; ACTION_OUTCOME: {lastObservation?.ActionOutcome.ToString() ?? "N/A"}; GOAL_PROGRESS: {lastObservation?.GoalProgress.ToString() ?? "N/A"}; OBSERVATION_CONFIDENCE: {(lastObservation is null ? "N/A" : lastObservation.Confidence.ToString("0.00"))}")
                            .AppendLine($"STAGNATION_STEPS: {stagnationSteps}")
                            .AppendLine($"REPEAT_COUNT: {repeatCount}")
                            .AppendLine($"REJECTED_PROPOSAL_CYCLE_COUNT: {rejectedProposalCycleCount}")
                            .AppendLine($"CONTINUOUS_IDLE_STEPS: {continuousIdleSteps}")
                            .AppendLine($"AMBIGUOUS_OBSERVATION_STEPS: {ambiguousObservationSteps}")
                            .AppendLine($"INSPECTION_ACTIONS: {observationActionGuard.ConsecutiveInspectionActions}/{(MaxConsecutiveInspectionActions > 0 ? MaxConsecutiveInspectionActions.ToString() : "unlimited")}; AIM_SINCE_INTERACTION: {observationActionGuard.AimIssuedSinceInteraction.ToString().ToLowerInvariant()}")
                            .AppendLine($"GOAL_MODE: {goalMode}")
                            .AppendLine($"REQUEST_REASONING_EFFORT: {RequestReasoningEffortOverride ?? ReasoningEffort ?? "default"}")
                            .AppendLine($"REQUEST_MAX_OUTPUT_TOKENS: {controlMaxOutputTokens}")
                            .AppendLine($"LAST_ACTION: {(prevAction == null ? "N/A" : prevAction.Description)}")
                            .AppendLine($"AIM_ACTIVE: {(lastAimRect is null ? "false" : $"true {FormatImageRect(CurrentScreenMap.ScreenToImageRect(lastAimRect.Value))}")}")
                            .AppendLine("CONTEXT_CHECKPOINT_VERSION: 1; the current screen and runtime state below are authoritative when they conflict with earlier model reasoning.");
                        var recentActionCheckpoint = BuildRecentActionCheckpoint(recentExecutedActions);
                        if (!string.IsNullOrWhiteSpace(recentActionCheckpoint))
                            metaSb.AppendLine(recentActionCheckpoint);
                        var shortTermPlanSummary = shortTermPlan.BuildPromptSummary();
                        if (!string.IsNullOrWhiteSpace(shortTermPlanSummary))
                            metaSb.AppendLine(shortTermPlanSummary);
                        if (repeatCount > 0 || stagnationSteps > 0)
                            metaSb.AppendLine("STRATEGY_HINT: The previous action did not visibly advance the screen. Do not repeat it; choose a different UI route or ask for a crop if the target is ambiguous.");
                        if (prevAction is not null &&
                            lastObservation is not null &&
                            ShouldRegisterImmediateNoEffectCooldown(
                                prevAction.Action,
                                lastObservation,
                                repeatCount))
                        {
                            metaSb.AppendLine(
                                "NO_EFFECT_ACTION_BLOCKED: LAST_ACTION is unavailable until the observed state changes. Do not propose it again; choose a different input modality, target, or diagnostic action.");
                        }
                        if (expectedContinuousIdle)
                            metaSb.AppendLine("CONTINUOUS_IDLE: The previous wait left the screen unchanged, which is valid for this open-ended goal. Reassess whether the requested state is still healthy or whether a new event is present; wait again only when continued observation is goal-aligned.");
                        if (ambiguousObservationSteps > 0)
                            metaSb.AppendLine("AMBIGUOUS_ACTION_OUTCOME: Visual activity was observed, but it could not be attributed confidently to the previous action. Reassess the state and expected effect; do not blindly repeat the same input.");
                        if (observationActionGuard.RequiresInteraction(MaxConsecutiveInspectionActions))
                            metaSb.AppendLine("INTERACTION_REQUIRED: The inspection budget is exhausted. request_crop and point are mechanically blocked until RDPilot executes a state-changing interaction. Use visible controls or one safe, reversible input and then observe its effect; use aim only when precise pointer targeting is required.");
                        else if (string.Equals(
                                     observationSession.EffectiveProfile,
                                     "turn_based_interaction",
                                     StringComparison.Ordinal) &&
                                 observationActionGuard.ConsecutiveInspectionActions >= 1)
                            metaSb.AppendLine("TURN_INSPECTION_COMPLETE: a detailed crop has already been supplied for the current state. Another nested request_crop is blocked until a state-changing input is attempted.");
                        if (string.Equals(observationSession.EffectiveProfile, "turn_based_interaction", StringComparison.Ordinal))
                        {
                            metaSb.AppendLine(turnBasedTransitions.CanProposeExecutionBatch
                                ? $"TURN_BASED_INTERACTION: aggressive execution_ready batching is the default from the first visible route. Commit to the strongest visible control hypothesis and test it with a reversible route; semantic uncertainty is not a reason to inspect HELP, narrate, or delay. Before the first directional input, request_crop is justified only when the board is physically unreadable at the supplied resolution, and HELP/instructions are justified only when no reversible control input is visible. Preliminary single-move calibration is forbidden when at least two reversible moves form a coherent route. Bind the route to TURN_STATE with planned_inputs, plan_waypoint, and plan_confidence. When fixed visible directional controls such as a D-pad exist and keyboard focus has not been confirmed by a successful move, prefer an ordered click sequence over keyboard keys. If a keyboard direction had no effect, preserve the logical route and immediately remap its longest valid prefix to those visible controls; do not test only one button. Send up to {turnBasedTransitions.AdvertisedMaxExecutionBatchLength} observed click/key inputs and include every visible unconditional turn before the next semantic uncertainty; do not pad a route merely to reach the cap. A small or distant recurring auxiliary UI change is recorded but does not stop execution. A blocked input, unavailable observation, broad screen/state transition, or novel local-to-distant causal change interrupts the batch. Do not use hold_keys for discrete movement."
                                : "TURN_BASED_INTERACTION: The interaction is in an exploration or reanalysis phase. Promptly perform exactly one maximally informative, reversible input, then inspect the resulting frame. Do not spend this turn narrating or solving a full route. Compare labeled temporal images and the reported change regions before acting. Prefer a visible directional/control button or a single Arrow/WASD/Space/F key when the interface suggests those controls. Keep planned_inputs null unless the route is already mechanically established. Do not use hold_keys or batch speculative moves.");
                            metaSb.AppendLine(
                                "TURN_GRID_ROUTE_HINT: On a visible grid, board, or maze, trace and count the route cell by cell from the controlled object's current footprint to an evidence-backed semantic waypoint before emitting planned_inputs. Every planned input must correspond to exactly one adjacent traversable cell, every intermediate footprint must remain on the connected visible path, and the number and order of repeated directions must match the counted grid intervals. Do not estimate distance from a qualitative relation such as 'near', 'below-right', or apparent screen proximity, guess through visible barriers, select an unverified marker only because it is salient, or repeat a closed route that returned to a known state. Aggressive batching means executing the longest route that passes this geometric check decisively.");
                            if (turnBasedTransitions.CanProposeExecutionBatch)
                            {
                                metaSb.AppendLine(
                                    "TURN_ROUTE_DECISION_MODE: commit_fast. Do not exhaustively solve hidden mechanics before acting. Choose the strongest visible unconditional route, emit its structured key sequence immediately, and let the per-input observation barrier expose mistaken assumptions.");
                            }
                            var transitionSummary = turnBasedTransitions.BuildPromptSummary();
                            if (!string.IsNullOrWhiteSpace(transitionSummary))
                                metaSb.AppendLine(transitionSummary);
                        }
                        if (recoveryEpisode != null &&
                            stagnationSteps < RecoveryMemoryTriggerSteps &&
                            repeatCount < 1)
                        {
                            metaSb.AppendLine("PROACTIVE_LOOP_SUSPECTED: recent actions indicate an emerging loop. Change strategy now, before a safety guard is reached.");
                        }
                        if (proactiveVisualCycle)
                            metaSb.AppendLine($"MULTI_STEP_LOOP_DETECTED: confidence={loopAssessment.Confidence:0.00}; the UI returned to a prior visual state after {visualCycleLength} steps with corroborating evidence. Do not repeat the intervening action sequence.");
                        else if (loopAssessment.IsProductiveCycle)
                            metaSb.AppendLine($"PRODUCTIVE_CYCLE_OBSERVED: confidence={loopAssessment.Confidence:0.00}; this state return matches an intentional recurring workflow and is not currently treated as a harmful loop. Continue only while the cycle remains goal-aligned and produces useful checks, maintenance, or event handling.");
                        else if (loopAssessment.Confidence >= 0.5)
                            metaSb.AppendLine($"LOOP_CANDIDATE: confidence={loopAssessment.Confidence:0.00}, below calibrated threshold={loopAssessment.DecisionThreshold:0.00}. This is not yet a confirmed loop; seek another recurrence signal before changing a valid strategy.");
                        if (planningLoopDetected)
                            metaSb.AppendLine($"PLANNING_LOOP_DETECTED: rejected proposal cycle length={rejectedProposalCycleLength}, repeated={rejectedProposalCycleCount}. Do not propose the blocked action sequence again; choose a materially different, currently executable route.");
                        if (!string.IsNullOrWhiteSpace(recoveryMemoryPrompt))
                            metaSb.AppendLine(recoveryMemoryPrompt);
                        if (!string.IsNullOrWhiteSpace(lastVerifierRejection))
                            metaSb.AppendLine($"LAST_VERIFY_REJECTION: {TrimForMeta(lastVerifierRejection, 240)}");
                        if (!string.IsNullOrWhiteSpace(lastExecutorFailure))
                            metaSb.AppendLine($"LAST_EXECUTOR_FAILURE: {TrimForMeta(lastExecutorFailure, 240)}");
                        if (lastPrecisionHintExpiresAfterStep > 0 && step > lastPrecisionHintExpiresAfterStep)
                        {
                            lastPrecisionHint = null;
                            lastPrecisionHintExpiresAfterStep = 0;
                        }
                        if (!string.IsNullOrWhiteSpace(lastPrecisionHint))
                            metaSb.AppendLine($"PRECISION_HINT: {TrimForMeta(lastPrecisionHint, 420)}");
                        if (lastTextInputHintExpiresAfterStep > 0 && step > lastTextInputHintExpiresAfterStep)
                        {
                            lastTextInputHint = null;
                            lastTextInputHintExpiresAfterStep = 0;
                        }
                        if (!string.IsNullOrWhiteSpace(lastTextInputHint))
                            metaSb.AppendLine($"TEXT_INPUT_HINT: {TrimForMeta(lastTextInputHint, 420)}");
                        if (textInputCooldownUntilStep >= step)
                            metaSb.AppendLine($"TEXT_INPUT_COOLDOWN: active until step {textInputCooldownUntilStep}; do not use paste_text, type_text, or paste shortcuts until focus/editability visibly changes.");
        
                        var omitFullScreenImage = OmitUnchangedScreenImageWithState &&
                                                  previousResponseIdForRequest != null &&
                                                  lastObservation?.VisualChange == VisualChangeState.Stable;
                        if (omitFullScreenImage)
                            metaSb.AppendLine("SCREEN_IMAGE: omitted because screen fingerprint is unchanged; use previous_response_id state plus current metadata.");

                        var turnChangeImages = new List<TurnChangeImagePair>();
                        if (hasPersistentTurnFocus &&
                            turnBasedTransitions.RequiresReanalysis &&
                            previousTurnFocusDataUrl is not null &&
                            currentTurnEvidenceDataUrl is not null)
                        {
                            var regionIndex = 1;
                            foreach (var changeRegion in turnBasedTransitions.SalientChangeRegions)
                            {
                                var pair = BuildTurnChangeImagePair(
                                    previousTurnFocusDataUrl,
                                    currentTurnEvidenceDataUrl,
                                    changeRegion,
                                    screensDir,
                                    commandId,
                                    step,
                                    regionIndex++);
                                if (pair is not null)
                                    turnChangeImages.Add(pair);
                            }
                            if (turnChangeImages.Count > 0)
                            {
                                activeTurnChangeImages = turnChangeImages;
                                Console.WriteLine(
                                    $"[turn-event] attached {turnChangeImages.Count} focused before/after change-region pair(s).");
                            }
                        }
                        IReadOnlyList<TurnChangeImagePair> requestTurnChangeImages =
                            turnBasedTransitions.RequiresReanalysis ||
                            turnBasedTransitions.HasActiveCausalEvent
                                ? activeTurnChangeImages
                                : [];
                        var includeTurnTemporalImages =
                            hasTurnVisualContext &&
                            (turnBasedTransitions.RequiresReanalysis ||
                             turnBasedTransitions.HasActiveCausalEvent ||
                             !turnBasedTransitions.CanUseExecutionBatch ||
                             !shortTermPlan.HasActivePlan);
                        if (hasTurnVisualContext && !includeTurnTemporalImages)
                        {
                            metaSb.AppendLine(
                                "TURN_TEMPORAL_IMAGES: omitted during predictable plan execution; use CURRENT_FOCUS_IMAGE and the transition ledger. Full before/reference evidence will return after an invalidating or salient change.");
                        }

                        var focusDataUrlForRequest = focusUrl ?? currentTurnEvidenceDataUrl;
                        var focusRectForRequest = appliedFocusRectForPrompt ??
                                                  (focusUrl is null &&
                                                   turnBasedInteractionRect is Rectangle evidenceRegion
                                                      ? ScreenRectToImage(evidenceRegion)
                                                      : null);
                        var focusPathForRequest = currentTurnFocusPath ?? currentTurnEvidencePath;
                        var previousTurnImageForRequest = includeTurnTemporalImages
                            ? previousTurnFocusDataUrl
                            : null;
                        var referenceTurnImageForRequest = includeTurnTemporalImages
                            ? turnReferenceFocusDataUrl
                            : null;
                        var previousTurnImageAttached =
                            previousTurnImageForRequest is not null &&
                            !string.Equals(
                                previousTurnImageForRequest,
                                focusDataUrlForRequest,
                                StringComparison.Ordinal);
                        var referenceTurnImageAttached =
                            referenceTurnImageForRequest is not null &&
                            !string.Equals(
                                referenceTurnImageForRequest,
                                focusDataUrlForRequest,
                                StringComparison.Ordinal) &&
                            !string.Equals(
                                referenceTurnImageForRequest,
                                previousTurnImageForRequest,
                                StringComparison.Ordinal);
                        metaSb.AppendLine(
                            $"TURN_TEMPORAL_EVIDENCE: current={(focusDataUrlForRequest is null ? "absent" : "attached")}; previous={(previousTurnImageAttached ? "attached" : "absent")}; reference={(referenceTurnImageAttached ? "attached" : "absent")}; change_pairs={requestTurnChangeImages.Count}.");
                        var reqBody = BuildRequestBody(Model, systemRules, goal, historyTail + "\n" + metaSb, dataUrl, imageW, imageH,
                                                       cx, cy, cnx, cny, focusDataUrlForRequest, focusRectForRequest, focusUiaRectForPrompt, focusUiaDataUrl,
                                                       promptContext, reuseUiaTargets, previousResponseIdForRequest, omitFullScreenImage, goalMode,
                                                       previousTurnImageForRequest,
                                                        referenceTurnImageForRequest,
                                                        requestTurnChangeImages,
                                                        controlMaxOutputTokens,
                                                        controlContextChain.CompactionEnabled,
                                                        controlContextChain.Enabled);
                        if (LogRequests)
                        {
                            var reqBodyForLog = BuildRequestBody_ForLog(Model, systemRules, goal, historyTail + "\n" + metaSb,
                                                                        omitFullScreenImage ? null : savedPath, imageW, imageH, cx, cy, cnx, cny,
                                                                        focusPathForRequest,
                                                                        focusRectForRequest,
                                                                        focusUiaRectForPrompt, focusUiaPath, promptContext, previousResponseIdForRequest, omitFullScreenImage, goalMode,
                                                                        includeTurnTemporalImages ? previousTurnFocusPath : null,
                                                                         includeTurnTemporalImages ? turnReferenceFocusPath : null,
                                                                         requestTurnChangeImages,
                                                                         controlMaxOutputTokens,
                                                                         controlContextChain.CompactionEnabled,
                                                                         controlContextChain.Enabled);
                            SaveJson(Path.Combine(requestsDir, $"{commandId}_{step}_request.json"), reqBodyForLog);
                        }
                        if (hasPersistentTurnFocus)
                        {
                            previousTurnFocusDataUrl = currentTurnEvidenceDataUrl;
                            previousTurnFocusPath = currentTurnEvidencePath;
                        }
        
                        controlContextChain.LogRequest(step);
                        var (action, raw, completedResponseId, contextFallbackUsed, compactionFallbackUsed, compactionOccurred) = await CallOpenAIAsync(
                            apiKey,
                            reqBody,
                            cancelCts.Token,
                            AllowsStableCanvasDrawBatch(
                                promptContext,
                                observationSession.EffectiveProfile),
                            turnBasedInteractionRect is not null &&
                            turnBasedTransitions.CanProposeExecutionBatch
                                ? Math.Max(
                                    0,
                                    turnBasedTransitions.AdvertisedMaxExecutionBatchLength - 1)
                                : 0);
                        controlContextChain.RecordResult(
                            completedResponseId,
                            contextFallbackUsed,
                            compactionFallbackUsed,
                            compactionOccurred);
                        RequestReasoningEffortOverride = null;
                        SaveRaw(Path.Combine(requestsDir, $"{commandId}_{step}_response.json"), raw);
        
                        if (action is null)
                        {
                            if (CancelRequested)
                            {
                                Console.WriteLine("Aborted (hotkey).");
                                runResult = new ControlRunResult(
                                    ControlRunOutcome.Cancelled,
                                    step,
                                    "cancelled with the emergency hotkey");
                                break;
                            }
        
                            consecutiveModelFailures++;
                            if (LastOpenAiFailureWasRetriable &&
                                MaxModelFailuresBeforeAbort > 0 &&
                                consecutiveModelFailures < MaxModelFailuresBeforeAbort)
                            {
                                Console.WriteLine($"[openai] transient {LastOpenAiFailureKind}; keeping goal alive ({consecutiveModelFailures}/{MaxModelFailuresBeforeAbort}).");
                                AddHistory(historyBuffer, $"[{step}] model_failure_retry: {LastOpenAiFailureKind}");
                                prevAction = null;
                                prevShotFingerprint = null;
                                prevActiveWindowFingerprint = null;
                                prevObservationFrame = null;
                                prevObservationContext = null;
                                await Task.Delay(750, cancelCts.Token);
                                continue;
                            }
        
                            Console.WriteLine("Could not parse action. Aborting this goal.");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.Failed,
                                step,
                                "the model response did not contain a valid action");
                            break;
                        }
                        consecutiveModelFailures = 0;
                        var modelResponseAction = action;
        
                        var currentAction = AttachFocusedTextObservationRegion(
                            CaptureResolvedAction(action, lastAimRect),
                            focusUiaRect);
                        var currentActionPolicy = observationSession.ResolveActionPolicy(
                            currentAction,
                            promptContext,
                            goal);
                        currentAction = AttachTurnBasedObservationRegion(
                            currentAction,
                            ResolveTurnBasedObservationRegion(
                                turnBasedInteractionRect,
                                currentActionPolicy),
                            currentActionPolicy);
                        var isTurnBasedAction = string.Equals(
                            currentActionPolicy,
                            "turn_based_interaction",
                            StringComparison.Ordinal);
                        shortTermPlan.Update(
                            action,
                            isTurnBasedAction
                                ? turnBasedTransitions.CurrentStateId
                                : null);
                        if (isTurnBasedAction &&
                            PendingSafeActions.Count == 0 &&
                            action.Type is "click" or "double_click" &&
                            action.PlannedInputs is { Length: > 1 } observedPlannedInputs &&
                            OpenAiResponsesService.TryBindObservedTurnActionToPlannedInput(
                                action,
                                observedPlannedInputs[0]) &&
                            shortTermPlan.TryExpandDirectionalSequence(
                                [observedPlannedInputs[0]],
                                turnBasedTransitions.CurrentStateId,
                                turnBasedTransitions.MaxExecutionBatchLength,
                                out var expandedObservedInputs) &&
                            turnBasedTransitions.TryBuildObservedDirectionalFollowUps(
                                action,
                                expandedObservedInputs.Skip(1).ToArray(),
                                out var inferredObservedFollowUps))
                        {
                            foreach (var followUp in inferredObservedFollowUps)
                                PendingSafeActions.Enqueue(followUp);
                            Console.WriteLine(
                                $"[turn-plan] expanded one observed click into {expandedObservedInputs.Length} route inputs using learned control positions.");
                        }
                        if (isTurnBasedAction &&
                            CanExpandTurnKeyActionFromPlan(
                                action,
                                PendingSafeActions.Count) &&
                            action.Keys is { Length: > 0 } proposedKeys &&
                            TryGetTurnBasedDirectionalSequenceLength(
                                action,
                                out _) &&
                            turnBasedTransitions.CanUseExecutionBatch &&
                            shortTermPlan.TryExpandDirectionalSequence(
                                proposedKeys,
                                turnBasedTransitions.CurrentStateId,
                                turnBasedTransitions.MaxExecutionBatchLength,
                                out var expandedKeys))
                        {
                            Console.WriteLine(
                                $"[turn-plan] expanded model batch {proposedKeys.Length}->{expandedKeys.Length} from the high-confidence structured route.");
                            action.Keys = expandedKeys;
                            currentAction = AttachTurnBasedObservationRegion(
                                AttachFocusedTextObservationRegion(
                                    CaptureResolvedAction(action, lastAimRect),
                                    focusUiaRect),
                                ResolveTurnBasedObservationRegion(
                                    turnBasedInteractionRect,
                                    currentActionPolicy),
                                currentActionPolicy);
                        }
                        Console.WriteLine($"[{step}] {currentAction.Description}");
                        if (action.Confidence is double confidence)
                            Console.WriteLine($"     confidence: {confidence:0.##}");
                        if (!string.IsNullOrWhiteSpace(action.Note))
                            Console.WriteLine($"     note: {action.Note}");
        
                        nextFocusRect = null; // reset – set by aim/point/request_crop
                        var actionExecutionFailed = false;
                        var actionExecuted = false;
                        var actionWasLocallyRejected = false;
                        var turnBatchObserved = false;
                        var turnReanalysisWasRequired =
                            turnBasedTransitions.RequiresReanalysis;
                        ScreenObservationFrame? actionBaselineFrame = null;
                        byte[]? actionBaselineLocalFingerprint = null;
                        try
                        {
                            if (!currentAction.IsValid)
                                throw new InvalidOperationException(currentAction.ValidationError);

                            if (string.Equals(
                                    currentActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal) &&
                                IsStateChangingInteractionAction(currentAction.Action) &&
                                turnBasedInteractionWindow != IntPtr.Zero &&
                                GetForegroundWindow() != turnBasedInteractionWindow &&
                                !TryRestoreForegroundWindow(turnBasedInteractionWindow))
                            {
                                throw new InvalidOperationException(
                                    "the remembered turn-based interaction window could not regain focus");
                            }

                            if (string.Equals(
                                    currentActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal) &&
                                TryGetTurnBasedDirectionalSequenceLength(
                                    currentAction.Action,
                                    out var directionalSequenceLength) &&
                                directionalSequenceLength > 1 &&
                                (!turnBasedTransitions.CanUseExecutionBatch ||
                                 directionalSequenceLength >
                                 turnBasedTransitions.MaxExecutionBatchLength))
                            {
                                throw new InvalidOperationException(
                                    turnBasedTransitions.CanUseExecutionBatch
                                        ? $"turn-based execution batches currently accept at most {turnBasedTransitions.MaxExecutionBatchLength} directional inputs"
                                        : "multiple directional inputs are blocked until turn-based exploration is complete");
                            }

                            if (string.Equals(
                                    currentActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal) &&
                                TryGetTurnBasedDirectionalSequenceLength(
                                    currentAction.Action,
                                    out _) &&
                                currentAction.Action.Keys is { Length: > 0 } routeKeys &&
                                !shortTermPlan.ProposedSequenceMatches(
                                    routeKeys,
                                    turnBasedTransitions.CurrentStateId,
                                    out var routeMismatchReason))
                            {
                                shortTermPlan.Invalidate(routeMismatchReason);
                                throw new InvalidOperationException(routeMismatchReason);
                            }

                            if (string.Equals(
                                    currentActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal) &&
                                turnBasedTransitions.RequiresReanalysis &&
                                IsStateChangingInteractionAction(currentAction.Action) &&
                                !turnBasedTransitions.HasRequiredSalientObservation(action))
                            {
                                const string missingObservationReason =
                                    "a salient world change must be described in salient_change_observation before another state-changing input";
                                Console.WriteLine(
                                    $"[guard] turn reanalysis action blocked: {missingObservationReason}");
                                AddHistory(
                                    historyBuffer,
                                    $"[{step}] IGNORED (salient_change_not_observed): {currentAction.Description}");
                                lastExecutorFailure = missingObservationReason;
                                PendingSafeActions.Clear();
                                RegisterRejectedProposal(currentAction);
                                prevAction = null;
                                prevShotFingerprint = null;
                                prevActiveWindowFingerprint = null;
                                prevObservationFrame = null;
                                prevObservationContext = null;
                                continue;
                            }

                            if (observationActionGuard.TryGetBlockReason(
                                    currentAction,
                                    MaxConsecutiveInspectionActions,
                                    string.Equals(
                                        currentActionPolicy,
                                        "turn_based_interaction",
                                        StringComparison.Ordinal),
                                    out var observationBlockReason))
                            {
                                Console.WriteLine($"[guard] observation action blocked: {observationBlockReason}");
                                AddHistory(historyBuffer, $"[{step}] IGNORED (observation_budget): {currentAction.Description}");
                                lastExecutorFailure = observationBlockReason;
                                PendingSafeActions.Clear();
                                RegisterRejectedProposal(currentAction);
                                prevAction = null;
                                prevShotFingerprint = null;
                                prevActiveWindowFingerprint = null;
                                prevObservationFrame = null;
                                prevObservationContext = null;
                                continue;
                            }

                            if (IsActionOnCooldown(currentAction, step, actionCooldownUntilStep, spatialActionCooldowns, out var cooldownUntil))
                            {
                                Console.WriteLine($"[guard] skipping repeated ineffective action until step {cooldownUntil}: {currentAction.Description}");
                                AddHistory(historyBuffer, $"[{step}] IGNORED (repeat_cooldown): {currentAction.Description}");
                                lastExecutorFailure = $"action temporarily blocked after no visible effect: {currentAction.Description}";
                                if (IsPointClickAction(action))
                                {
                                    lastPrecisionHint = BuildPrecisionHint(currentAction, double.NaN, "A repeated click was blocked after no visible effect.");
                                    lastPrecisionHintExpiresAfterStep = step + 1;
                                }
                                actionExecutionFailed = true;
                                RegisterRejectedProposal(currentAction);
                                prevAction = null;
                                prevShotFingerprint = null;
                                prevActiveWindowFingerprint = null;
                                prevObservationFrame = null;
                                prevObservationContext = null;
                                continue;
                            }
                            if (IsTextInputAttemptAction(action) && textInputCooldownUntilStep >= step)
                            {
                                Console.WriteLine($"[guard] skipping text input after repeated no-change attempts until step {textInputCooldownUntilStep}: {Describe(action)}");
                                AddHistory(historyBuffer, $"[{step}] IGNORED (text_input_no_progress): {Describe(action)}");
                                lastExecutorFailure = $"text input temporarily blocked after repeated no visible effect: {Describe(action)}";
                                lastTextInputHint = BuildTextInputCooldownHint(textInputNoChangeAttempts);
                                lastTextInputHintExpiresAfterStep = step + 3;
                                actionExecutionFailed = true;
                                RegisterRejectedProposal(currentAction);
                                prevAction = null;
                                prevShotFingerprint = null;
                                prevActiveWindowFingerprint = null;
                                prevObservationFrame = null;
                                prevObservationContext = null;
                                continue;
                            }

                            if (ActionNeedsEffectObservation(action) &&
                                (!IsMouseAction(action) || MouseEnabled))
                            {
                                actionBaselineFrame = CaptureObservationFrameProbe(
                                    currentAction.ObservationRegion is not null ||
                                    action.Type is "drag_drop" or "drag_path" or "type_text" or "paste_text");
                                if (currentAction.ObservationRegion is Rectangle localRegion)
                                {
                                    actionBaselineLocalFingerprint = CaptureRegionFingerprintProbe(localRegion);
                                }
                                if (string.Equals(
                                        currentActionPolicy,
                                        "turn_based_interaction",
                                        StringComparison.Ordinal) &&
                                    turnBasedInteractionRect is Rectangle persistentTurnRegion)
                                {
                                    var baselineAssessment = turnBasedTransitions.PrepareActionBaseline(
                                        actionBaselineFrame,
                                        persistentTurnRegion);
                                    if (baselineAssessment.ExternalStateChange)
                                    {
                                        Console.WriteLine(
                                            $"[turn] external state change detected while awaiting the model; discarding stale action; {baselineAssessment.Summary}");
                                        AddHistory(
                                            historyBuffer,
                                            $"[{step}] stale_action_discarded: external state changed while model was planning");
                                        PendingSafeActions.Clear();
                                        BeginExternalStateEpoch();
                                        lastExecutorFailure =
                                            "The interaction state changed externally while the model was planning. Reinspect the fresh state before acting.";
                                        prevAction = null;
                                        prevShotFingerprint = null;
                                        prevActiveWindowFingerprint = null;
                                        prevObservationFrame = null;
                                        prevObservationContext = null;
                                        continue;
                                    }
                                }
                                else if (ShouldCheckPreRegionTurnActionFreshness(
                                             currentActionPolicy,
                                             turnBasedInteractionRect,
                                             currentAction.Action))
                                {
                                    await Task.Delay(120, cancelCts.Token);
                                    var confirmationFrame = CaptureObservationFrameProbe(
                                        includeDetail: false);
                                    if (ShouldDiscardPreRegionTurnAction(
                                            observationFrame,
                                            actionBaselineFrame,
                                            confirmationFrame,
                                            out var promptDelta,
                                            out var stabilityDelta))
                                    {
                                        Console.WriteLine(
                                            $"[turn] pre-region state changed while awaiting the model; discarding stale action; prompt_delta={promptDelta:0.####}; stability_delta={stabilityDelta:0.####}");
                                        AddHistory(
                                            historyBuffer,
                                            $"[{step}] stale_action_discarded: pre-region screen changed while model was planning");
                                        PendingSafeActions.Clear();
                                        lastExecutorFailure =
                                            "The visible turn-based state changed while the model was planning. Reinspect the fresh screen and act on its current controls.";
                                        prevAction = null;
                                        prevShotFingerprint = null;
                                        prevActiveWindowFingerprint = null;
                                        prevObservationFrame = null;
                                        prevObservationContext = null;
                                        continue;
                                    }
                                    actionBaselineFrame = confirmationFrame;
                                }
                                observationSession.RecordAmbientMotion(
                                    observationFrame,
                                    actionBaselineFrame);
                            }
                            if (string.Equals(
                                    currentActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal))
                            {
                                turnBasedTransitions.UpdateWorkingMemory(
                                    modelResponseAction);
                            }
        
                            // ——— Mouse policy: global switch ———
                            if (IsMouseAction(action) && !MouseEnabled)
                            {
                                Console.WriteLine("[guard] mouse disabled → ignoring mouse action; use keyboard strategy or 'aim' without clicking.");
                                AddHistory(historyBuffer, $"[{step}] IGNORED (mouse_disabled)");
                                lastExecutorFailure = "mouse action was blocked because mouse input is disabled";
                                actionExecutionFailed = true;
                                actionWasLocallyRejected = true;
                            }
                            else if (action.Type == "aim")
                            {
                                var rect = ResolveAimRect(action);
                                if (rect is null) throw new InvalidOperationException("aim without parameters (bbox/crop/x/y/x_px/y_px).");
                                lastAimRect = rect.Value;
                                nextFocusRect = rect.Value; // show crop/overlay on next screenshot
                                actionExecuted = true;
                            }
                            else if (action.Type == "point")
                            {
                                // Visual pointer only – does NOT set AIM (clicks still blocked without 'aim')
                                var rect = ResolveCropRect(action);
                                if (rect is null) throw new InvalidOperationException("point without parameters.");
                                nextFocusRect = rect.Value;
                                actionExecuted = true;
                            }
                            else if (action.Type == "request_crop")
                            {
                                var rect = ResolveCropRect(action);
                                if (rect is null) throw new InvalidOperationException("request_crop without parameters.");
                                nextFocusRect = rect.Value;
                                if (string.Equals(
                                        currentActionPolicy,
                                        "turn_based_interaction",
                                        StringComparison.Ordinal))
                                {
                                    var clamped = ClampRect(rect.Value);
                                    var canReplaceTurnRegion =
                                        CanReplaceTurnBasedInteractionRegion(
                                            turnBasedInteractionRect is not null,
                                            turnBasedInteractionRegionIsAutomatic,
                                            turnBasedAutomaticRegionRefined,
                                            action,
                                            regionRequired: false);
                                    if (turnBasedInteractionRect is Rectangle existingTurnRegion &&
                                        !canReplaceTurnRegion)
                                    {
                                        var inspectionKind = IsOverlappingTurnInspection(
                                                existingTurnRegion,
                                                clamped)
                                            ? "overlapping"
                                            : "auxiliary";
                                        Console.WriteLine(
                                            $"[turn] transient {inspectionKind} inspection crop=({clamped.Left},{clamped.Top})–({clamped.Right},{clamped.Bottom}); preserving primary interaction region, transitions, and visual memory.");
                                    }
                                    else if (canReplaceTurnRegion)
                                    {
                                        turnBasedInteractionRect = clamped;
                                        turnBasedInteractionRegionIsAutomatic = false;
                                        turnBasedAutomaticRegionRefined = false;
                                        turnBasedInteractionWindow = GetForegroundWindow();
                                        turnBasedInteractionContext = currentTurnBasedContext;
                                        previousTurnFocusDataUrl = null;
                                        previousTurnFocusPath = null;
                                        turnReferenceFocusDataUrl = null;
                                        turnReferenceFocusPath = null;
                                        activeTurnChangeImages = [];
                                        turnBasedTransitions.Reset();
                                        Console.WriteLine(
                                            $"[turn] persistent interaction region set to ({clamped.Left},{clamped.Top})–({clamped.Right},{clamped.Bottom}).");
                                    }
                                }
                                actionExecuted = true;
                            }
                            else if (action.Type == "wait")
                            {
                                int secs = EffectiveWaitSeconds(action, out var requestedSecs);
                                if (secs < requestedSecs)
                                    Console.WriteLine($"[wait] Requested {requestedSecs}s capped to {secs}s.");
                                Console.WriteLine($"[wait] Sleeping {secs} s (long-running operation on screen)...");
                                await Task.Delay(secs * 1000, cancelCts.Token);
                                actionExecuted = true;
                            }
                            else if (action.Type == "done")
                            {
                                if (goalMode == "continuous")
                                {
                                    lastVerifierRejection =
                                        "The goal is continuous and has no natural completion. Continue meaningful activity; do not mark it complete. Termination is governed by user abort and configured runtime safety guards or limits.";
                                    Console.WriteLine($"[verify] continuous goal rejected 'done': {lastVerifierRejection}");
                                    AddHistory(historyBuffer, $"[{step}] done_rejected_continuous");
                                    actionExecutionFailed = true;
                                    RegisterRejectedProposal(currentAction);
                                    // This is a planning correction, not an executor
                                    // failure. An open-ended goal must remain alive even
                                    // if the model proposes done repeatedly.
                                }
                                else
                                {
                                    var verifyDataUrl = dataUrl;
                                    var verifyPath = savedPath;
                                    var verifyRealScreenW = screenW;
                                    var verifyRealScreenH = screenH;
                                    var verifyScreenW = imageW;
                                    var verifyScreenH = imageH;
                                    var verifyFocusUiaScreenRect = focusUiaRect;
                                    var verifyPromptContext = promptContext;
                                    if (RefreshScreenshotBeforeVerify)
                                    {
                                        await Task.Delay(UiSettleDelayMs, cancelCts.Token); // give UI time for slow apps when explicitly requested
                                        var (freshDataUrl, freshPath, freshW, freshH, freshImageW, freshImageH, _, _, freshFocusRect, freshFocusSummary, _, _, _, _, _) = ScreenshotToDataUrl(screensDir, commandId, step, null);
                                        verifyDataUrl = freshDataUrl;
                                        verifyPath = freshPath;
                                        verifyRealScreenW = freshW;
                                        verifyRealScreenH = freshH;
                                        verifyScreenW = freshImageW;
                                        verifyScreenH = freshImageH;
                                        verifyFocusUiaScreenRect = freshFocusRect;
                                        verifyPromptContext = CaptureUiPromptContext(freshFocusSummary, freshW, freshH);
                                    }
                                    verifyDataUrl = DownscaleDataUrlForHelperCall(verifyDataUrl, verifyPath, VerifyScreenshotMaxWidth);
                                    (verifyScreenW, verifyScreenH) = HelperImageSize(verifyPath, verifyScreenW, verifyScreenH, VerifyScreenshotMaxWidth);

                                    var previousScreenMap = CurrentScreenMap;
                                    SetCurrentScreenMap(verifyRealScreenW, verifyRealScreenH, verifyScreenW, verifyScreenH);
                                    var verifyFocusUiaRect = ScreenRectToImage(verifyFocusUiaScreenRect);
                                    VerifyDto? verify;
                                    var independentlyVerified = ShouldVerifyGoal(goal, step, action);
                                    try
                                    {
                                        verify = independentlyVerified
                                            ? await VerifyGoalAsync(apiKey, goal, verifyDataUrl, verifyPath, verifyScreenW, verifyScreenH, verifyFocusUiaRect, verifyPromptContext, requestsDir, commandId, step, cancelCts.Token)
                                            : new VerifyDto { Verdict = "yes", Reason = "verification skipped by mode" };
                                    }
                                    finally
                                    {
                                        CurrentScreenMap = previousScreenMap;
                                    }
                                    if (verify?.Verdict?.Equals("yes", StringComparison.OrdinalIgnoreCase) == true)
                                    {
                                        if (independentlyVerified)
                                        {
                                            Console.WriteLine($"[verify] ✅ Goal independently confirmed: {verify.Reason}");
                                            AddHistory(historyBuffer, $"[{step}] done_verified");
                                        }
                                        else
                                        {
                                            var confidenceText = action.Confidence is double doneConfidenceForLog
                                                ? doneConfidenceForLog.ToString("0.00")
                                                : "n/a";
                                            Console.WriteLine(
                                                $"[verify] ⏭ Independent verification skipped; " +
                                                $"{VerificationSkipReason(step, action)}; accepting model done (confidence={confidenceText}).");
                                            AddHistory(historyBuffer, $"[{step}] done_accepted_unverified");
                                        }
                                        lastVerifierRejection = null;
                                        lastAimRect = null;
                                        recoveryEpisode = ConfirmPendingRecovery(recoveryEpisode, recoveryLessons, independentlyVerified);
                                        Console.WriteLine("Finished (model returned 'done').");
                                        runResult = new ControlRunResult(
                                            ControlRunOutcome.Completed,
                                            step,
                                            independentlyVerified
                                                ? verify?.Reason ?? "the model completion was independently verified"
                                                : "model completion was accepted without independent verification");
                                        break;
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[verify] ❌ Goal NOT confirmed. Reason: {verify?.Reason ?? "n/a"}");
                                        lastVerifierRejection = verify?.Reason ?? "verifier rejected done without a reason";
                                        AddHistory(historyBuffer, $"[{step}] done_rejected: {verify?.Reason}");
                                        actionExecutionFailed = true;
                                    }
                                }
                            }
                            else if (action.Type is "drag_drop" or "drag_path")
                            {
                                if (action.Type == "drag_drop" &&
                                    (!HasExplicitPoint(action) || !HasExplicitDropPoint(action)))
                                    throw new InvalidOperationException("drag_drop requires an explicit source and destination.");
                                if (action.Type == "drag_path" &&
                                    (action.Path is null || action.Path.Length < 2))
                                    throw new InvalidOperationException("drag_path requires at least two path points.");

                                var source = action.Type == "drag_path"
                                    ? ResolveGesturePath(action)[0]
                                    : new Point(ResolvePoint(action).X, ResolvePoint(action).Y);
                                if (lastAimRect is null && !DirectClickWithoutAim)
                                {
                                    Console.WriteLine($"[guard] {action.Type} blocked: no active source AIM. Return 'aim' for the source first.");
                                    AddHistory(historyBuffer, $"[{step}] IGNORED (gesture_without_aim)");
                                    lastExecutorFailure = $"{action.Type} was blocked because its source had no active AIM";
                                    actionExecutionFailed = true;
                                    actionWasLocallyRejected = true;
                                }
                                else if (lastAimRect is Rectangle dragAim && !dragAim.Contains(source.X, source.Y))
                                {
                                    Console.WriteLine($"[guard] {action.Type} source outside active AIM → ignoring. Set AIM around the gesture start.");
                                    AddHistory(historyBuffer, $"[{step}] IGNORED (gesture_source_outside_aim)");
                                    lastExecutorFailure = $"{action.Type} was blocked because its source was outside the active AIM";
                                    actionExecutionFailed = true;
                                    actionWasLocallyRejected = true;
                                }
                                else
                                {
                                    ExecuteAction(action);
                                    lastAimRect = null;
                                    actionExecuted = true;
                                }
                            }
                            else if (string.Equals(
                                         currentActionPolicy,
                                         "turn_based_interaction",
                                         StringComparison.Ordinal) &&
                                     actionBaselineFrame is not null &&
                                     currentAction.ObservationRegion is Rectangle observedTurnRegion &&
                                     TryTakeQueuedObservedTurnActions(
                                         currentAction,
                                         observedTurnRegion,
                                         out var observedTurnActions))
                            {
                                var observedBatchResult =
                                    await ExecuteObservedTurnActionSequenceAsync(
                                        observedTurnActions,
                                        actionBaselineFrame,
                                        observedTurnRegion,
                                        turnBasedTransitions,
                                        turnBasedInteractionWindow,
                                        cancelCts.Token);
                                currentAction = observedBatchResult.ExecutedActions[^1];
                                action = currentAction.Action;
                                action.TurnSequenceObserved = true;
                                turnBatchObserved = true;
                                actionExecuted = true;
                            }
                            else if (action.Type is "click" or "double_click")
                            {
                                var hasExplicitClickPoint = HasExplicitPoint(action);
                                var explicitClickOutsideAim = false;
                                if (hasExplicitClickPoint && lastAimRect is Rectangle activeAim)
                                {
                                    var explicitPoint = ResolveClickPoint(
                                        action,
                                        activeAim,
                                        logAdjustment: false);
                                    explicitClickOutsideAim = !activeAim.Contains(
                                        explicitPoint.X,
                                        explicitPoint.Y);
                                }

                                if (DirectClickWithoutAim &&
                                    hasExplicitClickPoint &&
                                    (lastAimRect is null || explicitClickOutsideAim))
                                {
                                    if (explicitClickOutsideAim)
                                    {
                                        Console.WriteLine(
                                            "[aim] superseded by a new explicit click target after the model revised its intended interaction.");
                                        lastAimRect = null;
                                    }
                                    Console.WriteLine("[guard] direct explicit click allowed by profile.");
                                    ExecuteAction(action);
                                    actionExecuted = true;
                                }
                                else if (lastAimRect is null)
                                {
                                    Console.WriteLine("[guard] click blocked: no active AIM. Return 'aim' first.");
                                    AddHistory(historyBuffer, $"[{step}] IGNORED (click_without_aim)");
                                    lastExecutorFailure = "click was blocked because there was no active AIM";
                                    actionExecutionFailed = true;
                                    actionWasLocallyRejected = true;
                                }
                                else
                                {
                                    // Clicks must include explicit coordinates (don't default to AIM center)
                                    if (!HasExplicitPoint(action))
                                    {
                                        Console.WriteLine("[guard] click/double_click requires explicit coordinates (x/y or x_px/y_px) – provide an exact point within AIM.");
                                        AddHistory(historyBuffer, $"[{step}] IGNORED (click_missing_coords)");
                                        lastExecutorFailure = "click was blocked because explicit coordinates were missing";
                                        actionExecutionFailed = true;
                                        actionWasLocallyRejected = true;
                                    }
                                    else
                                    {
                                        var (xClick, yClick) = ResolveClickPoint(action, lastAimRect, logAdjustment: false);

                                        if (!lastAimRect.Value.Contains(xClick, yClick))
                                        {
                                            Console.WriteLine("[guard] click outside active AIM → ignoring. Set a proper 'aim' first.");
                                            AddHistory(historyBuffer, $"[{step}] IGNORED (click_outside_aim)");
                                            lastExecutorFailure = "click was blocked because its point was outside the active AIM";
                                            actionExecutionFailed = true;
                                            actionWasLocallyRejected = true;
                                        }
                                        else
                                        {
                                            ExecuteAction(action, lastAimRect);
                                            actionExecuted = true;
                                        }
                                    }
                                }
                            }
                            else if (string.Equals(
                                         currentActionPolicy,
                                         "turn_based_interaction",
                                         StringComparison.Ordinal) &&
                                     TryGetTurnBasedDirectionalSequenceLength(
                                         action,
                                         out var turnSequenceLength) &&
                                     turnSequenceLength > 1)
                            {
                                if (actionBaselineFrame is null ||
                                    currentAction.ObservationRegion is not Rectangle turnBatchRegion)
                                {
                                    throw new InvalidOperationException(
                                        "turn-based sequence requires a current interaction baseline");
                                }
                                var turnBatchResult = await ExecuteTurnBasedDirectionalSequenceAsync(
                                    action.Keys!,
                                    actionBaselineFrame,
                                    turnBatchRegion,
                                    turnBasedTransitions,
                                    turnBasedInteractionWindow,
                                    cancelCts.Token);
                                action.Keys = turnBatchResult.ExecutedKeys.ToArray();
                                action.TurnSequenceObserved = true;
                                currentAction = AttachTurnBasedObservationRegion(
                                    CaptureResolvedAction(action, lastAimRect),
                                    turnBasedInteractionRect,
                                    currentActionPolicy);
                                turnBatchObserved = true;
                                actionExecuted = true;
                            }
                            else
                            {
                                // Move/scroll/keys/type_text – execute normally
                                ExecuteAction(action);
                                actionExecuted = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex is OperationCanceledException && CancelRequested)
                            {
                                Console.WriteLine("Aborted (hotkey).");
                                runResult = new ControlRunResult(
                                    ControlRunOutcome.Cancelled,
                                    step,
                                    "cancelled with the emergency hotkey");
                                break;
                            }
                            Console.WriteLine($"Action execution error: {ex.Message}");
                            consecutiveActionFailures++;
                            actionExecutionFailed = true;
                            lastExecutorFailure = $"{currentAction.Description} failed: {ex.Message}";
                            AddHistory(historyBuffer, $"[{step}] EXECUTOR_FAILURE: {lastExecutorFailure}");
                            PendingSafeActions.Clear();
                            if (MaxActionFailuresBeforeAbort > 0 && consecutiveActionFailures >= MaxActionFailuresBeforeAbort)
                            {
                                Console.WriteLine($"[guard] stopping: local action execution failed {consecutiveActionFailures} time(s). Use --max-action-failures 0 to disable.");
                                runResult = new ControlRunResult(
                                    ControlRunOutcome.GuardStopped,
                                    step,
                                    $"local action execution failed {consecutiveActionFailures} times");
                                break;
                            }
                        }

                        if (actionWasLocallyRejected)
                        {
                            RegisterRejectedProposal(currentAction);
                            consecutiveActionFailures++;
                            if (MaxActionFailuresBeforeAbort > 0 &&
                                consecutiveActionFailures >= MaxActionFailuresBeforeAbort)
                            {
                                Console.WriteLine($"[guard] stopping: local action was rejected {consecutiveActionFailures} time(s). Use --max-action-failures 0 to disable.");
                                runResult = new ControlRunResult(
                                    ControlRunOutcome.GuardStopped,
                                    step,
                                    $"local action was rejected {consecutiveActionFailures} times");
                                break;
                            }
                        }
        
                        if (!actionExecutionFailed && actionExecuted)
                        {
                            if (string.Equals(
                                    currentActionPolicy,
                                    "turn_based_interaction",
                                    StringComparison.Ordinal) &&
                                IsStateChangingInteractionAction(currentAction.Action) &&
                                turnReanalysisWasRequired)
                            {
                                turnBasedTransitions.AcknowledgeReanalysis();
                                if (!turnBasedTransitions.HasActiveCausalEvent)
                                    activeTurnChangeImages = [];
                            }
                            consecutiveActionFailures = 0;
                            lastExecutorFailure = null;
                            AddHistory(historyBuffer, $"[{step}] {currentAction.Description}");
                            if (!string.IsNullOrWhiteSpace(action.Note))
                                AddHistory(historyBuffer, $"[{step}] note: {action.Note}");
                            if (action.Type != "done" && lastVerifierRejection != null)
                                lastVerifierRejection = null;

                            observationActionGuard.RecordExecuted(currentAction);
                            RecordRecoveryAction(recoveryEpisode, recentExecutedActions, currentAction, recoveryLessons);

                            var batchResult = await ExecuteQueuedSafeActionsAsync(
                                historyBuffer,
                                step,
                                action,
                                (actionBaselineFrame ?? observationFrame).GlobalFingerprint,
                                cancelCts.Token);
                            foreach (var batchedAction in batchResult.ExecutedActions)
                            {
                                observationActionGuard.RecordExecuted(batchedAction);
                                RecordRecoveryAction(recoveryEpisode, recentExecutedActions, batchedAction, recoveryLessons);
                                currentAction = batchedAction;
                            }
                            if (batchResult.ExecutedActions.Count > 0)
                            {
                                currentActionPolicy = observationSession.ResolveActionPolicy(
                                    currentAction,
                                    promptContext,
                                    goal);
                            }
                            if (!string.IsNullOrWhiteSpace(batchResult.Error))
                                lastExecutorFailure = batchResult.Error;
                        }
        
                        // Keep context for next step (delta/repeat metrics)
                        prevAction = actionExecutionFailed || !actionExecuted ? null : currentAction;
                        prevShotFingerprint = actionExecutionFailed || !actionExecuted
                            ? null
                            : (actionBaselineFrame ?? observationFrame).GlobalFingerprint;
                        prevActiveWindowFingerprint = actionExecutionFailed || !actionExecuted
                            ? null
                            : (actionBaselineFrame ?? observationFrame).ActiveWindowFingerprint;
                        prevObservationFrame = actionExecutionFailed || !actionExecuted
                            ? null
                            : actionBaselineFrame ?? observationFrame;
                        prevObservationContext = actionExecutionFailed || !actionExecuted
                            ? null
                            : promptContext;
        
                        if (CancelRequested)
                        {
                            Console.WriteLine("Aborted (hotkey).");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.Cancelled,
                                step,
                                "cancelled with the emergency hotkey");
                            break;
                        }
        
                        if (!actionExecutionFailed && actionExecuted && !turnBatchObserved)
                        {
                            try
                            {
                                _ = await WaitAfterActionAsync(
                                    currentAction.Action,
                                    (actionBaselineFrame ?? observationFrame).GlobalFingerprint,
                                    currentActionPolicy,
                                    actionBaselineLocalFingerprint,
                                    currentAction.ObservationRegion,
                                    cancelCts.Token);
                            }
                            catch (OperationCanceledException) when (CancelRequested)
                            {
                                Console.WriteLine("Aborted (hotkey).");
                                runResult = new ControlRunResult(
                                    ControlRunOutcome.Cancelled,
                                    step,
                                    "cancelled with the emergency hotkey");
                                break;
                            }
                        }
                    }
                }
                catch (OperationCanceledException) when (CancelRequested)
                {
                    Console.WriteLine("Aborted (hotkey).");
                    runResult = new ControlRunResult(
                        ControlRunOutcome.Cancelled,
                        currentStep,
                        "cancelled with the emergency hotkey");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(
                        $"Control run failed unexpectedly: {ex}");
                    runResult = new ControlRunResult(
                        ControlRunOutcome.Failed,
                        currentStep,
                        ex.Message);
                }
                finally
                {
                    ReleaseAllHeldKeys();
                    cancelCts?.Cancel();
                    cancelCts?.Dispose();
                    _ = FlushPendingRecoveryMemory();
                    if (LoopReplayAutoExportEnabled && RecoveryMemoryEnabled)
                        TryAutoExportLoopReplayCorpus();
                    controlContextChain?.LogClose(runResult.Outcome);
                    PrintRunMetrics();
                    Console.SetOut(prevOut);
                    Console.SetError(prevErr);
                    if (consoleHidden && RestoreConsoleAfterRun)
                        RestoreConsoleWindow();
                }
                return runResult with { Step = currentStep };
            }

            internal sealed class ControlContextChain
            {
                readonly string chainId;
                readonly int fallbackLimit;
                int fallbackCount;
                int compactionCount;

                internal ControlContextChain(
                    string commandId,
                    bool enabled,
                    bool compactionEnabled,
                    int fallbackLimit)
                {
                    chainId = commandId.Length <= 8 ? commandId : commandId[..8];
                    Enabled = enabled;
                    CompactionEnabled = enabled && compactionEnabled;
                    this.fallbackLimit = Math.Max(1, fallbackLimit);
                }

                internal bool Enabled { get; private set; }
                internal bool CompactionEnabled { get; private set; }
                internal string? PreviousResponseId { get; private set; }
                internal string? PreviousResponseIdForRequest =>
                    Enabled ? PreviousResponseId : null;
                internal int TurnCount { get; private set; }
                internal int RestartCount { get; private set; }

                internal void LogStart(string model)
                {
                    if (!Enabled)
                    {
                        Console.WriteLine("[context] control chaining disabled; explicit application history will be used.");
                        return;
                    }

                    var effectiveContext = SupportsReasoningContext(model)
                        ? ControlReasoningContext
                        : "model_default";
                    var compaction = CompactionEnabled
                        ? ControlContextCompactThreshold.ToString(CultureInfo.InvariantCulture)
                        : "off";
                    Console.WriteLine(
                        $"[context] started control chain id={chainId}; reasoning_context={effectiveContext}; compact_threshold={compaction}.");
                }

                internal void LogRequest(int step)
                {
                    if (!Enabled)
                        return;

                    if (PreviousResponseId is null)
                    {
                        Console.WriteLine($"[context] chain={chainId}; step={step}; starting API turn without previous_response_id.");
                        return;
                    }

                    Console.WriteLine(
                        $"[context] chain={chainId}; step={step}; continuing turn={TurnCount + 1}; previous={ShortResponseId(PreviousResponseId)}.");
                }

                internal void RecordResult(
                    string? completedResponseId,
                    bool contextFallbackUsed,
                    bool compactionFallbackUsed,
                    bool compactionOccurred)
                {
                    if (compactionFallbackUsed && CompactionEnabled)
                    {
                        CompactionEnabled = false;
                        Console.WriteLine(
                            $"[context] chain={chainId}; server-side compaction disabled for the remainder of this task.");
                    }

                    if (contextFallbackUsed)
                    {
                        PreviousResponseId = null;
                        fallbackCount++;
                        RestartCount++;
                        RunControlContextRestarts++;
                        Console.WriteLine(
                            $"[context] chain={chainId}; restarted from checkpoint; fallback={fallbackCount}/{fallbackLimit}.");
                        if (fallbackCount >= fallbackLimit)
                        {
                            Enabled = false;
                            CompactionEnabled = false;
                            Console.WriteLine(
                                $"[context] chain={chainId}; disabled after {fallbackCount} state failures; continuing with explicit application history.");
                        }
                    }

                    if (compactionOccurred)
                    {
                        compactionCount++;
                        Console.WriteLine(
                            $"[context] chain={chainId}; server-side compaction completed; count={compactionCount}.");
                    }

                    if (!Enabled || string.IsNullOrWhiteSpace(completedResponseId))
                        return;

                    PreviousResponseId = completedResponseId;
                    TurnCount++;
                    RunControlContextTurns++;
                    Console.WriteLine(
                        $"[context] chain={chainId}; finalized turn={TurnCount}; response={ShortResponseId(completedResponseId)}.");
                }

                internal void LogClose(ControlRunOutcome outcome)
                {
                    if (!Enabled && TurnCount == 0 && RestartCount == 0)
                        return;

                    Console.WriteLine(
                        $"[context] closed chain id={chainId}; outcome={outcome}; turns={TurnCount}; restarts={RestartCount}; compactions={compactionCount}.");
                }

                static string ShortResponseId(string responseId) =>
                    responseId.Length <= 18 ? responseId : responseId[..18] + "...";
            }
        
            internal static bool IsMouseAction(ActionDto a)
                => a.Type is "move" or "click" or "double_click" or "drag_drop" or "drag_path" or "scroll" or "focus_uia" or "click_uia";
        
            internal static bool IsPointClickAction(ActionDto? a)
                => a?.Type is "click" or "double_click";
        
            internal static bool IsTextInputAttemptAction(ActionDto? a)
            {
                if (a?.Type is "paste_text" or "type_text")
                    return true;
        
                if (a?.Type != "keys" || a.Keys is null || a.Keys.Length == 0)
                    return false;
        
                var compact = string.Join("+", a.Keys.Select(k => (k ?? "").Trim().ToLowerInvariant().Replace("control", "ctrl")));
                compact = compact.Replace(" ", "");
                return compact.Contains("ctrl+v", StringComparison.Ordinal) ||
                       compact.Contains("shift+insert", StringComparison.Ordinal);
            }
        
            internal static bool IsKnownAction(ActionDto? action) =>
                !string.IsNullOrWhiteSpace(action?.Type) && KnownActionTypes.Contains(action.Type);
        
            internal static string BuildPrecisionHint(ResolvedActionSnapshot action, double delta, string reason)
            {
                var target = "";
                if (action.ScreenPoint is Point screenPoint)
                {
                    var imagePoint = CurrentScreenMap.ScreenToImagePoint(screenPoint.X, screenPoint.Y);
                    target = $" Target was near SCREEN_SIZE ({imagePoint.X},{imagePoint.Y}).";
                }

                var deltaText = double.IsNaN(delta) ? "" : $" Screen delta={delta:0.####}.";
                return $"{reason}{target}{deltaText} Do not repeat the same point. For tiny controls, tree expanders, list rows, or menus, prefer request_crop/aim for that area or select the row and use keyboard expansion/activation such as ArrowRight or Enter.";
            }
        
            internal static string BuildTextInputHint(ActionDto action, double delta)
            {
                var deltaText = double.IsNaN(delta) ? "" : $" Screen delta={delta:0.####}.";
                var textLen = action.Text?.Length ?? 0;
                var textInfo = textLen > 0 ? $" Attempted {action.Type} with {textLen} character(s)." : $" Attempted {action.Type}.";
                return $"Previous text input produced little or no visible progress.{deltaText}{textInfo} The target is likely not focused, not editable, or the caret is elsewhere. Do not repeat text input or paste shortcuts immediately; first make the editable field active through visible UI, use request_crop/aim if needed, or switch to a different visible UI route.";
            }
        
            internal static string BuildTextInputCooldownHint(int attempts)
            {
                var countText = attempts > 0 ? $" after {attempts} no-change attempt(s)" : "";
                return $"Text input is temporarily blocked{countText} because it has not changed the screen. Do not choose paste_text, type_text, Ctrl+V, or Shift+Insert now. Use visible UI to establish focus/editability, open the correct editor/control, or choose a non-text navigation route first.";
            }
        
            internal static void AddHistory(StringBuilder historyBuffer, string line)
            {
                if (HistoryTailChars <= 0)
                    return;
        
                if (string.IsNullOrWhiteSpace(line))
                    return;
        
                if (historyBuffer.Length > 0)
                    historyBuffer.AppendLine();
                historyBuffer.Append(line);
        
                var maxBufferChars = Math.Max(HistoryTailChars * 2, HistoryTailChars + 2048);
                if (historyBuffer.Length > maxBufferChars)
                    historyBuffer.Remove(0, historyBuffer.Length - maxBufferChars);
            }
        
            internal static string TailHistory(StringBuilder historyBuffer, int maxChars, int maxLines)
            {
                if (maxChars <= 0 || historyBuffer.Length == 0)
                    return "";
        
                var tail = Tail(historyBuffer, maxChars);
                var firstNewLine = tail.IndexOf('\n');
                if (historyBuffer.Length > tail.Length && firstNewLine >= 0 && firstNewLine + 1 < tail.Length)
                    tail = tail[(firstNewLine + 1)..];
        
                if (maxLines <= 0)
                    return tail.Trim();
        
                var lines = tail.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length <= maxLines)
                    return string.Join('\n', lines).Trim();
        
                return string.Join('\n', lines.Skip(lines.Length - maxLines)).Trim();
            }

            internal static string BuildRecentActionCheckpoint(
                IReadOnlyCollection<ResolvedActionSnapshot> recentActions)
            {
                if (recentActions.Count == 0)
                    return "";

                var actions = recentActions
                    .TakeLast(12)
                    .Select(action => TrimForMeta(action.Description, 160));
                return "RECENT_EXECUTED_ACTIONS (oldest to newest): " +
                       string.Join(" | ", actions);
            }
        
            internal static bool IsLocalObservationAction(ActionDto a)
                => a.Type is "aim" or "point" or "request_crop";

            internal static bool IsInspectionAction(ActionDto a)
                => a.Type is "point" or "request_crop";

            internal static bool IsStateChangingInteractionAction(ActionDto a) =>
                ActionNeedsEffectObservation(a) && a.Type != "wait";

            internal static bool RequiresPrimaryTurnBasedRegion(ActionDto action)
            {
                if (!IsStateChangingInteractionAction(action) ||
                    action.Type is "click" or "double_click")
                {
                    return false;
                }

                if (action.Type == "keys" && action.Keys is { Length: 1 })
                {
                    var key = action.Keys[0]?.Trim().ToLowerInvariant();
                    if (key is "space" or "spacebar" or "enter" or "return")
                        return false;
                }

                return true;
            }

            internal static bool TryGetTurnBasedDirectionalSequenceLength(
                ActionDto action,
                out int length)
            {
                length = action.Keys?.Length ?? 0;
                if (action.Type != "keys" || length == 0)
                    return false;

                return action.Keys!.All(key =>
                {
                    var normalized = key?.Trim().ToLowerInvariant();
                    return normalized is "arrowleft" or "left" or "arrowright" or "right" or
                        "arrowup" or "up" or "arrowdown" or "down" or
                        "w" or "a" or "s" or "d";
                });
            }

            internal readonly record struct TurnBatchExecutionResult(
                IReadOnlyList<string> ExecutedKeys,
                bool Interrupted);

            internal readonly record struct ObservedTurnBatchExecutionResult(
                IReadOnlyList<ResolvedActionSnapshot> ExecutedActions,
                bool Interrupted);

            internal readonly record struct TurnNoEffectPolicy(
                bool ExtendObservation,
                bool ReplayInput);

            internal static TurnNoEffectPolicy ResolveTurnNoEffectPolicy(
                bool immediateReactionObserved) =>
                immediateReactionObserved
                    ? new TurnNoEffectPolicy(false, false)
                    : new TurnNoEffectPolicy(true, false);

            internal static bool CanExpandTurnKeyActionFromPlan(
                ActionDto action,
                int queuedFollowUpCount) =>
                queuedFollowUpCount == 0 &&
                action.Type == "keys" &&
                action.Keys is { Length: > 0 };

            internal static bool TryTakeQueuedObservedTurnActions(
                ResolvedActionSnapshot firstAction,
                Rectangle region,
                out IReadOnlyList<ResolvedActionSnapshot> actions)
            {
                actions = Array.Empty<ResolvedActionSnapshot>();
                if (PendingSafeActions.Count == 0 ||
                    firstAction.Action.PlannedInputs is not { Length: >= 2 } plannedInputs ||
                    (firstAction.Action.PlanConfidence ?? firstAction.Action.Confidence ?? 0) <
                        TurnBasedTransitionTracker.MinimumStructuredPlanConfidence ||
                    !OpenAiResponsesService.TryBindObservedTurnActionToPlannedInput(
                        firstAction.Action,
                        plannedInputs[0]))
                {
                    return false;
                }

                var queued = PendingSafeActions.ToArray();
                var maximumActions = Math.Min(
                    TurnBasedMaxBatchInputs,
                    plannedInputs.Length);
                var accepted = new List<ResolvedActionSnapshot>(maximumActions)
                {
                    firstAction with { ObservationRegion = ClampRect(region) }
                };
                for (var index = 0;
                     index < queued.Length && accepted.Count < maximumActions;
                     index++)
                {
                    var candidate = queued[index];
                    if (!OpenAiResponsesService.TryBindObservedTurnActionToPlannedInput(
                            candidate,
                            plannedInputs[index + 1]))
                    {
                        break;
                    }

                    var snapshot = CaptureResolvedAction(candidate, null) with
                    {
                        ObservationRegion = ClampRect(region)
                    };
                    if (!snapshot.IsValid)
                        break;
                    accepted.Add(snapshot);
                }

                if (accepted.Count <= 1)
                    return false;

                PendingSafeActions.Clear();
                actions = accepted;
                return true;
            }

            internal static async Task<ObservedTurnBatchExecutionResult> ExecuteObservedTurnActionSequenceAsync(
                IReadOnlyList<ResolvedActionSnapshot> actions,
                ScreenObservationFrame initialFrame,
                Rectangle region,
                TurnBasedTransitionTracker tracker,
                IntPtr interactionWindow,
                CancellationToken cancellationToken)
            {
                Console.WriteLine(
                    $"[turn-batch] mode={tracker.ExecutionBatchMode}; modality=mixed_observed; executing={actions.Count}; observation_barrier=per_input");
                var executed = new List<ResolvedActionSnapshot>(actions.Count);
                var beforeFrame = initialFrame;
                tracker.BeginBatch(initialFrame, region);
                for (var index = 0; index < actions.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (interactionWindow != IntPtr.Zero &&
                        GetForegroundWindow() != interactionWindow &&
                        !TryRestoreForegroundWindow(interactionWindow))
                    {
                        Console.WriteLine(
                            "[turn-focus] observed sequence interrupted because the interaction window could not regain focus.");
                        return new ObservedTurnBatchExecutionResult(executed, true);
                    }
                    var snapshot = actions[index];
                    var beforeLocal = CaptureRegionFingerprintProbe(region);
                    ExecuteAction(snapshot.Action);
                    var actionReactionObserved = await WaitAfterActionAsync(
                        snapshot.Action,
                        beforeFrame.GlobalFingerprint,
                        "turn_based_interaction",
                        beforeLocal,
                        region,
                        cancellationToken,
                        tracker.PreferFastBatchSettle);
                    var noEffectPolicy = ResolveTurnNoEffectPolicy(
                        actionReactionObserved);
                    if (noEffectPolicy.ExtendObservation)
                    {
                        Console.WriteLine(
                            $"[turn-batch] input={index + 1}/{actions.Count}; no immediate reaction; extending observation without replaying the input.");
                        actionReactionObserved = await WaitForLocalScreenReactionAndStableAsync(
                            beforeLocal,
                            region,
                            snapshot.Action,
                            cancellationToken,
                            preferFastSettle: false);
                    }
                    var afterFrame = CaptureObservationFrameProbe(includeDetail: true);
                    var assessment = tracker.RecordBatchActionStep(
                        beforeFrame,
                        afterFrame,
                        region,
                        snapshot,
                        actionReactionObserved);
                    beforeFrame = afterFrame;
                    executed.Add(snapshot);
                    Console.WriteLine(
                        $"[turn-batch] input={index + 1}/{actions.Count}; attempts=1; action={snapshot.Description}; known_transition={assessment.KnownTransition.ToString().ToLowerInvariant()}; salient={assessment.SalientChange.ToString().ToLowerInvariant()}; {assessment.Summary}");
                    if (assessment.NoEffect)
                    {
                        Console.WriteLine(
                            "[turn-batch] observed sequence interrupted after extended no_effect confirmation; the input was not replayed and the route cursor was not advanced.");
                        return new ObservedTurnBatchExecutionResult(executed, true);
                    }
                    if (index + 1 < actions.Count && !assessment.ContinueBatch)
                    {
                        Console.WriteLine(
                            "[turn-batch] observed sequence interrupted before the next input; a fresh model decision is required.");
                        return new ObservedTurnBatchExecutionResult(executed, true);
                    }
                }

                return new ObservedTurnBatchExecutionResult(executed, false);
            }

            internal static async Task<TurnBatchExecutionResult> ExecuteTurnBasedDirectionalSequenceAsync(
                IReadOnlyList<string> keys,
                ScreenObservationFrame initialFrame,
                Rectangle region,
                TurnBasedTransitionTracker tracker,
                IntPtr interactionWindow,
                CancellationToken cancellationToken)
            {
                Console.WriteLine(
                    $"[turn-batch] mode={tracker.ExecutionBatchMode}; advertised_cap={tracker.AdvertisedMaxExecutionBatchLength}; accepted_cap={tracker.MaxExecutionBatchLength}; executing={keys.Count}; observation_barrier=per_input");
                var executed = new List<string>(keys.Count);
                var beforeFrame = initialFrame;
                tracker.BeginBatch(initialFrame, region);
                for (var index = 0; index < keys.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (interactionWindow != IntPtr.Zero &&
                        GetForegroundWindow() != interactionWindow &&
                        !TryRestoreForegroundWindow(interactionWindow))
                    {
                        Console.WriteLine(
                            "[turn-focus] key sequence interrupted because the interaction window could not regain focus.");
                        return new TurnBatchExecutionResult(executed, true);
                    }
                    var key = keys[index];
                    var stepAction = new ActionDto
                    {
                        Type = "keys",
                        Keys = [key]
                    };
                    var beforeLocal = CaptureRegionFingerprintProbe(region);
                    PressKey(key);
                    var actionReactionObserved = await WaitAfterActionAsync(
                        stepAction,
                        beforeFrame.GlobalFingerprint,
                        "turn_based_interaction",
                        beforeLocal,
                        region,
                        cancellationToken,
                        tracker.PreferFastBatchSettle);
                    var noEffectPolicy = ResolveTurnNoEffectPolicy(
                        actionReactionObserved);
                    if (noEffectPolicy.ExtendObservation)
                    {
                        Console.WriteLine(
                            $"[turn-batch] input={index + 1}/{keys.Count}; no immediate reaction; extending observation without replaying the key.");
                        actionReactionObserved = await WaitForLocalScreenReactionAndStableAsync(
                            beforeLocal,
                            region,
                            stepAction,
                            cancellationToken,
                            preferFastSettle: false);
                    }
                    var afterFrame = CaptureObservationFrameProbe(includeDetail: true);
                    var assessment = tracker.RecordBatchStep(
                        beforeFrame,
                        afterFrame,
                        region,
                        key,
                        actionReactionObserved);
                    beforeFrame = afterFrame;
                    executed.Add(key);
                    Console.WriteLine(
                        $"[turn-batch] input={index + 1}/{keys.Count}; attempts=1; known_transition={assessment.KnownTransition.ToString().ToLowerInvariant()}; salient={assessment.SalientChange.ToString().ToLowerInvariant()}; {assessment.Summary}");
                    if (assessment.NoEffect)
                    {
                        Console.WriteLine(
                            "[turn-batch] sequence interrupted after extended no_effect confirmation; the key was not replayed and the route cursor was not advanced.");
                        return new TurnBatchExecutionResult(executed, true);
                    }
                    if (index + 1 < keys.Count && !assessment.ContinueBatch)
                    {
                        Console.WriteLine(
                            "[turn-batch] sequence interrupted before the next input; a fresh model decision is required.");
                        return new TurnBatchExecutionResult(executed, true);
                    }
                }
                return new TurnBatchExecutionResult(executed, false);
            }

            internal sealed class ObservationActionGuardState
            {
                readonly HashSet<string> inspectionSignatures = new(StringComparer.Ordinal);

                internal int ConsecutiveInspectionActions { get; private set; }
                internal bool AimIssuedSinceInteraction { get; private set; }

                internal bool RequiresInteraction(int maxConsecutiveInspectionActions) =>
                    maxConsecutiveInspectionActions > 0 &&
                    ConsecutiveInspectionActions >= maxConsecutiveInspectionActions;

                internal bool TryGetBlockReason(
                    ResolvedActionSnapshot action,
                    int maxConsecutiveInspectionActions,
                    out string reason) =>
                    TryGetBlockReason(
                        action,
                        maxConsecutiveInspectionActions,
                        singleInspectionBeforeInteraction: false,
                        out reason);

                internal bool TryGetBlockReason(
                    ResolvedActionSnapshot action,
                    int maxConsecutiveInspectionActions,
                    bool singleInspectionBeforeInteraction,
                    out string reason)
                {
                    reason = "";
                    if (IsInspectionAction(action.Action))
                    {
                        if (AimIssuedSinceInteraction)
                        {
                            reason = "an AIM is already active; interact with that target instead of requesting another inspection";
                            return true;
                        }

                        var signature = InspectionSignature(action.Action);
                        if (inspectionSignatures.Contains(signature))
                        {
                            reason = "the requested inspection revisits an area already inspected since the last interaction; choose a state-changing action";
                            return true;
                        }

                        if (singleInspectionBeforeInteraction &&
                            ConsecutiveInspectionActions >= 1)
                        {
                            reason = "a detailed turn-based inspection was already supplied; perform a state-changing input before requesting another nested crop";
                            return true;
                        }

                        if (RequiresInteraction(maxConsecutiveInspectionActions))
                        {
                            reason = $"the limit of {maxConsecutiveInspectionActions} consecutive crop/point inspections was reached; perform one safe state-changing interaction before inspecting again";
                            return true;
                        }
                    }
                    else if (action.Action.Type == "aim" && AimIssuedSinceInteraction)
                    {
                        reason = "an AIM is already active; click or perform the intended gesture instead of aiming again";
                        return true;
                    }

                    return false;
                }

                internal void RecordExecuted(ResolvedActionSnapshot action)
                {
                    if (IsInspectionAction(action.Action))
                    {
                        inspectionSignatures.Add(InspectionSignature(action.Action));
                        ConsecutiveInspectionActions++;
                        return;
                    }

                    if (action.Action.Type == "aim")
                    {
                        AimIssuedSinceInteraction = true;
                        return;
                    }

                    if (IsStateChangingInteractionAction(action.Action))
                        Reset();
                }

                internal void Reset()
                {
                    inspectionSignatures.Clear();
                    ConsecutiveInspectionActions = 0;
                    AimIssuedSinceInteraction = false;
                }

                static string InspectionSignature(ActionDto action)
                {
                    var rect = ResolveCropRect(action);
                    if (rect is not Rectangle region)
                        return IneffectiveActionSignature(action);

                    const int quantum = 16;
                    return $"{action.Type}:{region.Left / quantum},{region.Top / quantum},{region.Width / quantum},{region.Height / quantum}";
                }
            }

            internal sealed class ShortTermPlanTracker
            {
                string? plan;
                string status = "none";
                string? revisionReason;
                string[] plannedInputs = [];
                string? waypoint;
                string? originStateId;
                string? expectedStateId;
                double? planConfidence;
                int currentInputIndex;

                internal bool HasActivePlan =>
                    plan is not null && status == "active";

                internal bool HasExecutableDirectionalPlan =>
                    HasActivePlan && currentInputIndex < plannedInputs.Length;

                internal string Status => status;

                internal int RemainingInputCount =>
                    Math.Max(0, plannedInputs.Length - currentInputIndex);

                internal double EffectiveConfidence => planConfidence ?? 0;

                internal void Update(ActionDto action, string? currentStateId = null)
                {
                    var requestedStatus = action.PlanStatus?.Trim().ToLowerInvariant();
                    if (string.IsNullOrWhiteSpace(requestedStatus))
                        return;

                    if (requestedStatus == "none")
                    {
                        if (plan is not null || status != "none")
                            Console.WriteLine("[plan] cleared.");
                        plan = null;
                        status = "none";
                        revisionReason = null;
                        plannedInputs = [];
                        waypoint = null;
                        originStateId = null;
                        expectedStateId = null;
                        planConfidence = null;
                        currentInputIndex = 0;
                        return;
                    }

                    var nextPlan = NormalizePlanText(action.ShortTermPlan, 420);
                    if (nextPlan is null)
                    {
                        Invalidate("the model marked a plan active without supplying its steps");
                        return;
                    }

                    var nextReason = NormalizePlanText(action.PlanRevisionReason, 200);
                    var nextInputs = NormalizePlannedInputs(action.PlannedInputs);
                    if (OpenAiResponsesService.HasImmediateDirectionalReversal(nextInputs))
                    {
                        Invalidate(
                            "the proposed route contains an immediate opposing-direction reversal and is not a coherent path to one waypoint");
                        return;
                    }
                    var nextWaypoint = NormalizePlanText(action.PlanWaypoint, 160);
                    var revised = requestedStatus == "revised";
                    var hasStructuredIdentity =
                        plannedInputs.Length > 0 || nextInputs.Length > 0;
                    var changed = hasStructuredIdentity
                        ? !plannedInputs.SequenceEqual(
                              nextInputs,
                              StringComparer.Ordinal) ||
                          !string.Equals(
                              waypoint,
                              nextWaypoint,
                              StringComparison.Ordinal)
                        : !string.Equals(plan, nextPlan, StringComparison.Ordinal);
                    var preserveProgress = HasActivePlan && !revised && !changed;
                    plan = nextPlan;
                    status = "active";
                    revisionReason = revised ? nextReason : null;
                    planConfidence = action.PlanConfidence ?? action.Confidence ?? planConfidence;
                    if (!preserveProgress)
                    {
                        plannedInputs = nextInputs;
                        waypoint = nextWaypoint;
                        currentInputIndex = 0;
                        originStateId = currentStateId;
                        expectedStateId = currentStateId;
                        var requestedStateId = NormalizePlanText(action.PlanStateId, 32);
                        if (requestedStateId is not null &&
                            currentStateId is not null &&
                            !string.Equals(
                                requestedStateId,
                                currentStateId,
                                StringComparison.Ordinal))
                        {
                            Invalidate(
                                $"the proposed route was bound to stale state {requestedStateId}, current state is {currentStateId}");
                            return;
                        }
                    }
                    if (changed || revised)
                    {
                        Console.WriteLine(
                            revised
                                ? $"[plan] revised: {plan}; reason={revisionReason ?? "new evidence"}"
                                : $"[plan] active: {plan}");
                        if (plannedInputs.Length > 0)
                        {
                            Console.WriteLine(
                                $"[turn-plan] state={originStateId ?? "unknown"}; waypoint={waypoint ?? "unspecified"}; confidence={(planConfidence is double confidence ? confidence.ToString("0.00") : "n/a")}; inputs=[{string.Join(",", plannedInputs)}]");
                        }
                    }
                }

                internal bool ProposedSequenceMatches(
                    IReadOnlyList<string> proposedInputs,
                    string? currentStateId,
                    out string reason)
                {
                    reason = "";
                    if (!HasExecutableDirectionalPlan)
                        return true;
                    if (expectedStateId is not null &&
                        currentStateId is not null &&
                        !string.Equals(expectedStateId, currentStateId, StringComparison.Ordinal))
                    {
                        reason =
                            $"structured route expects state {expectedStateId}, current state is {currentStateId}";
                        return false;
                    }

                    var normalized = proposedInputs
                        .Select(NormalizePlannedInput)
                        .ToArray();
                    if (normalized.Any(input => input is null))
                    {
                        reason = "structured route contains a non-directional input";
                        return false;
                    }
                    var remaining = plannedInputs.Skip(currentInputIndex).ToArray();
                    if (normalized.Length > remaining.Length ||
                        !normalized.Select(input => input!).SequenceEqual(
                            remaining.Take(normalized.Length),
                            StringComparer.Ordinal))
                    {
                        reason =
                            $"proposed inputs do not match remaining structured route [{string.Join(",", remaining)}]";
                        return false;
                    }
                    return true;
                }

                internal bool TryExpandDirectionalSequence(
                    IReadOnlyList<string> proposedInputs,
                    string? currentStateId,
                    int maximumLength,
                    out string[] expandedInputs)
                {
                    expandedInputs = proposedInputs.ToArray();
                    if (!HasExecutableDirectionalPlan ||
                        EffectiveConfidence < TurnBasedTransitionTracker.MinimumStructuredPlanConfidence ||
                        maximumLength < 2 ||
                        !ProposedSequenceMatches(
                            proposedInputs,
                            currentStateId,
                            out _))
                    {
                        return false;
                    }

                    var remaining = plannedInputs
                        .Skip(currentInputIndex)
                        .Take(maximumLength)
                        .ToArray();
                    if (remaining.Length <= proposedInputs.Count)
                        return false;
                    expandedInputs = remaining;
                    return true;
                }

                internal bool IsExpectedInput(string input, string? currentStateId)
                {
                    if (!HasExecutableDirectionalPlan)
                        return false;
                    var normalized = NormalizePlannedInput(input);
                    return normalized is not null &&
                           string.Equals(
                               plannedInputs[currentInputIndex],
                               normalized,
                               StringComparison.Ordinal) &&
                           (expectedStateId is null || currentStateId is null ||
                            string.Equals(
                                expectedStateId,
                                currentStateId,
                                StringComparison.Ordinal));
                }

                internal void RecordDirectionalResult(
                    string input,
                    string beforeState,
                    string afterState,
                    bool changed)
                {
                    if (!HasExecutableDirectionalPlan)
                        return;
                    if (!IsExpectedInput(input, beforeState))
                    {
                        Invalidate("the executed directional input diverged from the structured route");
                        return;
                    }
                    if (!changed)
                    {
                        Invalidate("a structured route input was blocked or had no effect");
                        return;
                    }

                    currentInputIndex++;
                    expectedStateId = afterState;
                    if (currentInputIndex >= plannedInputs.Length)
                    {
                        status = "completed";
                        revisionReason = "the concrete route prefix reached its waypoint";
                        Console.WriteLine(
                            $"[turn-plan] completed {currentInputIndex}/{plannedInputs.Length} input(s); waypoint={waypoint ?? "unspecified"}.");
                        return;
                    }
                    Console.WriteLine(
                        $"[turn-plan] advanced to {currentInputIndex}/{plannedInputs.Length}; remaining=[{string.Join(",", plannedInputs.Skip(currentInputIndex))}]");
                }

                internal void Invalidate(string reason)
                {
                    if (!HasActivePlan)
                        return;
                    status = "invalidated";
                    revisionReason = NormalizePlanText(reason, 200);
                    Console.WriteLine(
                        $"[plan] invalidated: {revisionReason ?? "observed state contradicted the plan"}");
                }

                internal string BuildPromptSummary()
                {
                    if (plan is null || status == "none")
                        return "";
                    var builder = new StringBuilder()
                        .AppendLine($"SHORT_TERM_PLAN_STATUS: {status}")
                        .AppendLine($"SHORT_TERM_PLAN: {plan}");
                    if (plannedInputs.Length > 0)
                    {
                        builder
                            .AppendLine($"PLAN_STATE_ID: {originStateId ?? "unknown"}")
                            .AppendLine($"PLAN_EXPECTED_STATE_ID: {expectedStateId ?? "unknown"}")
                            .AppendLine($"PLAN_WAYPOINT: {waypoint ?? "unspecified"}")
                            .AppendLine($"PLAN_CONFIDENCE: {(planConfidence is double confidence ? confidence.ToString("0.00") : "N/A")}")
                            .AppendLine($"CURRENT_PLAN_INDEX: {currentInputIndex}/{plannedInputs.Length}")
                            .AppendLine($"PLANNED_INPUTS: [{string.Join(",", plannedInputs)}]")
                            .AppendLine($"REMAINING_PLANNED_INPUTS: [{string.Join(",", plannedInputs.Skip(currentInputIndex))}]");
                    }
                    if (!string.IsNullOrWhiteSpace(revisionReason))
                        builder.Append($"SHORT_TERM_PLAN_REVISION_REASON: {revisionReason}");
                    return builder.ToString();
                }

                static string[] NormalizePlannedInputs(string[]? values) =>
                    values?
                        .Select(NormalizePlannedInput)
                        .Where(value => value is not null)
                        .Select(value => value!)
                        .Take(TurnBasedMaxBatchInputs)
                        .ToArray() ?? [];

                static string? NormalizePlannedInput(string? value) =>
                    value?.Trim().ToLowerInvariant() switch
                    {
                        "arrowup" or "up" => "ArrowUp",
                        "arrowdown" or "down" => "ArrowDown",
                        "arrowleft" or "left" => "ArrowLeft",
                        "arrowright" or "right" => "ArrowRight",
                        "w" => "W",
                        "a" => "A",
                        "s" => "S",
                        "d" => "D",
                        _ => null
                    };

                static string? NormalizePlanText(string? value, int maxChars)
                {
                    if (string.IsNullOrWhiteSpace(value))
                        return null;
                    return TrimForMeta(value.Trim(), maxChars);
                }
            }

            internal sealed class TurnBasedTransitionTracker
            {
                const int StateFingerprintSide = 96;
                const double StateMatchMeanThreshold = 0.004;
                const double StateMatchChangedRatioThreshold = 0.0025;
                const double ActionStateMatchMeanThreshold = 0.0005;
                const double ActionStateMatchChangedRatioThreshold = 0.00075;
                const double MinimumDirectionalMotionChangedRatio = 0.00085;
                const double ExternalChangeMeanThreshold = 0.006;
                const double ExternalChangeRatioThreshold = 0.006;
                const double CoherentExternalChangeMeanThreshold = 0.003;
                const double CoherentExternalChangeRatioThreshold = 0.02;
                internal const double MinimumStructuredPlanConfidence = 0.55;
                const int ChangeTileSide = 4;
                const int RecurrentRegionLimit = 18;
                const int RecurrentRegionMaxAge = 12;
                const int StatePrototypeLimit = 512;

                sealed class StatePrototype(string id, byte[] fingerprint)
                {
                    internal string Id { get; } = id;
                    internal byte[] Fingerprint { get; set; } = fingerprint;
                }

                readonly record struct TurnTransition(
                    string From,
                    string Action,
                    string To,
                    string Result,
                    bool ReturnedToKnownState);

                readonly record struct NavigationEdge(
                    string FromState,
                    Point From,
                    string Action,
                    string ToState,
                    Point To,
                    bool Blocked);

                readonly record struct ChangeRegion(
                    int Left,
                    int Top,
                    int Right,
                    int Bottom,
                    int ChangedPixels);

                readonly record struct ChangeAnalysis(
                    double MeanDelta,
                    double ChangedRatio,
                    IReadOnlyList<ChangeRegion> Regions,
                    IReadOnlyList<ChangeRegion> NovelRegions,
                    int PredictableRegionCount,
                    int AuxiliaryRegionCount,
                    bool IsAuxiliaryOnly,
                    bool HasDistantRegions,
                    bool HasCausalDistantChange,
                    bool IsBroad)
                {
                    internal bool IsSalient => HasDistantRegions || IsBroad;
                    internal bool RequiresImmediateReanalysis =>
                        IsBroad || HasCausalDistantChange;
                }

                sealed class RecurrentChangeRegion(
                    ChangeRegion bounds,
                    int lastSeenTransition,
                    string? actionLabel)
                {
                    internal ChangeRegion Bounds { get; set; } = bounds;
                    internal int Hits { get; set; } = 1;
                    internal int LastSeenTransition { get; set; } = lastSeenTransition;
                    internal HashSet<string> ActionLabels { get; } =
                        string.IsNullOrWhiteSpace(actionLabel)
                            ? []
                            : [actionLabel];
                }

                sealed class CausalEventEvidence(
                    string baselineState,
                    string triggerAction,
                    string changedState,
                    string changeSummary)
                {
                    internal string BaselineState { get; } = baselineState;
                    internal string TriggerAction { get; } = triggerAction;
                    internal string ChangedState { get; } = changedState;
                    internal string ChangeSummary { get; } = changeSummary;
                    internal int TransitionAge { get; set; }
                    internal bool ReturnObserved { get; set; }
                    internal string? ReturnAction { get; set; }
                    internal bool ReenactmentObserved { get; set; }
                    internal string? ReenactmentAction { get; set; }
                }

                internal readonly record struct BaselineAssessment(
                    bool ExternalStateChange,
                    string Summary);

                internal readonly record struct BatchStepAssessment(
                    bool ContinueBatch,
                    bool KnownTransition,
                    bool SalientChange,
                    bool NoEffect,
                    string Summary);

                readonly List<StatePrototype> states = [];
                readonly Queue<TurnTransition> transitions = new();
                readonly HashSet<string> observedDirectionalInputs = new(StringComparer.Ordinal);
                readonly Queue<double> ordinaryTransitionRatios = new();
                readonly List<RecurrentChangeRegion> recurrentActionRegions = [];
                readonly Queue<string> mechanicsHypothesisHistory = new();
                readonly Queue<string> salientObservationHistory = new();
                readonly Queue<string> priorEpochMechanicsEvidence = new();
                readonly Dictionary<string, ActionDto> directionalClickTemplates = new(
                    StringComparer.Ordinal);
                readonly Dictionary<string, Point> navigationStatePositions = new(
                    StringComparer.Ordinal);
                readonly Queue<NavigationEdge> navigationEdges = new();
                readonly ShortTermPlanTracker shortTermPlan;
                bool[] volatilePixels = [];
                bool[] auxiliaryStatePixels = [];
                string? currentStateId;
                string? lastChangeSummary;
                string? worldStateSummary;
                string? mechanicsHypothesis;
                string? salientChangeObservation;
                IReadOnlyList<RectangleF> salientChangeRegions = [];
                int nextStateNumber = 1;
                int actionTransitionNumber;
                bool actionBaselineCalibrated;
                byte[]? batchOriginFingerprint;
                int consecutiveBatchNoEffect;
                int auxiliarySignalTransitions;
                int consecutiveAuxiliarySignalTransitions;
                bool auxiliarySignalChangedDuringKnownStateReturn;
                string? attemptOriginStateId;
                int attemptDirectionalInputs;
                int? inferredAttemptInputLimit;
                int boundedAttemptResetCount;
                CausalEventEvidence? causalEvent;
                bool previousProposalWasDiscardedAsStale;

                internal bool RequiresReanalysis { get; private set; }
                internal IReadOnlyList<RectangleF> SalientChangeRegions =>
                    salientChangeRegions;
                internal string? CurrentStateId => currentStateId;
                internal bool HasActiveCausalEvent => causalEvent is not null;
                internal bool PreferFastBatchSettle =>
                    HasRecentSuccessfulDirectionalEvidence(1);

                internal TurnBasedTransitionTracker(
                    ShortTermPlanTracker? shortTermPlan = null)
                {
                    this.shortTermPlan = shortTermPlan ?? new ShortTermPlanTracker();
                }

                internal bool CanUseExecutionBatch =>
                    CanProposeExecutionBatch &&
                    (HasHighConfidenceStructuredPlan ||
                     HasDirectionalExecutionEvidence);

                internal bool CanProposeExecutionBatch =>
                    states.Count >= 1 &&
                    !RequiresReanalysis;

                bool HasHighConfidenceStructuredPlan =>
                    shortTermPlan.HasExecutableDirectionalPlan &&
                    shortTermPlan.EffectiveConfidence >= MinimumStructuredPlanConfidence;

                bool HasDirectionalExecutionEvidence =>
                    transitions.Count >= 4 &&
                    observedDirectionalInputs.Count >= 2 ||
                    HasPredictableDirectionalRun() ||
                    HasPlanBackedDirectionalEvidence();

                int TransitionMemoryLimit =>
                    Math.Max(12, TurnBasedMaxBatchInputs);

                internal int AdvertisedMaxExecutionBatchLength =>
                    TurnBasedMaxBatchInputs;

                internal int MaxExecutionBatchLength =>
                    CanUseExecutionBatch ? TurnBasedMaxBatchInputs : 0;

                internal int PreferredAdvertisedExecutionBatchMinimum =>
                    Math.Min(6, AdvertisedMaxExecutionBatchLength);

                internal string ExecutionBatchMode =>
                    HasRecentSuccessfulDirectionalEvidence(4)
                        ? "mature_progressive"
                        : HasRecentSuccessfulDirectionalEvidence(1)
                            ? "progressive"
                            : "aggressive_immediate";

                internal void Reset()
                {
                    states.Clear();
                    transitions.Clear();
                    observedDirectionalInputs.Clear();
                    ordinaryTransitionRatios.Clear();
                    recurrentActionRegions.Clear();
                    mechanicsHypothesisHistory.Clear();
                    salientObservationHistory.Clear();
                    priorEpochMechanicsEvidence.Clear();
                    directionalClickTemplates.Clear();
                    navigationStatePositions.Clear();
                    navigationEdges.Clear();
                    volatilePixels = [];
                    auxiliaryStatePixels = [];
                    currentStateId = null;
                    lastChangeSummary = null;
                    worldStateSummary = null;
                    mechanicsHypothesis = null;
                    salientChangeObservation = null;
                    salientChangeRegions = [];
                    RequiresReanalysis = false;
                    shortTermPlan.Invalidate("the foreground interaction context changed");
                    nextStateNumber = 1;
                    actionTransitionNumber = 0;
                    actionBaselineCalibrated = false;
                    batchOriginFingerprint = null;
                    consecutiveBatchNoEffect = 0;
                    auxiliarySignalTransitions = 0;
                    consecutiveAuxiliarySignalTransitions = 0;
                    auxiliarySignalChangedDuringKnownStateReturn = false;
                    attemptOriginStateId = null;
                    attemptDirectionalInputs = 0;
                    inferredAttemptInputLimit = null;
                    boundedAttemptResetCount = 0;
                    causalEvent = null;
                    previousProposalWasDiscardedAsStale = false;
                }

                internal void BeginExternalStateEpoch()
                {
                    PreservePriorEpochMechanicsEvidence();
                    var current = states.FirstOrDefault(state =>
                        string.Equals(state.Id, currentStateId, StringComparison.Ordinal));
                    states.Clear();
                    if (current is not null)
                        states.Add(current);
                    transitions.Clear();
                    observedDirectionalInputs.Clear();
                    ordinaryTransitionRatios.Clear();
                    recurrentActionRegions.Clear();
                    mechanicsHypothesisHistory.Clear();
                    salientObservationHistory.Clear();
                    navigationStatePositions.Clear();
                    navigationEdges.Clear();
                    volatilePixels = [];
                    worldStateSummary = null;
                    mechanicsHypothesis = null;
                    salientChangeObservation = null;
                    actionBaselineCalibrated = true;
                    batchOriginFingerprint = null;
                    consecutiveBatchNoEffect = 0;
                    auxiliarySignalTransitions = 0;
                    consecutiveAuxiliarySignalTransitions = 0;
                    auxiliarySignalChangedDuringKnownStateReturn = false;
                    attemptOriginStateId = currentStateId;
                    attemptDirectionalInputs = 0;
                    inferredAttemptInputLimit = null;
                    boundedAttemptResetCount = 0;
                    causalEvent = null;
                    previousProposalWasDiscardedAsStale = true;
                    if (currentStateId is not null)
                        navigationStatePositions[currentStateId] = Point.Empty;
                }

                internal void BeginBatch(
                    ScreenObservationFrame frame,
                    Rectangle region)
                {
                    batchOriginFingerprint = ExtractRegionFingerprint(frame, region);
                    consecutiveBatchNoEffect = 0;
                }

                internal void ObserveState(
                    ScreenObservationFrame frame,
                    Rectangle region)
                {
                    var fingerprint = ExtractRegionFingerprint(frame, region);
                    if (fingerprint.Length == 0)
                        return;

                    EnsureVolatileMask(fingerprint.Length);
                    if (currentStateId is null)
                    {
                        currentStateId = ResolveState(fingerprint, out _);
                        navigationStatePositions[currentStateId] = Point.Empty;
                        attemptOriginStateId = currentStateId;
                        actionBaselineCalibrated = false;
                        Console.WriteLine(
                            $"[turn] primary state established; immediate structured batching is available up to {AdvertisedMaxExecutionBatchLength} inputs.");
                        return;
                    }

                    var current = states.FirstOrDefault(state =>
                        string.Equals(state.Id, currentStateId, StringComparison.Ordinal));
                    if (current is null)
                    {
                        currentStateId = ResolveState(fingerprint, out _);
                        navigationStatePositions.TryAdd(
                            currentStateId,
                            Point.Empty);
                        return;
                    }

                    LearnAmbientPixels(current.Fingerprint, fingerprint);
                    current.Fingerprint = fingerprint;
                    actionBaselineCalibrated = true;
                }

                internal BaselineAssessment PrepareActionBaseline(
                    ScreenObservationFrame frame,
                    Rectangle region)
                {
                    var fingerprint = ExtractRegionFingerprint(frame, region);
                    if (fingerprint.Length == 0 || currentStateId is null)
                    {
                        ObserveState(frame, region);
                        return new BaselineAssessment(false, "stable");
                    }

                    EnsureVolatileMask(fingerprint.Length);
                    var current = states.FirstOrDefault(state =>
                        string.Equals(state.Id, currentStateId, StringComparison.Ordinal));
                    if (current is null)
                    {
                        currentStateId = ResolveState(fingerprint, out _);
                        return new BaselineAssessment(false, "stable");
                    }

                    var earlierFingerprint = current.Fingerprint;
                    var difference = StateDifference(earlierFingerprint, fingerprint);
                    var analysis = AnalyzeChanges(
                        earlierFingerprint,
                        fingerprint,
                        learnActionPatterns: false);
                    var aggregateChange =
                        difference.MeanDelta >= ExternalChangeMeanThreshold &&
                        difference.ChangedRatio >= ExternalChangeRatioThreshold;
                    var coherentChange =
                        analysis.Regions.Count > 0 &&
                        difference.MeanDelta >= CoherentExternalChangeMeanThreshold &&
                        difference.ChangedRatio >= CoherentExternalChangeRatioThreshold;
                    var coldStartCosmeticDrift =
                        !actionBaselineCalibrated &&
                        difference.MeanDelta < ExternalChangeMeanThreshold &&
                        !analysis.IsSalient;
                    var peripheralOnlyChange =
                        IsPeripheralOnlyChange(analysis.Regions);
                    var externalStateChange =
                        !coldStartCosmeticDrift &&
                        !peripheralOnlyChange &&
                        (aggregateChange || coherentChange || analysis.IsSalient);
                    if (!externalStateChange)
                    {
                        LearnAmbientPixels(earlierFingerprint, fingerprint);
                        current.Fingerprint = fingerprint;
                        actionBaselineCalibrated = true;
                        salientChangeRegions = [];
                        var classification = coldStartCosmeticDrift
                            ? "cold_start_ambient_calibration"
                            : peripheralOnlyChange
                                ? "peripheral_window_chrome"
                            : analysis.Regions.Count == 0
                                ? "dispersed_ambient_drift"
                                : "stable_or_ambient";
                        return new BaselineAssessment(
                            false,
                            $"{classification}; {FormatChangeAnalysis(analysis)}");
                    }

                    var previousState = currentStateId;
                    SetSalientChangeRegions(analysis);
                    currentStateId = ResolveState(fingerprint, out _);
                    var rebased = states.FirstOrDefault(state =>
                        string.Equals(state.Id, currentStateId, StringComparison.Ordinal));
                    if (rebased is not null)
                        rebased.Fingerprint = fingerprint;
                    actionBaselineCalibrated = true;
                    lastChangeSummary =
                        $"external state change while awaiting the model: {previousState}->{currentStateId}; {FormatChangeAnalysis(analysis)}";
                    RequiresReanalysis = true;
                    shortTermPlan.Invalidate("the interaction state changed while the model was planning");
                    return new BaselineAssessment(true, lastChangeSummary);
                }

                internal bool RecordTransition(
                    ScreenObservationFrame before,
                    ScreenObservationFrame after,
                    Rectangle region,
                    ResolvedActionSnapshot action,
                    ObservationAssessment assessment)
                {
                    var wasExecutionReady = CanUseExecutionBatch;
                    var beforeFingerprint = ExtractRegionFingerprint(before, region);
                    var afterFingerprint = ExtractRegionFingerprint(after, region);
                    if (beforeFingerprint.Length == 0 || afterFingerprint.Length == 0)
                        return false;

                    EnsureVolatileMask(beforeFingerprint.Length);
                    var beforeState = currentStateId ?? ResolveState(beforeFingerprint, out _);
                    var beforePrototype = states.FirstOrDefault(state =>
                        string.Equals(state.Id, beforeState, StringComparison.Ordinal));
                    if (beforePrototype is not null)
                        beforePrototype.Fingerprint = beforeFingerprint;
                    var actionLabel = TurnActionLabel(action);
                    var changeAnalysis = AnalyzeChanges(
                        beforeFingerprint,
                        afterFingerprint,
                        learnActionPatterns: true,
                        actionLabel: actionLabel);
                    var persistentVisualChange =
                        HasPersistentVisualChange(changeAnalysis);
                    var movementEvidence = HasDirectionalMotionEvidence(
                        actionLabel,
                        changeAnalysis);
                    string afterState;
                    string result;
                    var returnedToKnownState = false;
                    if (assessment.ActionOutcome == ActionOutcomeState.NoEffect ||
                        !persistentVisualChange ||
                        IsCanonicalDirectionalInput(actionLabel) &&
                        !movementEvidence &&
                        !changeAnalysis.RequiresImmediateReanalysis)
                    {
                        afterState = beforeState;
                        result = "no_effect";
                    }
                    else if (assessment.ActionOutcome == ActionOutcomeState.Confirmed ||
                             assessment.VisualChange == VisualChangeState.Changed)
                    {
                        afterState = ResolveActionState(
                            afterFingerprint,
                            actionLabel,
                            changeAnalysis,
                            out var isNewState);
                        returnedToKnownState = !isNewState &&
                                               !string.Equals(
                                                   beforeState,
                                                   afterState,
                                                   StringComparison.Ordinal);
                        result = "changed";
                    }
                    else
                    {
                        var difference = StateDifference(beforeFingerprint, afterFingerprint);
                        afterState = IsSameState(difference)
                            ? beforeState
                            : ResolveState(afterFingerprint, out _);
                        result = "uncertain";
                    }

                    currentStateId = afterState;
                    var afterPrototype = states.FirstOrDefault(state =>
                        string.Equals(state.Id, afterState, StringComparison.Ordinal));
                    if (afterPrototype is not null)
                        afterPrototype.Fingerprint = afterFingerprint;
                    RememberDirectionalClickTemplate(action.Action, actionLabel);
                    var planExpectedAction =
                        shortTermPlan.IsExpectedInput(actionLabel, beforeState);
                    if (IsCanonicalDirectionalInput(actionLabel))
                        observedDirectionalInputs.Add(actionLabel);
                    transitions.Enqueue(new TurnTransition(
                        beforeState,
                        actionLabel,
                        afterState,
                        result,
                        returnedToKnownState));
                    while (transitions.Count > TransitionMemoryLimit)
                        transitions.Dequeue();
                    RecordNavigationTransition(
                        beforeState,
                        actionLabel,
                        afterState,
                        result);
                    RecordChangeEvidence(
                        beforeState,
                        actionLabel,
                        afterState,
                        changeAnalysis,
                        result,
                        movementEvidence,
                        returnedToKnownState);
                    if (result == "no_effect")
                        shortTermPlan.Invalidate("a planned directional input had no effect");
                    else if (changeAnalysis.RequiresImmediateReanalysis)
                        shortTermPlan.Invalidate("a planned input caused a novel or broad state change");
                    else if (planExpectedAction && IsCanonicalDirectionalInput(actionLabel))
                    {
                        shortTermPlan.RecordDirectionalResult(
                            actionLabel,
                            beforeState,
                            afterState,
                            result == "changed");
                    }
                    if (!wasExecutionReady && CanUseExecutionBatch)
                    {
                        Console.WriteLine(
                            "[turn] phase exploration -> execution_ready; bounded ordered key sequences are now allowed.");
                    }
                    return result == "no_effect";
                }

                internal BatchStepAssessment RecordBatchStep(
                    ScreenObservationFrame before,
                    ScreenObservationFrame after,
                    Rectangle region,
                    string key,
                    bool actionReactionObserved = false) =>
                    RecordBatchStepCore(
                        before,
                        after,
                        region,
                        CanonicalKeyLabel(key),
                        actionReactionObserved);

                internal BatchStepAssessment RecordBatchActionStep(
                    ScreenObservationFrame before,
                    ScreenObservationFrame after,
                    Rectangle region,
                    ResolvedActionSnapshot action,
                    bool actionReactionObserved)
                {
                    var actionLabel = TurnActionLabel(action);
                    RememberDirectionalClickTemplate(action.Action, actionLabel);
                    return RecordBatchStepCore(
                        before,
                        after,
                        region,
                        actionLabel,
                        actionReactionObserved);
                }

                internal bool TryBuildObservedDirectionalFollowUps(
                    ActionDto firstAction,
                    IReadOnlyList<string> remainingInputs,
                    out ActionDto[] followUps)
                {
                    followUps = [];
                    if (firstAction.Type is not ("click" or "double_click") ||
                        !OpenAiResponsesService.TryGetObservedTurnInputLabel(
                            firstAction,
                            out var firstLabel))
                    {
                        return false;
                    }

                    RememberDirectionalClickTemplate(firstAction, firstLabel);
                    var result = new List<ActionDto>(remainingInputs.Count);
                    foreach (var input in remainingInputs)
                    {
                        if (!OpenAiResponsesService.TryNormalizeDirectionalLabel(
                                input,
                                out var label))
                        {
                            break;
                        }

                        ActionDto? template = string.Equals(
                            label,
                            firstLabel,
                            StringComparison.Ordinal)
                            ? firstAction
                            : directionalClickTemplates.GetValueOrDefault(label);
                        if (template is null)
                            break;
                        result.Add(CloneDirectionalClick(template, label));
                    }

                    if (result.Count == 0)
                        return false;
                    followUps = result.ToArray();
                    return true;
                }

                void RememberDirectionalClickTemplate(
                    ActionDto action,
                    string actionLabel)
                {
                    if (action.Type is not ("click" or "double_click") ||
                        !HasExplicitPoint(action) ||
                        !OpenAiResponsesService.TryNormalizeDirectionalLabel(
                            actionLabel,
                            out var normalizedLabel))
                    {
                        return;
                    }

                    directionalClickTemplates[normalizedLabel] =
                        CloneDirectionalClick(action, normalizedLabel);
                }

                static ActionDto CloneDirectionalClick(
                    ActionDto source,
                    string resolvedLabel) =>
                    new()
                    {
                        Type = source.Type,
                        X = source.X,
                        Y = source.Y,
                        XPx = source.XPx,
                        YPx = source.YPx,
                        BBox = source.BBox is null
                            ? null
                            : new BBox
                            {
                                Left = source.BBox.Left,
                                Top = source.BBox.Top,
                                Right = source.BBox.Right,
                                Bottom = source.BBox.Bottom
                            },
                        Button = source.Button,
                        Confidence = source.Confidence,
                        Note = source.Note,
                        ResolvedTurnInputLabel = resolvedLabel
                    };

                BatchStepAssessment RecordBatchStepCore(
                    ScreenObservationFrame before,
                    ScreenObservationFrame after,
                    Rectangle region,
                    string actionLabel,
                    bool actionReactionObserved)
                {
                    var beforeFingerprint = ExtractRegionFingerprint(before, region);
                    var afterFingerprint = ExtractRegionFingerprint(after, region);
                    if (beforeFingerprint.Length == 0 || afterFingerprint.Length == 0)
                    {
                        return new BatchStepAssessment(
                            false,
                            false,
                            true,
                            false,
                            "intermediate observation was unavailable");
                    }

                    EnsureVolatileMask(beforeFingerprint.Length);
                    var beforeState = currentStateId ?? ResolveState(beforeFingerprint, out _);
                    var difference = StateDifference(beforeFingerprint, afterFingerprint);
                    var analysis = AnalyzeChanges(
                        beforeFingerprint,
                        afterFingerprint,
                        learnActionPatterns: true,
                        actionLabel: actionLabel);
                    var movementEvidence = HasDirectionalMotionEvidence(
                        actionLabel,
                        analysis);
                    var result = HasPersistentVisualChange(analysis) &&
                                 (!IsCanonicalDirectionalInput(actionLabel) ||
                                  movementEvidence ||
                                  analysis.RequiresImmediateReanalysis)
                        ? "changed"
                        : "no_effect";
                    var cumulativeDifference = batchOriginFingerprint is { Length: > 0 } origin &&
                                               origin.Length == afterFingerprint.Length
                        ? StateDifference(origin, afterFingerprint)
                        : difference;
                    var cumulativeChangeObserved = !IsSameState(cumulativeDifference);
                    if (result == "no_effect")
                        consecutiveBatchNoEffect++;
                    else
                        consecutiveBatchNoEffect = 0;
                    var confirmedNoEffect = result == "no_effect";
                    var returnedToKnownState = false;
                    var isNewState = false;
                    var afterState = result == "no_effect"
                        ? beforeState
                        : ResolveActionState(
                            afterFingerprint,
                            actionLabel,
                            analysis,
                            out isNewState);
                    if (result != "no_effect")
                    {
                        returnedToKnownState = !isNewState &&
                                               !string.Equals(
                                                   beforeState,
                                                   afterState,
                                                   StringComparison.Ordinal);
                    }
                    var knownTransition = transitions.Any(transition =>
                        string.Equals(transition.From, beforeState, StringComparison.Ordinal) &&
                        string.Equals(transition.Action, actionLabel, StringComparison.Ordinal) &&
                        string.Equals(transition.To, afterState, StringComparison.Ordinal) &&
                        string.Equals(transition.Result, result, StringComparison.Ordinal));
                    var predictableActionPattern =
                        IsPredictableDirectionalAction(actionLabel);
                    var planBackedAction =
                        shortTermPlan.IsExpectedInput(actionLabel, beforeState) &&
                        shortTermPlan.EffectiveConfidence >= MinimumStructuredPlanConfidence;

                    currentStateId = afterState;
                    var prototype = states.FirstOrDefault(state =>
                        string.Equals(state.Id, afterState, StringComparison.Ordinal));
                    if (prototype is not null)
                        prototype.Fingerprint = afterFingerprint;
                    observedDirectionalInputs.Add(actionLabel);
                    transitions.Enqueue(new TurnTransition(
                        beforeState,
                        actionLabel,
                        afterState,
                        result,
                        returnedToKnownState));
                    while (transitions.Count > TransitionMemoryLimit)
                        transitions.Dequeue();
                    RecordNavigationTransition(
                        beforeState,
                        actionLabel,
                        afterState,
                        result);

                    RecordChangeEvidence(
                        beforeState,
                        actionLabel,
                        afterState,
                        analysis,
                        result,
                        movementEvidence,
                        returnedToKnownState);
                    if (confirmedNoEffect)
                        shortTermPlan.Invalidate("a batched planned input had no persistent visual effect");
                    else if (analysis.RequiresImmediateReanalysis)
                        shortTermPlan.Invalidate("a batched plan input reached a novel or broad transition");
                    else if (planBackedAction && result == "changed")
                    {
                        shortTermPlan.RecordDirectionalResult(
                            actionLabel,
                            beforeState,
                            afterState,
                            changed: true);
                    }
                    var summary =
                        $"{beforeState} --{actionLabel}--> {afterState} [{result}]; transient_reaction={actionReactionObserved.ToString().ToLowerInvariant()}; movement_evidence={movementEvidence.ToString().ToLowerInvariant()}; predictable_action_pattern={predictableActionPattern.ToString().ToLowerInvariant()}; plan_backed={planBackedAction.ToString().ToLowerInvariant()}; no_effect_streak={consecutiveBatchNoEffect}; cumulative_change={cumulativeChangeObserved.ToString().ToLowerInvariant()}; cumulative_ratio={cumulativeDifference.ChangedRatio:0.####}; {FormatChangeAnalysis(analysis)}";
                    return new BatchStepAssessment(
                        !analysis.RequiresImmediateReanalysis && !confirmedNoEffect,
                        knownTransition,
                        analysis.RequiresImmediateReanalysis || confirmedNoEffect,
                        result == "no_effect",
                        summary);
                }

                internal void UpdateWorkingMemory(ActionDto action)
                {
                    previousProposalWasDiscardedAsStale = false;
                    var nextWorldState = NormalizeWorkingMemory(action.WorldStateSummary);
                    var nextHypothesis = NormalizeWorkingMemory(action.MechanicsHypothesis);
                    var nextSalientObservation = NormalizeWorkingMemory(
                        action.SalientChangeObservation);
                    if (!string.Equals(nextWorldState, worldStateSummary, StringComparison.Ordinal))
                    {
                        worldStateSummary = nextWorldState;
                        Console.WriteLine(
                            worldStateSummary is null
                                ? "[turn-memory] world_state cleared."
                                : $"[turn-memory] world_state={worldStateSummary}");
                    }
                    if (!string.Equals(nextHypothesis, mechanicsHypothesis, StringComparison.Ordinal))
                    {
                        mechanicsHypothesis = nextHypothesis;
                        if (mechanicsHypothesis is not null &&
                            !string.Equals(
                                mechanicsHypothesisHistory.LastOrDefault(),
                                mechanicsHypothesis,
                                StringComparison.Ordinal))
                        {
                            mechanicsHypothesisHistory.Enqueue(mechanicsHypothesis);
                            while (mechanicsHypothesisHistory.Count > 4)
                                mechanicsHypothesisHistory.Dequeue();
                        }
                        Console.WriteLine(
                            mechanicsHypothesis is null
                                ? "[turn-memory] hypothesis cleared."
                                : $"[turn-memory] hypothesis={mechanicsHypothesis}");
                    }
                    if (!string.Equals(
                            nextSalientObservation,
                            salientChangeObservation,
                            StringComparison.Ordinal))
                    {
                        salientChangeObservation = nextSalientObservation;
                        if (salientChangeObservation is not null &&
                            !string.Equals(
                                salientObservationHistory.LastOrDefault(),
                                salientChangeObservation,
                                StringComparison.Ordinal))
                        {
                            salientObservationHistory.Enqueue(
                                salientChangeObservation);
                            while (salientObservationHistory.Count > 6)
                                salientObservationHistory.Dequeue();
                        }
                        Console.WriteLine(
                            salientChangeObservation is null
                                ? "[turn-memory] salient_change_observation cleared."
                                : $"[turn-memory] salient_change={salientChangeObservation}");
                    }
                }

                void PreservePriorEpochMechanicsEvidence()
                {
                    AddPriorEpochEvidence(salientChangeObservation);
                    if (causalEvent is not null)
                    {
                        AddPriorEpochEvidence(
                            $"application-observed transition: {causalEvent.BaselineState} --{causalEvent.TriggerAction}--> {causalEvent.ChangedState}; {causalEvent.ChangeSummary}");
                    }
                }

                void AddPriorEpochEvidence(string? evidence)
                {
                    if (string.IsNullOrWhiteSpace(evidence) ||
                        priorEpochMechanicsEvidence.Contains(
                            evidence,
                            StringComparer.Ordinal))
                    {
                        return;
                    }

                    priorEpochMechanicsEvidence.Enqueue(evidence);
                    while (priorEpochMechanicsEvidence.Count > 8)
                        priorEpochMechanicsEvidence.Dequeue();
                }

                internal bool HasRequiredSalientObservation(ActionDto action) =>
                    !RequiresReanalysis ||
                    !string.IsNullOrWhiteSpace(action.SalientChangeObservation);

                internal void AcknowledgeReanalysis()
                {
                    if (!RequiresReanalysis)
                        return;
                    RequiresReanalysis = false;
                    salientChangeRegions = [];
                    Console.WriteLine("[turn] salient transition reanalysis acknowledged; one-step control resumed.");
                }

                internal string BuildPromptSummary()
                {
                    if (string.IsNullOrWhiteSpace(currentStateId))
                        return "";

                    var builder = new StringBuilder()
                        .AppendLine($"TURN_STATE: {currentStateId}");
                    if (transitions.Count > 0)
                    {
                        builder.AppendLine("TURN_TRANSITIONS (observed, oldest to newest):");
                        foreach (var transition in transitions)
                        {
                            builder.AppendLine(
                                $"- {transition.From} --{transition.Action}--> {transition.To} [{transition.Result}]" +
                                (transition.ReturnedToKnownState
                                    ? " [returned_to_known_state]"
                                    : ""));
                        }
                        builder.AppendLine("TURN_TOPOLOGY (latest observed directed edges):");
                        foreach (var stateGroup in transitions
                                     .Where(transition =>
                                         IsCanonicalDirectionalInput(transition.Action))
                                     .GroupBy(transition => transition.From))
                        {
                            var edges = stateGroup
                                .GroupBy(transition => transition.Action)
                                .Select(actionGroup => actionGroup.Last())
                                .Select(transition =>
                                    string.Equals(
                                        transition.Result,
                                        "no_effect",
                                        StringComparison.Ordinal)
                                        ? $"{transition.Action}=blocked"
                                        : $"{transition.Action}->{transition.To} [{transition.Result}]");
                            builder.AppendLine(
                                $"- {stateGroup.Key}: {string.Join("; ", edges)}");
                        }
                    }
                    AppendNavigationSummary(builder);
                    if (!string.IsNullOrWhiteSpace(lastChangeSummary))
                        builder.AppendLine($"TURN_VISUAL_CHANGE_REGIONS (percent of interaction region): {lastChangeSummary}");
                    if (recurrentActionRegions.Any(pattern => pattern.Hits >= 2))
                    {
                        builder.AppendLine(
                            "TURN_AUXILIARY_CHANGES: recurring action-correlated visual regions were observed and are treated as predictable evidence, not by themselves as a new world-rule event.");
                    }
                    if (auxiliarySignalTransitions > 0)
                    {
                        builder.AppendLine(
                            $"TURN_AUXILIARY_SIGNAL: a recurring peripheral visual indicator changed alongside {auxiliarySignalTransitions} observed input transition(s), including {consecutiveAuxiliarySignalTransitions} consecutively. It is excluded from board-state identity but retained as potentially task-relevant evidence; infer its meaning from its trajectory without assuming a semantic label.");
                        if (auxiliarySignalChangedDuringKnownStateReturn)
                        {
                            builder.AppendLine(
                                "TURN_AUXILIARY_RETURN_EVIDENCE: the peripheral indicator changed during a transition that returned the board to a previously observed state. Consider whether the indicator constrains the action sequence; its meaning is not predetermined.");
                        }
                        if (consecutiveAuxiliarySignalTransitions >= 6 &&
                            inferredAttemptInputLimit is null)
                        {
                            builder.AppendLine(
                                "TURN_POSSIBLE_ACTION_COST: the same peripheral indicator has changed across many consecutive inputs. Until disproved, treat interaction inputs as potentially consuming a bounded attempt resource: prefer the shortest visible or confirmed route, avoid redundant calibration and known closed paths, and use the per-input barrier only for genuine uncertainty.");
                        }
                    }
                    if (inferredAttemptInputLimit is int attemptLimit)
                    {
                        var estimatedRemaining = Math.Max(
                            0,
                            attemptLimit - attemptDirectionalInputs);
                        builder.AppendLine(
                            $"TURN_ATTEMPT_BUDGET_EVIDENCE: {boundedAttemptResetCount} broad return(s) to the attempt origin occurred after repeated directional inputs while distant world elements reappeared and the recurring peripheral indicator changed. This is evidence of a bounded attempt or action resource, not evidence that the last semantic target itself caused the reset.");
                        builder.AppendLine(
                            $"TURN_ATTEMPT_BUDGET_STATUS: observed_limit_about={attemptLimit}; current_attempt_inputs={attemptDirectionalInputs}; estimated_remaining_at_most={estimatedRemaining}. Complete the remaining goal through the shortest confirmed topology; do not reorder already confirmed semantic prerequisites or repeat exploratory detours merely because a reset occurred.");
                    }
                    if (transitions.LastOrDefault().ReturnedToKnownState)
                    {
                        builder.AppendLine(
                            "TURN_STATE_RETURN: the latest movement returned to a previously observed board state after auxiliary UI differences were excluded. Treat this recurrence as authoritative topology and do not repeat the same closed route unless it serves a deliberate new test.");
                    }
                    if (LatestTransitionHadNoEffect)
                    {
                        builder.AppendLine(
                            "TURN_HYPOTHESIS_CONTRADICTION: the latest state-changing input had no effect. First distinguish a blocked logical move from an unconfirmed input modality. If fixed visible directional controls exist and a keyboard direction failed, preserve the route and remap its longest valid prefix to an observed click batch; do not test one button at a time. Otherwise revise the causal model and prioritize a distinct reachable affordance.");
                    }
                    builder.AppendLine(
                        $"TURN_REANALYSIS_REQUIRED: {RequiresReanalysis.ToString().ToLowerInvariant()}");
                    if (!string.IsNullOrWhiteSpace(worldStateSummary))
                        builder.AppendLine($"TURN_WORLD_STATE_MEMORY: {worldStateSummary}");
                    if (previousProposalWasDiscardedAsStale)
                    {
                        builder.AppendLine(
                            "TURN_TRANSIENT_CONTEXT_WARNING: the immediately preceding model proposal was based on a screen that changed before execution. Its world description, mechanics hypothesis, plan, and proposed action are stale and must not be treated as evidence for the current state.");
                    }
                    if (priorEpochMechanicsEvidence.Count > 0)
                    {
                        builder.AppendLine(
                            "TURN_PRIOR_EPOCH_OBSERVED_EVIDENCE (before/after facts retained across a board or level boundary; when compatible with the current visuals, this evidence has priority over speculative semantic labels):");
                        foreach (var evidence in priorEpochMechanicsEvidence)
                            builder.AppendLine($"- {evidence}");
                    }
                    var priorHypotheses = mechanicsHypothesisHistory
                        .Where(item => !string.Equals(
                            item,
                            mechanicsHypothesis,
                            StringComparison.Ordinal))
                        .TakeLast(3)
                        .ToArray();
                    if (priorHypotheses.Length > 0)
                    {
                        builder.AppendLine(
                            "TURN_PRIOR_MECHANICS_HYPOTHESES (unverified model claims; observed transitions are authoritative):");
                        foreach (var priorHypothesis in priorHypotheses)
                            builder.AppendLine($"- {priorHypothesis}");
                    }
                    if (!string.IsNullOrWhiteSpace(mechanicsHypothesis))
                        builder.AppendLine($"TURN_MECHANICS_HYPOTHESIS: {mechanicsHypothesis}");
                    if (!string.IsNullOrWhiteSpace(salientChangeObservation))
                        builder.AppendLine($"TURN_LAST_SALIENT_OBSERVATION: {salientChangeObservation}");
                    if (causalEvent is not null)
                    {
                        builder.AppendLine(
                            $"TURN_CAUSAL_EVENT_LEDGER: retained independently of the model hypothesis; {causalEvent.BaselineState} --{causalEvent.TriggerAction}--> {causalEvent.ChangedState}; {causalEvent.ChangeSummary}");
                        if (causalEvent.ReturnObserved)
                        {
                            builder.AppendLine(
                                $"TURN_CAUSAL_RETURN: {causalEvent.ChangedState} --{causalEvent.ReturnAction}--> {causalEvent.BaselineState}; the earlier visual state recurred after a reversible input.");
                        }
                        if (causalEvent.ReenactmentObserved)
                        {
                            builder.AppendLine(
                                $"TURN_CAUSAL_REENACTMENT: {causalEvent.BaselineState} --{causalEvent.ReenactmentAction}--> {causalEvent.ChangedState}; the same changed state recurred, forming observed A-B-A-B evidence.");
                        }
                        if (!causalEvent.ReturnObserved)
                        {
                            builder.AppendLine(
                                "TURN_CAUSAL_NEXT_EVIDENCE: Prefer one safe reversible input that tests whether the changed state persists or returns before using an unrelated auxiliary control.");
                        }
                        else if (!causalEvent.ReenactmentObserved)
                        {
                            builder.AppendLine(
                                "TURN_CAUSAL_NEXT_EVIDENCE: A return to the baseline state is observed. Prefer one safe input that could reproduce the changed state before committing to an unrelated control or long route.");
                        }
                        builder.AppendLine(
                            "TURN_CAUSAL_ANALYSIS_HINT: Treat the retained state sequence and labeled temporal images as factual correlation. Infer the most general reversible world rule before trying unrelated auxiliary controls; do not assume a specific rule that is not visibly supported.");
                    }
                    builder.AppendLine(
                        $"TURN_PHASE: {(CanProposeExecutionBatch ? "execution_ready" : "exploration")}");
                    builder.Append(CanProposeExecutionBatch
                        ? $"TURN_EXECUTION_HINT: Aggressive batching is available immediately and is the default. Commit to the strongest visible control hypothesis even when the mechanics are not yet certain; the per-input barrier is the experiment. If compatible prior-epoch observed evidence identifies an affordance sequence or causal role, prefer testing or reusing that observed role before inventing a new meaning from color, position, or visual novelty alone. If the current screen is a gateway, overlay, dialog, title screen, or other obvious pre-task state with a prominent primary affordance, activate that visible affordance directly before trying keyboard aliases, inspecting similarly labelled auxiliary controls, or reasoning about the hidden task. Do not inspect HELP or request another crop merely to reduce semantic uncertainty before the first reversible route. Derive a concrete route to one visible semantic waypoint and put its longest reversible prefix in planned_inputs. In visual navigation or shape-matching tasks, compare the controllable object's current geometry, colors, and orientation with any apparent terminal. If they visibly mismatch and a distinct reachable marker or transformation affordance exists, prefer that informative intermediate waypoint and re-observe its causal effect instead of forcing the incompatible terminal. When a fixed D-pad or equivalent controls are visible and keyboard focus is unconfirmed, prefer an observed click sequence for the entire route. After a keyboard no_effect, remap the route to visible controls in one batch instead of issuing a single test click. Send up to {AdvertisedMaxExecutionBatchLength} inputs and include all visible unconditional turns before the next semantic uncertainty; do not pad a route merely to reach the cap. Confidence of {MinimumStructuredPlanConfidence:0.00} is sufficient. Small recurring auxiliary interface changes do not stop the route; a confirmed no_effect, unavailable observation, broad screen/state transition, or novel local-to-distant causal change interrupts it."
                        : RequiresReanalysis
                            ? "TURN_REANALYSIS_HINT: A non-routine or external state change occurred. Compare every explicitly labeled temporal image actually attached to this request and the reported changed regions. Revise the causal hypothesis, then choose exactly one evidence-driven input."
                            : "TURN_EXPLORATION_HINT: Treat this ledger as observed evidence. Promptly choose one maximally informative, goal-relevant, untried safe discrete input; avoid repeating an input recorded as no_effect. Do not narrate or solve a full route in this turn: observe this one input first. Compare the movable object with visible goal/target patterns before exhaustively testing every control, and keep planned_inputs null unless the route is already mechanically established.");
                    return builder.ToString();
                }

                void AppendNavigationSummary(StringBuilder builder)
                {
                    if (navigationEdges.Count == 0 ||
                        currentStateId is null ||
                        !navigationStatePositions.TryGetValue(
                            currentStateId,
                            out var currentPosition))
                    {
                        return;
                    }

                    builder.AppendLine(
                        $"TURN_NAVIGATION_POSE: state={currentStateId}; relative=({currentPosition.X},{currentPosition.Y}); origin=(0,0) is the first board state in this epoch; X increases right and Y increases down.");
                    builder.AppendLine(
                        "TURN_NAVIGATION_GRAPH (observed movement only; blocked means the input did not change board position):");
                    var latestEdges = navigationEdges
                        .GroupBy(edge => (
                            edge.From.X,
                            edge.From.Y,
                            edge.Action))
                        .Select(group => group.Last())
                        .GroupBy(edge => (
                            edge.From.X,
                            edge.From.Y,
                            edge.FromState))
                        .OrderBy(group => group.Key.Y)
                        .ThenBy(group => group.Key.X);
                    foreach (var group in latestEdges)
                    {
                        var descriptions = group
                            .OrderBy(edge => edge.Action, StringComparer.Ordinal)
                            .Select(edge => edge.Blocked
                                ? $"{edge.Action}=blocked"
                                : $"{edge.Action}->({edge.To.X},{edge.To.Y}) {edge.ToState}");
                        builder.AppendLine(
                            $"- ({group.Key.X},{group.Key.Y}) {group.Key.FromState}: {string.Join("; ", descriptions)}");
                    }
                }

                bool LatestTransitionHadNoEffect =>
                    transitions.Count > 0 &&
                    string.Equals(
                        transitions.Last().Result,
                        "no_effect",
                        StringComparison.Ordinal);

                bool HasPredictableDirectionalRun()
                {
                    var recent = transitions.Reverse().Take(2).ToArray();
                    return recent.Length == 2 &&
                           IsCanonicalDirectionalInput(recent[0].Action) &&
                           recent.All(transition =>
                               string.Equals(
                                   transition.Action,
                                   recent[0].Action,
                                   StringComparison.Ordinal) &&
                               string.Equals(
                                   transition.Result,
                                   "changed",
                                   StringComparison.Ordinal) &&
                               !string.Equals(
                                   transition.From,
                                   transition.To,
                                   StringComparison.Ordinal));
                }

                bool HasPlanBackedDirectionalEvidence() =>
                    shortTermPlan.HasExecutableDirectionalPlan &&
                    HasRecentSuccessfulDirectionalEvidence(1);

                bool HasRecentSuccessfulDirectionalEvidence(int requiredCount) =>
                    transitions
                        .Reverse()
                        .Take(6)
                        .Count(transition =>
                            IsCanonicalDirectionalInput(transition.Action) &&
                            string.Equals(
                                transition.Result,
                                "changed",
                                StringComparison.Ordinal) &&
                            !string.Equals(
                                transition.From,
                                transition.To,
                                StringComparison.Ordinal)) >= requiredCount;

                bool IsPredictableDirectionalAction(string actionLabel) =>
                    IsCanonicalDirectionalInput(actionLabel) &&
                    transitions
                        .Reverse()
                        .Take(6)
                        .Count(transition =>
                            string.Equals(
                                transition.Action,
                                actionLabel,
                                StringComparison.Ordinal) &&
                            string.Equals(
                                transition.Result,
                                "changed",
                                StringComparison.Ordinal) &&
                            !string.Equals(
                                transition.From,
                                transition.To,
                                StringComparison.Ordinal)) >= 3;

                string ResolveActionState(
                    byte[] fingerprint,
                    string actionLabel,
                    ChangeAnalysis analysis,
                    out bool isNew)
                {
                    if (!IsCanonicalDirectionalInput(actionLabel) ||
                        !HasPersistentVisualChange(analysis))
                    {
                        return ResolveState(fingerprint, out isNew);
                    }

                    return ResolveState(
                        fingerprint,
                        out isNew,
                        ActionStateMatchMeanThreshold,
                        ActionStateMatchChangedRatioThreshold);
                }

                string ResolveState(
                    byte[] fingerprint,
                    out bool isNew,
                    double meanThreshold = StateMatchMeanThreshold,
                    double changedRatioThreshold = StateMatchChangedRatioThreshold)
                {
                    StatePrototype? best = null;
                    var bestScore = double.MaxValue;
                    foreach (var state in states)
                    {
                        var difference = StateIdentityDifference(
                            state.Fingerprint,
                            fingerprint);
                        if (!IsSameState(
                                difference,
                                meanThreshold,
                                changedRatioThreshold))
                            continue;
                        var score = difference.ChangedRatio * 4 + difference.MeanDelta;
                        if (score < bestScore)
                        {
                            bestScore = score;
                            best = state;
                        }
                    }

                    if (best is not null)
                    {
                        isNew = false;
                        return best.Id;
                    }

                    var id = $"S{nextStateNumber++}";
                    states.Add(new StatePrototype(id, fingerprint));
                    if (states.Count > StatePrototypeLimit)
                    {
                        var removable = states.FirstOrDefault(state =>
                            !string.Equals(
                                state.Id,
                                currentStateId,
                                StringComparison.Ordinal) &&
                            !string.Equals(state.Id, id, StringComparison.Ordinal));
                        if (removable is not null)
                            states.Remove(removable);
                    }
                    isNew = true;
                    return id;
                }

                void RecordChangeEvidence(
                    string beforeState,
                    string actionLabel,
                    string afterState,
                    ChangeAnalysis analysis,
                    string result,
                    bool movementEvidence,
                    bool returnedToKnownState)
                {
                    lastChangeSummary =
                        $"{beforeState} --{actionLabel}--> {afterState}; {FormatChangeAnalysis(analysis)}";
                    RecordAuxiliarySignalEvidence(
                        analysis,
                        returnedToKnownState);
                    RecordAttemptBudgetEvidence(
                        actionLabel,
                        afterState,
                        analysis,
                        returnedToKnownState);
                    if (!string.Equals(result, "changed", StringComparison.Ordinal))
                    {
                        salientChangeRegions = [];
                        return;
                    }

                    UpdateCausalEvent(
                        beforeState,
                        actionLabel,
                        afterState,
                        analysis);
                    if (analysis.RequiresImmediateReanalysis)
                    {
                        SetSalientChangeRegions(analysis);
                        RequiresReanalysis = true;
                        Console.WriteLine(
                            $"[turn-event] salient visual change detected; reanalysis required; {lastChangeSummary}");
                        return;
                    }

                    if (analysis.HasDistantRegions)
                    {
                        Console.WriteLine(
                            $"[turn-event] distant change recorded without interrupting aggressive execution; {lastChangeSummary}");
                    }

                    salientChangeRegions = [];

                    if (IsCanonicalDirectionalInput(actionLabel) &&
                        movementEvidence &&
                        analysis.ChangedRatio >= MinimumDirectionalMotionChangedRatio)
                    {
                        ordinaryTransitionRatios.Enqueue(analysis.ChangedRatio);
                        while (ordinaryTransitionRatios.Count > TransitionMemoryLimit)
                            ordinaryTransitionRatios.Dequeue();
                    }
                }

                void RecordAuxiliarySignalEvidence(
                    ChangeAnalysis analysis,
                    bool returnedToKnownState)
                {
                    if (analysis.AuxiliaryRegionCount <= 0)
                    {
                        if (analysis.Regions.Count > 0)
                            consecutiveAuxiliarySignalTransitions = 0;
                        return;
                    }

                    auxiliarySignalTransitions++;
                    consecutiveAuxiliarySignalTransitions++;
                    if (returnedToKnownState)
                        auxiliarySignalChangedDuringKnownStateReturn = true;
                }

                void RecordAttemptBudgetEvidence(
                    string actionLabel,
                    string afterState,
                    ChangeAnalysis analysis,
                    bool returnedToKnownState)
                {
                    if (!IsCanonicalDirectionalInput(actionLabel))
                        return;

                    attemptDirectionalInputs++;
                    var returnedToAttemptOrigin =
                        returnedToKnownState &&
                        attemptOriginStateId is not null &&
                        string.Equals(
                            afterState,
                            attemptOriginStateId,
                            StringComparison.Ordinal);
                    if (!IsBoundedAttemptResetEvidence(
                            returnedToAttemptOrigin,
                            attemptDirectionalInputs,
                            analysis.IsBroad,
                            analysis.HasDistantRegions,
                            analysis.NovelRegions.Count,
                            analysis.AuxiliaryRegionCount))
                    {
                        return;
                    }

                    var observedLimit = attemptDirectionalInputs;
                    if (inferredAttemptInputLimit is null ||
                        observedLimit >= inferredAttemptInputLimit.Value * 0.70)
                    {
                        inferredAttemptInputLimit = inferredAttemptInputLimit is null
                            ? observedLimit
                            : Math.Min(
                                inferredAttemptInputLimit.Value,
                                observedLimit);
                    }
                    boundedAttemptResetCount++;
                    attemptDirectionalInputs = 0;
                    Console.WriteLine(
                        $"[turn-budget] bounded attempt evidence detected; observed_limit_about={inferredAttemptInputLimit}; resets={boundedAttemptResetCount}.");
                }

                internal static bool IsBoundedAttemptResetEvidence(
                    bool returnedToAttemptOrigin,
                    int directionalInputs,
                    bool broadChange,
                    bool distantChange,
                    int novelRegionCount,
                    int auxiliaryRegionCount) =>
                    returnedToAttemptOrigin &&
                    directionalInputs >= 8 &&
                    broadChange &&
                    distantChange &&
                    novelRegionCount >= 2 &&
                    auxiliaryRegionCount >= 1;

                void RecordNavigationTransition(
                    string beforeState,
                    string actionLabel,
                    string afterState,
                    string result)
                {
                    if (!TryDirectionalDelta(actionLabel, out var delta))
                        return;

                    if (!navigationStatePositions.TryGetValue(
                            beforeState,
                            out var beforePosition))
                    {
                        beforePosition = navigationStatePositions.Count == 0
                            ? Point.Empty
                            : navigationStatePositions.TryGetValue(
                                currentStateId ?? "",
                                out var currentPosition)
                                ? currentPosition
                                : Point.Empty;
                        navigationStatePositions[beforeState] = beforePosition;
                    }

                    var blocked = string.Equals(
                        result,
                        "no_effect",
                        StringComparison.Ordinal);
                    var expectedPosition = blocked
                        ? beforePosition
                        : new Point(
                            beforePosition.X + delta.X,
                            beforePosition.Y + delta.Y);
                    if (!blocked &&
                        navigationStatePositions.TryGetValue(
                            afterState,
                            out var knownPosition))
                    {
                        expectedPosition = knownPosition;
                    }
                    else if (!blocked)
                    {
                        navigationStatePositions[afterState] = expectedPosition;
                    }

                    navigationEdges.Enqueue(new NavigationEdge(
                        beforeState,
                        beforePosition,
                        actionLabel,
                        afterState,
                        expectedPosition,
                        blocked));
                    while (navigationEdges.Count > 128)
                        navigationEdges.Dequeue();
                }

                static bool TryDirectionalDelta(
                    string actionLabel,
                    out Point delta)
                {
                    delta = actionLabel switch
                    {
                        "ArrowUp" or "W" => new Point(0, -1),
                        "ArrowDown" or "S" => new Point(0, 1),
                        "ArrowLeft" or "A" => new Point(-1, 0),
                        "ArrowRight" or "D" => new Point(1, 0),
                        _ => Point.Empty
                    };
                    return delta != Point.Empty;
                }

                void UpdateCausalEvent(
                    string beforeState,
                    string actionLabel,
                    string afterState,
                    ChangeAnalysis analysis)
                {
                    if (causalEvent is not null)
                    {
                        causalEvent.TransitionAge++;
                        if (!causalEvent.ReturnObserved &&
                            string.Equals(
                                beforeState,
                                causalEvent.ChangedState,
                                StringComparison.Ordinal) &&
                            string.Equals(
                                afterState,
                                causalEvent.BaselineState,
                                StringComparison.Ordinal))
                        {
                            causalEvent.ReturnObserved = true;
                            causalEvent.ReturnAction = actionLabel;
                            Console.WriteLine(
                                $"[turn-causal] return observed: {causalEvent.ChangedState} --{actionLabel}--> {causalEvent.BaselineState}.");
                        }
                        else if (causalEvent.ReturnObserved &&
                                 !causalEvent.ReenactmentObserved &&
                                 string.Equals(
                                     beforeState,
                                     causalEvent.BaselineState,
                                     StringComparison.Ordinal) &&
                                 string.Equals(
                                     afterState,
                                     causalEvent.ChangedState,
                                     StringComparison.Ordinal))
                        {
                            causalEvent.ReenactmentObserved = true;
                            causalEvent.ReenactmentAction = actionLabel;
                            Console.WriteLine(
                                $"[turn-causal] A-B-A-B reenactment confirmed by action {actionLabel}.");
                        }

                        var maximumAge = causalEvent.ReenactmentObserved ? 10 : 4;
                        if (causalEvent.TransitionAge > maximumAge)
                        {
                            Console.WriteLine("[turn-causal] retained causal event expired after sufficient later evidence.");
                            causalEvent = null;
                        }
                    }

                    if (causalEvent is null &&
                        analysis.RequiresImmediateReanalysis &&
                        analysis.HasDistantRegions &&
                        !string.Equals(beforeState, afterState, StringComparison.Ordinal))
                    {
                        causalEvent = new CausalEventEvidence(
                            beforeState,
                            actionLabel,
                            afterState,
                            FormatChangeAnalysis(analysis));
                        Console.WriteLine(
                            $"[turn-causal] retained possible local-to-distant causal event: {beforeState} --{actionLabel}--> {afterState}.");
                    }
                }

                void SetSalientChangeRegions(ChangeAnalysis analysis)
                {
                    salientChangeRegions = analysis.NovelRegions
                        .Select(region => new RectangleF(
                            region.Left / (float)StateFingerprintSide,
                            region.Top / (float)StateFingerprintSide,
                            (region.Right - region.Left) / (float)StateFingerprintSide,
                            (region.Bottom - region.Top) / (float)StateFingerprintSide))
                        .Take(3)
                        .ToArray();
                }

                ChangeAnalysis AnalyzeChanges(
                    byte[] earlier,
                    byte[] later,
                    bool learnActionPatterns,
                    string? actionLabel = null)
                {
                    var difference = StateDifference(earlier, later);
                    var tileCount = StateFingerprintSide / ChangeTileSide;
                    var activeTiles = new bool[tileCount * tileCount];
                    for (var tileY = 0; tileY < tileCount; tileY++)
                    for (var tileX = 0; tileX < tileCount; tileX++)
                    {
                        var changed = 0;
                        for (var y = 0; y < ChangeTileSide; y++)
                        for (var x = 0; x < ChangeTileSide; x++)
                        {
                            var pixelX = tileX * ChangeTileSide + x;
                            var pixelY = tileY * ChangeTileSide + y;
                            var index = pixelY * StateFingerprintSide + pixelX;
                            if (index >= volatilePixels.Length || volatilePixels[index])
                                continue;
                            if (Math.Abs(earlier[index] - later[index]) >= 16)
                                changed++;
                        }
                        activeTiles[tileY * tileCount + tileX] = changed >= 2;
                    }

                    var regions = ConnectedChangeRegions(activeTiles, tileCount);
                    IReadOnlyList<ChangeRegion> novelRegions = regions;
                    IReadOnlyList<ChangeRegion> predictableRegions = [];
                    IReadOnlyList<ChangeRegion> auxiliaryRegions = [];
                    if (learnActionPatterns && regions.Count > 0)
                    {
                        actionTransitionNumber++;
                        var classification = ClassifyAndLearnActionRegions(
                            regions,
                            actionLabel);
                        novelRegions = classification.NovelRegions;
                        predictableRegions = classification.PredictableRegions;
                        auxiliaryRegions = classification.AuxiliaryRegions;
                        LearnAuxiliaryStateRegions(auxiliaryRegions);
                        novelRegions = novelRegions
                            .Where(region => !IsKnownAuxiliaryRegion(region))
                            .ToArray();
                    }
                    var isAuxiliaryOnly =
                        difference.ChangedRatio <= 0.012 &&
                        regions.Count > 0 &&
                        regions.All(region =>
                            IsKnownAuxiliaryRegion(region) ||
                            IsLikelyPeripheralAuxiliaryRegion(region));
                    var hasDistantRegions = false;
                    for (var novelIndex = 0;
                         novelIndex < novelRegions.Count && !hasDistantRegions;
                         novelIndex++)
                    for (var otherIndex = 0; otherIndex < regions.Count; otherIndex++)
                    {
                        if (novelRegions[novelIndex].Equals(regions[otherIndex]))
                            continue;
                        if (IsKnownAuxiliaryRegion(regions[otherIndex]))
                        {
                            continue;
                        }
                        if (RegionCenterDistance(
                                novelRegions[novelIndex],
                                regions[otherIndex]) >= 0.30 &&
                            !AreExpectedDirectionalMotionPair(
                                actionLabel,
                                novelRegions[novelIndex],
                                regions[otherIndex]))
                        {
                            hasDistantRegions = true;
                            break;
                        }
                    }

                    var typicalRatio = ordinaryTransitionRatios.Count == 0
                        ? 0.0
                        : ordinaryTransitionRatios.OrderBy(value => value)
                            .ElementAt(ordinaryTransitionRatios.Count / 2);
                    var broadThreshold = Math.Max(0.0075, typicalRatio * 2.5);
                    var isBroad = novelRegions.Count > 0 &&
                                  ordinaryTransitionRatios.Count >= 2 &&
                                  difference.ChangedRatio >= broadThreshold;
                    var hasCausalDistantChange =
                        hasDistantRegions &&
                        novelRegions.Count > 0 &&
                        predictableRegions.Any(region =>
                            !IsKnownAuxiliaryRegion(region));
                    return new ChangeAnalysis(
                        difference.MeanDelta,
                        difference.ChangedRatio,
                        regions,
                        novelRegions,
                        predictableRegions.Count,
                        auxiliaryRegions.Count,
                        isAuxiliaryOnly,
                        hasDistantRegions,
                        hasCausalDistantChange,
                        isBroad);
                }

                (IReadOnlyList<ChangeRegion> NovelRegions,
                    IReadOnlyList<ChangeRegion> PredictableRegions,
                    IReadOnlyList<ChangeRegion> AuxiliaryRegions)
                    ClassifyAndLearnActionRegions(
                        IReadOnlyList<ChangeRegion> regions,
                        string? actionLabel)
                {
                    recurrentActionRegions.RemoveAll(pattern =>
                        actionTransitionNumber - pattern.LastSeenTransition >
                        RecurrentRegionMaxAge);
                    var novel = new List<ChangeRegion>();
                    var predictable = new List<ChangeRegion>();
                    var auxiliary = new List<ChangeRegion>();
                    var matchedPatterns = new HashSet<RecurrentChangeRegion>();
                    foreach (var region in regions)
                    {
                        var match = recurrentActionRegions
                            .Where(pattern =>
                                !matchedPatterns.Contains(pattern) &&
                                RegionsFollowSamePattern(pattern.Bounds, region))
                            .OrderBy(pattern => RegionCenterDistance(
                                pattern.Bounds,
                                region))
                            .FirstOrDefault();
                        if (match is null)
                        {
                            novel.Add(region);
                            recurrentActionRegions.Add(new RecurrentChangeRegion(
                                region,
                                actionTransitionNumber,
                                actionLabel));
                            if (IsKnownAuxiliaryRegion(region) ||
                                IsStrongPeripheralIndicatorCandidate(region))
                                auxiliary.Add(region);
                            continue;
                        }

                        matchedPatterns.Add(match);
                        predictable.Add(region);
                        match.Hits++;
                        if (!string.IsNullOrWhiteSpace(actionLabel))
                            match.ActionLabels.Add(actionLabel);
                        match.Bounds = region;
                        match.LastSeenTransition = actionTransitionNumber;
                        if (IsKnownAuxiliaryRegion(region) ||
                            IsStrongPeripheralIndicatorCandidate(region) ||
                            IsLikelyPeripheralAuxiliaryRegion(region) &&
                            (match.Hits >= 3 || match.ActionLabels.Count >= 2))
                        {
                            auxiliary.Add(region);
                        }
                    }

                    if (recurrentActionRegions.Count > RecurrentRegionLimit)
                    {
                        recurrentActionRegions.RemoveRange(
                            0,
                            recurrentActionRegions.Count - RecurrentRegionLimit);
                    }
                    return (novel, predictable, auxiliary);
                }

                static bool RegionsFollowSamePattern(
                    ChangeRegion previous,
                    ChangeRegion current)
                {
                    var previousWidth = Math.Max(1, previous.Right - previous.Left);
                    var previousHeight = Math.Max(1, previous.Bottom - previous.Top);
                    var currentWidth = Math.Max(1, current.Right - current.Left);
                    var currentHeight = Math.Max(1, current.Bottom - current.Top);
                    var widthSimilarity = Math.Min(previousWidth, currentWidth) /
                                          (double)Math.Max(previousWidth, currentWidth);
                    var heightSimilarity = Math.Min(previousHeight, currentHeight) /
                                           (double)Math.Max(previousHeight, currentHeight);
                    var intersectionWidth = Math.Max(
                        0,
                        Math.Min(previous.Right, current.Right) -
                        Math.Max(previous.Left, current.Left));
                    var intersectionHeight = Math.Max(
                        0,
                        Math.Min(previous.Bottom, current.Bottom) -
                        Math.Max(previous.Top, current.Top));
                    var intersection = intersectionWidth * intersectionHeight;
                    var union = previousWidth * previousHeight +
                                currentWidth * currentHeight - intersection;
                    var intersectionOverUnion = union <= 0
                        ? 0
                        : intersection / (double)union;
                    var centerDistance = RegionCenterDistance(previous, current);
                    var verticalOverlap = intersectionHeight /
                                          (double)Math.Min(previousHeight, currentHeight);
                    return intersectionOverUnion >= 0.15 ||
                           widthSimilarity >= 0.50 &&
                           heightSimilarity >= 0.50 &&
                           centerDistance <= 0.12 ||
                           verticalOverlap >= 0.60 &&
                           heightSimilarity >= 0.50 &&
                           centerDistance <= 0.12;
                }

                static double RegionCenterDistance(
                    ChangeRegion left,
                    ChangeRegion right)
                {
                    var leftCenterX = (left.Left + left.Right) / 2.0;
                    var leftCenterY = (left.Top + left.Bottom) / 2.0;
                    var rightCenterX = (right.Left + right.Right) / 2.0;
                    var rightCenterY = (right.Top + right.Bottom) / 2.0;
                    return Math.Sqrt(
                        Math.Pow(
                            (leftCenterX - rightCenterX) /
                            StateFingerprintSide,
                            2) +
                        Math.Pow(
                            (leftCenterY - rightCenterY) /
                            StateFingerprintSide,
                            2));
                }

                static bool AreExpectedDirectionalMotionPair(
                    string? actionLabel,
                    ChangeRegion left,
                    ChangeRegion right)
                {
                    if (actionLabel is null ||
                        !IsCanonicalDirectionalInput(actionLabel))
                    {
                        return false;
                    }

                    var deltaX = Math.Abs(
                        (left.Left + left.Right) -
                        (right.Left + right.Right)) / 2.0;
                    var deltaY = Math.Abs(
                        (left.Top + left.Bottom) -
                        (right.Top + right.Bottom)) / 2.0;
                    var vertical = actionLabel is "ArrowUp" or "ArrowDown" or "W" or "S";
                    return vertical
                        ? deltaY >= ChangeTileSide && deltaY > deltaX * 1.5
                        : deltaX >= ChangeTileSide && deltaX > deltaY * 1.5;
                }

                static bool HasSpatiallySeparatedChangeRegions(
                    IReadOnlyList<ChangeRegion> regions)
                {
                    for (var first = 0; first < regions.Count; first++)
                    for (var second = first + 1; second < regions.Count; second++)
                    {
                        if (RegionCenterDistance(regions[first], regions[second]) >= 0.30)
                            return true;
                    }
                    return false;
                }

                static bool HasPersistentVisualChange(ChangeAnalysis analysis) =>
                    analysis.Regions.Count > 0 ||
                    analysis.ChangedRatio >= 0.00025 ||
                    analysis.MeanDelta >= 0.00005;

                bool HasDirectionalMotionEvidence(
                    string actionLabel,
                    ChangeAnalysis analysis)
                {
                    if (!IsCanonicalDirectionalInput(actionLabel))
                        return true;
                    if (analysis.IsAuxiliaryOnly)
                        return false;

                    var vertical = actionLabel is "ArrowUp" or "ArrowDown" or "W" or "S";
                    foreach (var region in analysis.Regions)
                    {
                        var width = Math.Max(1, region.Right - region.Left);
                        var height = Math.Max(1, region.Bottom - region.Top);
                        if (vertical &&
                            height >= ChangeTileSide * 2 &&
                            height >= width * 1.20 ||
                            !vertical &&
                            width >= ChangeTileSide * 2 &&
                            width >= height * 1.20)
                        {
                            return true;
                        }
                    }

                    for (var first = 0; first < analysis.Regions.Count; first++)
                    for (var second = first + 1; second < analysis.Regions.Count; second++)
                    {
                        var left = analysis.Regions[first];
                        var right = analysis.Regions[second];
                        var deltaX = Math.Abs(
                            (left.Left + left.Right) -
                            (right.Left + right.Right)) / 2.0;
                        var deltaY = Math.Abs(
                            (left.Top + left.Bottom) -
                            (right.Top + right.Bottom)) / 2.0;
                        if (vertical && deltaY >= ChangeTileSide && deltaY > deltaX * 1.5 ||
                            !vertical && deltaX >= ChangeTileSide && deltaX > deltaY * 1.5)
                        {
                            return true;
                        }
                    }

                    var typicalRatio = TypicalDirectionalMotionRatio();
                    var adaptiveThreshold = typicalRatio <= 0
                        ? MinimumDirectionalMotionChangedRatio
                        : Math.Max(
                            MinimumDirectionalMotionChangedRatio,
                            typicalRatio * 0.35);
                    return analysis.ChangedRatio >= adaptiveThreshold;
                }

                double TypicalDirectionalMotionRatio() =>
                    ordinaryTransitionRatios.Count == 0
                        ? 0.0
                        : ordinaryTransitionRatios
                            .OrderBy(value => value)
                            .ElementAt(ordinaryTransitionRatios.Count / 2);

                static bool IsLikelyPeripheralAuxiliaryRegion(ChangeRegion region)
                {
                    var width = Math.Max(1, region.Right - region.Left);
                    var height = Math.Max(1, region.Bottom - region.Top);
                    var outerBand = StateFingerprintSide / 8;
                    var nearOuterEdge =
                        region.Bottom <= outerBand ||
                        region.Top >= StateFingerprintSide - outerBand ||
                        region.Right <= outerBand ||
                        region.Left >= StateFingerprintSide - outerBand;
                    var compact =
                        width * height <= StateFingerprintSide * StateFingerprintSide * 0.025 &&
                        Math.Min(width, height) <= StateFingerprintSide / 8;
                    return nearOuterEdge && compact;
                }

                static bool IsStrongPeripheralIndicatorCandidate(
                    ChangeRegion region)
                {
                    if (!IsLikelyPeripheralAuxiliaryRegion(region))
                        return false;
                    var width = Math.Max(1, region.Right - region.Left);
                    var height = Math.Max(1, region.Bottom - region.Top);
                    return Math.Max(width, height) >= Math.Min(width, height) * 1.75;
                }

                void LearnAuxiliaryStateRegions(
                    IReadOnlyList<ChangeRegion> regions)
                {
                    EnsureAuxiliaryStateMask();
                    const int dilation = ChangeTileSide;
                    foreach (var region in regions)
                    {
                        var left = Math.Max(0, region.Left - dilation);
                        var top = Math.Max(0, region.Top - dilation);
                        var right = Math.Min(
                            StateFingerprintSide,
                            region.Right + dilation);
                        var bottom = Math.Min(
                            StateFingerprintSide,
                            region.Bottom + dilation);
                        for (var y = top; y < bottom; y++)
                        for (var x = left; x < right; x++)
                            auxiliaryStatePixels[y * StateFingerprintSide + x] = true;
                    }
                }

                bool IsKnownAuxiliaryRegion(ChangeRegion region)
                {
                    if (auxiliaryStatePixels.Length !=
                        StateFingerprintSide * StateFingerprintSide)
                    {
                        return false;
                    }

                    var covered = 0;
                    var area = 0;
                    for (var y = Math.Max(0, region.Top);
                         y < Math.Min(StateFingerprintSide, region.Bottom);
                         y++)
                    for (var x = Math.Max(0, region.Left);
                         x < Math.Min(StateFingerprintSide, region.Right);
                         x++)
                    {
                        area++;
                        if (auxiliaryStatePixels[y * StateFingerprintSide + x])
                            covered++;
                    }
                    return area > 0 && covered / (double)area >= 0.35;
                }

                static bool IsPeripheralOnlyChange(
                    IReadOnlyList<ChangeRegion> regions)
                {
                    if (regions.Count == 0)
                        return false;

                    const int topChromeBoundary = StateFingerprintSide / 8;
                    const int edgeBoundary = 2;
                    return regions.All(region =>
                        region.Bottom <= topChromeBoundary ||
                        region.Top >= StateFingerprintSide - edgeBoundary ||
                        region.Right <= edgeBoundary ||
                        region.Left >= StateFingerprintSide - edgeBoundary);
                }

                static List<ChangeRegion> ConnectedChangeRegions(
                    bool[] activeTiles,
                    int tileCount)
                {
                    var regions = new List<ChangeRegion>();
                    var visited = new bool[activeTiles.Length];
                    for (var start = 0; start < activeTiles.Length; start++)
                    {
                        if (!activeTiles[start] || visited[start])
                            continue;
                        var queue = new Queue<int>();
                        queue.Enqueue(start);
                        visited[start] = true;
                        var minX = tileCount;
                        var minY = tileCount;
                        var maxX = 0;
                        var maxY = 0;
                        var changedTiles = 0;
                        while (queue.Count > 0)
                        {
                            var current = queue.Dequeue();
                            var x = current % tileCount;
                            var y = current / tileCount;
                            minX = Math.Min(minX, x);
                            minY = Math.Min(minY, y);
                            maxX = Math.Max(maxX, x);
                            maxY = Math.Max(maxY, y);
                            changedTiles++;
                            for (var dy = -1; dy <= 1; dy++)
                            for (var dx = -1; dx <= 1; dx++)
                            {
                                if (dx == 0 && dy == 0)
                                    continue;
                                var nx = x + dx;
                                var ny = y + dy;
                                if (nx < 0 || ny < 0 || nx >= tileCount || ny >= tileCount)
                                    continue;
                                var neighbor = ny * tileCount + nx;
                                if (!activeTiles[neighbor] || visited[neighbor])
                                    continue;
                                visited[neighbor] = true;
                                queue.Enqueue(neighbor);
                            }
                        }
                        regions.Add(new ChangeRegion(
                            minX * ChangeTileSide,
                            minY * ChangeTileSide,
                            (maxX + 1) * ChangeTileSide,
                            (maxY + 1) * ChangeTileSide,
                            changedTiles * ChangeTileSide * ChangeTileSide));
                    }
                    return regions
                        .OrderByDescending(region => region.ChangedPixels)
                        .Take(6)
                        .ToList();
                }

                static string FormatChangeAnalysis(ChangeAnalysis analysis)
                {
                    var regions = analysis.Regions.Count == 0
                        ? "none"
                        : string.Join(", ", analysis.Regions.Select(region =>
                            $"[{Percent(region.Left)}..{Percent(region.Right)}% x, {Percent(region.Top)}..{Percent(region.Bottom)}% y]"));
                    return
                        $"changed_ratio={analysis.ChangedRatio:0.####}; mean_delta={analysis.MeanDelta:0.####}; regions={regions}; novel_regions={analysis.NovelRegions.Count}; predictable_regions={analysis.PredictableRegionCount}; auxiliary_regions={analysis.AuxiliaryRegionCount}; auxiliary_only={analysis.IsAuxiliaryOnly.ToString().ToLowerInvariant()}; distant={analysis.HasDistantRegions.ToString().ToLowerInvariant()}; causal_distant={analysis.HasCausalDistantChange.ToString().ToLowerInvariant()}; broad={analysis.IsBroad.ToString().ToLowerInvariant()}";
                }

                static int Percent(int coordinate) =>
                    Math.Clamp((int)Math.Round(coordinate * 100.0 / StateFingerprintSide), 0, 100);

                static string? NormalizeWorkingMemory(string? value)
                {
                    if (string.IsNullOrWhiteSpace(value))
                        return null;
                    return TrimForMeta(value.Trim(), 280);
                }

                void EnsureVolatileMask(int length)
                {
                    if (volatilePixels.Length != length)
                        volatilePixels = new bool[length];
                    EnsureAuxiliaryStateMask();
                }

                void EnsureAuxiliaryStateMask()
                {
                    var length = StateFingerprintSide * StateFingerprintSide;
                    if (auxiliaryStatePixels.Length != length)
                        auxiliaryStatePixels = new bool[length];
                }

                void LearnAmbientPixels(byte[] earlier, byte[] later)
                {
                    var count = Math.Min(
                        volatilePixels.Length,
                        Math.Min(earlier.Length, later.Length));
                    var changedPixels = new List<int>();
                    for (var index = 0; index < count; index++)
                    {
                        if (Math.Abs(earlier[index] - later[index]) >= 8)
                            changedPixels.Add(index);
                    }

                    const int ambientDilationRadius = 2;
                    foreach (var changedIndex in changedPixels)
                    {
                        var centerX = changedIndex % StateFingerprintSide;
                        var centerY = changedIndex / StateFingerprintSide;
                        for (var dy = -ambientDilationRadius; dy <= ambientDilationRadius; dy++)
                        for (var dx = -ambientDilationRadius; dx <= ambientDilationRadius; dx++)
                        {
                            var x = centerX + dx;
                            var y = centerY + dy;
                            if (x < 0 || y < 0 ||
                                x >= StateFingerprintSide ||
                                y >= StateFingerprintSide)
                            {
                                continue;
                            }
                            var neighbor = y * StateFingerprintSide + x;
                            if (neighbor < volatilePixels.Length)
                                volatilePixels[neighbor] = true;
                        }
                    }
                }

                (double MeanDelta, double ChangedRatio) StateDifference(
                    byte[] left,
                    byte[] right)
                {
                    var count = Math.Min(
                        volatilePixels.Length,
                        Math.Min(left.Length, right.Length));
                    if (count == 0)
                        return (1, 1);

                    double sum = 0;
                    var changed = 0;
                    var compared = 0;
                    for (var index = 0; index < count; index++)
                    {
                        if (volatilePixels[index])
                            continue;
                        var difference = Math.Abs(left[index] - right[index]);
                        sum += difference / 255.0;
                        if (difference >= 8)
                            changed++;
                        compared++;
                    }
                    return compared == 0
                        ? (1, 1)
                        : (sum / compared, changed / (double)compared);
                }

                (double MeanDelta, double ChangedRatio) StateIdentityDifference(
                    byte[] left,
                    byte[] right)
                {
                    var count = Math.Min(
                        volatilePixels.Length,
                        Math.Min(left.Length, right.Length));
                    if (count == 0)
                        return (1, 1);

                    double sum = 0;
                    var changed = 0;
                    var compared = 0;
                    for (var index = 0; index < count; index++)
                    {
                        if (volatilePixels[index] ||
                            index < auxiliaryStatePixels.Length &&
                            auxiliaryStatePixels[index])
                        {
                            continue;
                        }

                        var difference = Math.Abs(left[index] - right[index]);
                        sum += difference / 255.0;
                        if (difference >= 8)
                            changed++;
                        compared++;
                    }
                    return compared == 0
                        ? (1, 1)
                        : (sum / compared, changed / (double)compared);
                }

                static bool IsSameState(
                    (double MeanDelta, double ChangedRatio) difference,
                    double meanThreshold = StateMatchMeanThreshold,
                    double changedRatioThreshold = StateMatchChangedRatioThreshold) =>
                    difference.MeanDelta <= meanThreshold &&
                    difference.ChangedRatio <= changedRatioThreshold;

                static byte[] ExtractRegionFingerprint(
                    ScreenObservationFrame frame,
                    Rectangle region)
                {
                    if (frame.DetailWidth <= 0 ||
                        frame.DetailHeight <= 0 ||
                        frame.DetailFingerprint.Length != frame.DetailWidth * frame.DetailHeight)
                    {
                        return [];
                    }

                    region.Intersect(frame.ScreenBounds);
                    if (region.Width <= 0 || region.Height <= 0)
                        return [];

                    var output = new byte[StateFingerprintSide * StateFingerprintSide];
                    for (var y = 0; y < StateFingerprintSide; y++)
                    {
                        var screenY = region.Top +
                            (y + 0.5) * region.Height / StateFingerprintSide;
                        var detailY = Math.Clamp(
                            (int)((screenY - frame.ScreenBounds.Top) * frame.DetailHeight /
                                  Math.Max(1.0, frame.ScreenBounds.Height)),
                            0,
                            frame.DetailHeight - 1);
                        for (var x = 0; x < StateFingerprintSide; x++)
                        {
                            var screenX = region.Left +
                                (x + 0.5) * region.Width / StateFingerprintSide;
                            var detailX = Math.Clamp(
                                (int)((screenX - frame.ScreenBounds.Left) * frame.DetailWidth /
                                      Math.Max(1.0, frame.ScreenBounds.Width)),
                                0,
                                frame.DetailWidth - 1);
                            var value = frame.DetailFingerprint[
                                detailY * frame.DetailWidth + detailX];
                            output[y * StateFingerprintSide + x] =
                                (byte)(value / 8 * 8);
                        }
                    }
                    return output;
                }

                string TurnActionLabel(ResolvedActionSnapshot snapshot)
                {
                    var action = snapshot.Action;
                    if (OpenAiResponsesService.TryNormalizeDirectionalLabel(
                            action.ResolvedTurnInputLabel,
                            out var resolvedTurnInputLabel))
                    {
                        return resolvedTurnInputLabel;
                    }
                    if (action.Type == "keys" && action.Keys is { Length: > 0 })
                        return string.Join(",", action.Keys.Select(CanonicalKeyLabel));
                    if (action.Type is "click" or "double_click" or "click_uia" or "focus_uia")
                    {
                        if (TryResolveDirectionalClickTemplate(action, out var templateLabel))
                            return templateLabel;
                        var target = string.IsNullOrWhiteSpace(action.Note)
                            ? snapshot.SemanticTokens
                            : action.Note;
                        if (OpenAiResponsesService.TryNormalizeUnambiguousDirectionalLabel(
                                target,
                                out var direction))
                        {
                            return direction;
                        }
                        return string.IsNullOrWhiteSpace(target)
                            ? action.Type
                            : $"{action.Type}:{TrimForMeta(target, 48)}";
                    }
                    return action.Type;
                }

                bool TryResolveDirectionalClickTemplate(
                    ActionDto action,
                    out string label)
                {
                    label = "";
                    if (action.Type is not ("click" or "double_click") ||
                        !TryGetActionImagePoint(action, out var point))
                    {
                        return false;
                    }

                    const int maximumDistance = 8;
                    var nearest = directionalClickTemplates
                        .Select(entry => new
                        {
                            entry.Key,
                            Point = TryGetActionImagePoint(entry.Value, out var templatePoint)
                                ? templatePoint
                                : ((int X, int Y)?)null
                        })
                        .Where(entry => entry.Point is not null)
                        .Select(entry => new
                        {
                            entry.Key,
                            DistanceSquared =
                                (entry.Point!.Value.X - point.X) *
                                (entry.Point.Value.X - point.X) +
                                (entry.Point.Value.Y - point.Y) *
                                (entry.Point.Value.Y - point.Y)
                        })
                        .Where(entry =>
                            entry.DistanceSquared <= maximumDistance * maximumDistance)
                        .OrderBy(entry => entry.DistanceSquared)
                        .Take(2)
                        .ToArray();
                    if (nearest.Length == 0 ||
                        nearest.Length > 1 &&
                        nearest[0].DistanceSquared == nearest[1].DistanceSquared)
                    {
                        return false;
                    }

                    label = nearest[0].Key;
                    action.ResolvedTurnInputLabel = label;
                    return true;
                }

                static bool TryGetActionImagePoint(
                    ActionDto action,
                    out (int X, int Y) point)
                {
                    if (action.XPx is int x && action.YPx is int y)
                    {
                        point = (x, y);
                        return true;
                    }

                    if (action.BBox is
                        {
                            Left: int left,
                            Top: int top,
                            Right: int right,
                            Bottom: int bottom
                        })
                    {
                        point = ((left + right) / 2, (top + bottom) / 2);
                        return true;
                    }

                    point = default;
                    return false;
                }

                static string CanonicalKeyLabel(string key) =>
                    CanonicalDirectionLabel(key) ?? key.Trim();

                static string? CanonicalDirectionLabel(string? value)
                {
                    return OpenAiResponsesService.TryNormalizeDirectionalLabel(
                        value,
                        out var label)
                        ? label
                        : null;
                }

                static bool IsCanonicalDirectionalInput(string label) =>
                    label.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .All(item => item is "ArrowRight" or "ArrowLeft" or "ArrowUp" or "ArrowDown" or
                            "W" or "A" or "S" or "D" or "w" or "a" or "s" or "d");
            }

            internal static bool ActionNeedsEffectObservation(ActionDto action) =>
                action.Type is not ("aim" or "point" or "request_crop" or "move" or "done");

            internal static bool ShouldCheckPreRegionTurnActionFreshness(
                string actionPolicy,
                Rectangle? turnBasedInteractionRect,
                ActionDto action) =>
                string.Equals(
                    actionPolicy,
                    "turn_based_interaction",
                    StringComparison.Ordinal) &&
                turnBasedInteractionRect is null &&
                IsStateChangingInteractionAction(action) &&
                (IsSpatialPointerAction(action) ||
                 action.Type is "drag_drop" or "drag_path" or "click_uia" or "focus_uia");

            internal static bool ShouldDiscardPreRegionTurnAction(
                ScreenObservationFrame promptFrame,
                ScreenObservationFrame immediateFrame,
                ScreenObservationFrame confirmationFrame,
                out double promptDelta,
                out double stabilityDelta)
            {
                promptDelta = Math.Max(
                    ComputeImageDelta(
                        promptFrame.GlobalFingerprint,
                        immediateFrame.GlobalFingerprint),
                    ComputeImageDelta(
                        promptFrame.ActiveWindowFingerprint,
                        immediateFrame.ActiveWindowFingerprint));
                stabilityDelta = Math.Max(
                    ComputeImageDelta(
                        immediateFrame.GlobalFingerprint,
                        confirmationFrame.GlobalFingerprint),
                    ComputeImageDelta(
                        immediateFrame.ActiveWindowFingerprint,
                        confirmationFrame.ActiveWindowFingerprint));
                return promptDelta >= 0.0035 &&
                       stabilityDelta <= Math.Max(0.0015, promptDelta * 0.35);
            }

            internal static bool IsExpectedContinuousIdle(
                string goalMode,
                ActionDto? previousAction,
                bool noChange) =>
                noChange &&
                string.Equals(goalMode, "continuous", StringComparison.Ordinal) &&
                previousAction?.Type == "wait";

            internal static bool ShouldRegisterImmediateNoEffectCooldown(
                ActionDto action,
                ObservationAssessment assessment,
                int repeatCount) =>
                assessment.ActionOutcome == ActionOutcomeState.NoEffect &&
                ActionRepeatCooldownSteps > 0 &&
                action.Type is not ("wait" or "done") &&
                !IsLocalObservationAction(action) &&
                !IsTextInputAttemptAction(action) &&
                (!string.Equals(
                     assessment.ActionPolicy,
                     "turn_based_interaction",
                     StringComparison.Ordinal) ||
                 !IsStateChangingInteractionAction(action) ||
                 repeatCount > 0);

            internal static bool ShouldResetRejectedProposalLoop(
                ResolvedActionSnapshot? previousAction,
                bool noChange,
                bool expectedContinuousIdle) =>
                previousAction is not null &&
                !noChange &&
                !expectedContinuousIdle &&
                !IsLocalObservationAction(previousAction.Action);

            internal static bool IsSpatialPointerAction(ActionDto action)
                => action.Type is "move" or "click" or "double_click";

            internal static (int RepeatCount, string? LastSignature) UpdateRepeatDetection(
                ResolvedActionSnapshot previousAction,
                bool noChange,
                int repeatCount,
                string? lastSignature,
                List<ResolvedActionSnapshot> recentIneffectiveSpatialActions)
            {
                var family = RecoveryMemoryService.ActionFamily(previousAction.Action);
                var isSpatial = IsSpatialPointerAction(previousAction.Action) ||
                                family == "drag_drop";
                if (previousAction.ScreenPoint is Point pointerPoint && isSpatial)
                {
                    if (noChange)
                    {
                        var repeatsNearbyPoint = recentIneffectiveSpatialActions.Any(prior =>
                        {
                            if (!string.Equals(
                                    RecoveryMemoryService.ActionFamily(prior.Action),
                                    family,
                                    StringComparison.OrdinalIgnoreCase) ||
                                prior.ScreenPoint is not Point priorPoint ||
                                !ScreenPointsAreNearby(priorPoint, pointerPoint))
                            {
                                return false;
                            }

                            if (family != "drag_drop")
                                return true;
                            return prior.DestinationScreenPoint is Point priorDestination &&
                                   previousAction.DestinationScreenPoint is Point destination &&
                                   ScreenPointsAreNearby(priorDestination, destination);
                        });
                        repeatCount = repeatsNearbyPoint ? repeatCount + 1 : 0;
                        recentIneffectiveSpatialActions.Add(previousAction);
                        if (recentIneffectiveSpatialActions.Count > 16)
                            recentIneffectiveSpatialActions.RemoveAt(0);
                    }
                    else
                    {
                        repeatCount = 0;
                        recentIneffectiveSpatialActions.Clear();
                    }

                    return (repeatCount, null);
                }

                recentIneffectiveSpatialActions.Clear();
                if (!noChange)
                    return (0, null);
                var signature = previousAction.IneffectiveSignature;
                return (
                    string.Equals(signature, lastSignature, StringComparison.Ordinal)
                        ? repeatCount + 1
                        : 0,
                    signature);
            }

            internal static int RepeatedStringCycleLength(
                IReadOnlyList<string> values)
            {
                for (var period = 1;
                     period <= Math.Min(8, values.Count / 2);
                     period++)
                {
                    var start = values.Count - period * 2;
                    var matches = true;
                    for (var index = 0; index < period; index++)
                    {
                        if (string.Equals(
                                values[start + index],
                                values[start + period + index],
                                StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        matches = false;
                        break;
                    }

                    if (matches)
                        return period;
                }

                return 0;
            }

            internal static bool ScreenPointsAreNearby(Point a, Point b)
            {
                var radius = Math.Max(16, IneffectiveMouseClusterPx);
                var dx = (long)a.X - b.X;
                var dy = (long)a.Y - b.Y;
                return dx * dx + dy * dy <= (long)radius * radius;
            }

            internal static bool IsActionOnCooldown(
                ResolvedActionSnapshot action,
                int step,
                Dictionary<string, int> cooldowns,
                List<SpatialActionCooldown> spatialCooldowns,
                out int untilStep)
            {
                untilStep = 0;
                if (ActionRepeatCooldownSteps <= 0 || IsLocalObservationAction(action.Action))
                    return false;

                for (var i = spatialCooldowns.Count - 1; i >= 0; i--)
                {
                    if (spatialCooldowns[i].UntilStep < step)
                        spatialCooldowns.RemoveAt(i);
                }
                foreach (var expired in cooldowns
                             .Where(item => item.Value < step)
                             .Select(item => item.Key)
                             .ToArray())
                {
                    cooldowns.Remove(expired);
                }

                if (cooldowns.TryGetValue(action.IneffectiveSignature, out untilStep))
                {
                    if (step <= untilStep)
                        return true;
                    cooldowns.Remove(action.IneffectiveSignature);
                }

                var family = RecoveryMemoryService.ActionFamily(action.Action);
                if (action.ScreenPoint is Point screenPoint &&
                    (IsSpatialPointerAction(action.Action) || family == "drag_drop"))
                {
                    foreach (var cooldown in spatialCooldowns)
                    {
                        if (!string.Equals(cooldown.ActionFamily, family, StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (!ScreenPointsAreNearby(cooldown.ScreenPoint, screenPoint))
                            continue;
                        if (family == "drag_drop" &&
                            (cooldown.DestinationScreenPoint is not Point priorDestination ||
                             action.DestinationScreenPoint is not Point destination ||
                             !ScreenPointsAreNearby(priorDestination, destination)))
                        {
                            continue;
                        }

                        untilStep = Math.Max(untilStep, cooldown.UntilStep);
                    }

                    return untilStep >= step;
                }

                return false;
            }

            internal static void RegisterActionCooldown(
                ResolvedActionSnapshot action,
                int untilStep,
                Dictionary<string, int> cooldowns,
                List<SpatialActionCooldown> spatialCooldowns,
                bool clusterSpatially = true)
            {
                var family = RecoveryMemoryService.ActionFamily(action.Action);
                if (clusterSpatially &&
                    action.ScreenPoint is Point screenPoint &&
                    (IsSpatialPointerAction(action.Action) || family == "drag_drop"))
                {
                    spatialCooldowns.Add(new SpatialActionCooldown(
                        screenPoint,
                        action.DestinationScreenPoint,
                        family,
                        untilStep));
                    if (spatialCooldowns.Count >
                        Math.Max(16, RuntimeCooldownEntryLimit))
                    {
                        spatialCooldowns.RemoveRange(
                            0,
                            spatialCooldowns.Count -
                            Math.Max(16, RuntimeCooldownEntryLimit));
                    }
                    return;
                }

                cooldowns[action.IneffectiveSignature] = untilStep;
                if (cooldowns.Count >
                    Math.Max(16, RuntimeCooldownEntryLimit))
                {
                    foreach (var key in cooldowns
                                 .OrderBy(item => item.Value)
                                 .Take(
                                     cooldowns.Count -
                                     Math.Max(16, RuntimeCooldownEntryLimit))
                                 .Select(item => item.Key)
                                 .ToArray())
                    {
                        cooldowns.Remove(key);
                    }
                }
            }
        
            internal static bool IsPointerPositionOnlyAction(ActionDto a)
                => a.Type == "move";
        
            internal static int PostActionDelay(ActionDto a)
            {
                // These actions only adjust the next prompt/crop; the target UI has not changed.
                if (a.Type == "wait" || IsLocalObservationAction(a))
                    return 0;
        
                // A plain cursor move can affect hover state, but does not need the full app settle delay.
                if (IsPointerPositionOnlyAction(a))
                    return Math.Max(0, DelayFor(a));
        
                return Math.Max(UiSettleDelayMs, DelayFor(a));
            }
        
            internal static async Task<bool> WaitAfterActionAsync(
                ActionDto action,
                byte[] beforeFingerprint,
                string observationPolicy,
                byte[]? beforeLocalFingerprint,
                Rectangle? localRegion,
                CancellationToken cancellationToken,
                bool preferFastTurnSettle = false)
            {
                if (IsLocalObservationAction(action))
                    return false;
        
                if (!ScreenPollingEnabled)
                {
                    var fixedDelay = PostActionDelay(action);
                    if (fixedDelay > 0)
                        await Task.Delay(fixedDelay, cancellationToken);
                    return false;
                }
        
                if (action.Type == "wait")
                {
                    if (WaitNoChangeExtraMs <= 0)
                        return false;
        
                    var afterWait = CaptureScreenFingerprintProbe();
                    var delta = ComputeImageDelta(beforeFingerprint, afterWait);
                    if (delta < SettleThresholdFor(action))
                    {
                        Console.WriteLine($"[settle] screen unchanged after wait (coarse_delta={delta:0.####}); waiting extra {WaitNoChangeExtraMs} ms.");
                        await Task.Delay(WaitNoChangeExtraMs, cancellationToken);
                    }
                    return false;
                }
        
                var initialDelay = Math.Min(PostActionDelay(action), Math.Max(0, ScreenPollInitialDelayMs));
                if (initialDelay > 0)
                    await Task.Delay(initialDelay, cancellationToken);
        
                if (ScreenPollTimeoutMs <= 0)
                    return false;

                if (string.Equals(
                        observationPolicy,
                        "turn_based_interaction",
                        StringComparison.Ordinal) &&
                    beforeLocalFingerprint is { Length: > 0 } &&
                    localRegion is Rectangle turnRegion)
                {
                    return await WaitForLocalScreenReactionAndStableAsync(
                        beforeLocalFingerprint,
                        turnRegion,
                        action,
                        cancellationToken,
                        preferFastTurnSettle);
                }

                if (observationPolicy == "realtime_interaction" ||
                    action.Type == "hold_keys" ||
                    action.Type == "drag_path" && action.GestureKind is "game" or "pan")
                    return false;

                if (observationPolicy is "event_driven" or "streaming_output")
                {
                    await WaitForScreenReactionAsync(
                        beforeFingerprint,
                        action,
                        cancellationToken);
                    return false;
                }

                await WaitForScreenStableAsync(beforeFingerprint, action, cancellationToken);
                return false;
            }

            internal static async Task<bool> WaitForLocalScreenReactionAndStableAsync(
                byte[] beforeFingerprint,
                Rectangle region,
                ActionDto action,
                CancellationToken cancellationToken,
                bool preferFastSettle = false)
            {
                const double reactionThreshold = 0.0012;
                const double reactionRatioThreshold = 0.0015;
                const double stableThreshold = 0.0007;
                const double stableRatioThreshold = 0.0008;
                var requiredStableProbes = preferFastSettle ? 1 : 2;
                var pollIntervalMs = preferFastSettle
                    ? Math.Min(ScreenPollIntervalMs, 75)
                    : ScreenPollIntervalMs;
                var timeoutMs = preferFastSettle
                    ? Math.Min(ScreenPollTimeoutMs, 650)
                    : ScreenPollTimeoutMs;
                var settleMode = preferFastSettle ? "fast" : "adaptive";

                var sw = Stopwatch.StartNew();
                byte[]? previous = null;
                var sawReaction = false;
                var stableProbes = 0;
                var probes = 0;
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
                double lastFromBeforeRatio = double.NaN;
                double lastBetweenRatio = double.NaN;
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var current = CaptureRegionFingerprintProbe(region);
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    lastFromBeforeRatio = ComputeChangedPixelRatio(
                        beforeFingerprint,
                        current);
                    if (lastFromBefore >= reactionThreshold ||
                        lastFromBeforeRatio >= reactionRatioThreshold)
                        sawReaction = true;

                    if (sawReaction && previous is not null)
                    {
                        lastBetween = ComputeImageDelta(previous, current);
                        lastBetweenRatio = ComputeChangedPixelRatio(previous, current);
                        stableProbes = lastBetween < stableThreshold &&
                                       lastBetweenRatio < stableRatioThreshold
                            ? stableProbes + 1
                            : 0;
                        if (stableProbes >= requiredStableProbes)
                        {
                            Console.WriteLine(
                                $"[settle] local reaction stable after {sw.ElapsedMilliseconds} ms; probes={probes}; local_delta={lastFromBefore:0.####}; local_ratio={lastFromBeforeRatio:0.####}; action={action.Type}; mode={settleMode}");
                            return true;
                        }
                    }

                    previous = current;
                    var remaining = timeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await Task.Delay(
                        Math.Min(pollIntervalMs, remaining),
                        cancellationToken);
                }

                Console.WriteLine(
                    sawReaction
                        ? $"[settle] local reaction reached timeout after {sw.ElapsedMilliseconds} ms; probes={probes}; local_delta={lastFromBefore:0.####}; local_ratio={lastFromBeforeRatio:0.####}; local_between_delta={lastBetween:0.####}; local_between_ratio={lastBetweenRatio:0.####}; action={action.Type}; mode={settleMode}"
                        : $"[settle] no local reaction within {sw.ElapsedMilliseconds} ms; probes={probes}; local_delta={lastFromBefore:0.####}; local_ratio={lastFromBeforeRatio:0.####}; action={action.Type}; mode={settleMode}");
                return sawReaction;
            }

            internal static async Task WaitForScreenReactionAsync(
                byte[] beforeFingerprint,
                ActionDto action,
                CancellationToken cancellationToken)
            {
                var sw = Stopwatch.StartNew();
                var probes = 0;
                double delta = double.NaN;
                var threshold = SettleThresholdFor(action);
                while (sw.ElapsedMilliseconds < ScreenPollTimeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    delta = ComputeImageDelta(
                        beforeFingerprint,
                        CaptureScreenFingerprintProbe());
                    probes++;
                    if (delta >= threshold)
                    {
                        Console.WriteLine(
                            $"[settle] reaction started after {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={delta:0.####}; action={action.Type}");
                        return;
                    }
                    var remaining = ScreenPollTimeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await Task.Delay(
                        Math.Min(ScreenPollIntervalMs, remaining),
                        cancellationToken);
                }
                Console.WriteLine(
                    $"[settle] no visible reaction within {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={delta:0.####}; action={action.Type}");
            }
        
            internal static async Task WaitForScreenStableAsync(byte[] beforeFingerprint, ActionDto action, CancellationToken cancellationToken)
            {
                var sw = Stopwatch.StartNew();
                byte[]? previous = null;
                var sawChange = false;
                var probes = 0;
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
                var settleThreshold = SettleThresholdFor(action);
        
                while (sw.ElapsedMilliseconds < ScreenPollTimeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var current = CaptureScreenFingerprintProbe();
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    if (lastFromBefore >= settleThreshold)
                        sawChange = true;
        
                    if (previous != null)
                    {
                        lastBetween = ComputeImageDelta(previous, current);
                        if (lastBetween < settleThreshold && (sawChange || probes >= 2))
                        {
                            Console.WriteLine($"[settle] stable after {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={lastFromBefore:0.####}; action={action.Type}");
                            return;
                        }
                    }
        
                    previous = current;
                    var remaining = ScreenPollTimeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await Task.Delay(Math.Min(ScreenPollIntervalMs, remaining), cancellationToken);
                }
        
                if (probes > 0)
                    Console.WriteLine($"[settle] timeout after {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={lastFromBefore:0.####}; coarse_between_delta={lastBetween:0.####}; action={action.Type}");
            }
        
            internal static byte[] CaptureScreenFingerprintProbe()
            {
                var sw = Stopwatch.StartNew();
                try
                {
                    var (vx, vy, vw, vh) = GetPrimaryScreen();
                    using var bmp = new Bitmap(vw, vh, PixelFormat.Format24bppRgb);
                    using (var g = Graphics.FromImage(bmp))
                        g.CopyFromScreen(vx, vy, 0, 0, new Size(vw, vh), CopyPixelOperation.SourceCopy);
                    return BuildImageFingerprint(bmp);
                }
                finally
                {
                    sw.Stop();
                    RunScreenProbeCount++;
                    RunScreenProbeElapsed += sw.Elapsed;
                }
            }

            internal static double SettleThresholdFor(ActionDto action) => action.Type switch
            {
                "drag_path" => 0.0025,
                "click" or "double_click" or "focus_uia" or "click_uia" or
                    "type_text" or "paste_text" or "keys" => 0.0035,
                "hold_keys" => 0.008,
                _ => NoChangeThreshold
            };
        
            internal static async Task<BatchedActionExecutionResult> ExecuteQueuedSafeActionsAsync(
                StringBuilder historyBuffer,
                int step,
                ActionDto initialAction,
                byte[] initialBaselineFingerprint,
                CancellationToken cancellationToken)
            {
                var batchIndex = 0;
                var executed = new List<ResolvedActionSnapshot>();
                var previousAction = initialAction;
                var navigationLikeBatch = IsNavigationBatchPrelude(initialAction);
                byte[]? transitionBaseline = IsAdaptiveBatchWaitTrigger(initialAction)
                    ? initialBaselineFingerprint
                    : null;
                while (PendingSafeActions.Count > 0)
                {
                    var action = PendingSafeActions.Dequeue();
                    var snapshot = CaptureResolvedAction(action, null);
                    batchIndex++;
                    Console.WriteLine($"[{step}.{batchIndex}] batch {snapshot.Description}");
        
                    try
                    {
                        if (!snapshot.IsValid)
                            throw new InvalidOperationException(snapshot.ValidationError);

                        var transitionDelayMs = BatchTransitionDelayMs(previousAction, action);
                        if (transitionDelayMs > 0)
                        {
                            Console.WriteLine(
                                $"[batch] transition {previousAction.Type}->{action.Type}; delay={transitionDelayMs}ms");
                            await Task.Delay(transitionDelayMs, cancellationToken);
                        }

                        if (action.Type == "wait")
                        {
                            var secs = EffectiveWaitSeconds(action, out var requestedSecs);
                            if (secs < requestedSecs)
                                Console.WriteLine($"[wait] Requested {requestedSecs}s capped to {secs}s.");
                            if (transitionBaseline is not null && secs > 0)
                            {
                                await WaitForBatchedTransitionAsync(
                                    transitionBaseline,
                                    previousAction,
                                    secs * 1000,
                                    navigationLikeBatch,
                                    cancellationToken);
                            }
                            else
                            {
                                Console.WriteLine($"[wait] Sleeping {secs} s (batched)...");
                                await Task.Delay(secs * 1000, cancellationToken);
                            }
                        }
                        else
                        {
                            transitionBaseline = IsAdaptiveBatchWaitTrigger(action)
                                ? CaptureScreenFingerprintProbe()
                                : null;
                            ExecuteAction(action);
                        }
        
                        AddHistory(historyBuffer, $"[{step}.{batchIndex}] batch {Describe(action)}");
                        executed.Add(snapshot);
                        previousAction = action;
                    }
                    catch (Exception ex)
                    {
                        if (ex is OperationCanceledException && CancelRequested)
                        {
                            Console.WriteLine("Aborted (hotkey).");
                            PendingSafeActions.Clear();
                            return new BatchedActionExecutionResult(executed, "batched actions cancelled by emergency hotkey");
                        }
                        Console.WriteLine($"Batched action execution error: {ex.Message}");
                        PendingSafeActions.Clear();
                        return new BatchedActionExecutionResult(
                            executed,
                            $"{snapshot.Description} failed in action batch: {ex.Message}");
                    }
                }

                return new BatchedActionExecutionResult(executed, null);
            }

            internal static bool IsAdaptiveBatchWaitTrigger(ActionDto action) =>
                action.Type is "open_url" or "launch_app" or "run_command" ||
                OpenAiResponsesService.IsCommitKeys(action);

            internal static async Task WaitForBatchedTransitionAsync(
                byte[] beforeFingerprint,
                ActionDto triggerAction,
                int timeoutMs,
                bool navigationLikeBatch,
                CancellationToken cancellationToken)
            {
                if (timeoutMs <= 0)
                    return;

                var sw = Stopwatch.StartNew();
                byte[] previous = beforeFingerprint;
                var sawChange = false;
                var lastChangeAtMs = 0L;
                var probes = 0;
                var threshold = SettleThresholdFor(triggerAction);
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
                var minimumObservationMs = AdaptiveBatchMinimumObservationMs(
                    timeoutMs,
                    navigationLikeBatch);
                var quietWindowMs = Math.Min(350, timeoutMs);

                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var remaining = timeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await Task.Delay(Math.Min(100, remaining), cancellationToken);

                    var current = CaptureScreenFingerprintProbe();
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    lastBetween = ComputeImageDelta(previous, current);
                    if (lastFromBefore >= threshold)
                        sawChange = true;
                    if (lastBetween >= threshold)
                        lastChangeAtMs = sw.ElapsedMilliseconds;

                    previous = current;
                    if (sawChange &&
                        sw.ElapsedMilliseconds >= minimumObservationMs &&
                        sw.ElapsedMilliseconds - lastChangeAtMs >= quietWindowMs)
                    {
                        Console.WriteLine(
                            $"[batch] adaptive wait settled after {sw.ElapsedMilliseconds}ms; minimum={minimumObservationMs}ms; cap={timeoutMs}ms; probes={probes}; action={triggerAction.Type}");
                        return;
                    }
                }

                Console.WriteLine(
                    $"[batch] adaptive wait reached cap after {sw.ElapsedMilliseconds}ms; probes={probes}; coarse_delta={lastFromBefore:0.####}; coarse_between_delta={lastBetween:0.####}; action={triggerAction.Type}");
            }

            internal static int AdaptiveBatchMinimumObservationMs(
                int timeoutMs,
                bool navigationLikeBatch) =>
                navigationLikeBatch
                    ? Math.Max(0, timeoutMs)
                    : Math.Min(900, Math.Max(0, timeoutMs));

            internal static bool IsNavigationBatchPrelude(ActionDto action)
            {
                if (action.Type == "open_url")
                    return true;
                if (action.Type != "keys" || action.Keys is not { Length: > 0 } keys)
                    return false;

                var normalized = keys
                    .SelectMany(key => (key ?? "").Split(
                        '+',
                        StringSplitOptions.RemoveEmptyEntries |
                        StringSplitOptions.TrimEntries))
                    .Select(key => key.ToLowerInvariant())
                    .ToArray();
                return normalized.SequenceEqual(["win", "r"]) ||
                       normalized.SequenceEqual(["super", "r"]) ||
                       normalized.SequenceEqual(["meta", "r"]) ||
                       normalized.SequenceEqual(["ctrl", "l"]) ||
                       normalized.SequenceEqual(["ctrl", "k"]);
            }

            internal static int BatchTransitionDelayMs(ActionDto previous, ActionDto next)
            {
                if (next.Type == "wait")
                    return 0;

                if (previous.Type == "keys" && OpenAiResponsesService.IsTextEntryPrelude(previous) &&
                    next.Type is "type_text" or "paste_text")
                {
                    return 250;
                }

                if (previous.Type is "type_text" or "paste_text" && OpenAiResponsesService.IsCommitKeys(next))
                    return 500;

                if (previous.Type is "type_text" or "paste_text" &&
                    next.Type is "type_text" or "paste_text")
                {
                    return 100;
                }

                return Math.Max(80, DelayFor(previous));
            }

            internal static ResolvedActionSnapshot AttachFocusedTextObservationRegion(
                ResolvedActionSnapshot snapshot,
                Rectangle? focusedUiaRect)
            {
                if (snapshot.Action.Type is not ("type_text" or "paste_text") ||
                    focusedUiaRect is not Rectangle focus ||
                    focus.Width <= 0 ||
                    focus.Height <= 0)
                {
                    return snapshot;
                }

                focus.Inflate(8, 8);
                return snapshot with { ObservationRegion = ClampRect(focus) };
            }

            internal static ResolvedActionSnapshot AttachTurnBasedObservationRegion(
                ResolvedActionSnapshot snapshot,
                Rectangle? interactionRegion,
                string observationPolicy)
            {
                if (!string.Equals(
                        observationPolicy,
                        "turn_based_interaction",
                        StringComparison.Ordinal) ||
                    interactionRegion is not Rectangle region ||
                    region.Width <= 0 ||
                    region.Height <= 0 ||
                    !IsStateChangingInteractionAction(snapshot.Action))
                {
                    return snapshot;
                }

                return snapshot with { ObservationRegion = ClampRect(region) };
            }

            internal static Rectangle? ResolveTurnBasedObservationRegion(
                Rectangle? interactionRegion,
                string observationPolicy)
            {
                if (!string.Equals(
                        observationPolicy,
                        "turn_based_interaction",
                        StringComparison.Ordinal))
                {
                    return interactionRegion;
                }

                if (interactionRegion is Rectangle persistent &&
                    persistent.Width > 0 &&
                    persistent.Height > 0)
                {
                    return ClampRect(persistent);
                }

                var activeWindow = GetActiveWindowRectangle();
                if (activeWindow is Rectangle active &&
                    active.Width > 0 &&
                    active.Height > 0)
                {
                    return ClampRect(active);
                }

                var (screenX, screenY, screenWidth, screenHeight) = GetPrimaryScreen();
                return new Rectangle(screenX, screenY, screenWidth, screenHeight);
            }

            internal static Rectangle? InferTurnBasedObservationRegion(
                ScreenObservationFrame before,
                ScreenObservationFrame after,
                Rectangle currentRegion)
            {
                if (before.DetailWidth <= 0 ||
                    before.DetailHeight <= 0 ||
                    before.DetailWidth != after.DetailWidth ||
                    before.DetailHeight != after.DetailHeight ||
                    before.DetailFingerprint.Length != after.DetailFingerprint.Length ||
                    before.ScreenBounds != after.ScreenBounds)
                {
                    return null;
                }

                currentRegion.Intersect(before.ScreenBounds);
                if (currentRegion.Width <= 0 || currentRegion.Height <= 0)
                    return null;

                var detailWidth = before.DetailWidth;
                var detailHeight = before.DetailHeight;
                var bounds = before.ScreenBounds;
                var left = Math.Clamp(
                    (int)Math.Floor(
                        (currentRegion.Left - bounds.Left) * detailWidth /
                        (double)Math.Max(1, bounds.Width)),
                    0,
                    detailWidth - 1);
                var top = Math.Clamp(
                    (int)Math.Floor(
                        (currentRegion.Top - bounds.Top) * detailHeight /
                        (double)Math.Max(1, bounds.Height)),
                    0,
                    detailHeight - 1);
                var right = Math.Clamp(
                    (int)Math.Ceiling(
                        (currentRegion.Right - bounds.Left) * detailWidth /
                        (double)Math.Max(1, bounds.Width)),
                    left + 1,
                    detailWidth);
                var bottom = Math.Clamp(
                    (int)Math.Ceiling(
                        (currentRegion.Bottom - bounds.Top) * detailHeight /
                        (double)Math.Max(1, bounds.Height)),
                    top + 1,
                    detailHeight);

                var changed = new bool[detailWidth * detailHeight];
                var useColor = before.DetailColorFingerprint.Length ==
                                   detailWidth * detailHeight * 3 &&
                               after.DetailColorFingerprint.Length ==
                                   detailWidth * detailHeight * 3;
                for (var y = top; y < bottom; y++)
                for (var x = left; x < right; x++)
                {
                    var pixel = y * detailWidth + x;
                    var maximumDifference = Math.Abs(
                        before.DetailFingerprint[pixel] -
                        after.DetailFingerprint[pixel]);
                    if (useColor)
                    {
                        var color = pixel * 3;
                        maximumDifference = Math.Max(
                            maximumDifference,
                            Math.Max(
                                Math.Abs(before.DetailColorFingerprint[color] -
                                         after.DetailColorFingerprint[color]),
                                Math.Max(
                                    Math.Abs(before.DetailColorFingerprint[color + 1] -
                                             after.DetailColorFingerprint[color + 1]),
                                    Math.Abs(before.DetailColorFingerprint[color + 2] -
                                             after.DetailColorFingerprint[color + 2]))));
                    }
                    changed[pixel] = maximumDifference >= 12;
                }

                var regionPixelCount = (right - left) * (bottom - top);
                var minimumComponentSize = Math.Max(
                    4,
                    (int)Math.Ceiling(regionPixelCount * 0.0015));
                var visited = new bool[changed.Length];
                var meaningfulComponents = new List<(int Left, int Top, int Right, int Bottom, int Count)>();
                for (var start = 0; start < changed.Length; start++)
                {
                    if (!changed[start] || visited[start])
                        continue;

                    var startX = start % detailWidth;
                    var startY = start / detailWidth;
                    if (startX < left || startX >= right || startY < top || startY >= bottom)
                        continue;

                    var queue = new Queue<int>();
                    queue.Enqueue(start);
                    visited[start] = true;
                    var componentLeft = startX;
                    var componentTop = startY;
                    var componentRight = startX + 1;
                    var componentBottom = startY + 1;
                    var count = 0;
                    while (queue.Count > 0)
                    {
                        var current = queue.Dequeue();
                        var x = current % detailWidth;
                        var y = current / detailWidth;
                        count++;
                        componentLeft = Math.Min(componentLeft, x);
                        componentTop = Math.Min(componentTop, y);
                        componentRight = Math.Max(componentRight, x + 1);
                        componentBottom = Math.Max(componentBottom, y + 1);
                        for (var dy = -1; dy <= 1; dy++)
                        for (var dx = -1; dx <= 1; dx++)
                        {
                            if (dx == 0 && dy == 0)
                                continue;
                            var nx = x + dx;
                            var ny = y + dy;
                            if (nx < left || nx >= right || ny < top || ny >= bottom)
                                continue;
                            var neighbor = ny * detailWidth + nx;
                            if (!changed[neighbor] || visited[neighbor])
                                continue;
                            visited[neighbor] = true;
                            queue.Enqueue(neighbor);
                        }
                    }

                    if (count >= minimumComponentSize)
                    {
                        meaningfulComponents.Add((
                            componentLeft,
                            componentTop,
                            componentRight,
                            componentBottom,
                            count));
                    }
                }

                if (meaningfulComponents.Sum(component => component.Count) <
                    Math.Max(8, regionPixelCount * 0.004))
                {
                    return null;
                }

                var changedLeft = meaningfulComponents.Min(component => component.Left);
                var changedTop = meaningfulComponents.Min(component => component.Top);
                var changedRight = meaningfulComponents.Max(component => component.Right);
                var changedBottom = meaningfulComponents.Max(component => component.Bottom);
                var horizontalPadding = Math.Max(2, (right - left) / 40);
                var verticalPadding = Math.Max(2, (bottom - top) / 40);
                changedLeft = Math.Max(left, changedLeft - horizontalPadding);
                changedTop = Math.Max(top, changedTop - verticalPadding);
                changedRight = Math.Min(right, changedRight + horizontalPadding);
                changedBottom = Math.Min(bottom, changedBottom + verticalPadding);

                var inferred = Rectangle.FromLTRB(
                    bounds.Left + (int)Math.Floor(
                        changedLeft * bounds.Width / (double)detailWidth),
                    bounds.Top + (int)Math.Floor(
                        changedTop * bounds.Height / (double)detailHeight),
                    bounds.Left + (int)Math.Ceiling(
                        changedRight * bounds.Width / (double)detailWidth),
                    bounds.Top + (int)Math.Ceiling(
                        changedBottom * bounds.Height / (double)detailHeight));
                inferred.Intersect(currentRegion);
                if (inferred.Width < currentRegion.Width * 0.10 ||
                    inferred.Height < currentRegion.Height * 0.10)
                {
                    return null;
                }

                var currentArea = (long)currentRegion.Width * currentRegion.Height;
                var inferredArea = (long)inferred.Width * inferred.Height;
                var inferredRatio = inferredArea / (double)Math.Max(1, currentArea);
                return inferredRatio is >= 0.03 and <= 0.80
                    ? ClampRect(inferred)
                    : null;
            }

            internal static bool IsTurnBasedStateInspection(ActionDto action)
            {
                if (action.Type != "request_crop")
                    return false;

                var note = action.Note ?? "";
                if (Regex.IsMatch(
                        note,
                        @"\b(grid|playfield|game area|game state|puzzle|maze|map|canvas|siatk\w*|labirynt\w*|pole gr\w*)\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    return true;
                }
                if (IsTurnBasedAuxiliaryInspection(action))
                    return false;
                return Regex.IsMatch(
                    note,
                    @"\b(board|grid|playfield|level|game area|game state|puzzle|maze|map|canvas|plansz\w*|siatk\w*|poziom\w*|stan\w* gr\w*|map\w*|labirynt\w*|pole gr\w*)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            internal static bool ShouldEstablishTurnBasedInteractionRegion(
                ActionDto action,
                bool regionRequired) =>
                action.Type == "request_crop" &&
                (IsTurnBasedStateInspection(action) ||
                 regionRequired && !IsTurnBasedAuxiliaryInspection(action));

            internal static bool CanReplaceTurnBasedInteractionRegion(
                bool hasExistingRegion,
                bool existingRegionIsAutomatic,
                bool automaticRegionWasRefined,
                ActionDto action,
                bool regionRequired) =>
                ShouldEstablishTurnBasedInteractionRegion(action, regionRequired) &&
                (!hasExistingRegion ||
                 existingRegionIsAutomatic && !automaticRegionWasRefined);

            internal static bool IsTurnBasedAuxiliaryInspection(ActionDto action)
            {
                if (action.Type != "request_crop")
                    return false;

                return Regex.IsMatch(
                    action.Note ?? "",
                    @"\b(control\w*|button\w*|help\w*|instruction\w*|toolbar\w*|panel\w*|menu\w*|sterow\w*|przycisk\w*|pomoc\w*|instrukcj\w*)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            internal static bool IsOverlappingTurnInspection(
                Rectangle persistentRegion,
                Rectangle requestedRegion)
            {
                var intersection = Rectangle.Intersect(
                    persistentRegion,
                    requestedRegion);
                if (intersection.Width <= 0 || intersection.Height <= 0)
                    return false;
                var intersectionArea = (long)intersection.Width * intersection.Height;
                var smallerArea = Math.Min(
                    (long)persistentRegion.Width * persistentRegion.Height,
                    (long)requestedRegion.Width * requestedRegion.Height);
                return smallerArea > 0 &&
                       intersectionArea / (double)smallerArea >= 0.60;
            }

            internal static string TurnBasedContextKey(UiPromptContext context) =>
                $"{context.ActiveProcessName?.Trim().ToLowerInvariant()}|" +
                context.ActiveWindowTitle?.Trim().ToLowerInvariant();

            internal static bool AllowsStableCanvasDrawBatch(
                UiPromptContext context,
                string observationProfile)
            {
                if (!MouseEnabled ||
                    !DirectClickWithoutAim ||
                    observationProfile is "realtime_interaction" or "turn_based_interaction" or "streaming_output")
                {
                    return false;
                }

                if (string.Equals(
                        observationProfile,
                        "local_editing",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                var semantic = $"{context.ActiveProcessName} {context.ActiveWindowTitle} {context.FocusedUiaSummary}";
                return Regex.IsMatch(
                    semantic,
                    @"\b(mspaint|paint|canvas|drawing|sketch|krita|gimp|photoshop|inkscape|płótno|rysunek)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            internal static bool ShouldDeferFocusedTextStagnation(
                ResolvedActionSnapshot? previousAction,
                bool observedNoChange,
                int priorNoChangeAttempts) =>
                observedNoChange &&
                priorNoChangeAttempts == 0 &&
                previousAction is not null &&
                IsTextInputAttemptAction(previousAction.Action);
        
            internal static bool IsSafeBatchedAction(ActionDto action)
            {
                if (action.Type == "run_command" && !AllowRunCommand)
                    return false;
        
                if (action.Type is "open_url" or "launch_app")
                    return AllowHighLevelActions;
        
                return action.Type is "keys" or "type_text" or "paste_text" or "wait" or "run_command" or "drag_path";
            }
        
            internal static bool IsHighImpactGoal(string goal)
            {
                var g = goal.ToLowerInvariant();
                string[] highImpactTerms =
                {
                    "wyślij", "wyslij", "email", "e-mail", "mail", "outlook",
                    "zapisz", "plik", "usuń", "usun", "word", "excel", "powerpoint",
                    "uruchom skrypt", "powershell", "cmd"
                };
        
                return highImpactTerms.Any(t => g.Contains(t, StringComparison.OrdinalIgnoreCase));
            }
        
            internal static int DelayFor(ActionDto a) => a.Type switch
            {
                "open_url" => 500,
                "launch_app" => 500,
                "run_command" => 500,
                "paste_text" => 150,
                "move" => 120,
                "click" => 180,
                "double_click" => 250,
                "drag_drop" => 500,
                "drag_path" => 250,
                "keys" => 120,
                "hold_keys" => 80,
                "type_text" => 80,
                "scroll" => 80,
                "request_crop" => 80,
                "point" => 80,
                "aim" => 80,
                "wait" => 0,   // wait time handled separately
                "done" => 80,
                _ => 120
            };
        
            internal static int EffectiveWaitSeconds(ActionDto a, out int requestedSeconds)
            {
                requestedSeconds = Math.Max(0, a.WaitSeconds ?? 1);
                return MaxWaitSeconds > 0
                    ? Math.Min(requestedSeconds, MaxWaitSeconds)
                    : requestedSeconds;
            }
        
            internal static bool ShouldVerifyGoal(string goal, int step, ActionDto? action = null)
            {
                if (VerifyMode.Equals("always", StringComparison.OrdinalIgnoreCase))
                    return true;
                if (VerifyMode.Equals("off", StringComparison.OrdinalIgnoreCase))
                    return false;
        
                if (IsHighImpactGoal(goal))
                    return true;
        
                if (action?.Type == "done" &&
                    action.Confidence is double doneConfidence &&
                    doneConfidence >= SkipVerifyConfidenceThreshold &&
                    step > VerifyEarlySteps)
                    return false;
        
                if (VerifyEarlySteps > 0 && step <= VerifyEarlySteps)
                    return true;
        
                if (action?.Confidence is double confidence)
                    return confidence < VerifyLowConfidenceThreshold;
        
                return true;
            }

            internal static string VerificationSkipReason(int step, ActionDto? action)
            {
                if (VerifyMode.Equals("off", StringComparison.OrdinalIgnoreCase))
                    return "verify=off";
                if (action?.Type == "done" &&
                    action.Confidence is double doneConfidence &&
                    doneConfidence >= SkipVerifyConfidenceThreshold &&
                    step > VerifyEarlySteps)
                {
                    return $"verify=auto high-confidence completion " +
                           $"({doneConfidence:0.00} >= {SkipVerifyConfidenceThreshold:0.00})";
                }
                if (action?.Confidence is double confidence &&
                    confidence >= VerifyLowConfidenceThreshold)
                {
                    return $"verify=auto confidence above verification threshold " +
                           $"({confidence:0.00} >= {VerifyLowConfidenceThreshold:0.00})";
                }
                return $"verify={VerifyMode} policy";
            }
        
            internal static bool TryExecuteFastGoal(string goal)
            {
                var g = goal.Trim();
                if (string.IsNullOrWhiteSpace(g))
                    return false;
        
                if (TryExtractGoogleSearch(g, out var searchTerm))
                {
                    var url = "https://www.google.com/search?q=" + Uri.EscapeDataString(searchTerm);
                    Console.WriteLine($"[fast-path] open_url {url}");
                    OpenUrl(url);
                    return true;
                }
        
                if (TryExtractUrl(g, out var directUrl))
                {
                    Console.WriteLine($"[fast-path] open_url {directUrl}");
                    OpenUrl(directUrl);
                    return true;
                }
        
                if (TryExtractSimpleLaunchApp(g, out var app))
                {
                    Console.WriteLine($"[fast-path] launch_app {app}");
                    LaunchApp(app);
                    return true;
                }
        
                return false;
            }
        
            internal static bool TryExtractGoogleSearch(string goal, out string term)
            {
                term = "";
                var lower = goal.ToLowerInvariant();
                if (!lower.Contains("google") || !(lower.Contains("wyszukaj") || lower.Contains("search")))
                    return false;
        
                var quoted = ExtractFirstQuoted(goal);
                if (!string.IsNullOrWhiteSpace(quoted))
                {
                    term = quoted.Trim();
                    return true;
                }
        
                var markers = new[] { "frazę", "fraze", "term", "search for" };
                foreach (var marker in markers)
                {
                    var idx = lower.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase);
                    if (idx >= 0)
                    {
                        term = goal[(idx + marker.Length)..].Trim(' ', ':', '.', '"', '\'');
                        return !string.IsNullOrWhiteSpace(term);
                    }
                }
        
                return false;
            }
        
            internal static bool TryExtractUrl(string goal, out string url)
            {
                url = "";
                var match = System.Text.RegularExpressions.Regex.Match(goal, @"https?://[^\s""]+|(?<!@)\b[a-z0-9][a-z0-9-]*(?:\.[a-z0-9][a-z0-9-]*)+\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!match.Success)
                    return false;
        
                url = match.Value;
                if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    url = "https://" + url;
                return true;
            }
        
            internal static bool TryExtractSimpleLaunchApp(string goal, out string app)
            {
                app = "";
                var lower = goal.ToLowerInvariant();
                if (!(lower.StartsWith("otwórz aplikację ") || lower.StartsWith("otworz aplikacje ") || lower.StartsWith("open ")))
                    return false;
                if (lower.Contains("następnie") || lower.Contains("nastepnie") || lower.Contains(" i "))
                    return false;
        
                var known = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["edge"] = "msedge",
                    ["microsoft edge"] = "msedge",
                    ["outlook"] = "outlook",
                    ["paint"] = "mspaint",
                    ["word"] = "winword",
                    ["excel"] = "excel",
                    ["powerpoint"] = "powerpnt",
                    ["notepad"] = "notepad",
                    ["notatnik"] = "notepad"
                };
        
                foreach (var kv in known)
                {
                    if (lower.Contains(kv.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        app = kv.Value;
                        return true;
                    }
                }
        
                return false;
            }
        
            internal static string? ExtractFirstQuoted(string text)
            {
                var match = System.Text.RegularExpressions.Regex.Match(text, "[\"“”'„](.*?)[\"“”'”]");
                return match.Success ? match.Groups[1].Value : null;
            }
        
            internal static string? EffectiveReasoningEffort(
                int stagnationSteps,
                int repeatCount,
                int rejectedProposalCycleCount = 0)
            {
                if (!AdaptiveReasoningEffort || string.IsNullOrWhiteSpace(ReasoningEffort) || !SupportsReasoningEffort(Model))
                    return ReasoningEffort;
        
                var current = ReasoningEffort.ToLowerInvariant();
                if (current is "high" or "xhigh" or "max")
                    return ReasoningEffort;
        
                if (stagnationSteps >= 4 ||
                    repeatCount >= 3 ||
                    rejectedProposalCycleCount >= 3)
                    return "high";
        
                if (stagnationSteps >= 2 ||
                    repeatCount >= 1 ||
                    rejectedProposalCycleCount >= 1)
                    return "medium";
        
                return ReasoningEffort;
            }
    }
}

