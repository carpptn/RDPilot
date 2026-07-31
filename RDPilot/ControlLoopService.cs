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
                using var tee = new TeeTextWriter(prevOut, logFile);
                Console.SetOut(tee);
                Console.SetError(tee);
        
                CancellationTokenSource? cancelCts = null;
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
                        HasRecurringWorkflowIntent(goal);
                    Console.WriteLine($"Goal mode: {goalMode}");
                    if (recurringWorkflowIntent)
                        Console.WriteLine("Goal cycle policy: recurring workflow may contain productive state returns.");
                    Console.WriteLine("Loop start: one action -> screenshot -> next decision.");
                    Console.WriteLine("Emergency abort: Ctrl+Alt+Q\n");
        
                    if (AutoHideConsoleDuringRun || MinimizeConsoleDuringRun)
                        consoleHidden = ConcealConsoleWindow();
        
                    if (AllowHighLevelActions && TryExecuteFastGoal(goal))
                    {
                        Console.WriteLine("Finished (local fast path).");
                        return new ControlRunResult(
                            ControlRunOutcome.Completed,
                            0,
                            "completed through the local fast path");
                    }
        
                    var systemRules = BuildSystemRules(); // stable per run; dynamic action availability is enforced by schema
                    var historyBuffer = new StringBuilder();
                    var recoveryLessons = LoadRecoveryLessons();
                    var recentExecutedActions = new Queue<ResolvedActionSnapshot>();
                    var loopStateGraph = new LoopStateGraph { RunId = commandId };
                    RecoveryEpisodeState? recoveryEpisode = null;
                    cancelCts = StartCancelHotkeyListener();
        
                    Rectangle? nextFocusRect = null; // crop/overlay after 'aim'/'point'/'request_crop'
                    Rectangle? lastAimRect = null;   // active AIM – required before clicks
        
                    // change / strategy metrics
                    byte[]? prevShotFingerprint = null;
                    byte[]? prevActiveWindowFingerprint = null;
                    ResolvedActionSnapshot? prevAction = null;
                    string? lastSig = null;
                    var recentIneffectiveSpatialActions = new List<ResolvedActionSnapshot>();
                    int stagnationSteps = 0;
                    int repeatCount = 0;
                    int continuousIdleSteps = 0;
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
                    string? previousControlResponseId = null;
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
        
                        // screenshot at the beginning of a step (state after previous action)
                        var (dataUrl, savedPath, screenW, screenH, imageW, imageH, focusUrl, appliedFocusRect, focusUiaRect, focusUiaSummary, focusUiaDataUrl, focusUiaPath, shotFingerprint, activeWindowFingerprint) =
                            ScreenshotToDataUrl(screensDir, commandId, step, nextFocusRect);
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
        
                        // — Visual delta vs previous screenshot (effect of last action)
                        if (prevShotFingerprint != null)
                        {
                            lastGlobalDelta = ComputeImageDelta(prevShotFingerprint, shotFingerprint);
                            lastActiveWindowDelta = prevActiveWindowFingerprint is not null
                                ? ComputeImageDelta(prevActiveWindowFingerprint, activeWindowFingerprint)
                                : lastGlobalDelta;
                            // Prefer the foreground application. This prevents
                            // background animations from looking like task progress.
                            lastDelta = 0.80 * lastActiveWindowDelta + 0.20 * lastGlobalDelta;
                            bool noChange = lastDelta < NoChangeThreshold;
                            bool previousWasObservationOnly = prevAction != null && IsLocalObservationAction(prevAction.Action);
                            expectedContinuousIdle = IsExpectedContinuousIdle(
                                goalMode,
                                prevAction?.Action,
                                noChange);

                            if (ShouldResetRejectedProposalLoop(
                                    prevAction,
                                    noChange,
                                    expectedContinuousIdle))
                            {
                                ClearRejectedProposalLoop();
                            }

                            if (expectedContinuousIdle)
                            {
                                continuousIdleSteps++;
                                stagnationSteps = 0;
                                repeatCount = 0;
                                lastSig = null;
                                recentIneffectiveSpatialActions.Clear();
                            }
                            else if (!previousWasObservationOnly)
                            {
                                continuousIdleSteps = 0;
                                if (noChange) stagnationSteps++; else stagnationSteps = 0;
                            }
                            else
                            {
                                continuousIdleSteps = 0;
                                if (!noChange)
                                    stagnationSteps = 0;
                            }
        
                            if (prevAction != null && !expectedContinuousIdle)
                            {
                                (repeatCount, lastSig) = UpdateRepeatDetection(
                                    prevAction,
                                    noChange,
                                    repeatCount,
                                    lastSig,
                                    recentIneffectiveSpatialActions);

                                if (noChange &&
                                    repeatCount > 0 &&
                                    ActionRepeatCooldownSteps > 0 &&
                                    !IsLocalObservationAction(prevAction.Action))
                                {
                                    RegisterActionCooldown(
                                        prevAction,
                                        step + ActionRepeatCooldownSteps,
                                        actionCooldownUntilStep,
                                        spatialActionCooldowns);
                                }
                            }
        
                            // Expire AIM after a large visual change
                            if (lastAimRect is not null && lastDelta > AimExpireDelta)
                            {
                                Console.WriteLine($"[aim] expired (delta={lastDelta:0.###} > {AimExpireDelta:0.###})");
                                lastAimRect = null;
                            }
        
                            if (noChange && prevAction != null && IsPointClickAction(prevAction.Action))
                            {
                                lastPrecisionHint = BuildPrecisionHint(prevAction, lastDelta, "Previous click produced little or no visible progress.");
                                lastPrecisionHintExpiresAfterStep = step + 1;
                            }
                            if (noChange && prevAction?.Action.Type == "drag_drop")
                            {
                                lastPrecisionHint =
                                    "Previous drag_drop did not visibly move the source or change the destination. " +
                                    "Do not retry nearby coordinates; verify that the source is draggable, identify the semantic destination, or use a different interaction route.";
                                lastPrecisionHintExpiresAfterStep = step + 2;
                            }
                            if (noChange && prevAction != null && IsTextInputAttemptAction(prevAction.Action))
                            {
                                textInputNoChangeAttempts++;
                                lastTextInputHint = BuildTextInputHint(prevAction.Action, lastDelta);
                                lastTextInputHintExpiresAfterStep = step + 4;
                                if (textInputNoChangeAttempts >= 2)
                                    textInputCooldownUntilStep = Math.Max(textInputCooldownUntilStep, step + 4);
                            }
                            else if (!noChange)
                            {
                                lastPrecisionHint = null;
                                lastPrecisionHintExpiresAfterStep = 0;
                                lastTextInputHint = null;
                                lastTextInputHintExpiresAfterStep = 0;
                                textInputNoChangeAttempts = 0;
                                textInputCooldownUntilStep = 0;
                            }
                        }
        
                        var previousResponseIdForRequest = UsePreviousResponseState ? previousControlResponseId : null;
                        var historyTail = previousResponseIdForRequest != null || HistoryTailChars <= 0
                            ? ""
                            : TailHistory(historyBuffer, HistoryTailChars, HistoryTailLines);
                        var (screenCx, screenCy, _, _) = GetCursorPositionInPrimary();
                        var (cx, cy, cnx, cny) = CursorToImageCoordinates(screenCx, screenCy);
                        var promptContext = CaptureUiPromptContext(focusUiaSummary, screenW, screenH);
                        var appliedFocusRectForPrompt = ScreenRectToImage(appliedFocusRect);
                        var focusUiaRectForPrompt = ScreenRectToImage(focusUiaRect);
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
                                recurringWorkflowIntent);
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
                            goalMode,
                            recurringWorkflowIntent);

                        recoveryEpisode = await UpdateRecoveryEpisodeAsync(
                            recoveryEpisode,
                            step,
                            stagnationSteps,
                            repeatCount,
                            lastDelta,
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
                                              !double.IsNaN(lastDelta) &&
                                              lastDelta < NoChangeThreshold &&
                                              CurrentUiaTargets.Count > 0;
                        PrepareUiaTargetsForPrompt(reuseUiaTargets, screenW, screenH);
        
                        RequestReasoningEffortOverride = EffectiveReasoningEffort(
                            stagnationSteps,
                            repeatCount,
                            rejectedProposalCycleCount);
        
                        // inject observation metrics into the prompt
                        var metaSb = new StringBuilder()
                            .AppendLine($"LAST_STEP_DELTA: {(double.IsNaN(lastDelta) ? "N/A" : lastDelta.ToString("0.####"))} (threshold={NoChangeThreshold})")
                            .AppendLine($"LAST_GLOBAL_DELTA: {(double.IsNaN(lastGlobalDelta) ? "N/A" : lastGlobalDelta.ToString("0.####"))}; LAST_ACTIVE_WINDOW_DELTA: {(double.IsNaN(lastActiveWindowDelta) ? "N/A" : lastActiveWindowDelta.ToString("0.####"))}")
                            .AppendLine($"STAGNATION_STEPS: {stagnationSteps}")
                            .AppendLine($"REPEAT_COUNT: {repeatCount}")
                            .AppendLine($"REJECTED_PROPOSAL_CYCLE_COUNT: {rejectedProposalCycleCount}")
                            .AppendLine($"CONTINUOUS_IDLE_STEPS: {continuousIdleSteps}")
                            .AppendLine($"GOAL_MODE: {goalMode}")
                            .AppendLine($"REQUEST_REASONING_EFFORT: {RequestReasoningEffortOverride ?? ReasoningEffort ?? "default"}")
                            .AppendLine($"LAST_ACTION: {(prevAction == null ? "N/A" : prevAction.Description)}")
                            .AppendLine($"AIM_ACTIVE: {(lastAimRect is null ? "false" : $"true {FormatImageRect(CurrentScreenMap.ScreenToImageRect(lastAimRect.Value))}")}");
                        if (repeatCount > 0 || stagnationSteps > 0)
                            metaSb.AppendLine("STRATEGY_HINT: The previous action did not visibly advance the screen. Do not repeat it; choose a different UI route or ask for a crop if the target is ambiguous.");
                        if (expectedContinuousIdle)
                            metaSb.AppendLine("CONTINUOUS_IDLE: The previous wait left the screen unchanged, which is valid for this open-ended goal. Reassess whether the requested state is still healthy or whether a new event is present; wait again only when continued observation is goal-aligned.");
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
                                                  !double.IsNaN(lastDelta) &&
                                                  lastDelta < NoChangeThreshold;
                        if (omitFullScreenImage)
                            metaSb.AppendLine("SCREEN_IMAGE: omitted because screen fingerprint is unchanged; use previous_response_id state plus current metadata.");
        
                        var reqBody = BuildRequestBody(Model, systemRules, goal, historyTail + "\n" + metaSb, dataUrl, imageW, imageH,
                                                       cx, cy, cnx, cny, focusUrl, appliedFocusRectForPrompt, focusUiaRectForPrompt, focusUiaDataUrl,
                                                       promptContext, reuseUiaTargets, previousResponseIdForRequest, omitFullScreenImage, goalMode);
                        if (LogRequests)
                        {
                            var reqBodyForLog = BuildRequestBody_ForLog(Model, systemRules, goal, historyTail + "\n" + metaSb,
                                                                        omitFullScreenImage ? null : savedPath, imageW, imageH, cx, cy, cnx, cny,
                                                                        appliedFocusRect != null && LogScreens ? ScreenLogPath(screensDir, $"{commandId}_{step}_crop") : null,
                                                                        appliedFocusRectForPrompt,
                                                                        focusUiaRectForPrompt, focusUiaPath, promptContext, previousResponseIdForRequest, omitFullScreenImage, goalMode);
                            SaveJson(Path.Combine(requestsDir, $"{commandId}_{step}_request.json"), reqBodyForLog);
                        }
        
                        var (action, raw) = await CallOpenAIAsync(apiKey, reqBody, cancelCts.Token);
                        if (UsePreviousResponseState)
                            previousControlResponseId = LastOpenAiResponseId ?? previousControlResponseId;
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
        
                        var currentAction = CaptureResolvedAction(action, lastAimRect);
                        Console.WriteLine($"[{step}] {currentAction.Description}");
                        if (action.Confidence is double confidence)
                            Console.WriteLine($"     confidence: {confidence:0.##}");
                        if (!string.IsNullOrWhiteSpace(action.Note))
                            Console.WriteLine($"     note: {action.Note}");
        
                        nextFocusRect = null; // reset – set by aim/point/request_crop
                        var actionExecutionFailed = false;
                        var actionExecuted = false;
                        var actionWasLocallyRejected = false;
                        try
                        {
                            if (!currentAction.IsValid)
                                throw new InvalidOperationException(currentAction.ValidationError);

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
                                continue;
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
                                        var (freshDataUrl, freshPath, freshW, freshH, freshImageW, freshImageH, _, _, freshFocusRect, freshFocusSummary, _, _, _, _) = ScreenshotToDataUrl(screensDir, commandId, step, null);
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
                                        Console.WriteLine($"[verify] ✅ Goal confirmed: {verify.Reason}");
                                        AddHistory(historyBuffer, $"[{step}] done_verified");
                                        lastVerifierRejection = null;
                                        lastAimRect = null;
                                        recoveryEpisode = ConfirmPendingRecovery(recoveryEpisode, recoveryLessons, independentlyVerified);
                                        Console.WriteLine("Finished (model returned 'done').");
                                        runResult = new ControlRunResult(
                                            ControlRunOutcome.Completed,
                                            step,
                                            verify?.Reason ??
                                            "the model completion was verified");
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
                            else if (action.Type == "drag_drop")
                            {
                                if (!HasExplicitPoint(action) || !HasExplicitDropPoint(action))
                                    throw new InvalidOperationException("drag_drop requires an explicit source and destination.");

                                var source = ResolvePoint(action);
                                if (lastAimRect is null && !DirectClickWithoutAim)
                                {
                                    Console.WriteLine("[guard] drag_drop blocked: no active source AIM. Return 'aim' for the source first.");
                                    AddHistory(historyBuffer, $"[{step}] IGNORED (drag_without_aim)");
                                    lastExecutorFailure = "drag_drop was blocked because its source had no active AIM";
                                    actionExecutionFailed = true;
                                    actionWasLocallyRejected = true;
                                }
                                else if (lastAimRect is Rectangle dragAim && !dragAim.Contains(source.X, source.Y))
                                {
                                    Console.WriteLine("[guard] drag_drop source outside active AIM → ignoring. Set AIM around the source object.");
                                    AddHistory(historyBuffer, $"[{step}] IGNORED (drag_source_outside_aim)");
                                    lastExecutorFailure = "drag_drop was blocked because its source was outside the active AIM";
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
                            else if (action.Type is "click" or "double_click")
                            {
                                if (lastAimRect is null && DirectClickWithoutAim && HasExplicitPoint(action))
                                {
                                    Console.WriteLine("[guard] direct click without AIM allowed by profile.");
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
                            consecutiveActionFailures = 0;
                            lastExecutorFailure = null;
                            AddHistory(historyBuffer, $"[{step}] {currentAction.Description}");
                            if (!string.IsNullOrWhiteSpace(action.Note))
                                AddHistory(historyBuffer, $"[{step}] note: {action.Note}");
                            if (action.Type != "done" && lastVerifierRejection != null)
                                lastVerifierRejection = null;

                            RecordRecoveryAction(recoveryEpisode, recentExecutedActions, currentAction, recoveryLessons);

                            var batchResult = await ExecuteQueuedSafeActionsAsync(historyBuffer, step, cancelCts.Token);
                            foreach (var batchedAction in batchResult.ExecutedActions)
                            {
                                RecordRecoveryAction(recoveryEpisode, recentExecutedActions, batchedAction, recoveryLessons);
                                currentAction = batchedAction;
                            }
                            if (!string.IsNullOrWhiteSpace(batchResult.Error))
                                lastExecutorFailure = batchResult.Error;
                        }
        
                        // Keep context for next step (delta/repeat metrics)
                        prevAction = actionExecutionFailed || !actionExecuted ? null : currentAction;
                        prevShotFingerprint = actionExecutionFailed || !actionExecuted ? null : shotFingerprint;
                        prevActiveWindowFingerprint = actionExecutionFailed || !actionExecuted
                            ? null
                            : activeWindowFingerprint;
        
                        if (CancelRequested)
                        {
                            Console.WriteLine("Aborted (hotkey).");
                            runResult = new ControlRunResult(
                                ControlRunOutcome.Cancelled,
                                step,
                                "cancelled with the emergency hotkey");
                            break;
                        }
        
                        if (!actionExecutionFailed && actionExecuted)
                        {
                            try { await WaitAfterActionAsync(currentAction.Action, shotFingerprint, cancelCts.Token); }
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
                        $"Control run failed unexpectedly: {ex.Message}");
                    runResult = new ControlRunResult(
                        ControlRunOutcome.Failed,
                        currentStep,
                        ex.Message);
                }
                finally
                {
                    cancelCts?.Cancel();
                    cancelCts?.Dispose();
                    _ = FlushPendingRecoveryMemory();
                    if (LoopReplayAutoExportEnabled && RecoveryMemoryEnabled)
                        TryAutoExportLoopReplayCorpus();
                    PrintRunMetrics();
                    Console.SetOut(prevOut);
                    Console.SetError(prevErr);
                    if (consoleHidden && RestoreConsoleAfterRun)
                        RestoreConsoleWindow();
                }
                return runResult with { Step = currentStep };
            }
        
            internal static bool IsMouseAction(ActionDto a)
                => a.Type is "move" or "click" or "double_click" or "drag_drop" or "scroll" or "focus_uia" or "click_uia";
        
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
        
            internal static bool IsLocalObservationAction(ActionDto a)
                => a.Type is "aim" or "point" or "request_crop";

            internal static bool IsExpectedContinuousIdle(
                string goalMode,
                ActionDto? previousAction,
                bool noChange) =>
                noChange &&
                string.Equals(goalMode, "continuous", StringComparison.Ordinal) &&
                previousAction?.Type == "wait";

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

                if (!cooldowns.TryGetValue(action.IneffectiveSignature, out untilStep))
                    return false;
        
                if (step <= untilStep)
                    return true;
        
                cooldowns.Remove(action.IneffectiveSignature);
                return false;
            }

            internal static void RegisterActionCooldown(
                ResolvedActionSnapshot action,
                int untilStep,
                Dictionary<string, int> cooldowns,
                List<SpatialActionCooldown> spatialCooldowns)
            {
                var family = RecoveryMemoryService.ActionFamily(action.Action);
                if (action.ScreenPoint is Point screenPoint &&
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
        
            internal static async Task WaitAfterActionAsync(ActionDto action, byte[] beforeFingerprint, CancellationToken cancellationToken)
            {
                if (IsLocalObservationAction(action))
                    return;
        
                if (!ScreenPollingEnabled)
                {
                    var fixedDelay = PostActionDelay(action);
                    if (fixedDelay > 0)
                        await Task.Delay(fixedDelay, cancellationToken);
                    return;
                }
        
                if (action.Type == "wait")
                {
                    if (WaitNoChangeExtraMs <= 0)
                        return;
        
                    var afterWait = CaptureScreenFingerprintProbe();
                    var delta = ComputeImageDelta(beforeFingerprint, afterWait);
                    if (delta < NoChangeThreshold)
                    {
                        Console.WriteLine($"[settle] screen unchanged after wait (delta={delta:0.####}); waiting extra {WaitNoChangeExtraMs} ms.");
                        await Task.Delay(WaitNoChangeExtraMs, cancellationToken);
                    }
                    return;
                }
        
                var initialDelay = Math.Min(PostActionDelay(action), Math.Max(0, ScreenPollInitialDelayMs));
                if (initialDelay > 0)
                    await Task.Delay(initialDelay, cancellationToken);
        
                if (ScreenPollTimeoutMs <= 0)
                    return;
        
                await WaitForScreenStableAsync(beforeFingerprint, action, cancellationToken);
            }
        
            internal static async Task WaitForScreenStableAsync(byte[] beforeFingerprint, ActionDto action, CancellationToken cancellationToken)
            {
                var sw = Stopwatch.StartNew();
                byte[]? previous = null;
                var sawChange = false;
                var probes = 0;
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
        
                while (sw.ElapsedMilliseconds < ScreenPollTimeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var current = CaptureScreenFingerprintProbe();
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    if (lastFromBefore >= NoChangeThreshold)
                        sawChange = true;
        
                    if (previous != null)
                    {
                        lastBetween = ComputeImageDelta(previous, current);
                        if (lastBetween < NoChangeThreshold && (sawChange || probes >= 2))
                        {
                            Console.WriteLine($"[settle] stable after {sw.ElapsedMilliseconds} ms; probes={probes}; delta={lastFromBefore:0.####}; action={action.Type}");
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
                    Console.WriteLine($"[settle] timeout after {sw.ElapsedMilliseconds} ms; probes={probes}; delta={lastFromBefore:0.####}; between={lastBetween:0.####}; action={action.Type}");
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
        
            internal static async Task<BatchedActionExecutionResult> ExecuteQueuedSafeActionsAsync(
                StringBuilder historyBuffer,
                int step,
                CancellationToken cancellationToken)
            {
                var batchIndex = 0;
                var executed = new List<ResolvedActionSnapshot>();
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

                        if (action.Type == "wait")
                        {
                            var secs = EffectiveWaitSeconds(action, out var requestedSecs);
                            if (secs < requestedSecs)
                                Console.WriteLine($"[wait] Requested {requestedSecs}s capped to {secs}s.");
                            Console.WriteLine($"[wait] Sleeping {secs} s (batched)...");
                            await Task.Delay(secs * 1000, cancellationToken);
                        }
                        else
                        {
                            ExecuteAction(action);
                            await Task.Delay(Math.Max(0, DelayFor(action)), cancellationToken);
                        }
        
                        AddHistory(historyBuffer, $"[{step}.{batchIndex}] batch {Describe(action)}");
                        executed.Add(snapshot);
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
        
            internal static bool IsSafeBatchedAction(ActionDto action)
            {
                if (action.Type == "run_command" && !AllowRunCommand)
                    return false;
        
                if (action.Type is "open_url" or "launch_app")
                    return AllowHighLevelActions;
        
                return action.Type is "keys" or "type_text" or "paste_text" or "wait" or "run_command";
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
                "keys" => 120,
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
                if (current is "high" or "xhigh")
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

