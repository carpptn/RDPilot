internal static partial class RDPilotApplication
{
    /// <summary>
    /// Runs the goal-driven desktop control loop and its safety guards.
    /// </summary>
    internal static class ControlLoopService
    {
            // === Control loop ===
            internal static async Task RunOnce(string apiKey, string goal)
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
                try
                {
                    ResetRunMetrics();
                    PendingSafeActions.Clear();
                    Console.WriteLine($"Command ID: {commandId}");
                    Console.WriteLine($"Goal: {goal}");
                    Console.WriteLine("Loop start: one action -> screenshot -> next decision.");
                    Console.WriteLine("Emergency abort: Ctrl+Alt+Q\n");
        
                    if (AutoHideConsoleDuringRun || MinimizeConsoleDuringRun)
                        consoleHidden = ConcealConsoleWindow();
        
                    if (AllowHighLevelActions && TryExecuteFastGoal(goal))
                    {
                        Console.WriteLine("Finished (local fast path).");
                        return;
                    }
        
                    var systemRules = BuildSystemRules(); // stable per run; dynamic action availability is enforced by schema
                    var historyBuffer = new StringBuilder();
                    cancelCts = StartCancelHotkeyListener();
        
                    Rectangle? nextFocusRect = null; // crop/overlay after 'aim'/'point'/'request_crop'
                    Rectangle? lastAimRect = null;   // active AIM – required before clicks
        
                    // change / strategy metrics
                    byte[]? prevShotFingerprint = null;
                    ActionDto? prevAction = null;
                    string? lastSig = null;
                    int stagnationSteps = 0;
                    int repeatCount = 0;
                    double lastDelta = double.NaN;
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
        
                    for (int step = 1; step <= MaxSteps; step++)
                    {
                        if (CancelRequested) { Console.WriteLine("Aborted (hotkey)."); break; }
        
                        // screenshot at the beginning of a step (state after previous action)
                        var (dataUrl, savedPath, screenW, screenH, imageW, imageH, focusUrl, appliedFocusRect, focusUiaRect, focusUiaSummary, focusUiaDataUrl, focusUiaPath, shotFingerprint) =
                            ScreenshotToDataUrl(screensDir, commandId, step, nextFocusRect);
                        SetCurrentScreenMap(screenW, screenH, imageW, imageH);
                        Console.WriteLine($"[shot] {ShotLabel(savedPath, commandId, step)}");
                        if (CurrentScreenMap.IsScaled)
                            Console.WriteLine($"[coords] model image {imageW}x{imageH} -> screen {screenW}x{screenH}");
        
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
                            lastDelta = ComputeImageDelta(prevShotFingerprint, shotFingerprint); // 0..1
                            bool noChange = lastDelta < NoChangeThreshold;
                            bool previousWasObservationOnly = prevAction != null && IsLocalObservationAction(prevAction);
        
                            if (!previousWasObservationOnly)
                            {
                                if (noChange) stagnationSteps++; else stagnationSteps = 0;
                            }
                            else if (!noChange)
                            {
                                stagnationSteps = 0;
                            }
        
                            if (prevAction != null)
                            {
                                var sig = IneffectiveActionSignature(prevAction);
                                if (noChange && sig == lastSig) repeatCount++;
                                else { repeatCount = 0; lastSig = sig; }
                            }
        
                            // Expire AIM after a large visual change
                            if (lastAimRect is not null && lastDelta > AimExpireDelta)
                            {
                                Console.WriteLine($"[aim] expired (delta={lastDelta:0.###} > {AimExpireDelta:0.###})");
                                lastAimRect = null;
                            }
        
                            if (MaxStagnationStepsBeforeAbort > 0 && stagnationSteps >= MaxStagnationStepsBeforeAbort)
                            {
                                Console.WriteLine($"[guard] stopping: no visible progress for {stagnationSteps} consecutive step(s). Use --max-stagnation 0 to disable.");
                                break;
                            }
        
                            if (MaxRepeatedActionBeforeAbort > 0 && repeatCount >= MaxRepeatedActionBeforeAbort)
                            {
                                Console.WriteLine($"[guard] stopping: repeated ineffective action {repeatCount} time(s). Use --max-repeated-actions 0 to disable.");
                                break;
                            }
        
                            if (noChange && prevAction != null && IsPointClickAction(prevAction))
                            {
                                lastPrecisionHint = BuildPrecisionHint(prevAction, lastDelta, "Previous click produced little or no visible progress.");
                                lastPrecisionHintExpiresAfterStep = step + 1;
                            }
                            if (noChange && prevAction != null && IsTextInputAttemptAction(prevAction))
                            {
                                textInputNoChangeAttempts++;
                                lastTextInputHint = BuildTextInputHint(prevAction, lastDelta);
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
        
                        var reuseUiaTargets = ReuseUiaTargetsWhenScreenUnchanged &&
                                              !double.IsNaN(lastDelta) &&
                                              lastDelta < NoChangeThreshold &&
                                              CurrentUiaTargets.Count > 0;
                        PrepareUiaTargetsForPrompt(reuseUiaTargets, screenW, screenH);
        
                        RequestReasoningEffortOverride = EffectiveReasoningEffort(stagnationSteps, repeatCount);
        
                        // inject observation metrics into the prompt
                        var metaSb = new StringBuilder()
                            .AppendLine($"LAST_STEP_DELTA: {(double.IsNaN(lastDelta) ? "N/A" : lastDelta.ToString("0.####"))} (threshold={NoChangeThreshold})")
                            .AppendLine($"STAGNATION_STEPS: {stagnationSteps}")
                            .AppendLine($"REPEAT_COUNT: {repeatCount}")
                            .AppendLine($"REQUEST_REASONING_EFFORT: {RequestReasoningEffortOverride ?? ReasoningEffort ?? "default"}")
                            .AppendLine($"LAST_ACTION: {(prevAction == null ? "N/A" : Describe(prevAction))}")
                            .AppendLine($"AIM_ACTIVE: {(lastAimRect is null ? "false" : $"true {FormatImageRect(CurrentScreenMap.ScreenToImageRect(lastAimRect.Value))}")}");
                        if (repeatCount > 0 || stagnationSteps > 0)
                            metaSb.AppendLine("STRATEGY_HINT: The previous action did not visibly advance the screen. Do not repeat it; choose a different UI route or ask for a crop if the target is ambiguous.");
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
                                                       promptContext, reuseUiaTargets, previousResponseIdForRequest, omitFullScreenImage);
                        if (LogRequests)
                        {
                            var reqBodyForLog = BuildRequestBody_ForLog(Model, systemRules, goal, historyTail + "\n" + metaSb,
                                                                        omitFullScreenImage ? null : savedPath, imageW, imageH, cx, cy, cnx, cny,
                                                                        appliedFocusRect != null && LogScreens ? ScreenLogPath(screensDir, $"{commandId}_{step}_crop") : null,
                                                                        appliedFocusRectForPrompt,
                                                                        focusUiaRectForPrompt, focusUiaPath, promptContext, previousResponseIdForRequest, omitFullScreenImage);
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
                                break;
                            }
        
                            consecutiveModelFailures++;
                            if (LastOpenAiFailureWasRetriable &&
                                MaxModelFailuresBeforeAbort > 0 &&
                                consecutiveModelFailures < MaxModelFailuresBeforeAbort)
                            {
                                Console.WriteLine($"[openai] transient {LastOpenAiFailureKind}; keeping goal alive ({consecutiveModelFailures}/{MaxModelFailuresBeforeAbort}).");
                                AddHistory(historyBuffer, $"[{step}] model_failure_retry: {LastOpenAiFailureKind}");
                                await Task.Delay(750, cancelCts.Token);
                                continue;
                            }
        
                            Console.WriteLine("Could not parse action. Aborting this goal.");
                            break;
                        }
                        consecutiveModelFailures = 0;
        
                        Console.WriteLine($"[{step}] {Describe(action)}");
                        if (action.Confidence is double confidence)
                            Console.WriteLine($"     confidence: {confidence:0.##}");
                        if (!string.IsNullOrWhiteSpace(action.Note))
                            Console.WriteLine($"     note: {action.Note}");
        
                        var currentActionSignature = IneffectiveActionSignature(action);
                        nextFocusRect = null; // reset – set by aim/point/request_crop
                        var actionExecutionFailed = false;
                        try
                        {
                            if (IsActionOnCooldown(action, currentActionSignature, step, actionCooldownUntilStep, out var cooldownUntil))
                            {
                                Console.WriteLine($"[guard] skipping repeated ineffective action until step {cooldownUntil}: {Describe(action)}");
                                AddHistory(historyBuffer, $"[{step}] IGNORED (repeat_cooldown): {Describe(action)}");
                                lastExecutorFailure = $"action temporarily blocked after no visible effect: {Describe(action)}";
                                if (IsPointClickAction(action))
                                {
                                    lastPrecisionHint = BuildPrecisionHint(action, double.NaN, "A repeated click was blocked after no visible effect.");
                                    lastPrecisionHintExpiresAfterStep = step + 1;
                                }
                                actionExecutionFailed = true;
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
                                continue;
                            }
        
                            // ——— Mouse policy: global switch ———
                            if (IsMouseAction(action) && !MouseEnabled)
                            {
                                Console.WriteLine("[guard] mouse disabled → ignoring mouse action; use keyboard strategy or 'aim' without clicking.");
                                AddHistory(historyBuffer, $"[{step}] IGNORED (mouse_disabled)");
                            }
                            else if (action.Type == "aim")
                            {
                                var rect = ResolveAimRect(action);
                                if (rect is null) throw new InvalidOperationException("aim without parameters (bbox/crop/x/y/x_px/y_px).");
                                lastAimRect = rect.Value;
                                nextFocusRect = rect.Value; // show crop/overlay on next screenshot
                            }
                            else if (action.Type == "point")
                            {
                                // Visual pointer only – does NOT set AIM (clicks still blocked without 'aim')
                                var rect = ResolveCropRect(action);
                                if (rect is null) throw new InvalidOperationException("point without parameters.");
                                nextFocusRect = rect.Value;
                            }
                            else if (action.Type == "request_crop")
                            {
                                var rect = ResolveCropRect(action);
                                if (rect is null) throw new InvalidOperationException("request_crop without parameters.");
                                nextFocusRect = rect.Value;
                            }
                            else if (action.Type == "wait")
                            {
                                int secs = EffectiveWaitSeconds(action, out var requestedSecs);
                                if (secs < requestedSecs)
                                    Console.WriteLine($"[wait] Requested {requestedSecs}s capped to {secs}s.");
                                Console.WriteLine($"[wait] Sleeping {secs} s (long-running operation on screen)...");
                                await Task.Delay(secs * 1000, cancelCts.Token);
                            }
                            else if (action.Type == "done")
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
                                    var (freshDataUrl, freshPath, freshW, freshH, freshImageW, freshImageH, _, _, freshFocusRect, freshFocusSummary, _, _, _) = ScreenshotToDataUrl(screensDir, commandId, step, null);
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
                                try
                                {
                                    verify = ShouldVerifyGoal(goal, step, action)
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
                                    Console.WriteLine("Finished (model returned 'done').");
                                    break;
                                }
                                else
                                {
                                    Console.WriteLine($"[verify] ❌ Goal NOT confirmed. Reason: {verify?.Reason ?? "n/a"}");
                                    lastVerifierRejection = verify?.Reason ?? "verifier rejected done without a reason";
                                    AddHistory(historyBuffer, $"[{step}] done_rejected: {verify?.Reason}");
                                }
                            }
                            else if (action.Type is "click" or "double_click")
                            {
                                if (lastAimRect is null && DirectClickWithoutAim && HasExplicitPoint(action))
                                {
                                    Console.WriteLine("[guard] direct click without AIM allowed by profile.");
                                    ExecuteAction(action);
                                }
                                else if (lastAimRect is null)
                                {
                                    Console.WriteLine("[guard] click blocked: no active AIM. Return 'aim' first.");
                                    AddHistory(historyBuffer, $"[{step}] IGNORED (click_without_aim)");
                                }
                                else
                                {
                                    // Clicks must include explicit coordinates (don't default to AIM center)
                                    if (!HasExplicitPoint(action))
                                    {
                                        Console.WriteLine("[guard] click/double_click requires explicit coordinates (x/y or x_px/y_px) – provide an exact point within AIM.");
                                        AddHistory(historyBuffer, $"[{step}] IGNORED (click_missing_coords)");
                                        continue; // go to next round
                                    }
        
                                    var (xClick, yClick) = ResolveClickPoint(action, lastAimRect, logAdjustment: false);
        
                                    if (!lastAimRect.Value.Contains(xClick, yClick))
                                    {
                                        Console.WriteLine("[guard] click outside active AIM → ignoring. Set a proper 'aim' first.");
                                        AddHistory(historyBuffer, $"[{step}] IGNORED (click_outside_aim)");
                                    }
                                    else
                                    {
                                        ExecuteAction(action, lastAimRect);
                                    }
                                }
                            }
                            else
                            {
                                // Move/scroll/keys/type_text – execute normally
                                ExecuteAction(action);
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex is OperationCanceledException && CancelRequested)
                            {
                                Console.WriteLine("Aborted (hotkey).");
                                break;
                            }
                            Console.WriteLine($"Action execution error: {ex.Message}");
                            consecutiveActionFailures++;
                            actionExecutionFailed = true;
                            lastExecutorFailure = $"{Describe(action)} failed: {ex.Message}";
                            AddHistory(historyBuffer, $"[{step}] EXECUTOR_FAILURE: {lastExecutorFailure}");
                            PendingSafeActions.Clear();
                            if (MaxActionFailuresBeforeAbort > 0 && consecutiveActionFailures >= MaxActionFailuresBeforeAbort)
                            {
                                Console.WriteLine($"[guard] stopping: local action execution failed {consecutiveActionFailures} time(s). Use --max-action-failures 0 to disable.");
                                break;
                            }
                        }
        
                        if (!actionExecutionFailed)
                        {
                            consecutiveActionFailures = 0;
                            lastExecutorFailure = null;
                            AddHistory(historyBuffer, $"[{step}] {Describe(action)}");
                            if (!string.IsNullOrWhiteSpace(action.Note))
                                AddHistory(historyBuffer, $"[{step}] note: {action.Note}");
                            if (action.Type != "done" && lastVerifierRejection != null)
                                lastVerifierRejection = null;
        
                            await ExecuteQueuedSafeActionsAsync(historyBuffer, step, cancelCts.Token);
                        }
        
                        // Keep context for next step (delta/repeat metrics)
                        prevAction = action;
                        prevShotFingerprint = shotFingerprint;
                        if (!actionExecutionFailed && ActionRepeatCooldownSteps > 0 && repeatCount > 0 && !IsLocalObservationAction(action))
                            actionCooldownUntilStep[currentActionSignature] = step + ActionRepeatCooldownSteps;
        
                        if (CancelRequested) { Console.WriteLine("Aborted (hotkey)."); break; }
        
                        if (!actionExecutionFailed)
                        {
                            try { await WaitAfterActionAsync(action, shotFingerprint, cancelCts.Token); }
                            catch (OperationCanceledException) when (CancelRequested)
                            {
                                Console.WriteLine("Aborted (hotkey).");
                                break;
                            }
                        }
                    }
                }
                finally
                {
                    cancelCts?.Cancel();
                    cancelCts?.Dispose();
                    PrintRunMetrics();
                    Console.SetOut(prevOut);
                    Console.SetError(prevErr);
                    if (consoleHidden && RestoreConsoleAfterRun)
                        RestoreConsoleWindow();
                }
            }
        
            internal static bool IsMouseAction(ActionDto a)
                => a.Type is "move" or "click" or "double_click" or "scroll" or "focus_uia" or "click_uia";
        
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
        
            internal static string BuildPrecisionHint(ActionDto action, double delta, string reason)
            {
                var target = "";
                try
                {
                    var p = ResolvePoint(action);
                    var imagePoint = CurrentScreenMap.ScreenToImagePoint(p.X, p.Y);
                    target = $" Target was near SCREEN_SIZE ({imagePoint.X},{imagePoint.Y}).";
                }
                catch
                {
                    // Coordinate-free click actions are rare; keep the hint useful anyway.
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
        
            internal static bool IsActionOnCooldown(ActionDto action, string signature, int step, Dictionary<string, int> cooldowns, out int untilStep)
            {
                untilStep = 0;
                if (ActionRepeatCooldownSteps <= 0 || IsLocalObservationAction(action))
                    return false;
        
                if (!cooldowns.TryGetValue(signature, out untilStep))
                    return false;
        
                if (step <= untilStep)
                    return true;
        
                cooldowns.Remove(signature);
                return false;
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
        
            internal static async Task ExecuteQueuedSafeActionsAsync(StringBuilder historyBuffer, int step, CancellationToken cancellationToken)
            {
                var batchIndex = 0;
                while (PendingSafeActions.Count > 0)
                {
                    var action = PendingSafeActions.Dequeue();
                    batchIndex++;
                    Console.WriteLine($"[{step}.{batchIndex}] batch {Describe(action)}");
        
                    try
                    {
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
                    }
                    catch (Exception ex)
                    {
                        if (ex is OperationCanceledException && CancelRequested)
                        {
                            Console.WriteLine("Aborted (hotkey).");
                            PendingSafeActions.Clear();
                            return;
                        }
                        Console.WriteLine($"Batched action execution error: {ex.Message}");
                        PendingSafeActions.Clear();
                        return;
                    }
                }
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
        
            internal static string? EffectiveReasoningEffort(int stagnationSteps, int repeatCount)
            {
                if (!AdaptiveReasoningEffort || string.IsNullOrWhiteSpace(ReasoningEffort) || !SupportsReasoningEffort(Model))
                    return ReasoningEffort;
        
                var current = ReasoningEffort.ToLowerInvariant();
                if (current is "high" or "xhigh")
                    return ReasoningEffort;
        
                if (stagnationSteps >= 4 || repeatCount >= 3)
                    return "high";
        
                if (stagnationSteps >= 2 || repeatCount >= 1)
                    return "medium";
        
                return ReasoningEffort;
            }
    }
}

