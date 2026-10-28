namespace RDPilot.Control;

/// <summary>Resolves completion proposals, keeping verifier failures separate from negative verdicts.</summary>
internal static class GoalCompletionService
{
    internal static async Task<StepFlow> RunAsync(ControlRunState run, ControlStepState turn)
    {
        if (CancelRequested || run.Cancellation.IsCancellationRequested)
            return ControlActionOutcome.Cancel(run, turn);

        if (run.goalMode == "continuous")
            return Reject(run, turn, "The goal is continuous and has no natural completion. Continue meaningful activity; do not mark it complete. Termination is governed by user abort and configured runtime safety guards or limits.");

        var independentlyVerified = ShouldVerifyGoal(run.goal, turn.step, turn.action);
        var verify = independentlyVerified
            ? await VerifyAsync(run, turn)
            : new VerifyDto { Verdict = "yes", Reason = "verification skipped by mode" };

        // A response can arrive while the emergency hotkey is being processed.
        if (CancelRequested || run.Cancellation.IsCancellationRequested)
            return ControlActionOutcome.Cancel(run, turn);

        if (verify?.Verdict?.Equals("yes", StringComparison.OrdinalIgnoreCase) == true)
        {
            run.consecutiveVerificationFailures = 0;
            if (independentlyVerified)
            {
                Console.WriteLine($"[verify] ✅ Goal independently confirmed: {verify.Reason}");
                AddHistory(run.historyBuffer, $"[{turn.step}] done_verified");
            }
            else
            {
                var confidence = turn.action.Confidence?.ToString("0.00") ?? "n/a";
                Console.WriteLine($"[verify] ⏭ Independent verification skipped; {VerificationSkipReason(turn.step, turn.action)}; accepting model done (confidence={confidence}).");
                AddHistory(run.historyBuffer, $"[{turn.step}] done_accepted_unverified");
            }
            run.lastVerifierRejection = null;
            run.InvalidateAim();
            run.recoveryEpisode = ConfirmPendingRecovery(run.recoveryEpisode, run.recoveryLessons, independentlyVerified);
            Console.WriteLine("Finished (model returned 'done').");
            run.runResult = new ControlRunResult(ControlRunOutcome.Completed, turn.step,
                independentlyVerified ? verify.Reason ?? "the model completion was independently verified"
                    : "model completion was accepted without independent verification");
            return StepFlow.Stop;
        }

        if (verify?.Verdict?.Equals("no", StringComparison.OrdinalIgnoreCase) == true)
        {
            run.consecutiveVerificationFailures = 0;
            return Reject(run, turn, verify.Reason ?? "verifier rejected done without a reason");
        }

        // Controller successes must not reset a series of failed verifier calls.
        run.consecutiveVerificationFailures++;
        turn.actionExecutionFailed = true;
        turn.decision.FollowUps.Clear();
        var failure = string.IsNullOrWhiteSpace(CurrentRun.LastOpenAiFailureKind)
            ? "invalid_verifier_response" : CurrentRun.LastOpenAiFailureKind;
        var message = $"Goal verification failed: {failure}";
        Console.WriteLine($"[verify] {message}");
        AddHistory(run.historyBuffer, $"[{turn.step}] verification_failure: {failure}");
        if (CurrentRun.LastOpenAiFailureWasRetriable &&
            (Options.MaxModelFailuresBeforeAbort == 0 ||
             run.consecutiveVerificationFailures < Options.MaxModelFailuresBeforeAbort))
        {
            run.lastVerifierRejection = $"Completion could not be verified because of a transient verifier failure ({failure}). Retry verification before claiming success.";
            await RuntimeClock.Delay(750, run.Cancellation);
            return StepFlow.Next;
        }

        run.runResult = new ControlRunResult(ControlRunOutcome.Failed, turn.step, message);
        return StepFlow.Stop;
    }

    static StepFlow Reject(ControlRunState run, ControlStepState turn, string reason)
    {
        run.lastVerifierRejection = reason;
        Console.WriteLine($"[verify] ❌ Goal NOT confirmed. Reason: {reason}");
        AddHistory(run.historyBuffer, $"[{turn.step}] done_rejected: {reason}");
        turn.actionExecutionFailed = true;
        turn.decision.FollowUps.Clear();
        run.RegisterRejectedProposal(turn.currentAction);
        return StepFlow.Next;
    }

    static async Task<VerifyDto?> VerifyAsync(ControlRunState run, ControlStepState turn)
    {
        var dataUrl = turn.dataUrl;
        var savedPath = turn.savedPath;
        var screenW = turn.screenW;
        var screenH = turn.screenH;
        var focusRect = turn.focusUiaRect;
        var context = turn.promptContext;
        if (Options.RefreshScreenshotBeforeVerify)
        {
            await RuntimeClock.Delay(Options.UiSettleDelayMs, run.Cancellation);
            var capture = ScreenshotToDataUrl(run.screensDir, run.commandId, turn.step, null);
            dataUrl = capture.DataUrl;
            savedPath = capture.SavedPath;
            screenW = capture.ScreenWidth;
            screenH = capture.ScreenHeight;
            focusRect = capture.FocusUiaRect;
            context = CaptureUiPromptContext(capture.FocusUiaSummary, screenW, screenH);
        }

        var prepared = PrepareHelperImage(dataUrl, Options.VerifyScreenshotMaxWidth);
        var previousScreenMap = CurrentRun.CurrentScreenMap;
        try
        {
            SetCurrentScreenMap(screenW, screenH, prepared.Width, prepared.Height);
            return await VerifyGoalAsync(run.apiKey, run.goal, prepared.DataUrl, savedPath,
                prepared.Width, prepared.Height, ScreenRectToImage(focusRect), context,
                run.requestsDir, run.commandId, turn.step, run.Cancellation);
        }
        finally { CurrentRun.CurrentScreenMap = previousScreenMap; }
    }
}
