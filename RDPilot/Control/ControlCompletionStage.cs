namespace RDPilot.Control;

/// <summary>Records action outcomes, executes allowed follow-ups and waits for UI settling.</summary>
internal static class ControlCompletionStage
{
    internal static async Task<StepFlow> RunAsync(ControlRunState run, ControlStepState turn)
    {
        if (turn.actionWasLocallyRejected)
        {
            run.RegisterRejectedProposal(turn.currentAction);
            if (ControlActionOutcome.Fail(run, turn, run.lastExecutorFailure ?? "local action rejected") == StepFlow.Stop)
                return StepFlow.Stop;
        }

        if (!turn.actionExecutionFailed && turn.actionExecuted)
        {
            if (string.Equals(turn.currentActionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
                IsStateChangingInteractionAction(turn.currentAction.Action) &&
                turn.turnReanalysisWasRequired)
            {
                run.turnBasedTransitions.AcknowledgeReanalysis();
                if (!run.turnBasedTransitions.HasActiveCausalEvent)
                    run.activeTurnChangeImages = [];
            }

            if (!turn.executionHistoryRecorded)
                ControlActionOutcome.ApplyExecuted(run, turn, new(turn.currentAction,
                    turn.actionBaselineFrame ?? turn.observationFrame, turn.actionBaselineLocalFingerprint,
                    BeforeContext: turn.actionBaselineContext ?? turn.promptContext), turn.step.ToString());
            if (turn.action.Type != "done" && run.lastVerifierRejection != null)
                run.lastVerifierRejection = null;
            var batchResult = await ExecuteQueuedSafeActionsAsync(
                turn.decision.FollowUps,
                turn.step,
                turn.action,
                (turn.actionBaselineFrame ?? turn.observationFrame).GlobalFingerprint,
                run.Cancellation,
                snapshot => ActionExecutionPolicy.EnsureAllowed(run, turn.step, snapshot, turn.currentActionPolicy));
            var batchIndex = 0;
            foreach (var execution in batchResult.Executions)
            {
                ControlActionOutcome.ApplyExecuted(run, turn, execution, $"{turn.step}.{++batchIndex}", batched: true);
            }

            if (batchResult.ExecutedActions.Count > 0)
            {
                turn.currentActionPolicy = run.observationSession.ResolveActionPolicy(turn.currentAction, turn.promptContext, run.goal);
            }

            if (batchResult.Cancelled)
                return ControlActionOutcome.Cancel(run, turn);
            if (!string.IsNullOrWhiteSpace(batchResult.Error))
            {
                if (ControlActionOutcome.Fail(run, turn, batchResult.Error) == StepFlow.Stop)
                    return StepFlow.Stop;
            }
            else
                ControlActionOutcome.Succeed(run);
        }

        // Keep context for next step (delta/repeat metrics)
        run.prevObservedAction = turn.actionExecutionFailed || !turn.actionExecuted ? null : turn.lastObservedAction;
        run.prevAction = turn.actionExecutionFailed || !turn.actionExecuted ? null : turn.currentAction;
        run.prevShotFingerprint = turn.actionExecutionFailed || !turn.actionExecuted

            ? null

            : (turn.actionBaselineFrame ?? turn.observationFrame).GlobalFingerprint;
        run.prevActiveWindowFingerprint = turn.actionExecutionFailed || !turn.actionExecuted

            ? null

            : (turn.actionBaselineFrame ?? turn.observationFrame).ActiveWindowFingerprint;
        run.prevObservationFrame = turn.actionExecutionFailed || !turn.actionExecuted ? null : turn.actionBaselineFrame ?? turn.observationFrame;
        run.prevObservationContext = turn.actionExecutionFailed || !turn.actionExecuted ? null : turn.actionBaselineContext ?? turn.promptContext;
        if (CancelRequested || run.Cancellation.IsCancellationRequested)
        {
            return ControlActionOutcome.Cancel(run, turn);
        }

        if (!turn.actionExecutionFailed && turn.actionExecuted && !turn.turnBatchObserved)
        {
            try
            {
                _ = await WaitAfterActionAsync(
                    turn.currentAction.Action,
                    (turn.actionBaselineFrame ?? turn.observationFrame).GlobalFingerprint,
                    turn.currentActionPolicy,
                    turn.actionBaselineLocalFingerprint,
                    turn.currentAction.ObservationRegion,
                    run.Cancellation);
            }
            catch (OperationCanceledException) when (CancelRequested || run.Cancellation.IsCancellationRequested)
            {
                return ControlActionOutcome.Cancel(run, turn);
            }
        }

        return StepFlow.Next;
    }
}
