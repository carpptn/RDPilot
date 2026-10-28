namespace RDPilot.Control;

/// <summary>Accounts for one decision, including failures of its follow-up actions.</summary>
internal static class ControlActionOutcome
{
    internal static StepFlow Fail(ControlRunState run, ControlStepState turn, string error)
    {
        turn.actionExecutionFailed = true;
        turn.decision.FollowUps.Clear();
        run.lastExecutorFailure = error;
        if (!turn.failureRecorded)
        {
            turn.failureRecorded = true;
            run.consecutiveActionFailures++;
            AddHistory(run.historyBuffer, $"[{turn.step}] EXECUTOR_FAILURE: {error}");
        }
        if (Options.MaxActionFailuresBeforeAbort == 0 ||
            run.consecutiveActionFailures < Options.MaxActionFailuresBeforeAbort)
            return StepFlow.Next;
        Console.WriteLine($"[guard] stopping after {run.consecutiveActionFailures} consecutive action failures.");
        run.runResult = new ControlRunResult(ControlRunOutcome.GuardStopped, turn.step, error);
        return StepFlow.Stop;
    }

    internal static void RecordHistory(StringBuilder history, string stepId, ResolvedActionSnapshot action, bool batched = false)
    {
        AddHistory(history, $"[{stepId}] {(batched ? "batch " : "")}{action.Description}");
        if (!string.IsNullOrWhiteSpace(action.Action.Note))
            AddHistory(history, $"[{stepId}] note: {action.Action.Note}");
    }

    internal static void RecordExecuted(ControlRunState run, string stepId, ResolvedActionSnapshot action, bool batched = false)
    {
        RecordHistory(run.historyBuffer, stepId, action, batched);
        run.observationActionGuard.RecordExecuted(action);
        RecordRecoveryAction(run.recoveryEpisode, run.recentExecutedActions, action, run.recoveryLessons);
    }

    internal static void ApplyExecuted(ControlRunState run, ControlStepState turn,
        ActionExecutionRecord execution, string stepId, bool batched = false)
    {
        RecordExecuted(run, stepId, execution.Action, batched);
        turn.currentAction = execution.Action;
        turn.action = execution.Action.Action;
        turn.actionExecuted = true;
        turn.actionBaselineFrame = execution.Before;
        turn.actionBaselineContext = execution.BeforeContext;
        turn.actionBaselineLocalFingerprint = execution.LocalBefore;
        turn.lastObservedAction = execution.Observation;
    }

    internal static StepFlow ApplyObservedBatch(ControlRunState run, ControlStepState turn, BatchedActionExecutionResult result)
    {
        for (var index = 0; index < result.Executions.Count; index++)
            ApplyExecuted(run, turn, result.Executions[index], $"{turn.step}.{index + 1}", batched: true);
        turn.executionHistoryRecorded = true;
        turn.turnBatchObserved = true;
        if (result.Cancelled)
            return Cancel(run, turn);
        if (result.Error is not null)
            return Fail(run, turn, result.Error);
        if (result.ExecutedActions.Count == 0)
            return Fail(run, turn, "observed sequence stopped before its first input");
        return StepFlow.Next;
    }

    internal static void Succeed(ControlRunState run)
    {
        run.consecutiveActionFailures = 0;
        run.lastExecutorFailure = null;
    }

    internal static StepFlow Cancel(ControlRunState run, ControlStepState turn)
    {
        turn.decision.FollowUps.Clear();
        run.runResult = new ControlRunResult(ControlRunOutcome.Cancelled, turn.step, "desktop action cancelled");
        return StepFlow.Stop;
    }
}
