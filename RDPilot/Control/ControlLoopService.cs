namespace RDPilot.Control;

/// <summary>
/// Runs the goal-driven desktop control loop and its safety guards.
/// </summary>
internal static class ControlLoopService
{
    // === Control loop ===
    internal static async Task<ControlRunResult> RunOnce(string apiKey, string goal)
    {
        using var runScope = BeginRun();
        var run = new ControlRunState(apiKey, goal, EnsureScreensDir(), EnsureRequestsDir());
        var logPath = Path.Combine(EnsureLogDir(), $"{run.commandId}.log");
        using var logSession = new RunLogSession(logPath);
        try
        {
            ResetRunMetrics();
            Console.WriteLine($"Command ID: {run.commandId}");
            Console.WriteLine($"Goal: {run.goal}");
            Console.WriteLine($"Goal mode: {run.goalMode}");
            if (run.recurringWorkflowIntent)
                Console.WriteLine("Goal cycle policy: recurring workflow may contain productive state returns.");
            Console.WriteLine("Loop start: one action -> screenshot -> next decision.");
            Console.WriteLine("Emergency abort: Ctrl+Alt+Q\n");
            if (!Options.BatchMode &&
                (Options.AutoHideConsoleDuringRun || Options.MinimizeConsoleDuringRun))
                run.consoleHidden = ConcealConsoleWindow();
            if (Options.AllowHighLevelActions && TryExecuteFastGoal(run.goal))
            {
                Console.WriteLine("Finished (local fast path).");
                run.runResult = new ControlRunResult(ControlRunOutcome.Completed, 0, "completed through the local fast path");
                return run.runResult;
            }

            run.controlContextChain.LogStart(Options.Model);
            run.recoveryLessons = LoadRecoveryLessons();
            run.observationSession.LogInitialProfile();
            run.cancelCts = StartCancelHotkeyListener();
            for (int step = 1; Options.MaxSteps == 0 || step <= Options.MaxSteps; step++)
            {
                var turn = new ControlStepState(step);
                var observationFlow = await ControlObservationStage.RunAsync(run, turn);
                if (observationFlow == StepFlow.Stop)
                    break;
                if (observationFlow == StepFlow.Retry)
                    continue;
                var decisionFlow = await ControlDecisionStage.RunAsync(run, turn);
                if (decisionFlow == StepFlow.Stop)
                    break;
                if (decisionFlow == StepFlow.Retry)
                    continue;
                var executionFlow = await ControlExecutionStage.RunAsync(run, turn);
                if (executionFlow == StepFlow.Stop)
                    break;
                if (executionFlow == StepFlow.Retry)
                    continue;
                var completionFlow = await ControlCompletionStage.RunAsync(run, turn);
                if (completionFlow == StepFlow.Stop)
                    break;
                if (completionFlow == StepFlow.Retry)
                    continue;
            }
        }
        catch (OperationCanceledException) when (CancelRequested)
        {
            Console.WriteLine("Aborted (hotkey).");
            run.runResult = new ControlRunResult(ControlRunOutcome.Cancelled, run.currentStep, "cancelled with the emergency hotkey");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Control run failed unexpectedly: {ex}");
            run.runResult = new ControlRunResult(ControlRunOutcome.Failed, run.currentStep, ex.Message);
        }
        finally
        {
            ReleaseAllHeldInputs();
            if (run.cancelCts is not null)
                await run.cancelCts.DisposeAsync();
            _ = FlushPendingRecoveryMemory();
            if (Options.LoopReplayAutoExportEnabled && Options.RecoveryMemoryEnabled)
                TryAutoExportLoopReplayCorpus();
            run.controlContextChain?.LogClose(run.runResult.Outcome);
            PrintRunMetrics();
            if (run.consoleHidden && Options.RestoreConsoleAfterRun)
                RestoreConsoleWindow();
        }

        return run.runResult with
        {
            Step = run.currentStep
        };
    }
}
