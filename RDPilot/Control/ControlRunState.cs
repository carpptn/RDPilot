namespace RDPilot.Control;

/// <summary>State shared by consecutive control steps; initialized once by the run coordinator.</summary>
internal sealed class ControlRunState
{
    internal ControlRunState(string apiKey, string goal, string screensDir, string requestsDir)
    {
        this.commandId = Guid.NewGuid().ToString("N");
        this.screensDir = screensDir;
        this.requestsDir = requestsDir;
        this.apiKey = apiKey;
        this.goal = goal;
        this.runResult = new ControlRunResult(ControlRunOutcome.StepLimitReached, Options.MaxSteps, Options.MaxSteps > 0 ? $"configured step limit {Options.MaxSteps} was reached" : "continuous run ended unexpectedly");
        this.goalMode = ResolveGoalMode(goal, Options.GoalMode);
        this.controlContextChain = new ControlContextChain(commandId, Options.UsePreviousResponseState, Options.ControlContextCompactionEnabled, Options.ControlContextFallbackLimit);
        this.systemRules = BuildSystemRules();
        this.historyBuffer = new StringBuilder();
        this.recoveryLessons = [];
        this.recentExecutedActions = new();
        this.loopStateGraph = new LoopStateGraph { RunId = commandId };
        this.observationSession = new();
        this.observationActionGuard = new();
        this.shortTermPlan = new();
        this.turnBasedTransitions = new TurnBasedTransitionTracker(shortTermPlan);
        this.activeTurnChangeImages = [];
        this.recentIneffectiveSpatialActions = [];
        this.actionCooldownUntilStep = new(StringComparer.OrdinalIgnoreCase);
        this.spatialActionCooldowns = [];
        this.recentRejectedActions = new();
        this.recentRejectedProposalSignatures = new();
        recurringWorkflowIntent = string.Equals(goalMode, "continuous", StringComparison.Ordinal) && HasRecurringWorkflowIntent(goal);
        lastDelta = lastGlobalDelta = lastActiveWindowDelta = double.NaN;
    }

    internal readonly string commandId;
    internal readonly string screensDir;
    internal readonly string requestsDir;
    internal CancellationMonitor? cancelCts;
    internal readonly ControlContextChain controlContextChain;
    internal bool consoleHidden;
    internal int currentStep;
    internal ControlRunResult runResult;
    internal readonly string goalMode;
    internal bool recurringWorkflowIntent;
    internal readonly string systemRules;
    internal readonly StringBuilder historyBuffer;
    internal List<RecoveryLesson> recoveryLessons;
    internal Queue<ResolvedActionSnapshot> recentExecutedActions;
    internal LoopStateGraph loopStateGraph;
    internal readonly AdaptiveObservationSession observationSession;
    internal readonly ObservationActionGuardState observationActionGuard;
    internal readonly ShortTermPlanTracker shortTermPlan;
    internal readonly PromptVisualState promptVisualState = new();
    internal readonly TurnBasedTransitionTracker turnBasedTransitions;
    internal RecoveryEpisodeState? recoveryEpisode;
    internal Rectangle? nextFocusRect;
    internal Rectangle? lastAimRect => observationActionGuard.AimRect;
    internal Rectangle? turnBasedInteractionRect;
    internal bool turnBasedInteractionRegionIsAutomatic;
    internal bool turnBasedAutomaticRegionRefined;
    internal nint turnBasedInteractionWindow;
    internal TurnInteractionContext? turnBasedInteractionContext;
    internal string? previousTurnFocusDataUrl;
    internal string? previousTurnFocusPath;
    internal string? turnReferenceFocusDataUrl;
    internal string? turnReferenceFocusPath;
    internal IReadOnlyList<TurnChangeImagePair> activeTurnChangeImages;
    internal byte[]? prevShotFingerprint;
    internal byte[]? prevActiveWindowFingerprint;
    internal ScreenObservationFrame? prevObservationFrame;
    internal ObservedActionResult? prevObservedAction;
    internal int stateEpoch;
    internal UiPromptContext? prevObservationContext;
    internal ObservationAssessment? lastObservation;
    internal ResolvedActionSnapshot? prevAction;
    internal string? lastSig;
    internal List<ResolvedActionSnapshot> recentIneffectiveSpatialActions;
    internal int stagnationSteps;
    internal int repeatCount;
    internal int continuousIdleSteps;
    internal int ambiguousObservationSteps;
    internal double lastDelta;
    internal double lastGlobalDelta;
    internal double lastActiveWindowDelta;
    internal string? lastVerifierRejection;
    internal string? lastExecutorFailure;
    internal string? lastPrecisionHint;
    internal int lastPrecisionHintExpiresAfterStep;
    internal string? lastTextInputHint;
    internal int lastTextInputHintExpiresAfterStep;
    internal int textInputNoChangeAttempts;
    internal int textInputCooldownUntilStep;
    internal int consecutiveModelFailures;
    internal int consecutiveVerificationFailures;
    internal int consecutiveActionFailures;
    internal Dictionary<string, int> actionCooldownUntilStep;
    internal List<SpatialActionCooldown> spatialActionCooldowns;
    internal Queue<ResolvedActionSnapshot> recentRejectedActions;
    internal Queue<string> recentRejectedProposalSignatures;
    internal int rejectedProposalCycleCount;
    internal int rejectedProposalCycleLength;
    internal bool planningLoopDetected;
    internal ResolvedActionSnapshot? previousRejectedAction;
    internal readonly string apiKey;
    internal CancellationToken Cancellation => cancelCts?.Token ?? throw new InvalidOperationException("Control run has not started cancellation monitoring.");

    internal readonly string goal;
    internal void RegisterRejectedProposal(ResolvedActionSnapshot rejected)
    {
        this.previousRejectedAction = rejected;
        this.recentRejectedActions.Enqueue(rejected);
        while (this.recentRejectedActions.Count > 24)
            this.recentRejectedActions.Dequeue();
        this.recentRejectedProposalSignatures.Enqueue(rejected.IneffectiveSignature);
        while (this.recentRejectedProposalSignatures.Count > 32)
            this.recentRejectedProposalSignatures.Dequeue();
        var cycleLength = RepeatedStringCycleLength(this.recentRejectedProposalSignatures.ToArray());
        if (cycleLength <= 0)
        {
            this.rejectedProposalCycleCount = 0;
            this.rejectedProposalCycleLength = 0;
            this.planningLoopDetected = false;
            return;
        }

        this.rejectedProposalCycleCount = this.rejectedProposalCycleLength == cycleLength ? this.rejectedProposalCycleCount + 1 : 1;
        this.rejectedProposalCycleLength = cycleLength;
        this.planningLoopDetected = true;
        Console.WriteLine($"[loop] rejected-proposal cycle detected; length={cycleLength}; recurrence={this.rejectedProposalCycleCount}; action={rejected.Description}");
    }

    internal void ClearRejectedProposalLoop()
    {
        this.recentRejectedActions.Clear();
        this.recentRejectedProposalSignatures.Clear();
        this.rejectedProposalCycleCount = 0;
        this.rejectedProposalCycleLength = 0;
        this.planningLoopDetected = false;
        this.previousRejectedAction = null;
    }

    internal void SetAim(Rectangle region)
    {
        observationActionGuard.SetAim(region);
        nextFocusRect = region;
    }

    internal void InvalidateAim()
    {
        if (nextFocusRect == lastAimRect)
            nextFocusRect = null;
        observationActionGuard.InvalidateAim();
    }

    internal void ResetTurnInteraction()
    {
        BeginExternalStateEpoch();
        turnBasedInteractionRect = null;
        turnBasedInteractionRegionIsAutomatic = false;
        turnBasedAutomaticRegionRefined = false;
        turnBasedInteractionWindow = IntPtr.Zero;
        turnBasedInteractionContext = null;
        prevAction = null;
        prevShotFingerprint = prevActiveWindowFingerprint = null;
        prevObservationFrame = null;
        prevObservationContext = null;
        turnBasedTransitions.Reset();
        Console.WriteLine("[turn] cleared interaction evidence because the window or desktop context changed.");
    }

    internal void BeginExternalStateEpoch(ObservationAssessment? boundaryObservation = null)
    {
        this.stateEpoch++;
        this.observationActionGuard.Reset();
        this.nextFocusRect = null;
        this.prevObservedAction = null;
        this.loopStateGraph = new LoopStateGraph
        {
            RunId = this.commandId
        };
        this.recoveryEpisode = null;
        this.recentExecutedActions.Clear();
        this.recentIneffectiveSpatialActions.Clear();
        this.ClearRejectedProposalLoop();
        this.stagnationSteps = 0;
        this.repeatCount = 0;
        this.continuousIdleSteps = 0;
        this.ambiguousObservationSteps = 0;
        this.consecutiveActionFailures = 0;
        this.lastSig = null;
        this.lastObservation = boundaryObservation;
        this.lastDelta = double.NaN;
        this.lastGlobalDelta = double.NaN;
        this.lastActiveWindowDelta = double.NaN;
        this.actionCooldownUntilStep.Clear();
        this.spatialActionCooldowns.Clear();
        this.textInputNoChangeAttempts = 0;
        this.textInputCooldownUntilStep = 0;
        this.previousTurnFocusDataUrl = null;
        this.previousTurnFocusPath = null;
        this.turnReferenceFocusDataUrl = null;
        this.turnReferenceFocusPath = null;
        this.activeTurnChangeImages = [];
        this.turnBasedTransitions.BeginExternalStateEpoch();
        Console.WriteLine("[loop] external state boundary started a fresh visual and loop-detection epoch; local topology and temporal images were cleared while prior task-mechanics evidence was retained.");
    }
}
