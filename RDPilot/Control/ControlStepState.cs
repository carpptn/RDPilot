namespace RDPilot.Control;

/// <summary>Values passed between stages of one step; stage-local values stay in their stage.</summary>
internal sealed class ControlStepState
{
    internal ControlStepState(int step)
    {
        if (step < 1) throw new ArgumentOutOfRangeException(nameof(step));
        this.step = step;
    }

    internal bool expectedContinuousIdle;
    internal IReadOnlyList<string>? plannedInputsForExecution;
    private UiPromptContext? _promptContext;
    internal UiPromptContext promptContext
    {
        get => _promptContext ?? throw new InvalidOperationException("Step promptContext has not been initialized by its producer stage.");
        set => _promptContext = value;
    }
    internal TurnInteractionContext? currentTurnBasedContext;
    internal string? previousResponseIdForRequest;
    internal string historyTail = "";
    internal Rectangle? appliedFocusRectForPrompt;
    internal Rectangle? focusUiaRectForPrompt;
    internal string? currentTurnFocusPath;
    internal string? currentTurnEvidenceDataUrl;
    internal string? currentTurnEvidencePath;
    internal bool hasPersistentTurnFocus;
    internal bool hasTurnVisualContext;
    private LoopDetectionAssessment? _loopAssessment;
    internal LoopDetectionAssessment loopAssessment
    {
        get => _loopAssessment ?? throw new InvalidOperationException("Step loopAssessment has not been initialized by its producer stage.");
        set => _loopAssessment = value;
    }
    internal bool proactiveVisualCycle;
    internal int visualCycleLength;
    private ResolvedActionSnapshot? _currentAction;
    internal ResolvedActionSnapshot currentAction
    {
        get => _currentAction ?? throw new InvalidOperationException("Step currentAction has not been initialized by its producer stage.");
        set => _currentAction = value;
    }
    internal string currentActionPolicy = "";
    internal bool actionExecutionFailed;
    internal bool failureRecorded;
    internal bool actionExecuted;
    internal bool actionWasLocallyRejected;
    internal bool turnBatchObserved;
    internal ObservedActionResult? lastObservedAction;
    internal bool executionHistoryRecorded;
    internal bool turnReanalysisWasRequired;
    internal ScreenObservationFrame? actionBaselineFrame;
    internal UiPromptContext? actionBaselineContext;
    internal byte[]? actionBaselineLocalFingerprint;
    internal string dataUrl = "";
    internal string? savedPath;
    internal int screenW;
    internal int screenH;
    internal int imageW;
    internal int imageH;
    internal string? focusUrl;
    internal Rectangle? focusUiaRect;
    internal string? focusUiaDataUrl;
    internal byte[] shotFingerprint = [];
    internal byte[] activeWindowFingerprint = [];
    private ScreenObservationFrame? _observationFrame;
    internal ScreenObservationFrame observationFrame
    {
        get => _observationFrame ?? throw new InvalidOperationException("Step observationFrame has not been initialized by its producer stage.");
        set => _observationFrame = value;
    }
    internal int cx;
    internal int cy;
    internal double cnx;
    internal double cny;
    internal ControlDecision? Decision { get; set; }
    internal ControlDecision decision => Decision ?? throw new InvalidOperationException("Execution requires a model decision.");
    internal ActionDto action { get => decision.Action; set => decision.Action = value; }

    internal readonly int step;
}

internal enum StepFlow
{
    Next,
    Retry,
    Stop
}
