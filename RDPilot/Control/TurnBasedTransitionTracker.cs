using static RDPilot.Control.TurnChangeRegions;

namespace RDPilot.Control;

            internal sealed class TurnBasedTransitionTracker
            {
                const int StateFingerprintSide = TurnVisualAnalysis.Side;
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
                const int RecurrentRegionLimit = 18;
                const int RecurrentRegionMaxAge = 12;

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

                readonly TurnStateStore stateStore = new();
                IReadOnlyList<TurnStateStore.Prototype> states => stateStore.States;
                IReadOnlyDictionary<string, Point> navigationStatePositions => stateStore.Positions;
                IReadOnlyCollection<TurnStateStore.NavigationEdge> navigationEdges => stateStore.Edges;
                readonly Queue<TurnTransition> transitions = new();
                readonly HashSet<string> observedDirectionalInputs = new(StringComparer.Ordinal);
                readonly Queue<double> ordinaryTransitionRatios = new();
                readonly List<RecurrentChangeRegion> recurrentActionRegions = [];
                readonly Queue<string> mechanicsHypothesisHistory = new();
                readonly Queue<string> salientObservationHistory = new();
                readonly Queue<string> priorEpochMechanicsEvidence = new();
                readonly DirectionalControlMemory directionalControls = new();
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
                    Math.Max(12, Options.TurnBasedMaxBatchInputs);

                internal int AdvertisedMaxExecutionBatchLength =>
                    Options.TurnBasedMaxBatchInputs;

                internal int MaxExecutionBatchLength =>
                    CanUseExecutionBatch ? Options.TurnBasedMaxBatchInputs : 0;

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
                    stateStore.Clear();
                    transitions.Clear();
                    observedDirectionalInputs.Clear();
                    ordinaryTransitionRatios.Clear();
                    recurrentActionRegions.Clear();
                    mechanicsHypothesisHistory.Clear();
                    salientObservationHistory.Clear();
                    priorEpochMechanicsEvidence.Clear();
                    directionalControls.Clear();
                    stateStore.ClearNavigation();
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
                    stateStore.Clear();
                    if (current is not null)
                        stateStore.Add(current, currentStateId);
                    transitions.Clear();
                    observedDirectionalInputs.Clear();
                    ordinaryTransitionRatios.Clear();
                    recurrentActionRegions.Clear();
                    mechanicsHypothesisHistory.Clear();
                    salientObservationHistory.Clear();
                    stateStore.ClearNavigation();
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
                        stateStore.Anchor(currentStateId);
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

                    EnsureVolatileMask(TurnVisualAnalysis.PixelCount(fingerprint));
                    if (currentStateId is null)
                    {
                        currentStateId = ResolveState(fingerprint, out _);
                        stateStore.Anchor(currentStateId);
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
                        stateStore.Anchor(currentStateId);
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

                    EnsureVolatileMask(TurnVisualAnalysis.PixelCount(fingerprint));
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

                    EnsureVolatileMask(TurnVisualAnalysis.PixelCount(beforeFingerprint));
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
                        TurnInput.IsDirectionalSequence(actionLabel) &&
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
                    directionalControls.Remember(action, actionLabel);
                    var planExpectedAction =
                        shortTermPlan.IsExpectedInput(actionLabel, beforeState);
                    if (TurnInput.IsDirectionalSequence(actionLabel))
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
                    else if (planExpectedAction && TurnInput.IsDirectionalSequence(actionLabel))
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
                    directionalControls.Remember(action, actionLabel);
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
                        !BatchPlanner.TryGetObservedTurnInputLabel(
                            firstAction,
                            out var firstLabel))
                    {
                        return false;
                    }

                    directionalControls.Remember(firstAction, firstLabel);
                    var result = new List<ActionDto>(remainingInputs.Count);
                    foreach (var input in remainingInputs)
                    {
                        if (!BatchPlanner.TryNormalizeDirectionalLabel(
                                input,
                                out var label))
                        {
                            break;
                        }

                        var template = directionalControls.Recall(label);
                        if (template is null) break;
                        result.Add(template);
                    }

                    if (result.Count == 0)
                        return false;
                    followUps = result.ToArray();
                    return true;
                }

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

                    EnsureVolatileMask(TurnVisualAnalysis.PixelCount(beforeFingerprint));
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
                                 (!TurnInput.IsDirectionalSequence(actionLabel) ||
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

                internal string BuildPromptSummary() =>
                    TurnPromptFormatter.Build(new(
                        CurrentStateId: currentStateId,
                        Transitions: transitions.ToArray(),
                        NavigationStatePositions: new Dictionary<string, Point>(navigationStatePositions),
                        NavigationEdges: navigationEdges.ToArray(),
                        LastChangeSummary: lastChangeSummary,
                        HasRecurrentActionRegions: recurrentActionRegions.Any(pattern => pattern.Hits >= 2),
                        AuxiliarySignalTransitions: auxiliarySignalTransitions,
                        ConsecutiveAuxiliarySignalTransitions: consecutiveAuxiliarySignalTransitions,
                        AuxiliarySignalChangedDuringKnownStateReturn: auxiliarySignalChangedDuringKnownStateReturn,
                        InferredAttemptInputLimit: inferredAttemptInputLimit,
                        AttemptDirectionalInputs: attemptDirectionalInputs,
                        BoundedAttemptResetCount: boundedAttemptResetCount,
                        LatestTransitionHadNoEffect: LatestTransitionHadNoEffect,
                        RequiresReanalysis: RequiresReanalysis,
                        WorldStateSummary: worldStateSummary,
                        PreviousProposalWasDiscardedAsStale: previousProposalWasDiscardedAsStale,
                        PriorEpochMechanicsEvidence: priorEpochMechanicsEvidence.ToArray(),
                        MechanicsHypothesisHistory: mechanicsHypothesisHistory.ToArray(),
                        MechanicsHypothesis: mechanicsHypothesis,
                        SalientChangeObservation: salientChangeObservation,
                        CausalEvent: causalEvent is null ? null : new(
                            causalEvent.BaselineState, causalEvent.TriggerAction,
                            causalEvent.ChangedState, causalEvent.ChangeSummary,
                            causalEvent.ReturnObserved, causalEvent.ReturnAction,
                            causalEvent.ReenactmentObserved, causalEvent.ReenactmentAction),
                        CanProposeExecutionBatch: CanProposeExecutionBatch,
                        AdvertisedMaxExecutionBatchLength: AdvertisedMaxExecutionBatchLength));

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
                           TurnInput.IsDirectionalSequence(recent[0].Action) &&
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
                            TurnInput.IsDirectionalSequence(transition.Action) &&
                            string.Equals(
                                transition.Result,
                                "changed",
                                StringComparison.Ordinal) &&
                            !string.Equals(
                                transition.From,
                                transition.To,
                                StringComparison.Ordinal)) >= requiredCount;

                bool IsPredictableDirectionalAction(string actionLabel) =>
                    TurnInput.IsDirectionalSequence(actionLabel) &&
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
                    if (!TurnInput.IsDirectionalSequence(actionLabel) ||
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
                    TurnStateStore.Prototype? best = null;
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
                    stateStore.Add(new(id, fingerprint), currentStateId);
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

                    if (TurnInput.IsDirectionalSequence(actionLabel) &&
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
                    if (!TurnInput.IsDirectionalSequence(actionLabel))
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

                void RecordNavigationTransition(string beforeState, string actionLabel, string afterState, string result) =>
                    stateStore.RecordTransition(beforeState, actionLabel, afterState, result, currentStateId);

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
                    var activeTiles = TurnVisualAnalysis.ChangedTiles(earlier, later, volatilePixels, ChangeTileSide);

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


                static bool HasPersistentVisualChange(ChangeAnalysis analysis) =>
                    analysis.Regions.Count > 0 ||
                    analysis.ChangedRatio >= 0.00025 ||
                    analysis.MeanDelta >= 0.00005;

                bool HasDirectionalMotionEvidence(
                    string actionLabel,
                    ChangeAnalysis analysis)
                {
                    if (!TurnInput.IsDirectionalSequence(actionLabel))
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

                void LearnAmbientPixels(byte[] earlier, byte[] later) =>
                    TurnVisualAnalysis.LearnAmbient(earlier, later, volatilePixels);

                (double MeanDelta, double ChangedRatio) StateDifference(byte[] left, byte[] right) =>
                    TurnVisualAnalysis.Compare(left, right, volatilePixels);

                (double MeanDelta, double ChangedRatio) StateIdentityDifference(byte[] left, byte[] right) =>
                    TurnVisualAnalysis.Compare(left, right, volatilePixels, auxiliaryStatePixels);

                static bool IsSameState(
                    (double MeanDelta, double ChangedRatio) difference,
                    double meanThreshold = StateMatchMeanThreshold,
                    double changedRatioThreshold = StateMatchChangedRatioThreshold) =>
                    difference.MeanDelta <= meanThreshold &&
                    difference.ChangedRatio <= changedRatioThreshold;

                static byte[] ExtractRegionFingerprint(ScreenObservationFrame frame, Rectangle region) =>
                    TurnVisualAnalysis.Extract(frame, region);

                string TurnActionLabel(ResolvedActionSnapshot snapshot)
                {
                    var action = snapshot.Action;
                    if (TurnInput.Key(action.ResolvedTurnInputLabel) is string resolvedTurnInputLabel)
                    {
                        return resolvedTurnInputLabel;
                    }
                    if (action.Type == "keys" && action.Keys is { Length: > 0 })
                        return string.Join(",", action.Keys.Select(CanonicalKeyLabel));
                    if (action.Type is "click" or "double_click" or "click_uia" or "focus_uia")
                    {
                        if (directionalControls.TryResolve(snapshot, out var templateLabel))
                            return templateLabel;
                        var target = string.IsNullOrWhiteSpace(action.Note)
                            ? snapshot.SemanticTokens
                            : action.Note;
                        if (BatchPlanner.TryNormalizeUnambiguousDirectionalLabel(
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

                static string CanonicalKeyLabel(string key) =>
                    TurnInput.Key(key) ?? key.Trim();

            }
