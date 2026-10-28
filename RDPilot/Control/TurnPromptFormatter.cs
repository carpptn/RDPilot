namespace RDPilot.Control;

/// <summary>Formats a captured transition ledger without changing execution state.</summary>
internal static class TurnPromptFormatter
{
    const double MinimumStructuredPlanConfidence = TurnBasedTransitionTracker.MinimumStructuredPlanConfidence;

    internal sealed record CausalEvidence(
        string BaselineState, string TriggerAction, string ChangedState, string ChangeSummary,
        bool ReturnObserved, string? ReturnAction, bool ReenactmentObserved, string? ReenactmentAction);

    internal sealed record Snapshot(
        string? CurrentStateId,
        IReadOnlyList<TurnTransition> Transitions,
        IReadOnlyDictionary<string, Point> NavigationStatePositions,
        IReadOnlyCollection<TurnStateStore.NavigationEdge> NavigationEdges,
        string? LastChangeSummary,
        bool HasRecurrentActionRegions,
        int AuxiliarySignalTransitions,
        int ConsecutiveAuxiliarySignalTransitions,
        bool AuxiliarySignalChangedDuringKnownStateReturn,
        int? InferredAttemptInputLimit,
        int AttemptDirectionalInputs,
        int BoundedAttemptResetCount,
        bool LatestTransitionHadNoEffect,
        bool RequiresReanalysis,
        string? WorldStateSummary,
        bool PreviousProposalWasDiscardedAsStale,
        IReadOnlyList<string> PriorEpochMechanicsEvidence,
        IReadOnlyList<string> MechanicsHypothesisHistory,
        string? MechanicsHypothesis,
        string? SalientChangeObservation,
        TurnPromptFormatter.CausalEvidence? CausalEvent,
        bool CanProposeExecutionBatch,
        int AdvertisedMaxExecutionBatchLength);

    internal static string Build(Snapshot snapshot)
    {
        var (
            currentStateId,
            transitions,
            navigationStatePositions,
            navigationEdges,
            lastChangeSummary,
            hasRecurrentActionRegions,
            auxiliarySignalTransitions,
            consecutiveAuxiliarySignalTransitions,
            auxiliarySignalChangedDuringKnownStateReturn,
            inferredAttemptInputLimit,
            attemptDirectionalInputs,
            boundedAttemptResetCount,
            LatestTransitionHadNoEffect,
            RequiresReanalysis,
            worldStateSummary,
            previousProposalWasDiscardedAsStale,
            priorEpochMechanicsEvidence,
            mechanicsHypothesisHistory,
            mechanicsHypothesis,
            salientChangeObservation,
            causalEvent,
            CanProposeExecutionBatch,
            AdvertisedMaxExecutionBatchLength) = snapshot;

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
                             TurnInput.IsDirectionalSequence(transition.Action))
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
        AppendNavigationSummary(builder, currentStateId, navigationStatePositions, navigationEdges);
        if (!string.IsNullOrWhiteSpace(lastChangeSummary))
            builder.AppendLine($"TURN_VISUAL_CHANGE_REGIONS (percent of interaction region): {lastChangeSummary}");
        if (hasRecurrentActionRegions)
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

    static void AppendNavigationSummary(
        StringBuilder builder,
        string? currentStateId,
        IReadOnlyDictionary<string, Point> navigationStatePositions,
        IReadOnlyCollection<TurnStateStore.NavigationEdge> navigationEdges)
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

}
