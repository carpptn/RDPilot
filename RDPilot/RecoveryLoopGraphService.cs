internal static partial class RDPilotApplication
{
    /// <summary>
    /// Detects repeated desktop states with a bounded action-labelled graph.
    /// </summary>
    internal static partial class RecoveryMemoryService
    {
        internal static LoopDetectionAssessment AssessVisualStateCycle(
            LoopStateGraph graph,
            byte[] currentFingerprint,
            byte[] currentActiveWindowFingerprint,
            UiPromptContext context,
            int currentStep,
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot? previousAction,
            double lastDelta,
            bool recordLearning = true,
            string goalMode = "finite",
            bool recurringWorkflowIntent = false)
        {
            if (graph.PendingCandidateStep is int pendingCandidateStep &&
                currentStep - pendingCandidateStep >
                Math.Max(2, GraphCandidateTtlSteps))
            {
                if (recordLearning)
                {
                    var expiredAssessment = new LoopDetectionAssessment(
                        false,
                        graph.PendingCandidateConfidence,
                        0,
                        graph.PendingCandidateVisitCount,
                        false,
                        false,
                        "candidate expired before an independent recurrence supplied a reliable label")
                    {
                        LoopTopology = graph.PendingCandidateTopology,
                        InteractionDomain = graph.PendingCandidateDomain,
                        DecisionThreshold = CalibratedLoopThreshold(
                            graph.PendingCandidateCalibrationKey,
                            MultiStepStateCycle),
                        CycleDisposition = "inconclusive",
                        CalibrationKey =
                            graph.PendingCandidateCalibrationKey,
                        RunId = graph.RunId
                    };
                    RegisterCalibrationInconclusive(
                        graph.PendingCandidateCalibrationKey);
                    AppendLoopTelemetry(
                        "graph_candidate_expired",
                        currentStep,
                        MultiStepStateCycle,
                        graph.PendingCandidateTopology,
                        graph.PendingCandidateDomain,
                        expiredAssessment,
                        null,
                        new
                        {
                            labelSource = "censored_candidate_ttl",
                            candidateWasActionable =
                                graph.PendingCandidateWasActionable
                        });
                    PersistCalibrationSnapshot();
                }
                ClearPendingGraphCandidate(graph);
            }

            var process = NormalizeText(context.ActiveProcessName);
            var titleTokens = NormalizeTokens(context.ActiveWindowTitle);
            var focusTokens = NormalizeTokens(context.FocusedUiaSummary);
            var matchingNode = graph.Nodes
                .Where(node =>
                    string.Equals(node.ActiveProcess, process, StringComparison.Ordinal) &&
                    ContextTokensCompatible(node.WindowTitleTokens, titleTokens) &&
                    ContextTokensCompatible(node.FocusTokens, focusTokens, minimum: 0.20) &&
                    node.ScreenFingerprint.Length == currentFingerprint.Length &&
                    node.ActiveWindowFingerprint.Length == currentActiveWindowFingerprint.Length)
                .Select(node => (
                    Node: node,
                    GlobalDelta: ComputeStabilityWeightedDelta(
                        node.ScreenFingerprint,
                        currentFingerprint,
                        node.ScreenInstability),
                    ActiveDelta: ComputeStabilityWeightedDelta(
                        node.ActiveWindowFingerprint,
                        currentActiveWindowFingerprint,
                        node.ActiveWindowInstability)))
                .Where(item =>
                    item.ActiveDelta < NoChangeThreshold &&
                    item.GlobalDelta < NoChangeThreshold * 3)
                .OrderBy(item => 0.75 * item.ActiveDelta + 0.25 * item.GlobalDelta)
                .FirstOrDefault();

            var isExistingNode = matchingNode.Node is not null;
            var node = matchingNode.Node ?? new LoopStateNode
            {
                Id = graph.NextNodeId++,
                ScreenFingerprint = currentFingerprint.ToArray(),
                ActiveWindowFingerprint = currentActiveWindowFingerprint.ToArray(),
                ActiveProcess = process,
                WindowTitleTokens = titleTokens,
                FocusTokens = focusTokens
            };
            if (!isExistingNode)
                graph.Nodes.Add(node);

            var priorStep = node.LastSeenStep;
            var priorVisitCount = node.VisitCount;
            var previousNodeId = graph.LastNodeId;
            var lastPathIndex = graph.RecentNodePath.FindLastIndex(id => id == node.Id);
            var departedFromState = lastPathIndex >= 0 &&
                                    graph.RecentNodePath
                                        .Skip(lastPathIndex + 1)
                                        .Any(id => id != node.Id);
            var cycleLength = priorStep > 0 ? currentStep - priorStep : 0;
            var plausibleDistance = cycleLength is >= 2 and <= 24;
            var graphCycle = previousNodeId is int previousId &&
                             previousId != node.Id &&
                             GraphHasPath(graph, node.Id, previousId, maxDepth: 16);

            if (previousNodeId is int fromId)
                AddOrUpdateGraphEdge(graph, fromId, node.Id, previousAction, currentStep);

            node.VisitCount++;
            node.LastSeenStep = currentStep;
            node.RecentVisitSteps.Add(currentStep);
            if (node.RecentVisitSteps.Count > 6)
                node.RecentVisitSteps.RemoveAt(0);
            node.ScreenInstability = UpdateInstabilityMask(
                node.ScreenFingerprint,
                currentFingerprint,
                node.ScreenInstability);
            node.ActiveWindowInstability = UpdateInstabilityMask(
                node.ActiveWindowFingerprint,
                currentActiveWindowFingerprint,
                node.ActiveWindowInstability);
            // Slowly adapt a node prototype while keeping state clusters stable.
            if (matchingNode.ActiveDelta < NoChangeThreshold * 0.5)
            {
                node.ScreenFingerprint = currentFingerprint.ToArray();
                node.ActiveWindowFingerprint = currentActiveWindowFingerprint.ToArray();
            }
            graph.RecentNodePath.Add(node.Id);
            if (graph.RecentNodePath.Count > 32)
                graph.RecentNodePath.RemoveAt(0);
            graph.LastNodeId = node.Id;
            PruneStateGraph(graph);

            var returnedToPriorState = isExistingNode && departedFromState && plausibleDistance;
            var matchingSteps = node.RecentVisitSteps.ToArray();
            var consistentReturnPeriod = matchingSteps.Length >= 3 &&
                                         Math.Abs(
                                             (matchingSteps[^1] - matchingSteps[^2]) -
                                             (matchingSteps[^2] - matchingSteps[^3])) <= 1;

            var actions = recentActions.TakeLast(16).ToList();
            if (previousAction != null &&
                (actions.Count == 0 || !ReferenceEquals(actions[^1], previousAction)))
            {
                actions.Add(previousAction);
            }
            var repeatedActionCycle = RepeatedActionCycleLength(actions) > 0;
            var semanticStateKey = BuildSemanticStateKey(context, previousAction);
            graph.SemanticVisitSteps.TryGetValue(semanticStateKey, out var semanticVisits);
            semanticVisits ??= [];
            var priorSemanticStep = semanticVisits.Count > 0 ? semanticVisits[^1] : 0;
            var lastSemanticIndex = graph.RecentSemanticPath.FindLastIndex(key =>
                string.Equals(key, semanticStateKey, StringComparison.Ordinal));
            var departedFromSemanticState = lastSemanticIndex >= 0 &&
                                            graph.RecentSemanticPath
                                                .Skip(lastSemanticIndex + 1)
                                                .Any(key => !string.Equals(
                                                    key,
                                                    semanticStateKey,
                                                    StringComparison.Ordinal));
            var semanticCycleLength = priorSemanticStep > 0
                ? currentStep - priorSemanticStep
                : 0;
            semanticVisits.Add(currentStep);
            if (semanticVisits.Count > 6)
                semanticVisits.RemoveAt(0);
            graph.SemanticVisitSteps[semanticStateKey] = semanticVisits;
            PruneSemanticStateHistory(graph);
            graph.RecentSemanticPath.Add(semanticStateKey);
            if (graph.RecentSemanticPath.Count > 32)
                graph.RecentSemanticPath.RemoveAt(0);
            var semanticConsistentPeriod = semanticVisits.Count >= 3 &&
                                           Math.Abs(
                                               (semanticVisits[^1] - semanticVisits[^2]) -
                                               (semanticVisits[^2] - semanticVisits[^3])) <= 1;
            var semanticCycle = departedFromSemanticState &&
                                semanticCycleLength is >= 2 and <= 24 &&
                                repeatedActionCycle &&
                                (semanticVisits.Count >= 3 || semanticConsistentPeriod);
            var unchangedLastStep = !double.IsNaN(lastDelta) && lastDelta < NoChangeThreshold;
            var stateSimilarity = isExistingNode
                ? Math.Clamp(
                    1.0 -
                    (0.75 * matchingNode.ActiveDelta + 0.25 * matchingNode.GlobalDelta) /
                    Math.Max(0.000001, NoChangeThreshold),
                    0,
                    1)
                : 0;
            var interactionDomain = ClassifyInteractionDomain(recentActions, previousAction, context);
            var loopTopology = semanticCycle && !returnedToPriorState
                ? "semantic-cycle"
                : graphCycle
                    ? "graph-cycle"
                    : "state-return";
            var calibrationKey = BuildCalibrationKey(
                MultiStepStateCycle,
                process,
                interactionDomain,
                loopTopology,
                goalMode);
            var decisionThreshold = CalibratedLoopThreshold(
                calibrationKey,
                MultiStepStateCycle);

            if (!returnedToPriorState && !semanticCycle)
            {
                var noReturnAssessment = new LoopDetectionAssessment(
                    false, 0, 0, priorVisitCount, repeatedActionCycle, false,
                    "no context-matched visual or semantic state return with an intervening different state")
                {
                    LoopTopology = "",
                    InteractionDomain = interactionDomain,
                    DecisionThreshold = decisionThreshold,
                    GraphCycle = false,
                    CycleDisposition = "none",
                    CalibrationKey = calibrationKey,
                    RunId = graph.RunId
                };
                if (graph.PendingCandidateStep is int candidateStep &&
                    currentStep - candidateStep >= 3 &&
                    !double.IsNaN(lastDelta) &&
                    lastDelta >= NoChangeThreshold)
                {
                    var outcome = graph.PendingCandidateWasActionable
                        ? "graph_candidate_intervened"
                        : "graph_candidate_dissipated";
                    if (recordLearning)
                    {
                        if (graph.PendingCandidateWasActionable)
                        {
                            RegisterCalibrationInconclusive(
                                graph.PendingCandidateCalibrationKey);
                        }
                        else
                        {
                            RegisterCalibrationOutcome(
                                graph.PendingCandidateCalibrationKey,
                                confirmed: false);
                        }
                    }
                    if (recordLearning)
                    {
                        AppendLoopTelemetry(
                            outcome,
                            currentStep,
                            MultiStepStateCycle,
                            graph.PendingCandidateTopology,
                            graph.PendingCandidateDomain,
                            noReturnAssessment,
                            graph.PendingCandidateWasActionable ? null : false,
                            new
                            {
                                labelSource = graph.PendingCandidateWasActionable
                                    ? "inconclusive_after_intervention"
                                    : "shadow_candidate_dissipated"
                            });
                    }
                    ClearPendingGraphCandidate(graph);
                    if (recordLearning)
                        PersistCalibrationSnapshot();
                }
                return noReturnAssessment;
            }

            // A single return may change recovery behavior, but it is not a
            // calibration label. A later independent recurrence supplies that label.
            var confidence = returnedToPriorState
                ? 0.32 + 0.18 * stateSimilarity
                : 0.38;
            if (graphCycle)
                confidence += 0.20;
            if (priorVisitCount >= 2)
                confidence += 0.12;
            if (consistentReturnPeriod)
                confidence += 0.15;
            if (repeatedActionCycle)
                confidence += 0.18;
            if (unchangedLastStep)
                confidence += 0.06;
            if (semanticCycle)
                confidence += semanticConsistentPeriod ? 0.18 : 0.12;
            confidence = Math.Clamp(confidence, 0, 1);
            var recurrenceDetected = confidence >= decisionThreshold;
            var productiveCycle = recurrenceDetected &&
                                  IsLikelyProductiveCycle(
                                      goalMode,
                                      recurringWorkflowIntent,
                                      actions,
                                      previousAction,
                                      unchangedLastStep);

            var evidence = $"graph_return={stateSimilarity:0.00}; node_visits={priorVisitCount + 1}; graph_cycle={graphCycle}; " +
                           $"consistent_period={consistentReturnPeriod}; action_cycle={repeatedActionCycle}; " +
                           $"semantic_cycle={semanticCycle}; semantic_period={semanticConsistentPeriod}; " +
                           $"unchanged_last_step={unchangedLastStep}; productive_cycle={productiveCycle}";
            var independentlyConfirmed =
                confidence >= 0.5 &&
                !productiveCycle &&
                graph.PendingCandidateWasActionable &&
                graph.PendingCandidateStep is int pendingStep &&
                currentStep - pendingStep >= 2 &&
                string.Equals(
                    graph.PendingCandidateCalibrationKey,
                    calibrationKey,
                    StringComparison.OrdinalIgnoreCase) &&
                (graph.PendingCandidateNodeId == node.Id ||
                 (semanticCycle &&
                  string.Equals(
                      graph.PendingCandidateSemanticStateKey,
                      semanticStateKey,
                      StringComparison.Ordinal)));
            var assessment = new LoopDetectionAssessment(
                recurrenceDetected && !productiveCycle,
                confidence,
                returnedToPriorState ? cycleLength : semanticCycleLength,
                returnedToPriorState ? priorVisitCount : semanticVisits.Count - 1,
                repeatedActionCycle,
                consistentReturnPeriod || semanticConsistentPeriod,
                evidence)
            {
                LoopTopology = loopTopology,
                InteractionDomain = interactionDomain,
                DecisionThreshold = decisionThreshold,
                GraphCycle = graphCycle,
                SemanticCycle = semanticCycle,
                IsProductiveCycle = productiveCycle,
                CycleDisposition = productiveCycle
                    ? "productive"
                    : recurrenceDetected
                        ? "harmful"
                        : "candidate",
                IndependentlyConfirmed = independentlyConfirmed,
                CalibrationKey = calibrationKey,
                RunId = graph.RunId
            };
            if (productiveCycle)
            {
                if (recordLearning)
                {
                    if (graph.PendingCandidateStep is null ||
                        !string.Equals(
                            graph.PendingCandidateCalibrationKey,
                            calibrationKey,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        if (graph.PendingCandidateStep is not null)
                        {
                            RegisterCalibrationInconclusive(
                                graph.PendingCandidateCalibrationKey);
                        }
                        RegisterCalibrationCandidate(calibrationKey);
                    }
                    RegisterCalibrationInconclusive(calibrationKey);
                    AppendLoopTelemetry(
                        "graph_productive_cycle",
                        currentStep,
                        MultiStepStateCycle,
                        loopTopology,
                        interactionDomain,
                        assessment,
                        null,
                        new
                        {
                            labelSource =
                                "goal_aligned_recurring_workflow",
                            goalMode
                        });
                    PersistCalibrationSnapshot();
                }
                ClearPendingGraphCandidate(graph);
            }
            else if (confidence >= 0.5)
            {
                if (independentlyConfirmed)
                {
                    if (recordLearning)
                        RegisterCalibrationOutcome(calibrationKey, confirmed: true);
                    ClearPendingGraphCandidate(graph);
                    if (recordLearning)
                        PersistCalibrationSnapshot();
                }
                else if (graph.PendingCandidateStep is null)
                {
                    if (recordLearning)
                        RegisterCalibrationCandidate(calibrationKey);
                    SetPendingGraphCandidate(
                        graph,
                        currentStep,
                        confidence,
                        loopTopology,
                        interactionDomain,
                        calibrationKey,
                        node.Id,
                        semanticStateKey,
                        node.VisitCount,
                        assessment.IsLoop);
                }
                else if (string.Equals(
                             graph.PendingCandidateCalibrationKey,
                             calibrationKey,
                             StringComparison.OrdinalIgnoreCase))
                {
                    graph.PendingCandidateConfidence = Math.Max(
                        graph.PendingCandidateConfidence,
                        confidence);
                    if (assessment.IsLoop && !graph.PendingCandidateWasActionable)
                    {
                        // Promote a shadow observation to the first actionable
                        // occurrence. Calibration still waits for a later recurrence
                        // of this same phase instead of labeling the threshold
                        // crossing with its own evidence.
                        graph.PendingCandidateStep = currentStep;
                        graph.PendingCandidateTopology = loopTopology;
                        graph.PendingCandidateDomain = interactionDomain;
                        graph.PendingCandidateNodeId = node.Id;
                        graph.PendingCandidateSemanticStateKey = semanticStateKey;
                        graph.PendingCandidateVisitCount = node.VisitCount;
                        graph.PendingCandidateWasActionable = true;
                    }
                }
                else
                {
                    // A candidate from another context must not block learning
                    // in the current context. Treat the displaced candidate as
                    // censored rather than as evidence of a false loop.
                    if (recordLearning)
                    {
                        RegisterCalibrationInconclusive(
                            graph.PendingCandidateCalibrationKey);
                        RegisterCalibrationCandidate(calibrationKey);
                    }
                    SetPendingGraphCandidate(
                        graph,
                        currentStep,
                        confidence,
                        loopTopology,
                        interactionDomain,
                        calibrationKey,
                        node.Id,
                        semanticStateKey,
                        node.VisitCount,
                        assessment.IsLoop);
                }
                if (recordLearning)
                {
                    AppendLoopTelemetry(
                        independentlyConfirmed
                            ? "graph_loop_independently_confirmed"
                            : assessment.IsLoop
                                ? "graph_loop_detected"
                                : "graph_loop_shadow_candidate",
                        currentStep,
                        MultiStepStateCycle,
                        loopTopology,
                        interactionDomain,
                        assessment,
                        independentlyConfirmed ? true : null,
                        new
                        {
                            labelSource = independentlyConfirmed
                                ? "second_context_matched_recurrence"
                                : assessment.IsLoop
                                    ? "intervention_candidate_not_labeled"
                                    : "shadow_candidate"
                        });
                }
            }
            return assessment;
        }

        static void ClearPendingGraphCandidate(LoopStateGraph graph)
        {
            graph.PendingCandidateStep = null;
            graph.PendingCandidateConfidence = 0;
            graph.PendingCandidateTopology = "";
            graph.PendingCandidateDomain = "";
            graph.PendingCandidateCalibrationKey = "";
            graph.PendingCandidateNodeId = null;
            graph.PendingCandidateSemanticStateKey = "";
            graph.PendingCandidateVisitCount = 0;
            graph.PendingCandidateWasActionable = false;
        }

        static void SetPendingGraphCandidate(
            LoopStateGraph graph,
            int step,
            double confidence,
            string topology,
            string domain,
            string calibrationKey,
            int nodeId,
            string semanticStateKey,
            int visitCount,
            bool actionable)
        {
            graph.PendingCandidateStep = step;
            graph.PendingCandidateConfidence = confidence;
            graph.PendingCandidateTopology = topology;
            graph.PendingCandidateDomain = domain;
            graph.PendingCandidateCalibrationKey = calibrationKey;
            graph.PendingCandidateNodeId = nodeId;
            graph.PendingCandidateSemanticStateKey = semanticStateKey;
            graph.PendingCandidateVisitCount = visitCount;
            graph.PendingCandidateWasActionable = actionable;
        }

        static bool ContextTokensCompatible(
            string left,
            string right,
            double minimum = 0.5) =>
            (string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right)) ||
            TokenSimilarity(left, right) >= minimum;

        static void AddOrUpdateGraphEdge(
            LoopStateGraph graph,
            int fromNodeId,
            int toNodeId,
            ResolvedActionSnapshot? previousAction,
            int step)
        {
            var actionKey = previousAction is null ? "unknown" : ActionCycleKey(previousAction);
            var edge = graph.Edges.FirstOrDefault(item =>
                item.FromNodeId == fromNodeId &&
                item.ToNodeId == toNodeId &&
                string.Equals(item.ActionKey, actionKey, StringComparison.Ordinal));
            if (edge is null)
            {
                graph.Edges.Add(new LoopStateEdge
                {
                    FromNodeId = fromNodeId,
                    ToNodeId = toNodeId,
                    ActionKey = actionKey,
                    TraversalCount = 1,
                    LastSeenStep = step
                });
            }
            else
            {
                edge.TraversalCount++;
                edge.LastSeenStep = step;
            }
            PruneGraphEdges(graph);
        }

        static bool GraphHasPath(
            LoopStateGraph graph,
            int fromNodeId,
            int targetNodeId,
            int maxDepth)
        {
            var queue = new Queue<(int NodeId, int Depth)>();
            var visited = new HashSet<int> { fromNodeId };
            queue.Enqueue((fromNodeId, 0));
            while (queue.Count > 0)
            {
                var (nodeId, depth) = queue.Dequeue();
                if (depth >= maxDepth)
                    continue;
                foreach (var edge in graph.Edges.Where(edge => edge.FromNodeId == nodeId))
                {
                    if (edge.ToNodeId == targetNodeId)
                        return true;
                    if (visited.Add(edge.ToNodeId))
                        queue.Enqueue((edge.ToNodeId, depth + 1));
                }
            }
            return false;
        }

        static void PruneStateGraph(LoopStateGraph graph)
        {
            const int maximumNodes = 64;
            if (graph.Nodes.Count <= maximumNodes)
                return;

            var recent = graph.RecentNodePath.ToHashSet();
            var remove = graph.Nodes
                .Where(node => !recent.Contains(node.Id))
                .OrderBy(node => node.LastSeenStep)
                .Take(graph.Nodes.Count - maximumNodes)
                .Select(node => node.Id)
                .ToHashSet();
            graph.Nodes.RemoveAll(node => remove.Contains(node.Id));
            graph.Edges.RemoveAll(edge =>
                remove.Contains(edge.FromNodeId) || remove.Contains(edge.ToNodeId));
            PruneGraphEdges(graph);
        }

        static void PruneGraphEdges(LoopStateGraph graph)
        {
            var limit = Math.Max(32, RuntimeGraphEdgeLimit);
            if (graph.Edges.Count <= limit)
                return;

            var keep = graph.Edges
                .OrderByDescending(edge => edge.LastSeenStep)
                .ThenByDescending(edge => edge.TraversalCount)
                .Take(limit)
                .ToHashSet();
            graph.Edges.RemoveAll(edge => !keep.Contains(edge));
        }

        static void PruneSemanticStateHistory(LoopStateGraph graph)
        {
            var limit = Math.Max(32, RuntimeSemanticStateLimit);
            if (graph.SemanticVisitSteps.Count <= limit)
                return;

            var remove = graph.SemanticVisitSteps
                .OrderBy(item =>
                    item.Value.Count == 0 ? int.MinValue : item.Value[^1])
                .Take(graph.SemanticVisitSteps.Count - limit)
                .Select(item => item.Key)
                .ToArray();
            foreach (var key in remove)
                graph.SemanticVisitSteps.Remove(key);
        }

        internal static RecentLoopState CreateRecentLoopState(
            int step,
            byte[] screenFingerprint,
            UiPromptContext context) =>
            new(
                step,
                screenFingerprint.ToArray(),
                NormalizeText(context.ActiveProcessName),
                NormalizeTokens(context.ActiveWindowTitle))
            {
                FocusTokens = NormalizeTokens(context.FocusedUiaSummary)
            };

        static string ActionCycleKey(ResolvedActionSnapshot action)
        {
            var family = ActionFamily(action.Action);
            if (family is "click" or "move" && action.ScreenPoint is Point point)
            {
                const int regionSize = 48;
                return $"{family}:{(int)Math.Round(point.X / (double)regionSize)},{(int)Math.Round(point.Y / (double)regionSize)}";
            }
            if (family == "drag_drop" &&
                action.ScreenPoint is Point source &&
                action.DestinationScreenPoint is Point destination)
            {
                return $"drag_drop:{source.X},{source.Y}->{destination.X},{destination.Y}";
            }

            return string.IsNullOrWhiteSpace(action.IneffectiveSignature)
                ? family
                : action.IneffectiveSignature.ToLowerInvariant();
        }

        static double ComputeStabilityWeightedDelta(
            byte[] reference,
            byte[] current,
            double[] instability)
        {
            if (reference.Length != current.Length || reference.Length == 0)
                return 1;
            if (instability.Length != reference.Length)
                return ComputeImageDelta(reference, current);

            double weightedDifference = 0;
            double totalWeight = 0;
            for (var index = 0; index < reference.Length; index++)
            {
                // Frequently animated pixels retain at least 25% weight, so a
                // learned mask cannot hide a genuinely changing application.
                var weight = Math.Clamp(1.0 - instability[index] * 4.0, 0.25, 1.0);
                weightedDifference += Math.Abs(reference[index] - current[index]) / 255.0 * weight;
                totalWeight += weight;
            }
            return totalWeight <= 0 ? 1 : weightedDifference / totalWeight;
        }

        static double[] UpdateInstabilityMask(
            byte[] reference,
            byte[] current,
            double[] existing)
        {
            if (reference.Length != current.Length || reference.Length == 0)
                return [];
            var result = existing.Length == reference.Length
                ? existing.ToArray()
                : new double[reference.Length];
            for (var index = 0; index < reference.Length; index++)
            {
                var observed = Math.Abs(reference[index] - current[index]) / 255.0;
                result[index] = 0.85 * result[index] + 0.15 * observed;
            }
            return result;
        }

        static string BuildSemanticStateKey(
            UiPromptContext context,
            ResolvedActionSnapshot? previousAction)
        {
            var actionFamily = ActionFamily(previousAction?.Action);
            var actionTarget = NormalizeTokens(previousAction?.SemanticTokens);
            return string.Join(
                '|',
                NormalizeText(context.ActiveProcessName),
                NormalizeTokens(context.ActiveWindowTitle),
                NormalizeTokens(context.FocusedUiaSummary),
                actionFamily,
                actionTarget);
        }

        static string BuildCalibrationKey(
            string loopKind,
            string process,
            string interactionDomain,
            string loopTopology,
            string goalMode) =>
            string.Join(
                '|',
                loopKind,
                string.IsNullOrWhiteSpace(interactionDomain) ? "unknown-domain" : interactionDomain,
                string.IsNullOrWhiteSpace(loopTopology) ? "unknown-topology" : loopTopology,
                string.IsNullOrWhiteSpace(process) ? "unknown-process" : process,
                string.Equals(
                    goalMode,
                    "continuous",
                    StringComparison.OrdinalIgnoreCase)
                    ? "continuous"
                    : "finite");

        internal static bool HasRecurringWorkflowIntent(string? goal)
        {
            var normalized = NormalizeText(goal);
            return Regex.IsMatch(
                normalized,
                @"\b(monitor|monitoring|watch|observe|poll|supervise|maintain|repeatedly|periodically|regularly|cyclically|recurring|keep\s+(?:it\s+)?(?:active|running|healthy)|respond\s+to\s+new|handle\s+incoming|check\s+(?:repeatedly|periodically)|monitoruj|obserwuj|pilnuj|nadzoruj|utrzymuj|sprawdzaj|powtarzaj|regularnie|okresowo|cyklicznie|śledź|reaguj\s+na|obsługuj\s+nowe)\b",
                RegexOptions.CultureInvariant);
        }

        static bool IsLikelyProductiveCycle(
            string goalMode,
            bool recurringWorkflowIntent,
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot? previousAction,
            bool unchangedLastStep)
        {
            if (!string.Equals(
                    goalMode,
                    "continuous",
                    StringComparison.OrdinalIgnoreCase) ||
                !recurringWorkflowIntent ||
                unchangedLastStep)
            {
                return false;
            }

            var actions = recentActions.TakeLast(12).ToList();
            if (previousAction is not null &&
                (actions.Count == 0 ||
                 !ReferenceEquals(actions[^1], previousAction)))
            {
                actions.Add(previousAction);
            }
            if (actions.Count == 0)
                return false;

            var families = actions
                .Select(action => ActionFamily(action.Action))
                .ToArray();
            if (families.Any(family =>
                    family is "drag_drop" or "text_input" or "run_command" or
                        "open_url" or "launch_app"))
            {
                return false;
            }

            var hasObservationCadence = families.Any(family =>
                family is "wait" or "observe");
            var semanticEvidence = string.Join(
                ' ',
                actions.Select(action =>
                    $"{action.SemanticTokens} {action.Action.Note}"));
            var hasRecurringSemanticEvidence = Regex.IsMatch(
                NormalizeText(semanticEvidence),
                @"\b(monitor|watch|observe|poll|check|refresh|status|health|incoming|new\s+event|monitoruj|obserwuj|sprawdź|sprawdzaj|odśwież|status|nowe\s+zdarzenie)\b",
                RegexOptions.CultureInvariant);
            return hasObservationCadence || hasRecurringSemanticEvidence;
        }

        static bool IsTerminalProcess(string process) =>
            process is "pwsh" or "powershell" or "powershell_ise" or "cmd" or "wt" or "windowsterminal" or "conhost";

        static bool IsBrowserProcess(string process) =>
            process is "msedge" or "chrome" or "firefox" or "brave" or "opera";

        static bool LooksLikeRasterSurface(
            string process,
            string visualContext,
            IReadOnlyCollection<ResolvedActionSnapshot> actions)
        {
            if (Regex.IsMatch(visualContext, @"\b(canvas|raster|bitmap|pixel|sprite)\b", RegexOptions.CultureInvariant))
                return true;

            if (!IsBrowserProcess(process))
                return false;

            var isDocumentSurface =
                visualContext.Contains("rootwebarea", StringComparison.Ordinal) ||
                visualContext.Contains("controltype.document", StringComparison.Ordinal) ||
                visualContext.Contains(" document", StringComparison.Ordinal);
            if (!isDocumentSurface)
                return false;

            var actionIntent = string.Join(
                ' ',
                actions.TakeLast(6).Select(a => $"{a.Action.Note} {a.Description}"));
            return Regex.IsMatch(
                $"{visualContext} {actionIntent}".ToLowerInvariant(),
                @"\b(game|puzzle|board|map|image|diagram|drawing|paint|piece|tile)\b",
                RegexOptions.CultureInvariant);
        }

        static bool HasPlacementIntent(ResolvedActionSnapshot snapshot)
        {
            var intent = $"{snapshot.Action.Note} {snapshot.Description}".ToLowerInvariant();
            return Regex.IsMatch(
                intent,
                @"\b(drag|drop|place|placement|position|reposition|move\s+(?:the\s+)?(?:piece|item|object|token|tile|card))\b",
                RegexOptions.CultureInvariant);
        }

        static string NormalizeTokens(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";

            return string.Join(' ', Regex.Matches(value.ToLowerInvariant(), @"[\p{L}\p{Nd}]{2,}").Cast<Match>()
                .Select(match => match.Value)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(token => token, StringComparer.Ordinal)
                .Take(24));
        }

        static double TokenSimilarity(string left, string right)
        {
            var a = left.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet(StringComparer.Ordinal);
            var b = right.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToHashSet(StringComparer.Ordinal);
            if (a.Count == 0 || b.Count == 0)
                return 0;

            var intersection = a.Count(b.Contains);
            var union = a.Count + b.Count - intersection;
            return union == 0 ? 0 : intersection / (double)union;
        }

        static bool TryDecodeFingerprint(string value, out byte[] fingerprint)
        {
            try
            {
                fingerprint = Convert.FromBase64String(value);
                return true;
            }
            catch
            {
                fingerprint = [];
                return false;
            }
        }
    }
}
