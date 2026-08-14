internal static partial class RDPilotApplication
{
    /// <summary>
    /// Learns coordinate-independent recovery strategies from successfully escaped UI stalls.
    /// </summary>
    internal static partial class RecoveryMemoryService
    {
        const double MinimumLessonSimilarity = 0.55;
        const int RecentActionLimit = 24;
        const int CurrentMemoryVersion = 7;
        internal static readonly object RecoveryFileGate = new();
        static readonly string CalibrationWriterId =
            $"{DateTime.UtcNow.Ticks:x16}-{Environment.ProcessId}-{Guid.NewGuid():N}";
        static Dictionary<string, LoopCalibrationBucket> RecoveryCalibration =
            new(StringComparer.OrdinalIgnoreCase);
        static bool RecoveryMemoryReadOnly;
        static bool LastRecoveryLoadUsedBackup;
        static bool RecoveryMemoryDirty;
        static List<RecoveryLesson>? PendingRecoveryLessons;

        internal const string AlternatingActionCycle = "AlternatingActionCycle";
        internal const string MultiStepActionCycle = "MultiStepActionCycle";
        internal const string MultiStepStateCycle = "MultiStepStateCycle";
        internal const string ConsoleInteractionLoop = "ConsoleInteractionLoop";
        internal const string CommandFailureLoop = "CommandFailureLoop";
        internal const string TextInputNoChange = "TextInputNoChange";
        internal const string RepeatedKeySequence = "RepeatedKeySequence";
        internal const string PlacementOrDragLoop = "PlacementOrDragLoop";
        internal const string RasterCanvasPointerLoop = "RasterCanvasPointerLoop";
        internal const string PointerRegionLoop = "PointerRegionLoop";
        internal const string UiaInteractionLoop = "UiaInteractionLoop";
        internal const string NavigationBounce = "NavigationBounce";
        internal const string ScrollNoProgress = "ScrollNoProgress";
        internal const string ObservationLoop = "ObservationLoop";
        internal const string WaitNoChange = "WaitNoChange";
        internal const string GenericNoProgressLoop = "GenericNoProgressLoop";
        internal const string RejectedProposalLoop = "RejectedProposalLoop";

        internal static async Task<RecoveryEpisodeState?> UpdateRecoveryEpisodeAsync(
            RecoveryEpisodeState? episode,
            int step,
            int stagnationSteps,
            int repeatCount,
            double lastDelta,
            ObservationAssessment? observationAssessment,
            byte[] shotFingerprint,
            byte[] activeWindowFingerprint,
            ResolvedActionSnapshot? previousAction,
            UiPromptContext context,
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            List<RecoveryLesson> lessons,
            LoopDetectionAssessment loopAssessment,
            bool planningLoopDetected,
            int planningCycleLength,
            ResolvedActionSnapshot? previousRejectedAction,
            IReadOnlyCollection<ResolvedActionSnapshot> recentRejectedActions,
            string goal,
            string goalMode,
            string currentImageDataUrl,
            string? currentImagePath,
            Func<RecoveryEpisodeState, Task<RecoveryProgressDto?>> verifyProgressAsync)
        {
            if (!RecoveryMemoryEnabled ||
                (double.IsNaN(lastDelta) && !planningLoopDetected))
                return episode;

            var noProgress = observationAssessment is not null
                ? observationAssessment.GoalProgress == GoalProgressState.NoProgress
                : lastDelta < NoChangeThreshold;
            var proactivePattern = noProgress &&
                                   HasEmergingLoopPattern(recentActions, previousAction);
            var proactiveVisualCycle = loopAssessment.IsLoop;
            var triggerReached = stagnationSteps >= Math.Max(1, RecoveryMemoryTriggerSteps) ||
                                 repeatCount >= 1 ||
                                 proactivePattern ||
                                 proactiveVisualCycle ||
                                 planningLoopDetected;
            if (episode is null)
            {
                if (!triggerReached)
                    return null;

                var loopKind = planningLoopDetected
                    ? RejectedProposalLoop
                    : proactiveVisualCycle
                        ? MultiStepStateCycle
                        : ClassifyLoopKind(recentActions, previousAction, context);
                var loopTopology = planningLoopDetected
                    ? planningCycleLength > 1
                        ? "rejected-proposal-cycle"
                        : "rejected-proposal-repeat"
                    : proactiveVisualCycle
                        ? loopAssessment.LoopTopology
                        : ClassifyLoopTopology(recentActions, previousAction, stagnationSteps);
                var triggerAction = planningLoopDetected
                    ? previousRejectedAction
                    : previousAction;
                var failedActions = planningLoopDetected
                    ? recentRejectedActions
                    : recentActions;
                var interactionDomain =
                    planningLoopDetected ||
                    string.IsNullOrWhiteSpace(
                        loopAssessment.InteractionDomain)
                        ? ClassifyInteractionDomain(
                            failedActions,
                            triggerAction,
                            context)
                        : loopAssessment.InteractionDomain;
                var started = new RecoveryEpisodeState
                {
                    StartedAtStep = step,
                    TriggerContext = context,
                    TriggerFingerprint = shotFingerprint.ToArray(),
                    TriggerActiveWindowFingerprint = activeWindowFingerprint.ToArray(),
                    GoalTokens = NormalizeGoalTokens(goal),
                    GoalDomain = ClassifyGoalDomain(goal),
                    GoalIntentTokens = NormalizeGoalIntentTokens(goal),
                    GoalMode = goalMode,
                    TriggerActionFamily = ActionFamily(triggerAction?.Action),
                    LoopKind = loopKind,
                    LoopTopology = loopTopology,
                    InteractionDomain = interactionDomain,
                    FailedActions = failedActions.TakeLast(RecentActionLimit).ToList(),
                    MaxStagnationSteps = stagnationSteps,
                    TriggerImageDataUrl = currentImageDataUrl,
                    TriggerImagePath = currentImagePath
                };
                AppendLoopTelemetry("episode_started", step, loopKind, loopTopology, interactionDomain, loopAssessment, null);
                if (!proactiveVisualCycle &&
                    !planningLoopDetected &&
                    loopAssessment.Confidence < 0.5 &&
                    (stagnationSteps >= Math.Max(1, RecoveryMemoryTriggerSteps) || repeatCount >= 1))
                {
                    AppendLoopTelemetry(
                        "loop_detected_by_action_guard_without_visual_candidate",
                        step,
                        loopKind,
                        loopTopology,
                        interactionDomain,
                        loopAssessment,
                        true);
                }
                Console.WriteLine($"[memory] recovery episode started at step {step}; kind={started.LoopKind}; topology={loopTopology}; domain={interactionDomain}; trigger={started.TriggerActionFamily}; stagnation={stagnationSteps}; repeat={repeatCount}; proactive={proactivePattern || proactiveVisualCycle || planningLoopDetected}");
                return started;
            }

            episode.MaxStagnationSteps = Math.Max(episode.MaxStagnationSteps, stagnationSteps);
            if (episode.IsValidating)
            {
                if (triggerReached)
                {
                    Console.WriteLine("[memory] candidate recovery rejected because the loop returned during validation.");
                    RegisterAppliedStrategyFailure(episode, lessons, "loop returned during validation");
                    SaveRecoveryLessons(lessons);
                    AppendLoopTelemetry("validation_rejected", step, episode.LoopKind, episode.LoopTopology, episode.InteractionDomain, loopAssessment, false);
                    episode.IsValidating = false;
                    episode.ValidationRemaining = 0;
                    episode.CandidateLesson = null;
                    return episode;
                }

                var meaningfulValidationTransition = IsMeaningfulStateTransition(
                    episode.TriggerContext,
                    episode.TriggerFingerprint,
                    episode.TriggerActiveWindowFingerprint,
                    context,
                    shotFingerprint,
                    activeWindowFingerprint,
                    lastDelta);
                episode.ValidationObservedSemanticProgress |=
                    IsRecoveryValidationProgress(
                        previousAction,
                        observationAssessment,
                        meaningfulValidationTransition);
                episode.ValidationRemaining--;
                if (episode.ValidationRemaining > 0)
                    return episode;

                if (episode.CandidateLesson is RecoveryLesson confirmed &&
                    episode.ValidationObservedSemanticProgress)
                {
                    StoreConfirmedLesson(confirmed, lessons);
                    AppendLoopTelemetry("recovery_confirmed", step, episode.LoopKind, episode.LoopTopology, episode.InteractionDomain, loopAssessment, true);
                }
                else
                {
                    SaveRecoveryLessons(lessons);
                    AppendLoopTelemetry("recovery_rejected_no_semantic_progress", step, episode.LoopKind, episode.LoopTopology, episode.InteractionDomain, loopAssessment, false);
                }
                return null;
            }

            var visibleProgress = stagnationSteps == 0 &&
                                  previousAction != null &&
                                  !IsLocalObservationAction(previousAction.Action) &&
                                  (observationAssessment?.IsProgress == true ||
                                   observationAssessment is null &&
                                   lastDelta >= NoChangeThreshold &&
                                   IsMeaningfulStateTransition(
                                       episode.TriggerContext,
                                       episode.TriggerFingerprint,
                                       episode.TriggerActiveWindowFingerprint,
                                       context,
                                       shotFingerprint,
                                       activeWindowFingerprint,
                                       lastDelta));
            if (!visibleProgress)
            {
                if (episode.AppliedLessonId is not null)
                    episode.AppliedLessonNoProgressObservations++;
                if (triggerReached &&
                    (episode.AppliedLessonId is null ||
                     episode.AppliedLessonNoProgressObservations >=
                     Math.Max(2, RecoveryMemoryValidationSteps)))
                {
                    RegisterAppliedStrategyFailure(episode, lessons, "strategy did not produce visible progress");
                }
                return episode;
            }

            RecoveryProgressDto? progressAssessment;
            if (RecoveryProgressVerificationEnabled)
            {
                progressAssessment = await verifyProgressAsync(episode);
                var verifiedProgress = IsVerifiedRecoveryProgress(progressAssessment);
                if (!verifiedProgress)
                {
                    var verdict = progressAssessment?.Verdict ?? "unavailable";
                    var confidence = progressAssessment?.Confidence ?? 0;
                    var evidence = progressAssessment?.Evidence ?? "independent progress verification unavailable";
                    if (episode.AppliedLessonId is not null &&
                        progressAssessment?.Verdict?.Equals("no", StringComparison.OrdinalIgnoreCase) == true &&
                        confidence >= RecoveryProgressConfidenceThreshold)
                    {
                        RegisterAppliedStrategyFailure(
                            episode,
                            lessons,
                            $"goal-progress verifier rejected the outcome: {TrimForMeta(evidence, 180)}");
                    }
                    AppendLoopTelemetry(
                        "recovery_progress_not_confirmed",
                        step,
                        episode.LoopKind,
                        episode.LoopTopology,
                        episode.InteractionDomain,
                        loopAssessment,
                        false,
                        new
                        {
                            goalMode,
                            verdict,
                            confidence,
                            evidence
                        });
                    Console.WriteLine(
                        $"[memory] visible change was not learned as progress; verifier={verdict}/{confidence:0.00}; " +
                        $"evidence={TrimForMeta(evidence, 220)}");
                    return episode;
                }
            }
            else
            {
                progressAssessment = new RecoveryProgressDto
                {
                    Verdict = "yes",
                    Confidence = Math.Clamp(
                        Math.Max(lastDelta, NoChangeThreshold) /
                        Math.Max(NoChangeThreshold * 4, 0.02),
                        0.5,
                        0.85),
                    Evidence = "local semantic and foreground visual transition; model verifier disabled",
                    StateLabel = NormalizeTokens(
                        $"{context.ActiveWindowTitle} {context.FocusedUiaSummary}")
                };
            }

            episode.LastProgressConfidence = progressAssessment!.Confidence;
            episode.LastProgressEvidence = progressAssessment.Evidence ?? "";

            episode.LastOutcomeTokens = NormalizeTokens(
                $"{context.ActiveWindowTitle} {context.FocusedUiaSummary} {progressAssessment.StateLabel}");
            if (episode.AppliedLessonId is not null &&
                !AppliedLessonOutcomeMatches(
                    episode,
                    lessons,
                    episode.LastOutcomeTokens,
                    shotFingerprint,
                    activeWindowFingerprint))
            {
                RegisterAppliedStrategyFailure(
                    episode,
                    lessons,
                    "screen changed but the learned semantic outcome was not observed");
            }
            episode.AppliedLessonMadeProgress = episode.AppliedLessonId != null;
            episode.CandidateLesson = BuildCandidateLesson(
                episode,
                previousAction!,
                lastDelta,
                episode.LastOutcomeTokens,
                shotFingerprint,
                activeWindowFingerprint);
            episode.IsValidating = true;
            episode.ValidationRemaining = Math.Max(1, RecoveryMemoryValidationSteps);
            episode.ValidationObservedSemanticProgress = false;
            Console.WriteLine($"[memory] candidate recovery found; validating for {episode.ValidationRemaining} additional step(s).");
            return episode;
        }

        internal static bool IsRecoveryValidationProgress(
            ResolvedActionSnapshot? previousAction,
            ObservationAssessment? observationAssessment,
            bool meaningfulStateTransition) =>
            previousAction is not null &&
            !IsLocalObservationAction(previousAction.Action) &&
            (observationAssessment?.IsProgress == true ||
             observationAssessment is null && meaningfulStateTransition);

        internal static bool IsVerifiedRecoveryProgress(
            RecoveryProgressDto? assessment) =>
            assessment?.Verdict?.Equals(
                "yes",
                StringComparison.OrdinalIgnoreCase) == true &&
            assessment.Confidence >= RecoveryProgressConfidenceThreshold;

        internal static void RecordRecoveryAction(
            RecoveryEpisodeState? episode,
            Queue<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot action,
            IReadOnlyCollection<RecoveryLesson> lessons)
        {
            if (action.Action.Type == "done")
                return;

            recentActions.Enqueue(action);
            while (recentActions.Count > RecentActionLimit)
                recentActions.Dequeue();

            if (episode is { IsValidating: false })
            {
                episode.RecoveryActions.Add(action);
                while (episode.RecoveryActions.Count >
                       Math.Max(8, RuntimeRecoveryActionLimit))
                {
                    episode.RecoveryActions.RemoveAt(0);
                }
                TrackSuggestedStrategyAttempt(episode, action, lessons);
            }
        }

        internal static RecoveryEpisodeState? ConfirmPendingRecovery(
            RecoveryEpisodeState? episode,
            List<RecoveryLesson> lessons,
            bool independentlyVerified)
        {
            if (episode?.CandidateLesson is RecoveryLesson candidate && independentlyVerified)
            {
                StoreConfirmedLesson(candidate, lessons);
            }
            else if (episode?.CandidateLesson is not null)
            {
                Console.WriteLine("[memory] pending recovery was not persisted because goal verification was skipped.");
            }
            return null;
        }

        internal static string BuildRecoveryMemoryPrompt(
            IReadOnlyCollection<RecoveryLesson> lessons,
            UiPromptContext context,
            byte[] shotFingerprint,
            byte[] activeWindowFingerprint,
            ResolvedActionSnapshot? previousAction,
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            RecoveryEpisodeState? episode,
            int stagnationSteps,
            int repeatCount,
            string goal)
        {
            if (!RecoveryMemoryEnabled ||
                RecoveryMemoryPromptMaxLessons <= 0 ||
                (episode is null &&
                 stagnationSteps < Math.Max(1, RecoveryMemoryTriggerSteps) &&
                 repeatCount < 1))
            {
                return "";
            }

            var actionFamily = ActionFamily(previousAction?.Action);
            var loopKind = episode?.LoopKind ?? ClassifyLoopKind(recentActions, previousAction, context);
            var loopTopology = episode?.LoopTopology ??
                               ClassifyLoopTopology(recentActions, previousAction, stagnationSteps);
            var interactionDomain = episode?.InteractionDomain ??
                                    ClassifyInteractionDomain(recentActions, previousAction, context);
            var goalTokens = episode?.GoalTokens ?? NormalizeGoalTokens(goal);
            var goalDomain = episode?.GoalDomain ?? ClassifyGoalDomain(goal);
            var goalIntentTokens = episode?.GoalIntentTokens ?? NormalizeGoalIntentTokens(goal);
            var goalMode = episode?.GoalMode ?? ClassifyGoalMode(goal);
            var ranked = lessons
                .Where(IsLessonActive)
                .Where(lesson => episode is null || !episode.RejectedLessonIds.Contains(lesson.Id))
                .Select(lesson =>
                {
                    var similarity = LessonSimilarity(
                        lesson,
                        context,
                        shotFingerprint,
                        activeWindowFingerprint,
                        actionFamily,
                        loopKind,
                        loopTopology,
                        interactionDomain,
                        goalTokens,
                        goalDomain,
                        goalIntentTokens,
                        goalMode);
                    return (
                        Lesson: lesson,
                        Similarity: similarity,
                        SelectionScore: ContextualBanditScore(lesson, similarity));
                })
                .Where(item => item.Similarity >= MinimumLessonSimilarity)
                .OrderByDescending(item => item.SelectionScore)
                .ThenByDescending(item => item.Similarity)
                .ThenByDescending(item => item.Lesson.SuccessCount)
                .ThenByDescending(item => item.Lesson.UpdatedUtc)
                .GroupBy(
                    item => string.IsNullOrWhiteSpace(item.Lesson.StrategySignature)
                        ? item.Lesson.Id
                        : item.Lesson.StrategySignature,
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .Take(RecoveryMemoryPromptMaxLessons)
                .ToArray();

            var negative = lessons
                .Where(lesson => !IsLessonActive(lesson))
                .Select(lesson => (
                    Lesson: lesson,
                    Similarity: LessonSimilarity(
                        lesson,
                        context,
                        shotFingerprint,
                        activeWindowFingerprint,
                        actionFamily,
                        loopKind,
                        loopTopology,
                        interactionDomain,
                        goalTokens,
                        goalDomain,
                        goalIntentTokens,
                        goalMode)))
                .Where(item => item.Similarity >= MinimumLessonSimilarity)
                .OrderByDescending(item => item.Similarity)
                .Take(2)
                .ToArray();

            if (ranked.Length == 0 && negative.Length == 0)
                return "";

            if (episode != null)
            {
                foreach (var item in ranked)
                {
                    if (!episode.SuggestedLessonIds.Contains(item.Lesson.Id, StringComparer.Ordinal))
                        episode.SuggestedLessonIds.Add(item.Lesson.Id);
                }
            }

            var sb = new StringBuilder()
                .AppendLine("RECOVERY_MEMORY:")
                .AppendLine($"DETECTED_LOOP_KIND: {loopKind}")
                .AppendLine($"LOOP_TOPOLOGY: {loopTopology}; INTERACTION_DOMAIN: {interactionDomain}; GOAL_DOMAIN: {goalDomain}; GOAL_MODE: {goalMode}")
                .AppendLine("Previously confirmed, materially distinct recoveries for similar loop states are listed below. Try the safest highest-ranked strategy whose visible preconditions match. If its preconditions do not match or its expected effect fails, try the next distinct route instead of a coordinate variation. Adapt targets to the current UI and never reuse old coordinates blindly.");
            for (var i = 0; i < ranked.Length; i++)
            {
                var item = ranked[i];
                sb.AppendLine($"{i + 1}. strategy_id={item.Lesson.Id}; loop_kind={item.Lesson.LoopKind}; similarity={item.Similarity:0.00}; expected_value={item.SelectionScore:0.00}; reliability={RecoveryReliability(item.Lesson):0.00}; successes={item.Lesson.SuccessCount}; failures={item.Lesson.FailureCount}; avoid: {TrimForMeta(item.Lesson.AvoidPattern, 220)}");
                sb.AppendLine($"   successful strategy: {TrimForMeta(item.Lesson.WinningStrategy, 320)}");
                foreach (var strategyStep in item.Lesson.StrategySteps.Take(4))
                {
                    sb.AppendLine(
                        $"   - {strategyStep.ActionFamily}: intent={TrimForMeta(strategyStep.Intent, 100)}; " +
                        $"target={TrimForMeta(strategyStep.TargetTokens, 100)}; " +
                        $"precondition={TrimForMeta(strategyStep.Preconditions, 120)}; " +
                        $"expected={TrimForMeta(strategyStep.ExpectedEffect, 120)}");
                }
            }
            if (negative.Length > 0)
            {
                sb.AppendLine("NEGATIVE_MEMORY:");
                foreach (var item in negative)
                {
                    var failedPreconditions = string.Join(
                        "; ",
                        item.Lesson.StrategySteps
                            .Select(step => step.Preconditions)
                            .Where(value => !string.IsNullOrWhiteSpace(value))
                            .Distinct(StringComparer.Ordinal)
                            .Take(3));
                    sb.AppendLine(
                        $"- Do not reuse quarantined strategy in this context: " +
                        $"{TrimForMeta(item.Lesson.WinningStrategy, 240)}; " +
                        $"preconditions={TrimForMeta(failedPreconditions, 180)}; " +
                        $"reason={TrimForMeta(item.Lesson.LastFailureReason, 160)}");
                }
            }
            return sb.ToString().TrimEnd();
        }

        internal static double RecoveryReliability(RecoveryLesson lesson) =>
            (lesson.SuccessCount + 1.0) / (lesson.SuccessCount + lesson.FailureCount + 2.0);

        internal static double ContextualBanditScore(RecoveryLesson lesson, double similarity)
        {
            if (similarity <= 0)
                return 0;

            var attempts = Math.Max(0, lesson.SuccessCount + lesson.FailureCount);
            var reliability = RecoveryReliability(lesson);
            var uncertainty = Math.Sqrt(Math.Log(attempts + 2.0) / (attempts + 1.0));
            var ageDays = Math.Max(0, (DateTime.UtcNow - lesson.UpdatedUtc).TotalDays);
            var recency = Math.Exp(-ageDays / 180.0);
            var rewardMean = lesson.RewardObservationCount > 0
                ? lesson.CumulativeReward / lesson.RewardObservationCount
                : reliability;
            var exploration = Math.Sqrt(
                Math.Log(Math.Max(2, lesson.SelectionCount + attempts + 2.0)) /
                Math.Max(1, lesson.SelectionCount + 1.0));
            var normalizedCost = Math.Clamp(lesson.AverageActionCost / 2.0, 0, 1);
            return Math.Clamp(
                similarity * (
                    0.54 +
                    0.16 * reliability +
                    0.16 * Math.Clamp(rewardMean, 0, 1) +
                    0.06 * recency -
                    0.05 * normalizedCost) +
                0.04 * uncertainty +
                0.03 * Math.Clamp(exploration, 0, 1),
                0,
                1);
        }

        static void TrackSuggestedStrategyAttempt(
            RecoveryEpisodeState episode,
            ResolvedActionSnapshot action,
            IReadOnlyCollection<RecoveryLesson> lessons)
        {
            if (episode.AppliedLessonId != null)
                return;

            var explicitLessonId = action.Action.RecoveryStrategyId?.Trim();
            var candidateIds = string.IsNullOrWhiteSpace(explicitLessonId)
                ? episode.SuggestedLessonIds.ToArray()
                : [explicitLessonId];
            foreach (var lessonId in candidateIds)
            {
                if (episode.RejectedLessonIds.Contains(lessonId))
                    continue;
                if (!episode.SuggestedLessonIds.Contains(lessonId, StringComparer.Ordinal))
                {
                    Console.WriteLine($"[memory] ignored unknown or non-suggested recovery_strategy_id={lessonId}.");
                    return;
                }

                var lesson = lessons.FirstOrDefault(item => string.Equals(item.Id, lessonId, StringComparison.Ordinal));
                if (lesson is null || !IsLessonActive(lesson))
                    continue;
                var expected = lesson.StrategySteps.Count > 0
                    ? lesson.StrategySteps
                    : lesson.WinningActionTypes
                        .Select(family => new RecoveryStrategyStep { ActionFamily = family })
                        .ToList();
                if (expected.Count == 0)
                    continue;

                episode.SuggestedLessonProgress.TryGetValue(lessonId, out var progress);
                if (progress >= expected.Count)
                    continue;

                if (action.Action.RecoveryStrategyStep is int explicitStep &&
                    explicitStep != progress + 1)
                {
                    Console.WriteLine(
                        $"[memory] ignored out-of-order recovery strategy step {explicitStep}; expected {progress + 1} for {lessonId}.");
                    return;
                }

                if (StrategyStepMatches(
                        expected[progress],
                        action,
                        explicitReference: !string.IsNullOrWhiteSpace(explicitLessonId)))
                {
                    progress++;
                }
                else if (string.IsNullOrWhiteSpace(explicitLessonId) &&
                         StrategyStepMatches(expected[0], action, explicitReference: false))
                {
                    progress = 1;
                }
                else if (IsLocalObservationAction(action.Action) &&
                         !expected.Any(step => string.Equals(
                             step.ActionFamily,
                             "observe",
                             StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(explicitLessonId))
                    {
                        Console.WriteLine(
                            $"[memory] recovery_strategy_id={lessonId} did not match the declared semantic step; attribution skipped.");
                        return;
                    }
                    progress = 0;
                }

                episode.SuggestedLessonProgress[lessonId] = progress;
                if (progress < expected.Count)
                    continue;

                episode.AppliedLessonId = lessonId;
                episode.AppliedLessonNoProgressObservations = 0;
                RecordLessonSelection(lesson);
                Console.WriteLine($"[memory] detected application of suggested lesson {lessonId}; awaiting visible outcome.");
                return;
            }
        }

        static bool StrategyStepMatches(
            RecoveryStrategyStep expected,
            ResolvedActionSnapshot actual,
            bool explicitReference = false)
        {
            var actualFamily = ActionFamily(actual.Action);
            if (!string.Equals(expected.ActionFamily, actualFamily, StringComparison.OrdinalIgnoreCase))
                return false;

            if (!string.IsNullOrWhiteSpace(expected.ParameterSignature) &&
                !string.Equals(
                    expected.ParameterSignature,
                    SemanticParameterSignature(actual.Action),
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(expected.TargetTokens) &&
                (string.IsNullOrWhiteSpace(actual.SemanticTokens) ||
                 TokenSimilarity(expected.TargetTokens, actual.SemanticTokens) <
                 (explicitReference ? 0.18 : 0.25)))
            {
                return false;
            }

            return true;
        }

        static bool AppliedLessonOutcomeMatches(
            RecoveryEpisodeState episode,
            IReadOnlyCollection<RecoveryLesson> lessons,
            string outcomeTokens,
            byte[] outcomeFingerprint,
            byte[] outcomeActiveWindowFingerprint)
        {
            if (episode.AppliedLessonId is not string lessonId)
                return true;
            var lesson = lessons.FirstOrDefault(item =>
                string.Equals(item.Id, lessonId, StringComparison.Ordinal));
            if (lesson is null)
                return true;

            var triggerTokens = NormalizeTokens(
                $"{lesson.WindowTitleTokens} {lesson.FocusTokens}");
            var tokensAreDiscriminative =
                TokenSimilarity(lesson.ExpectedOutcomeTokens, triggerTokens) < 0.80;
            var tokenMatch = tokensAreDiscriminative &&
                             !string.IsNullOrWhiteSpace(lesson.ExpectedOutcomeTokens) &&
                             !string.IsNullOrWhiteSpace(outcomeTokens) &&
                             TokenSimilarity(lesson.ExpectedOutcomeTokens, outcomeTokens) >= 0.20;
            var visualMatch = false;
            if (TryDecodeFingerprint(
                    lesson.ExpectedOutcomeActiveWindowFingerprintBase64,
                    out var expectedWindow) &&
                expectedWindow.Length == outcomeActiveWindowFingerprint.Length)
            {
                visualMatch = ComputeImageDelta(
                    expectedWindow,
                    outcomeActiveWindowFingerprint) < 0.08;
            }
            else if (TryDecodeFingerprint(
                         lesson.ExpectedOutcomeFingerprintBase64,
                         out var expectedScreen) &&
                     expectedScreen.Length == outcomeFingerprint.Length)
            {
                visualMatch = ComputeImageDelta(expectedScreen, outcomeFingerprint) < 0.08;
            }

            return tokenMatch || visualMatch ||
                   (string.IsNullOrWhiteSpace(lesson.ExpectedOutcomeTokens) &&
                    string.IsNullOrWhiteSpace(lesson.ExpectedOutcomeFingerprintBase64) &&
                    string.IsNullOrWhiteSpace(lesson.ExpectedOutcomeActiveWindowFingerprintBase64));
        }

        static void RegisterAppliedStrategyFailure(
            RecoveryEpisodeState episode,
            List<RecoveryLesson> lessons,
            string reason)
        {
            if (episode.AppliedLessonId is not string lessonId)
                return;

            var lesson = lessons.FirstOrDefault(item => string.Equals(item.Id, lessonId, StringComparison.Ordinal));
            episode.RejectedLessonIds.Add(lessonId);
            episode.SuggestedLessonIds.RemoveAll(id => string.Equals(id, lessonId, StringComparison.Ordinal));
            episode.SuggestedLessonProgress.Remove(lessonId);
            episode.AppliedLessonId = null;
            episode.AppliedLessonNoProgressObservations = 0;

            if (lesson is null)
                return;

            RecordLessonFailure(lesson);
            lesson.ConsecutiveFailureCount++;
            lesson.LastFailureUtc = DateTime.UtcNow;
            lesson.UpdatedUtc = DateTime.UtcNow;
            lesson.LastFailureReason = reason;

            if (lesson.ConsecutiveFailureCount >= Math.Max(1, RecoveryMemoryFailureLimit))
            {
                lesson.Status = "quarantined";
                lesson.QuarantinedUtc = DateTime.UtcNow;
                Console.WriteLine($"[memory] quarantined recovery lesson {lesson.Id} after {lesson.ConsecutiveFailureCount} consecutive failures; reason={reason}");
            }
            else
            {
                Console.WriteLine($"[memory] recovery lesson {lesson.Id} failed ({lesson.ConsecutiveFailureCount}/{RecoveryMemoryFailureLimit}); reason={reason}");
            }
            SaveRecoveryLessons(lessons);
        }

        internal static string ActionFamily(ActionDto? action)
        {
            if (action is null || string.IsNullOrWhiteSpace(action.Type))
                return "unknown";

            return action.Type.ToLowerInvariant() switch
            {
                "click" or "double_click" => "click",
                "drag_drop" => "drag_drop",
                "drag_path" => "path_gesture",
                "move" => "move",
                "aim" or "point" or "request_crop" => "observe",
                "focus_uia" or "click_uia" => "uia",
                "type_text" or "paste_text" => "text_input",
                "keys" => "keys",
                "hold_keys" => "key_hold",
                "scroll" => "scroll",
                "wait" => "wait",
                var type => type
            };
        }

        internal static string ClassifyLoopKind(
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot? previousAction,
            UiPromptContext context)
        {
            var actions = recentActions.TakeLast(16).ToList();
            if (previousAction != null &&
                (actions.Count == 0 || !ReferenceEquals(actions[^1], previousAction)))
            {
                actions.Add(previousAction);
            }

            var currentFamily = ActionFamily(previousAction?.Action ?? actions.LastOrDefault()?.Action);
            var process = NormalizeText(context.ActiveProcessName);
            if (IsTerminalProcess(process) &&
                (currentFamily is "text_input" or "keys" or "run_command" ||
                 actions.TakeLast(5).Any(a => ActionFamily(a.Action) is "text_input" or "keys" or "run_command")))
            {
                return ConsoleInteractionLoop;
            }

            if (currentFamily == "drag_drop")
                return PlacementOrDragLoop;

            if (currentFamily == "path_gesture")
                return previousAction?.Action.GestureKind is "draw" or "lasso"
                    ? RasterCanvasPointerLoop
                    : PointerRegionLoop;

            var visualContext = $"{context.ActiveWindowTitle} {context.FocusedUiaSummary}".ToLowerInvariant();
            if (currentFamily is "click" or "move" &&
                LooksLikeRasterSurface(process, visualContext, actions))
            {
                return RasterCanvasPointerLoop;
            }

            var cycleLength = RepeatedActionCycleLength(actions);
            if (cycleLength == 2)
                return AlternatingActionCycle;
            if (cycleLength >= 3)
                return MultiStepActionCycle;

            if (currentFamily == "run_command")
                return CommandFailureLoop;
            if (currentFamily == "text_input")
                return TextInputNoChange;
            if (currentFamily == "keys")
                return RepeatedKeySequence;

            if (currentFamily is "click" or "move")
            {
                var recentPointerActions = actions.TakeLast(6)
                    .Where(a => ActionFamily(a.Action) is "move" or "click")
                    .ToArray();
                if (recentPointerActions.Any(a => ActionFamily(a.Action) == "move") &&
                    recentPointerActions.Any(a => ActionFamily(a.Action) == "click") &&
                    recentPointerActions.Any(HasPlacementIntent))
                {
                    return PlacementOrDragLoop;
                }
                return PointerRegionLoop;
            }

            if (currentFamily == "uia")
                return UiaInteractionLoop;
            if (currentFamily == "scroll")
                return ScrollNoProgress;
            if (currentFamily is "open_url" or "launch_app")
                return NavigationBounce;
            if (currentFamily == "observe")
                return ObservationLoop;
            if (currentFamily == "wait")
                return WaitNoChange;
            return GenericNoProgressLoop;
        }

        internal static double LessonSimilarity(
            RecoveryLesson lesson,
            UiPromptContext context,
            byte[] shotFingerprint,
            byte[] activeWindowFingerprint,
            string actionFamily,
            string loopKind,
            string loopTopology,
            string interactionDomain,
            string goalTokens,
            string goalDomain,
            string goalIntentTokens,
            string goalMode)
        {
            var process = NormalizeText(context.ActiveProcessName);
            if (!string.IsNullOrEmpty(lesson.ActiveProcess) &&
                !string.Equals(lesson.ActiveProcess, process, StringComparison.Ordinal))
            {
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(lesson.LoopKind) &&
                !string.Equals(lesson.LoopKind, loopKind, StringComparison.Ordinal))
            {
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(lesson.InteractionDomain) &&
                !string.Equals(lesson.InteractionDomain, interactionDomain, StringComparison.Ordinal))
            {
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(lesson.GoalDomain) &&
                !string.Equals(lesson.GoalDomain, goalDomain, StringComparison.Ordinal))
            {
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(lesson.GoalMode) &&
                !string.Equals(lesson.GoalMode, goalMode, StringComparison.Ordinal))
            {
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(lesson.GoalIntentTokens) &&
                !string.IsNullOrWhiteSpace(goalIntentTokens) &&
                TokenSimilarity(lesson.GoalIntentTokens, goalIntentTokens) < 0.5)
            {
                return 0;
            }

            var titleSimilarity = TokenSimilarity(lesson.WindowTitleTokens, NormalizeTokens(context.ActiveWindowTitle));
            var focusSimilarity = TokenSimilarity(lesson.FocusTokens, NormalizeTokens(context.FocusedUiaSummary));
            var goalSimilarity = string.IsNullOrWhiteSpace(lesson.GoalTokens)
                ? 0.5
                : TokenSimilarity(lesson.GoalTokens, goalTokens);
            if (!string.IsNullOrWhiteSpace(lesson.GoalTokens) && goalSimilarity < 0.15)
                return 0;
            var imageSimilarity = 0.0;
            if (TryDecodeFingerprint(lesson.ScreenFingerprintBase64, out var stored) &&
                stored.Length == shotFingerprint.Length &&
                stored.Length > 0)
            {
                var delta = ComputeImageDelta(stored, shotFingerprint);
                imageSimilarity = Math.Max(0, 1.0 - delta / 0.15);
            }
            var activeWindowSimilarity = 0.0;
            if (TryDecodeFingerprint(lesson.ActiveWindowFingerprintBase64, out var storedWindow) &&
                storedWindow.Length == activeWindowFingerprint.Length &&
                storedWindow.Length > 0)
            {
                var delta = ComputeImageDelta(storedWindow, activeWindowFingerprint);
                activeWindowSimilarity = Math.Max(0, 1.0 - delta / 0.15);
            }

            // A shared host process (for example msedge) and the same action type are
            // not enough to call two loops analogous.
            if (titleSimilarity < 0.25 &&
                focusSimilarity < 0.25 &&
                imageSimilarity < 0.55 &&
                activeWindowSimilarity < 0.55)
                return 0;

            var score = string.Equals(lesson.ActiveProcess, process, StringComparison.Ordinal) ? 0.10 : 0.03;
            score += 0.10 * titleSimilarity;
            score += 0.08 * focusSimilarity;
            if (string.Equals(lesson.TriggerActionFamily, actionFamily, StringComparison.OrdinalIgnoreCase))
                score += 0.08;
            if (string.Equals(lesson.LoopKind, loopKind, StringComparison.Ordinal))
                score += 0.16;
            if (string.IsNullOrWhiteSpace(lesson.LoopTopology) ||
                string.Equals(lesson.LoopTopology, loopTopology, StringComparison.Ordinal))
                score += 0.10;
            if (string.IsNullOrWhiteSpace(lesson.InteractionDomain) ||
                string.Equals(lesson.InteractionDomain, interactionDomain, StringComparison.Ordinal))
                score += 0.10;
            score += 0.08 * imageSimilarity;
            score += 0.12 * activeWindowSimilarity;
            score += 0.08 * goalSimilarity;

            return Math.Clamp(score, 0, 1);
        }

        static RecoveryLesson BuildCandidateLesson(
            RecoveryEpisodeState episode,
            ResolvedActionSnapshot winningAction,
            double lastDelta,
            string expectedOutcomeTokens,
            byte[] expectedOutcomeFingerprint,
            byte[] expectedOutcomeActiveWindowFingerprint)
        {
            var recoveryActions = episode.RecoveryActions.Count > 0
                ? episode.RecoveryActions
                : [winningAction];
            if (!ReferenceEquals(recoveryActions.LastOrDefault(), winningAction) &&
                recoveryActions.LastOrDefault()?.Description != winningAction.Description)
            {
                recoveryActions = recoveryActions.Append(winningAction).ToList();
            }

            var winningTail = RecoveryTail(recoveryActions, episode.TriggerActionFamily);
            var strategySteps = winningTail.Select(BuildStrategyStep).ToList();
            return new RecoveryLesson
            {
                ActiveProcess = NormalizeText(episode.TriggerContext.ActiveProcessName),
                GoalTokens = episode.GoalTokens,
                GoalDomain = episode.GoalDomain,
                GoalIntentTokens = episode.GoalIntentTokens,
                GoalMode = episode.GoalMode,
                WindowTitleTokens = NormalizeTokens(episode.TriggerContext.ActiveWindowTitle),
                FocusTokens = NormalizeTokens(episode.TriggerContext.FocusedUiaSummary),
                TriggerActionFamily = episode.TriggerActionFamily,
                LoopKind = episode.LoopKind,
                LoopTopology = episode.LoopTopology,
                InteractionDomain = episode.InteractionDomain,
                AvoidPattern = DescribeAvoidPattern(episode.LoopKind, episode.FailedActions, episode.MaxStagnationSteps),
                WinningStrategy = DescribeWinningStrategy(strategySteps, lastDelta),
                WinningActionTypes = winningTail.Select(a => ActionFamily(a.Action)).ToArray(),
                StrategySteps = strategySteps,
                StrategySignature = StrategySignature(strategySteps),
                ScreenFingerprintBase64 = Convert.ToBase64String(episode.TriggerFingerprint),
                ActiveWindowFingerprintBase64 = Convert.ToBase64String(
                    episode.TriggerActiveWindowFingerprint),
                ExpectedOutcomeTokens = expectedOutcomeTokens,
                ExpectedOutcomeFingerprintBase64 = Convert.ToBase64String(
                    expectedOutcomeFingerprint),
                ExpectedOutcomeActiveWindowFingerprintBase64 = Convert.ToBase64String(
                    expectedOutcomeActiveWindowFingerprint),
                LastSuccessUtc = DateTime.UtcNow,
                AverageActionCost = StrategyActionCost(strategySteps),
                LastProgressConfidence = episode.LastProgressConfidence,
                LastProgressEvidence = episode.LastProgressEvidence,
                ValidationSource = RecoveryProgressVerificationEnabled
                    ? "independent-progress-verifier"
                    : "local-transition-fallback",
                RDPilotVersion = typeof(RDPilotApplication).Assembly.GetName().Version?.ToString() ?? "",
                PromptVersion = PromptCacheKey ?? "rdpilot-control",
                ModelName = Model
            };
        }

        static List<ResolvedActionSnapshot> RecoveryTail(
            IReadOnlyList<ResolvedActionSnapshot> actions,
            string triggerFamily)
        {
            if (actions.Count == 0)
                return [];

            var start = Math.Max(0, actions.Count - 4);
            for (var i = actions.Count - 2; i >= start; i--)
            {
                if (!string.Equals(ActionFamily(actions[i].Action), triggerFamily, StringComparison.OrdinalIgnoreCase))
                {
                    start = i;
                    break;
                }
            }
            return actions.Skip(start).Take(4).ToList();
        }

        static string DescribeAvoidPattern(
            string loopKind,
            IReadOnlyCollection<ResolvedActionSnapshot> failedActions,
            int maxStagnation)
        {
            var families = failedActions
                .Where(a => !IsLocalObservationAction(a.Action))
                .Select(a => ActionFamily(a.Action))
                .ToArray();
            if (families.Length == 0)
                return $"{loopKind}: the previous strategy produced no visible progress for {Math.Max(1, maxStagnation)} steps";

            var groups = families
                .GroupBy(x => x)
                .OrderByDescending(g => g.Count())
                .Select(g => $"{g.Count()}× {g.Key}")
                .Take(3);
            return $"{loopKind}: repeating {string.Join(", ", groups)} produced no visible progress for up to {Math.Max(1, maxStagnation)} steps";
        }

        static string DescribeWinningStrategy(
            IReadOnlyList<RecoveryStrategyStep> steps,
            double delta)
        {
            var descriptions = steps.Select(DescribeStrategyStep).Where(x => x.Length > 0).ToArray();
            if (descriptions.Length == 0)
                return $"choose a materially different UI route; it previously produced visible progress (delta {delta:0.####})";

            return $"{string.Join(" -> ", descriptions)}; this sequence previously produced visible progress (delta {delta:0.####})";
        }

        static RecoveryStrategyStep BuildStrategyStep(ResolvedActionSnapshot snapshot)
        {
            var action = snapshot.Action;
            var family = ActionFamily(action);
            var semanticIntent = NormalizeTokens(action.Note);
            var intent = DefaultActionIntent(family, action);
            if (!string.IsNullOrWhiteSpace(semanticIntent))
                intent += $"; semantic intent: {TrimForMeta(semanticIntent, 80)}";
            return new RecoveryStrategyStep
            {
                ActionFamily = family,
                Intent = intent,
                TargetTokens = snapshot.SemanticTokens,
                Preconditions = StrategyPreconditions(family),
                ExpectedEffect = StrategyExpectedEffect(family),
                ParameterSignature = SemanticParameterSignature(action)
            };
        }

        static string DescribeStrategyStep(RecoveryStrategyStep step)
        {
            var target = string.IsNullOrWhiteSpace(step.TargetTokens)
                ? ""
                : $" target[{step.TargetTokens}]";
            var parameters = string.IsNullOrWhiteSpace(step.ParameterSignature)
                ? ""
                : $" ({step.ParameterSignature})";
            return $"{step.ActionFamily}{parameters}{target}: {step.Intent}";
        }

        static string DefaultActionIntent(string family, ActionDto action) =>
            family switch
            {
                "observe" => "re-observe the ambiguous target with a fresh AIM or crop",
                "move" => "move the pointer first and inspect the hover or placement state",
                "click" => "click once only after the target state is visibly valid",
                "drag_drop" => "drag the current source object to the semantically correct destination",
                "path_gesture" => $"perform the bounded {action.GestureKind ?? "pointer"} path once on the intended surface",
                "uia" => "use the matching current UI Automation target",
                "keys" => $"use the keyboard action [{string.Join("+", action.Keys ?? [])}]",
                "key_hold" => $"hold [{string.Join("+", action.Keys ?? [])}] for {EffectiveKeyHoldDurationMs(action)} ms",
                "scroll" => (action.ScrollDy ?? 0) >= 0 ? "scroll down to a different visible route" : "scroll up to a different visible route",
                "text_input" => "establish editable focus, then enter the text once",
                "wait" => "wait for the visible operation to settle",
                "open_url" => "navigate through the currently valid URL route",
                "launch_app" => "launch the required application through the current UI",
                _ => $"switch to action type {family}"
            };

        static string StrategyPreconditions(string family) =>
            family switch
            {
                "click" => "target is visible, enabled, and unambiguous",
                "drag_drop" => "source and destination are both visible; source is draggable",
                "path_gesture" => "the intended gesture surface and starting point are visible",
                "text_input" => "editable control has focus",
                "keys" => "the intended window/control has focus",
                "key_hold" => "the realtime target has focus and is ready for held input",
                "uia" => "a semantically matching UIA target exists",
                "scroll" => "the intended scroll container is active",
                _ => "current visible state matches the learned context"
            };

        static string StrategyExpectedEffect(string family) =>
            family switch
            {
                "click" => "target state or active view changes",
                "drag_drop" => "source leaves its old location and destination accepts it",
                "path_gesture" => "the expected local path or realtime response appears",
                "text_input" => "text appears once in the intended control",
                "keys" => "focus or application state changes as intended",
                "key_hold" => "the application reacts during the bounded hold and all keys are released",
                "scroll" => "new content becomes visible",
                "observe" => "ambiguity is reduced without mutating the UI",
                _ => "the loop state is left and visible goal progress persists"
            };

        static double StrategyActionCost(IReadOnlyCollection<RecoveryStrategyStep> steps)
        {
            if (steps.Count == 0)
                return 1;
            return steps.Average(step => step.ActionFamily switch
            {
                "observe" or "move" => 0.25,
                "click" or "uia" or "keys" or "scroll" => 0.5,
                "key_hold" => 0.65,
                "text_input" => 0.7,
                "drag_drop" => 0.8,
                "path_gesture" => 0.8,
                "wait" => 0.9,
                "open_url" or "launch_app" => 1.2,
                "run_command" => 1.5,
                _ => 0.75
            });
        }

        static string SemanticParameterSignature(ActionDto action) =>
            ActionFamily(action) switch
            {
                "keys" => string.Join("+", action.Keys ?? []).ToLowerInvariant(),
                "key_hold" => $"{string.Join("+", action.Keys ?? []).ToLowerInvariant()}:{EffectiveKeyHoldDurationMs(action)}ms",
                "scroll" => (action.ScrollDy ?? 0) >= 0 ? "down" : "up",
                "text_input" => string.IsNullOrEmpty(action.Text) ? "" : $"text-length:{Math.Min(32, action.Text.Length / 8)}",
                "drag_drop" => action.Button?.ToLowerInvariant() ?? "left",
                "path_gesture" => $"{action.GestureKind ?? "other"}:{action.Path?.Length ?? 0}-points",
                "click" => action.Button?.ToLowerInvariant() ?? "left",
                _ => ""
            };

        static string StrategySignature(IReadOnlyCollection<RecoveryStrategyStep> steps) =>
            string.Join(
                "->",
                steps.Select(step =>
                    $"{step.ActionFamily}:{step.ParameterSignature}:{NormalizeTokens(step.TargetTokens)}"));

        static void StoreConfirmedLesson(RecoveryLesson candidate, List<RecoveryLesson> lessons)
        {
            candidate.LastSuccessUtc ??= DateTime.UtcNow;
            candidate.UpdatedUtc = DateTime.UtcNow;
            var existing = lessons.FirstOrDefault(lesson =>
                string.Equals(lesson.ActiveProcess, candidate.ActiveProcess, StringComparison.Ordinal) &&
                string.Equals(lesson.LoopKind, candidate.LoopKind, StringComparison.Ordinal) &&
                string.Equals(lesson.LoopTopology, candidate.LoopTopology, StringComparison.Ordinal) &&
                string.Equals(lesson.InteractionDomain, candidate.InteractionDomain, StringComparison.Ordinal) &&
                string.Equals(lesson.GoalDomain, candidate.GoalDomain, StringComparison.Ordinal) &&
                string.Equals(
                    string.IsNullOrWhiteSpace(lesson.GoalMode) ? "finite" : lesson.GoalMode,
                    string.IsNullOrWhiteSpace(candidate.GoalMode) ? "finite" : candidate.GoalMode,
                    StringComparison.Ordinal) &&
                (string.IsNullOrWhiteSpace(lesson.GoalIntentTokens) ||
                 string.IsNullOrWhiteSpace(candidate.GoalIntentTokens) ||
                 TokenSimilarity(lesson.GoalIntentTokens, candidate.GoalIntentTokens) >= 0.5) &&
                string.Equals(lesson.StrategySignature, candidate.StrategySignature, StringComparison.Ordinal) &&
                TokenSimilarity(lesson.GoalTokens, candidate.GoalTokens) >= 0.35 &&
                (string.IsNullOrWhiteSpace(lesson.ExpectedOutcomeTokens) ||
                 string.IsNullOrWhiteSpace(candidate.ExpectedOutcomeTokens) ||
                 TokenSimilarity(lesson.ExpectedOutcomeTokens, candidate.ExpectedOutcomeTokens) >= 0.25) &&
                (TokenSimilarity(lesson.FocusTokens, candidate.FocusTokens) >= 0.35 ||
                 TokenSimilarity(lesson.WindowTitleTokens, candidate.WindowTitleTokens) >= 0.5));

            if (existing is null)
            {
                RecordLessonSuccess(candidate, candidate.LastProgressConfidence);
                lessons.Add(candidate);
                existing = candidate;
            }
            else
            {
                existing.UpdatedUtc = DateTime.UtcNow;
                RecordLessonSuccess(existing, candidate.LastProgressConfidence);
                existing.ConsecutiveFailureCount = 0;
                existing.LastSuccessUtc = DateTime.UtcNow;
                existing.Status = "active";
                existing.QuarantinedUtc = null;
                existing.LastFailureReason = "";
                existing.AvoidPattern = candidate.AvoidPattern;
                existing.WinningStrategy = candidate.WinningStrategy;
                existing.GoalTokens = candidate.GoalTokens;
                existing.GoalDomain = candidate.GoalDomain;
                existing.GoalIntentTokens = candidate.GoalIntentTokens;
                existing.FocusTokens = candidate.FocusTokens;
                existing.ScreenFingerprintBase64 = candidate.ScreenFingerprintBase64;
                existing.ActiveWindowFingerprintBase64 = candidate.ActiveWindowFingerprintBase64;
                existing.ExpectedOutcomeTokens = candidate.ExpectedOutcomeTokens;
                existing.ExpectedOutcomeFingerprintBase64 = candidate.ExpectedOutcomeFingerprintBase64;
                existing.ExpectedOutcomeActiveWindowFingerprintBase64 =
                    candidate.ExpectedOutcomeActiveWindowFingerprintBase64;
                existing.StrategySteps = candidate.StrategySteps;
                existing.AverageActionCost = candidate.AverageActionCost;
                existing.LastProgressConfidence = candidate.LastProgressConfidence;
                existing.LastProgressEvidence = candidate.LastProgressEvidence;
                existing.ValidationSource = candidate.ValidationSource;
                existing.RDPilotVersion = candidate.RDPilotVersion;
                existing.PromptVersion = candidate.PromptVersion;
                existing.ModelName = candidate.ModelName;
            }

            var persisted = SaveRecoveryLessons(lessons);
            Console.WriteLine(
                persisted
                    ? $"[memory] confirmed and persisted recovery lesson {existing.Id}; successes={existing.SuccessCount}; strategy={existing.WinningStrategy}"
                    : $"[memory] confirmed recovery lesson {existing.Id} in memory; durable persistence is pending; strategy={existing.WinningStrategy}");
        }

        static string NormalizeText(string? value) =>
            string.IsNullOrWhiteSpace(value) ? "" : value.Trim().ToLowerInvariant();

        static string NormalizeGoalTokens(string? goal)
        {
            if (string.IsNullOrWhiteSpace(goal))
                return "";

            var stopWords = new HashSet<string>(
                [
                    "the", "and", "for", "with", "this", "that", "from", "into",
                    "oraz", "aby", "żeby", "sie", "się", "jest", "tego", "tym",
                    "proszę", "please", "zrób", "wykonaj"
                ],
                StringComparer.Ordinal);
            return string.Join(' ', Regex.Matches(goal.ToLowerInvariant(), @"[\p{L}\p{Nd}]{2,}")
                .Cast<Match>()
                .Select(match => match.Value)
                .Where(token => !stopWords.Contains(token))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(token => token, StringComparer.Ordinal)
                .Take(32));
        }

        static string ClassifyGoalDomain(string? goal)
        {
            var normalized = NormalizeText(goal);
            var tokens = Regex.Matches(normalized, @"[\p{L}\p{Nd}]{2,}")
                .Cast<Match>()
                .Select(match => match.Value)
                .ToArray();
            static bool Has(
                IReadOnlyCollection<string> values,
                IReadOnlyCollection<string> exact,
                params string[] prefixes) =>
                values.Any(token =>
                    exact.Contains(token, StringComparer.Ordinal) ||
                    prefixes.Any(prefix => token.StartsWith(prefix, StringComparison.Ordinal)));

            if (Has(tokens, ["game", "puzzle", "gra", "grę", "gry", "graj", "tile", "piece", "level"], "układank", "kafel", "plansz"))
                return "game";
            if (Has(tokens, ["password", "hasło", "login", "logowanie", "authenticate"], "uwierzyteln"))
                return "authentication";
            if (Has(tokens, ["pay", "payment", "bank", "transfer", "purchase", "kup"], "płat", "przelew", "finans"))
                return "financial";
            if (Has(tokens, ["send", "email", "mail", "message", "wiadomość", "wyślij", "slack", "teams"], "wiadomo", "wysył"))
                return "communication";
            if (Has(tokens, ["delete", "remove", "erase", "format", "usuń", "kasuj", "wyczyść"], "usuw", "kasow"))
                return "destructive";
            if (Has(tokens, ["code", "build", "compile", "test", "kod", "program", "projekt"], "kompil", "programow"))
                return "development";
            return "general";
        }

        internal static string ClassifyGoalMode(string? goal)
        {
            var normalized = NormalizeText(goal);
            if (string.IsNullOrWhiteSpace(normalized))
                return "finite";

            var explicitlyUserTerminated = Regex.IsMatch(
                normalized,
                @"\b(until\s+(?:i|the\s+user)\s+stop|until\s+stopped|until\s+cancelled|until\s+canceled|aż\s+(?:ja\s+)?przerwę|dopóki\s+(?:ja\s+)?nie\s+przerwę|do\s+odwołania)\b",
                RegexOptions.CultureInvariant);
            if (explicitlyUserTerminated)
                return "continuous";

            var hasExplicitEnd = Regex.IsMatch(
                normalized,
                @"\b(until|aż|dopóki|complete\w*|finish\w*|ukończ\p{L}*|zakończ\p{L}*|wygraj\p{L}*|win|winning|won|one\s+round|jedną\s+rundę|first\s+level|pierwsz\p{L}*\s+poziom\p{L}*|level\s+\d+|poziom\s+\d+|when\s+done|gdy\s+zakończ\p{L}*|do\s+godziny|przez\s+\d+|for\s+\d+\s+(?:seconds?|minutes?|hours?)|\d+\s+(?:times?|razy|iterations?|iteracj\p{L}*))\b",
                RegexOptions.CultureInvariant);
            if (hasExplicitEnd)
                return "finite";

            return Regex.IsMatch(
                normalized,
                @"\b(keep\s+\w+ing|keep\s+running|continue\s+\w+ing|indefinitely|continuously|ongoing|repeatedly|periodically|regularly|cyclically|recurring|monitor|watch|observe|maintain|supervise|respond\s+to\s+new|handle\s+incoming|stay\s+active|graj|pograj|kontynuuj|bez\s+końca|ciągle|stale|regularnie|okresowo|cyklicznie|powtarzaj|sprawdzaj|monitoruj|obserwuj|pilnuj|nadzoruj|utrzymuj|reaguj\s+na|obsługuj\s+nowe|śledź|pracuj\s+dalej)\b",
                RegexOptions.CultureInvariant)
                ? "continuous"
                : "finite";
        }

        internal static string ResolveGoalMode(
            string? goal,
            string? configuredMode)
        {
            var normalized = NormalizeText(configuredMode);
            return normalized is "finite" or "continuous"
                ? normalized
                : ClassifyGoalMode(goal);
        }

        static string NormalizeGoalIntentTokens(string? goal)
        {
            if (string.IsNullOrWhiteSpace(goal))
                return "";
            var intentPrefixes = new[]
            {
                "add", "dod", "create", "utw", "open", "otw", "close", "zamk",
                "delete", "remove", "usu", "kas", "move", "przen", "drag", "przeci",
                "copy", "kopi", "send", "wyśl", "left", "lew", "right", "praw",
                "up", "gór", "down", "dół", "before", "przed", "after", "po",
                "enable", "włącz", "disable", "wyłącz", "start", "uruch", "stop", "zatrz"
            };
            return string.Join(' ', Regex.Matches(goal.ToLowerInvariant(), @"[\p{L}\p{Nd}]{2,}")
                .Cast<Match>()
                .Select(match => match.Value)
                .Where(token => intentPrefixes.Any(prefix =>
                    token.StartsWith(prefix, StringComparison.Ordinal)))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(token => token, StringComparer.Ordinal)
                .Take(12));
        }

        static bool IsMeaningfulStateTransition(
            UiPromptContext triggerContext,
            byte[] triggerFingerprint,
            byte[] triggerActiveWindowFingerprint,
            UiPromptContext currentContext,
            byte[] currentFingerprint,
            byte[] currentActiveWindowFingerprint,
            double lastDelta)
        {
            if (!string.Equals(
                    NormalizeText(triggerContext.ActiveProcessName),
                    NormalizeText(currentContext.ActiveProcessName),
                    StringComparison.Ordinal))
            {
                return true;
            }

            var titleSimilarity = TokenSimilarity(
                NormalizeTokens(triggerContext.ActiveWindowTitle),
                NormalizeTokens(currentContext.ActiveWindowTitle));
            var focusSimilarity = TokenSimilarity(
                NormalizeTokens(triggerContext.FocusedUiaSummary),
                NormalizeTokens(currentContext.FocusedUiaSummary));
            if (titleSimilarity < 0.55 &&
                (!string.IsNullOrWhiteSpace(triggerContext.ActiveWindowTitle) ||
                 !string.IsNullOrWhiteSpace(currentContext.ActiveWindowTitle)))
            {
                return true;
            }
            if (focusSimilarity < 0.35 &&
                (!string.IsNullOrWhiteSpace(triggerContext.FocusedUiaSummary) ||
                 !string.IsNullOrWhiteSpace(currentContext.FocusedUiaSummary)))
            {
                return true;
            }

            var triggerDelta = triggerFingerprint.Length == currentFingerprint.Length
                ? ComputeImageDelta(triggerFingerprint, currentFingerprint)
                : lastDelta;
            var activeWindowDelta =
                triggerActiveWindowFingerprint.Length == currentActiveWindowFingerprint.Length
                    ? ComputeImageDelta(
                        triggerActiveWindowFingerprint,
                        currentActiveWindowFingerprint)
                    : triggerDelta;
            return activeWindowDelta >= Math.Max(NoChangeThreshold * 1.5, 0.0075) ||
                   triggerDelta >= Math.Max(NoChangeThreshold * 3.0, 0.015);
        }

        static string ClassifyLoopTopology(
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot? previousAction,
            int stagnationSteps)
        {
            var actions = recentActions.TakeLast(24).ToList();
            if (previousAction != null &&
                (actions.Count == 0 || !ReferenceEquals(actions[^1], previousAction)))
            {
                actions.Add(previousAction);
            }

            var cycleLength = RepeatedActionCycleLength(actions);
            if (cycleLength == 2)
                return "alternating-cycle";
            if (cycleLength >= 3)
                return "multi-step-cycle";
            var meaningful = actions.Where(action => !IsLocalObservationAction(action.Action)).TakeLast(2).ToArray();
            if (meaningful.Length == 2 && ActionsCycleEquivalent(meaningful[0], meaningful[1]))
                return "direct-repeat";
            return stagnationSteps > 0 ? "stagnation" : "state-return";
        }

        static string ClassifyInteractionDomain(
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot? previousAction,
            UiPromptContext context)
        {
            var actions = recentActions.TakeLast(16).ToList();
            if (previousAction != null &&
                (actions.Count == 0 || !ReferenceEquals(actions[^1], previousAction)))
            {
                actions.Add(previousAction);
            }
            var process = NormalizeText(context.ActiveProcessName);
            var family = ActionFamily(previousAction?.Action ?? actions.LastOrDefault()?.Action);
            if (IsTerminalProcess(process))
                return "terminal";
            if (family == "drag_drop" ||
                family != "path_gesture" && actions.Any(HasPlacementIntent))
                return "placement";
            if (family == "path_gesture")
                return previousAction?.Action.GestureKind is "draw" or "lasso"
                    ? "raster-canvas"
                    : "gesture-surface";
            if (family == "uia")
                return "uia";
            var visualContext = $"{context.ActiveWindowTitle} {context.FocusedUiaSummary}".ToLowerInvariant();
            if (LooksLikeRasterSurface(process, visualContext, actions))
                return "raster-canvas";
            if (IsBrowserProcess(process))
                return "browser";
            if (family is "click" or "move" or "scroll")
                return "pointer-ui";
            if (family is "keys" or "key_hold" or "text_input")
                return "keyboard-ui";
            return "generic-ui";
        }

        static int RepeatedActionCycleLength(IReadOnlyList<ResolvedActionSnapshot> actions)
        {
            for (var period = 2; period <= Math.Min(12, actions.Count / 2); period++)
            {
                var tail = actions.TakeLast(period * 2).ToArray();
                var matches = true;
                for (var i = 0; i < period; i++)
                {
                    if (ActionsCycleEquivalent(tail[i], tail[i + period]))
                        continue;
                    matches = false;
                    break;
                }
                if (matches &&
                    tail.Take(period).Skip(1).Any(action =>
                        !ActionsCycleEquivalent(tail[0], action)))
                {
                    return period;
                }
            }
            return 0;
        }

        static bool ActionsCycleEquivalent(
            ResolvedActionSnapshot left,
            ResolvedActionSnapshot right)
        {
            var leftFamily = ActionFamily(left.Action);
            var rightFamily = ActionFamily(right.Action);
            if (!string.Equals(leftFamily, rightFamily, StringComparison.OrdinalIgnoreCase))
                return false;

            if (leftFamily is "click" or "move")
            {
                return left.ScreenPoint is Point leftPoint &&
                       right.ScreenPoint is Point rightPoint &&
                       ScreenPointsAreNearby(leftPoint, rightPoint);
            }

            if (leftFamily == "drag_drop")
            {
                return left.ScreenPoint is Point leftSource &&
                       right.ScreenPoint is Point rightSource &&
                       left.DestinationScreenPoint is Point leftDestination &&
                       right.DestinationScreenPoint is Point rightDestination &&
                       ScreenPointsAreNearby(leftSource, rightSource) &&
                       ScreenPointsAreNearby(leftDestination, rightDestination);
            }

            return string.Equals(
                left.IneffectiveSignature,
                right.IneffectiveSignature,
                StringComparison.OrdinalIgnoreCase);
        }

        internal static bool HasEmergingLoopPattern(
            IReadOnlyCollection<ResolvedActionSnapshot> recentActions,
            ResolvedActionSnapshot? previousAction)
        {
            var actions = recentActions.TakeLast(24).ToList();
            if (previousAction != null &&
                (actions.Count == 0 || !ReferenceEquals(actions[^1], previousAction)))
            {
                actions.Add(previousAction);
            }

            if (RepeatedActionCycleLength(actions) > 0)
                return true;

            var meaningful = actions
                .Where(action => !IsLocalObservationAction(action.Action))
                .TakeLast(2)
                .ToArray();
            if (meaningful.Length < 2)
                return false;

            return ActionsCycleEquivalent(meaningful[0], meaningful[1]);
        }
    }
}
