
namespace RDPilot.Diagnostics;

    /// <summary>
    /// Replays precomputed screen-state sequences without desktop input or API calls.
    /// This gives loop-detector precision/recall measurements on a growing corpus.
    /// </summary>
    internal static class LoopReplayService
    {
        internal static LoopReplayResult ExecuteLoopReplay(string corpusPath)
        {
            var path = Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(corpusPath));
            if (!File.Exists(path))
                throw new FileNotFoundException("Loop replay corpus was not found.", path);

            var corpus = JsonSerializer.Deserialize<LoopReplayCorpus>(
                             File.ReadAllText(path),
                             PrettyJson)
                         ?? throw new InvalidOperationException(
                             "Loop replay corpus could not be parsed.");
            if (corpus.Cases.Count == 0)
                throw new InvalidOperationException(
                    "Loop replay corpus does not contain any cases.");

            var truePositive = 0;
            var falsePositive = 0;
            var trueNegative = 0;
            var falseNegative = 0;
            var independentTruePositive = 0;
            var independentFalsePositive = 0;
            var independentTrueNegative = 0;
            var independentFalseNegative = 0;
            var independentCases = 0;
            foreach (var replayCase in corpus.Cases)
            {
                var result = ReplayCase(replayCase);
                if (replayCase.ExpectedLoop && result.Detected) truePositive++;
                else if (!replayCase.ExpectedLoop && result.Detected) falsePositive++;
                else if (replayCase.ExpectedLoop) falseNegative++;
                else trueNegative++;

                if (replayCase.HasIndependentLabel)
                {
                    independentCases++;
                    if (replayCase.ExpectedLoop && result.Detected)
                        independentTruePositive++;
                    else if (!replayCase.ExpectedLoop && result.Detected)
                        independentFalsePositive++;
                    else if (replayCase.ExpectedLoop)
                        independentFalseNegative++;
                    else
                        independentTrueNegative++;
                }

                Console.WriteLine(
                    $"[{(result.Detected == replayCase.ExpectedLoop ? "PASS" : "FAIL")}] " +
                    $"{(string.IsNullOrWhiteSpace(replayCase.Name) ? "unnamed" : replayCase.Name)}: " +
                    $"expected_loop={replayCase.ExpectedLoop}; detected={result.Detected}; " +
                    $"step={(result.Step > 0 ? result.Step : 0)}; confidence={result.Confidence:0.00}; " +
                    $"label={(replayCase.HasIndependentLabel ? "independent" : "telemetry-regression")}");
            }

            var resultSummary = new LoopReplayResult(truePositive, falsePositive, trueNegative, falseNegative);
            var precision = truePositive + falsePositive == 0
                ? 1
                : truePositive / (double)(truePositive + falsePositive);
            var recall = truePositive + falseNegative == 0
                ? 1
                : truePositive / (double)(truePositive + falseNegative);
            Console.WriteLine(
                $"Regression summary: cases={corpus.Cases.Count}; TP={truePositive}; FP={falsePositive}; " +
                $"TN={trueNegative}; FN={falseNegative}; precision={precision:0.000}; recall={recall:0.000}");
            if (independentCases == 0)
            {
                Console.WriteLine(
                    "Independent accuracy summary: unavailable. Add manually reviewed cases with a non-telemetry labelSource; telemetry-derived labels are regression signals, not an unbiased accuracy estimate.");
                return resultSummary;
            }

            var independentPrecision =
                independentTruePositive + independentFalsePositive == 0
                    ? 1
                    : independentTruePositive /
                      (double)(independentTruePositive +
                               independentFalsePositive);
            var independentRecall =
                independentTruePositive + independentFalseNegative == 0
                    ? 1
                    : independentTruePositive /
                      (double)(independentTruePositive +
                               independentFalseNegative);
            Console.WriteLine(
                $"Independent accuracy summary: cases={independentCases}; TP={independentTruePositive}; FP={independentFalsePositive}; " +
                $"TN={independentTrueNegative}; FN={independentFalseNegative}; precision={independentPrecision:0.000}; recall={independentRecall:0.000}");
            return resultSummary;
        }

        internal static void ImportIndependentLoopReplayCorpus(string sourcePath) =>
            LoopReplayCorpusStore.Import(sourcePath);

        internal static (bool Detected, int Step, double Confidence) ReplayCase(
            LoopReplayCase replayCase)
        {
            ReplayCaseValidator.Validate(replayCase);
            var graph = new LoopStateGraph();
            var recentActions = new Queue<ResolvedActionSnapshot>();
            int? previousEpoch = null;
            var highestConfidence = 0.0;
            var highestConfidenceStep = 0;
            for (var index = 0; index < replayCase.Frames.Count; index++)
            {
                var frame = replayCase.Frames[index];
                if (frame.StateEpoch is int epoch)
                {
                    if (previousEpoch is int prior && prior != epoch)
                    {
                        graph = new LoopStateGraph();
                        recentActions.Clear();
                    }
                    previousEpoch = epoch;
                }
                var resolved = ReplayCaseValidator.ResolveFrame(replayCase, index);
                var previousAction = resolved.PreviousAction;
                if (resolved.RecentActions is not null)
                    recentActions = new Queue<ResolvedActionSnapshot>(resolved.RecentActions);
                else if (previousAction is not null)
                {
                    recentActions.Enqueue(previousAction);
                    while (recentActions.Count > RecentActionLimit) recentActions.Dequeue();
                }

                var context = new UiPromptContext(
                    frame.WindowTitle,
                    frame.ActiveProcess,
                    frame.FocusSummary,
                    null,
                    null,
                    null) { SemanticFeatures = frame.ContextFeatures };
                var assessment = AssessVisualStateCycle(
                    graph,
                    resolved.Screen,
                    resolved.ActiveWindow,
                    context,
                    index + 1,
                    recentActions,
                    previousAction,
                    frame.LastDelta ?? double.NaN,
                    recordLearning: false,
                    goalMode: replayCase.GoalMode,
                    recurringWorkflowIntent:
                        replayCase.RecurringWorkflowIntent,
                    goalProgress: Enum.TryParse<GoalProgressState>(frame.GoalProgress, true, out var progress) &&
                        Enum.IsDefined(progress) ? progress : GoalProgressState.Unknown);
                if (assessment.Confidence > highestConfidence)
                {
                    highestConfidence = assessment.Confidence;
                    highestConfidenceStep = index + 1;
                }
                if (assessment.IsLoop)
                    return (true, index + 1, assessment.Confidence);
            }
            return (false, highestConfidenceStep, highestConfidence);
        }

    }
