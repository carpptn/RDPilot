internal static partial class RDPilotApplication
{
    /// <summary>
    /// Replays precomputed screen-state sequences without desktop input or API calls.
    /// This gives loop-detector precision/recall measurements on a growing corpus.
    /// </summary>
    internal static class LoopReplayService
    {
        internal static void ExecuteLoopReplay(string corpusPath)
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
                return;
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
        }

        internal static void ImportIndependentLoopReplayCorpus(
            string sourcePath)
        {
            var source = Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(sourcePath));
            if (!File.Exists(source))
                throw new FileNotFoundException(
                    "Independent loop replay corpus was not found.",
                    source);

            var imported = JsonSerializer.Deserialize<LoopReplayCorpus>(
                               File.ReadAllText(source),
                               PrettyJson)
                           ?? throw new InvalidOperationException(
                               "Independent loop replay corpus could not be parsed.");
            if (imported.Cases.Count == 0)
                throw new InvalidOperationException(
                    "Independent loop replay corpus does not contain any cases.");
            foreach (var replayCase in imported.Cases)
            {
                if (string.IsNullOrWhiteSpace(replayCase.Name))
                    throw new InvalidOperationException(
                        "Every imported replay case requires a stable name.");
                if (!replayCase.HasIndependentLabel)
                {
                    throw new InvalidOperationException(
                        $"Replay case '{replayCase.Name}' requires a non-telemetry labelSource such as 'manual:reviewed'.");
                }
                _ = ReplayCase(replayCase);
            }

            var destination = EffectiveLoopReplayCorpusPath();
            lock (RecoveryMemoryService.RecoveryFileGate)
            {
                using var mutex =
                    RecoveryMemoryService.CreateRecoveryFileMutex(
                        destination);
                var lockTaken =
                    RecoveryMemoryService.WaitForRecoveryMutex(mutex);
                if (!lockTaken)
                    throw new IOException(
                        "Could not acquire the loop replay corpus lock.");
                try
                {
                    var existing = File.Exists(destination)
                        ? JsonSerializer.Deserialize<LoopReplayCorpus>(
                              File.ReadAllText(destination),
                              PrettyJson)
                          ?? new LoopReplayCorpus()
                        : new LoopReplayCorpus();
                    var importedNames = imported.Cases
                        .Select(item => item.Name)
                        .ToHashSet(StringComparer.Ordinal);
                    var merged = existing.Cases
                        .Where(item => !importedNames.Contains(item.Name))
                        .Concat(imported.Cases)
                        .ToList();
                    var output = new LoopReplayCorpus { Cases = merged };
                    var directory = Path.GetDirectoryName(destination);
                    if (!string.IsNullOrWhiteSpace(directory))
                        Directory.CreateDirectory(directory);
                    var temporary =
                        $"{destination}.tmp.{Environment.ProcessId}.{Guid.NewGuid():N}";
                    try
                    {
                        File.WriteAllText(
                            temporary,
                            JsonSerializer.Serialize(output, PrettyJson),
                            Encoding.UTF8);
                        File.Move(
                            temporary,
                            destination,
                            overwrite: true);
                    }
                    finally
                    {
                        if (File.Exists(temporary))
                            File.Delete(temporary);
                    }
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }

            Console.WriteLine(
                $"[loop-replay] imported {imported.Cases.Count} independently labelled case(s) into {destination}");
        }

        internal static (bool Detected, int Step, double Confidence) ReplayCase(
            LoopReplayCase replayCase)
        {
            if (replayCase.Frames.Count < 2)
                throw new InvalidOperationException(
                    $"Replay case '{replayCase.Name}' needs at least two frames.");

            SetCurrentScreenMap(
                0,
                0,
                Math.Max(1, replayCase.ScreenWidth),
                Math.Max(1, replayCase.ScreenHeight),
                Math.Max(1, replayCase.ScreenWidth),
                Math.Max(1, replayCase.ScreenHeight));
            var graph = new LoopStateGraph();
            var recentActions = new Queue<ResolvedActionSnapshot>();
            var highestConfidence = 0.0;
            var highestConfidenceStep = 0;
            for (var index = 0; index < replayCase.Frames.Count; index++)
            {
                var frame = replayCase.Frames[index];
                var screen = DecodeReplayFingerprint(
                    frame.ScreenFingerprintBase64,
                    replayCase.Name,
                    index,
                    "screen");
                var active = string.IsNullOrWhiteSpace(
                    frame.ActiveWindowFingerprintBase64)
                    ? screen.ToArray()
                    : DecodeReplayFingerprint(
                        frame.ActiveWindowFingerprintBase64,
                        replayCase.Name,
                        index,
                        "active-window");
                ResolvedActionSnapshot? previousAction = null;
                if (frame.PreviousAction is not null)
                {
                    previousAction = CaptureResolvedAction(
                        frame.PreviousAction,
                        null);
                    if (!previousAction.IsValid)
                    {
                        throw new InvalidOperationException(
                            $"Replay case '{replayCase.Name}', frame {index + 1}: " +
                            previousAction.ValidationError);
                    }
                    recentActions.Enqueue(previousAction);
                    while (recentActions.Count > 24)
                        recentActions.Dequeue();
                }

                var context = new UiPromptContext(
                    frame.WindowTitle,
                    frame.ActiveProcess,
                    frame.FocusSummary,
                    null,
                    null,
                    null);
                var assessment = AssessVisualStateCycle(
                    graph,
                    screen,
                    active,
                    context,
                    index + 1,
                    recentActions,
                    previousAction,
                    frame.LastDelta ?? double.NaN,
                    recordLearning: false,
                    goalMode: replayCase.GoalMode,
                    recurringWorkflowIntent:
                        replayCase.RecurringWorkflowIntent);
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

        static byte[] DecodeReplayFingerprint(
            string encoded,
            string caseName,
            int frameIndex,
            string kind)
        {
            try
            {
                var fingerprint = Convert.FromBase64String(encoded);
                if (fingerprint.Length == 0)
                    throw new FormatException("fingerprint is empty");
                return fingerprint;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Replay case '{caseName}', frame {frameIndex + 1} has an invalid {kind} fingerprint.",
                    ex);
            }
        }
    }
}
