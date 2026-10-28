namespace RDPilot.Persistence;

/// <summary>Builds redacted telemetry envelopes and derives replay cases.</summary>
internal static partial class RecoveryMemoryService
{
    static void AppendLoopTelemetry(
        string eventName,
        int step,
        string loopKind,
        string loopTopology,
        string interactionDomain,
        LoopDetectionAssessment assessment,
        bool? confirmed,
        object? details = null)
    {
        if (!Options.RecoveryMemoryEnabled)
            return;
        try
        {
            var paths = RecoveryStoragePaths.ResolveValidated(Options);
            var payload = JsonSerializer.Serialize(new
            {
                timestampUtc = RuntimeClock.UtcNow,
                @event = eventName,
                runId = assessment.RunId,
                step,
                loopKind,
                loopTopology,
                interactionDomain,
                assessment.Confidence,
                assessment.DecisionThreshold,
                assessment.CycleLength,
                assessment.MatchingPriorStates,
                assessment.GraphCycle,
                assessment.RepeatedActionCycle,
                assessment.ConsistentReturnPeriod,
                assessment.SemanticCycle,
                assessment.IsProductiveCycle,
                assessment.CycleDisposition,
                assessment.IndependentlyConfirmed,
                assessment.CalibrationKey,
                assessment.Evidence,
                confirmed,
                details
            });
            LoopTelemetryStore.Append(paths, payload);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[memory] could not append loop telemetry: {ex.Message}");
        }
    }

    internal static void AppendLoopReplayObservation(
        string runId,
        int step,
        int screenWidth,
        int screenHeight,
        byte[] screenFingerprint,
        byte[] activeWindowFingerprint,
        UiPromptContext context,
        ResolvedActionSnapshot? previousAction,
        double lastDelta,
        LoopDetectionAssessment assessment,
        ObservationAssessment? observationAssessment,
        string goalMode,
        bool recurringWorkflowIntent,
        int stateEpoch = 0,
        IReadOnlyCollection<ResolvedActionSnapshot>? recentActions = null)
    {
        if (!Options.RecoveryMemoryEnabled)
            return;
        try
        {
            var paths = RecoveryStoragePaths.ResolveValidated(Options);
            var history = recentActions?.TakeLast(RecentActionLimit).ToArray();
            var previousIndex = history is null ? -1 : Array.FindLastIndex(history, item => ReferenceEquals(item, previousAction));
            var replayFrame = new LoopReplayFrame
            {
                ReplayStateVersion = history is null ? 1 : 2,
                ContextFeatures = LoopSemanticClassifier.Context(context),
                RecentActions = history?.Select(item => new ReplayExecutedAction
                {
                    Action = ReplaySafeAction(item)!,
                    Geometry = ReplayActionResolver.Record(item)
                }).ToList(),
                PreviousActionIndex = previousIndex < 0 ? null : previousIndex,
                StateEpoch = stateEpoch,
                ScreenFingerprintBase64 =
                    Convert.ToBase64String(screenFingerprint),
                ActiveWindowFingerprintBase64 =
                    Convert.ToBase64String(activeWindowFingerprint),
                ActiveProcess = NormalizeText(context.ActiveProcessName),
                WindowTitle = ReplaySafeTokens(context.ActiveWindowTitle),
                FocusSummary = ReplaySafeTokens(context.FocusedUiaSummary),
                PreviousAction = ReplaySafeAction(previousAction),
                PreviousActionGeometry = ReplayActionResolver.Record(previousAction),
                LastDelta = double.IsFinite(lastDelta)
                    ? lastDelta
                    : null,
                ObservationPolicyVersion = observationAssessment is null ? 0 : 1,
                ObservationProfile = observationAssessment?.Profile ?? "",
                ObservationActionPolicy = observationAssessment?.ActionPolicy ?? "",
                VisualChange = observationAssessment?.VisualChange.ToString() ?? "",
                ActionOutcome = observationAssessment?.ActionOutcome.ToString() ?? "",
                GoalProgress = observationAssessment?.GoalProgress.ToString() ?? "",
                SemanticStateChanged = observationAssessment?.SemanticStateChanged,
                ObservationConfidence = observationAssessment?.Confidence,
                LocalDelta = observationAssessment is not null &&
                             double.IsFinite(observationAssessment.LocalDelta)
                    ? observationAssessment.LocalDelta
                    : null,
                LocalChangedRatio = observationAssessment?.LocalChangedRatio,
                ObservationThreshold = observationAssessment?.ChangeThreshold
            };
            var payload = JsonSerializer.Serialize(new
            {
                timestampUtc = RuntimeClock.UtcNow,
                @event = "observation",
                runId,
                step,
                screenWidth,
                screenHeight,
                assessment.Confidence,
                assessment.IndependentlyConfirmed,
                observationProfile = observationAssessment?.Profile,
                visualChange = observationAssessment?.VisualChange.ToString(),
                actionOutcome = observationAssessment?.ActionOutcome.ToString(),
                goalProgress = observationAssessment?.GoalProgress.ToString(),
                goalMode,
                recurringWorkflowIntent,
                replayFrame
            });
            LoopTelemetryStore.Append(paths, payload);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[memory] could not append replay observation: {ex.Message}");
        }
    }

    static string ReplaySafeTokens(string? value)
    {
        var tokens = Regex.Matches(
                NormalizeText(value),
                @"[\p{L}\p{Nd}]{2,}")
            .Cast<Match>()
            .Select(match => match.Value)
            .Distinct(StringComparer.Ordinal)
            .Take(24)
            .Select(token =>
            {
                var hash = SHA256.HashData(Encoding.UTF8.GetBytes(token));
                return Convert.ToHexString(hash.AsSpan(0, 5))
                    .ToLowerInvariant();
            });
        return string.Join(' ', tokens);
    }

    static ActionDto? ReplaySafeAction(
        ResolvedActionSnapshot? snapshot)
    {
        if (snapshot is null)
            return null;

        var source = snapshot.Action;
        return new ActionDto
        {
            Type = source.Type,
            X = source.X,
            Y = source.Y,
            XPx = source.XPx,
            YPx = source.YPx,
            ToX = source.ToX,
            ToY = source.ToY,
            ToXPx = source.ToXPx,
            ToYPx = source.ToYPx,
            Button = source.Button,
            Keys = source.Keys?
                .Select(key => source.Type == "hold_keys"
                    ? ReplaySafeHeldKey(key)
                    : ReplaySafeKey(key))
                .ToArray(),
            Path = ReplaySafePath(source.Path),
            GestureKind = source.GestureKind,
            DurationMs = source.DurationMs,
            Text = source.Text is null ? null : "<redacted>",
            UiaIndex = source.UiaIndex,
            ScrollDy = source.ScrollDy,
            BBox = source.BBox,
            ToBBox = source.ToBBox,
            Crop = source.Crop,
            DragDurationMs = source.DragDurationMs,
            WaitSeconds = source.WaitSeconds,
            Note = ReplaySafeTokens(snapshot.SemanticTokens)
        };
    }

    static GesturePointDto[]? ReplaySafePath(GesturePointDto[]? path)
    {
        if (path is null)
            return null;
        if (path.Length == 0)
            return [];

        const int maximumStoredPoints = 32;
        var indices = path.Length <= maximumStoredPoints
            ? Enumerable.Range(0, path.Length)
            : Enumerable.Range(0, maximumStoredPoints)
                .Select(index => (int)Math.Round(
                    index * (path.Length - 1) /
                    (double)(maximumStoredPoints - 1)));
        return indices
            .Distinct()
            .Select(index => new GesturePointDto
            {
                XPx = path[index].XPx / 4 * 4,
                YPx = path[index].YPx / 4 * 4
            })
            .ToArray();
    }

    static string ReplaySafeKey(string? key)
    {
        var normalized = NormalizeText(key);
        return Regex.IsMatch(
            normalized,
            @"^(ctrl|control|alt|shift|win|enter|escape|esc|tab|space|backspace|delete|del|insert|ins|home|end|pageup|pagedown|pgup|pgdn|up|down|left|right|arrowup|arrowdown|arrowleft|arrowright|f(?:[1-9]|1[0-2]))$",
            RegexOptions.CultureInvariant)
            ? normalized
            : "<text-key>";
    }

    static string ReplaySafeHeldKey(string? key)
    {
        var safe = ReplaySafeKey(key);
        return safe is "win" or "<text-key>"
            ? "space"
            : safe;
    }

    internal static void TryAutoExportLoopReplayCorpus()
    {
        try
        {
            ExportLoopTelemetryToReplayCorpus(
                EffectiveLoopReplayCorpusPath(),
                quiet: true);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[memory] could not auto-export loop replay corpus: {ex.Message}");
        }
    }

    internal static void ExportLoopTelemetryToReplayCorpus(string destinationPath, bool quiet = false) =>
        LoopReplayCorpusStore.Export(destinationPath, quiet);

    internal static LoopReplayCorpus BuildLoopReplayCorpus(
        IEnumerable<string> telemetryLines,
        LoopReplayCorpus? existing = null)
    {
        var envelopes = new List<LoopTelemetryReplayEnvelope>();
        foreach (var line in telemetryLines)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;
            try
            {
                var envelope =
                    JsonSerializer.Deserialize<LoopTelemetryReplayEnvelope>(
                        line,
                        PrettyJson);
                if (envelope is not null &&
                    !string.IsNullOrWhiteSpace(envelope.RunId))
                {
                    envelopes.Add(envelope);
                }
            }
            catch (JsonException)
            {
                // A partially written or older line must not prevent export
                // of the remaining telemetry.
            }
        }

        var existingCases = existing?.Cases ?? [];
        var manualCases = existingCases
            .Where(replayCase =>
                string.IsNullOrWhiteSpace(replayCase.LabelSource) ||
                !replayCase.LabelSource.StartsWith(
                    "telemetry:",
                    StringComparison.Ordinal))
            .ToList();
        var previousGeneratedCases = existingCases
            .Where(replayCase =>
                replayCase.LabelSource?.StartsWith(
                    "telemetry:",
                    StringComparison.Ordinal) == true)
            .ToList();
        var generatedCases = new List<LoopReplayCase>();
        foreach (var run in envelopes.GroupBy(item => item.RunId))
        {
            var observations = run
                .Where(item =>
                    item.Event == "observation" &&
                    item.ReplayFrame is not null)
                .GroupBy(item => item.Step)
                .Select(group => group
                    .OrderByDescending(item => item.TimestampUtc)
                    .First())
                .OrderBy(item => item.Step)
                .ToList();
            if (observations.Count < 2)
                continue;

            var firstPositiveStep = run
                .Where(item =>
                    item.IndependentlyConfirmed ||
                    item.Confirmed == true)
                .Select(item => item.Step)
                .Where(step => step > 0)
                .DefaultIfEmpty(0)
                .Min();
            var expectedLoop = firstPositiveStep > 0;
            var quietNegative =
                observations.Count >= 6 &&
                observations.All(item => item.Confidence < 0.5);
            if (!expectedLoop && !quietNegative)
                continue;

            var endStep = expectedLoop
                ? firstPositiveStep
                : observations[^1].Step;
            var selected = observations
                .Where(item => item.Step <= endStep)
                .TakeLast(24)
                .ToList();
            if (selected.Count < 2)
                continue;

            var capturedUtc = selected
                .Select(item => item.TimestampUtc)
                .Where(value => value != default)
                .DefaultIfEmpty(RuntimeClock.UtcNow)
                .Max();
            generatedCases.Add(new LoopReplayCase
            {
                Name =
                    $"telemetry:{run.Key}:{(expectedLoop ? "positive" : "negative")}",
                ExpectedLoop = expectedLoop,
                LabelSource = expectedLoop
                    ? "telemetry:independent_confirmation"
                    : "telemetry:quiet_run",
                CapturedUtc = capturedUtc,
                GoalMode = string.IsNullOrWhiteSpace(
                    selected[^1].GoalMode)
                    ? "finite"
                    : selected[^1].GoalMode,
                RecurringWorkflowIntent =
                    selected[^1].RecurringWorkflowIntent,
                ScreenWidth = Math.Max(1, selected[^1].ScreenWidth),
                ScreenHeight = Math.Max(1, selected[^1].ScreenHeight),
                Frames = selected
                    .Select(item => item.ReplayFrame!)
                    .ToList()
            });
        }

        return new LoopReplayCorpus
        {
            Cases = manualCases
                .Concat(previousGeneratedCases
                    .Concat(generatedCases)
                    .GroupBy(
                        replayCase =>
                        {
                            var separator =
                                replayCase.Name.LastIndexOf(':');
                            return separator > 0
                                ? replayCase.Name[..separator]
                                : replayCase.Name;
                        },
                        StringComparer.Ordinal)
                    .Select(group => group
                        .OrderByDescending(item => item.ExpectedLoop)
                        .ThenByDescending(item => item.CapturedUtc)
                        .First())
                    .OrderByDescending(item => item.CapturedUtc)
                    .Take(200))
                .ToList()
        };
    }

}
