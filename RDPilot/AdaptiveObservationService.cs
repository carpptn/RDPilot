internal static partial class RDPilotApplication
{
    /// <summary>
    /// Selects an observation profile for the current action and separates visual
    /// motion from action success and goal progress.
    /// </summary>
    internal sealed class AdaptiveObservationSession
    {
        const int NoiseHistoryLimit = 15;

        readonly Queue<double> stableNoiseSamples = new();
        readonly Queue<double> ambientMotionSamples = new();
        string effectiveProfile = "general";
        string? pendingProfile;
        int pendingProfileSamples;
        string contextProcess = "";

        internal string EffectiveProfile => effectiveProfile;

        internal void PrepareForPrompt(
            UiPromptContext context,
            string? goal = null)
        {
            var (candidate, confidence, reason, immediate) = SelectProfile(
                null,
                context,
                Math.Max(StableNoiseBaseline(), AmbientMotionBaseline()),
                goal);
            ConsiderProfile(candidate, confidence, reason, immediate);
        }

        internal string ResolveActionPolicy(
            ResolvedActionSnapshot action,
            UiPromptContext context,
            string? goal = null)
        {
            if (!string.Equals(ObservationMode, "auto", StringComparison.OrdinalIgnoreCase))
                return effectiveProfile;
            return SelectProfile(
                action,
                context,
                Math.Max(StableNoiseBaseline(), AmbientMotionBaseline()),
                goal).Profile;
        }

        internal AdaptiveObservationSession()
        {
            if (!string.Equals(ObservationMode, "auto", StringComparison.OrdinalIgnoreCase))
                effectiveProfile = NormalizeProfile(ObservationMode);
        }

        internal void RecordAmbientMotion(
            ScreenObservationFrame earlier,
            ScreenObservationFrame immediatelyBeforeAction)
        {
            var active = ComputeImageDelta(
                earlier.ActiveWindowFingerprint,
                immediatelyBeforeAction.ActiveWindowFingerprint);
            var global = ComputeImageDelta(
                earlier.GlobalFingerprint,
                immediatelyBeforeAction.GlobalFingerprint);
            var observed = 0.8 * active + 0.2 * global;
            if (!double.IsFinite(observed))
                return;
            ambientMotionSamples.Enqueue(Math.Clamp(observed, 0, 1));
            while (ambientMotionSamples.Count > NoiseHistoryLimit)
                ambientMotionSamples.Dequeue();
        }

        internal void LogInitialProfile()
        {
            Console.WriteLine(
                $"[observation] mode={ObservationMode}; initial profile={effectiveProfile}");
        }

        internal ObservationAssessment Assess(
            ScreenObservationFrame before,
            ScreenObservationFrame after,
            ResolvedActionSnapshot? action,
            UiPromptContext? beforeContext,
            UiPromptContext context,
            string goalMode,
            string? goal = null)
        {
            var globalDelta = ComputeImageDelta(
                before.GlobalFingerprint,
                after.GlobalFingerprint);
            var activeDelta = ComputeImageDelta(
                before.ActiveWindowFingerprint,
                after.ActiveWindowFingerprint);
            var (localDelta, localRatio) = ComputeLocalMetrics(
                before,
                after,
                action);

            var process = NormalizeObservationText(context.ActiveProcessName);
            var semanticChanged = HasSemanticStateChange(
                beforeContext,
                context,
                action?.Action);
            if (!string.Equals(contextProcess, process, StringComparison.Ordinal))
            {
                if (!string.IsNullOrWhiteSpace(contextProcess))
                {
                    SwitchProfile("general", 1, "foreground process changed");
                    stableNoiseSamples.Clear();
                    ambientMotionSamples.Clear();
                }
                contextProcess = process;
                pendingProfile = null;
                pendingProfileSamples = 0;
            }

            var noise = Math.Max(StableNoiseBaseline(), AmbientMotionBaseline());
            var (candidate, candidateConfidence, reason, immediate) = SelectProfile(
                action,
                context,
                noise,
                goal);
            ConsiderProfile(candidate, candidateConfidence, reason, immediate);

            var actionPolicy = string.Equals(
                    ObservationMode,
                    "auto",
                    StringComparison.OrdinalIgnoreCase) &&
                candidateConfidence >= 0.7
                    ? candidate
                    : effectiveProfile;
            var threshold = ThresholdFor(actionPolicy, noise);
            var effectiveDelta = 0.8 * activeDelta + 0.2 * globalDelta;
            var hasLocalEvidence = double.IsFinite(localDelta);
            var localChanged = hasLocalEvidence &&
                               (localDelta >= threshold ||
                                localRatio >= LocalRatioThreshold(actionPolicy));
            var broadChanged = effectiveDelta >= threshold;
            var expectsLocal = action?.ObservationRegion is not null &&
                               ((actionPolicy == "local_editing" &&
                                 action.Action.Type is ("drag_drop" or "drag_path" or "type_text" or "paste_text")) ||
                                (actionPolicy == "turn_based_interaction" &&
                                 ControlLoopService.IsStateChangingInteractionAction(action.Action)));

            VisualChangeState visual;
            if (expectsLocal && localChanged)
                visual = VisualChangeState.Changed;
            else if (expectsLocal && broadChanged && !localChanged)
                visual = VisualChangeState.Ambiguous;
            else if (broadChanged)
                visual = actionPolicy is "streaming_output" or "realtime_interaction" &&
                         effectiveDelta <= Math.Max(threshold, noise * 1.35)
                    ? VisualChangeState.Unstable
                    : VisualChangeState.Changed;
            else
                visual = VisualChangeState.Stable;

            var outcome = AssessActionOutcome(
                action?.Action,
                visual,
                localChanged,
                broadChanged,
                semanticChanged);
            var progress = AssessGoalProgress(
                action?.Action,
                outcome,
                goalMode,
                actionPolicy);
            var confidence = AssessmentConfidence(
                visual,
                effectiveDelta,
                localDelta,
                threshold,
                candidateConfidence);
            var evidence = BuildEvidence(
                $"{effectiveProfile}; policy={actionPolicy}",
                globalDelta,
                activeDelta,
                localDelta,
                localRatio,
                threshold,
                noise,
                visual,
                outcome,
                progress) + $"; semantic_changed={semanticChanged.ToString().ToLowerInvariant()}";

            LearnStableNoise(action?.Action, visual, effectiveDelta);

            if (ObservationLogVerbose)
                Console.WriteLine($"[observation] {evidence}");

            return new ObservationAssessment(
                effectiveProfile,
                visual,
                outcome,
                progress,
                confidence,
                expectsLocal && hasLocalEvidence ? localDelta : effectiveDelta,
                globalDelta,
                activeDelta,
                localDelta,
                localRatio,
                threshold,
                evidence)
            {
                SemanticStateChanged = semanticChanged,
                ActionPolicy = actionPolicy
            };
        }

        void ConsiderProfile(
            string candidate,
            double confidence,
            string reason,
            bool immediate)
        {
            if (!string.Equals(ObservationMode, "auto", StringComparison.OrdinalIgnoreCase))
            {
                SwitchProfile(NormalizeProfile(ObservationMode), 1, "manually selected");
                return;
            }

            if (string.Equals(candidate, effectiveProfile, StringComparison.Ordinal))
            {
                pendingProfile = null;
                pendingProfileSamples = 0;
                return;
            }

            if (immediate && confidence >= 0.8)
            {
                SwitchProfile(candidate, confidence, reason);
                return;
            }

            if (!string.Equals(candidate, pendingProfile, StringComparison.Ordinal))
            {
                pendingProfile = candidate;
                pendingProfileSamples = 1;
                return;
            }

            pendingProfileSamples++;
            if (pendingProfileSamples >= 2 && confidence >= 0.65)
                SwitchProfile(candidate, confidence, reason);
        }

        void SwitchProfile(string profile, double confidence, string reason)
        {
            profile = NormalizeProfile(profile);
            if (string.Equals(profile, effectiveProfile, StringComparison.Ordinal))
                return;

            var previous = effectiveProfile;
            effectiveProfile = profile;
            pendingProfile = null;
            pendingProfileSamples = 0;
            Console.WriteLine(
                $"[observation] profile {previous} -> {profile}; confidence={confidence:0.00}; reason={reason}");
        }

        (string Profile, double Confidence, string Reason, bool Immediate) SelectProfile(
            ResolvedActionSnapshot? snapshot,
            UiPromptContext context,
            double noise,
            string? goal = null)
        {
            var action = snapshot?.Action;
            var type = NormalizeObservationText(action?.Type);
            var gesture = NormalizeObservationText(action?.GestureKind);
            var process = NormalizeObservationText(context.ActiveProcessName);
            var semantic = NormalizeObservationText(
                $"{action?.Note} {snapshot?.SemanticTokens} {context.FocusedUiaSummary}");

            if (type == "drag_path")
            {
                if (gesture is "pan" or "game" or "realtime" ||
                    Regex.IsMatch(semantic, @"\b(game|camera|pan|joystick|steer|gra|kamera)\b"))
                {
                    return ("realtime_interaction", 0.94, "explicit realtime path gesture", true);
                }
                return ("local_editing", 0.96, "path gesture with a local visual footprint", true);
            }

            if (type == "hold_keys")
                return ("realtime_interaction", 0.98, "bounded key hold", true);

            if (type == "drag_drop")
                return ("local_editing", 0.86, "source/destination gesture", true);

            if (type is "type_text" or "paste_text" && snapshot?.ObservationRegion is not null)
                return ("local_editing", 0.9, "focused text-field edit", true);

            if (IsTurnBasedInteractionContext(
                    goal,
                    context.ActiveWindowTitle,
                    context.FocusedUiaSummary,
                    semantic))
            {
                return ("turn_based_interaction", 0.92, "puzzle or board interface with discrete controls", true);
            }

            if (action is null)
                return ("general", 0.7, "no executed action", false);

            if (noise >= 0.012 && HasPersistentHighAmbientMotion() &&
                IsRealtimeAmbientContext(type, process, semantic, context.ActiveWindowTitle))
            {
                return ("realtime_interaction", 0.88, "persistent motion in a realtime interaction context", true);
            }

            var terminal = IsObservationTerminalProcess(process);
            if (terminal && type is "keys" or "run_command")
            {
                return noise >= NoChangeThreshold
                    ? ("streaming_output", 0.86, "terminal output remains in motion", true)
                    : ("event_driven", 0.84, "terminal input expects output activity", true);
            }

            if (type is "open_url" or "launch_app" or "wait" or "run_command")
                return ("event_driven", 0.84, "asynchronous action", true);

            if (type is "click" or "double_click" or "click_uia" &&
                Regex.IsMatch(
                    semantic,
                    @"\b(submit|send|save|open|navigate|load|upload|download|wyślij|wyslij|zapisz|otwórz|otworz|przejdź|przejdz)\b"))
            {
                return ("event_driven", 0.82, "action intent expects an asynchronous transition", true);
            }

            if (type is "click" or "double_click" or "focus_uia" or "click_uia" or
                "type_text" or "paste_text" or "keys" or "scroll")
            {
                return noise >= 0.0045
                    ? ("static_ui", 0.84, "discrete UI action retains settle despite foreground motion", false)
                    : ("static_ui", 0.76, "discrete UI action on a stable foreground", false);
            }

            if (noise >= 0.0045)
                return ("streaming_output", 0.72, "persistent foreground motion", false);

            return ("general", 0.65, "no specialized policy matched", false);
        }

        bool HasPersistentHighAmbientMotion() =>
            ambientMotionSamples.Count >= 2 &&
            ambientMotionSamples.TakeLast(2).All(sample => sample >= 0.012);

        static bool IsRealtimeAmbientContext(
            string actionType,
            string process,
            string semantic,
            string windowTitle)
        {
            if (actionType is not ("keys" or "scroll"))
                return false;

            var context = NormalizeObservationText($"{process} {windowTitle} {semantic}");
            return Regex.IsMatch(
                context,
                @"\b(game|gaming|joystick|steer|camera control|realtime|real-time|gra|kamera|canvas animation|animated canvas)\b");
        }

        internal static bool IsTurnBasedInteractionContext(
            string? goal,
            string? windowTitle,
            string? focusedUiaSummary,
            string? actionSemantic)
        {
            var intent = NormalizeObservationText($"{goal} {actionSemantic}");
            var visibleContext = NormalizeObservationText(
                $"{windowTitle} {focusedUiaSummary} {actionSemantic}");
            if (Regex.IsMatch(
                    intent + " " + visibleContext,
                    @"\b(realtime|real-time|racing|race|shooter|fps|steer|wyścig|wyscig|strzelanka)\b"))
            {
                return false;
            }

            var puzzleIntent = Regex.IsMatch(
                intent,
                @"\b(puzzle|logic|logical|turn-based|turn based|grid-based|grid based|board puzzle|zagad\w*|logicz\w*|turow\w*|planszow\w*|arc-agi)\b");
            var interactiveSurface = Regex.IsMatch(
                visibleContext,
                @"\b(game|puzzle|board|grid|level|task|control|d-pad|arc-agi|gra|zagad\w*|plansz\w*|poziom\w*|sterow\w*)\b");
            return puzzleIntent && interactiveSurface;
        }

        static ActionOutcomeState AssessActionOutcome(
            ActionDto? action,
            VisualChangeState visual,
            bool localChanged,
            bool broadChanged,
            bool semanticChanged)
        {
            if (action is null || IsLocalObservationAction(action) || action.Type == "move")
                return ActionOutcomeState.NotObserved;

            if (semanticChanged)
                return ActionOutcomeState.Confirmed;

            if (visual == VisualChangeState.Ambiguous)
                return broadChanged && !localChanged
                    ? ActionOutcomeState.UnexpectedChange
                    : ActionOutcomeState.Ambiguous;
            if (visual == VisualChangeState.Unstable)
                return ActionOutcomeState.Ambiguous;
            if (visual == VisualChangeState.Stable)
                return ActionOutcomeState.NoEffect;
            return ActionOutcomeState.Confirmed;
        }

        static GoalProgressState AssessGoalProgress(
            ActionDto? action,
            ActionOutcomeState outcome,
            string goalMode,
            string actionPolicy)
        {
            if (action is null || IsLocalObservationAction(action) || action.Type == "move")
                return GoalProgressState.Neutral;
            if (actionPolicy == "turn_based_interaction")
            {
                return outcome switch
                {
                    ActionOutcomeState.UnexpectedChange or
                    ActionOutcomeState.Ambiguous => GoalProgressState.Unknown,
                    _ => GoalProgressState.Neutral
                };
            }
            if (action.Type == "wait" &&
                string.Equals(goalMode, "continuous", StringComparison.OrdinalIgnoreCase) &&
                outcome == ActionOutcomeState.NoEffect)
            {
                return GoalProgressState.Neutral;
            }
            return outcome switch
            {
                ActionOutcomeState.Confirmed => GoalProgressState.Progress,
                ActionOutcomeState.NoEffect => GoalProgressState.NoProgress,
                ActionOutcomeState.UnexpectedChange => GoalProgressState.Unknown,
                ActionOutcomeState.Ambiguous => GoalProgressState.Unknown,
                _ => GoalProgressState.Neutral
            };
        }

        void LearnStableNoise(ActionDto? action, VisualChangeState visual, double delta)
        {
            if (visual != VisualChangeState.Stable ||
                action is not null && action.Type != "wait" && !IsLocalObservationAction(action) ||
                !double.IsFinite(delta) || delta > 0.02)
            {
                return;
            }

            stableNoiseSamples.Enqueue(Math.Max(0, delta));
            while (stableNoiseSamples.Count > NoiseHistoryLimit)
                stableNoiseSamples.Dequeue();
        }

        double StableNoiseBaseline()
        {
            if (stableNoiseSamples.Count == 0)
                return 0;
            var ordered = stableNoiseSamples.OrderBy(value => value).ToArray();
            return ordered[ordered.Length / 2];
        }

        double AmbientMotionBaseline()
        {
            if (ambientMotionSamples.Count == 0)
                return 0;
            var ordered = ambientMotionSamples
                .TakeLast(5)
                .OrderBy(value => value)
                .ToArray();
            return ordered[ordered.Length / 2];
        }

        static double ThresholdFor(string profile, double noise) => profile switch
        {
            "static_ui" => Math.Clamp(noise * 3 + 0.0025, 0.0025, 0.012),
            "local_editing" => Math.Clamp(noise * 2 + 0.0012, 0.0012, 0.008),
            "event_driven" => Math.Clamp(noise * 3 + 0.0035, 0.0035, 0.018),
            "turn_based_interaction" => Math.Clamp(noise * 2.5 + 0.0035, 0.004, 0.025),
            "streaming_output" => Math.Clamp(noise * 2.5 + 0.004, 0.006, 0.03),
            "realtime_interaction" => Math.Clamp(noise * 2 + 0.006, 0.008, 0.04),
            _ => Math.Clamp(noise * 3 + NoChangeThreshold, NoChangeThreshold, 0.02)
        };

        static double LocalRatioThreshold(string profile) => profile switch
        {
            "local_editing" => 0.003,
            "turn_based_interaction" => 0.0015,
            _ => 0.01
        };

        static (double Delta, double ChangedRatio) ComputeLocalMetrics(
            ScreenObservationFrame before,
            ScreenObservationFrame after,
            ResolvedActionSnapshot? snapshot)
        {
            var screenRegion = snapshot?.ObservationRegion;
            if (screenRegion is not Rectangle region ||
                before.DetailWidth <= 0 ||
                before.DetailHeight <= 0 ||
                before.DetailWidth != after.DetailWidth ||
                before.DetailHeight != after.DetailHeight ||
                before.DetailFingerprint.Length != after.DetailFingerprint.Length)
            {
                return (double.NaN, 0);
            }

            var bounds = before.ScreenBounds;
            region.Intersect(bounds);
            if (region.Width <= 0 || region.Height <= 0)
                return (double.NaN, 0);

            var left = Math.Clamp(
                (int)Math.Floor((region.Left - bounds.Left) * before.DetailWidth / (double)Math.Max(1, bounds.Width)),
                0,
                before.DetailWidth - 1);
            var top = Math.Clamp(
                (int)Math.Floor((region.Top - bounds.Top) * before.DetailHeight / (double)Math.Max(1, bounds.Height)),
                0,
                before.DetailHeight - 1);
            var right = Math.Clamp(
                (int)Math.Ceiling((region.Right - bounds.Left) * before.DetailWidth / (double)Math.Max(1, bounds.Width)),
                left + 1,
                before.DetailWidth);
            var bottom = Math.Clamp(
                (int)Math.Ceiling((region.Bottom - bounds.Top) * before.DetailHeight / (double)Math.Max(1, bounds.Height)),
                top + 1,
                before.DetailHeight);

            double sum = 0;
            var changed = 0;
            var count = 0;
            var useColor = before.DetailColorFingerprint.Length ==
                               before.DetailWidth * before.DetailHeight * 3 &&
                           after.DetailColorFingerprint.Length ==
                               after.DetailWidth * after.DetailHeight * 3;
            for (var y = top; y < bottom; y++)
            {
                var row = y * before.DetailWidth;
                for (var x = left; x < right; x++)
                {
                    if (!IsInsideObservationFootprint(
                            snapshot!,
                            before,
                            x,
                            y))
                    {
                        continue;
                    }
                    var pixelIndex = row + x;
                    double difference;
                    var maximumChannelDifference = 0;
                    if (useColor)
                    {
                        var colorIndex = pixelIndex * 3;
                        var red = Math.Abs(before.DetailColorFingerprint[colorIndex] - after.DetailColorFingerprint[colorIndex]);
                        var green = Math.Abs(before.DetailColorFingerprint[colorIndex + 1] - after.DetailColorFingerprint[colorIndex + 1]);
                        var blue = Math.Abs(before.DetailColorFingerprint[colorIndex + 2] - after.DetailColorFingerprint[colorIndex + 2]);
                        difference = (red + green + blue) / 3.0;
                        maximumChannelDifference = Math.Max(red, Math.Max(green, blue));
                    }
                    else
                    {
                        maximumChannelDifference = Math.Abs(
                            before.DetailFingerprint[pixelIndex] -
                            after.DetailFingerprint[pixelIndex]);
                        difference = maximumChannelDifference;
                    }
                    sum += difference / 255.0;
                    if (maximumChannelDifference >= 8)
                        changed++;
                    count++;
                }
            }

            return count == 0
                ? (double.NaN, 0)
                : (sum / count, changed / (double)count);
        }

        static bool IsInsideObservationFootprint(
            ResolvedActionSnapshot snapshot,
            ScreenObservationFrame frame,
            int detailX,
            int detailY)
        {
            if (snapshot.ScreenPath.Count == 0)
                return true;

            var bounds = frame.ScreenBounds;
            var screenX = bounds.Left +
                          (detailX + 0.5) * bounds.Width /
                          Math.Max(1, frame.DetailWidth);
            var screenY = bounds.Top +
                          (detailY + 0.5) * bounds.Height /
                          Math.Max(1, frame.DetailHeight);

            if (snapshot.Action.Type == "drag_drop")
            {
                const double endpointRadius = 64;
                return DistanceSquared(screenX, screenY, snapshot.ScreenPath[0]) <=
                           endpointRadius * endpointRadius ||
                       DistanceSquared(screenX, screenY, snapshot.ScreenPath[^1]) <=
                           endpointRadius * endpointRadius;
            }

            const double pathRadius = 28;
            for (var index = 1; index < snapshot.ScreenPath.Count; index++)
            {
                if (DistanceToSegmentSquared(
                        screenX,
                        screenY,
                        snapshot.ScreenPath[index - 1],
                        snapshot.ScreenPath[index]) <= pathRadius * pathRadius)
                {
                    return true;
                }
            }
            return false;
        }

        static double DistanceSquared(double x, double y, Point point)
        {
            var dx = x - point.X;
            var dy = y - point.Y;
            return dx * dx + dy * dy;
        }

        static double DistanceToSegmentSquared(
            double x,
            double y,
            Point start,
            Point end)
        {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var lengthSquared = dx * (double)dx + dy * (double)dy;
            if (lengthSquared <= 0)
                return DistanceSquared(x, y, start);
            var projection = Math.Clamp(
                ((x - start.X) * dx + (y - start.Y) * dy) / lengthSquared,
                0,
                1);
            var nearestX = start.X + projection * dx;
            var nearestY = start.Y + projection * dy;
            var distanceX = x - nearestX;
            var distanceY = y - nearestY;
            return distanceX * distanceX + distanceY * distanceY;
        }

        static double AssessmentConfidence(
            VisualChangeState visual,
            double effectiveDelta,
            double localDelta,
            double threshold,
            double profileConfidence)
        {
            var measured = double.IsFinite(localDelta)
                ? Math.Max(effectiveDelta, localDelta)
                : effectiveDelta;
            var separation = Math.Abs(measured - threshold) /
                             Math.Max(0.0001, threshold);
            var evidenceConfidence = visual == VisualChangeState.Ambiguous
                ? 0.45
                : Math.Clamp(0.62 + separation * 0.12, 0.62, 0.96);
            return Math.Clamp(0.55 * evidenceConfidence + 0.45 * profileConfidence, 0, 1);
        }

        static string BuildEvidence(
            string profile,
            double global,
            double active,
            double local,
            double ratio,
            double threshold,
            double noise,
            VisualChangeState visual,
            ActionOutcomeState outcome,
            GoalProgressState progress) =>
            $"profile={profile}; visual={visual}; outcome={outcome}; progress={progress}; " +
            $"global={global:0.####}; active={active:0.####}; " +
            $"local={(double.IsFinite(local) ? local.ToString("0.####") : "n/a")}; " +
            $"local_ratio={ratio:0.####}; threshold={threshold:0.####}; noise={noise:0.####}";

        static string NormalizeProfile(string? profile) => NormalizeObservationText(profile) switch
        {
            "static" => "static_ui",
            "local" => "local_editing",
            "event" => "event_driven",
            "streaming" => "streaming_output",
            "realtime" => "realtime_interaction",
            "turn_based" => "turn_based_interaction",
            "static_ui" or "local_editing" or "event_driven" or
                "streaming_output" or "realtime_interaction" or "turn_based_interaction" => NormalizeObservationText(profile),
            _ => "general"
        };

        static bool IsObservationTerminalProcess(string process) =>
            process is "pwsh" or "powershell" or "powershell_ise" or "cmd" or
                "wt" or "windowsterminal" or "conhost";

        static bool HasSemanticStateChange(
            UiPromptContext? before,
            UiPromptContext after,
            ActionDto? action)
        {
            if (before is null || action is null ||
                action.Type is "drag_drop" or "drag_path" or "hold_keys" or
                    "move" or "wait")
            {
                return false;
            }

            return !string.Equals(
                       NormalizeObservationText(before.ActiveProcessName),
                       NormalizeObservationText(after.ActiveProcessName),
                       StringComparison.Ordinal) ||
                   !string.Equals(
                       NormalizeObservationText(before.ActiveWindowTitle),
                       NormalizeObservationText(after.ActiveWindowTitle),
                       StringComparison.Ordinal) ||
                   !string.Equals(
                       NormalizeObservationText(before.FocusedUiaSummary),
                       NormalizeObservationText(after.FocusedUiaSummary),
                       StringComparison.Ordinal);
        }

        static string NormalizeObservationText(string? value) =>
            string.IsNullOrWhiteSpace(value)
                ? ""
                : value.Trim().ToLowerInvariant();
    }
}
