
namespace RDPilot.Control;

    internal static class ControlBatchService
    {

            internal readonly record struct TurnNoEffectPolicy(
                bool ExtendObservation,
                bool ReplayInput);

            internal static bool TryTakeQueuedObservedTurnActions(
                Queue<ActionDto> followUps,
                ResolvedActionSnapshot firstAction,
                Rectangle region,
                out IReadOnlyList<ResolvedActionSnapshot> actions,
                IReadOnlyList<string>? plannedInputs = null)
            {
                plannedInputs ??= firstAction.Action.PlannedInputs;
                actions = Array.Empty<ResolvedActionSnapshot>();
                if (followUps.Count == 0 ||
                    plannedInputs is not { Count: >= 2 } ||
                    (firstAction.Action.PlanConfidence ?? firstAction.Action.Confidence ?? 0) <
                        TurnBasedTransitionTracker.MinimumStructuredPlanConfidence ||
                    !BatchPlanner.TryBindObservedTurnActionToPlannedInput(
                        firstAction.Action,
                        plannedInputs[0]))
                {
                    return false;
                }

                var queued = followUps.ToArray();
                var maximumActions = Math.Min(
                    Options.TurnBasedMaxBatchInputs,
                    plannedInputs.Count);
                var accepted = new List<ResolvedActionSnapshot>(maximumActions)
                {
                    firstAction with { ObservationRegion = ClampRect(region) }
                };
                for (var index = 0;
                     index < queued.Length && accepted.Count < maximumActions;
                     index++)
                {
                    var candidate = queued[index];
                    if (!BatchPlanner.TryBindObservedTurnActionToPlannedInput(
                            candidate,
                            plannedInputs[index + 1]))
                    {
                        break;
                    }

                    var snapshot = CaptureResolvedAction(candidate, null) with
                    {
                        ObservationRegion = ClampRect(region)
                    };
                    if (!snapshot.IsValid)
                        break;
                    accepted.Add(snapshot);
                }

                if (accepted.Count <= 1)
                    return false;

                followUps.Clear();
                actions = accepted;
                return true;
            }

            internal static async Task<BatchedActionExecutionResult> ExecuteObservedTurnActionSequenceAsync(
                IReadOnlyList<ResolvedActionSnapshot> actions,
                ScreenObservationFrame initialFrame,
                Rectangle region,
                TurnBasedTransitionTracker tracker,
                IntPtr interactionWindow,
                CancellationToken cancellationToken,
                Action<ResolvedActionSnapshot>? validateAction = null)
            {
                Console.WriteLine(
                    $"[turn-batch] mode={tracker.ExecutionBatchMode}; modality=mixed_observed; executing={actions.Count}; observation_barrier=per_input");
                var executed = new List<ActionExecutionRecord>(actions.Count);
                BatchedActionExecutionResult Result(string? error = null, bool cancelled = false, bool interrupted = false,
                    ResolvedActionSnapshot? failed = null) => new(executed, error, cancelled, interrupted, failed);
                var beforeFrame = initialFrame;
                tracker.BeginBatch(initialFrame, region);
                for (var index = 0; index < actions.Count; index++)
                {
                    var snapshot = actions[index];
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (interactionWindow != IntPtr.Zero &&
                            GetForegroundWindow() != interactionWindow &&
                            !TryRestoreForegroundWindow(interactionWindow))
                            throw new InvalidOperationException("interaction window could not regain focus");
                        if (!snapshot.IsValid) throw new InvalidOperationException(snapshot.ValidationError);
                        validateAction?.Invoke(snapshot);
                        var beforeLocal = CaptureRegionFingerprintProbe(region);
                        var beforeContext = ActionObservationCapture.Context(beforeFrame);
                        ExecuteAction(snapshot.Action);
                        snapshot.Action.TurnSequenceObserved = true;
                        executed.Add(new(snapshot, beforeFrame, beforeLocal, BeforeContext: beforeContext));
                        var actionReactionObserved = await WaitAfterActionAsync(
                            snapshot.Action,
                            beforeFrame.GlobalFingerprint,
                            "turn_based_interaction",
                            beforeLocal,
                            region,
                            cancellationToken,
                            tracker.PreferFastBatchSettle);
                        var noEffectPolicy = ResolveTurnNoEffectPolicy(
                            actionReactionObserved);
                        if (noEffectPolicy.ExtendObservation)
                        {
                            Console.WriteLine(
                                $"[turn-batch] input={index + 1}/{actions.Count}; no immediate reaction; extending observation without replaying the input.");
                            actionReactionObserved = await WaitForLocalScreenReactionAndStableAsync(
                                beforeLocal,
                                region,
                                snapshot.Action,
                                cancellationToken,
                                preferFastSettle: false);
                        }
                        var afterFrame = CaptureObservationFrameProbe(includeDetail: true);
                        var assessment = tracker.RecordBatchActionStep(
                            beforeFrame,
                            afterFrame,
                            region,
                            snapshot,
                            actionReactionObserved);
                        executed[^1] = executed[^1] with { Observation = new(snapshot, beforeFrame, afterFrame, assessment) };
                        beforeFrame = afterFrame;
                        Console.WriteLine(
                            $"[turn-batch] input={index + 1}/{actions.Count}; attempts=1; action={snapshot.Description}; known_transition={assessment.KnownTransition.ToString().ToLowerInvariant()}; salient={assessment.SalientChange.ToString().ToLowerInvariant()}; {assessment.Summary}");
                        if (assessment.NoEffect)
                        {
                            Console.WriteLine(
                                "[turn-batch] observed sequence interrupted after extended no_effect confirmation; the input was not replayed and the route cursor was not advanced.");
                            return Result(interrupted: true);
                        }
                        if (index + 1 < actions.Count && !assessment.ContinueBatch)
                        {
                            Console.WriteLine(
                                "[turn-batch] observed sequence interrupted before the next input; a fresh model decision is required.");
                            return Result(interrupted: true);
                        }
                    }
                    catch (OperationCanceledException) when (CancelRequested || cancellationToken.IsCancellationRequested)
                    {
                        return Result("observed sequence cancelled", cancelled: true, interrupted: true, failed: snapshot);
                    }
                    catch (Exception ex)
                    {
                        return Result($"{snapshot.Description} failed in observed sequence: {ex.Message}", interrupted: true, failed: snapshot);
                    }
                }

                return Result();
            }

            internal static Task<BatchedActionExecutionResult> ExecuteTurnBasedDirectionalSequenceAsync(
                IReadOnlyList<string> keys,
                ScreenObservationFrame initialFrame,
                Rectangle region,
                TurnBasedTransitionTracker tracker,
                IntPtr interactionWindow,
                CancellationToken cancellationToken,
                Action<ResolvedActionSnapshot>? validateAction = null)
            {
                var actions = keys.Select(key => CaptureResolvedAction(new ActionDto
                {
                    Type = "keys",
                    Keys = [key]
                }, null) with { ObservationRegion = region }).ToArray();
                return ExecuteObservedTurnActionSequenceAsync(
                    actions, initialFrame, region, tracker, interactionWindow, cancellationToken, validateAction);
            }

            internal static async Task<BatchedActionExecutionResult> ExecuteQueuedSafeActionsAsync(
                Queue<ActionDto> followUps,
                int step,
                ActionDto initialAction,
                byte[] initialBaselineFingerprint,
                CancellationToken cancellationToken,
                Action<ResolvedActionSnapshot>? validateAction = null)
            {
                var batchIndex = 0;
                var executed = new List<ActionExecutionRecord>();
                var previousAction = initialAction;
                var navigationLikeBatch = IsNavigationBatchPrelude(initialAction);
                byte[]? transitionBaseline = IsAdaptiveBatchWaitTrigger(initialAction)
                    ? initialBaselineFingerprint
                    : null;
                while (followUps.Count > 0)
                {
                    var action = followUps.Dequeue();
                    var snapshot = new ResolvedActionSnapshot(action, action.Type, action.Type, null);
                    batchIndex++;
        
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        snapshot = CaptureResolvedAction(action, null);
                        Console.WriteLine($"[{step}.{batchIndex}] batch {snapshot.Description}");
                        if (!snapshot.IsValid)
                            throw new InvalidOperationException(snapshot.ValidationError);

                        validateAction?.Invoke(snapshot);
                        var transitionDelayMs = BatchTransitionDelayMs(previousAction, action);
                        if (transitionDelayMs > 0)
                        {
                            Console.WriteLine(
                                $"[batch] transition {previousAction.Type}->{action.Type}; delay={transitionDelayMs}ms");
                            await RuntimeClock.Delay(transitionDelayMs, cancellationToken);
                        }

                        var beforeFrame = CaptureObservationFrameProbe(includeDetail: true);
                        var beforeLocal = snapshot.ObservationRegion is Rectangle region
                            ? CaptureRegionFingerprintProbe(region) : null;
                        var beforeContext = ActionObservationCapture.Context(beforeFrame);
                        if (action.Type == "wait")
                        {
                            var secs = EffectiveWaitSeconds(action, out var requestedSecs);
                            if (secs < requestedSecs)
                                Console.WriteLine($"[wait] Requested {requestedSecs}s capped to {secs}s.");
                            if (transitionBaseline is not null && secs > 0)
                            {
                                await WaitForBatchedTransitionAsync(
                                    transitionBaseline,
                                    previousAction,
                                    secs * 1000,
                                    navigationLikeBatch,
                                    cancellationToken);
                            }
                            else
                            {
                                Console.WriteLine($"[wait] Sleeping {secs} s (batched)...");
                                await RuntimeClock.Delay(secs * 1000, cancellationToken);
                            }
                        }
                        else
                        {
                            transitionBaseline = IsAdaptiveBatchWaitTrigger(action)
                                ? CaptureScreenFingerprintProbe()
                                : null;
                            ExecuteAction(action);
                        }
        
                        executed.Add(new(snapshot, beforeFrame, beforeLocal, BeforeContext: beforeContext));
                        previousAction = action;
                    }
                    catch (Exception ex)
                    {
                        if (ex is OperationCanceledException && (CancelRequested || cancellationToken.IsCancellationRequested))
                        {
                            Console.WriteLine("Aborted (hotkey).");
                            followUps.Clear();
                            return new BatchedActionExecutionResult(executed, "batched actions cancelled", Cancelled: true, Interrupted: true, FailedAction: snapshot);
                        }
                        Console.WriteLine($"Batched action execution error: {ex.Message}");
                        followUps.Clear();
                        return new BatchedActionExecutionResult(
                            executed,
                            $"{snapshot.Description} failed in action batch: {ex.Message}", Interrupted: true, FailedAction: snapshot);
                    }
                }

                return new BatchedActionExecutionResult(executed, null);
            }

            internal static bool IsAdaptiveBatchWaitTrigger(ActionDto action) =>
                action.Type is "open_url" or "launch_app" or "run_command" ||
                BatchPlanner.IsCommitKeys(action);

            internal static async Task WaitForBatchedTransitionAsync(
                byte[] beforeFingerprint,
                ActionDto triggerAction,
                int timeoutMs,
                bool navigationLikeBatch,
                CancellationToken cancellationToken)
            {
                if (timeoutMs <= 0)
                    return;

                var sw = RuntimeClock.StartTimer();
                byte[] previous = beforeFingerprint;
                var sawChange = false;
                var lastChangeAtMs = 0L;
                var probes = 0;
                var threshold = SettleThresholdFor(triggerAction);
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
                var minimumObservationMs = AdaptiveBatchMinimumObservationMs(
                    timeoutMs,
                    navigationLikeBatch);
                var quietWindowMs = Math.Min(350, timeoutMs);

                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var remaining = timeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await RuntimeClock.Delay(Math.Min(100, remaining), cancellationToken);

                    var current = CaptureScreenFingerprintProbe();
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    lastBetween = ComputeImageDelta(previous, current);
                    if (lastFromBefore >= threshold)
                        sawChange = true;
                    if (lastBetween >= threshold)
                        lastChangeAtMs = sw.ElapsedMilliseconds;

                    previous = current;
                    if (sawChange &&
                        sw.ElapsedMilliseconds >= minimumObservationMs &&
                        sw.ElapsedMilliseconds - lastChangeAtMs >= quietWindowMs)
                    {
                        Console.WriteLine(
                            $"[batch] adaptive wait settled after {sw.ElapsedMilliseconds}ms; minimum={minimumObservationMs}ms; cap={timeoutMs}ms; probes={probes}; action={triggerAction.Type}");
                        return;
                    }
                }

                Console.WriteLine(
                    $"[batch] adaptive wait reached cap after {sw.ElapsedMilliseconds}ms; probes={probes}; coarse_delta={lastFromBefore:0.####}; coarse_between_delta={lastBetween:0.####}; action={triggerAction.Type}");
            }

            internal static int AdaptiveBatchMinimumObservationMs(
                int timeoutMs,
                bool navigationLikeBatch) =>
                navigationLikeBatch
                    ? Math.Max(0, timeoutMs)
                    : Math.Min(900, Math.Max(0, timeoutMs));

            internal static bool IsNavigationBatchPrelude(ActionDto action)
            {
                if (action.Type == "open_url")
                    return true;
                if (action.Type != "keys" || action.Keys is not { Length: > 0 } keys)
                    return false;

                var normalized = keys
                    .SelectMany(key => (key ?? "").Split(
                        '+',
                        StringSplitOptions.RemoveEmptyEntries |
                        StringSplitOptions.TrimEntries))
                    .Select(key => key.ToLowerInvariant())
                    .ToArray();
                return normalized.SequenceEqual(["win", "r"]) ||
                       normalized.SequenceEqual(["super", "r"]) ||
                       normalized.SequenceEqual(["meta", "r"]) ||
                       normalized.SequenceEqual(["ctrl", "l"]) ||
                       normalized.SequenceEqual(["ctrl", "k"]);
            }

            internal static int BatchTransitionDelayMs(ActionDto previous, ActionDto next)
            {
                if (next.Type == "wait")
                    return 0;

                if (previous.Type == "keys" && BatchPlanner.IsTextEntryPrelude(previous) &&
                    next.Type is "type_text" or "paste_text")
                {
                    return 250;
                }

                if (previous.Type is "type_text" or "paste_text" && BatchPlanner.IsCommitKeys(next))
                    return 500;

                if (previous.Type is "type_text" or "paste_text" &&
                    next.Type is "type_text" or "paste_text")
                {
                    return 100;
                }

                return Math.Max(80, DelayFor(previous));
            }

            internal static bool AllowsStableCanvasDrawBatch(
                UiPromptContext context,
                string observationProfile)
            {
                if (!Options.MouseEnabled ||
                    !Options.DirectClickWithoutAim ||
                    observationProfile is "realtime_interaction" or "turn_based_interaction" or "streaming_output")
                {
                    return false;
                }

                if (string.Equals(
                        observationProfile,
                        "local_editing",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                var semantic = $"{context.ActiveProcessName} {context.ActiveWindowTitle} {context.FocusedUiaSummary}";
                return Regex.IsMatch(
                    semantic,
                    @"\b(mspaint|paint|canvas|drawing|sketch|krita|gimp|photoshop|inkscape|płótno|rysunek)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }
        
            internal static bool IsSafeBatchedAction(ActionDto action)
            {
                if (action.Type == "run_command" && !Options.AllowRunCommand)
                    return false;
        
                if (action.Type is "open_url" or "launch_app")
                    return Options.AllowHighLevelActions;
        
                return action.Type is "keys" or "type_text" or "paste_text" or "wait" or "run_command" or "drag_path";
            }
    }
