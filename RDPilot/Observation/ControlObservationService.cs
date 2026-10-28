
namespace RDPilot.Observation;

    internal static class ControlObservationService
    {
        
            internal static int PostActionDelay(ActionDto a)
            {
                // These actions only adjust the next prompt/crop; the target UI has not changed.
                if (a.Type == "wait" || IsLocalObservationAction(a))
                    return 0;
        
                // A plain cursor move can affect hover state, but does not need the full app settle delay.
                if (IsPointerPositionOnlyAction(a))
                    return Math.Max(0, DelayFor(a));
        
                return Math.Max(Options.UiSettleDelayMs, DelayFor(a));
            }
        
            internal static async Task<bool> WaitAfterActionAsync(
                ActionDto action,
                byte[] beforeFingerprint,
                string observationPolicy,
                byte[]? beforeLocalFingerprint,
                Rectangle? localRegion,
                CancellationToken cancellationToken,
                bool preferFastTurnSettle = false)
            {
                if (IsLocalObservationAction(action))
                    return false;
        
                if (!Options.ScreenPollingEnabled)
                {
                    var fixedDelay = PostActionDelay(action);
                    if (fixedDelay > 0)
                        await RuntimeClock.Delay(fixedDelay, cancellationToken);
                    return false;
                }
        
                if (action.Type == "wait")
                {
                    if (Options.WaitNoChangeExtraMs <= 0)
                        return false;
        
                    var afterWait = CaptureScreenFingerprintProbe();
                    var delta = ComputeImageDelta(beforeFingerprint, afterWait);
                    if (delta < SettleThresholdFor(action))
                    {
                        Console.WriteLine($"[settle] screen unchanged after wait (coarse_delta={delta:0.####}); waiting extra {Options.WaitNoChangeExtraMs} ms.");
                        await RuntimeClock.Delay(Options.WaitNoChangeExtraMs, cancellationToken);
                    }
                    return false;
                }
        
                var initialDelay = Math.Min(PostActionDelay(action), Math.Max(0, Options.ScreenPollInitialDelayMs));
                if (initialDelay > 0)
                    await RuntimeClock.Delay(initialDelay, cancellationToken);
        
                if (Options.ScreenPollTimeoutMs <= 0)
                    return false;

                if (string.Equals(
                        observationPolicy,
                        "turn_based_interaction",
                        StringComparison.Ordinal) &&
                    beforeLocalFingerprint is { Length: > 0 } &&
                    localRegion is Rectangle turnRegion)
                {
                    return await WaitForLocalScreenReactionAndStableAsync(
                        beforeLocalFingerprint,
                        turnRegion,
                        action,
                        cancellationToken,
                        preferFastTurnSettle);
                }

                if (observationPolicy == "realtime_interaction" ||
                    action.Type == "hold_keys" ||
                    action.Type == "drag_path" && action.GestureKind is "game" or "pan")
                    return false;

                if (observationPolicy is "event_driven" or "streaming_output")
                {
                    await WaitForScreenReactionAsync(
                        beforeFingerprint,
                        action,
                        cancellationToken);
                    return false;
                }

                await WaitForScreenStableAsync(beforeFingerprint, action, cancellationToken);
                return false;
            }

            internal static async Task<bool> WaitForLocalScreenReactionAndStableAsync(
                byte[] beforeFingerprint,
                Rectangle region,
                ActionDto action,
                CancellationToken cancellationToken,
                bool preferFastSettle = false)
            {
                const double reactionThreshold = 0.0012;
                const double reactionRatioThreshold = 0.0015;
                const double stableThreshold = 0.0007;
                const double stableRatioThreshold = 0.0008;
                var requiredStableProbes = preferFastSettle ? 1 : 2;
                var pollIntervalMs = preferFastSettle
                    ? Math.Min(Options.ScreenPollIntervalMs, 75)
                    : Options.ScreenPollIntervalMs;
                var timeoutMs = preferFastSettle
                    ? Math.Min(Options.ScreenPollTimeoutMs, 650)
                    : Options.ScreenPollTimeoutMs;
                var settleMode = preferFastSettle ? "fast" : "adaptive";

                var sw = RuntimeClock.StartTimer();
                byte[]? previous = null;
                var sawReaction = false;
                var stableProbes = 0;
                var probes = 0;
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
                double lastFromBeforeRatio = double.NaN;
                double lastBetweenRatio = double.NaN;
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var current = CaptureRegionFingerprintProbe(region);
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    lastFromBeforeRatio = ComputeChangedPixelRatio(
                        beforeFingerprint,
                        current);
                    if (lastFromBefore >= reactionThreshold ||
                        lastFromBeforeRatio >= reactionRatioThreshold)
                        sawReaction = true;

                    if (sawReaction && previous is not null)
                    {
                        lastBetween = ComputeImageDelta(previous, current);
                        lastBetweenRatio = ComputeChangedPixelRatio(previous, current);
                        stableProbes = lastBetween < stableThreshold &&
                                       lastBetweenRatio < stableRatioThreshold
                            ? stableProbes + 1
                            : 0;
                        if (stableProbes >= requiredStableProbes)
                        {
                            Console.WriteLine(
                                $"[settle] local reaction stable after {sw.ElapsedMilliseconds} ms; probes={probes}; local_delta={lastFromBefore:0.####}; local_ratio={lastFromBeforeRatio:0.####}; action={action.Type}; mode={settleMode}");
                            return true;
                        }
                    }

                    previous = current;
                    var remaining = timeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await RuntimeClock.Delay(
                        Math.Min(pollIntervalMs, remaining),
                        cancellationToken);
                }

                Console.WriteLine(
                    sawReaction
                        ? $"[settle] local reaction reached timeout after {sw.ElapsedMilliseconds} ms; probes={probes}; local_delta={lastFromBefore:0.####}; local_ratio={lastFromBeforeRatio:0.####}; local_between_delta={lastBetween:0.####}; local_between_ratio={lastBetweenRatio:0.####}; action={action.Type}; mode={settleMode}"
                        : $"[settle] no local reaction within {sw.ElapsedMilliseconds} ms; probes={probes}; local_delta={lastFromBefore:0.####}; local_ratio={lastFromBeforeRatio:0.####}; action={action.Type}; mode={settleMode}");
                return sawReaction;
            }

            internal static async Task WaitForScreenReactionAsync(
                byte[] beforeFingerprint,
                ActionDto action,
                CancellationToken cancellationToken)
            {
                var sw = RuntimeClock.StartTimer();
                var probes = 0;
                double delta = double.NaN;
                var threshold = SettleThresholdFor(action);
                while (sw.ElapsedMilliseconds < Options.ScreenPollTimeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    delta = ComputeImageDelta(
                        beforeFingerprint,
                        CaptureScreenFingerprintProbe());
                    probes++;
                    if (delta >= threshold)
                    {
                        Console.WriteLine(
                            $"[settle] reaction started after {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={delta:0.####}; action={action.Type}");
                        return;
                    }
                    var remaining = Options.ScreenPollTimeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await RuntimeClock.Delay(
                        Math.Min(Options.ScreenPollIntervalMs, remaining),
                        cancellationToken);
                }
                Console.WriteLine(
                    $"[settle] no visible reaction within {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={delta:0.####}; action={action.Type}");
            }
        
            internal static async Task WaitForScreenStableAsync(byte[] beforeFingerprint, ActionDto action, CancellationToken cancellationToken)
            {
                var sw = RuntimeClock.StartTimer();
                byte[]? previous = null;
                var sawChange = false;
                var probes = 0;
                double lastFromBefore = double.NaN;
                double lastBetween = double.NaN;
                var settleThreshold = SettleThresholdFor(action);
        
                while (sw.ElapsedMilliseconds < Options.ScreenPollTimeoutMs)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var current = CaptureScreenFingerprintProbe();
                    probes++;
                    lastFromBefore = ComputeImageDelta(beforeFingerprint, current);
                    if (lastFromBefore >= settleThreshold)
                        sawChange = true;
        
                    if (previous != null)
                    {
                        lastBetween = ComputeImageDelta(previous, current);
                        if (lastBetween < settleThreshold && (sawChange || probes >= 2))
                        {
                            Console.WriteLine($"[settle] stable after {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={lastFromBefore:0.####}; action={action.Type}");
                            return;
                        }
                    }
        
                    previous = current;
                    var remaining = Options.ScreenPollTimeoutMs - (int)sw.ElapsedMilliseconds;
                    if (remaining <= 0)
                        break;
                    await RuntimeClock.Delay(Math.Min(Options.ScreenPollIntervalMs, remaining), cancellationToken);
                }
        
                if (probes > 0)
                    Console.WriteLine($"[settle] timeout after {sw.ElapsedMilliseconds} ms; probes={probes}; coarse_delta={lastFromBefore:0.####}; coarse_between_delta={lastBetween:0.####}; action={action.Type}");
            }
        
            internal static byte[] CaptureScreenFingerprintProbe()
            {
                var sw = RuntimeClock.StartTimer();
                try
                {
                    var (vx, vy, vw, vh) = GetPrimaryScreen();
                    using var bmp = new Bitmap(vw, vh, PixelFormat.Format24bppRgb);
                    using (var g = Graphics.FromImage(bmp))
                        DesktopCapture.Backend.CopyPixels(g, new Rectangle(vx, vy, vw, vh));
                    return BuildImageFingerprint(bmp);
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunScreenProbeCount++;
                    CurrentRun.RunScreenProbeElapsed += sw.Elapsed;
                }
            }

            internal static double SettleThresholdFor(ActionDto action) => action.Type switch
            {
                "drag_path" => 0.0025,
                "click" or "double_click" or "focus_uia" or "click_uia" or
                    "type_text" or "paste_text" or "keys" => 0.0035,
                "hold_keys" => 0.008,
                _ => Options.NoChangeThreshold
            };

            internal static ResolvedActionSnapshot AttachFocusedTextObservationRegion(
                ResolvedActionSnapshot snapshot,
                Rectangle? focusedUiaRect)
            {
                if (snapshot.Action.Type is not ("type_text" or "paste_text") ||
                    focusedUiaRect is not Rectangle focus ||
                    focus.Width <= 0 ||
                    focus.Height <= 0)
                {
                    return snapshot;
                }

                focus.Inflate(8, 8);
                return snapshot with { ObservationRegion = ClampRect(focus) };
            }

            internal static ResolvedActionSnapshot AttachTurnBasedObservationRegion(
                ResolvedActionSnapshot snapshot,
                Rectangle? interactionRegion,
                string observationPolicy)
            {
                if (!string.Equals(
                        observationPolicy,
                        "turn_based_interaction",
                        StringComparison.Ordinal) ||
                    interactionRegion is not Rectangle region ||
                    region.Width <= 0 ||
                    region.Height <= 0 ||
                    !IsStateChangingInteractionAction(snapshot.Action))
                {
                    return snapshot;
                }

                return snapshot with { ObservationRegion = ClampRect(region) };
            }

            internal static Rectangle? ResolveTurnBasedObservationRegion(
                Rectangle? interactionRegion,
                string observationPolicy)
            {
                if (!string.Equals(
                        observationPolicy,
                        "turn_based_interaction",
                        StringComparison.Ordinal))
                {
                    return interactionRegion;
                }

                if (interactionRegion is Rectangle persistent &&
                    persistent.Width > 0 &&
                    persistent.Height > 0)
                {
                    return ClampRect(persistent);
                }

                var activeWindow = ScreenshotService.ActiveWindowRectangleProvider();
                if (activeWindow is Rectangle active &&
                    active.Width > 0 &&
                    active.Height > 0)
                {
                    return ClampRect(active);
                }

                var (screenX, screenY, screenWidth, screenHeight) = GetPrimaryScreen();
                return new Rectangle(screenX, screenY, screenWidth, screenHeight);
            }

            internal static Rectangle? InferTurnBasedObservationRegion(
                ScreenObservationFrame before,
                ScreenObservationFrame after,
                Rectangle currentRegion)
            {
                if (before.DetailWidth <= 0 ||
                    before.DetailHeight <= 0 ||
                    before.DetailWidth != after.DetailWidth ||
                    before.DetailHeight != after.DetailHeight ||
                    before.DetailFingerprint.Length != after.DetailFingerprint.Length ||
                    before.ScreenBounds != after.ScreenBounds)
                {
                    return null;
                }

                currentRegion.Intersect(before.ScreenBounds);
                if (currentRegion.Width <= 0 || currentRegion.Height <= 0)
                    return null;

                var detailWidth = before.DetailWidth;
                var detailHeight = before.DetailHeight;
                var bounds = before.ScreenBounds;
                var left = Math.Clamp(
                    (int)Math.Floor(
                        (currentRegion.Left - bounds.Left) * detailWidth /
                        (double)Math.Max(1, bounds.Width)),
                    0,
                    detailWidth - 1);
                var top = Math.Clamp(
                    (int)Math.Floor(
                        (currentRegion.Top - bounds.Top) * detailHeight /
                        (double)Math.Max(1, bounds.Height)),
                    0,
                    detailHeight - 1);
                var right = Math.Clamp(
                    (int)Math.Ceiling(
                        (currentRegion.Right - bounds.Left) * detailWidth /
                        (double)Math.Max(1, bounds.Width)),
                    left + 1,
                    detailWidth);
                var bottom = Math.Clamp(
                    (int)Math.Ceiling(
                        (currentRegion.Bottom - bounds.Top) * detailHeight /
                        (double)Math.Max(1, bounds.Height)),
                    top + 1,
                    detailHeight);

                var changed = new bool[detailWidth * detailHeight];
                var useColor = before.DetailColorFingerprint.Length ==
                                   detailWidth * detailHeight * 3 &&
                               after.DetailColorFingerprint.Length ==
                                   detailWidth * detailHeight * 3;
                for (var y = top; y < bottom; y++)
                for (var x = left; x < right; x++)
                {
                    var pixel = y * detailWidth + x;
                    var maximumDifference = Math.Abs(
                        before.DetailFingerprint[pixel] -
                        after.DetailFingerprint[pixel]);
                    if (useColor)
                    {
                        var color = pixel * 3;
                        maximumDifference = Math.Max(
                            maximumDifference,
                            Math.Max(
                                Math.Abs(before.DetailColorFingerprint[color] -
                                         after.DetailColorFingerprint[color]),
                                Math.Max(
                                    Math.Abs(before.DetailColorFingerprint[color + 1] -
                                             after.DetailColorFingerprint[color + 1]),
                                    Math.Abs(before.DetailColorFingerprint[color + 2] -
                                             after.DetailColorFingerprint[color + 2]))));
                    }
                    changed[pixel] = maximumDifference >= 12;
                }

                var regionPixelCount = (right - left) * (bottom - top);
                var minimumComponentSize = Math.Max(
                    4,
                    (int)Math.Ceiling(regionPixelCount * 0.0015));
                var visited = new bool[changed.Length];
                var meaningfulComponents = new List<(int Left, int Top, int Right, int Bottom, int Count)>();
                for (var start = 0; start < changed.Length; start++)
                {
                    if (!changed[start] || visited[start])
                        continue;

                    var startX = start % detailWidth;
                    var startY = start / detailWidth;
                    if (startX < left || startX >= right || startY < top || startY >= bottom)
                        continue;

                    var queue = new Queue<int>();
                    queue.Enqueue(start);
                    visited[start] = true;
                    var componentLeft = startX;
                    var componentTop = startY;
                    var componentRight = startX + 1;
                    var componentBottom = startY + 1;
                    var count = 0;
                    while (queue.Count > 0)
                    {
                        var current = queue.Dequeue();
                        var x = current % detailWidth;
                        var y = current / detailWidth;
                        count++;
                        componentLeft = Math.Min(componentLeft, x);
                        componentTop = Math.Min(componentTop, y);
                        componentRight = Math.Max(componentRight, x + 1);
                        componentBottom = Math.Max(componentBottom, y + 1);
                        for (var dy = -1; dy <= 1; dy++)
                        for (var dx = -1; dx <= 1; dx++)
                        {
                            if (dx == 0 && dy == 0)
                                continue;
                            var nx = x + dx;
                            var ny = y + dy;
                            if (nx < left || nx >= right || ny < top || ny >= bottom)
                                continue;
                            var neighbor = ny * detailWidth + nx;
                            if (!changed[neighbor] || visited[neighbor])
                                continue;
                            visited[neighbor] = true;
                            queue.Enqueue(neighbor);
                        }
                    }

                    if (count >= minimumComponentSize)
                    {
                        meaningfulComponents.Add((
                            componentLeft,
                            componentTop,
                            componentRight,
                            componentBottom,
                            count));
                    }
                }

                if (meaningfulComponents.Sum(component => component.Count) <
                    Math.Max(8, regionPixelCount * 0.004))
                {
                    return null;
                }

                var changedLeft = meaningfulComponents.Min(component => component.Left);
                var changedTop = meaningfulComponents.Min(component => component.Top);
                var changedRight = meaningfulComponents.Max(component => component.Right);
                var changedBottom = meaningfulComponents.Max(component => component.Bottom);
                var horizontalPadding = Math.Max(2, (right - left) / 40);
                var verticalPadding = Math.Max(2, (bottom - top) / 40);
                changedLeft = Math.Max(left, changedLeft - horizontalPadding);
                changedTop = Math.Max(top, changedTop - verticalPadding);
                changedRight = Math.Min(right, changedRight + horizontalPadding);
                changedBottom = Math.Min(bottom, changedBottom + verticalPadding);

                var inferred = Rectangle.FromLTRB(
                    bounds.Left + (int)Math.Floor(
                        changedLeft * bounds.Width / (double)detailWidth),
                    bounds.Top + (int)Math.Floor(
                        changedTop * bounds.Height / (double)detailHeight),
                    bounds.Left + (int)Math.Ceiling(
                        changedRight * bounds.Width / (double)detailWidth),
                    bounds.Top + (int)Math.Ceiling(
                        changedBottom * bounds.Height / (double)detailHeight));
                inferred.Intersect(currentRegion);
                if (inferred.Width < currentRegion.Width * 0.10 ||
                    inferred.Height < currentRegion.Height * 0.10)
                {
                    return null;
                }

                var currentArea = (long)currentRegion.Width * currentRegion.Height;
                var inferredArea = (long)inferred.Width * inferred.Height;
                var inferredRatio = inferredArea / (double)Math.Max(1, currentArea);
                return inferredRatio is >= 0.03 and <= 0.80
                    ? ClampRect(inferred)
                    : null;
            }

            internal static bool IsTurnBasedStateInspection(ActionDto action)
            {
                if (action.Type != "request_crop")
                    return false;

                var note = action.Note ?? "";
                if (Regex.IsMatch(
                        note,
                        @"\b(grid|playfield|game area|game state|puzzle|maze|map|canvas|siatk\w*|labirynt\w*|pole gr\w*)\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    return true;
                }
                if (IsTurnBasedAuxiliaryInspection(action))
                    return false;
                return Regex.IsMatch(
                    note,
                    @"\b(board|grid|playfield|level|game area|game state|puzzle|maze|map|canvas|plansz\w*|siatk\w*|poziom\w*|stan\w* gr\w*|map\w*|labirynt\w*|pole gr\w*)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            internal static bool ShouldEstablishTurnBasedInteractionRegion(
                ActionDto action,
                bool regionRequired) =>
                action.Type == "request_crop" &&
                (IsTurnBasedStateInspection(action) ||
                 regionRequired && !IsTurnBasedAuxiliaryInspection(action));

            internal static bool CanReplaceTurnBasedInteractionRegion(
                bool hasExistingRegion,
                bool existingRegionIsAutomatic,
                bool automaticRegionWasRefined,
                ActionDto action,
                bool regionRequired) =>
                ShouldEstablishTurnBasedInteractionRegion(action, regionRequired) &&
                (!hasExistingRegion ||
                 existingRegionIsAutomatic && !automaticRegionWasRefined);

            internal static bool IsTurnBasedAuxiliaryInspection(ActionDto action)
            {
                if (action.Type != "request_crop")
                    return false;

                return Regex.IsMatch(
                    action.Note ?? "",
                    @"\b(control\w*|button\w*|help\w*|instruction\w*|toolbar\w*|panel\w*|menu\w*|sterow\w*|przycisk\w*|pomoc\w*|instrukcj\w*)\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            internal static bool IsOverlappingTurnInspection(
                Rectangle persistentRegion,
                Rectangle requestedRegion)
            {
                var intersection = Rectangle.Intersect(
                    persistentRegion,
                    requestedRegion);
                if (intersection.Width <= 0 || intersection.Height <= 0)
                    return false;
                var intersectionArea = (long)intersection.Width * intersection.Height;
                var smallerArea = Math.Min(
                    (long)persistentRegion.Width * persistentRegion.Height,
                    (long)requestedRegion.Width * requestedRegion.Height);
                return smallerArea > 0 &&
                       intersectionArea / (double)smallerArea >= 0.60;
            }

            internal static string TurnBasedContextKey(UiPromptContext context) =>
                $"{context.ActiveProcessName?.Trim().ToLowerInvariant()}|" +
                context.ActiveWindowTitle?.Trim().ToLowerInvariant();
    }
