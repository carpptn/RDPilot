namespace RDPilot.Control;

/// <summary>Plans bounded follow-ups and resolves directional inputs independently of API parsing.</summary>
internal static class BatchPlanner
{
    internal static ControlDecision CreateDecision(
        List<ActionDto> candidates,
        bool allowDrawGestureBatch = false,
        int observedTurnBatchLimit = 0,
        IReadOnlyList<string>? plannedInputs = null)
    {
        if (!Options.ExecuteMultiActionCandidates || Options.MaxQueuedBatchActions <= 0 || candidates.Count <= 1)
            return new ControlDecision(candidates[0]);

        var queuedActions = SafeBatchFollowUps(
            candidates,
            allowDrawGestureBatch,
            observedTurnBatchLimit, plannedInputs);

        if (queuedActions.Length > 0)
        {
            Console.WriteLine($"[batch] queued {queuedActions.Length} safe follow-up action(s) from the same response.");
            if (candidates[0].Type == "drag_path")
            {
                var acceptedGestures = new[] { candidates[0] }.Concat(queuedActions).ToArray();
                Console.WriteLine(
                    $"[batch] stable-canvas draw sequence accepted; strokes={acceptedGestures.Length}; points={acceptedGestures.Sum(action => action.Path?.Length ?? 0)}; duration_ms={acceptedGestures.Sum(action => action.DurationMs ?? 0)}");
            }
        }
        if (queuedActions.Length < candidates.Count - 1)
        {
            Console.WriteLine(
                $"[batch] observation barrier discarded {candidates.Count - 1 - queuedActions.Length} unsafe or state-dependent follow-up action(s).");
        }
        return new ControlDecision(candidates[0], queuedActions);
    }

    internal static ActionDto[] SafeBatchFollowUps(
        List<ActionDto> candidates,
        bool allowDrawGestureBatch = false,
        int observedTurnBatchLimit = 0,
        IReadOnlyList<string>? plannedInputs = null)
    {
        if (Options.MaxQueuedBatchActions <= 0 || candidates.Count <= 1)
            return Array.Empty<ActionDto>();

        if (observedTurnBatchLimit > 0)
        {
            var observedTurnActions = SafeObservedTurnFollowUps(
                candidates,
                observedTurnBatchLimit, plannedInputs);
            if (observedTurnActions.Length > 0)
                return observedTurnActions;
        }

        var previous = candidates[0];
        if (!IsWellFormedSafeBatchedAction(previous))
            return Array.Empty<ActionDto>();

        var result = new List<ActionDto>();

        var gestureBatch = previous.Type == "drag_path";
        var gesturePoints = gestureBatch ? previous.Path?.Length ?? 0 : 0;
        var gestureDurationMs = gestureBatch ? previous.DurationMs ?? 0 : 0;
        if (gestureBatch &&
            (!allowDrawGestureBatch || !IsSafeBatchedDrawGesture(previous)))
            return Array.Empty<ActionDto>();

        foreach (var candidate in candidates.Skip(1))
        {
            if (result.Count >= Options.MaxQueuedBatchActions)
                break;

            if (!IsSafeBatchedAction(candidate))
                break;

            if (!IsWellFormedSafeBatchedAction(candidate))
                break;

            if (!CanBatchWithoutObservation(previous, candidate))
                break;

            if (gestureBatch)
            {
                if (!IsSafeBatchedDrawGesture(candidate))
                    break;
                var nextPoints = gesturePoints + (candidate.Path?.Length ?? 0);
                var nextDurationMs = gestureDurationMs + (candidate.DurationMs ?? 0);
                if (nextPoints > Options.MaxBatchedGesturePoints ||
                    nextDurationMs > Options.MaxBatchedGestureDurationMs)
                {
                    break;
                }
                gesturePoints = nextPoints;
                gestureDurationMs = nextDurationMs;
            }

            result.Add(candidate);
            previous = candidate;
            if (candidate.Type == "wait")
                break;
        }

        return result.ToArray();
    }

    internal static ActionDto[] SafeObservedTurnFollowUps(
        IReadOnlyList<ActionDto> candidates,
        int observedTurnBatchLimit,
        IReadOnlyList<string>? plannedInputs = null)
    {
        plannedInputs ??= candidates.Count > 0 ? candidates[0].PlannedInputs : null;
        if (candidates.Count <= 1 ||
            observedTurnBatchLimit <= 0 ||
            plannedInputs is not { Count: >= 2 } ||
            HasImmediateDirectionalReversal(plannedInputs) ||
            (candidates[0].PlanConfidence ?? candidates[0].Confidence ?? 0) <
            TurnBasedTransitionTracker.MinimumStructuredPlanConfidence ||
            !TryBindObservedTurnActionToPlannedInput(
                candidates[0],
                plannedInputs[0]))
        {
            return Array.Empty<ActionDto>();
        }

        var limit = Math.Min(
            observedTurnBatchLimit,
            plannedInputs.Count - 1);
        var result = new List<ActionDto>(limit);
        var directionalClickPoints = new Dictionary<string, (int X, int Y)>(
            StringComparer.Ordinal);
        if (!TryReconcileObservedDirectionalClick(
                candidates[0],
                plannedInputs[0],
                directionalClickPoints))
        {
            return Array.Empty<ActionDto>();
        }
        for (var index = 1; index < candidates.Count && result.Count < limit; index++)
        {
            var candidate = candidates[index];
            if (!TryReconcileObservedDirectionalClick(
                    candidate,
                    plannedInputs[index],
                    directionalClickPoints) ||
                !TryBindObservedTurnActionToPlannedInput(
                    candidate,
                    plannedInputs[index]) ||
                !CaptureResolvedAction(candidate, null).IsValid)
            {
                break;
            }

            result.Add(candidate);
        }

        return result.ToArray();
    }

    internal static bool HasImmediateDirectionalReversal(
        IReadOnlyList<string> inputs)
    {
        string? previous = null;
        foreach (var input in inputs)
        {
            if (!TryNormalizeDirectionalLabel(input, out var current))
            {
                previous = null;
                continue;
            }

            if (previous is not null &&
                (previous == "ArrowLeft" && current == "ArrowRight" ||
                 previous == "ArrowRight" && current == "ArrowLeft" ||
                 previous == "ArrowUp" && current == "ArrowDown" ||
                 previous == "ArrowDown" && current == "ArrowUp"))
            {
                return true;
            }

            previous = current;
        }

        return false;
    }

    static bool TryReconcileObservedDirectionalClick(
        ActionDto action,
        string? plannedInput,
        IDictionary<string, (int X, int Y)> knownPoints)
    {
        if (action.Type is not ("click" or "double_click"))
            return true;
        if (!TryNormalizeDirectionalLabel(plannedInput, out var plannedLabel) ||
            !TryGetActionImagePoint(action, out var actionPoint))
        {
            return false;
        }

        const int sameControlTolerance = 4;
        if (knownPoints.TryGetValue(plannedLabel, out var knownPoint))
        {
            if (PointDistanceSquared(actionPoint, knownPoint) >
                sameControlTolerance * sameControlTolerance)
            {
                RecenterClickAction(action, actionPoint, knownPoint);
                Console.WriteLine(
                    $"[batch] corrected a repeated {plannedLabel} click to its previously observed control position.");
            }
            return true;
        }

        var conflictingPoint = knownPoints.FirstOrDefault(entry =>
            !string.Equals(entry.Key, plannedLabel, StringComparison.Ordinal) &&
            PointDistanceSquared(actionPoint, entry.Value) <=
            sameControlTolerance * sameControlTolerance);
        if (!string.IsNullOrWhiteSpace(conflictingPoint.Key))
        {
            if (!TryInferDirectionalPadPoint(
                    knownPoints,
                    plannedLabel,
                    out var inferredPoint))
            {
                Console.WriteLine(
                    $"[batch] rejected inconsistent directional click: {plannedLabel} reused the {conflictingPoint.Key} control position.");
                return false;
            }

            RecenterClickAction(action, actionPoint, inferredPoint);
            actionPoint = inferredPoint;
            Console.WriteLine(
                $"[batch] repaired inconsistent directional click geometry: {plannedLabel} had reused the {conflictingPoint.Key} control position.");
        }

        knownPoints[plannedLabel] = actionPoint;
        return true;
    }

    static bool TryGetActionImagePoint(
        ActionDto action,
        out (int X, int Y) point)
    {
        if (action.BBox is { Left: { } left, Top: { } top, Right: { } right, Bottom: { } bottom })
        {
            point = ((left + right) / 2, (top + bottom) / 2);
            return true;
        }
        if (action.XPx is int x && action.YPx is int y)
        {
            point = (x, y);
            return true;
        }
        if (action.X is double normalizedX && action.Y is double normalizedY)
        {
            point = (
                (int)Math.Round(normalizedX * Math.Max(1, CurrentRun.CurrentScreenMap.ImageW - 1)),
                (int)Math.Round(normalizedY * Math.Max(1, CurrentRun.CurrentScreenMap.ImageH - 1)));
            return true;
        }
        point = default;
        return false;
    }

    static long PointDistanceSquared(
        (int X, int Y) first,
        (int X, int Y) second)
    {
        var dx = (long)first.X - second.X;
        var dy = (long)first.Y - second.Y;
        return dx * dx + dy * dy;
    }

    static bool TryInferDirectionalPadPoint(
        IDictionary<string, (int X, int Y)> knownPoints,
        string targetLabel,
        out (int X, int Y) point)
    {
        var vertical = knownPoints.FirstOrDefault(entry =>
            entry.Key is "ArrowUp" or "ArrowDown");
        var horizontal = knownPoints.FirstOrDefault(entry =>
            entry.Key is "ArrowLeft" or "ArrowRight");
        if (string.IsNullOrWhiteSpace(vertical.Key) ||
            string.IsNullOrWhiteSpace(horizontal.Key))
        {
            point = default;
            return false;
        }

        var centerX = vertical.Value.X;
        var centerY = horizontal.Value.Y;
        var verticalOffset = Math.Abs(vertical.Value.Y - centerY);
        var horizontalOffset = Math.Abs(horizontal.Value.X - centerX);
        var orientationMatches =
            (vertical.Key == "ArrowUp" && vertical.Value.Y < centerY ||
             vertical.Key == "ArrowDown" && vertical.Value.Y > centerY) &&
            (horizontal.Key == "ArrowLeft" && horizontal.Value.X < centerX ||
             horizontal.Key == "ArrowRight" && horizontal.Value.X > centerX);
        if (!orientationMatches ||
            verticalOffset < 4 ||
            horizontalOffset < 4 ||
            verticalOffset > horizontalOffset * 2 ||
            horizontalOffset > verticalOffset * 2)
        {
            point = default;
            return false;
        }

        point = targetLabel switch
        {
            "ArrowUp" => (centerX, centerY - verticalOffset),
            "ArrowDown" => (centerX, centerY + verticalOffset),
            "ArrowLeft" => (centerX - horizontalOffset, centerY),
            "ArrowRight" => (centerX + horizontalOffset, centerY),
            _ => default
        };
        return targetLabel is "ArrowUp" or "ArrowDown" or "ArrowLeft" or "ArrowRight";
    }

    static void RecenterClickAction(
        ActionDto action,
        (int X, int Y) oldPoint,
        (int X, int Y) newPoint)
    {
        var dx = newPoint.X - oldPoint.X;
        var dy = newPoint.Y - oldPoint.Y;
        if (action.BBox is { Left: { } left, Top: { } top, Right: { } right, Bottom: { } bottom })
        {
            action.BBox.Left = left + dx;
            action.BBox.Top = top + dy;
            action.BBox.Right = right + dx;
            action.BBox.Bottom = bottom + dy;
        }
        action.XPx = newPoint.X;
        action.YPx = newPoint.Y;
    }

    internal static bool TryBindObservedTurnActionToPlannedInput(
        ActionDto action,
        string? plannedInput)
    {
        if (!TryNormalizeDirectionalLabel(plannedInput, out var plannedLabel))
            return false;

        if (action.Type is "click" or "double_click")
        {
            if (!Options.MouseEnabled ||
                !Options.DirectClickWithoutAim ||
                !HasExplicitPoint(action))
            {
                return false;
            }

            action.ResolvedTurnInputLabel = plannedLabel;
            return true;
        }

        if (action.Type != "keys" || action.Keys is not { Length: 1 } ||
            TurnInput.Key(action.Keys[0]) is not string actionKey ||
            actionKey != TurnInput.Key(plannedInput))
        {
            return false;
        }

        action.ResolvedTurnInputLabel = actionKey;
        return true;
    }

    internal static bool TryGetObservedTurnInputLabel(
        ActionDto action,
        out string label)
    {
        label = "";
        if (TryNormalizeDirectionalLabel(
                action.ResolvedTurnInputLabel,
                out label))
        {
            return true;
        }

        if (action.Type == "keys" && action.Keys is { Length: 1 })
            return TryNormalizeDirectionalLabel(action.Keys[0], out label);

        if (action.Type is not ("click" or "double_click") ||
            !Options.MouseEnabled ||
            !Options.DirectClickWithoutAim ||
            !HasExplicitPoint(action))
        {
            return false;
        }

        return TryNormalizeUnambiguousDirectionalLabel(action.Note, out label);
    }

    internal static bool TryNormalizeUnambiguousDirectionalLabel(
        string? value,
        out string label)
    {
        var text = value ?? "";
        var matches = new List<string>(4);
        if (Regex.IsMatch(text, @"\b(arrowright|right(?:ward|wards)?|w prawo|praw\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            matches.Add("ArrowRight");
        if (Regex.IsMatch(text, @"\b(arrowleft|left(?:ward|wards)?|w lewo|lew\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            matches.Add("ArrowLeft");
        if (Regex.IsMatch(text, @"\b(arrowup|up(?:ward|wards)?|w gór\w*|gór\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            matches.Add("ArrowUp");
        if (Regex.IsMatch(text, @"\b(arrowdown|down(?:ward|wards)?|w dół|dol\w*|dół)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            matches.Add("ArrowDown");

        if (matches.Count == 1)
        {
            label = matches[0];
            return true;
        }

        label = "";
        return false;
    }

    internal static bool TryNormalizeDirectionalLabel(
        string? value,
        out string label)
    {
        if (TurnInput.Direction(value) is string direction)
        {
            label = direction;
            return true;
        }
        var text = value ?? "";
        if (Regex.IsMatch(text, @"\b(arrowright|right(?:ward|wards)?|w prawo|praw\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            label = "ArrowRight";
        else if (Regex.IsMatch(text, @"\b(arrowleft|left(?:ward|wards)?|w lewo|lew\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            label = "ArrowLeft";
        else if (Regex.IsMatch(text, @"\b(arrowup|up(?:ward|wards)?|w gór\w*|gór\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            label = "ArrowUp";
        else if (Regex.IsMatch(text, @"\b(arrowdown|down(?:ward|wards)?|w dół|dol\w*|dół)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            label = "ArrowDown";
        else
        {
            label = "";
            return false;
        }

        return true;
    }

    internal static bool IsWellFormedSafeBatchedAction(ActionDto action)
    {
        if (!CaptureResolvedAction(action, null).IsValid)
            return false;

        return action.Type switch
        {
            "keys" => action.Keys is { Length: > 0 } &&
                      action.Keys.All(key => !string.IsNullOrWhiteSpace(key)),
            "type_text" or "paste_text" => !string.IsNullOrEmpty(action.Text),
            "wait" => action.WaitSeconds is > 0,
            "run_command" => !string.IsNullOrWhiteSpace(action.Command),
            "open_url" => !string.IsNullOrWhiteSpace(action.Url),
            "launch_app" => !string.IsNullOrWhiteSpace(action.App),
            "drag_path" => IsSafeBatchedDrawGesture(action),
            _ => false
        };
    }

    internal static bool IsSafeBatchedDrawGesture(ActionDto action) =>
        Options.MouseEnabled &&
        Options.DirectClickWithoutAim &&
        action.Type == "drag_path" &&
        string.Equals(action.GestureKind, "draw", StringComparison.OrdinalIgnoreCase) &&
        action.Path is { Length: >= 2 } &&
        action.DurationMs is >= 100 &&
        action.DurationMs <= Options.MaxGestureDurationMs;

    internal static bool CanBatchWithoutObservation(
        ActionDto previous,
        ActionDto next)
    {
        if (!IsSafeBatchedAction(previous) || !IsSafeBatchedAction(next))
            return false;
        if (previous.Type == "wait")
            return false;
        if (next.Type == "wait")
            return true;

        if (previous.Type == "drag_path" || next.Type == "drag_path")
            return IsSafeBatchedDrawGesture(previous) &&
                   IsSafeBatchedDrawGesture(next);

        if (previous.Type is "open_url" or "launch_app" or "run_command" ||
            IsCommitKeys(previous))
        {
            return false;
        }

        if (next.Type is "type_text" or "paste_text")
        {
            return previous.Type is "type_text" or "paste_text" ||
                   previous.Type == "keys" && IsTextEntryPrelude(previous);
        }

        if (previous.Type == "keys" && IsFieldAdvanceKeys(previous) &&
            next.Type == "keys" && IsTextSelectionKeys(next))
        {
            return true;
        }

        return next.Type == "keys" &&
               previous.Type is "type_text" or "paste_text" &&
               (IsCommitKeys(next) || IsFieldAdvanceKeys(next));
    }

    internal static bool IsCommitKeys(ActionDto action)
    {
        if (action.Type != "keys" || action.Keys is null)
            return false;
        var signature = NormalizeBatchKeySignature(action.Keys);
        return signature is "enter" or "return";
    }

    internal static bool IsFieldAdvanceKeys(ActionDto action)
    {
        if (action.Type != "keys" || action.Keys is null)
            return false;
        var signature = NormalizeBatchKeySignature(action.Keys);
        return signature is "tab" or "shift+tab";
    }

    internal static bool IsTextSelectionKeys(ActionDto action)
    {
        if (action.Type != "keys" || action.Keys is null)
            return false;
        return NormalizeBatchKeySignature(action.Keys) == "ctrl+a";
    }

    internal static bool IsTextEntryPrelude(ActionDto action)
    {
        if (action.Type != "keys" || action.Keys is null)
            return false;
        var signature = NormalizeBatchKeySignature(action.Keys);
        return signature is "win" or "win+r" or "ctrl+l" or "ctrl+k" or
            "ctrl+e" or "ctrl+a" or "tab" or "shift+tab" or "f6";
    }

    static string NormalizeBatchKeySignature(IEnumerable<string> keys) =>
        KeyboardActionParser.Signature(keys.ToArray());

}
