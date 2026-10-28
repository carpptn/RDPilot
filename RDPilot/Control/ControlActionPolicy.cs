namespace RDPilot.Control;

internal static class ControlActionPolicy
{
    internal static string MouseButton(ActionDto action) => action.Type is "move" or "scroll" ? "" :
        action.Type == "focus_uia" ? "left" : NormalizeMouseButton(action.Button);
    internal static string NormalizeMouseButton(string? button) => button?.Trim().ToLowerInvariant() switch
    {
        "right" => "right",
        "middle" => "middle",
        _ => "left"
    };
    internal static bool IsMouseAction(ActionDto a) => a.Type is "move" or "click" or "double_click" or "drag_drop" or "drag_path" or "scroll" or "focus_uia" or "click_uia";
    internal static bool IsPointClickAction(ActionDto? a) => a?.Type is "click" or "double_click";
    internal static bool IsTextInputAttemptAction(ActionDto? a)
    {
        if (a?.Type is "paste_text" or "type_text")
            return true;
        if (a?.Type != "keys" || a.Keys is null || a.Keys.Length == 0)
            return false;
        try
        {
            return KeyboardActionParser.Parse(a.Keys).Any(stroke =>
                stroke.Key == "v" && stroke.Modifiers.Contains("ctrl") ||
                stroke.Key == "insert" && stroke.Modifiers.Contains("shift"));
        }
        catch (ArgumentException) { return false; }
    }

    internal static bool IsKnownAction(ActionDto? action) => !string.IsNullOrWhiteSpace(action?.Type) && KnownActionTypes.Contains(action.Type);
    internal static string BuildPrecisionHint(ResolvedActionSnapshot action, double delta, string reason)
    {
        var target = "";
        if (action.ScreenPoint is Point screenPoint)
        {
            var imagePoint = CurrentRun.CurrentScreenMap.ScreenToImagePoint(screenPoint.X, screenPoint.Y);
            target = $" Target was near SCREEN_SIZE ({imagePoint.X},{imagePoint.Y}).";
        }

        var deltaText = double.IsNaN(delta) ? "" : $" Screen delta={delta:0.####}.";
        return $"{reason}{target}{deltaText} Do not repeat the same point. For tiny controls, tree expanders, list rows, or menus, prefer request_crop/aim for that area or select the row and use keyboard expansion/activation such as ArrowRight or Enter.";
    }

    internal static string BuildTextInputHint(ActionDto action, double delta)
    {
        var deltaText = double.IsNaN(delta) ? "" : $" Screen delta={delta:0.####}.";
        var textLen = action.Text?.Length ?? 0;
        var textInfo = textLen > 0 ? $" Attempted {action.Type} with {textLen} character(s).": $" Attempted {action.Type}.";
        return $"Previous text input produced little or no visible progress.{deltaText}{textInfo} The target is likely not focused, not editable, or the caret is elsewhere. Do not repeat text input or paste shortcuts immediately; first make the editable field active through visible UI, use request_crop/aim if needed, or switch to a different visible UI route.";
    }

    internal static string BuildTextInputCooldownHint(int attempts)
    {
        var countText = attempts > 0 ? $" after {attempts} no-change attempt(s)": "";
        return $"Text input is temporarily blocked{countText} because it has not changed the screen. Do not choose paste_text, type_text, Ctrl+V, or Shift+Insert now. Use visible UI to establish focus/editability, open the correct editor/control, or choose a non-text navigation route first.";
    }

    internal static bool IsLocalObservationAction(ActionDto a) => a.Type is "aim" or "point" or "request_crop";
    internal static bool IsInspectionAction(ActionDto a) => a.Type is "point" or "request_crop";
    internal static bool IsStateChangingInteractionAction(ActionDto a) => ActionNeedsEffectObservation(a) && a.Type != "wait";
    internal static bool RequiresPrimaryTurnBasedRegion(ActionDto action)
    {
        if (!IsStateChangingInteractionAction(action) || action.Type is "click" or "double_click")
        {
            return false;
        }

        if (action.Type == "keys" && action.Keys is { Length: 1 })
        {
            var key = action.Keys[0]?.Trim().ToLowerInvariant();
            if (key is "space" or "spacebar" or "enter" or "return")
                return false;
        }

        return true;
    }

    internal static bool TryGetTurnBasedDirectionalSequenceLength(ActionDto action, out int length)
    {
        length = action.Keys?.Length ?? 0;
        if (action.Type != "keys" || length == 0)
            return false;
        return action.Keys!.All(key => TurnInput.Key(key) is not null);
    }

    internal static TurnNoEffectPolicy ResolveTurnNoEffectPolicy(bool immediateReactionObserved) => immediateReactionObserved ? new TurnNoEffectPolicy(false, false) : new TurnNoEffectPolicy(true, false);
    internal static bool CanExpandTurnKeyActionFromPlan(ActionDto action, int queuedFollowUpCount) => queuedFollowUpCount == 0 && action.Type == "keys" && action.Keys is { Length: > 0 };
    internal static bool ActionNeedsEffectObservation(ActionDto action) => action.Type is not ("aim" or "point" or "request_crop" or "move" or "done");
    internal static bool ShouldCheckPreRegionTurnActionFreshness(string actionPolicy, Rectangle? turnBasedInteractionRect, ActionDto action) => string.Equals(actionPolicy, "turn_based_interaction", StringComparison.Ordinal) &&
        turnBasedInteractionRect is null &&
        IsStateChangingInteractionAction(action) &&
        (IsSpatialPointerAction(action) ||
        action.Type is "drag_drop" or "drag_path" or "click_uia" or "focus_uia");
    internal static bool ShouldDiscardPreRegionTurnAction(
        ScreenObservationFrame promptFrame,
        ScreenObservationFrame immediateFrame,
        ScreenObservationFrame confirmationFrame,
        out double promptDelta,
        out double stabilityDelta)
    {
        promptDelta = Math.Max(
            ComputeImageDelta(promptFrame.GlobalFingerprint, immediateFrame.GlobalFingerprint),
            ComputeImageDelta(promptFrame.ActiveWindowFingerprint, immediateFrame.ActiveWindowFingerprint));
        stabilityDelta = Math.Max(
            ComputeImageDelta(immediateFrame.GlobalFingerprint, confirmationFrame.GlobalFingerprint),
            ComputeImageDelta(immediateFrame.ActiveWindowFingerprint, confirmationFrame.ActiveWindowFingerprint));
        return promptDelta >= 0.0035 && stabilityDelta <= Math.Max(0.0015, promptDelta * 0.35);
    }

    internal static bool IsExpectedContinuousIdle(string goalMode, ActionDto? previousAction, bool noChange) => noChange && string.Equals(goalMode, "continuous", StringComparison.Ordinal) &&
        previousAction?.Type == "wait";
    internal static bool ShouldRegisterImmediateNoEffectCooldown(ActionDto action, ObservationAssessment assessment, int repeatCount) => assessment.ActionOutcome == ActionOutcomeState.NoEffect &&
        Options.ActionRepeatCooldownSteps > 0 &&
        action.Type is not ("wait" or "done") &&
        !IsLocalObservationAction(action) &&
        !IsTextInputAttemptAction(action) &&
        (!string.Equals(assessment.ActionPolicy, "turn_based_interaction", StringComparison.Ordinal) ||
        !IsStateChangingInteractionAction(action) ||
        repeatCount > 0);
    internal static bool ShouldResetRejectedProposalLoop(ResolvedActionSnapshot? previousAction, bool noChange, bool expectedContinuousIdle) => previousAction is not null &&
        !noChange &&
        !expectedContinuousIdle &&
        !IsLocalObservationAction(previousAction.Action);
    internal static bool IsSpatialPointerAction(ActionDto action) => action.Type is "move" or "click" or "double_click";
    internal static bool IsPointerPositionOnlyAction(ActionDto a) => a.Type == "move";
    internal static bool ShouldDeferFocusedTextStagnation(ResolvedActionSnapshot? previousAction, bool observedNoChange, int priorNoChangeAttempts) => observedNoChange &&
        priorNoChangeAttempts == 0 &&
        previousAction is not null &&
        IsTextInputAttemptAction(previousAction.Action);
    internal static int DelayFor(ActionDto a) => a.Type switch
    {
        "open_url" => 500,
        "launch_app" => 500,
        "run_command" => 500,
        "paste_text" => 150,
        "move" => 120,
        "click" => 180,
        "double_click" => 250,
        "drag_drop" => 500,
        "drag_path" => 250,
        "keys" => 120,
        "hold_keys" => 80,
        "type_text" => 80,
        "scroll" => 80,
        "request_crop" => 80,
        "point" => 80,
        "aim" => 80,
        "wait" => 0, // wait time handled separately
        "done" => 80,
        _ => 120
    };
    internal static int EffectiveWaitSeconds(ActionDto a, out int requestedSeconds)
    {
        requestedSeconds = Math.Max(0, a.WaitSeconds ?? 1);
        return Options.MaxWaitSeconds > 0 ? Math.Min(requestedSeconds, Options.MaxWaitSeconds) : requestedSeconds;
    }
}
