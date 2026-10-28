namespace RDPilot.Control;

/// <summary>Run-level restrictions checked for each input, including batch follow-ups.</summary>
internal static class ActionExecutionPolicy
{
    internal static string? BlockReason(ControlRunState run, int step, ResolvedActionSnapshot snapshot, string policy)
    {
        if (!snapshot.IsValid) return snapshot.ValidationError;
        if (policy == "turn_based_interaction" && IsStateChangingInteractionAction(snapshot.Action) &&
            run.turnBasedInteractionContext is { GeometryIsCurrent: false })
            return "interaction window or desktop geometry changed while the action was being planned";
        if (IsMouseAction(snapshot.Action) && !Options.MouseEnabled)
            return "mouse action was blocked because mouse input is disabled";
        if (run.observationActionGuard.TryGetBlockReason(snapshot, Options.MaxConsecutiveInspectionActions,
            policy == "turn_based_interaction", out var inspectionReason))
            return inspectionReason;
        if (IsActionOnCooldown(snapshot, step, run.actionCooldownUntilStep, run.spatialActionCooldowns, out var until))
        {
            if (IsPointClickAction(snapshot.Action))
            {
                run.lastPrecisionHint = BuildPrecisionHint(snapshot, double.NaN, "A repeated click was blocked after no visible effect.");
                run.lastPrecisionHintExpiresAfterStep = step + 1;
            }
            return $"action temporarily blocked after no visible effect until step {until}: {snapshot.Description}";
        }
        if (IsTextInputAttemptAction(snapshot.Action) && run.textInputCooldownUntilStep >= step)
        {
            run.lastTextInputHint = BuildTextInputCooldownHint(run.textInputNoChangeAttempts);
            run.lastTextInputHintExpiresAfterStep = step + 3;
            return $"text input temporarily blocked after repeated no visible effect: {snapshot.Description}";
        }
        return null;
    }

    internal static void EnsureAllowed(ControlRunState run, int step, ResolvedActionSnapshot snapshot, string policy)
    {
        var reason = BlockReason(run, step, snapshot, policy);
        if (reason is not null) throw new InvalidOperationException(reason);
    }
}
