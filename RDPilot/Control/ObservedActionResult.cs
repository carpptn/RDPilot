namespace RDPilot.Control;

/// <summary>Evidence belonging to exactly one input in an observed sequence.</summary>
internal sealed record ObservedActionResult(
    ResolvedActionSnapshot Action,
    ScreenObservationFrame Before,
    ScreenObservationFrame After,
    TurnBasedTransitionTracker.BatchStepAssessment Assessment)
{
    internal ObservationAssessment Reconcile(ObservationAssessment current, ScreenObservationFrame latest)
    {
        // A delayed reaction must remain observable. Reuse the tracker verdict only
        // while the screen still matches the frame on which that verdict was based.
        var threshold = Math.Max(0.000001, current.ChangeThreshold);
        if (!Assessment.NoEffect || After.ScreenBounds != latest.ScreenBounds ||
            ComputeImageDelta(After.GlobalFingerprint, latest.GlobalFingerprint) >= threshold ||
            ComputeImageDelta(After.ActiveWindowFingerprint, latest.ActiveWindowFingerprint) >= threshold ||
            After.DetailWidth != latest.DetailWidth || After.DetailHeight != latest.DetailHeight ||
            After.DetailColorFingerprint.Length > 0 &&
            ComputeImageDelta(After.DetailColorFingerprint, latest.DetailColorFingerprint) >= threshold)
            return current;
        return current with
        {
            ActionOutcome = ActionOutcomeState.NoEffect,
            GoalProgress = GoalProgressState.NoProgress,
            Evidence = current.Evidence + "; observed_batch_input=no_effect"
        };
    }
}
