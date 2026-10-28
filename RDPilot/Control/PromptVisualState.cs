namespace RDPilot.Control;

/// <summary>Tracks evidence actually used by the model and by the last UIA scan independently of action outcomes.</summary>
internal sealed class PromptVisualState
{
    sealed record Evidence(string Image, Rectangle Bounds, UiPromptContext Context)
    {
        internal static Evidence From(ControlStepState turn) => new(turn.dataUrl, turn.observationFrame.ScreenBounds, turn.promptContext);
    }
    Evidence? modelEvidence;
    string? modelResponseId;
    Evidence? uiaEvidence;
    IReadOnlyList<UiaTarget>? uiaTargets;

    internal bool CanOmitImage(ControlStepState turn) =>
        turn.previousResponseIdForRequest is not null &&
        turn.previousResponseIdForRequest == modelResponseId &&
        modelEvidence == Evidence.From(turn);

    internal bool CanReuseTargets(ControlStepState turn, IReadOnlyList<UiaTarget> targets) =>
        targets.Count > 0 && ReferenceEquals(targets, uiaTargets) && uiaEvidence == Evidence.From(turn);

    internal void RecordTargets(ControlStepState turn, IReadOnlyList<UiaTarget> targets)
    {
        uiaEvidence = Evidence.From(turn);
        uiaTargets = targets;
    }

    internal void RecordResponse(ControlStepState turn, string? responseId, bool imageAttached)
    {
        // Failed requests cannot advance the evidence associated with an existing response chain.
        if (string.IsNullOrWhiteSpace(responseId)) return;
        if (imageAttached) modelEvidence = Evidence.From(turn);
        modelResponseId = responseId;
    }
}
