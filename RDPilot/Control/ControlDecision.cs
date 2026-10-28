namespace RDPilot.Control;

/// <summary>One model decision and the follow-ups that are valid only after its primary action.</summary>
internal sealed class ControlDecision(ActionDto action, IEnumerable<ActionDto>? followUps = null)
{
    internal ActionDto Action { get; set; } = action;
    internal Queue<ActionDto> FollowUps { get; } = new(followUps ?? []);
}
