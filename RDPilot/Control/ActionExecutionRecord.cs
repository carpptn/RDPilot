namespace RDPilot.Control;

/// <summary>Evidence for one successfully sent input, whether standalone or batched.</summary>
internal sealed record ActionExecutionRecord(
    ResolvedActionSnapshot Action,
    ScreenObservationFrame Before,
    byte[]? LocalBefore = null,
    ObservedActionResult? Observation = null,
    UiPromptContext? BeforeContext = null);
