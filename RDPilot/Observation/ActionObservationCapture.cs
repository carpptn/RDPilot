namespace RDPilot.Observation;

/// <summary>Captures semantic context at an input boundary, including current UIA focus.</summary>
internal static class ActionObservationCapture
{
    internal static UiPromptContext Context(ScreenObservationFrame frame) =>
        CaptureUiPromptContext(Options.IncludeFocusUia ? GetFocusedUiaSummary() : null,
            frame.ScreenBounds.Width, frame.ScreenBounds.Height);
}