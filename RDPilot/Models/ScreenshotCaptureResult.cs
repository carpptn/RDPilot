namespace RDPilot.Models;

/// <summary>Encoded screenshot, optional crops and the clean observation captured alongside them.</summary>
internal sealed record ScreenshotCaptureResult
{
    public required string DataUrl { get; init; }
    public string? SavedPath { get; init; }

    // Physical desktop dimensions and the encoded overview dimensions may differ.
    public required int ScreenWidth { get; init; }
    public required int ScreenHeight { get; init; }
    public required int ImageWidth { get; init; }
    public required int ImageHeight { get; init; }

    // Crop rectangles use physical desktop coordinates, including any nonzero origin.
    public string? FocusDataUrl { get; init; }
    public Rectangle? FocusRect { get; init; }
    public Rectangle? FocusUiaRect { get; init; }
    public string? FocusUiaSummary { get; init; }
    public string? FocusUiaDataUrl { get; init; }
    public string? FocusUiaImagePath { get; init; }

    // Fingerprints describe the captured screen before debug overlays are drawn.
    public required byte[] DeltaFingerprint { get; init; }
    public required byte[] ActiveWindowFingerprint { get; init; }
    public required ScreenObservationFrame ObservationFrame { get; init; }
}
