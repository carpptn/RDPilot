namespace RDPilot.Models;

/// <summary>Content-free features needed to replay semantic loop decisions.</summary>
internal sealed record LoopActionFeatures(
    bool RecurringObservation,
    bool InspectionEntry,
    bool InspectionReturn,
    bool Placement,
    bool RasterInteraction);

internal sealed record LoopContextFeatures(bool RasterSurface, bool DocumentSurface, bool RasterInteraction);