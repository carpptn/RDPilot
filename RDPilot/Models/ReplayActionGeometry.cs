namespace RDPilot.Models;

/// <summary>Recorded physical geometry; replay never needs the current desktop or UIA tree.</summary>
internal sealed class ReplayActionGeometry
{
    public Point? Source { get; set; }
    public Point? Destination { get; set; }
    public List<Point> Path { get; set; } = [];
    public Rectangle? ObservationRegion { get; set; }
    public string SignatureHash { get; set; } = "";
    public string? ValidationError { get; set; }
    public LoopActionFeatures? SemanticFeatures { get; set; }
}
