namespace RDPilot.Control;

/// <summary>Labels image and physical desktop coordinates explicitly in question answers.</summary>
internal static class QaAnswerPresentation
{
    internal static void WriteLocation(QaLocateDto answer, ScreenCoordinateMapper map)
    {
        if (answer.X is double x && answer.Y is double y)
        {
            var desktop = map.NormalizedToScreenPoint(x, y);
            Console.WriteLine($"Location (desktop): {desktop.X}:{desktop.Y}");
        }
        else if (answer.XPx is int px && answer.YPx is int py)
        {
            var desktop = map.ImageToScreenPoint(px, py);
            Console.WriteLine($"Location (image): {px}:{py}; desktop: {desktop.X}:{desktop.Y}");
        }
        if (answer.BBox is { Left: int left, Top: int top, Right: int right, Bottom: int bottom })
        {
            var desktop = map.ImageToScreenRect(Rectangle.FromLTRB(left, top, right, bottom));
            Console.WriteLine($"bbox (image)=({left},{top})-({right},{bottom}); desktop=({desktop.Left},{desktop.Top})-({desktop.Right},{desktop.Bottom})");
        }
    }
}
