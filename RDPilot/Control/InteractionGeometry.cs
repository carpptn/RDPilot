namespace RDPilot.Control;

/// <summary>Physical geometry; independent of image resizing and prompt formatting.</summary>
internal sealed record InteractionGeometry(Rectangle Desktop, Rectangle? Window)
{
    internal static InteractionGeometry Capture()
    {
        var (x, y, width, height) = GetPrimaryScreen();
        return new(new(x, y, width, height), ScreenshotService.ActiveWindowRectangleProvider());
    }
}

internal sealed record TurnInteractionContext(string Identity, InteractionGeometry Geometry)
{
    internal static TurnInteractionContext Capture(UiPromptContext context) =>
        new(TurnBasedContextKey(context), InteractionGeometry.Capture());

    internal bool GeometryIsCurrent => Geometry == InteractionGeometry.Capture();
}
