namespace RDPilot.Control;

/// <summary>Remembers physical control positions and projects them into each request's image.</summary>
internal sealed class DirectionalControlMemory
{
    sealed record Control(Point Position, string Type, string? Button, double? Confidence, string? Note);
    readonly Dictionary<string, Control> controls = new(StringComparer.Ordinal);
    InteractionGeometry? geometry;

    internal void Clear() { controls.Clear(); geometry = null; }

    void RefreshGeometry()
    {
        var current = InteractionGeometry.Capture();
        if (geometry != current) controls.Clear();
        geometry = current;
    }

    internal void Remember(ActionDto action, string direction)
        => Remember(CaptureResolvedAction(action, null), direction);

    internal void Remember(ResolvedActionSnapshot resolved, string direction)
    {
        RefreshGeometry();
        var action = resolved.Action;
        if (action.Type is not ("click" or "double_click") ||
            !BatchPlanner.TryNormalizeDirectionalLabel(direction, out var label)) return;
        if (!resolved.IsValid || resolved.ScreenPoint is not Point point) return;
        controls[label] = new(point, action.Type, action.Button, action.Confidence, action.Note);
    }

    internal ActionDto? Recall(string direction)
    {
        RefreshGeometry();
        if (!controls.TryGetValue(direction, out var control)) return null;
        var point = CurrentRun.CurrentScreenMap.ScreenToImagePoint(control.Position.X, control.Position.Y);
        return new()
        {
            Type = control.Type, XPx = point.X, YPx = point.Y, Button = control.Button,
            Confidence = control.Confidence, Note = control.Note, ResolvedTurnInputLabel = direction
        };
    }

    internal bool TryResolve(ResolvedActionSnapshot resolved, out string label)
    {
        label = "";
        RefreshGeometry();
        var action = resolved.Action;
        if (action.Type is not ("click" or "double_click")) return false;
        if (!resolved.IsValid || resolved.ScreenPoint is not Point screenPoint) return false;
        var point = CurrentRun.CurrentScreenMap.ScreenToImagePoint(screenPoint.X, screenPoint.Y);
        var nearest = controls.Select(entry =>
        {
            var candidate = CurrentRun.CurrentScreenMap.ScreenToImagePoint(entry.Value.Position.X, entry.Value.Position.Y);
            var dx = (long)candidate.X - point.X;
            var dy = (long)candidate.Y - point.Y;
            return (entry.Key, Distance: dx * dx + dy * dy);
        }).Where(entry => entry.Distance <= 8 * 8).OrderBy(entry => entry.Distance).Take(2).ToArray();
        if (nearest.Length == 0 || nearest.Length > 1 && nearest[0].Distance == nearest[1].Distance) return false;
        label = nearest[0].Key;
        action.ResolvedTurnInputLabel = label;
        return true;
    }
}
