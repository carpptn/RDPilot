namespace RDPilot.Diagnostics;

/// <summary>Resolves recorded actions without reading or changing live desktop state.</summary>
internal static class ReplayActionResolver
{
    internal static ReplayActionGeometry? Record(ResolvedActionSnapshot? snapshot) => snapshot is null ? null : new()
    {
        Source = snapshot.ScreenPoint,
        Destination = snapshot.DestinationScreenPoint,
        Path = snapshot.ScreenPath.Count <= 32 ? snapshot.ScreenPath.ToList() : Enumerable.Range(0, 32)
            .Select(index => snapshot.ScreenPath[(int)Math.Round(index * (snapshot.ScreenPath.Count - 1) / 31.0)]).ToList(),
        ObservationRegion = snapshot.ObservationRegion,
        SignatureHash = Hash(snapshot.IneffectiveSignature),
        SemanticFeatures = LoopSemanticClassifier.Action(snapshot),
        ValidationError = snapshot.ValidationError
    };

    internal static ResolvedActionSnapshot Resolve(LoopReplayFrame frame, int width, int height)
    {
        return Resolve(frame.PreviousAction ?? throw new InvalidOperationException("Replay action is missing."),
            frame.PreviousActionGeometry, width, height);
    }

    internal static ResolvedActionSnapshot Resolve(ActionDto action, ReplayActionGeometry? geometry, int width, int height)
    {
        Point? source = geometry?.Source;
        Point? destination = geometry?.Destination;
        IReadOnlyList<Point> path = geometry?.Path ?? [];
        var error = geometry?.ValidationError;
        if (geometry is null)
        {
            // Older corpora contain only action coordinates. Their pixel coordinates are
            // interpreted in the recorded screen. UIA indices have no recoverable geometry.
            if (action.Type is "click" or "double_click" or "move" or "drag_drop")
            {
                source = PointFrom(action.X, action.Y, action.XPx, action.YPx, action.BBox);
                if (source is null) error = "Missing or invalid replay source coordinates.";
            }
            if (action.Type == "drag_drop")
            {
                destination = PointFrom(action.ToX, action.ToY, action.ToXPx, action.ToYPx, action.ToBBox);
                if (destination is null) error = "Missing or invalid replay destination coordinates.";
                if (source is Point from && destination is Point to) path = new[] { from, to };
            }
            if (action.Type == "drag_path")
            {
                path = action.Path?.Select(p => new Point(p.XPx, p.YPx)).ToArray() ?? [];
                if (path.Count < 2 || path.Any(p => p.X < 0 || p.Y < 0 || p.X >= width || p.Y >= height))
                    error = "Invalid replay gesture path.";
                else { source = path[0]; destination = path[^1]; }
            }
            if (action.Type is "click_uia" or "focus_uia" && action.UiaIndex is null)
                error = "Missing replay uia_index.";
        }
        var tokens = string.Join(' ', Regex.Matches($"{action.Note} {action.GestureKind}".ToLowerInvariant(), @"[\p{L}\p{Nd}]{2,}")
            .Cast<Match>().Select(m => m.Value).Where(t => !int.TryParse(t, out _)).Distinct().Take(20));
        Rectangle? signatureRegion = null;
        if (action.Type is "aim" or "point" or "request_crop")
        {
            var box = action.Type == "aim" ? action.BBox ?? action.Crop : action.Crop ?? action.BBox;
            if (box is { Left: int left, Top: int top, Right: int right, Bottom: int bottom } &&
                PointFrom(null, null, null, null, box) is not null)
                signatureRegion = Rectangle.FromLTRB(left, top, right, bottom);
            else if (box is null && PointFrom(action.X, action.Y, action.XPx, action.YPx, null) is Point center)
            {
                var half = Options.FocusCropSize / 2;
                signatureRegion = Rectangle.FromLTRB(Math.Max(0, center.X - half), Math.Max(0, center.Y - half),
                    Math.Min(width, center.X - half + Options.FocusCropSize), Math.Min(height, center.Y - half + Options.FocusCropSize));
            }
        }
        var signature = string.IsNullOrEmpty(geometry?.SignatureHash)
            ? Hash(ActionSignatureService.Create(action, source, destination, path, signatureRegion, tokens)) : geometry.SignatureHash;
        return new ResolvedActionSnapshot(action, action.Type, $"{action.Type}:{signature}", source)
        {
            DestinationScreenPoint = destination,
            ScreenPath = path,
            ObservationRegion = geometry?.ObservationRegion,
            SemanticTokens = tokens,
            SemanticFeatures = geometry?.SemanticFeatures,
            ValidationError = error
        };

        Point? PointFrom(double? x, double? y, int? px, int? py, BBox? box)
        {
            if (box is not null)
            {
                if (box is not { Left: int left, Top: int top, Right: int right, Bottom: int bottom } ||
                    left < 0 || top < 0 || right <= left || bottom <= top || right > width || bottom > height) return null;
                return new Point(left + (right - left) / 2, top + (bottom - top) / 2);
            }
            if (px.HasValue && py.HasValue)
                return px >= 0 && px < width && py >= 0 && py < height ? new Point(px.Value, py.Value) : null;
            if (x.HasValue && y.HasValue && double.IsFinite(x.Value) && double.IsFinite(y.Value) &&
                x >= 0 && x <= 1 && y >= 0 && y <= 1)
                return new Point((int)Math.Round(x.Value * (width - 1)), (int)Math.Round(y.Value * (height - 1)));
            return null;
        }
    }

    static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
}
