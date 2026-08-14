internal static partial class RDPilotApplication
{
    /// <summary>
    /// Maps image coordinates and executes constrained mouse, keyboard, clipboard, and local actions.
    /// </summary>
    internal static class DesktopInputService
    {
            internal readonly record struct AbsoluteMouseMoveData(int Dx, int Dy, uint Flags);
            internal sealed record MouseDragPlan(
                Point[] Moves,
                int EffectiveDurationMs,
                int InitialHoldMs);

            const uint MouseEventMove = 0x0001;
            const uint MouseEventMoveNoCoalesce = 0x2000;
            const uint MouseEventVirtualDesk = 0x4000;
            const uint MouseEventAbsolute = 0x8000;

            // Controlled desktop region. Multi-monitor capture is opt-in because
            // it increases image size and may expose unrelated secondary screens.
            internal static (int X, int Y, int W, int H) GetPrimaryScreen()
            {
                if (MultiMonitorEnabled)
                {
                    var x = GetSystemMetrics((int)SystemMetric.SM_XVIRTUALSCREEN);
                    var y = GetSystemMetrics((int)SystemMetric.SM_YVIRTUALSCREEN);
                    var width = GetSystemMetrics((int)SystemMetric.SM_CXVIRTUALSCREEN);
                    var height = GetSystemMetrics((int)SystemMetric.SM_CYVIRTUALSCREEN);
                    if (width > 0 && height > 0)
                        return (x, y, width, height);
                }

                int w = GetSystemMetrics((int)SystemMetric.SM_CXSCREEN);
                int h = GetSystemMetrics((int)SystemMetric.SM_CYSCREEN);
                return (0, 0, w, h);
            }
            [DllImport("user32.dll")] internal static extern int GetSystemMetrics(int nIndex);
        
            // --- Cursor pos (PRIMARY) ---
            [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT lpPoint);
            [StructLayout(LayoutKind.Sequential)] struct POINT { public int X; public int Y; }
            internal static (int X, int Y, double Nx, double Ny) GetCursorPositionInPrimary()
            {
                if (!GetCursorPos(out var p)) return (0, 0, 0, 0);
                var (vx, vy, vw, vh) = GetPrimaryScreen();
                int relX = Math.Max(0, Math.Min(vw - 1, p.X - vx));
                int relY = Math.Max(0, Math.Min(vh - 1, p.Y - vy));
                double nx = vw > 1 ? (double)relX / (vw - 1) : 0.0;
                double ny = vh > 1 ? (double)relY / (vh - 1) : 0.0;
                return (vx + relX, vy + relY, nx, ny);
            }
        
            internal static void SetCurrentScreenMap(int screenW, int screenH, int imageW, int imageH)
            {
                var (screenX, screenY, _, _) = GetPrimaryScreen();
                CurrentScreenMap = ScreenCoordinateMapper.Create(
                    screenX,
                    screenY,
                    screenW,
                    screenH,
                    imageW,
                    imageH);
            }

            internal static void SetCurrentScreenMap(
                int screenX,
                int screenY,
                int screenW,
                int screenH,
                int imageW,
                int imageH) =>
                CurrentScreenMap = ScreenCoordinateMapper.Create(
                    screenX,
                    screenY,
                    screenW,
                    screenH,
                    imageW,
                    imageH);
        
            internal static Rectangle? ScreenRectToImage(Rectangle? rect) =>
                rect.HasValue ? CurrentScreenMap.ScreenToImageRect(rect.Value) : null;
        
            internal static (int X, int Y, double Nx, double Ny) CursorToImageCoordinates(int screenX, int screenY)
            {
                var (ix, iy) = CurrentScreenMap.ScreenToImagePoint(screenX, screenY);
                var nx = CurrentScreenMap.ImageW > 1 ? ix / (double)(CurrentScreenMap.ImageW - 1) : 0.0;
                var ny = CurrentScreenMap.ImageH > 1 ? iy / (double)(CurrentScreenMap.ImageH - 1) : 0.0;
                return (ix, iy, nx, ny);
            }
        
            internal static string FormatImageRect(Rectangle r) =>
                $"[{r.Left},{r.Top}]–[{r.Right},{r.Bottom}]";
        
            internal static string FormatActionRect(BBox box)
            {
                var imageRect = RectFromBBox(box);
                var screenRect = CurrentScreenMap.ImageToScreenRect(imageRect);
                return CurrentScreenMap.RequiresMapping
                    ? $"bbox={FormatImageRect(imageRect)}→screen{FormatImageRect(screenRect)}"
                    : $"bbox=({screenRect.Left},{screenRect.Top})–({screenRect.Right},{screenRect.Bottom})";
            }
        
            internal static string FormatActionPoint(int imageX, int imageY)
            {
                var screenPoint = CurrentScreenMap.ImageToScreenPoint(imageX, imageY);
                return CurrentScreenMap.RequiresMapping
                    ? $"({imageX},{imageY})→screen({screenPoint.X},{screenPoint.Y})"
                    : $"({screenPoint.X},{screenPoint.Y})";
            }
        
            internal static string Describe(ActionDto a)
            {
                if (a is null) return "null";
        
                if (a.Type == "open_url") return $"open_url \"{a.Url}\"";
                if (a.Type == "launch_app") return $"launch_app \"{a.App}\"";
                if (a.Type == "run_command") return $"run_command \"{Tail(a.Command ?? "", 120)}\"";
                if (a.Type == "paste_text") return $"paste_text ({a.Text?.Length ?? 0} chars)";
                if (a.Type == "focus_uia") return $"focus_uia #{a.UiaIndex}";
                if (a.Type == "click_uia") return $"click_uia #{a.UiaIndex}";
                if (a.Type == "drag_drop")
                {
                    if (!HasExplicitPoint(a) || !HasExplicitDropPoint(a))
                        return "drag_drop (missing source or destination)";
                    var source = ResolvePoint(a);
                    var destination = ResolveDropPoint(a);
                    return $"drag_drop ({source.X},{source.Y})→({destination.X},{destination.Y}) {a.Button ?? "left"} {EffectiveDragDurationMs(a)}ms";
                }
        
                // request_crop / point
                if (a.Type is "request_crop" or "point")
                {
                    string prefix = a.Type == "point" ? "point" : "request_crop";
                    if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                        return $"{prefix} {FormatActionRect(a.Crop)}";
                    if (a.XPx is not null && a.YPx is not null)
                        return $"{prefix} center={FormatActionPoint(a.XPx.Value, a.YPx.Value)}";
                    if (a.X is not null && a.Y is not null)
                    {
                        var (pxN, pyN) = NormalizedToPixels(a.X.Value, a.Y.Value);
                        return $"{prefix} center≈({pxN},{pyN})";
                    }
                    return $"{prefix} (missing parameters)";
                }
        
                // aim
                if (a.Type == "aim")
                {
                    if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                        return $"aim {FormatActionRect(a.BBox)}";
                    if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                        return $"aim(crop) {FormatActionRect(a.Crop)}";
                    if (a.XPx is not null && a.YPx is not null)
                        return $"aim center={FormatActionPoint(a.XPx.Value, a.YPx.Value)}";
                    if (a.X is not null && a.Y is not null)
                    {
                        var (pxN, pyN) = NormalizedToPixels(a.X.Value, a.Y.Value);
                        return $"aim center≈({pxN},{pyN})";
                    }
                    return "aim (missing parameters)";
                }
        
                // move / click / double_click
                if (a.Type is "move" or "click" or "double_click")
                {
                    // by bbox
                    if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                    {
                        var p = ResolvePoint(a);
                        string sfxB = a.Type == "click" ? $" {a.Button ?? "left"}"
                                    : a.Type == "double_click" ? " (double)" : "";
                        return $"{a.Type} bbox→({p.X},{p.Y}){sfxB}";
                    }
        
                    // by explicit coordinates
                    if (a.XPx is not null && a.YPx is not null)
                    {
                        string sfxP = a.Type == "click" ? $" {a.Button ?? "left"}"
                                    : a.Type == "double_click" ? " (double)" : "";
                        return $"{a.Type} {FormatActionPoint(a.XPx.Value, a.YPx.Value)}{sfxP}";
                    }
                    if (a.X is not null && a.Y is not null)
                    {
                        var (px, py) = NormalizedToPixels(a.X.Value, a.Y.Value);
                        string sfxN = a.Type == "click" ? $" {a.Button ?? "left"}"
                                    : a.Type == "double_click" ? " (double)" : "";
                        return $"{a.Type} ({px},{py}){sfxN}";
                    }
        
                    // missing coords
                    string sfx = a.Type == "click" ? $" {a.Button ?? "left"}"
                               : a.Type == "double_click" ? " (double)" : "";
                    return $"{a.Type} (coords: missing){sfx}";
                }
        
                // scroll
                if (a.Type == "scroll")
                {
                    int dy = a.ScrollDy ?? 0;
                    return $"scroll dy={dy}";
                }
        
                if (a.Type == "drag_path")
                {
                    var points = a.Path ?? Array.Empty<GesturePointDto>();
                    return $"drag_path kind={a.GestureKind ?? "other"} points={points.Length} duration={EffectiveGestureDurationMs(a)}ms";
                }

                // keys / type_text / wait / done
                if (a.Type == "keys") return $"keys [{string.Join("+", a.Keys ?? Array.Empty<string>())}]";
                if (a.Type == "hold_keys") return $"hold_keys [{string.Join("+", a.Keys ?? Array.Empty<string>())}] {EffectiveKeyHoldDurationMs(a)}ms";
                if (a.Type == "type_text") return $"type_text \"{Tail(a.Text ?? "", 120)}\"";
                if (a.Type == "wait")
                {
                    var secs = EffectiveWaitSeconds(a, out var requested);
                    return secs == requested ? $"wait {secs}s" : $"wait {secs}s (requested {requested}s)";
                }
                if (a.Type == "done") return "done";
        
                return $"unknown {a.Type}";
            }
        
            internal static string ActionSignature(ActionDto a)
            {
                if (a == null || string.IsNullOrEmpty(a.Type)) return "null";
                string t = a.Type.ToLowerInvariant();
        
                if (t == "open_url") return $"open_url:{a.Url}";
                if (t == "launch_app") return $"launch_app:{a.App}";
                if (t == "run_command")
                    return $"run_command:{StableSensitiveSignature(a.Command)}:{StableActionContext(a)}";
                if (t == "paste_text")
                    return $"paste_text:{StableSensitiveSignature(a.Text)}:{StableActionContext(a)}";
                if (t is "focus_uia" or "click_uia") return $"{t}:{a.UiaIndex}";
                if (t == "drag_drop")
                {
                    if (!HasExplicitPoint(a) || !HasExplicitDropPoint(a))
                        return "drag_drop:missing";
                    var source = ResolvePoint(a);
                    var destination = ResolveDropPoint(a);
                    var cluster = Math.Max(16, IneffectiveMouseClusterPx);
                    return $"drag_drop:{source.X / cluster},{source.Y / cluster}->{destination.X / cluster},{destination.Y / cluster}";
                }
                if (t == "drag_path")
                {
                    if (a.Path is null || a.Path.Length < 2)
                        return "drag_path:missing";
                    var path = ResolveGesturePath(a);
                    var cluster = Math.Max(8, IneffectiveMouseClusterPx / 2);
                    var sampled = path
                        .Where((_, index) => index == 0 || index == path.Count - 1 || index % Math.Max(1, path.Count / 10) == 0)
                        .Take(12)
                        .Select(point => $"{point.X / cluster},{point.Y / cluster}");
                    return $"drag_path:{a.GestureKind ?? "other"}:{string.Join(';', sampled)}";
                }
                if (t == "keys") return $"keys:{string.Join("+", (a.Keys ?? Array.Empty<string>()).Select(k => k.ToLowerInvariant()))}";
                if (t == "hold_keys") return $"hold_keys:{string.Join("+", (a.Keys ?? Array.Empty<string>()).Select(k => k.ToLowerInvariant()))}:{EffectiveKeyHoldDurationMs(a)}";
                if (t == "type_text")
                    return $"type_text:{StableSensitiveSignature(a.Text)}:{StableActionContext(a)}";
                if (t == "scroll") return $"scroll:{a.ScrollDy ?? 0}";
                if (t == "wait") return $"wait:{EffectiveWaitSeconds(a, out _)}";
        
                if (t is "move" or "click" or "double_click")
                {
                    int cx, cy;
                    var p = ResolvePoint(a);
                    cx = p.X; cy = p.Y;
                    cx = (cx / 16) * 16;
                    cy = (cy / 16) * 16;
                    return $"{t}:{cx},{cy}";
                }
        
                if (t == "aim")
                {
                    var rect = ResolveAimRect(a);
                    if (rect is Rectangle r)
                    {
                        int cx = (r.Left + r.Right) / 2;
                        int cy = (r.Top + r.Bottom) / 2;
                        cx = (cx / 16) * 16;
                        cy = (cy / 16) * 16;
                        return $"aim:{cx},{cy}";
                    }
                    return "aim";
                }
        
                if (t is "request_crop" or "point")
                {
                    var rect = ResolveCropRect(a);
                    if (rect is Rectangle r)
                    {
                        int cx = ((r.Left + r.Right) / 2 / 16) * 16;
                        int cy = ((r.Top + r.Bottom) / 2 / 16) * 16;
                        return $"{t}:{cx},{cy}";
                    }
                    return t;
                }
        
                return t;
            }

            static string StableSensitiveSignature(string? value)
            {
                if (string.IsNullOrEmpty(value))
                    return "empty";
                var hash = SHA256.HashData(Encoding.UTF8.GetBytes(value));
                return Convert.ToHexString(hash.AsSpan(0, 6)).ToLowerInvariant();
            }

            static string StableActionContext(ActionDto action)
            {
                var tokens = ActionSemanticTokens(action);
                return string.IsNullOrWhiteSpace(tokens)
                    ? "unspecified-target"
                    : StableSensitiveSignature(tokens);
            }
        
            internal static string IneffectiveActionSignature(ActionDto a)
            {
                if (a == null || string.IsNullOrEmpty(a.Type))
                    return "null";
                return ActionSignature(a);
            }

            internal static ResolvedActionSnapshot CaptureResolvedAction(ActionDto action, Rectangle? clickContextRect)
            {
                Point? screenPoint = null;
                Point? destinationScreenPoint = null;
                string? validationError = ValidateActionCoordinates(action);
                if (validationError is null && action.Type is ("click" or "double_click"))
                {
                    var p = ResolveClickPoint(action, clickContextRect, logAdjustment: false);
                    screenPoint = new Point(p.X, p.Y);
                }
                else if (validationError is null && action.Type == "move")
                {
                    var p = ResolvePoint(action);
                    screenPoint = new Point(p.X, p.Y);
                }
                else if (validationError is null && action.Type == "drag_drop")
                {
                    var source = ResolvePoint(action);
                    var destination = ResolveDropPoint(action);
                    screenPoint = new Point(source.X, source.Y);
                    destinationScreenPoint = new Point(destination.X, destination.Y);
                }
                else if (validationError is null && action.Type == "drag_path")
                {
                    var path = ResolveGesturePath(action);
                    screenPoint = path[0];
                    destinationScreenPoint = path[^1];
                }
                else if (validationError is null && action.Type is "focus_uia" or "click_uia")
                {
                    var target = ResolveUiaTarget(action);
                    screenPoint = new Point(target.CenterX, target.CenterY);
                }

                IReadOnlyList<Point> screenPath = Array.Empty<Point>();
                if (validationError is null && action.Type == "drag_path")
                    screenPath = ResolveGesturePath(action);
                else if (validationError is null && action.Type == "drag_drop" &&
                         screenPoint is Point dragSource &&
                         destinationScreenPoint is Point dragDestination)
                    screenPath = new[] { dragSource, dragDestination };

                var observationRegion = validationError is null
                    ? BuildActionObservationRegion(action, screenPoint, screenPath)
                    : null;

                string description;
                string signature;
                try { description = Describe(action); }
                catch { description = $"{action.Type} (invalid parameters)"; }
                try { signature = validationError is null ? IneffectiveActionSignature(action) : $"{action.Type}:invalid"; }
                catch { signature = $"{action.Type}:invalid"; }

                return new ResolvedActionSnapshot(
                    action,
                    description,
                    signature,
                    screenPoint)
                {
                    DestinationScreenPoint = destinationScreenPoint,
                    ScreenPath = screenPath,
                    ObservationRegion = observationRegion,
                    SemanticTokens = ActionSemanticTokens(action),
                    ValidationError = validationError
                };
            }

            internal static string? ValidateActionCoordinates(ActionDto action)
            {
                if (action.BBox is not null && !IsValidBox(action.BBox))
                    return "bbox must have complete coordinates with right > left and bottom > top.";
                if (action.ToBBox is not null && !IsValidBox(action.ToBBox))
                    return "to_bbox must have complete coordinates with right > left and bottom > top.";
                if (action.Crop is not null && !IsValidBox(action.Crop))
                    return "crop must have complete coordinates with right > left and bottom > top.";
                if (action.BBox is not null && !BoxFitsCurrentImage(action.BBox))
                    return "bbox must stay inside the current screenshot.";
                if (action.ToBBox is not null && !BoxFitsCurrentImage(action.ToBBox))
                    return "to_bbox must stay inside the current screenshot.";
                if (action.Crop is not null && !BoxFitsCurrentImage(action.Crop))
                    return "crop must stay inside the current screenshot.";

                var sourceNormalizedError = ValidateNormalizedPair(action.X, action.Y, "x/y");
                if (sourceNormalizedError is not null)
                    return sourceNormalizedError;
                var destinationNormalizedError = ValidateNormalizedPair(action.ToX, action.ToY, "to_x/to_y");
                if (destinationNormalizedError is not null)
                    return destinationNormalizedError;
                var sourcePixelError = ValidatePixelPair(
                    action.XPx,
                    action.YPx,
                    CurrentScreenMap.ImageW,
                    CurrentScreenMap.ImageH,
                    "x_px/y_px");
                if (sourcePixelError is not null)
                    return sourcePixelError;
                var destinationPixelError = ValidatePixelPair(
                    action.ToXPx,
                    action.ToYPx,
                    CurrentScreenMap.ImageW,
                    CurrentScreenMap.ImageH,
                    "to_x_px/to_y_px");
                if (destinationPixelError is not null)
                    return destinationPixelError;

                if (action.Type is "click" or "double_click" or "move" &&
                    !HasExplicitPoint(action))
                {
                    return $"{action.Type} requires bbox, x_px/y_px, or x/y.";
                }

                if (action.Type == "drag_drop")
                {
                    if (!HasExplicitPoint(action))
                        return "drag_drop requires a valid source bbox, x_px/y_px, or x/y.";
                    if (!HasExplicitDropPoint(action))
                        return "drag_drop requires a valid destination to_bbox, to_x_px/to_y_px, or to_x/to_y.";
                    var source = ResolvePoint(action);
                    var destination = ResolveDropPoint(action);
                    var dx = (long)source.X - destination.X;
                    var dy = (long)source.Y - destination.Y;
                    if (dx * dx + dy * dy < 16)
                        return "drag_drop source and destination are effectively identical.";
                }

                if (action.Type == "drag_path")
                {
                    if (action.Path is null || action.Path.Length < 2)
                        return "drag_path requires at least two path points.";
                    if (action.Path.Length > MaxGesturePathPoints)
                        return $"drag_path accepts at most {MaxGesturePathPoints} points.";
                    foreach (var point in action.Path)
                    {
                        if (point is null)
                            return "drag_path points cannot be null.";
                        if (point.XPx < 0 || point.YPx < 0 ||
                            point.XPx >= CurrentScreenMap.ImageW ||
                            point.YPx >= CurrentScreenMap.ImageH)
                        {
                            return "drag_path points must stay inside the current screenshot.";
                        }
                    }
                    var path = ResolveGesturePath(action);
                    double totalLength = 0;
                    for (var index = 1; index < path.Count; index++)
                        totalLength += Distance(path[index - 1], path[index]);
                    var diagonal = Math.Sqrt(
                        CurrentScreenMap.ScreenW * (double)CurrentScreenMap.ScreenW +
                        CurrentScreenMap.ScreenH * (double)CurrentScreenMap.ScreenH);
                    if (totalLength < 4)
                        return "drag_path must contain visible movement.";
                    if (totalLength > diagonal * 8)
                        return "drag_path is longer than the bounded gesture limit.";
                    if (action.DurationMs is < 100 or > MaxGestureDurationMs)
                        return $"drag_path duration_ms must be between 100 and {MaxGestureDurationMs}.";
                    if (action.GestureKind is not null &&
                        action.GestureKind.ToLowerInvariant() is not ("draw" or "lasso" or "pan" or "slider" or "game" or "other"))
                        return "drag_path gesture_kind is not supported.";
                }

                if (action.Type == "hold_keys")
                {
                    if (action.Keys is null || action.Keys.Length == 0)
                        return "hold_keys requires at least one key.";
                    if (action.Keys.Length > MaxHeldKeys)
                        return $"hold_keys accepts at most {MaxHeldKeys} simultaneous keys.";
                    if (action.DurationMs is null or < 100 or > MaxKeyHoldDurationMs)
                        return $"hold_keys duration_ms must be between 100 and {MaxKeyHoldDurationMs}.";
                    try
                    {
                        foreach (var key in action.Keys)
                        {
                            if (string.IsNullOrWhiteSpace(key) || key.Contains('+'))
                                return "hold_keys keys must be individual key names, not chord strings.";
                            if (!IsHoldableKey(key))
                                return $"hold_keys does not allow the key '{key}'.";
                            _ = KeyNameToVk(key);
                        }
                    }
                    catch (Exception ex)
                    {
                        return $"hold_keys contains an unsupported key: {ex.Message}";
                    }
                }

                return null;
            }

            static string? ValidateNormalizedPair(
                double? x,
                double? y,
                string label)
            {
                if (x.HasValue != y.HasValue)
                    return $"{label} must be supplied as a complete pair.";
                if (!x.HasValue)
                    return null;
                if (!double.IsFinite(x.Value) ||
                    !double.IsFinite(y!.Value) ||
                    x.Value is < 0 or > 1 ||
                    y.Value is < 0 or > 1)
                {
                    return $"{label} must contain finite values in the 0..1 range.";
                }
                return null;
            }

            static string? ValidatePixelPair(
                int? x,
                int? y,
                int width,
                int height,
                string label)
            {
                if (x.HasValue != y.HasValue)
                    return $"{label} must be supplied as a complete pair.";
                if (!x.HasValue)
                    return null;
                if (x.Value < 0 ||
                    y!.Value < 0 ||
                    x.Value >= Math.Max(1, width) ||
                    y.Value >= Math.Max(1, height))
                {
                    return $"{label} must stay inside the current screenshot.";
                }
                return null;
            }

            static bool BoxFitsCurrentImage(BBox box) =>
                IsValidBox(box) &&
                box.Left >= 0 &&
                box.Top >= 0 &&
                box.Right <= CurrentScreenMap.ImageW &&
                box.Bottom <= CurrentScreenMap.ImageH;

            internal static bool IsValidBox(BBox box) =>
                box is { Left: { } left, Top: { } top, Right: { } right, Bottom: { } bottom } &&
                right > left &&
                bottom > top;

            internal static IReadOnlyList<Point> ResolveGesturePath(ActionDto action)
            {
                if (action.Path is null)
                    return Array.Empty<Point>();
                return action.Path
                    .Select(point => CurrentScreenMap.ImageToScreenPoint(point.XPx, point.YPx))
                    .Select(point => new Point(point.X, point.Y))
                    .ToArray();
            }

            static Rectangle? BuildActionObservationRegion(
                ActionDto action,
                Point? screenPoint,
                IReadOnlyList<Point> path)
            {
                if (path.Count > 0)
                {
                    var left = path.Min(point => point.X);
                    var top = path.Min(point => point.Y);
                    var right = path.Max(point => point.X);
                    var bottom = path.Max(point => point.Y);
                    var padding = action.Type == "drag_path" ? 24 : 48;
                    return ClampRect(Rectangle.FromLTRB(
                        left - padding,
                        top - padding,
                        right + padding + 1,
                        bottom + padding + 1));
                }
                if (screenPoint is Point point &&
                    action.Type is "click" or "double_click" or "focus_uia" or "click_uia")
                    return ClampRect(SquareAround(point.X, point.Y, 128));
                return null;
            }

            static double Distance(Point left, Point right)
            {
                var dx = (long)right.X - left.X;
                var dy = (long)right.Y - left.Y;
                return Math.Sqrt(dx * dx + dy * dy);
            }

            static string ActionSemanticTokens(ActionDto action)
            {
                var raw = action.Note ?? "";
                if (!string.IsNullOrWhiteSpace(action.GestureKind))
                    raw += $" {action.GestureKind}";
                if (action.UiaIndex is int index)
                {
                    var target = CurrentUiaTargets.FirstOrDefault(item => item.Index == index);
                    if (target is not null)
                        raw += $" {target.Name} {target.ControlType}";
                }
                return string.Join(' ', Regex.Matches(raw.ToLowerInvariant(), @"[\p{L}\p{Nd}]{2,}")
                    .Cast<Match>()
                    .Select(match => match.Value)
                    .Where(token => !int.TryParse(token, out _))
                    .Distinct(StringComparer.Ordinal)
                    .Take(20));
            }
        
            internal static Rectangle? ResolveAimRect(ActionDto a)
            {
                if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                    return ClampRect(CurrentScreenMap.ImageToScreenRect(RectFromBBox(a.BBox)));
                if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                    return ClampRect(CurrentScreenMap.ImageToScreenRect(RectFromBBox(a.Crop)));
                if (a.XPx.HasValue && a.YPx.HasValue)
                {
                    var p = CurrentScreenMap.ImageToScreenPoint(a.XPx.Value, a.YPx.Value);
                    return ClampRect(SquareAround(p.X, p.Y, FocusCropSize));
                }
                if (a.X.HasValue && a.Y.HasValue)
                {
                    var (px, py) = NormalizedToPixels(a.X.Value, a.Y.Value);
                    return ClampRect(SquareAround(px, py, FocusCropSize));
                }
                return null;
            }
        
            internal static Rectangle? ResolveCropRect(ActionDto a)
            {
                if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                    return ClampRect(CurrentScreenMap.ImageToScreenRect(RectFromBBox(a.Crop)));
                if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                    return ClampRect(CurrentScreenMap.ImageToScreenRect(RectFromBBox(a.BBox)));
                if (a.XPx.HasValue && a.YPx.HasValue)
                {
                    var p = CurrentScreenMap.ImageToScreenPoint(a.XPx.Value, a.YPx.Value);
                    return ClampRect(SquareAround(p.X, p.Y, FocusCropSize));
                }
                if (a.X.HasValue && a.Y.HasValue)
                {
                    var (px, py) = NormalizedToPixels(a.X.Value, a.Y.Value);
                    return ClampRect(SquareAround(px, py, FocusCropSize));
                }
                return null;
            }
        
            internal static Rectangle RectFromBBox(BBox b)
            {
                if (!IsValidBox(b))
                    throw new InvalidOperationException("Invalid bbox: right must be greater than left and bottom greater than top.");
                return Rectangle.FromLTRB(b.Left!.Value, b.Top!.Value, b.Right!.Value, b.Bottom!.Value);
            }
        
            // ===== Action execution =====
            [DllImport("user32.dll")] internal static extern bool SetCursorPos(int x, int y);
            [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        
            internal static void ExecuteAction(ActionDto action, Rectangle? clickContextRect = null)
            {
                var sw = Stopwatch.StartNew();
                try
                {
                    switch (action.Type)
                    {
                        case "open_url":
                            {
                                if (!AllowHighLevelActions)
                                    throw new InvalidOperationException("open_url is disabled. Use --allow-high-level-actions to enable it.");
                                if (string.IsNullOrWhiteSpace(action.Url))
                                    throw new InvalidOperationException("Missing url");
                                OpenUrl(action.Url);
                                break;
                            }
                        case "launch_app":
                            {
                                if (!AllowHighLevelActions)
                                    throw new InvalidOperationException("launch_app is disabled. Use --allow-high-level-actions to enable it.");
                                if (string.IsNullOrWhiteSpace(action.App))
                                    throw new InvalidOperationException("Missing app");
                                LaunchApp(action.App);
                                break;
                            }
                        case "run_command":
                            {
                                if (!AllowRunCommand)
                                    throw new InvalidOperationException("run_command is disabled. Use --allow-run-command to enable it.");
                                if (string.IsNullOrWhiteSpace(action.Command))
                                    throw new InvalidOperationException("Missing command");
                                RunCommand(action.Command);
                                break;
                            }
                        case "paste_text":
                            {
                                if (action.Text is null) throw new InvalidOperationException("Missing text");
                                PasteText(action.Text);
                                break;
                            }
                        case "focus_uia":
                            {
                                var target = ResolveUiaTarget(action);
                                SetCursorPos(target.CenterX, target.CenterY);
                                MouseClick("left");
                                break;
                            }
                        case "click_uia":
                            {
                                var target = ResolveUiaTarget(action);
                                SetCursorPos(target.CenterX, target.CenterY);
                                MouseClick(action.Button ?? "left");
                                break;
                            }
                        case "move":
                            {
                                var (x, y) = ResolvePoint(action);
                                SetCursorPos(x, y);
                                break;
                            }
                        case "click":
                            {
                                var (x, y) = ResolveClickPoint(action, clickContextRect, logAdjustment: true);
                                SetCursorPos(x, y);
                                MouseClick(action.Button ?? "left");
                                break;
                            }
                        case "double_click":
                            {
                                var (x, y) = ResolveClickPoint(action, clickContextRect, logAdjustment: true);
                                SetCursorPos(x, y);
                                MouseDoubleClick(action.Button ?? "left");
                                break;
                            }
                        case "drag_drop":
                            {
                                var source = ResolvePoint(action);
                                var destination = ResolveDropPoint(action);
                                MouseDragPath(
                                    new[] { new Point(source.X, source.Y), new Point(destination.X, destination.Y) },
                                    action.Button ?? "left",
                                    EffectiveDragDurationMs(action));
                                break;
                            }
                        case "drag_path":
                            {
                                MouseDragPath(
                                    ResolveGesturePath(action),
                                    action.Button ?? "left",
                                    EffectiveGestureDurationMs(action));
                                break;
                            }
                        case "keys":
                            {
                                if (action.Keys is null || action.Keys.Length == 0)
                                    throw new InvalidOperationException("Missing keys");
                                PressKeysSmart(action.Keys);
                                break;
                            }
                        case "hold_keys":
                            {
                                HoldKeys(action.Keys ?? Array.Empty<string>(), EffectiveKeyHoldDurationMs(action));
                                break;
                            }
                        case "type_text":
                            {
                                if (action.Text is null) throw new InvalidOperationException("Missing text");
                                TypeText(action.Text);
                                break;
                            }
                        case "scroll":
                            {
                                int dy = action.ScrollDy ?? 0;
                                if (dy != 0) MouseScroll(dy);
                                break;
                            }
                        case "request_crop":
                        case "point":
                        case "aim":
                        case "wait": // handled in loop (await Task.Delay), nothing here
                        case "done":
                            break;
        
                        default:
                            throw new InvalidOperationException($"Unknown action type: {action.Type}");
                    }
                }
                finally
                {
                    sw.Stop();
                    RunLocalActions++;
                    RunLocalActionElapsed += sw.Elapsed;
                }
            }
        
            internal static void OpenUrl(string url)
            {
                if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https"))
                    throw new InvalidOperationException($"Unsupported URL: {url}");
        
                Process.Start(new ProcessStartInfo
                {
                    FileName = uri.ToString(),
                    UseShellExecute = true
                });
            }
        
            internal static void LaunchApp(string app)
            {
                var trimmed = app.Trim();
                if (trimmed.Length == 0)
                    throw new InvalidOperationException("Empty app");
        
                Process.Start(new ProcessStartInfo
                {
                    FileName = trimmed,
                    UseShellExecute = true
                });
            }
        
            internal static void RunCommand(string command)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c " + command,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Minimized
                });
            }
        
            internal static void TypeText(string text)
            {
                if (ClipboardPasteThreshold > 0 && text.Length >= ClipboardPasteThreshold)
                {
                    PasteText(text);
                    return;
                }
        
                TypeUnicodeString(text);
            }
        
            internal static void PasteText(string text)
            {
                SetClipboardText(text);
                PressKeysSmart(new[] { "ctrl+v" });
            }
        
            internal static UiaTarget ResolveUiaTarget(ActionDto action)
            {
                if (!action.UiaIndex.HasValue)
                    throw new InvalidOperationException("Missing uia_index");
        
                var target = CurrentUiaTargets.FirstOrDefault(t => t.Index == action.UiaIndex.Value);
                if (target is null)
                    throw new InvalidOperationException($"Unknown uia_index {action.UiaIndex.Value}");
        
                return target;
            }
        
            internal static (int X, int Y) ResolvePoint(ActionDto a)
            {
                if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                {
                    var rect = CurrentScreenMap.ImageToScreenRect(RectFromBBox(a.BBox));
                    int cx = (rect.Left + rect.Right) / 2;
                    int cy = (rect.Top + rect.Bottom) / 2;
                    return ClampPoint(cx, cy);
                }
                if (a.XPx.HasValue && a.YPx.HasValue)
                {
                    var p = CurrentScreenMap.ImageToScreenPoint(a.XPx.Value, a.YPx.Value);
                    return ClampPoint(p.X, p.Y);
                }
                if (a.X.HasValue && a.Y.HasValue)
                {
                    var p = NormalizedToPixels(a.X.Value, a.Y.Value);
                    return ClampPoint(p.X, p.Y);
                }
                throw new InvalidOperationException("Missing coordinates (bbox or x_px/y_px or x/y).");
            }
        
            internal static (int X, int Y) ResolveClickPoint(ActionDto a, Rectangle? clickContextRect, bool logAdjustment)
            {
                var point = ResolvePoint(a);
                if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } } || clickContextRect is not Rectangle aim)
                    return point;
        
                if (!aim.Contains(point.X, point.Y) || !ShouldAdjustEdgeClick(point, aim))
                    return point;
        
                var adjusted = SafeClickPoint(aim);
                if (logAdjustment && (adjusted.X != point.X || adjusted.Y != point.Y))
                    Console.WriteLine($"[click] adjusted edge AIM click ({point.X},{point.Y}) -> ({adjusted.X},{adjusted.Y})");
                return adjusted;
            }

            internal static (int X, int Y) ResolveDropPoint(ActionDto action)
            {
                if (action.ToBBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                {
                    var rect = CurrentScreenMap.ImageToScreenRect(RectFromBBox(action.ToBBox));
                    return ClampPoint(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2);
                }
                if (action.ToXPx.HasValue && action.ToYPx.HasValue)
                {
                    var point = CurrentScreenMap.ImageToScreenPoint(action.ToXPx.Value, action.ToYPx.Value);
                    return ClampPoint(point.X, point.Y);
                }
                if (action.ToX.HasValue && action.ToY.HasValue)
                {
                    var point = NormalizedToPixels(action.ToX.Value, action.ToY.Value);
                    return ClampPoint(point.X, point.Y);
                }

                throw new InvalidOperationException("Missing drag destination (to_bbox or to_x_px/to_y_px or to_x/to_y).");
            }
        
            internal static bool ShouldAdjustEdgeClick((int X, int Y) point, Rectangle region)
            {
                if (region.Width <= 2 || region.Height <= 2)
                    return false;
        
                var (_, _, screenW, screenH) = GetPrimaryScreen();
                var areaRatio = (double)region.Width * region.Height / Math.Max(1, screenW * screenH);
                if (areaRatio > ClickAimEdgeAdjustMaxAreaRatio)
                    return false;
        
                var marginX = Math.Clamp((int)Math.Round(region.Width * ClickAimEdgeMarginRatio), ClickAimEdgeMinMarginPx, ClickAimEdgeMaxMarginPx);
                var marginY = Math.Clamp((int)Math.Round(region.Height * ClickAimEdgeMarginRatio), ClickAimEdgeMinMarginPx, ClickAimEdgeMaxMarginPx);
        
                return point.X <= region.Left + marginX ||
                       point.X >= region.Right - marginX ||
                       point.Y <= region.Top + marginY ||
                       point.Y >= region.Bottom - marginY;
            }
        
            internal static (int X, int Y) SafeClickPoint(Rectangle region)
            {
                var insetX = Math.Min(Math.Max(1, region.Width / 10), Math.Max(0, (region.Width - 1) / 2));
                var insetY = Math.Min(Math.Max(1, region.Height / 10), Math.Max(0, (region.Height - 1) / 2));
                var safe = Rectangle.FromLTRB(region.Left + insetX, region.Top + insetY, region.Right - insetX, region.Bottom - insetY);
                var x = safe.Width > 0 ? safe.Left + safe.Width / 2 : region.Left + region.Width / 2;
                var y = safe.Height > 0 ? safe.Top + safe.Height / 2 : region.Top + region.Height / 2;
                return ClampPoint(x, y);
            }
        
            internal static (int X, int Y) ClampPoint(int x, int y)
            {
                var (vx, vy, vw, vh) = GetPrimaryScreen();
                return (
                    Math.Max(vx, Math.Min(vx + vw - 1, x)),
                    Math.Max(vy, Math.Min(vy + vh - 1, y))
                );
            }
        
            internal static bool HasExplicitPoint(ActionDto a) =>
                (a.BBox is not null && IsValidBox(a.BBox))
                || (a.XPx is not null && a.YPx is not null)
                || (a.X is not null && a.Y is not null);

            internal static bool HasExplicitDropPoint(ActionDto action) =>
                (action.ToBBox is not null && IsValidBox(action.ToBBox))
                || (action.ToXPx is not null && action.ToYPx is not null)
                || (action.ToX is not null && action.ToY is not null);

            internal static int EffectiveDragDurationMs(ActionDto action) =>
                Math.Clamp(action.DragDurationMs ?? 500, 100, 3000);

            internal static int EffectiveGestureDurationMs(ActionDto action) =>
                Math.Clamp(action.DurationMs ?? 800, 100, MaxGestureDurationMs);

            internal static int EffectiveKeyHoldDurationMs(ActionDto action) =>
                Math.Clamp(action.DurationMs ?? 250, 100, MaxKeyHoldDurationMs);
        
            internal static (int X, int Y) NormalizedToPixels(double nx, double ny)
            {
                var (vx, vy, vw, vh) = GetPrimaryScreen();
                int x = vx + (int)Math.Round(nx * (vw - 1));
                int y = vy + (int)Math.Round(ny * (vh - 1));
                return (x, y);
            }
        
            // ==== Mouse ====
            internal static void MouseClick(string button)
            {
                var (down, up) = MouseButtonFlags(button);
                var inputs = new INPUT[]
                {
                    new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = down } } },
                    new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = up } } },
                };
                SendInputWithRetry(inputs, "mouse click");
            }

            internal static void MouseDragDrop(
                (int X, int Y) source,
                (int X, int Y) destination,
                string button,
                int durationMs)
            {
                MouseDragPath(
                    new[] { new Point(source.X, source.Y), new Point(destination.X, destination.Y) },
                    button,
                    durationMs);
            }

            internal static void MouseDragPath(
                IReadOnlyList<Point> path,
                string button,
                int durationMs)
            {
                if (path is null || path.Count < 2)
                    throw new InvalidOperationException("A pointer path requires at least two points.");

                var (down, up) = MouseButtonFlags(button);
                var plan = BuildMouseDragPlan(path, durationMs);
                SendAbsoluteMouseMove(path[0].X, path[0].Y, "drag initial move");

                var pressed = false;
                var movesSent = 0;
                Exception? dragFailure = null;
                var dragStopwatch = Stopwatch.StartNew();
                double buttonDownAtMs = 0;
                try
                {
                    SendInputWithRetry(
                        [new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = down } } }],
                        "drag button down");
                    pressed = true;
                    buttonDownAtMs = dragStopwatch.Elapsed.TotalMilliseconds;
                    Console.WriteLine(
                        $"[input] drag_path button={button} phase=down; points={path.Count}; " +
                        $"moves_planned={plan.Moves.Length}; duration={plan.EffectiveDurationMs}ms");

                    for (var index = 0; index < plan.Moves.Length; index++)
                    {
                        var targetElapsedMs = GestureMoveTargetElapsedMs(
                            plan,
                            index,
                            buttonDownAtMs);
                        WaitUntilGestureTime(dragStopwatch, targetElapsedMs);
                        var point = plan.Moves[index];
                        SendAbsoluteMouseMove(point.X, point.Y, "drag path move");
                        movesSent++;
                    }
                }
                catch (Exception ex)
                {
                    dragFailure = ex;
                    throw;
                }
                finally
                {
                    if (pressed)
                    {
                        try
                        {
                            SendInputWithRetry(
                                [new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = up } } }],
                                "drag button up");
                            dragStopwatch.Stop();
                            Console.WriteLine(
                                $"[input] drag_path button={button} phase=up; " +
                                $"moves_sent={movesSent}/{plan.Moves.Length}; elapsed={dragStopwatch.ElapsedMilliseconds}ms");
                        }
                        catch (Exception releaseFailure) when (dragFailure is not null)
                        {
                            Console.WriteLine($"[input] drag failed and button release also failed: {releaseFailure.Message}");
                        }
                    }
                }
            }

            internal static MouseDragPlan BuildMouseDragPlan(
                IReadOnlyList<Point> path,
                int durationMs)
            {
                if (path is null || path.Count < 2)
                    throw new InvalidOperationException("A pointer path requires at least two points.");

                var effectiveDuration = Math.Clamp(durationMs, 100, MaxGestureDurationMs);
                var lengths = new double[path.Count - 1];
                double totalLength = 0;
                for (var index = 1; index < path.Count; index++)
                {
                    lengths[index - 1] = Distance(path[index - 1], path[index]);
                    totalLength += lengths[index - 1];
                }
                if (totalLength < 1)
                    throw new InvalidOperationException("Pointer path has no movement.");

                var targetSteps = Math.Clamp(
                    Math.Max(effectiveDuration / 8, (int)Math.Ceiling(totalLength / 4)),
                    path.Count - 1,
                    600);
                var segmentStepCounts = lengths
                    .Select(length => Math.Max(
                        1,
                        (int)Math.Round(targetSteps * length / totalLength)))
                    .ToArray();
                var moves = new List<Point>(segmentStepCounts.Sum());
                for (var segment = 1; segment < path.Count; segment++)
                {
                    var segmentSteps = segmentStepCounts[segment - 1];
                    var source = path[segment - 1];
                    var destination = path[segment];
                    for (var index = 1; index <= segmentSteps; index++)
                    {
                        var progress = index / (double)segmentSteps;
                        moves.Add(new Point(
                            (int)Math.Round(source.X + (destination.X - source.X) * progress),
                            (int)Math.Round(source.Y + (destination.Y - source.Y) * progress)));
                    }
                }

                var initialHoldMs = Math.Clamp(effectiveDuration / 10, 20, 80);
                return new MouseDragPlan(
                    moves.ToArray(),
                    effectiveDuration,
                    initialHoldMs);
            }

            internal static double GestureMoveTargetElapsedMs(
                MouseDragPlan plan,
                int moveIndex,
                double buttonDownAtMs = 0)
            {
                if (moveIndex < 0 || moveIndex >= plan.Moves.Length)
                    throw new ArgumentOutOfRangeException(nameof(moveIndex));

                var movementDuration = plan.EffectiveDurationMs - plan.InitialHoldMs;
                return buttonDownAtMs +
                    plan.InitialHoldMs +
                    movementDuration * (moveIndex + 1.0) / plan.Moves.Length;
            }

            internal static void WaitUntilGestureTime(
                Stopwatch stopwatch,
                double targetElapsedMs)
            {
                while (true)
                {
                    if (CancelRequested)
                        throw new OperationCanceledException(
                            "Pointer gesture cancelled by the emergency hotkey.");

                    var remainingMs = targetElapsedMs - stopwatch.Elapsed.TotalMilliseconds;
                    if (remainingMs <= 0)
                        return;
                    if (remainingMs >= 3)
                    {
                        Thread.Sleep(Math.Max(1, (int)Math.Floor(remainingMs) - 1));
                        continue;
                    }
                    Thread.SpinWait(64);
                }
            }

            internal static AbsoluteMouseMoveData BuildAbsoluteMouseMoveData(
                int screenX,
                int screenY,
                int desktopX,
                int desktopY,
                int desktopW,
                int desktopH)
            {
                static int Normalize(int value, int origin, int size) =>
                    (int)Math.Round(
                        Math.Clamp(value - origin, 0, Math.Max(0, size - 1)) *
                        65535.0 /
                        Math.Max(1, size - 1));

                return new AbsoluteMouseMoveData(
                    Normalize(screenX, desktopX, desktopW),
                    Normalize(screenY, desktopY, desktopH),
                    MouseEventMove |
                    MouseEventMoveNoCoalesce |
                    MouseEventVirtualDesk |
                    MouseEventAbsolute);
            }

            static void SendAbsoluteMouseMove(int screenX, int screenY, string label)
            {
                var desktopX = GetSystemMetrics((int)SystemMetric.SM_XVIRTUALSCREEN);
                var desktopY = GetSystemMetrics((int)SystemMetric.SM_YVIRTUALSCREEN);
                var desktopW = GetSystemMetrics((int)SystemMetric.SM_CXVIRTUALSCREEN);
                var desktopH = GetSystemMetrics((int)SystemMetric.SM_CYVIRTUALSCREEN);
                if (desktopW <= 0 || desktopH <= 0)
                    (desktopX, desktopY, desktopW, desktopH) = GetPrimaryScreen();

                var move = BuildAbsoluteMouseMoveData(
                    screenX,
                    screenY,
                    desktopX,
                    desktopY,
                    desktopW,
                    desktopH);
                SendInputWithRetry(
                    [new INPUT
                    {
                        type = 0,
                        U = new InputUnion
                        {
                            mi = new MOUSEINPUT
                            {
                                dx = move.Dx,
                                dy = move.Dy,
                                dwFlags = move.Flags
                            }
                        }
                    }],
                    label);
            }

            static (uint Down, uint Up) MouseButtonFlags(string? button) =>
                (button ?? "left").ToLowerInvariant() switch
                {
                    "right" => (0x0008u, 0x0010u),
                    "middle" => (0x0020u, 0x0040u),
                    _ => (0x0002u, 0x0004u)
                };
        
            internal static void MouseDoubleClick(string button)
            {
                MouseClick(button);
                Thread.Sleep(80);
                MouseClick(button);
            }
        
            internal static void MouseScroll(int dyLines)
            {
                const int WHEEL_DELTA = 120;
                int delta = -Math.Clamp(dyLines, -20, 20) * WHEEL_DELTA; // action schema: positive is down
                var inputs = new INPUT[]
                {
                    new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = 0x0800, mouseData = (uint)delta } } } // MOUSEEVENTF_WHEEL
                };
                SendInputWithRetry(inputs, "mouse wheel");
            }
        
            // ==== Keyboard ====
            internal static void PressKeysSmart(string[] keys)
            {
                if (keys is null || keys.Length == 0)
                    throw new InvalidOperationException("Missing keys");
        
                // Recognized modifiers in chords
                var modifiers = new HashSet<string>(
                    new[] { "ctrl", "alt", "shift", "win", "super", "meta", "cmd" },
                    StringComparer.OrdinalIgnoreCase
                );
        
                // 1) Back-compat: ["ctrl","esc"] → treat as one chord
                if (keys.Length >= 2 && keys.Take(keys.Length - 1).All(k => modifiers.Contains(k)))
                {
                    var mods = keys.Take(keys.Length - 1).ToArray();
                    var main = keys.Last();
                    PressChord(mods, main);
                    return;
                }
        
                if (TrySendVirtualKeySequence(keys))
                    return;
        
                // 2) New mode: items may contain pluses, e.g., ["ctrl+shift+esc", "win+r", "tab"]
                foreach (var item in keys)
                {
                    if (string.IsNullOrWhiteSpace(item))
                        continue;
        
                    var parts = item.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        
                    // Single-string chord: modifiers + main key
                    if (parts.Length >= 2 && parts.Take(parts.Length - 1).All(p => modifiers.Contains(p)))
                    {
                        var mods = parts.Take(parts.Length - 1).ToArray();
                        var main = parts.Last();
                        PressChord(mods, main);
                    }
                    else
                    {
                        // Single key (e.g., "tab", "esc", "f5")
                        PressKey(item);
                    }
                }
            }
        
            internal static bool TrySendVirtualKeySequence(string[] keys)
            {
                if (keys.Length <= 1 || keys.Any(k => string.IsNullOrWhiteSpace(k) || k.Contains('+')))
                    return false;
        
                var inputs = new List<INPUT>(keys.Length * 2);
                try
                {
                    foreach (var key in keys)
                        if (!TryAddVirtualKeyPress(key, inputs))
                            return false;
                }
                catch
                {
                    return false;
                }
        
                SendKeyboardBatch(inputs);
                return true;
            }
        
            internal static void PressChord(string[] modifiers, string main)
            {
                var inputs = new List<INPUT>();
                try
                {
                    foreach (var m in modifiers)
                        inputs.Add(KeyboardInput(KeyNameToVk(m), keyUp: false));
        
                    if (!TryAddVirtualKeyPress(main, inputs))
                        throw new ArgumentException($"Chord main key cannot be batched: {main}");
        
                    foreach (var m in modifiers.Reverse())
                        inputs.Add(KeyboardInput(KeyNameToVk(m), keyUp: true));
        
                    SendKeyboardBatch(inputs);
                }
                catch
                {
                    foreach (var m in modifiers) KeyDown(m);
                    try { PressKey(main); }
                    finally
                    {
                        foreach (var m in modifiers.Reverse()) KeyUp(m);
                    }
                }
            }
        
            internal static void PressKey(string key)
            {
                if (string.IsNullOrEmpty(key))
                    throw new ArgumentException("Empty key");
        
                if (key == " " || key.Equals("space", StringComparison.OrdinalIgnoreCase))
                {
                    TypeUnicodeString(" ");
                    return;
                }
        
                if (key.Length == 1)
                {
                    char ch = key[0];
                    bool asciiLetter = (ch is >= 'a' and <= 'z') || (ch is >= 'A' and <= 'Z');
                    bool asciiDigit = (ch is >= '0' and <= '9');
                    if (!asciiLetter && !asciiDigit)
                    {
                        TypeUnicodeString(key);
                        return;
                    }
                }
        
                KeyDown(key);
                KeyUp(key);
            }
        
            private static bool TryAddVirtualKeyPress(string key, List<INPUT> inputs)
            {
                if (string.IsNullOrEmpty(key))
                    throw new ArgumentException("Empty key");
        
                if (key.Length == 1)
                {
                    char ch = key[0];
                    bool asciiLetter = (ch is >= 'a' and <= 'z') || (ch is >= 'A' and <= 'Z');
                    bool asciiDigit = (ch is >= '0' and <= '9');
                    if (!asciiLetter && !asciiDigit && ch != ' ')
                        return false;
                }
        
                var vk = KeyNameToVk(key);
                inputs.Add(KeyboardInput(vk, keyUp: false));
                inputs.Add(KeyboardInput(vk, keyUp: true));
                return true;
            }
        
            internal static void KeyDown(string key) { SendKeyboard(KeyNameToVk(key), false); }
            internal static void KeyUp(string key) { SendKeyboard(KeyNameToVk(key), true); }

            static readonly object HeldKeysGate = new();
            static readonly HashSet<string> HeldKeys = new(StringComparer.OrdinalIgnoreCase);

            internal static void HoldKeys(string[] keys, int durationMs)
            {
                if (keys is null || keys.Length == 0 || keys.Length > MaxHeldKeys)
                    throw new InvalidOperationException($"hold_keys requires 1..{MaxHeldKeys} keys.");

                var normalized = keys
                    .Select(key => key?.Trim().ToLowerInvariant() ?? "")
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (normalized.Any(key => !IsHoldableKey(key)))
                    throw new InvalidOperationException(
                        "hold_keys accepts only letters, digits, arrows, Space, Shift, Ctrl, and Alt.");
                var pressed = new List<string>(normalized.Length);
                Exception? holdFailure = null;
                try
                {
                    foreach (var key in normalized)
                    {
                        KeyDown(key);
                        pressed.Add(key);
                        lock (HeldKeysGate)
                            HeldKeys.Add(key);
                    }

                    var remaining = Math.Clamp(durationMs, 100, MaxKeyHoldDurationMs);
                    while (remaining > 0)
                    {
                        if (CancelRequested)
                            throw new OperationCanceledException("Key hold cancelled by the emergency hotkey.");
                        var slice = Math.Min(25, remaining);
                        Thread.Sleep(slice);
                        remaining -= slice;
                    }
                }
                catch (Exception ex)
                {
                    holdFailure = ex;
                    throw;
                }
                finally
                {
                    Exception? firstReleaseFailure = null;
                    foreach (var key in pressed.AsEnumerable().Reverse())
                    {
                        var released = false;
                        try
                        {
                            KeyUp(key);
                            released = true;
                        }
                        catch (Exception releaseFailure)
                        {
                            Console.WriteLine(
                                holdFailure is null
                                    ? $"[input] held key '{key}' could not be released: {releaseFailure.Message}"
                                    : $"[input] key hold failed and '{key}' release also failed: {releaseFailure.Message}");
                            firstReleaseFailure ??= releaseFailure;
                        }
                        finally
                        {
                            if (released)
                            {
                                lock (HeldKeysGate)
                                    HeldKeys.Remove(key);
                            }
                        }
                    }
                    if (holdFailure is null && firstReleaseFailure is not null)
                        throw new InvalidOperationException(
                            "One or more held keys could not be released.",
                            firstReleaseFailure);
                }
            }

            internal static void ReleaseAllHeldKeys()
            {
                string[] held;
                lock (HeldKeysGate)
                {
                    held = HeldKeys.ToArray();
                    HeldKeys.Clear();
                }
                foreach (var key in held.Reverse())
                {
                    try { KeyUp(key); }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[input] could not release held key '{key}': {ex.Message}");
                    }
                }
            }

            static bool IsHoldableKey(string? key)
            {
                var normalized = key?.Trim().ToLowerInvariant() ?? "";
                if (normalized.Length == 1 &&
                    (char.IsLetterOrDigit(normalized[0]) || normalized[0] == ' '))
                    return true;
                return normalized is "left" or "right" or "up" or "down" or
                    "arrowleft" or "arrowright" or "arrowup" or "arrowdown" or
                    "space" or "spacebar" or "shift" or "ctrl" or "control" or
                    "alt" or "option";
            }
        
            internal static ushort KeyNameToVk(string key)
            {
                key = key.ToLowerInvariant();
                if (key is "super" or "meta" or "cmd") key = "win";
                if (key == " ") return 0x20; // VK_SPACE
        
                return key switch
                {
                    "ctrl" or "control" => 0x11,
                    "alt" or "option" => 0x12,
                    "shift" => 0x10,
                    "win" or "windows" => 0x5B,
                    "enter" or "return" => 0x0D,
                    "tab" => 0x09,
                    "esc" or "escape" => 0x1B,
                    "space" or "spacebar" => 0x20,
                    "backspace" or "bksp" => 0x08,
                    "delete" or "del" => 0x2E,
                    "insert" or "ins" => 0x2D,
                    "left" or "arrowleft" => 0x25,
                    "up" or "arrowup" => 0x26,
                    "right" or "arrowright" => 0x27,
                    "down" or "arrowdown" => 0x28,
                    "home" => 0x24,
                    "end" => 0x23,
                    "pageup" or "pgup" => 0x21,
                    "pagedown" or "pgdn" => 0x22,
                    _ => MapAlnumToVk(key)
                };
            }
        
            internal static ushort MapAlnumToVk(string key)
            {
                if (key.Length == 1)
                {
                    char c = key[0];
                    if (char.IsLetter(c)) return (ushort)char.ToUpperInvariant(c); // 'A'..'Z'
                    if (char.IsDigit(c)) return (ushort)c;                         // '0'..'9'
                }
                if (key.StartsWith("f", StringComparison.OrdinalIgnoreCase) &&
                    int.TryParse(key[1..], out int f) && f is >= 1 and <= 24)
                    return (ushort)(0x70 + (f - 1)); // F1..F24
        
                throw new ArgumentException($"Unknown key: {key}");
            }
        
            internal static void TypeUnicodeString(string s)
            {
                var batch = new List<INPUT>(Math.Min(Math.Max(2, s.Length * 2), 128));
                foreach (var ch in s)
                {
                    batch.Add(UnicodeInput(ch, keyUp: false));
                    batch.Add(UnicodeInput(ch, keyUp: true));
        
                    if (batch.Count >= 128)
                        FlushUnicodeInputBatch(batch);
                }
        
                FlushUnicodeInputBatch(batch);
            }
        
            const uint CF_UNICODETEXT = 13;
            const uint GMEM_MOVEABLE = 0x0002;
        
            [DllImport("user32.dll", SetLastError = true)] internal static extern bool OpenClipboard(IntPtr hWndNewOwner);
            [DllImport("user32.dll", SetLastError = true)] internal static extern bool EmptyClipboard();
            [DllImport("user32.dll", SetLastError = true)] internal static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
            [DllImport("user32.dll", SetLastError = true)] internal static extern bool CloseClipboard();
            [DllImport("kernel32.dll", SetLastError = true)] internal static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);
            [DllImport("kernel32.dll", SetLastError = true)] internal static extern IntPtr GlobalLock(IntPtr hMem);
            [DllImport("kernel32.dll", SetLastError = true)] internal static extern bool GlobalUnlock(IntPtr hMem);
            [DllImport("kernel32.dll", SetLastError = true)] internal static extern IntPtr GlobalFree(IntPtr hMem);
        
            internal static void SetClipboardText(string text)
            {
                var bytes = Encoding.Unicode.GetBytes(text + "\0");
                var hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytes.Length);
                if (hGlobal == IntPtr.Zero)
                    throw new InvalidOperationException("GlobalAlloc for clipboard failed.");
        
                try
                {
                    var target = GlobalLock(hGlobal);
                    if (target == IntPtr.Zero)
                        throw new InvalidOperationException("GlobalLock for clipboard failed.");
        
                    try { Marshal.Copy(bytes, 0, target, bytes.Length); }
                    finally { GlobalUnlock(hGlobal); }
        
                    if (!OpenClipboardWithRetry())
                        throw new InvalidOperationException("OpenClipboard failed.");
        
                    try
                    {
                        EmptyClipboard();
                        if (SetClipboardData(CF_UNICODETEXT, hGlobal) == IntPtr.Zero)
                            throw new InvalidOperationException("SetClipboardData failed.");
                        hGlobal = IntPtr.Zero; // clipboard owns the handle now
                    }
                    finally
                    {
                        CloseClipboard();
                    }
                }
                finally
                {
                    if (hGlobal != IntPtr.Zero)
                        GlobalFree(hGlobal);
                }
            }
        
            internal static bool OpenClipboardWithRetry(int attempts = 6, int delayMs = 50)
            {
                for (var i = 0; i < attempts; i++)
                {
                    if (OpenClipboard(IntPtr.Zero))
                        return true;
                    Thread.Sleep(delayMs);
                }
                return false;
            }
        
            // ===== WinAPI / SendInput P/Invoke =====
            [StructLayout(LayoutKind.Sequential)]
            struct INPUT { public uint type; public InputUnion U; }
        
            [StructLayout(LayoutKind.Explicit)]
            struct InputUnion
            {
                [FieldOffset(0)] public MOUSEINPUT mi;
                [FieldOffset(0)] public KEYBDINPUT ki;
            }
        
            [StructLayout(LayoutKind.Sequential)]
            struct MOUSEINPUT
            {
                public int dx, dy;
                public uint mouseData, dwFlags, time;
                public nint dwExtraInfo;
            }
        
            [StructLayout(LayoutKind.Sequential)]
            struct KEYBDINPUT
            {
                public ushort wVk, wScan;
                public uint dwFlags, time;
                public nint dwExtraInfo;
            }
        
            internal static void SendKeyboard(ushort vk, bool keyUp)
            {
                var input = KeyboardInput(vk, keyUp);
                SendInputWithRetry(new[] { input }, "keyboard VK");
            }
        
            private static INPUT KeyboardInput(ushort vk, bool keyUp)
            {
                return new INPUT
                {
                    type = 1,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = vk,
                            wScan = 0,
                            dwFlags = keyUp ? 0x0002u : 0u, // KEYEVENTF_KEYUP
                            time = 0,
                            dwExtraInfo = 0
                        }
                    }
                };
            }
        
            private static void SendKeyboardBatch(List<INPUT> inputs)
            {
                if (inputs.Count == 0)
                    return;
        
                SendInputWithRetry(inputs.ToArray(), "keyboard VK batch");
            }
        
            internal static void SendUnicodeKey(char ch, bool keyUp)
            {
                var input = UnicodeInput(ch, keyUp);
                SendInputWithRetry(new[] { input }, "keyboard UNICODE");
            }
        
            private static INPUT UnicodeInput(char ch, bool keyUp)
            {
                return new INPUT
                {
                    type = 1,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0,
                            wScan = ch,
                            dwFlags = 0x0004u | (keyUp ? 0x0002u : 0u), // KEYEVENTF_UNICODE | KEYUP?
                            time = 0,
                            dwExtraInfo = 0
                        }
                    }
                };
            }
        
            private static void FlushUnicodeInputBatch(List<INPUT> batch)
            {
                if (batch.Count == 0)
                    return;
        
                var inputs = batch.ToArray();
                batch.Clear();
                SendInputWithRetry(inputs, "keyboard UNICODE batch");
            }
        
            private static void SendInputWithRetry(INPUT[] inputs, string label)
            {
                if (inputs.Length == 0)
                    return;
        
                var expected = (uint)inputs.Length;
                var size = Marshal.SizeOf<INPUT>();
                var lastError = 0;
        
                for (var attempt = 0; attempt <= SendInputMaxRetries; attempt++)
                {
                    var sent = SendInput(expected, inputs, size);
                    if (sent == expected)
                        return;
        
                    lastError = Marshal.GetLastWin32Error();
                    if (attempt < SendInputMaxRetries && SendInputRetryDelayMs > 0)
                        Thread.Sleep(SendInputRetryDelayMs);
                }
        
                throw new InvalidOperationException($"SendInput {label} failed after {SendInputMaxRetries + 1} attempt(s), last_error={lastError}.");
            }
    }
}



