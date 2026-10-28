namespace RDPilot.Control;

/// <summary>Canonical action identity shared by desktop execution and offline replay.</summary>
internal static class ActionSignatureService
{
    internal static string Create(ActionDto action, Point? source = null, Point? destination = null,
        IReadOnlyList<Point>? path = null, Rectangle? region = null, string semanticTokens = "")
    {
        var type = action.Type?.ToLowerInvariant() ?? "null";
        var keys = string.Join("+", (action.Keys ?? []).Select(KeyboardActionParser.NormalizeKey).Distinct(StringComparer.Ordinal));
        var context = string.IsNullOrWhiteSpace(semanticTokens) ? "unspecified-target" : SensitiveHash(semanticTokens);
        switch (type)
        {
            case "open_url": return $"open_url:{action.Url}";
            case "launch_app": return $"launch_app:{action.App}";
            case "run_command": return $"run_command:{SensitiveHash(action.Command)}:{context}";
            case "type_text":
            case "paste_text": return $"{type}:{SensitiveHash(action.Text)}:{context}";
            case "focus_uia":
            case "click_uia": return $"{type}:{MouseButton(action)}:{action.UiaIndex}";
            case "keys": return $"keys:{KeyboardActionParser.Signature(action.Keys)}";
            case "hold_keys": return $"hold_keys:{keys}:{EffectiveKeyHoldDurationMs(action)}";
            case "scroll": return $"scroll:{action.ScrollDy ?? 0}";
            case "wait": return $"wait:{EffectiveWaitSeconds(action, out _)}";
            case "drag_drop":
                if (source is not Point from || destination is not Point to) return "drag_drop:missing";
                var cluster = Math.Max(16, Options.IneffectiveMouseClusterPx);
                return $"drag_drop:{MouseButton(action)}:{from.X / cluster},{from.Y / cluster}->{to.X / cluster},{to.Y / cluster}";
            case "drag_path":
                if (path is null || path.Count < 2) return "drag_path:missing";
                var pathCluster = Math.Max(8, Options.IneffectiveMouseClusterPx / 2);
                var sampled = path.Where((_, i) => i == 0 || i == path.Count - 1 || i % Math.Max(1, path.Count / 10) == 0)
                    .Take(12).Select(p => $"{p.X / pathCluster},{p.Y / pathCluster}");
                return $"drag_path:{MouseButton(action)}:{action.GestureKind ?? "other"}:{string.Join(';', sampled)}";
            case "move":
            case "click":
            case "double_click":
                return source is Point p ? $"{type}:{MouseButton(action)}:{p.X / 16 * 16},{p.Y / 16 * 16}" : $"{type}:missing";
            case "aim":
            case "point":
            case "request_crop":
                return region is Rectangle r ? $"{type}:{(r.Left + r.Width / 2) / 16 * 16},{(r.Top + r.Height / 2) / 16 * 16}" : type;
            default: return type;
        }
    }

    static string SensitiveHash(string? value) => string.IsNullOrEmpty(value) ? "empty" :
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)).AsSpan(0, 6)).ToLowerInvariant();
}
