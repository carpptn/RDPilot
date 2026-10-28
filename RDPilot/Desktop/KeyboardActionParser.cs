namespace RDPilot.Desktop;

/// <summary>One key/chord grammar for validation, execution and action identity.</summary>
internal static class KeyboardActionParser
{
    internal sealed record Stroke(string[] Modifiers, string Key)
    {
        internal string Signature => string.Join("+", Modifiers.Append(Key));
    }

    internal static string NormalizeKey(string? key)
    {
        if (key == " ") return "space";
        if (key is { Length: 1 } && key[0] > 127) return key;
        return key?.Trim().ToLowerInvariant() switch
        {
            "control" => "ctrl",
            "option" => "alt",
            "windows" or "super" or "meta" or "cmd" => "win",
            "spacebar" => "space",
            "return" => "enter",
            "escape" => "esc",
            "arrowleft" => "left",
            "arrowright" => "right",
            "arrowup" => "up",
            "arrowdown" => "down",
            "bksp" => "backspace",
            "del" => "delete",
            "ins" => "insert",
            "pgup" => "pageup",
            "pgdn" => "pagedown",
            var value => value ?? ""
        };
    }

    internal static bool IsModifier(string key) => key is "ctrl" or "alt" or "shift" or "win";

    internal static Stroke[] Parse(string[]? keys, bool validate = true)
    {
        if (keys is null || keys.Length == 0) throw new ArgumentException("Missing keys");
        var normalized = keys.Select(NormalizeKey).ToArray();
        Stroke[] strokes;
        if (normalized.Length > 1 && normalized.Take(normalized.Length - 1).All(IsModifier))
            strokes = [new(normalized[..^1], normalized[^1])];
        else
            strokes = normalized.Select(ParseStroke).ToArray();
        foreach (var stroke in strokes)
        {
            if (string.IsNullOrEmpty(stroke.Key)) throw new ArgumentException("Empty key");
            // Unmodified single characters may still use Unicode text input.
            if (validate && stroke.Modifiers.Length > 0)
                _ = KeyboardInputEncoder.Chord(stroke.Modifiers, stroke.Key);
            else if (validate && stroke.Key.Length != 1) _ = KeyNameToVk(stroke.Key);
        }
        return strokes;
    }

    static Stroke ParseStroke(string value)
    {
        if (value.Length == 1 || !value.Contains('+')) return new([], value);
        var parts = (value.EndsWith("++", StringComparison.Ordinal)
            ? value[..^2].Split('+', StringSplitOptions.TrimEntries).Append("+")
            : value.Split('+', StringSplitOptions.TrimEntries)).Select(NormalizeKey).ToArray();
        if (parts.Length < 2 || parts.Any(string.IsNullOrEmpty) || !parts[..^1].All(IsModifier))
            throw new ArgumentException($"Invalid key chord: {value}");
        return new(parts[..^1], parts[^1]);
    }

    internal static string Signature(string[]? keys) =>
        string.Join(";", Parse(keys, validate: false).Select(stroke => stroke.Signature));
}
