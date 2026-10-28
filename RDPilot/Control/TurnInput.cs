namespace RDPilot.Control;

/// <summary>Separates the physical directional key from its logical direction on a board.</summary>
internal static class TurnInput
{
    internal static string? Key(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        "arrowup" or "up" => "ArrowUp", "arrowdown" or "down" => "ArrowDown",
        "arrowleft" or "left" => "ArrowLeft", "arrowright" or "right" => "ArrowRight",
        "w" => "W", "a" => "A", "s" => "S", "d" => "D", _ => null
    };

    internal static bool IsDirectionalSequence(string label) =>
        label.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .All(item => Key(item) is not null);

    internal static string? Direction(string? value) => Key(value) switch
    {
        "W" => "ArrowUp", "A" => "ArrowLeft", "S" => "ArrowDown", "D" => "ArrowRight",
        var key => key
    };
}
