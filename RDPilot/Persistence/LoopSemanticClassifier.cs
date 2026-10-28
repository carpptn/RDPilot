namespace RDPilot.Persistence;

/// <summary>Computes the same semantic features for live input and recorded replay.</summary>
internal static class LoopSemanticClassifier
{
    const string RasterIntent = @"\b(game|puzzle|board|map|image|diagram|drawing|paint|piece|tile)\b";
    internal static LoopActionFeatures Action(ResolvedActionSnapshot snapshot)
    {
        if (snapshot.SemanticFeatures is { } recorded) return recorded;
        var intent = $"{snapshot.Action.Note} {snapshot.Description} {snapshot.SemanticTokens}";
        var actionIntent = $"{snapshot.Action.Note} {snapshot.Description}";
        return new(
            Match($"{snapshot.SemanticTokens} {snapshot.Action.Note}", @"\b(monitor|watch|observe|poll|check|refresh|status|health|incoming|new\s+event|monitoruj|obserwuj|sprawdź|sprawdzaj|odśwież|status|nowe\s+zdarzenie)\b"),
            Match(intent, @"\b(help|instruction\w*|guide|documentation|details|information|pomoc\w*|instrukcj\w*|wskazówk\w*|szczegół\w*|informacj\w*)\b"),
            Match(intent, @"\b(back|return\w*|close\w*|dismiss\w*|resume\w*|wróć\w*|powr\w*|zamkn\w*|wznów\w*)\b") ||
                snapshot.Action.Type == "keys" && snapshot.Action.Keys?.Any(key =>
                    KeyboardActionParser.NormalizeKey(key) is "esc" or "browserback") == true,
            Match(actionIntent, @"\b(drag|drop|place|placement|position|reposition|move\s+(?:the\s+)?(?:piece|item|object|token|tile|card))\b"),
            Match(actionIntent, RasterIntent));
    }

    internal static LoopContextFeatures Context(UiPromptContext context)
    {
        if (context.SemanticFeatures is { } recorded) return recorded;
        var visual = $"{context.ActiveWindowTitle} {context.FocusedUiaSummary}".ToLowerInvariant();
        return new(Match(visual, @"\b(canvas|raster|bitmap|pixel|sprite)\b"),
            visual.Contains("rootwebarea", StringComparison.Ordinal) ||
            visual.Contains("controltype.document", StringComparison.Ordinal) ||
            visual.Contains(" document", StringComparison.Ordinal), Match(visual, RasterIntent));
    }

    static bool Match(string value, string pattern) => Regex.IsMatch(value.ToLowerInvariant(), pattern, RegexOptions.CultureInvariant);
}