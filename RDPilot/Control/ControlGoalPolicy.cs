namespace RDPilot.Control;

internal static class ControlGoalPolicy
{
    internal static bool IsHighImpactGoal(string goal)
    {
        var g = goal.ToLowerInvariant();
        string[] highImpactTerms =
        {
            "delete", "remove", "erase", "save", "send", "submit", "publish",
            "purchase", "buy", "payment", "run script", "execute script",
            "skasuj", "kasuj", "opublikuj", "kup", "zapłać", "zaplac",
            "wyślij",
            "wyslij",
            "email",
            "e-mail",
            "mail",
            "outlook",
            "zapisz",
            "plik",
            "usuń",
            "usun",
            "word",
            "excel",
            "powerpoint",
            "uruchom skrypt",
            "powershell",
            "cmd"
        };
        return highImpactTerms.Any(t => g.Contains(t, StringComparison.OrdinalIgnoreCase));
    }

    internal static bool ShouldVerifyGoal(string goal, int step, ActionDto? action = null)
    {
        if (Options.VerifyMode.Equals("always", StringComparison.OrdinalIgnoreCase))
            return true;
        if (Options.VerifyMode.Equals("off", StringComparison.OrdinalIgnoreCase))
            return false;
        if (IsHighImpactGoal(goal))
            return true;
        if (Options.VerifyEarlySteps > 0 && step <= Options.VerifyEarlySteps)
            return true;
        if (action?.Confidence is not double confidence || !double.IsFinite(confidence) || confidence is < 0 or > 1)
            return true;
        // Low-confidence results always require verification. A completion may skip it
        // only at the explicit skip threshold; the interval between the thresholds verifies.
        if (confidence < Options.VerifyLowConfidenceThreshold)
            return true;
        return action.Type != "done" || confidence < Options.SkipVerifyConfidenceThreshold;
    }

    internal static string VerificationSkipReason(int step, ActionDto? action) =>
        Options.VerifyMode.Equals("off", StringComparison.OrdinalIgnoreCase)
            ? "verify=off"
            : $"verify=auto high-confidence completion ({action?.Confidence:0.00} >= {Options.SkipVerifyConfidenceThreshold:0.00}); step={step}";

    internal static bool TryExecuteFastGoal(string goal)
    {
        var g = goal.Trim();
        if (string.IsNullOrWhiteSpace(g))
            return false;
        if (TryExtractUrl(g, out var directUrl))
        {
            Console.WriteLine($"[fast-path] open_url {directUrl}");
            OpenUrl(directUrl);
            return true;
        }

        return false;
    }

    internal static bool TryExtractUrl(string goal, out string url)
    {
        url = "";
        var match = Regex.Match(
            goal.Trim(),
            @"^(?:(?:open|go to|navigate to|otwórz|otworz|wejdź na|wejdz na)\s+)?(?<url>https?://[^\s""<>]+|[a-z0-9][a-z0-9-]*(?:\.[a-z0-9][a-z0-9-]*)+(?:/[^\s""<>]*)?)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success)
            return false;
        url = match.Groups["url"].Value;
        if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
            !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            url = "https://" + url;
        return true;
    }

    internal static string? EffectiveReasoningEffort(int stagnationSteps, int repeatCount, int rejectedProposalCycleCount = 0)
    {
        if (!Options.AdaptiveReasoningEffort ||
            string.IsNullOrWhiteSpace(Options.ReasoningEffort) ||
            !SupportsReasoningEffort(Options.Model))
            return Options.ReasoningEffort;
        var current = Options.ReasoningEffort.ToLowerInvariant();
        if (current is "high" or "xhigh" or "max")
            return Options.ReasoningEffort;
        if (stagnationSteps >= 4 || repeatCount >= 3 || rejectedProposalCycleCount >= 3)
            return "high";
        if (stagnationSteps >= 2 || repeatCount >= 1 || rejectedProposalCycleCount >= 1)
            return "medium";
        return Options.ReasoningEffort;
    }
}
