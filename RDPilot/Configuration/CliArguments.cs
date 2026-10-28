namespace RDPilot.Configuration;

internal sealed record CliValidationResult(string? Task, IReadOnlyList<string> Errors)
{
    internal bool Success => Errors.Count == 0;
}

internal sealed record CliArgument(string Name, string? Value)
{
    internal bool TryRead(string name, out string value)
    {
        value = Value ?? "";
        return Value is not null && Name.Equals(name, StringComparison.OrdinalIgnoreCase);
    }
}

/// <summary>Tokenizes value options once for both profile selection and individual overrides.</summary>
internal static class CliArguments
{
    private static readonly HashSet<string> ValueOptions = new(StringComparer.OrdinalIgnoreCase)
    {
        "--context-compact-threshold",
        "--context-fallback-limit",
        "--control-reasoning-context",
        "--crop-format",
        "--crop-max-width",
        "--delay",
        "--effort",
        "--focus-crop-size",
        "--focused-overview-max-width",
        "--goal",
        "--goal-mode",
        "--graph-candidate-ttl",
        "--grid",
        "--history-chars",
        "--history-lines",
        "--incomplete-max-output-retries",
        "--incomplete-max-output-token-cap",
        "--jpeg-quality",
        "--loop-confidence-threshold",
        "--loop-replay",
        "--loop-replay-corpus",
        "--loop-replay-export",
        "--loop-replay-import",
        "--max-action-failures",
        "--max-action-text-chars",
        "--max-artifacts",
        "--max-batch-actions",
        "--max-focus-crop-pixels",
        "--max-inspection-actions",
        "--max-model-failures",
        "--max-output-tokens",
        "--max-rejected-proposals",
        "--max-repeated-actions",
        "--max-stagnation",
        "--max-steps",
        "--max-uia-nodes",
        "--max-uia-targets",
        "--max-wait",
        "--memory-export",
        "--model",
        "--observation-profile",
        "--openai-timeout",
        "--paste-threshold",
        "--profile",
        "--prompt-cache-key",
        "--qa-effort",
        "--qa-max-output-tokens",
        "--qa-model",
        "--qa-screenshot-max-width",
        "--reasoning-effort",
        "--recovery-archive-max-bytes",
        "--recovery-archive-path",
        "--recovery-archive-retained-files",
        "--recovery-failure-limit",
        "--recovery-max-file-bytes",
        "--recovery-max-lessons",
        "--recovery-max-quarantined",
        "--recovery-memory-path",
        "--recovery-progress-confidence",
        "--recovery-prompt-lessons",
        "--recovery-reserved-per-context",
        "--recovery-soft-max-per-context",
        "--recovery-telemetry-max-bytes",
        "--recovery-telemetry-retained-files",
        "--recovery-trigger",
        "--recovery-validation",
        "--repeat-cooldown",
        "--replay-request",
        "--replay-response",
        "--retries",
        "--runtime-cooldowns",
        "--runtime-graph-edges",
        "--runtime-recovery-actions",
        "--runtime-semantic-states",
        "--screen-log-format",
        "--screen-log-max-width",
        "--screen-poll-initial-delay",
        "--screen-poll-interval",
        "--screen-poll-timeout",
        "--screenshot-format",
        "--screenshot-max-width",
        "--sendinput-retries",
        "--sendinput-retry-delay",
        "--skip-verify-confidence",
        "--task",
        "--task-file",
        "--turn-batch-inputs",
        "--turn-reanalysis-max-output-tokens",
        "--uia-candidate-multiplier",
        "--uia-max-area-ratio",
        "--uia-name-chars",
        "--uia-scan-ms",
        "--uia-summary-chars",
        "--verbosity",
        "--verify",
        "--verify-early-steps",
        "--verify-effort",
        "--verify-low-confidence",
        "--verify-max-output-tokens",
        "--verify-model",
        "--verify-screenshot-max-width",
        "--wait-no-change-extra",
    };

    internal static IReadOnlyList<CliArgument> Parse(string[] args, Action<string> reportError)
    {
        var result = new List<CliArgument>();
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (string.IsNullOrWhiteSpace(arg)) continue;
            var equals = arg.StartsWith("--", StringComparison.Ordinal) ? arg.IndexOf('=') : -1;
            var name = equals < 0 ? arg : arg[..equals];
            if (!ValueOptions.Contains(name))
            {
                if (equals >= 0)
                    reportError($"Unknown option or unexpected value: {arg}");
                else
                    result.Add(new CliArgument(name, null));
                continue;
            }
            string? value = equals >= 0 ? arg[(equals + 1)..] : null;
            if (equals < 0 && i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                value = args[++i];
            if (string.IsNullOrWhiteSpace(value))
                reportError($"Missing value for {name}.");
            else
                result.Add(new CliArgument(name, value));
        }
        return result;
    }
}
