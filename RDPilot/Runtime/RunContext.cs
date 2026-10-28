
namespace RDPilot.Runtime;

/// <summary>Mutable state owned by one control or question-answer run.</summary>
internal sealed class RunContext
{
    internal AppOptions? Configuration { get; init; }
    internal volatile bool CancelRequested;
    internal bool CliArgumentError = false;
    internal List<string> ConfigurationErrors { get; } = [];
    internal string? RequestReasoningEffortOverride = null;
    internal List<UiaTarget> CurrentUiaTargets = [];
    internal ScreenCoordinateMapper CurrentScreenMap = ScreenCoordinateMapper.Create(1, 1, 1, 1);
    internal int RunOpenAiCalls = 0;
    internal int RunOpenAiRetries = 0;
    internal TimeSpan RunOpenAiElapsed = TimeSpan.Zero;
    internal long RunOpenAiRequestBytes = 0;
    internal long RunOpenAiBytes = 0;
    internal int RunMultiCandidateResponses = 0;
    internal int RunWebSearchCalls = 0;
    internal long RunInputTokens = 0;
    internal long RunCachedTokens = 0;
    internal long RunOutputTokens = 0;
    internal long RunReasoningTokens = 0;
    internal int RunEarlyAcceptedControlStreams = 0;
    internal int RunControlContextTurns = 0;
    internal int RunControlContextRestarts = 0;
    internal int RunControlContextFallbacks = 0;
    internal int RunControlContextCompactions = 0;
    internal int RunControlCompactionFallbacks = 0;
    internal int RunScreenshotCount = 0;
    internal TimeSpan RunScreenshotElapsed = TimeSpan.Zero;
    internal int RunScreenProbeCount = 0;
    internal TimeSpan RunScreenProbeElapsed = TimeSpan.Zero;
    internal int RunScreenSanityWarnings = 0;
    internal int RunImageEncodeCount = 0;
    internal TimeSpan RunImageEncodeElapsed = TimeSpan.Zero;
    internal int RunScreenLogCount = 0;
    internal TimeSpan RunScreenLogElapsed = TimeSpan.Zero;
    internal int RunArtifactLogWrites = 0;
    internal TimeSpan RunArtifactLogElapsed = TimeSpan.Zero;
    internal int RunUiaCalls = 0;
    internal TimeSpan RunUiaElapsed = TimeSpan.Zero;
    internal int RunLocalActions = 0;
    internal TimeSpan RunLocalActionElapsed = TimeSpan.Zero;
    internal bool LastOpenAiFailureWasRetriable = false;
    internal string LastOpenAiFailureKind = "";
    internal string? LastOpenAiResponseId = null;
    internal HashSet<string> ReportedSanityWarnings = new(StringComparer.OrdinalIgnoreCase);
    internal HashSet<string> ReportedWebSearchCallIds = new(StringComparer.Ordinal);
}
