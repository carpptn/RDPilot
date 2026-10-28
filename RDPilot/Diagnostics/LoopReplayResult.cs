namespace RDPilot.Diagnostics;

/// <summary>Regression counts returned to callers as well as rendered in the replay report.</summary>
internal sealed record LoopReplayResult(int TruePositive, int FalsePositive, int TrueNegative, int FalseNegative)
{
    internal int CaseCount => TruePositive + FalsePositive + TrueNegative + FalseNegative;
    internal int FailureCount => FalsePositive + FalseNegative;
    internal CommandOutcome Outcome => FailureCount == 0 ? CommandOutcome.Success : CommandOutcome.Failed;
}
