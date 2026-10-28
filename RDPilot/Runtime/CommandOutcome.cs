namespace RDPilot.Runtime;

/// <summary>Shared terminal command status; numeric values are process exit codes.</summary>
internal enum CommandOutcome
{
    Success = 0,
    Failed = 1,
    Cancelled = 130
}
