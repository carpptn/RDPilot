namespace RDPilot.Runtime;

/// <summary>Renders Q&A status from the same outcome used for the batch exit code.</summary>
internal static class CommandPresentation
{
    internal static void WriteQuestionResult(CommandOutcome outcome, bool batch)
    {
        var message = (batch, outcome) switch
        {
            (true, CommandOutcome.Success) => "[batch] completed",
            (true, CommandOutcome.Cancelled) => "[batch] cancelled",
            (true, _) => "[batch] failed: no usable answer was returned",
            (false, CommandOutcome.Success) => "✅ Answer completed. Enter next (ENTER = exit).",
            (false, CommandOutcome.Cancelled) => "Answer cancelled. Enter next (ENTER = exit).",
            (false, _) => "Answer failed: no usable answer was returned. Enter next (ENTER = exit)."
        };
        if (outcome == CommandOutcome.Success) Console.WriteLine(message);
        else Console.Error.WriteLine(message);
    }
}
