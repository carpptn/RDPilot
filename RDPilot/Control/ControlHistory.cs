namespace RDPilot.Control;

internal static class ControlHistory
{
    internal static void AddHistory(StringBuilder historyBuffer, string line)
    {
        if (Options.HistoryTailChars <= 0)
            return;
        if (string.IsNullOrWhiteSpace(line))
            return;
        if (historyBuffer.Length > 0)
            historyBuffer.AppendLine();
        historyBuffer.Append(line);
        var maxBufferChars = Math.Max(Options.HistoryTailChars * 2, Options.HistoryTailChars + 2048);
        if (historyBuffer.Length > maxBufferChars)
            historyBuffer.Remove(0, historyBuffer.Length - maxBufferChars);
    }

    internal static string TailHistory(StringBuilder historyBuffer, int maxChars, int maxLines)
    {
        if (maxChars <= 0 || historyBuffer.Length == 0)
            return "";
        var tail = Tail(historyBuffer, maxChars);
        var firstNewLine = tail.IndexOf('\n');
        if (historyBuffer.Length > tail.Length && firstNewLine >= 0 &&
            firstNewLine + 1 < tail.Length)
            tail = tail[(firstNewLine + 1)..];
        if (maxLines <= 0)
            return tail.Trim();
        var lines = tail.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length <= maxLines)
            return string.Join('\n', lines).Trim();
        return string.Join('\n', lines.Skip(lines.Length - maxLines)).Trim();
    }

    internal static string BuildRecentActionCheckpoint(IReadOnlyCollection<ResolvedActionSnapshot> recentActions)
    {
        if (recentActions.Count == 0)
            return "";
        var actions = recentActions.TakeLast(12).Select(action => TrimForMeta(action.Description, 160));
        return "RECENT_EXECUTED_ACTIONS (oldest to newest): " + string.Join(" | ", actions);
    }
}
