namespace RDPilot.Diagnostics;

/// <summary>Validates every recorded frame independently of loop detection or live desktop state.</summary>
internal static class ReplayCaseValidator
{
    internal sealed record ResolvedFrame(
        byte[] Screen,
        byte[] ActiveWindow,
        ResolvedActionSnapshot? PreviousAction,
        IReadOnlyList<ResolvedActionSnapshot>? RecentActions);

    internal static void Validate(LoopReplayCase replayCase)
    {
        if (replayCase is null || replayCase.Frames is null || replayCase.Frames.Count < 2)
            throw new InvalidOperationException("Replay case needs at least two frames.");
        if (replayCase.ScreenWidth < 1 || replayCase.ScreenHeight < 1)
            throw new InvalidOperationException("Replay screen dimensions must be positive.");

        // Do not retain all decoded images: validation memory is bounded to one frame.
        // Replay resolves its visited prefix again after the whole case has been checked.
        for (var index = 0; index < replayCase.Frames.Count; index++)
            _ = ResolveFrame(replayCase, index);
    }

    internal static ResolvedFrame ResolveFrame(LoopReplayCase replayCase, int index)
    {
        var frame = replayCase.Frames[index]
            ?? throw new InvalidOperationException($"Replay case '{replayCase.Name}' contains a null frame at {index + 1}.");
        if (frame.ReplayStateVersion is < 0 or > 2 || frame.StateEpoch is < 0)
            throw new InvalidOperationException($"Unsupported replay state metadata at frame {index + 1}.");
        var screen = DecodeReplayFingerprint(
            frame.ScreenFingerprintBase64,
            replayCase.Name,
            index,
            "screen");
        var active = string.IsNullOrWhiteSpace(
            frame.ActiveWindowFingerprintBase64)
            ? screen.ToArray()
            : DecodeReplayFingerprint(
                frame.ActiveWindowFingerprintBase64,
                replayCase.Name,
                index,
                "active-window");
        ResolvedActionSnapshot? previousAction = null;
        IReadOnlyList<ResolvedActionSnapshot>? recentActions = null;
        if (frame.ReplayStateVersion == 2)
        {
            if (frame.RecentActions is null || frame.RecentActions.Count > RecentActionLimit)
                throw new InvalidOperationException($"Invalid replay execution history at frame {index + 1}.");
            var history = frame.RecentActions.Select(item =>
                item?.Action is null ? throw new InvalidOperationException("Replay history contains a null action.") :
                ReplayActionResolver.Resolve(item.Action, item.Geometry, replayCase.ScreenWidth, replayCase.ScreenHeight)).ToArray();
            if (history.Any(item => !item.IsValid))
                throw new InvalidOperationException($"Invalid replay action in history at frame {index + 1}.");
            recentActions = history;
            if (frame.PreviousActionIndex is int actionIndex)
            {
                if (actionIndex < 0 || actionIndex >= history.Length || frame.PreviousAction is null)
                    throw new InvalidOperationException($"Invalid previous action index at frame {index + 1}.");
                var recordedPrevious = ReplayActionResolver.Resolve(frame, replayCase.ScreenWidth, replayCase.ScreenHeight);
                if (!recordedPrevious.IsValid || recordedPrevious.IneffectiveSignature != history[actionIndex].IneffectiveSignature)
                    throw new InvalidOperationException($"Previous action differs from execution history at frame {index + 1}.");
                previousAction = history[actionIndex];
            }
        }
        if (previousAction is null && frame.PreviousAction is not null)
            previousAction = ReplayActionResolver.Resolve(frame, replayCase.ScreenWidth, replayCase.ScreenHeight);
        if (previousAction is { IsValid: false })
            throw new InvalidOperationException($"Replay case '{replayCase.Name}', frame {index + 1}: {previousAction.ValidationError}");
        return new(screen, active, previousAction, recentActions);
    }

    static byte[] DecodeReplayFingerprint(
        string encoded,
        string caseName,
        int frameIndex,
        string kind)
    {
        try
        {
            var fingerprint = Convert.FromBase64String(encoded);
            if (fingerprint.Length == 0)
                throw new FormatException("fingerprint is empty");
            return fingerprint;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Replay case '{caseName}', frame {frameIndex + 1} has an invalid {kind} fingerprint.",
                ex);
        }
    }
}
