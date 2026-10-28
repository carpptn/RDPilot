using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static ResolvedActionSnapshot Snapshot(string type, Point point) => new(new ActionDto { Type = type, XPx = point.X, YPx = point.Y, Note = "test target" }, $"{type} {point}", $"{type}:{point.X},{point.Y}", point)
    {
        SemanticTokens = "test target"
    };
    static RecoveryLesson SemanticLesson(string target, string signature) => new()
    {
        ActiveProcess = "game",
        GoalTokens = "place tile slot",
        GoalDomain = "game",
        GoalIntentTokens = "move",
        WindowTitleTokens = "puzzle",
        FocusTokens = "canvas board",
        TriggerActionFamily = "click",
        LoopKind = RDPilot.Persistence.RecoveryMemoryService.PointerRegionLoop,
        LoopTopology = "direct-repeat",
        InteractionDomain = "raster-canvas",
        StrategySignature = signature,
        StrategySteps = [new RecoveryStrategyStep
        {
            ActionFamily = "click",
            TargetTokens = target,
            ParameterSignature = "left"
        }

        ],
        WinningActionTypes = ["click"],
        WinningStrategy = $"click {target}",
        ExpectedOutcomeTokens = "canvas board",
        UpdatedUtc = DateTime.UtcNow
    };
    static byte[] Fingerprint(byte value) => Enumerable.Repeat(value, 96 * 54).ToArray();
    static string TelemetryObservation(string runId, int step, byte[] fingerprint, double confidence, bool independentlyConfirmed, string goalMode = "finite", bool recurringWorkflowIntent = false) => System.Text.Json.JsonSerializer.Serialize(new { timestampUtc = DateTime.UtcNow.AddSeconds(step), @event = "observation", runId, step, screenWidth = 1280, screenHeight = 720, confidence, independentlyConfirmed, goalMode, recurringWorkflowIntent, replayFrame = new LoopReplayFrame { ScreenFingerprintBase64 = Convert.ToBase64String(fingerprint), ActiveWindowFingerprintBase64 = Convert.ToBase64String(fingerprint), ActiveProcess = "test", WindowTitle = "hashed title", FocusSummary = "hashed focus", LastDelta = step == 1 ? null : 0.1 } });
    static T GetStatic<T>(Type type, string name) => (T)TestState.Field(type, name, BindingFlags.Static | BindingFlags.NonPublic)!.GetValue(null)!;
    static object? InvokePrivate(Type type, string name, params object? [] args)
    {
        var method = type.GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic) ?? throw new MissingMethodException(type.FullName, name);
        return method.Invoke(null, args);
    }

    internal static Exception RootException(Exception ex)
    {
        while (ex.InnerException is not null)
            ex = ex.InnerException;
        return ex;
    }

    static void Assert(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
