/// <summary>
/// Converts coordinates and rectangles between a captured image and the physical screen.
/// </summary>
internal readonly record struct ScreenCoordinateMapper(
    int ScreenX,
    int ScreenY,
    int ScreenW,
    int ScreenH,
    int ImageW,
    int ImageH)
{
    public static ScreenCoordinateMapper Create(int screenW, int screenH, int imageW, int imageH)
        => Create(0, 0, screenW, screenH, imageW, imageH);

    public static ScreenCoordinateMapper Create(
        int screenX,
        int screenY,
        int screenW,
        int screenH,
        int imageW,
        int imageH)
    {
        screenW = Math.Max(1, screenW);
        screenH = Math.Max(1, screenH);
        imageW = Math.Max(1, imageW);
        imageH = Math.Max(1, imageH);
        return new ScreenCoordinateMapper(
            screenX,
            screenY,
            screenW,
            screenH,
            imageW,
            imageH);
    }

    public bool IsScaled => ScreenW != ImageW || ScreenH != ImageH;
    public bool HasNonZeroOrigin => ScreenX != 0 || ScreenY != 0;
    public bool RequiresMapping => IsScaled || HasNonZeroOrigin;

    public (int X, int Y) ImageToScreenPoint(int x, int y)
    {
        var sx = ScaleInclusive(x, ImageW, ScreenW);
        var sy = ScaleInclusive(y, ImageH, ScreenH);
        return (
            ScreenX + Clamp(sx, ScreenW),
            ScreenY + Clamp(sy, ScreenH));
    }

    public (int X, int Y) ScreenToImagePoint(int x, int y)
    {
        var ix = ScaleInclusive(x - ScreenX, ScreenW, ImageW);
        var iy = ScaleInclusive(y - ScreenY, ScreenH, ImageH);
        return (Clamp(ix, ImageW), Clamp(iy, ImageH));
    }

    public Rectangle ImageToScreenRect(Rectangle r) =>
        OffsetRect(
            ScaleRect(r, ImageW, ImageH, ScreenW, ScreenH),
            ScreenX,
            ScreenY);

    public Rectangle ScreenToImageRect(Rectangle r) =>
        ScaleRect(
            OffsetRect(r, -ScreenX, -ScreenY),
            ScreenW,
            ScreenH,
            ImageW,
            ImageH);

    static Rectangle OffsetRect(Rectangle rect, int dx, int dy) =>
        new(rect.X + dx, rect.Y + dy, rect.Width, rect.Height);

    static Rectangle ScaleRect(Rectangle r, int fromW, int fromH, int toW, int toH)
    {
        var left = (int)Math.Floor(r.Left * (double)toW / Math.Max(1, fromW));
        var top = (int)Math.Floor(r.Top * (double)toH / Math.Max(1, fromH));
        var right = (int)Math.Ceiling(r.Right * (double)toW / Math.Max(1, fromW));
        var bottom = (int)Math.Ceiling(r.Bottom * (double)toH / Math.Max(1, fromH));

        left = Clamp(left, toW);
        top = Clamp(top, toH);
        right = Math.Max(left + 1, Math.Min(toW, right));
        bottom = Math.Max(top + 1, Math.Min(toH, bottom));
        return Rectangle.FromLTRB(left, top, right, bottom);
    }

    static int ScaleInclusive(int value, int fromSize, int toSize)
    {
        if (fromSize <= 1 || toSize <= 1)
            return 0;
        return (int)Math.Round(value * (double)(toSize - 1) / (fromSize - 1));
    }

    static int Clamp(int value, int size) =>
        Math.Max(0, Math.Min(Math.Max(0, size - 1), value));
}

/// <summary>
/// Win32 system-metric identifiers used to determine the primary-screen dimensions.
/// </summary>
public enum SystemMetric
{
    SM_CXSCREEN = 0,
    SM_CYSCREEN = 1,
    SM_XVIRTUALSCREEN = 76,
    SM_YVIRTUALSCREEN = 77,
    SM_CXVIRTUALSCREEN = 78,
    SM_CYVIRTUALSCREEN = 79
}

/// <summary>
/// Describes one structured desktop action returned by the control model.
/// </summary>
public sealed class ActionDto
{
    [JsonPropertyName("type")] public string Type { get; set; } = "noop";

    [JsonPropertyName("x")] public double? X { get; set; }        // 0..1
    [JsonPropertyName("y")] public double? Y { get; set; }        // 0..1
    [JsonPropertyName("x_px")] public int? XPx { get; set; }      // pixels
    [JsonPropertyName("y_px")] public int? YPx { get; set; }      // pixels

    [JsonPropertyName("to_x")] public double? ToX { get; set; }       // drag destination, 0..1
    [JsonPropertyName("to_y")] public double? ToY { get; set; }       // drag destination, 0..1
    [JsonPropertyName("to_x_px")] public int? ToXPx { get; set; }     // drag destination, pixels
    [JsonPropertyName("to_y_px")] public int? ToYPx { get; set; }     // drag destination, pixels

    [JsonPropertyName("button")] public string? Button { get; set; }
    [JsonPropertyName("keys")] public string[]? Keys { get; set; }
    [JsonPropertyName("text")] public string? Text { get; set; }
    [JsonPropertyName("url")] public string? Url { get; set; }
    [JsonPropertyName("app")] public string? App { get; set; }
    [JsonPropertyName("command")] public string? Command { get; set; }
    [JsonPropertyName("uia_index")] public int? UiaIndex { get; set; }

    [JsonPropertyName("scroll_dy")] public int? ScrollDy { get; set; } // positive = down

    [JsonPropertyName("bbox")] public BBox? BBox { get; set; }      // target/click/drag source
    [JsonPropertyName("to_bbox")] public BBox? ToBBox { get; set; } // drag destination
    [JsonPropertyName("crop")] public BBox? Crop { get; set; }      // request_crop / aim / point

    [JsonPropertyName("drag_duration_ms")] public int? DragDurationMs { get; set; }
    [JsonPropertyName("wait_seconds")] public int? WaitSeconds { get; set; } // wait duration for 'wait'
    [JsonPropertyName("confidence")] public double? Confidence { get; set; } // 0..1 certainty for next action
    [JsonPropertyName("note")] public string? Note { get; set; }    // short comment (required in schema)
    [JsonPropertyName("recovery_strategy_id")] public string? RecoveryStrategyId { get; set; }
    [JsonPropertyName("recovery_strategy_step")] public int? RecoveryStrategyStep { get; set; } // 1-based
}

internal enum ControlRunOutcome
{
    Completed,
    Cancelled,
    GuardStopped,
    StepLimitReached,
    Failed
}

internal sealed record ControlRunResult(
    ControlRunOutcome Outcome,
    int Step,
    string Message)
{
    public bool Completed => Outcome == ControlRunOutcome.Completed;
}

/// <summary>
/// Preserves action coordinates and diagnostics in the screen coordinate space
/// that was active when the action was selected.
/// </summary>
internal sealed record ResolvedActionSnapshot(
    ActionDto Action,
    string Description,
    string IneffectiveSignature,
    Point? ScreenPoint)
{
    public string SemanticTokens { get; init; } = "";
    public Point? DestinationScreenPoint { get; init; }
    public string? ValidationError { get; init; }
    public bool IsValid => string.IsNullOrWhiteSpace(ValidationError);
}

/// <summary>
/// Temporarily blocks pointer actions near a recently ineffective screen point.
/// </summary>
internal readonly record struct SpatialActionCooldown(
    Point ScreenPoint,
    Point? DestinationScreenPoint,
    string ActionFamily,
    int UntilStep);

internal sealed record BatchedActionExecutionResult(
    IReadOnlyList<ResolvedActionSnapshot> ExecutedActions,
    string? Error);

/// <summary>
/// Keeps a short-lived visual state history for detecting multi-step loops.
/// </summary>
internal sealed record RecentLoopState(
    int Step,
    byte[] ScreenFingerprint,
    string ActiveProcess,
    string WindowTitleTokens)
{
    public string FocusTokens { get; init; } = "";
    public int GraphNodeId { get; init; } = -1;
}

/// <summary>
/// A bounded state-transition graph used to find both fixed and variable-length cycles.
/// </summary>
internal sealed class LoopStateGraph
{
    public string RunId { get; set; } = "";
    public List<LoopStateNode> Nodes { get; } = [];
    public List<LoopStateEdge> Edges { get; } = [];
    public List<int> RecentNodePath { get; } = [];
    public int? LastNodeId { get; set; }
    public int NextNodeId { get; set; }
    public int? PendingCandidateStep { get; set; }
    public double PendingCandidateConfidence { get; set; }
    public string PendingCandidateTopology { get; set; } = "";
    public string PendingCandidateDomain { get; set; } = "";
    public string PendingCandidateCalibrationKey { get; set; } = "";
    public int? PendingCandidateNodeId { get; set; }
    public string PendingCandidateSemanticStateKey { get; set; } = "";
    public int PendingCandidateVisitCount { get; set; }
    public bool PendingCandidateWasActionable { get; set; }
    public List<string> RecentSemanticPath { get; } = [];
    public Dictionary<string, List<int>> SemanticVisitSteps { get; } =
        new(StringComparer.Ordinal);
}

internal sealed class LoopStateNode
{
    public int Id { get; init; }
    public byte[] ScreenFingerprint { get; set; } = [];
    public byte[] ActiveWindowFingerprint { get; set; } = [];
    public string ActiveProcess { get; set; } = "";
    public string WindowTitleTokens { get; set; } = "";
    public string FocusTokens { get; set; } = "";
    public int VisitCount { get; set; }
    public int LastSeenStep { get; set; }
    public List<int> RecentVisitSteps { get; } = [];
    public double[] ScreenInstability { get; set; } = [];
    public double[] ActiveWindowInstability { get; set; } = [];
}

internal sealed class LoopStateEdge
{
    public int FromNodeId { get; init; }
    public int ToNodeId { get; init; }
    public string ActionKey { get; init; } = "";
    public int TraversalCount { get; set; }
    public int LastSeenStep { get; set; }
}

/// <summary>
/// Explains the weighted evidence behind a proactive multi-step loop decision.
/// </summary>
internal sealed record LoopDetectionAssessment(
    bool IsLoop,
    double Confidence,
    int CycleLength,
    int MatchingPriorStates,
    bool RepeatedActionCycle,
    bool ConsistentReturnPeriod,
    string Evidence)
{
    public string LoopTopology { get; init; } = "";
    public string InteractionDomain { get; init; } = "";
    public double DecisionThreshold { get; init; }
    public bool GraphCycle { get; init; }
    public bool SemanticCycle { get; init; }
    public bool IsProductiveCycle { get; init; }
    public string CycleDisposition { get; init; } = "none";
    public bool IndependentlyConfirmed { get; init; }
    public string CalibrationKey { get; init; } = "";
    public string RunId { get; init; } = "";
}

/// <summary>
/// A coordinate-independent, semantic step of a learned recovery strategy.
/// </summary>
internal sealed class RecoveryStrategyStep
{
    public string ActionFamily { get; set; } = "";
    public string Intent { get; set; } = "";
    public string TargetTokens { get; set; } = "";
    public string Preconditions { get; set; } = "";
    public string ExpectedEffect { get; set; } = "";
    public string ParameterSignature { get; set; } = "";
}

/// <summary>
/// A persistent, coordinate-independent lesson learned after escaping a control-loop stall.
/// </summary>
internal sealed class RecoveryLesson
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
    public string ActiveProcess { get; set; } = "";
    public string GoalTokens { get; set; } = "";
    public string GoalDomain { get; set; } = "";
    public string GoalIntentTokens { get; set; } = "";
    public string GoalMode { get; set; } = "finite";
    public string WindowTitleTokens { get; set; } = "";
    public string FocusTokens { get; set; } = "";
    public string TriggerActionFamily { get; set; } = "";
    public string LoopKind { get; set; } = "";
    public string LoopTopology { get; set; } = "";
    public string InteractionDomain { get; set; } = "";
    public string AvoidPattern { get; set; } = "";
    public string WinningStrategy { get; set; } = "";
    public string[] WinningActionTypes { get; set; } = [];
    public List<RecoveryStrategyStep> StrategySteps { get; set; } = [];
    public string StrategySignature { get; set; } = "";
    public string ScreenFingerprintBase64 { get; set; } = "";
    public string ActiveWindowFingerprintBase64 { get; set; } = "";
    public string ExpectedOutcomeTokens { get; set; } = "";
    public string ExpectedOutcomeFingerprintBase64 { get; set; } = "";
    public string ExpectedOutcomeActiveWindowFingerprintBase64 { get; set; } = "";
    public string Status { get; set; } = "active";
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
    public Dictionary<string, int> SuccessByWriter { get; set; } = [];
    public Dictionary<string, int> FailureByWriter { get; set; } = [];
    public int CompactedSuccessCount { get; set; }
    public int CompactedFailureCount { get; set; }
    public DateTime? CountersCompactedBeforeUtc { get; set; }
    public int ConsecutiveFailureCount { get; set; }
    public DateTime? LastSuccessUtc { get; set; }
    public DateTime? LastFailureUtc { get; set; }
    public DateTime? QuarantinedUtc { get; set; }
    public string LastFailureReason { get; set; } = "";
    public int SelectionCount { get; set; }
    public Dictionary<string, int> SelectionByWriter { get; set; } = [];
    public int CompactedSelectionCount { get; set; }
    public double CumulativeReward { get; set; }
    public Dictionary<string, double> RewardByWriter { get; set; } = [];
    public double CompactedCumulativeReward { get; set; }
    public int RewardObservationCount { get; set; }
    public Dictionary<string, int> RewardObservationByWriter { get; set; } = [];
    public int CompactedRewardObservationCount { get; set; }
    public double AverageActionCost { get; set; }
    public double LastProgressConfidence { get; set; }
    public string LastProgressEvidence { get; set; } = "";
    public string ValidationSource { get; set; } = "";
    public string RDPilotVersion { get; set; } = "";
    public string PromptVersion { get; set; } = "";
    public string ModelName { get; set; } = "";
}

/// <summary>
/// Versioned root object for the persistent recovery-memory file.
/// </summary>
internal sealed class RecoveryLessonStore
{
    public int Version { get; set; } = 7;
    public List<RecoveryLesson> Lessons { get; set; } = [];
    public Dictionary<string, LoopCalibrationBucket> Calibration { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);
}

internal sealed class RecoveryLessonArchiveStore
{
    public int Version { get; set; } = 1;
    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
    public List<RecoveryLesson> Lessons { get; set; } = [];
}

internal sealed class LoopCalibrationBucket
{
    public int CandidateCount { get; set; }
    public int ConfirmedCount { get; set; }
    public int RejectedCount { get; set; }
    public int InconclusiveCount { get; set; }
    public Dictionary<string, int> CandidateByWriter { get; set; } = [];
    public Dictionary<string, int> ConfirmedByWriter { get; set; } = [];
    public Dictionary<string, int> RejectedByWriter { get; set; } = [];
    public Dictionary<string, int> InconclusiveByWriter { get; set; } = [];
    public int CompactedCandidateCount { get; set; }
    public int CompactedConfirmedCount { get; set; }
    public int CompactedRejectedCount { get; set; }
    public int CompactedInconclusiveCount { get; set; }
    public DateTime? CountersCompactedBeforeUtc { get; set; }
    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Tracks one in-progress stagnation episode until a recovery is durably confirmed.
/// </summary>
internal sealed class RecoveryEpisodeState
{
    public int StartedAtStep { get; init; }
    public UiPromptContext TriggerContext { get; init; } = new("", "", null, null, null, null);
    public byte[] TriggerFingerprint { get; init; } = [];
    public byte[] TriggerActiveWindowFingerprint { get; init; } = [];
    public string GoalTokens { get; init; } = "";
    public string GoalDomain { get; init; } = "";
    public string GoalIntentTokens { get; init; } = "";
    public string GoalMode { get; init; } = "finite";
    public string TriggerActionFamily { get; init; } = "";
    public string LoopKind { get; init; } = "";
    public string LoopTopology { get; init; } = "";
    public string InteractionDomain { get; init; } = "";
    public List<ResolvedActionSnapshot> FailedActions { get; init; } = [];
    public List<ResolvedActionSnapshot> RecoveryActions { get; } = [];
    public int MaxStagnationSteps { get; set; }
    public bool IsValidating { get; set; }
    public int ValidationRemaining { get; set; }
    public RecoveryLesson? CandidateLesson { get; set; }
    public List<string> SuggestedLessonIds { get; } = [];
    public HashSet<string> RejectedLessonIds { get; } = new(StringComparer.Ordinal);
    public Dictionary<string, int> SuggestedLessonProgress { get; } = new(StringComparer.Ordinal);
    public string? AppliedLessonId { get; set; }
    public int AppliedLessonNoProgressObservations { get; set; }
    public bool AppliedLessonMadeProgress { get; set; }
    public bool ValidationObservedSemanticProgress { get; set; }
    public string LastOutcomeTokens { get; set; } = "";
    public string TriggerImageDataUrl { get; init; } = "";
    public string? TriggerImagePath { get; init; }
    public double LastProgressConfidence { get; set; }
    public string LastProgressEvidence { get; set; } = "";
}

/// <summary>
/// Independent assessment of whether a recovery produced goal-aligned progress.
/// For continuous goals, progress means resumed viable activity rather than completion.
/// </summary>
internal sealed class RecoveryProgressDto
{
    [JsonPropertyName("verdict")] public string? Verdict { get; set; } // yes|no|uncertain
    [JsonPropertyName("confidence")] public double Confidence { get; set; }
    [JsonPropertyName("evidence")] public string? Evidence { get; set; }
    [JsonPropertyName("state_label")] public string? StateLabel { get; set; }
}

internal sealed class LoopReplayCorpus
{
    public List<LoopReplayCase> Cases { get; set; } = [];
}

internal sealed class LoopReplayCase
{
    public string Name { get; set; } = "";
    public bool ExpectedLoop { get; set; }
    public string LabelSource { get; set; } = "";
    public DateTime CapturedUtc { get; set; }
    public string GoalMode { get; set; } = "finite";
    public bool RecurringWorkflowIntent { get; set; }
    public int ScreenWidth { get; set; } = 1280;
    public int ScreenHeight { get; set; } = 720;
    public List<LoopReplayFrame> Frames { get; set; } = [];

    [JsonIgnore]
    public bool HasIndependentLabel =>
        !string.IsNullOrWhiteSpace(LabelSource) &&
        !LabelSource.StartsWith("telemetry:", StringComparison.OrdinalIgnoreCase);
}

internal sealed class LoopReplayFrame
{
    public string ScreenFingerprintBase64 { get; set; } = "";
    public string ActiveWindowFingerprintBase64 { get; set; } = "";
    public string ActiveProcess { get; set; } = "";
    public string WindowTitle { get; set; } = "";
    public string FocusSummary { get; set; } = "";
    public ActionDto? PreviousAction { get; set; }
    public double? LastDelta { get; set; }
}

internal sealed class LoopTelemetryReplayEnvelope
{
    public DateTime TimestampUtc { get; set; }
    public string Event { get; set; } = "";
    public string RunId { get; set; } = "";
    public int Step { get; set; }
    public int ScreenWidth { get; set; }
    public int ScreenHeight { get; set; }
    public double Confidence { get; set; }
    public bool IndependentlyConfirmed { get; set; }
    public bool? Confirmed { get; set; }
    public string GoalMode { get; set; } = "finite";
    public bool RecurringWorkflowIntent { get; set; }
    public LoopReplayFrame? ReplayFrame { get; set; }
}

/// <summary>
/// Contains an answer to a screenshot question and an optional screen location.
/// </summary>
public sealed class QaLocateDto
{
    [JsonPropertyName("answer_text")] public string? AnswerText { get; set; }

    [JsonPropertyName("x")] public double? X { get; set; }       // 0..1
    [JsonPropertyName("y")] public double? Y { get; set; }       // 0..1
    [JsonPropertyName("x_px")] public int? XPx { get; set; }     // hint
    [JsonPropertyName("y_px")] public int? YPx { get; set; }     // hint
    [JsonPropertyName("bbox")] public BBox? BBox { get; set; }
    [JsonPropertyName("note")] public string? Note { get; set; }
}

/// <summary>
/// Contains the model's binary verification verdict for a completed goal.
/// </summary>
public sealed class VerifyDto
{
    [JsonPropertyName("verdict")] public string? Verdict { get; set; } // "yes"|"no"
    [JsonPropertyName("reason")] public string? Reason { get; set; }
}

/// <summary>
/// Represents an image-space bounding rectangle supplied by the model.
/// </summary>
public sealed class BBox
{
    [JsonPropertyName("left")] public int? Left { get; set; }
    [JsonPropertyName("top")] public int? Top { get; set; }
    [JsonPropertyName("right")] public int? Right { get; set; }
    [JsonPropertyName("bottom")] public int? Bottom { get; set; }
}

/// <summary>
/// Represents an actionable UI Automation element visible on the desktop.
/// </summary>
public sealed record UiaTarget(int Index, string Name, string ControlType, Rectangle Rect)
{
    public int CenterX => (Rect.Left + Rect.Right) / 2;
    public int CenterY => (Rect.Top + Rect.Bottom) / 2;
}

/// <summary>
/// Associates a UI Automation target with its local relevance score.
/// </summary>
public sealed record UiaTargetCandidate(UiaTarget Target, int Score);

/// <summary>
/// Captures the currently focused UI Automation element for a screenshot.
/// </summary>
public sealed record FocusedUiaSnapshot(Rectangle? Rect, string? Summary);

/// <summary>
/// Provides desktop state that enriches a control or question-answering prompt.
/// </summary>
public sealed record UiPromptContext(
    string ActiveWindowTitle,
    string ActiveProcessName,
    string? FocusedUiaSummary,
    string? BlockingPromptHint,
    string? ActiveWindowGeometry,
    string? WindowVisibilityHint);

/// <summary>
/// Summarizes one logged OpenAI response for offline analysis.
/// </summary>
public sealed record ResponseAnalysisRow(string FileName, string Kind, string Model, long Seconds, long InputTokens, long CachedTokens, long TotalTokens, long ReasoningTokens, int CandidateCount);

/// <summary>
/// Summarizes one logged OpenAI request for offline analysis.
/// </summary>
public sealed record RequestAnalysisRow(string FileName, string Kind, long Bytes, string Model, long TextChars, int ImageRefs, string CacheKey, int MaxOutputTokens, string Effort);

/// <summary>
/// Tracks the number and total size of images restored during request replay.
/// </summary>
public sealed record ReplayImageStats
{
    public int Images { get; set; }
    public long ImageBytes { get; set; }
}

/// <summary>
/// Writes output to two text writers, typically the console and a run-log file.
/// </summary>
public sealed class TeeTextWriter : TextWriter
{
    private readonly TextWriter _a;
    private readonly TextWriter _b;
    public TeeTextWriter(TextWriter a, TextWriter b) { _a = a; _b = b; }

    public override Encoding Encoding => Encoding.UTF8;
    public override void Flush() { _a.Flush(); _b.Flush(); }
    public override void Write(char value) { _a.Write(value); _b.Write(value); }
    public override void Write(string? value) { _a.Write(value); _b.Write(value); }
    public override void WriteLine(string? value) { _a.WriteLine(value); _b.WriteLine(value); }
    public override void WriteLine() { _a.WriteLine(); _b.WriteLine(); }
}

