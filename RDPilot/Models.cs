/// <summary>
/// Converts coordinates and rectangles between a captured image and the physical screen.
/// </summary>
internal readonly record struct ScreenCoordinateMapper(int ScreenW, int ScreenH, int ImageW, int ImageH)
{
    public static ScreenCoordinateMapper Create(int screenW, int screenH, int imageW, int imageH)
    {
        screenW = Math.Max(1, screenW);
        screenH = Math.Max(1, screenH);
        imageW = Math.Max(1, imageW);
        imageH = Math.Max(1, imageH);
        return new ScreenCoordinateMapper(screenW, screenH, imageW, imageH);
    }

    public bool IsScaled => ScreenW != ImageW || ScreenH != ImageH;

    public (int X, int Y) ImageToScreenPoint(int x, int y)
    {
        var sx = ScaleInclusive(x, ImageW, ScreenW);
        var sy = ScaleInclusive(y, ImageH, ScreenH);
        return (Clamp(sx, ScreenW), Clamp(sy, ScreenH));
    }

    public (int X, int Y) ScreenToImagePoint(int x, int y)
    {
        var ix = ScaleInclusive(x, ScreenW, ImageW);
        var iy = ScaleInclusive(y, ScreenH, ImageH);
        return (Clamp(ix, ImageW), Clamp(iy, ImageH));
    }

    public Rectangle ImageToScreenRect(Rectangle r) =>
        ScaleRect(r, ImageW, ImageH, ScreenW, ScreenH);

    public Rectangle ScreenToImageRect(Rectangle r) =>
        ScaleRect(r, ScreenW, ScreenH, ImageW, ImageH);

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

    [JsonPropertyName("button")] public string? Button { get; set; }
    [JsonPropertyName("keys")] public string[]? Keys { get; set; }
    [JsonPropertyName("text")] public string? Text { get; set; }
    [JsonPropertyName("url")] public string? Url { get; set; }
    [JsonPropertyName("app")] public string? App { get; set; }
    [JsonPropertyName("command")] public string? Command { get; set; }
    [JsonPropertyName("uia_index")] public int? UiaIndex { get; set; }

    [JsonPropertyName("scroll_dy")] public int? ScrollDy { get; set; } // positive = down

    [JsonPropertyName("bbox")] public BBox? BBox { get; set; }      // target/click
    [JsonPropertyName("crop")] public BBox? Crop { get; set; }      // request_crop / aim / point

    [JsonPropertyName("wait_seconds")] public int? WaitSeconds { get; set; } // wait duration for 'wait'
    [JsonPropertyName("confidence")] public double? Confidence { get; set; } // 0..1 certainty for next action
    [JsonPropertyName("note")] public string? Note { get; set; }    // short comment (required in schema)
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

