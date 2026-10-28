namespace RDPilot.Control;

/// <summary>Accepts usable answers and removes invalid optional location metadata before display.</summary>
internal static class QaAnswerValidator
{
    internal static QaLocateDto? Validate(QaLocateDto? answer, int imageWidth, int imageHeight)
    {
        if (answer is null) return null;
        var text = string.IsNullOrWhiteSpace(answer.AnswerText) ? null : answer.AnswerText.Trim();
        var normalizedPoint = IsNormalized(answer.X) && IsNormalized(answer.Y);
        var pixelPoint = answer.XPx is int x && answer.YPx is int y &&
            x >= 0 && x < imageWidth && y >= 0 && y < imageHeight;
        var validBox = answer.BBox is { Left: int left, Top: int top, Right: int right, Bottom: int bottom } &&
            left >= 0 && top >= 0 && right > left && bottom > top &&
            right <= imageWidth && bottom <= imageHeight;
        // A note supplements an answer; it is not sufficient to complete a question on its own.
        if (text is null && !normalizedPoint && !pixelPoint && !validBox) return null;
        return new QaLocateDto
        {
            AnswerText = text,
            Note = string.IsNullOrWhiteSpace(answer.Note) ? null : answer.Note.Trim(),
            X = normalizedPoint ? answer.X : null,
            Y = normalizedPoint ? answer.Y : null,
            XPx = pixelPoint ? answer.XPx : null,
            YPx = pixelPoint ? answer.YPx : null,
            BBox = validBox ? answer.BBox : null
        };
    }

    static bool IsNormalized(double? value) =>
        value is double number && double.IsFinite(number) && number >= 0 && number <= 1;
}
