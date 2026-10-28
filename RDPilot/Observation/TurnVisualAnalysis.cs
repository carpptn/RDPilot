namespace RDPilot.Observation;

/// <summary>Pure, color-aware sampling and change metrics for turn-state tracking.</summary>
internal static class TurnVisualAnalysis
{
    internal const int Side = 96;
    const int Channels = 3;

    internal static byte[] Extract(ScreenObservationFrame frame, Rectangle region)
    {
        if (frame.DetailWidth <= 0 || frame.DetailHeight <= 0 ||
            frame.DetailFingerprint.Length != frame.DetailWidth * frame.DetailHeight) return [];
        region.Intersect(frame.ScreenBounds);
        if (region.Width <= 0 || region.Height <= 0) return [];
        var color = frame.DetailColorFingerprint.Length == frame.DetailFingerprint.Length * Channels;
        var output = new byte[Side * Side * Channels];
        for (var y = 0; y < Side; y++)
        for (var x = 0; x < Side; x++)
        {
            var screenY = region.Top + (y + 0.5) * region.Height / Side;
            var screenX = region.Left + (x + 0.5) * region.Width / Side;
            var detailY = Math.Clamp((int)((screenY - frame.ScreenBounds.Top) * frame.DetailHeight /
                Math.Max(1.0, frame.ScreenBounds.Height)), 0, frame.DetailHeight - 1);
            var detailX = Math.Clamp((int)((screenX - frame.ScreenBounds.Left) * frame.DetailWidth /
                Math.Max(1.0, frame.ScreenBounds.Width)), 0, frame.DetailWidth - 1);
            var source = detailY * frame.DetailWidth + detailX;
            for (var channel = 0; channel < Channels; channel++)
            {
                var value = color ? frame.DetailColorFingerprint[source * Channels + channel] : frame.DetailFingerprint[source];
                output[(y * Side + x) * Channels + channel] = (byte)(value / 8 * 8);
            }
        }
        return output;
    }

    internal static int PixelCount(byte[] fingerprint) => fingerprint.Length / Channels;

    static int Difference(byte[] left, byte[] right, int pixel)
    {
        var offset = pixel * Channels;
        return Math.Max(Math.Abs(left[offset] - right[offset]),
            Math.Max(Math.Abs(left[offset + 1] - right[offset + 1]), Math.Abs(left[offset + 2] - right[offset + 2])));
    }

    internal static (double MeanDelta, double ChangedRatio) Compare(
        byte[] left, byte[] right, bool[] volatilePixels, bool[]? auxiliaryPixels = null)
    {
        var count = Math.Min(volatilePixels.Length, Math.Min(PixelCount(left), PixelCount(right)));
        double sum = 0;
        var changed = 0;
        var compared = 0;
        for (var index = 0; index < count; index++)
        {
            if (volatilePixels[index] || auxiliaryPixels is not null && index < auxiliaryPixels.Length && auxiliaryPixels[index]) continue;
            var difference = Difference(left, right, index);
            sum += difference / 255.0;
            if (difference >= 8) changed++;
            compared++;
        }
        return compared == 0 ? (1, 1) : (sum / compared, changed / (double)compared);
    }

    internal static bool[] ChangedTiles(byte[] earlier, byte[] later, bool[] volatilePixels, int tileSide)
    {
        var tileCount = Side / tileSide;
        var result = new bool[tileCount * tileCount];
        var count = Math.Min(volatilePixels.Length, Math.Min(PixelCount(earlier), PixelCount(later)));
        for (var tileY = 0; tileY < tileCount; tileY++)
        for (var tileX = 0; tileX < tileCount; tileX++)
        {
            var changed = 0;
            for (var y = 0; y < tileSide; y++)
            for (var x = 0; x < tileSide; x++)
            {
                var index = (tileY * tileSide + y) * Side + tileX * tileSide + x;
                if (index < count && !volatilePixels[index] && Difference(earlier, later, index) >= 16) changed++;
            }
            result[tileY * tileCount + tileX] = changed >= 2;
        }
        return result;
    }

    internal static void LearnAmbient(byte[] earlier, byte[] later, bool[] volatilePixels)
    {
        var count = Math.Min(volatilePixels.Length, Math.Min(PixelCount(earlier), PixelCount(later)));
        for (var index = 0; index < count; index++)
        {
            if (Difference(earlier, later, index) < 8) continue;
            var centerX = index % Side;
            var centerY = index / Side;
            for (var dy = -2; dy <= 2; dy++)
            for (var dx = -2; dx <= 2; dx++)
            {
                var x = centerX + dx;
                var y = centerY + dy;
                if (x >= 0 && y >= 0 && x < Side && y < Side && y * Side + x < volatilePixels.Length)
                    volatilePixels[y * Side + x] = true;
            }
        }
    }
}