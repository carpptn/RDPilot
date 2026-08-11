internal static partial class RDPilotApplication
{
    /// <summary>
    /// Builds reduced image fingerprints and compares screen changes.
    /// </summary>
    internal static class ImageComparisonService
    {
            // ==== Image delta (0..1) ====
            internal static byte[] BuildImageFingerprint(Bitmap bmp, int w = 96, int h = 54)
            {
                using var resized = ResizeToFast(bmp, w, h);
                var rect = new Rectangle(0, 0, w, h);
                var bits = resized.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                try
                {
                    var stride = Math.Abs(bits.Stride);
                    var buffer = new byte[stride * h];
                    Marshal.Copy(bits.Scan0, buffer, 0, buffer.Length);
        
                    var fingerprint = new byte[w * h];
                    var i = 0;
                    for (int y = 0; y < h; y++)
                    {
                        var rowOffset = bits.Stride >= 0 ? y * bits.Stride : (h - 1 - y) * stride;
                        for (int x = 0; x < w; x++)
                        {
                            var px = rowOffset + (x * 3);
                            var b = buffer[px];
                            var g = buffer[px + 1];
                            var r = buffer[px + 2];
                            fingerprint[i++] = (byte)Math.Round(0.299 * r + 0.587 * g + 0.114 * b);
                        }
                    }
        
                    return fingerprint;
                }
                finally
                {
                    resized.UnlockBits(bits);
                }
            }
        
            internal static double ComputeImageDelta(byte[] a, byte[] b)
            {
                var n = Math.Min(a.Length, b.Length);
                if (n == 0)
                    return 1.0;
        
                double sum = 0;
                for (var i = 0; i < n; i++)
                    sum += Math.Abs(a[i] - b[i]) / 255.0;
        
                return sum / n;
            }
        
            internal static double ComputeImageDelta(string pathA, string pathB, int w = 96, int h = 54)
            {
                using var a = new Bitmap(pathA);
                using var b = new Bitmap(pathB);
                return ComputeImageDelta(BuildImageFingerprint(a, w, h), BuildImageFingerprint(b, w, h));
            }
        
            internal static Bitmap ResizeTo(Bitmap src, int w, int h)
            {
                var dst = new Bitmap(w, h, PixelFormat.Format24bppRgb);
                using var g = Graphics.FromImage(dst);
                g.InterpolationMode = InterpolationMode.HighQualityBilinear;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.SmoothingMode = SmoothingMode.None;
                g.DrawImage(src, new Rectangle(0, 0, w, h));
                return dst;
            }
        
            internal static Bitmap ResizeToFast(Bitmap src, int w, int h)
            {
                var dst = new Bitmap(w, h, PixelFormat.Format24bppRgb);
                using var g = Graphics.FromImage(dst);
                g.CompositingQuality = CompositingQuality.HighSpeed;
                g.InterpolationMode = InterpolationMode.Low;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.SmoothingMode = SmoothingMode.None;
                g.DrawImage(src, new Rectangle(0, 0, w, h));
                return dst;
            }
    }
}

