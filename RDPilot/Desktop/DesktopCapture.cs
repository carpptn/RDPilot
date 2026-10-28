
namespace RDPilot.Desktop;

/// <summary>Native screen capture boundary; tests can supply synthetic frames.</summary>
internal static class DesktopCapture
{
    internal interface IBackend
    {
        Rectangle Bounds(bool multipleMonitors);
        void CopyPixels(Graphics destination, Rectangle source);
    }

    sealed class WindowsCapture : IBackend
    {
        public Rectangle Bounds(bool multipleMonitors)
        {
            if (multipleMonitors)
            {
                var bounds = new Rectangle(GetSystemMetrics((int)SystemMetric.SM_XVIRTUALSCREEN),
                    GetSystemMetrics((int)SystemMetric.SM_YVIRTUALSCREEN),
                    GetSystemMetrics((int)SystemMetric.SM_CXVIRTUALSCREEN),
                    GetSystemMetrics((int)SystemMetric.SM_CYVIRTUALSCREEN));
                if (bounds.Width > 0 && bounds.Height > 0) return bounds;
            }
            return new Rectangle(0, 0, GetSystemMetrics((int)SystemMetric.SM_CXSCREEN), GetSystemMetrics((int)SystemMetric.SM_CYSCREEN));
        }
        public void CopyPixels(Graphics destination, Rectangle source) =>
            destination.CopyFromScreen(source.X, source.Y, 0, 0, source.Size, CopyPixelOperation.SourceCopy);
    }

    internal static IBackend Backend { get; set; } = new WindowsCapture();
}
