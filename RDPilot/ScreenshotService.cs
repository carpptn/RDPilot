internal static partial class RDPilotApplication
{
    /// <summary>
    /// Captures, annotates, encodes, and validates desktop screenshots.
    /// </summary>
    internal static class ScreenshotService
    {
            // ====== Screenshot (PRIMARY) -> PNG + data URLs ======
            // Always draw the current UIA focus overlay on the full screenshot.
            internal static (string dataUrl, string? savedPath, int screenW, int screenH, int imageW, int imageH,
                    string? focusDataUrl, Rectangle? focusRect,
                    Rectangle? focusUiaRect, string? focusUiaSummary, string? focusUiaDataUrl, string? focusUiaImagePath,
                    byte[] deltaFingerprint, byte[] activeWindowFingerprint)
                ScreenshotToDataUrl(string saveDir, string commandId, int step, Rectangle? explicitFocusRect)
            {
                var screenshotSw = Stopwatch.StartNew();
                var (vx, vy, vw, vh) = GetPrimaryScreen();
        
                using var bmp = new Bitmap(vw, vh, PixelFormat.Format24bppRgb);
                Rectangle? focusRectAbs = null;
                string? focusUiaSummary = null;
                if (IncludeFocusUia)
                {
                    var uiaSw = Stopwatch.StartNew();
                    var focusSnapshot = CaptureFocusedUiaSnapshot();
                    focusRectAbs = focusSnapshot.Rect;
                    focusUiaSummary = focusSnapshot.Summary;
                    RecordUiaMetric(uiaSw);
                }
                using (var g = Graphics.FromImage(bmp))
                    g.CopyFromScreen(vx, vy, 0, 0, new Size(vw, vh), CopyPixelOperation.SourceCopy);
        
                // Delta/polling fingerprints must describe the real screen, not RDPilot's debug overlays.
                var deltaFingerprint = BuildImageFingerprint(bmp);
                var activeWindowFingerprint = BuildActiveWindowFingerprint(
                    bmp,
                    vx,
                    vy,
                    deltaFingerprint);
                ReportScreenshotSanity(deltaFingerprint, vw, vh);
        
                using (var g = Graphics.FromImage(bmp))
                {
                    // - UIA overlay (white+red rounded ring)
                    if (focusRectAbs is Rectangle frAbs)
                    {
                        // convert to primary-local coords (screenshot origin (vx,vy))
                        var r = Rectangle.FromLTRB(
                            frAbs.Left - vx - FocusRingPadding,
                            frAbs.Top - vy - FocusRingPadding,
                            frAbs.Right - vx + FocusRingPadding,
                            frAbs.Bottom - vy + FocusRingPadding);
                        r.Intersect(new Rectangle(0, 0, vw, vh));
        
                        using var path = RoundedRect(r, FocusCornerRadius);
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        var largeFocusOverlay = (double)r.Width * r.Height / Math.Max(1, vw * vh) >= LargeFocusOverlayAreaRatio;
                        using var glow = new Pen(Color.White, FocusGlowThickness) { Alignment = PenAlignment.Center };
                        using var ring = new Pen(Color.Red, FocusRingThickness) { Alignment = PenAlignment.Center };
                        if (!largeFocusOverlay)
                            g.DrawPath(glow, path);
                        g.DrawPath(ring, path);
                    }
        
                    // - Grid overlay (optional)
                    DrawGridOverlay(g, vw, vh);
                }
        
                string? fullPath = LogScreens ? ScreenLogPath(saveDir, $"{commandId}_{step}") : null;
                var shareFullScreenLogEncoding = fullPath != null && CanShareFullScreenLogEncoding();
                if (fullPath != null && !shareFullScreenLogEncoding)
                    SaveScreenLogImage(bmp, fullPath);
        
                string? focusUrl = null;
                Rectangle? rect = null;
        
                if (SendFocusCrop && explicitFocusRect is Rectangle rWant)
                {
                    var rOk = ClampRect(rWant);
                    var localCrop = new Rectangle(
                        rOk.Left - vx,
                        rOk.Top - vy,
                        rOk.Width,
                        rOk.Height);
                    using var crop = bmp.Clone(localCrop, PixelFormat.Format24bppRgb);
        
                    var cropPath = ScreenLogPath(saveDir, $"{commandId}_{step}_crop");
                    var shareCropLogEncoding = LogScreens && CanShareCropScreenLogEncoding(crop);
                    if (LogScreens && !shareCropLogEncoding)
                        SaveScreenLogImage(crop, cropPath);
        
                    focusUrl = shareCropLogEncoding
                        ? EncodeBitmapToDataUrl(crop, MaxCropSendWidth, CropSendFormat, ScreenshotJpegQuality, cropPath)
                        : EncodeCropBitmapToDataUrl(crop);
                    rect = rOk;
                }
        
                var (fullImageW, fullImageH) = FullScreenshotRequestImageSize(bmp, focusUrl != null);
        
                if (fullPath != null && shareFullScreenLogEncoding && focusUrl != null && FocusedOverviewMaxWidth > 0)
                    SaveScreenLogImage(bmp, fullPath);
                var fullDataUrl = EncodeFullScreenshotForRequest(bmp, fullPath, shareFullScreenLogEncoding, focusUrl != null);
        
                // small FOCUS_UIA crop (optional)
                string? focusUiaDataUrl = null;
                string? focusUiaPath = null;
                var focusUiaRectAbs = focusRectAbs;
                if (IncludeFocusUia && IncludeFocusUiaCrop && focusUiaRectAbs is Rectangle frAbs2)
                {
                    var clamped = ClampRect(frAbs2);
                    var local = new Rectangle(
                        clamped.Left - vx,
                        clamped.Top - vy,
                        clamped.Width,
                        clamped.Height);
                    if (MaxFocusUiaCropPixels <= 0 || local.Width * local.Height <= MaxFocusUiaCropPixels)
                    {
                        using var crop = bmp.Clone(local, PixelFormat.Format24bppRgb);
                        focusUiaPath = ScreenLogPath(saveDir, $"{commandId}_{step}_focus_uia");
                        var shareCropLogEncoding = LogScreens && CanShareCropScreenLogEncoding(crop);
                        focusUiaDataUrl = shareCropLogEncoding
                            ? EncodeBitmapToDataUrl(crop, MaxCropSendWidth, CropSendFormat, ScreenshotJpegQuality, focusUiaPath)
                            : EncodeCropBitmapToDataUrl(crop);
                        if (LogScreens && !shareCropLogEncoding)
                            SaveScreenLogImage(crop, focusUiaPath);
                        if (!LogScreens)
                            focusUiaPath = null;
                    }
                }
        
                screenshotSw.Stop();
                RunScreenshotCount++;
                RunScreenshotElapsed += screenshotSw.Elapsed;
                return (fullDataUrl, fullPath, vw, vh, fullImageW, fullImageH, focusUrl, rect, focusUiaRectAbs, focusUiaSummary, focusUiaDataUrl, focusUiaPath, deltaFingerprint, activeWindowFingerprint);
            }

            internal static byte[] BuildActiveWindowFingerprint(
                Bitmap screenshot,
                int screenOriginX,
                int screenOriginY,
                byte[] fallback)
            {
                if (IsOwnConsoleForeground())
                    return fallback.ToArray();
                var activeRect = GetActiveWindowRectangle();
                if (activeRect is not Rectangle absolute)
                    return fallback.ToArray();

                var local = new Rectangle(
                    absolute.Left - screenOriginX,
                    absolute.Top - screenOriginY,
                    absolute.Width,
                    absolute.Height);
                local.Intersect(new Rectangle(0, 0, screenshot.Width, screenshot.Height));
                if (local.Width < 16 || local.Height < 16)
                    return fallback.ToArray();

                using var crop = screenshot.Clone(local, PixelFormat.Format24bppRgb);
                return BuildImageFingerprint(crop);
            }
        
            internal static (int W, int H) FullScreenshotRequestImageSize(Bitmap bmp, bool hasFocusCrop)
            {
                var maxWidth = hasFocusCrop && FocusedOverviewMaxWidth > 0
                    ? EffectiveFocusedOverviewMaxWidth()
                    : MaxScreenshotSendWidth;
                return EffectiveEncodedSize(bmp, maxWidth);
            }
        
            internal static string EncodeFullScreenshotForRequest(Bitmap bmp, string? fullPath, bool shareFullScreenLogEncoding, bool hasFocusCrop)
            {
                if (hasFocusCrop && FocusedOverviewMaxWidth > 0)
                    return EncodeBitmapToDataUrl(bmp, EffectiveFocusedOverviewMaxWidth(), ScreenshotSendFormat, ScreenshotJpegQuality);
        
                return shareFullScreenLogEncoding
                    ? EncodeBitmapToDataUrl(bmp, MaxScreenshotSendWidth, ScreenshotSendFormat, ScreenshotJpegQuality, fullPath)
                    : EncodeBitmapToDataUrl(bmp, applySendProfile: true);
            }
        
            internal static void SaveAimOverlay(string baseShotPath, Rectangle rect, string overlayPath)
            {
                using var bmp = new Bitmap(baseShotPath);
                using var g = Graphics.FromImage(bmp);
                g.SmoothingMode = SmoothingMode.None;
                var (_, _, screenW, screenH) = GetPrimaryScreen();
                var scaleX = screenW > 0 ? bmp.Width / (double)screenW : 1.0;
                var scaleY = screenH > 0 ? bmp.Height / (double)screenH : 1.0;
                var (screenX, screenY, _, _) = GetPrimaryScreen();
        
                using var penYellow = new Pen(Color.Yellow, 3f) { DashStyle = DashStyle.Dash };
                using var penBlue = new Pen(Color.DeepSkyBlue, 1.5f) { DashStyle = DashStyle.Dash };
        
                var r = Rectangle.FromLTRB(
                    (int)Math.Round((rect.Left - screenX) * scaleX),
                    (int)Math.Round((rect.Top - screenY) * scaleY),
                    (int)Math.Round((rect.Right - screenX - 1) * scaleX),
                    (int)Math.Round((rect.Bottom - screenY - 1) * scaleY));
                r.Intersect(new Rectangle(0, 0, bmp.Width, bmp.Height));
                g.DrawRectangle(penYellow, r);
                var r2 = Rectangle.FromLTRB(Math.Max(0, r.Left - 1), Math.Max(0, r.Top - 1),
                                            Math.Min(bmp.Width - 1, r.Right + 1), Math.Min(bmp.Height - 1, r.Bottom + 1));
                g.DrawRectangle(penBlue, r2);
        
                bmp.Save(overlayPath, ImageFormat.Png);
            }
        
            internal static void ReportScreenshotSanity(byte[] fingerprint, int screenW, int screenH)
            {
                if (!ScreenSanityChecks)
                    return;
        
                var warnings = new List<string>();
                if (screenW < 640 || screenH < 480)
                    warnings.Add($"unusual primary screen size {screenW}x{screenH}");
        
                if (fingerprint.Length > 0)
                {
                    var min = 255;
                    var max = 0;
                    long sum = 0;
                    foreach (var px in fingerprint)
                    {
                        if (px < min) min = px;
                        if (px > max) max = px;
                        sum += px;
                    }
        
                    var avg = sum / (double)fingerprint.Length;
                    var range = max - min;
                    if (avg < 4)
                        warnings.Add("screenshot appears nearly black");
                    else if (avg > 251)
                        warnings.Add("screenshot appears nearly white");
                    else if (range < 4)
                        warnings.Add("screenshot appears nearly uniform");
                }
        
                if (IsOwnConsoleForeground())
                    warnings.Add("RDPilot console is the foreground window; it may cover the target UI");
        
                foreach (var warning in warnings)
                {
                    if (!ReportedSanityWarnings.Add(warning))
                        continue;
        
                    RunScreenSanityWarnings++;
                    Console.WriteLine($"[sanity] {warning}");
                }
            }
        
            internal static bool IsOwnConsoleForeground()
            {
                try
                {
                    var console = GetConsoleWindow();
                    return console != IntPtr.Zero && console == GetForegroundWindow();
                }
                catch
                {
                    return false;
                }
            }
        
            internal static GraphicsPath RoundedRect(Rectangle r, int radius)
            {
                var p = new GraphicsPath();
                if (radius <= 0) { p.AddRectangle(r); return p; }
                int d = radius * 2;
                p.AddArc(r.Left, r.Top, d, d, 180, 90);
                p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
                p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                p.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
                p.CloseFigure();
                return p;
            }
        
            internal static void DrawGridOverlay(Graphics g, int w, int h)
            {
                if (GridStepPx <= 0) return;
        
                using var minorPen = new Pen(Color.FromArgb(80, 128, 128, 128), 2f);
                using var majorPen = new Pen(Color.FromArgb(140, 64, 64, 64), 4f);
                using var font = new Font("Consolas", 9, FontStyle.Regular, GraphicsUnit.Pixel);
                using var bg = new SolidBrush(Color.FromArgb(220, 0, 0, 0));
                using var fg = new SolidBrush(Color.White);
        
                g.SmoothingMode = SmoothingMode.None;
        
                // vertical lines
                for (int x = 0; x < w; x += GridStepPx)
                {
                    var pen = (GridMajorEveryPx > 0 && x % GridMajorEveryPx == 0) ? majorPen : minorPen;
                    g.DrawLine(pen, x, 0, x, h);
        
                    if (GridLabelEveryPx > 0 && x % GridLabelEveryPx == 0)
                        DrawLabel(g, font, bg, fg, $"{x}", x + 2, 2);
                }
        
                // horizontal lines
                for (int y = 0; y < h; y += GridStepPx)
                {
                    var pen = (GridMajorEveryPx > 0 && y % GridMajorEveryPx == 0) ? majorPen : minorPen;
                    g.DrawLine(pen, 0, y, w, y);
        
                    if (GridLabelEveryPx > 0 && y % GridLabelEveryPx == 0)
                        DrawLabel(g, font, bg, fg, $"{y}", 2, y + 2);
                }
        
                static void DrawLabel(Graphics g, Font font, Brush bg, Brush fg, string text, int x, int y)
                {
                    var sz = g.MeasureString(text, font, new SizeF(100, 20), StringFormat.GenericTypographic);
                    var rect = new RectangleF(x, y, sz.Width + 4, sz.Height + 2);
                    g.FillRectangle(bg, rect);
                    g.DrawString(text, font, fg, x + 2, y + 1, StringFormat.GenericTypographic);
                }
            }
        
            internal static string EncodeBitmapToDataUrl(Bitmap bmp, bool applySendProfile = true)
            {
                return applySendProfile
                    ? EncodeBitmapToDataUrl(bmp, MaxScreenshotSendWidth, ScreenshotSendFormat, ScreenshotJpegQuality)
                    : EncodeBitmapToDataUrl(bmp, 0, "png", ScreenshotJpegQuality);
            }
        
            internal static string EncodeCropBitmapToDataUrl(Bitmap bmp) =>
                EncodeBitmapToDataUrl(bmp, MaxCropSendWidth, CropSendFormat, ScreenshotJpegQuality);
        
            internal static string DownscaleDataUrlForHelperCall(string dataUrl, string? savedPath, int helperMaxWidth)
            {
                if (helperMaxWidth <= 0 || string.IsNullOrWhiteSpace(savedPath) || !File.Exists(savedPath))
                    return dataUrl;
        
                try
                {
                    using var bmp = new Bitmap(savedPath);
                    if (!ShouldResizeBitmap(bmp, helperMaxWidth))
                        return dataUrl;
                    return EncodeBitmapToDataUrl(bmp, helperMaxWidth, ScreenshotSendFormat, ScreenshotJpegQuality);
                }
                catch
                {
                    return dataUrl;
                }
            }
        
            internal static (int W, int H) HelperImageSize(string? savedPath, int fallbackW, int fallbackH, int helperMaxWidth)
            {
                if (helperMaxWidth <= 0 || string.IsNullOrWhiteSpace(savedPath) || !File.Exists(savedPath))
                    return (fallbackW, fallbackH);
        
                try
                {
                    using var bmp = new Bitmap(savedPath);
                    return EffectiveEncodedSize(bmp, helperMaxWidth);
                }
                catch
                {
                    return (fallbackW, fallbackH);
                }
            }
        
            internal static bool CanShareFullScreenLogEncoding() =>
                ScreenshotSendFormat.Equals(ScreenLogFormat, StringComparison.OrdinalIgnoreCase) &&
                MaxScreenshotSendWidth == MaxScreenLogWidth;
        
            internal static bool CanShareCropScreenLogEncoding(Bitmap crop) =>
                CropSendFormat.Equals(ScreenLogFormat, StringComparison.OrdinalIgnoreCase) &&
                EffectiveEncodedWidth(crop, MaxCropSendWidth) == EffectiveEncodedWidth(crop, MaxScreenLogWidth);
        
            internal static int EffectiveEncodedWidth(Bitmap bmp, int maxWidth) =>
                ShouldResizeBitmap(bmp, maxWidth) ? maxWidth : bmp.Width;
        
            internal static (int W, int H) EffectiveEncodedSize(Bitmap bmp, int maxWidth)
            {
                if (!ShouldResizeBitmap(bmp, maxWidth))
                    return (bmp.Width, bmp.Height);
        
                var scale = (double)maxWidth / bmp.Width;
                return (maxWidth, Math.Max(1, (int)Math.Round(bmp.Height * scale)));
            }
        
            internal static string EncodeBitmapToDataUrl(Bitmap bmp, int maxWidth, string format, long jpegQuality, string? writeEncodedImagePath = null)
            {
                var sw = Stopwatch.StartNew();
                try
                {
                    using var resized = ShouldResizeBitmap(bmp, maxWidth) ? PrepareOutboundBitmap(bmp, maxWidth) : null;
                    var outbound = resized ?? bmp;
                    using var ms = new MemoryStream();
        
                    if (format.Equals("jpeg", StringComparison.OrdinalIgnoreCase))
                    {
                        SaveJpeg(outbound, ms, jpegQuality);
                        if (writeEncodedImagePath != null)
                            SaveEncodedScreenLogImage(ms, writeEncodedImagePath);
                        return DataUrlFromMemoryStream("image/jpeg", ms);
                    }
        
                    outbound.Save(ms, ImageFormat.Png);
                    if (writeEncodedImagePath != null)
                        SaveEncodedScreenLogImage(ms, writeEncodedImagePath);
                    return DataUrlFromMemoryStream("image/png", ms);
                }
                finally
                {
                    sw.Stop();
                    RunImageEncodeCount++;
                    RunImageEncodeElapsed += sw.Elapsed;
                }
            }
        
            internal static bool ShouldResizeBitmap(Bitmap bmp, int maxWidth) =>
                maxWidth > 0 && bmp.Width > maxWidth;
        
            internal static string DataUrlFromMemoryStream(string mimeType, MemoryStream ms) =>
                $"data:{mimeType};base64,{Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length)}";
        
            internal static Bitmap PrepareOutboundBitmap(Bitmap bmp, int maxWidth)
            {
                if (maxWidth <= 0 || bmp.Width <= maxWidth)
                    return new Bitmap(bmp);
        
                var scale = (double)maxWidth / bmp.Width;
                var targetW = maxWidth;
                var targetH = Math.Max(1, (int)Math.Round(bmp.Height * scale));
                return ResizeTo(bmp, targetW, targetH);
            }
        
            internal static void SaveJpeg(Bitmap bmp, Stream stream, long quality)
            {
                var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => string.Equals(c.MimeType, "image/jpeg", StringComparison.OrdinalIgnoreCase));
                if (codec is null)
                {
                    bmp.Save(stream, ImageFormat.Jpeg);
                    return;
                }
        
                using var encoderParameters = new EncoderParameters(1);
                encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, Math.Clamp(quality, 1L, 100L));
                bmp.Save(stream, codec, encoderParameters);
            }
        
            internal static string ScreenLogExtension() =>
                ScreenLogFormat.Equals("jpeg", StringComparison.OrdinalIgnoreCase) ? ".jpg" : ".png";
        
            internal static string ScreenLogPath(string dir, string baseName) =>
                Path.Combine(dir, baseName + ScreenLogExtension());
        
            internal static void SaveScreenLogImage(Bitmap bmp, string path)
            {
                var sw = Stopwatch.StartNew();
                try
                {
                    if (ShouldResizeBitmap(bmp, MaxScreenLogWidth))
                    {
                        using var outbound = PrepareOutboundBitmap(bmp, MaxScreenLogWidth);
                        SaveScreenLogImageCore(outbound, path);
                        return;
                    }
        
                    SaveScreenLogImageCore(bmp, path);
                }
                finally
                {
                    sw.Stop();
                    RunScreenLogCount++;
                    RunScreenLogElapsed += sw.Elapsed;
                }
            }
        
            internal static void SaveEncodedScreenLogImage(MemoryStream ms, string path)
            {
                var sw = Stopwatch.StartNew();
                try
                {
                    using var fs = File.Create(path);
                    fs.Write(ms.GetBuffer(), 0, (int)ms.Length);
                }
                finally
                {
                    sw.Stop();
                    RunScreenLogCount++;
                    RunScreenLogElapsed += sw.Elapsed;
                }
            }
        
            internal static void SaveScreenLogImageCore(Bitmap bmp, string path)
            {
                if (ScreenLogFormat.Equals("jpeg", StringComparison.OrdinalIgnoreCase))
                {
                    using var fs = File.Create(path);
                    SaveJpeg(bmp, fs, ScreenshotJpegQuality);
                    return;
                }
        
                bmp.Save(path, ImageFormat.Png);
            }
        
            internal static Rectangle SquareAround(int cx, int cy, int size)
            {
                int half = size / 2;
                return new Rectangle(cx - half, cy - half, size, size);
            }
        
            internal static Rectangle ClampRect(Rectangle r)
            {
                var (x, y, w, h) = GetPrimaryScreen();
                int left = Math.Max(x, Math.Min(x + w - 1, r.Left));
                int top = Math.Max(y, Math.Min(y + h - 1, r.Top));
                int right = Math.Max(left + 1, Math.Min(x + w, r.Right));
                int bottom = Math.Max(top + 1, Math.Min(y + h, r.Bottom));
                return Rectangle.FromLTRB(left, top, right, bottom);
            }
    }
}



