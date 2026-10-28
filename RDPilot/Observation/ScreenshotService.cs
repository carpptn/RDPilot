
namespace RDPilot.Observation;

    /// <summary>
    /// Captures, annotates, encodes, and validates desktop screenshots.
    /// </summary>
    internal static class ScreenshotService
    {
            // ====== Screenshot (PRIMARY) -> PNG + data URLs ======
            // Always draw the current UIA focus overlay on the full screenshot.
            internal static ScreenshotCaptureResult ScreenshotToDataUrl(
                    string saveDir,
                    string commandId,
                    int step,
                    Rectangle? explicitFocusRect,
                    bool includeObservationDetail = false)
            {
                var screenshotSw = RuntimeClock.StartTimer();
                var (vx, vy, vw, vh) = GetPrimaryScreen();
        
                using var bmp = new Bitmap(vw, vh, PixelFormat.Format24bppRgb);
                Rectangle? focusRectAbs = null;
                string? focusUiaSummary = null;
                if (Options.IncludeFocusUia)
                {
                    var uiaSw = RuntimeClock.StartTimer();
                    var focusSnapshot = CaptureFocusedUiaSnapshot();
                    focusRectAbs = focusSnapshot.Rect;
                    focusUiaSummary = focusSnapshot.Summary;
                    RecordUiaMetric(uiaSw);
                }
                using (var g = Graphics.FromImage(bmp))
                    DesktopCapture.Backend.CopyPixels(g, new Rectangle(vx, vy, vw, vh));
        
                // Delta/polling fingerprints must describe the real screen, not RDPilot's debug overlays.
                var deltaFingerprint = BuildImageFingerprint(bmp);
                var activeWindowFingerprint = BuildActiveWindowFingerprint(
                    bmp,
                    vx,
                    vy,
                    deltaFingerprint);
                var detailWidth = 0;
                var detailHeight = 0;
                var detailFingerprints = (Grayscale: Array.Empty<byte>(), Color: Array.Empty<byte>());
                if (includeObservationDetail)
                {
                    var detailScale = Math.Min(
                        1.0,
                        Math.Min(640.0 / Math.Max(1, vw), 360.0 / Math.Max(1, vh)));
                    detailWidth = Math.Max(1, (int)Math.Round(vw * detailScale));
                    detailHeight = Math.Max(1, (int)Math.Round(vh * detailScale));
                    detailFingerprints = BuildDetailFingerprints(
                        bmp,
                        detailWidth,
                        detailHeight);
                }
                var observationFrame = new ScreenObservationFrame(
                    deltaFingerprint,
                    activeWindowFingerprint,
                    detailFingerprints.Grayscale,
                    detailWidth,
                    detailHeight,
                    new Rectangle(vx, vy, vw, vh))
                {
                    DetailColorFingerprint = detailFingerprints.Color
                };
                ReportScreenshotSanity(deltaFingerprint, vw, vh);
        
                using (var g = Graphics.FromImage(bmp))
                {
                    // - UIA overlay (white+red rounded ring)
                    if (focusRectAbs is Rectangle frAbs)
                    {
                        // convert to primary-local coords (screenshot origin (vx,vy))
                        var r = Rectangle.FromLTRB(
                            frAbs.Left - vx - Options.FocusRingPadding,
                            frAbs.Top - vy - Options.FocusRingPadding,
                            frAbs.Right - vx + Options.FocusRingPadding,
                            frAbs.Bottom - vy + Options.FocusRingPadding);
                        r.Intersect(new Rectangle(0, 0, vw, vh));
        
                        using var path = RoundedRect(r, Options.FocusCornerRadius);
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        var largeFocusOverlay = (double)r.Width * r.Height / Math.Max(1, vw * vh) >= Options.LargeFocusOverlayAreaRatio;
                        using var glow = new Pen(Color.White, Options.FocusGlowThickness) { Alignment = PenAlignment.Center };
                        using var ring = new Pen(Color.Red, Options.FocusRingThickness) { Alignment = PenAlignment.Center };
                        if (!largeFocusOverlay)
                            g.DrawPath(glow, path);
                        g.DrawPath(ring, path);
                    }
        
                    // - Grid overlay (optional)
                    DrawGridOverlay(g, vw, vh);
                }
        
                string? fullPath = Options.LogScreens ? ScreenLogPath(saveDir, $"{commandId}_{step}") : null;
                var shareFullScreenLogEncoding = fullPath != null && CanShareFullScreenLogEncoding();
                if (fullPath != null && !shareFullScreenLogEncoding)
                    SaveScreenLogImage(bmp, fullPath);
        
                string? focusUrl = null;
                Rectangle? rect = null;
        
                if (Options.SendFocusCrop && explicitFocusRect is Rectangle rWant)
                {
                    var rOk = ClampRect(rWant);
                    var localCrop = new Rectangle(
                        rOk.Left - vx,
                        rOk.Top - vy,
                        rOk.Width,
                        rOk.Height);
                    using var crop = bmp.Clone(localCrop, PixelFormat.Format24bppRgb);
        
                    var cropPath = ScreenLogPath(saveDir, $"{commandId}_{step}_crop");
                    var shareCropLogEncoding = Options.LogScreens && CanShareCropScreenLogEncoding(crop);
                    if (Options.LogScreens && !shareCropLogEncoding)
                        SaveScreenLogImage(crop, cropPath);
        
                    focusUrl = shareCropLogEncoding
                        ? EncodeBitmapToDataUrl(crop, Options.MaxCropSendWidth, Options.CropSendFormat, Options.ScreenshotJpegQuality, cropPath)
                        : EncodeCropBitmapToDataUrl(crop);
                    rect = rOk;
                }
        
                var (fullImageW, fullImageH) = FullScreenshotRequestImageSize(bmp, focusUrl != null);
        
                if (fullPath != null && shareFullScreenLogEncoding && focusUrl != null && Options.FocusedOverviewMaxWidth > 0)
                    SaveScreenLogImage(bmp, fullPath);
                var fullDataUrl = EncodeFullScreenshotForRequest(bmp, fullPath, shareFullScreenLogEncoding, focusUrl != null);
        
                // small FOCUS_UIA crop (optional)
                string? focusUiaDataUrl = null;
                string? focusUiaPath = null;
                var focusUiaRectAbs = focusRectAbs;
                if (Options.IncludeFocusUia && Options.IncludeFocusUiaCrop && focusUiaRectAbs is Rectangle frAbs2)
                {
                    var clamped = ClampRect(frAbs2);
                    var local = new Rectangle(
                        clamped.Left - vx,
                        clamped.Top - vy,
                        clamped.Width,
                        clamped.Height);
                    if (Options.MaxFocusUiaCropPixels <= 0 || local.Width * local.Height <= Options.MaxFocusUiaCropPixels)
                    {
                        using var crop = bmp.Clone(local, PixelFormat.Format24bppRgb);
                        focusUiaPath = ScreenLogPath(saveDir, $"{commandId}_{step}_focus_uia");
                        var shareCropLogEncoding = Options.LogScreens && CanShareCropScreenLogEncoding(crop);
                        focusUiaDataUrl = shareCropLogEncoding
                            ? EncodeBitmapToDataUrl(crop, Options.MaxCropSendWidth, Options.CropSendFormat, Options.ScreenshotJpegQuality, focusUiaPath)
                            : EncodeCropBitmapToDataUrl(crop);
                        if (Options.LogScreens && !shareCropLogEncoding)
                            SaveScreenLogImage(crop, focusUiaPath);
                        if (!Options.LogScreens)
                            focusUiaPath = null;
                    }
                }
        
                screenshotSw.Stop();
                CurrentRun.RunScreenshotCount++;
                CurrentRun.RunScreenshotElapsed += screenshotSw.Elapsed;
                return new ScreenshotCaptureResult
                {
                    DataUrl = fullDataUrl,
                    SavedPath = fullPath,
                    ScreenWidth = vw,
                    ScreenHeight = vh,
                    ImageWidth = fullImageW,
                    ImageHeight = fullImageH,
                    FocusDataUrl = focusUrl,
                    FocusRect = rect,
                    FocusUiaRect = focusUiaRectAbs,
                    FocusUiaSummary = focusUiaSummary,
                    FocusUiaDataUrl = focusUiaDataUrl,
                    FocusUiaImagePath = focusUiaPath,
                    DeltaFingerprint = deltaFingerprint,
                    ActiveWindowFingerprint = activeWindowFingerprint,
                    ObservationFrame = observationFrame
                };
            }

            internal static Func<Rectangle?> ActiveWindowRectangleProvider { get; set; } = GetActiveWindowRectangle;

            internal static byte[] BuildActiveWindowFingerprint(
                Bitmap screenshot,
                int screenOriginX,
                int screenOriginY,
                byte[] fallback)
            {
                if (IsOwnConsoleForeground())
                    return fallback.ToArray();
                var activeRect = ActiveWindowRectangleProvider();
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

            internal static ScreenObservationFrame CaptureObservationFrameProbe(
                bool includeDetail = false)
            {
                var sw = RuntimeClock.StartTimer();
                try
                {
                    var (vx, vy, vw, vh) = GetPrimaryScreen();
                    using var bmp = new Bitmap(vw, vh, PixelFormat.Format24bppRgb);
                    using (var graphics = Graphics.FromImage(bmp))
                        DesktopCapture.Backend.CopyPixels(graphics, new Rectangle(vx, vy, vw, vh));

                    var global = BuildImageFingerprint(bmp);
                    var active = BuildActiveWindowFingerprint(bmp, vx, vy, global);
                    var detailWidth = 0;
                    var detailHeight = 0;
                    var detailFingerprints = (Grayscale: Array.Empty<byte>(), Color: Array.Empty<byte>());
                    if (includeDetail)
                    {
                        var scale = Math.Min(
                            1.0,
                            Math.Min(640.0 / Math.Max(1, vw), 360.0 / Math.Max(1, vh)));
                        detailWidth = Math.Max(1, (int)Math.Round(vw * scale));
                        detailHeight = Math.Max(1, (int)Math.Round(vh * scale));
                        detailFingerprints = BuildDetailFingerprints(
                            bmp,
                            detailWidth,
                            detailHeight);
                    }
                    return new ScreenObservationFrame(
                        global,
                        active,
                        detailFingerprints.Grayscale,
                        detailWidth,
                        detailHeight,
                        new Rectangle(vx, vy, vw, vh))
                    {
                        DetailColorFingerprint = detailFingerprints.Color
                    };
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunScreenProbeCount++;
                    CurrentRun.RunScreenProbeElapsed += sw.Elapsed;
                }
            }

            internal static byte[] CaptureRegionFingerprintProbe(Rectangle region)
            {
                var sw = RuntimeClock.StartTimer();
                try
                {
                    var clamped = ClampRect(region);
                    using var bitmap = new Bitmap(
                        clamped.Width,
                        clamped.Height,
                        PixelFormat.Format24bppRgb);
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        DesktopCapture.Backend.CopyPixels(graphics, clamped);
                    }

                    const int maximumSide = 192;
                    var scale = Math.Min(
                        1.0,
                        Math.Min(
                            maximumSide / (double)Math.Max(1, clamped.Width),
                            maximumSide / (double)Math.Max(1, clamped.Height)));
                    var width = Math.Max(1, (int)Math.Round(clamped.Width * scale));
                    var height = Math.Max(1, (int)Math.Round(clamped.Height * scale));
                    return BuildDetailFingerprints(bitmap, width, height).Color;
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunScreenProbeCount++;
                    CurrentRun.RunScreenProbeElapsed += sw.Elapsed;
                }
            }
        
            internal static (int W, int H) FullScreenshotRequestImageSize(Bitmap bmp, bool hasFocusCrop)
            {
                var maxWidth = hasFocusCrop && Options.FocusedOverviewMaxWidth > 0
                    ? EffectiveFocusedOverviewMaxWidth()
                    : Options.MaxScreenshotSendWidth;
                return EffectiveEncodedSize(bmp, maxWidth);
            }
        
            internal static string EncodeFullScreenshotForRequest(Bitmap bmp, string? fullPath, bool shareFullScreenLogEncoding, bool hasFocusCrop)
            {
                if (hasFocusCrop && Options.FocusedOverviewMaxWidth > 0)
                    return EncodeBitmapToDataUrl(bmp, EffectiveFocusedOverviewMaxWidth(), Options.ScreenshotSendFormat, Options.ScreenshotJpegQuality);
        
                return shareFullScreenLogEncoding
                    ? EncodeBitmapToDataUrl(bmp, Options.MaxScreenshotSendWidth, Options.ScreenshotSendFormat, Options.ScreenshotJpegQuality, fullPath)
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
                if (!Options.ScreenSanityChecks)
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
                    if (!CurrentRun.ReportedSanityWarnings.Add(warning))
                        continue;
        
                    CurrentRun.RunScreenSanityWarnings++;
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
                if (Options.GridStepPx <= 0) return;
        
                using var minorPen = new Pen(Color.FromArgb(80, 128, 128, 128), 2f);
                using var majorPen = new Pen(Color.FromArgb(140, 64, 64, 64), 4f);
                using var font = new Font("Consolas", 9, FontStyle.Regular, GraphicsUnit.Pixel);
                using var bg = new SolidBrush(Color.FromArgb(220, 0, 0, 0));
                using var fg = new SolidBrush(Color.White);
        
                g.SmoothingMode = SmoothingMode.None;
        
                // vertical lines
                for (int x = 0; x < w; x += Options.GridStepPx)
                {
                    var pen = (Options.GridMajorEveryPx > 0 && x % Options.GridMajorEveryPx == 0) ? majorPen : minorPen;
                    g.DrawLine(pen, x, 0, x, h);
        
                    if (Options.GridLabelEveryPx > 0 && x % Options.GridLabelEveryPx == 0)
                        DrawLabel(g, font, bg, fg, $"{x}", x + 2, 2);
                }
        
                // horizontal lines
                for (int y = 0; y < h; y += Options.GridStepPx)
                {
                    var pen = (Options.GridMajorEveryPx > 0 && y % Options.GridMajorEveryPx == 0) ? majorPen : minorPen;
                    g.DrawLine(pen, 0, y, w, y);
        
                    if (Options.GridLabelEveryPx > 0 && y % Options.GridLabelEveryPx == 0)
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
                    ? EncodeBitmapToDataUrl(bmp, Options.MaxScreenshotSendWidth, Options.ScreenshotSendFormat, Options.ScreenshotJpegQuality)
                    : EncodeBitmapToDataUrl(bmp, 0, "png", Options.ScreenshotJpegQuality);
            }
        
            internal static string EncodeCropBitmapToDataUrl(Bitmap bmp) =>
                EncodeBitmapToDataUrl(bmp, Options.MaxCropSendWidth, Options.CropSendFormat, Options.ScreenshotJpegQuality);

            internal static bool TryBuildTurnRegionEvidenceImage(
                string fullScreenDataUrl,
                Rectangle screenRegion,
                Rectangle screenBounds,
                string saveDir,
                string commandId,
                int step,
                out string? dataUrl,
                out string? imagePath)
            {
                dataUrl = null;
                imagePath = null;
                try
                {
                    screenRegion.Intersect(screenBounds);
                    if (screenRegion.Width <= 0 || screenRegion.Height <= 0)
                        return false;

                    var normalizedRegion = new RectangleF(
                        (screenRegion.Left - screenBounds.Left) /
                        (float)Math.Max(1, screenBounds.Width),
                        (screenRegion.Top - screenBounds.Top) /
                        (float)Math.Max(1, screenBounds.Height),
                        screenRegion.Width /
                        (float)Math.Max(1, screenBounds.Width),
                        screenRegion.Height /
                        (float)Math.Max(1, screenBounds.Height));
                    using var source = DecodeDataUrlBitmap(fullScreenDataUrl);
                    using var overviewCrop = CropNormalizedRegion(
                        source,
                        normalizedRegion);
                    using var nativeCrop = CaptureNativeScreenRegion(screenRegion);
                    var crop = nativeCrop is not null &&
                               AreTurnRegionCapturesConsistent(
                                   overviewCrop,
                                   nativeCrop)
                        ? nativeCrop
                        : overviewCrop;
                    if (nativeCrop is not null &&
                        ReferenceEquals(crop, overviewCrop))
                    {
                        Console.WriteLine(
                            "[turn-memory] native interaction-region capture changed after the primary screenshot; using the coherent overview crop for this turn.");
                    }
                    imagePath = Options.LogScreens
                        ? ScreenLogPath(saveDir, $"{commandId}_{step}_turn_region")
                        : null;
                    if (imagePath is not null)
                        SaveScreenLogImage(crop, imagePath);
                    dataUrl = EncodeCropBitmapToDataUrl(crop);
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(
                        $"[turn-memory] could not build automatic interaction-region image: {ex.Message}");
                    dataUrl = null;
                    imagePath = null;
                    return false;
                }
            }

            static Bitmap? CaptureNativeScreenRegion(Rectangle screenRegion)
            {
                Bitmap? crop = null;
                try
                {
                    crop = new Bitmap(screenRegion.Width, screenRegion.Height, PixelFormat.Format24bppRgb);
                    using var graphics = Graphics.FromImage(crop);
                    DesktopCapture.Backend.CopyPixels(graphics, screenRegion);
                    return crop;
                }
                catch
                {
                    crop?.Dispose();
                    return null;
                }
            }

            internal static bool AreTurnRegionCapturesConsistent(
                Bitmap earlier,
                Bitmap later)
            {
                if (earlier.Width <= 0 || earlier.Height <= 0 ||
                    later.Width <= 0 || later.Height <= 0)
                {
                    return false;
                }

                using var normalizedLater = new Bitmap(
                    earlier.Width,
                    earlier.Height,
                    PixelFormat.Format24bppRgb);
                using (var graphics = Graphics.FromImage(normalizedLater))
                {
                    graphics.InterpolationMode =
                        System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                    graphics.DrawImage(
                        later,
                        new Rectangle(0, 0, earlier.Width, earlier.Height));
                }

                var stepX = Math.Max(1, earlier.Width / 96);
                var stepY = Math.Max(1, earlier.Height / 96);
                long difference = 0;
                long samples = 0;
                for (var y = 0; y < earlier.Height; y += stepY)
                for (var x = 0; x < earlier.Width; x += stepX)
                {
                    var before = earlier.GetPixel(x, y);
                    var after = normalizedLater.GetPixel(x, y);
                    difference += Math.Abs(before.R - after.R) +
                                  Math.Abs(before.G - after.G) +
                                  Math.Abs(before.B - after.B);
                    samples++;
                }

                var normalizedDifference = difference /
                    (double)Math.Max(1, samples * 3 * 255);
                return normalizedDifference <= 0.08;
            }

            internal static TurnChangeImagePair? BuildTurnChangeImagePair(
                string beforeDataUrl,
                string afterDataUrl,
                RectangleF normalizedRegion,
                string saveDir,
                string commandId,
                int step,
                int regionIndex)
            {
                try
                {
                    using var beforeSource = DecodeDataUrlBitmap(beforeDataUrl);
                    using var afterSource = DecodeDataUrlBitmap(afterDataUrl);
                    using var beforeCrop = CropNormalizedRegion(beforeSource, normalizedRegion);
                    using var afterCrop = CropNormalizedRegion(afterSource, normalizedRegion);
                    var beforePath = Options.LogScreens
                        ? ScreenLogPath(saveDir, $"{commandId}_{step}_change_{regionIndex}_before")
                        : null;
                    var afterPath = Options.LogScreens
                        ? ScreenLogPath(saveDir, $"{commandId}_{step}_change_{regionIndex}_after")
                        : null;
                    if (beforePath is not null)
                        SaveScreenLogImage(beforeCrop, beforePath);
                    if (afterPath is not null)
                        SaveScreenLogImage(afterCrop, afterPath);
                    return new TurnChangeImagePair(
                        EncodeCropBitmapToDataUrl(beforeCrop),
                        EncodeCropBitmapToDataUrl(afterCrop),
                        beforePath,
                        afterPath,
                        regionIndex);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(
                        $"[turn-event] could not build change-region images: {ex.Message}");
                    return null;
                }
            }

            static Bitmap DecodeDataUrlBitmap(string dataUrl)
            {
                var separator = dataUrl.IndexOf(',');
                if (separator < 0 || separator + 1 >= dataUrl.Length)
                    throw new InvalidOperationException("invalid image data URL");
                var bytes = Convert.FromBase64String(dataUrl[(separator + 1)..]);
                using var stream = new MemoryStream(bytes, writable: false);
                using var image = Image.FromStream(stream);
                return new Bitmap(image);
            }

            static Bitmap CropNormalizedRegion(
                Bitmap source,
                RectangleF normalizedRegion)
            {
                const double paddingRatio = 0.04;
                var left = (int)Math.Floor(
                    (normalizedRegion.Left - paddingRatio) * source.Width);
                var top = (int)Math.Floor(
                    (normalizedRegion.Top - paddingRatio) * source.Height);
                var right = (int)Math.Ceiling(
                    (normalizedRegion.Right + paddingRatio) * source.Width);
                var bottom = (int)Math.Ceiling(
                    (normalizedRegion.Bottom + paddingRatio) * source.Height);
                var crop = Rectangle.FromLTRB(
                    Math.Clamp(left, 0, Math.Max(0, source.Width - 1)),
                    Math.Clamp(top, 0, Math.Max(0, source.Height - 1)),
                    Math.Clamp(right, 1, source.Width),
                    Math.Clamp(bottom, 1, source.Height));
                if (crop.Width < 2 || crop.Height < 2)
                    throw new InvalidOperationException("change region is too small");
                return source.Clone(crop, PixelFormat.Format24bppRgb);
            }
        
            internal sealed record PreparedImage(string DataUrl, int Width, int Height);

            internal static PreparedImage PrepareHelperImage(string dataUrl, int maxWidth)
            {
                using var bitmap = DecodeDataUrlBitmap(dataUrl);
                var (width, height) = EffectiveEncodedSize(bitmap, maxWidth);
                var encoded = ShouldResizeBitmap(bitmap, maxWidth)
                    ? EncodeBitmapToDataUrl(bitmap, maxWidth, Options.ScreenshotSendFormat, Options.ScreenshotJpegQuality)
                    : dataUrl;
                return new PreparedImage(encoded, width, height);
            }

            internal static bool CanShareFullScreenLogEncoding() =>
                Options.ScreenshotSendFormat.Equals(Options.ScreenLogFormat, StringComparison.OrdinalIgnoreCase) &&
                Options.MaxScreenshotSendWidth == Options.MaxScreenLogWidth;
        
            internal static bool CanShareCropScreenLogEncoding(Bitmap crop) =>
                Options.CropSendFormat.Equals(Options.ScreenLogFormat, StringComparison.OrdinalIgnoreCase) &&
                EffectiveEncodedWidth(crop, Options.MaxCropSendWidth) == EffectiveEncodedWidth(crop, Options.MaxScreenLogWidth);
        
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
                var sw = RuntimeClock.StartTimer();
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
                    CurrentRun.RunImageEncodeCount++;
                    CurrentRun.RunImageEncodeElapsed += sw.Elapsed;
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
                Options.ScreenLogFormat.Equals("jpeg", StringComparison.OrdinalIgnoreCase) ? ".jpg" : ".png";
        
            internal static string ScreenLogPath(string dir, string baseName) =>
                Path.Combine(dir, baseName + ScreenLogExtension());
        
            internal static void SaveScreenLogImage(Bitmap bmp, string path)
            {
                var sw = RuntimeClock.StartTimer();
                try
                {
                    if (ShouldResizeBitmap(bmp, Options.MaxScreenLogWidth))
                    {
                        using var outbound = PrepareOutboundBitmap(bmp, Options.MaxScreenLogWidth);
                        SaveScreenLogImageCore(outbound, path);
                        return;
                    }
        
                    SaveScreenLogImageCore(bmp, path);
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunScreenLogCount++;
                    CurrentRun.RunScreenLogElapsed += sw.Elapsed;
                }
            }
        
            internal static void SaveEncodedScreenLogImage(MemoryStream ms, string path)
            {
                var sw = RuntimeClock.StartTimer();
                try
                {
                    using var fs = File.Create(path);
                    fs.Write(ms.GetBuffer(), 0, (int)ms.Length);
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunScreenLogCount++;
                    CurrentRun.RunScreenLogElapsed += sw.Elapsed;
                }
            }
        
            internal static void SaveScreenLogImageCore(Bitmap bmp, string path)
            {
                if (Options.ScreenLogFormat.Equals("jpeg", StringComparison.OrdinalIgnoreCase))
                {
                    using var fs = File.Create(path);
                    SaveJpeg(bmp, fs, Options.ScreenshotJpegQuality);
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
