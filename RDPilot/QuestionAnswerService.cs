internal static partial class RDPilotApplication
{
    /// <summary>
    /// Answers screenshot-based questions without executing desktop actions.
    /// </summary>
    internal static class QuestionAnswerService
    {
            // === Q&A detection ===
            internal static bool IsQuestion(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return false;
                s = s.Trim();
                if (s.StartsWith("/ask ", StringComparison.OrdinalIgnoreCase)) return true;
                if (s.EndsWith("?")) return true;
                return false;
            }
        
            // === Q&A ===
            internal static async Task RunAskOnce(string apiKey, string question)
            {
                var commandId = Guid.NewGuid().ToString("N");
                var screensDir = EnsureScreensDir();
                var requestsDir = EnsureRequestsDir();
                var logDir = EnsureLogDir();
        
                var prevOut = Console.Out;
                var prevErr = Console.Error;
                var logPath = Path.Combine(logDir, $"{commandId}.log");
                using var logFile = new StreamWriter(logPath, append: false, Encoding.UTF8) { AutoFlush = true };
                using var tee = new TeeTextWriter(prevOut, logFile);
                Console.SetOut(tee);
                Console.SetError(tee);
        
                CancellationTokenSource? cancelCts = null;
                var consoleHidden = false;
                try
                {
                    ResetRunMetrics();
                    Console.WriteLine($"[Q&A] ID: {commandId}");
                    Console.WriteLine($"Question: {question}");
                    cancelCts = StartCancelHotkeyListener();
        
                    if (AutoHideConsoleDuringRun && IsOwnConsoleForeground())
                        consoleHidden = ConcealConsoleWindow();
        
                    var (dataUrl, savedPath, screenW, screenH, imageW, imageH, _, _, focusUiaRect, focusUiaSummary, focusUiaDataUrl, focusUiaPath, _, _) =
                        ScreenshotToDataUrl(screensDir, commandId, 1, null);
                    dataUrl = DownscaleDataUrlForHelperCall(dataUrl, savedPath, QaScreenshotMaxWidth);
                    (imageW, imageH) = HelperImageSize(savedPath, imageW, imageH, QaScreenshotMaxWidth);
                    SetCurrentScreenMap(screenW, screenH, imageW, imageH);
                    Console.WriteLine($"[shot] {ShotLabel(savedPath, commandId, 1)}");
        
                    var (screenCx, screenCy, _, _) = GetCursorPositionInPrimary();
                    var (cx, cy, cnx, cny) = CursorToImageCoordinates(screenCx, screenCy);
                    var promptContext = CaptureUiPromptContext(focusUiaSummary, screenW, screenH);
                    var focusUiaRectForPrompt = ScreenRectToImage(focusUiaRect);
        
                    var systemRules = BuildQaSystemRules();
                    var qaModel = EffectiveQaModel();
                    var reqBody = BuildQARequestBody(qaModel, systemRules, question, dataUrl, imageW, imageH, cx, cy, cnx, cny, focusUiaRectForPrompt, focusUiaDataUrl, promptContext);
                    if (LogRequests)
                    {
                        var reqBodyForLog = BuildQARequestBody_ForLog(qaModel, systemRules, question, savedPath, imageW, imageH, cx, cy, cnx, cny, focusUiaRectForPrompt, focusUiaPath, promptContext);
                        SaveJson(Path.Combine(requestsDir, $"{commandId}_qa_request.json"), reqBodyForLog);
                    }
        
                    var (qa, raw) = await CallOpenAIParsedAsync<QaLocateDto>(apiKey, reqBody, cancelCts.Token);
                    SaveRaw(Path.Combine(requestsDir, $"{commandId}_qa_response.json"), raw);
        
                    if (qa == null)
                    {
                        Console.WriteLine(CancelRequested ? "Aborted (hotkey)." : "No response.");
                        return;
                    }
        
                    if (!string.IsNullOrWhiteSpace(qa.AnswerText))
                    {
                        Console.WriteLine("🟩 Answer:");
                        Console.WriteLine(qa.AnswerText.Trim());
                    }
                    if (!string.IsNullOrWhiteSpace(qa.Note))
                        Console.WriteLine($"ℹ️ note: {qa.Note}");
        
                    if (qa.X.HasValue && qa.Y.HasValue)
                    {
                        int px = (int)Math.Round(qa.X.Value * (screenW - 1));
                        int py = (int)Math.Round(qa.Y.Value * (screenH - 1));
                        Console.WriteLine($"📍 Location (from 0..1): {px}:{py}");
                    }
                    else if (qa.XPx.HasValue && qa.YPx.HasValue)
                    {
                        Console.WriteLine($"📍 Location (hint x_px/y_px): {qa.XPx}:{qa.YPx}");
                    }
        
                    if (qa.BBox is not null)
                        Console.WriteLine($"🧰 bbox=({qa.BBox.Left},{qa.BBox.Top})–({qa.BBox.Right},{qa.BBox.Bottom})");
                }
                finally
                {
                    cancelCts?.Cancel();
                    cancelCts?.Dispose();
                    PrintRunMetrics();
                    Console.SetOut(prevOut);
                    Console.SetError(prevErr);
                    if (consoleHidden && RestoreConsoleAfterRun)
                        RestoreConsoleWindow();
                }
            }
    }
}

