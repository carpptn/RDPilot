
namespace RDPilot.Control;

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
            internal static async Task<CommandOutcome> RunAskOnce(string apiKey, string question)
            {
                using var runScope = BeginRun();
                var commandId = Guid.NewGuid().ToString("N");
                var screensDir = EnsureScreensDir();
                var requestsDir = EnsureRequestsDir();
                var logDir = EnsureLogDir();
        
                using var logSession = new RunLogSession(Path.Combine(logDir, $"{commandId}.log"));

                CancellationMonitor? cancelCts = null;
                var consoleHidden = false;
                try
                {
                    ResetRunMetrics();
                    Console.WriteLine($"[Q&A] ID: {commandId}");
                    Console.WriteLine($"Question: {question}");
                    cancelCts = StartCancelHotkeyListener();
        
                    if (!Options.BatchMode &&
                        Options.AutoHideConsoleDuringRun &&
                        IsOwnConsoleForeground())
                        consoleHidden = ConcealConsoleWindow();
        
                    var capture = ScreenshotToDataUrl(screensDir, commandId, 1, null);
                    var preparedImage = PrepareHelperImage(capture.DataUrl, Options.QaScreenshotMaxWidth);
                    var dataUrl = preparedImage.DataUrl;
                    var imageW = preparedImage.Width;
                    var imageH = preparedImage.Height;
                    var screenW = capture.ScreenWidth;
                    var screenH = capture.ScreenHeight;
                    SetCurrentScreenMap(screenW, screenH, imageW, imageH);
                    Console.WriteLine($"[shot] {ShotLabel(capture.SavedPath, commandId, 1)}");
        
                    var (screenCx, screenCy, _, _) = GetCursorPositionInPrimary();
                    var (cx, cy, cnx, cny) = CursorToImageCoordinates(screenCx, screenCy);
                    var promptContext = CaptureUiPromptContext(capture.FocusUiaSummary, screenW, screenH);
                    var focusUiaRectForPrompt = ScreenRectToImage(capture.FocusUiaRect);
        
                    var systemRules = BuildQaSystemRules();
                    var qaModel = EffectiveQaModel();
                    var reqBody = BuildQARequestBody(qaModel, systemRules, question, dataUrl, imageW, imageH, cx, cy, cnx, cny, focusUiaRectForPrompt, capture.FocusUiaDataUrl, promptContext);
                    SaveRequestLog(Path.Combine(requestsDir, $"{commandId}_qa_request.json"), reqBody);

                    var (qa, raw) = await CallOpenAIParsedAsync<QaLocateDto>(apiKey, reqBody, cancelCts.Token);
                    SaveRaw(Path.Combine(requestsDir, $"{commandId}_qa_response.json"), raw);
        
                    if (CancelRequested || cancelCts.Token.IsCancellationRequested)
                        return CommandOutcome.Cancelled;
                    qa = QaAnswerValidator.Validate(qa, imageW, imageH);
                    if (qa == null)
                    {
                        Console.WriteLine(CancelRequested ? "Aborted (hotkey)." : "No usable answer was returned.");
                        return CancelRequested ? CommandOutcome.Cancelled : CommandOutcome.Failed;
                    }
        
                    if (!string.IsNullOrWhiteSpace(qa.AnswerText))
                    {
                        Console.WriteLine("🟩 Answer:");
                        Console.WriteLine(qa.AnswerText.Trim());
                    }
                    if (!string.IsNullOrWhiteSpace(qa.Note))
                        Console.WriteLine($"ℹ️ note: {qa.Note}");
        
                    QaAnswerPresentation.WriteLocation(qa, CurrentRun.CurrentScreenMap);
                    return CommandOutcome.Success;
                }
                finally
                {
                    if (cancelCts is not null) await cancelCts.DisposeAsync();
                    PrintRunMetrics();
                    if (consoleHidden && Options.RestoreConsoleAfterRun)
                        RestoreConsoleWindow();
                }
            }
    }
