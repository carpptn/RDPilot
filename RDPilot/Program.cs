using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Automation;

internal class RDPilot
{
    // === CONFIG ===
    const string Model = "gpt-5";                       // e.g., gpt-4o-mini, gpt-4o, gpt-5
    const string ApiUrl = "https://api.openai.com/v1/responses";
    
    const int MaxStepsDefault = 10000;

    // Mouse: enable/disable (default DISABLED here)
    static bool MouseEnabled = false;

    // Global, configurable time to let UI "settle" after an action before taking the next screenshot
    static int UiSettleDelayMs = 1000;                  // can override via POST_ACTION_DELAY_MS or --delay <ms>

    // Include a small crop of the current UIA focus area as an extra image?
    const bool IncludeFocusUiaCrop = true;
    const bool SendFocusCrop = true;                    // attach crop from AIM/request_crop
    const int FocusCropSize = 320;                      // px

    // Focus overlay (look)
    const int FocusRingPadding = 6;
    const int FocusRingThickness = 4;
    const int FocusCornerRadius = 10;

    // === Grid overlay on screenshots ===
    static int GridStepPx = 0;          // 0 = off; e.g., 100 = lines every 100 px
    static int GridLabelEveryPx = 100;  // label every N px
    static int GridMajorEveryPx = 100;  // thicker line every N px

    // AIM expiration after a large visual change
    const double AimExpireDelta = 0.08;

    // Stagnation / verification
    const double NoChangeThreshold = 0.005;             // 0..1 (avg pixel diff after downsampling)


    static readonly JsonSerializerOptions PrettyJson = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    // --- HOTKEY: Ctrl+Alt+Q (works when console has focus) ---
    [DllImport("user32.dll")] static extern short GetAsyncKeyState(int vKey);
    static volatile bool CancelRequested = false;
    static bool IsPressed(int vk) => (GetAsyncKeyState(vk) & 0x8000) != 0;

    static CancellationTokenSource StartCancelHotkeyListener()
    {
        CancelRequested = false;
        var cts = new CancellationTokenSource();
        _ = Task.Run(async () =>
        {
            while (!cts.IsCancellationRequested)
            {
                if (IsPressed(0x11) && IsPressed(0x12) && IsPressed(0x51)) // Ctrl+Alt+Q
                {
                    CancelRequested = true;
                    Console.WriteLine("\n⛔ Aborted (Ctrl+Alt+Q)");
                    break;
                }
                await Task.Delay(50);
            }
        }, cts.Token);
        return cts;
    }

    // --- DPI awareness ---
    [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(nint value);

    // --- Console always-on-top (Win32) ---
    [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    const uint SWP_NOSIZE = 0x0001;
    const uint SWP_NOMOVE = 0x0002;
    const uint SWP_NOACTIVATE = 0x0010;
    const uint SWP_SHOWWINDOW = 0x0040;

    static CancellationTokenSource? _topMostCts;


    static async Task Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        // MOUSE_ENABLED from env/arg
        var envME = Environment.GetEnvironmentVariable("MOUSE_ENABLED");
        if (!string.IsNullOrWhiteSpace(envME))
        {
            MouseEnabled = envME.Equals("1", StringComparison.OrdinalIgnoreCase)
                        || envME.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || envME.Equals("yes", StringComparison.OrdinalIgnoreCase);
        }
        if (args.Any(a => a.Equals("--mouse", StringComparison.OrdinalIgnoreCase)))
            MouseEnabled = true;

        // POST_ACTION_DELAY_MS from env / --delay <ms>
        var envDelay = Environment.GetEnvironmentVariable("POST_ACTION_DELAY_MS");
        if (int.TryParse(envDelay, out var d1) && d1 >= 0) UiSettleDelayMs = d1;
        var delayArgIdx = Array.FindIndex(args, a => a.Equals("--delay", StringComparison.OrdinalIgnoreCase));
        if (delayArgIdx >= 0 && delayArgIdx + 1 < args.Length && int.TryParse(args[delayArgIdx + 1], out var d2) && d2 >= 0)
            UiSettleDelayMs = d2;

        // GRID_STEP_PX from env / --grid <px|0|off>
        var envGrid = Environment.GetEnvironmentVariable("GRID_STEP_PX");
        if (int.TryParse(envGrid, out var g1) && g1 >= 0) GridStepPx = g1;

        var gridArgIdx = Array.FindIndex(args, a => a.Equals("--grid", StringComparison.OrdinalIgnoreCase));
        if (gridArgIdx >= 0)
        {
            if (gridArgIdx + 1 < args.Length)
            {
                var gridVal = args[gridArgIdx + 1];
                if (int.TryParse(gridVal, out var g2) && g2 >= 0)
                    GridStepPx = g2;
                else if (string.Equals(gridVal, "off", StringComparison.OrdinalIgnoreCase) || gridVal == "0")
                    GridStepPx = 0;
            }
        }

        var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            Console.Write("Enter OPENAI_API_KEY: ");
            apiKey = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                Console.Error.WriteLine("Missing API key.");
                return;
            }
        }

        try { SetProcessDpiAwarenessContext((nint)(-4)); } catch { /* best effort */ }

        string? pending = args
            .Where(a => !a.Equals("--mouse", StringComparison.OrdinalIgnoreCase))
            .Where((a, i) => i != delayArgIdx && i != delayArgIdx + 1)
            .DefaultIfEmpty(null).LastOrDefault();

        Console.WriteLine("Interactive mode. Enter commands or questions.");
        Console.WriteLine("Empty input or '/exit' will quit. Prefix '/ask ' forces Q&A mode.");
        Console.WriteLine("Emergency abort while running: Ctrl+Alt+Q\n");

        while (true)
        {
            string goal;
            if (!string.IsNullOrEmpty(pending))
            {
                goal = pending!;
                pending = null;
                Console.WriteLine($"Command (from args): {goal}");
            }
            else
            {
                Console.Write("Command/Question: ");
                goal = Console.ReadLine() ?? "";
            }

            if (string.IsNullOrWhiteSpace(goal) || goal.Trim().Equals("/exit", StringComparison.OrdinalIgnoreCase))
                break;

            if (IsQuestion(goal))
                await RunAskOnce(apiKey!, goal);
            else
                await RunOnce(apiKey!, goal);

            Console.WriteLine();
            Console.WriteLine("✅ Done. Enter next (ENTER = exit).");
        }
    }

    // === Q&A detection ===
    static bool IsQuestion(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s.Trim();
        if (s.StartsWith("/ask ", StringComparison.OrdinalIgnoreCase)) return true;
        if (s.EndsWith("?")) return true;
        return false;
    }

    // === Q&A ===
    static async Task RunAskOnce(string apiKey, string question)
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

        try
        {
            Console.WriteLine($"[Q&A] ID: {commandId}");
            Console.WriteLine($"Question: {question}");

            var (dataUrl, savedPath, _, _, focusUiaRect, focusUiaDataUrl, focusUiaPath) =
                ScreenshotToDataUrl(screensDir, commandId, 1, null);
            Console.WriteLine($"[shot] {Path.GetFileName(savedPath)}");

            var (_, _, screenW, screenH) = GetPrimaryScreen();
            var (cx, cy, cnx, cny) = GetCursorPositionInPrimary();

            var systemRules = BuildQaSystemRules();
            var reqBody = BuildQARequestBody(Model, systemRules, question, dataUrl, screenW, screenH, cx, cy, cnx, cny, focusUiaRect, focusUiaDataUrl);
            var reqBodyForLog = BuildQARequestBody_ForLog(Model, systemRules, question, savedPath, screenW, screenH, cx, cy, cnx, cny, focusUiaRect, focusUiaPath);

            SaveJson(Path.Combine(requestsDir, $"{commandId}_qa_request.json"), reqBodyForLog);

            var (qa, raw) = await CallOpenAIParsedAsync<QaLocateDto>(apiKey, reqBody);
            SaveRaw(Path.Combine(requestsDir, $"{commandId}_qa_response.json"), raw);

            if (qa == null)
            {
                Console.WriteLine("No response.");
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
            Console.SetOut(prevOut);
            Console.SetError(prevErr);
        }
    }

    // === Control loop ===
    static async Task RunOnce(string apiKey, string goal)
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

        try
        {
            Console.WriteLine($"Command ID: {commandId}");
            Console.WriteLine($"Goal: {goal}");
            Console.WriteLine("Loop start: one action -> screenshot -> next decision.");
            Console.WriteLine("Emergency abort: Ctrl+Alt+Q\n");

            var history = new List<string>();
            using var cancelCts = StartCancelHotkeyListener();

            Rectangle? nextFocusRect = null; // crop/overlay after 'aim'/'point'/'request_crop'
            Rectangle? lastAimRect = null;   // active AIM – required before clicks

            // change / strategy metrics
            string? prevShotPath = null;
            ActionDto? prevAction = null;
            string? lastSig = null;
            int stagnationSteps = 0;
            int repeatCount = 0;
            double lastDelta = double.NaN;

            for (int step = 1; step <= MaxStepsDefault; step++)
            {
                if (CancelRequested) { Console.WriteLine("Aborted (hotkey)."); break; }

                // screenshot at the beginning of a step (state after previous action)
                var (dataUrl, savedPath, focusUrl, appliedFocusRect, focusUiaRect, focusUiaDataUrl, focusUiaPath) =
                    ScreenshotToDataUrl(screensDir, commandId, step, nextFocusRect);
                Console.WriteLine($"[shot] {Path.GetFileName(savedPath)}");

                if (appliedFocusRect is Rectangle rCrop)
                {
                    var cropPath = Path.Combine(screensDir, $"{commandId}_{step}_crop.png");
                    Console.WriteLine($"[crop] {Path.GetFileName(cropPath)}  bbox=({rCrop.Left},{rCrop.Top})–({rCrop.Right},{rCrop.Bottom})");
                    SaveAimOverlay(savedPath, rCrop, Path.Combine(screensDir, $"{commandId}_{step}_aim_overlay.png"));
                    Console.WriteLine($"[aim-overlay] {Path.GetFileName(Path.Combine(screensDir, $"{commandId}_{step}_aim_overlay.png"))}");
                }
                if (focusUiaRect is Rectangle fr)
                    Console.WriteLine($"[focus_uia] bbox=({fr.Left},{fr.Top})–({fr.Right},{fr.Bottom})");

                // — Visual delta vs previous screenshot (effect of last action)
                if (prevShotPath != null)
                {
                    lastDelta = ComputeImageDelta(prevShotPath, savedPath); // 0..1
                    bool noChange = lastDelta < NoChangeThreshold;

                    if (noChange) stagnationSteps++; else stagnationSteps = 0;

                    if (prevAction != null)
                    {
                        var sig = ActionSignature(prevAction);
                        if (noChange && sig == lastSig) repeatCount++;
                        else { repeatCount = 0; lastSig = sig; }
                    }

                    // Expire AIM after a large visual change
                    if (lastAimRect is not null && lastDelta > AimExpireDelta)
                    {
                        Console.WriteLine($"[aim] expired (delta={lastDelta:0.###} > {AimExpireDelta:0.###})");
                        lastAimRect = null;
                    }
                }

                var historyTail = Tail(string.Join(Environment.NewLine, history), 2000);
                var (_, _, screenW, screenH) = GetPrimaryScreen();
                var (cx, cy, cnx, cny) = GetCursorPositionInPrimary();

                var systemRules = BuildSystemRules(); // includes AIM rules and overlay description

                // inject observation metrics into the prompt
                var metaSb = new StringBuilder()
                    .AppendLine($"LAST_STEP_DELTA: {(double.IsNaN(lastDelta) ? "N/A" : lastDelta.ToString("0.####"))} (threshold={NoChangeThreshold})")
                    .AppendLine($"STAGNATION_STEPS: {stagnationSteps}")
                    .AppendLine($"REPEAT_COUNT: {repeatCount}")
                    .AppendLine($"LAST_ACTION: {(prevAction == null ? "N/A" : Describe(prevAction))}")
                    .AppendLine($"AIM_ACTIVE: {(lastAimRect is null ? "false" : $"true [{lastAimRect.Value.Left},{lastAimRect.Value.Top}]–[{lastAimRect.Value.Right},{lastAimRect.Value.Bottom}]")}");

                var reqBody = BuildRequestBody(Model, systemRules, goal, historyTail + "\n" + metaSb, dataUrl, screenW, screenH,
                                               cx, cy, cnx, cny, focusUrl, appliedFocusRect, focusUiaRect, focusUiaDataUrl);
                var reqBodyForLog = BuildRequestBody_ForLog(Model, systemRules, goal, historyTail + "\n" + metaSb,
                                                            savedPath, screenW, screenH, cx, cy, cnx, cny,
                                                            appliedFocusRect != null ? Path.Combine(screensDir, $"{commandId}_{step}_crop.png") : null,
                                                            appliedFocusRect,
                                                            focusUiaRect, focusUiaPath);

                SaveJson(Path.Combine(requestsDir, $"{commandId}_{step}_request.json"), reqBodyForLog);

                var (action, raw) = await CallOpenAIAsync(apiKey, reqBody);
                SaveRaw(Path.Combine(requestsDir, $"{commandId}_{step}_response.json"), raw);

                if (action is null)
                {
                    Console.WriteLine("Could not parse action. Aborting this goal.");
                    break;
                }

                Console.WriteLine($"[{step}] {Describe(action)}");
                if (!string.IsNullOrWhiteSpace(action.Note))
                    Console.WriteLine($"     note: {action.Note}");

                nextFocusRect = null; // reset – set by aim/point/request_crop
                bool executed = false;

                try
                {
                    // ——— Mouse policy: global switch ———
                    if (IsMouseAction(action) && !MouseEnabled)
                    {
                        Console.WriteLine("[guard] mouse disabled → ignoring mouse action; use keyboard strategy or 'aim' without clicking.");
                        history.Add($"[{step}] IGNORED (mouse_disabled)");
                    }
                    else if (action.Type == "aim")
                    {
                        var rect = ResolveAimRect(action);
                        if (rect is null) throw new InvalidOperationException("aim without parameters (bbox/crop/x/y/x_px/y_px).");
                        lastAimRect = rect.Value;
                        nextFocusRect = rect.Value; // show crop/overlay on next screenshot
                    }
                    else if (action.Type == "point")
                    {
                        // Visual pointer only – does NOT set AIM (clicks still blocked without 'aim')
                        var rect = ResolveCropRect(action);
                        if (rect is null) throw new InvalidOperationException("point without parameters.");
                        nextFocusRect = rect.Value;
                    }
                    else if (action.Type == "request_crop")
                    {
                        var rect = ResolveCropRect(action);
                        if (rect is null) throw new InvalidOperationException("request_crop without parameters.");
                        nextFocusRect = rect.Value;
                    }
                    else if (action.Type == "wait")
                    {
                        int secs = Math.Max(0, action.WaitSeconds ?? 1);
                        Console.WriteLine($"[wait] Sleeping {secs} s (long-running operation on screen)...");
                        await Task.Delay(secs * 1000);
                        executed = true;
                    }
                    else if (action.Type == "done")
                    {
                        await Task.Delay(UiSettleDelayMs); // give UI time
                        var (_, freshPath, _, _, _, _, _) = ScreenshotToDataUrl(screensDir, commandId, step, null);
                        var verify = await VerifyGoalAsync(apiKey, goal, freshPath, requestsDir, commandId, step);
                        if (verify?.Verdict?.Equals("yes", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            Console.WriteLine($"[verify] ✅ Goal confirmed: {verify.Reason}");
                            history.Add($"[{step}] done_verified");
                            lastAimRect = null;
                            Console.WriteLine("Finished (model returned 'done').");
                            break;
                        }
                        else
                        {
                            Console.WriteLine($"[verify] ❌ Goal NOT confirmed. Reason: {verify?.Reason ?? "n/a"}");
                            history.Add($"[{step}] done_rejected: {verify?.Reason}");
                        }
                    }
                    else if (action.Type is "click" or "double_click")
                    {
                        if (lastAimRect is null)
                        {
                            Console.WriteLine("[guard] click blocked: no active AIM. Return 'aim' first.");
                            history.Add($"[{step}] IGNORED (click_without_aim)");
                        }
                        else
                        {
                            // Clicks must include explicit coordinates (don't default to AIM center)
                            if (!HasExplicitPoint(action))
                            {
                                Console.WriteLine("[guard] click/double_click requires explicit coordinates (x/y or x_px/y_px) – provide an exact point within AIM.");
                                history.Add($"[{step}] IGNORED (click_missing_coords)");
                                continue; // go to next round
                            }

                            var (xClick, yClick) = ResolvePoint(action);

                            if (!lastAimRect.Value.Contains(xClick, yClick))
                            {
                                Console.WriteLine("[guard] click outside active AIM → ignoring. Set a proper 'aim' first.");
                                history.Add($"[{step}] IGNORED (click_outside_aim)");
                            }
                            else
                            {
                                ExecuteAction(action);
                                executed = true;
                            }
                        }
                    }
                    else
                    {
                        // Move/scroll/keys/type_text – execute normally
                        ExecuteAction(action);
                        executed = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Action execution error: {ex.Message}");
                    break;
                }

                history.Add($"[{step}] {Describe(action)}");
                if (!string.IsNullOrWhiteSpace(action.Note))
                    history.Add($"[{step}] note: {action.Note}");

                // Keep context for next step (delta/repeat metrics)
                prevAction = action;
                prevShotPath = savedPath;

                if (CancelRequested) { Console.WriteLine("Aborted (hotkey)."); break; }

                // After 'wait' we don't add the artificial UI settle delay
                int postDelay = action.Type == "wait" ? 0 : Math.Max(UiSettleDelayMs, DelayFor(action));
                if (postDelay > 0) await Task.Delay(postDelay);
            }
        }
        finally
        {
            Console.SetOut(prevOut);
            Console.SetError(prevErr);
        }
    }

    static bool IsMouseAction(ActionDto a)
        => a.Type is "move" or "click" or "double_click" or "scroll";

    static int DelayFor(ActionDto a) => a.Type switch
    {
        "move" => 120,
        "click" => 180,
        "double_click" => 250,
        "keys" => 120,
        "type_text" => 80,
        "scroll" => 80,
        "request_crop" => 80,
        "point" => 80,
        "aim" => 80,
        "wait" => 0,   // wait time handled separately
        "done" => 80,
        _ => 120
    };

    // === System rules (control) ===
    static string BuildSystemRules()
    {
        var sb = new StringBuilder();
        sb.AppendLine("You are an agent that controls a Windows 10/11 computer strictly via the UI.");
        sb.AppendLine("Return EXACTLY ONE action per round as JSON:");
        sb.AppendLine("- keys, type_text, move, click, double_click, scroll, request_crop, point, aim, wait, done");
        sb.AppendLine();
        sb.AppendLine("Important: The screenshot may contain a white+red rounded rectangle overlay – that's the element with current keyboard focus (FOCUS_UIA). Treat it as a reliable source of truth.");
        sb.AppendLine();
        sb.AppendLine("Every action MUST include a short 'note' (1–2 sentences explaining your decision).");
        sb.AppendLine();
        sb.AppendLine("Guidelines:");
        sb.AppendLine("- Work in pixel coordinates (0,0 is top-left). When clicking, aim at the center of the bbox.");
        sb.AppendLine($"- MOUSE_ALLOWED: {(MouseEnabled ? "true" : "false")} (when false – do NOT use the mouse).");
        sb.AppendLine("- When MOUSE_ALLOWED is false, move/click/double_click/scroll are not allowed – use keyboard instead.");
        sb.AppendLine("- Strongly prefer the keyboard. Use TAB/Shift+TAB, Ctrl+L/F6 (address), Ctrl+K/E (search), Ctrl+W, etc.");
        sb.AppendLine("- For text input use 'type_text' (full UNICODE string). Use 'keys' only for shortcuts and function keys.");
        sb.AppendLine();
        sb.AppendLine("- One action per round. Decide solely from the screenshot and metadata (SCREEN_SIZE, CURSOR_POS, FOCUS_UIA/FOCUS_CROP, DELTA/REPEAT).");
        sb.AppendLine("- If the target is ambiguous, prefer actions relative to FOCUS_UIA (e.g., TAB/Shift+TAB or aim at the center of FOCUS_UIA).");
        sb.AppendLine();
        sb.AppendLine("- Return 'done' ONLY when the screen state clearly confirms the goal (there will be an additional verification).");
        sb.AppendLine("- DO NOT use machine-specific taskbar/app-number shortcuts: Win+1..9, Super+1..9, etc.");
        sb.AppendLine("- Prefer deterministic strategies. If a proposed action may be nondeterministic – choose an alternative.");
        sb.AppendLine();
        sb.AppendLine("- BEFORE any 'click'/'double_click' you MUST set an 'aim' (the target region). Clicks outside the active AIM are ignored.");
        sb.AppendLine("- After setting AIM, ensure the intended target is visible within the AIM frame. If not – re-aim until it is.");
        sb.AppendLine($"- After a large visual change (LAST_STEP_DELTA > {AimExpireDelta:0.###}) the previous AIM expires – set a new one before clicking.");
        sb.AppendLine("- Define 'aim' via 'bbox' (preferred) or a point (x/y or x_px/y_px) – in the latter case the crop is a square of ~ FocusCropSize.");
        sb.AppendLine("- 'request_crop' and 'point' are only for requesting zoom/homing; they do NOT replace 'aim'.");
        sb.AppendLine("- If an active AIM exists, in 'click'/'double_click' you MUST PROVIDE COORDINATES inside AIM (do not rely on implicit centering).");
        sb.AppendLine("- 'double_click' means a standard double click (e.g., launch an app). 'button:right' in 'click' = e.g., context menu.");
        sb.AppendLine();
        sb.AppendLine("- Use 'wait' when a long-running process is visible (e.g., progress bar, installer, render, upload).");
        sb.AppendLine("  Set 'wait_seconds' to a realistic duration; during 'wait' no screenshots are taken. Reassess the screen afterwards.");
        if (GridStepPx > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"- The screenshot contains a semi-transparent grid every {GridStepPx} pixels; use it to provide precise x_px/y_px.");
            sb.AppendLine("- The origin (0,0) is the top-left corner of the screen.");
        }
        return sb.ToString();
    }

    // === Q&A rules ===
    static string BuildQaSystemRules()
    {
        var sb = new StringBuilder();
        sb.AppendLine("You are a screen analyst. Answer strictly based on the screenshot, metadata (SCREEN_SIZE, CURSOR_POS), and the user's question.");
        sb.AppendLine("The image may include a white+red rounded rectangle overlay – that's the element with current keyboard focus (FOCUS_UIA). Treat it as a reliable focus indicator.");
        sb.AppendLine("Return BOTH: a short textual answer and location metadata for the most relevant element. Add a short 'note'.");
        sb.AppendLine("Always think in pixel coordinates (0,0 top-left). If a location makes sense, choose the center of the visible bbox.");
        sb.AppendLine("If a location would not make sense, the location fields may be null.");
        if (GridStepPx > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"- The screenshot contains a semi-transparent grid every {GridStepPx} pixels; use it to provide precise x_px/y_px.");
            sb.AppendLine("- The origin (0,0) is the top-left corner of the screen.");
        }
        return sb.ToString();
    }

    // === Rectangle schema ===
    static object BoxSchema() => new
    {
        type = new object[] { "object", "null" },
        additionalProperties = false,
        properties = new Dictionary<string, object>
        {
            ["left"] = new { type = new object[] { "integer", "null" } },
            ["top"] = new { type = new object[] { "integer", "null" } },
            ["right"] = new { type = new object[] { "integer", "null" } },
            ["bottom"] = new { type = new object[] { "integer", "null" } },
        },
        required = new[] { "left", "top", "right", "bottom" }
    };

    // === Request build (control) ===
    static object BuildRequestBody(
        string model, string systemRules, string goal, string historyPlusMeta,
        string imageDataUrl, int screenW, int screenH,
        int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
        string? focusDataUrl, Rectangle? focusRect,
        Rectangle? focusUiaRect, string? focusUiaDataUrl)
    {
        var format = new
        {
            type = "json_schema",
            name = "SingleAction",
            strict = true,
            schema = new
            {
                type = "object",
                additionalProperties = false,
                properties = new Dictionary<string, object>
                {
                    ["type"] = new { type = "string", @enum = new[] { "move", "click", "double_click", "keys", "type_text", "scroll", "request_crop", "point", "aim", "wait", "done" } },
                    ["x"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                    ["y"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                    ["x_px"] = new { type = new object[] { "integer", "null" } },
                    ["y_px"] = new { type = new object[] { "integer", "null" } },
                    ["button"] = new { type = new object[] { "string", "null" }, @enum = new object[] { "left", "right", "middle", null! } },
                    ["keys"] = new { type = new object[] { "array", "null" }, items = new { type = "string" } },
                    ["text"] = new { type = new object[] { "string", "null" } },
                    ["scroll_dy"] = new { type = new object[] { "integer", "null" } }, // positive = down
                    ["bbox"] = BoxSchema(),
                    ["crop"] = BoxSchema(),
                    ["wait_seconds"] = new { type = new object[] { "integer", "null" }, minimum = 0 },
                    ["note"] = new { type = new object[] { "string", "null" } }
                },
                required = new[] { "type", "x", "y", "x_px", "y_px", "button", "keys", "text", "scroll_dy", "bbox", "crop", "wait_seconds", "note" }
            }
        };

        var userText = new StringBuilder()
            .AppendLine($"GOAL: {goal}")
            .AppendLine("HISTORY:")
            .AppendLine(historyPlusMeta)
            .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px)")
            .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})")
            .AppendLine($"MOUSE_ALLOWED: {(MouseEnabled ? "true" : "false")}");

        if (focusUiaRect.HasValue)
        {
            var r = focusUiaRect.Value;
            int cx = (r.Left + r.Right) / 2;
            int cy = (r.Top + r.Bottom) / 2;
            userText.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
            userText.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
        }
        else
        {
            userText.AppendLine("FOCUS_UIA: none");
        }

        if (focusRect.HasValue)
        {
            var r = focusRect.Value;
            userText.AppendLine($"FOCUS_CROP: left={r.Left}, top={r.Top}, width={r.Width}, height={r.Height} (px).");
        }

        var userContent = new List<object>
        {
            new { type = "input_text",  text = userText.ToString() },
            new { type = "input_image", image_url = imageDataUrl }
        };
        if (IncludeFocusUiaCrop && focusUiaDataUrl != null)
            userContent.Add(new { type = "input_image", image_url = focusUiaDataUrl });
        if (focusDataUrl != null)
            userContent.Add(new { type = "input_image", image_url = focusDataUrl });

        var req = new Dictionary<string, object>
        {
            ["model"] = model,
            ["input"] = new object[]
            {
                new { role = "system", content = new object[] { new { type = "input_text", text = systemRules } } },
                new { role = "user",   content = userContent.ToArray() }
            },
            ["text"] = new { format }
        };
        if (SupportsTemperature(model))
            req["temperature"] = 0.0;

        return req;
    }

    // Logging variant (no base64)
    static object BuildRequestBody_ForLog(
        string model, string systemRules, string goal, string historyPlusMeta,
        string fullImagePath, int screenW, int screenH,
        int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
        string? focusImagePath, Rectangle? focusRect,
        Rectangle? focusUiaRect, string? focusUiaImagePath)
    {
        var format = new { name = "SingleAction" };
        var userText = new StringBuilder()
            .AppendLine($"GOAL: {goal}")
            .AppendLine("HISTORY:")
            .AppendLine(historyPlusMeta)
            .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px)")
            .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})")
            .AppendLine($"MOUSE_ALLOWED: {(MouseEnabled ? "true" : "false")}");

        if (focusUiaRect.HasValue)
        {
            var r = focusUiaRect.Value;
            int cx = (r.Left + r.Right) / 2;
            int cy = (r.Top + r.Bottom) / 2;
            userText.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
            userText.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
        }
        else
        {
            userText.AppendLine("FOCUS_UIA: none");
        }

        if (focusRect.HasValue)
        {
            var r = focusRect.Value;
            userText.AppendLine($"FOCUS_CROP: left={r.Left}, top={r.Top}, width={r.Width}, height={r.Height} (px).");
        }

        var userContent = new List<object>
        {
            new { type = "input_text",  text = userText.ToString() },
            new { type = "input_image", image_url = $"file://{fullImagePath}" }
        };
        if (IncludeFocusUiaCrop && focusUiaImagePath != null)
            userContent.Add(new { type = "input_image", image_url = $"file://{focusUiaImagePath}" });
        if (focusImagePath != null)
            userContent.Add(new { type = "input_image", image_url = $"file://{focusImagePath}" });

        var req = new Dictionary<string, object>
        {
            ["model"] = model,
            ["input"] = new object[]
            {
                new { role = "system", content = new object[] { new { type = "input_text", text = systemRules } } },
                new { role = "user",   content = userContent.ToArray() }
            },
            ["text"] = new { format }
        };
        if (SupportsTemperature(model))
            req["temperature"] = 0.0;

        return req;
    }

    // === Request build (Q&A) ===
    static object BuildQARequestBody(string model, string qaRules, string question, string imageDataUrl,
                                     int screenW, int screenH, int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
                                     Rectangle? focusUiaRect, string? focusUiaDataUrl)
    {
        if (question.StartsWith("/ask ", StringComparison.OrdinalIgnoreCase))
            question = question[5..];

        var format = new
        {
            type = "json_schema",
            name = "QAWithLocate",
            strict = true,
            schema = new
            {
                type = "object",
                additionalProperties = false,
                properties = new Dictionary<string, object>
                {
                    ["answer_text"] = new { type = new object[] { "string", "null" } },
                    ["x"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                    ["y"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                    ["x_px"] = new { type = new object[] { "integer", "null" } },
                    ["y_px"] = new { type = new object[] { "integer", "null" } },
                    ["bbox"] = BoxSchema(),
                    ["note"] = new { type = new object[] { "string", "null" } },
                },
                required = new[] { "answer_text", "x", "y", "x_px", "y_px", "bbox", "note" }
            }
        };

        var meta = new StringBuilder()
            .AppendLine($"QUESTION: {question}")
            .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px)")
            .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})");

        if (focusUiaRect.HasValue)
        {
            var r = focusUiaRect.Value;
            int cx = (r.Left + r.Right) / 2;
            int cy = (r.Top + r.Bottom) / 2;
            meta.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
            meta.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
        }
        else
        {
            meta.AppendLine("FOCUS_UIA: none");
        }

        var userContent = new List<object>
        {
            new { type = "input_text",  text = meta.ToString() },
            new { type = "input_image", image_url = imageDataUrl }
        };
        if (IncludeFocusUiaCrop && focusUiaDataUrl != null)
            userContent.Add(new { type = "input_image", image_url = focusUiaDataUrl });

        var req = new Dictionary<string, object>
        {
            ["model"] = model,
            ["input"] = new object[]
            {
                new { role = "system", content = new object[] { new { type = "input_text", text = qaRules } } },
                new { role = "user",   content = userContent.ToArray() }
            },
            ["text"] = new { format }
        };
        if (SupportsTemperature(model))
            req["temperature"] = 0.0;

        return req;
    }

    static object BuildQARequestBody_ForLog(string model, string qaRules, string question, string fullImagePath,
                                            int screenW, int screenH, int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
                                            Rectangle? focusUiaRect, string? focusUiaImagePath)
    {
        if (question.StartsWith("/ask ", StringComparison.OrdinalIgnoreCase))
            question = question[5..];

        var meta = new StringBuilder()
            .AppendLine($"QUESTION: {question}")
            .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px)")
            .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})");

        if (focusUiaRect.HasValue)
        {
            var r = focusUiaRect.Value;
            int cx = (r.Left + r.Right) / 2;
            int cy = (r.Top + r.Bottom) / 2;
            meta.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
            meta.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
        }
        else
        {
            meta.AppendLine("FOCUS_UIA: none");
        }

        var content = new List<object>
        {
            new { type = "input_text",  text = meta.ToString() },
            new { type = "input_image", image_url = $"file://{fullImagePath}" }
        };
        if (IncludeFocusUiaCrop && focusUiaImagePath != null)
            content.Add(new { type = "input_image", image_url = $"file://{focusUiaImagePath}" });

        var req = new Dictionary<string, object>
        {
            ["model"] = model,
            ["input"] = new object[]
            {
                new { role = "system", content = new object[] { new { type = "input_text", text = qaRules } } },
                new { role = "user",   content = content.ToArray() }
            },
            ["text"] = new { format = new { name = "QAWithLocate" } }
        };
        if (SupportsTemperature(model))
            req["temperature"] = 0.0;

        return req;
    }

    // === VerifyGoal (Q&A yes/no) + request logs ===
    static async Task<VerifyDto?> VerifyGoalAsync(string apiKey, string goal, string currentShotPath,
                                                  string requestsDir, string commandId, int step)
    {
        var (vx, vy, vw, vh) = GetPrimaryScreen();
        var dataUrl = EncodeBitmapToDataUrl(new Bitmap(currentShotPath));

        var format = new
        {
            type = "json_schema",
            name = "VerifyGoal",
            strict = true,
            schema = new
            {
                type = "object",
                additionalProperties = false,
                properties = new Dictionary<string, object>
                {
                    ["verdict"] = new { type = "string", @enum = new[] { "yes", "no" } },
                    ["reason"] = new { type = new object[] { "string", "null" } }
                },
                required = new[] { "verdict", "reason" }
            }
        };

        var rules = "You are a strict verifier. Based on the image, decide whether the GOAL is achieved. Return 'yes' only if the screen makes it unambiguous.";

        var userText = new StringBuilder()
            .AppendLine($"GOAL: {goal}")
            .AppendLine($"SCREEN_SIZE: width={vw}, height={vh} (px)");

        var req = new Dictionary<string, object>
        {
            ["model"] = Model,
            ["input"] = new object[]
            {
                new { role = "system", content = new object[] { new { type = "input_text", text = rules } } },
                new { role = "user",   content = new object[] {
                        new { type = "input_text", text = userText.ToString() },
                        new { type = "input_image", image_url = dataUrl }
                    }
                }
            },
            ["text"] = new { format }
        };
        if (SupportsTemperature(Model))
            req["temperature"] = 0.0;

        // log request
        SaveJson(Path.Combine(requestsDir, $"{commandId}_{step}_verify_request.json"),
                 new { request = "verify", body = new { rules, meta = userText.ToString(), screenshot = $"file://{currentShotPath}" } });

        var (parsed, raw) = await CallOpenAIParsedAsync<VerifyDto>(apiKey, req);

        // log response
        SaveRaw(Path.Combine(requestsDir, $"{commandId}_{step}_verify_response.json"), raw);

        return parsed;
    }

    static bool SupportsTemperature(string modelName) =>
        !(modelName?.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase) ?? false);

    // ==== API calls ====
    static async Task<(ActionDto? parsed, string raw)> CallOpenAIAsync(string apiKey, object body)
    {
        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        http.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");

        var json = JsonSerializer.Serialize(body);
        var resp = await http.PostAsync(ApiUrl, new StringContent(json, Encoding.UTF8, "application/json"));
        var raw = await resp.Content.ReadAsStringAsync();

        if (!resp.IsSuccessStatusCode)
        {
            Console.WriteLine($"OpenAI HTTP {(int)resp.StatusCode}: {raw}");
            return (null, raw);
        }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;

            if (root.TryGetProperty("output_parsed", out var op))
                return (JsonSerializer.Deserialize<ActionDto>(op.GetRawText()), raw);

            if (root.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in output.EnumerateArray())
                {
                    if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array) continue;
                    foreach (var c in content.EnumerateArray())
                        if (c.TryGetProperty("parsed", out var parsedEl))
                            return (JsonSerializer.Deserialize<ActionDto>(parsedEl.GetRawText()), raw);
                }

                foreach (var item in output.EnumerateArray())
                {
                    if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array) continue;
                    foreach (var c in content.EnumerateArray())
                    {
                        if (c.TryGetProperty("type", out var tEl) &&
                            tEl.ValueKind == JsonValueKind.String &&
                            tEl.GetString() == "output_text" &&
                            c.TryGetProperty("text", out var txtEl) &&
                            txtEl.ValueKind == JsonValueKind.String)
                        {
                            var maybe = TryExtractJsonObject(txtEl.GetString() ?? "");
                            if (!string.IsNullOrWhiteSpace(maybe))
                                return (JsonSerializer.Deserialize<ActionDto>(maybe), raw);
                        }
                    }
                }
            }

            if (root.TryGetProperty("output_text", out var outText) && outText.ValueKind == JsonValueKind.String)
            {
                var maybe = TryExtractJsonObject(outText.GetString() ?? "");
                if (!string.IsNullOrWhiteSpace(maybe))
                    return (JsonSerializer.Deserialize<ActionDto>(maybe), raw);
            }
            if (root.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
            {
                var maybe = TryExtractJsonObject(textEl.GetString() ?? "");
                if (!string.IsNullOrWhiteSpace(maybe))
                    return (JsonSerializer.Deserialize<ActionDto>(maybe), raw);
            }

            Console.WriteLine("No parsable JSON found (parsed/output_text). RAW:");
            Console.WriteLine(raw);
            return (null, raw);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Response parsing error: {ex.Message}");
            Console.WriteLine(raw);
            return (null, raw);
        }
    }

    static async Task<(T? parsed, string raw)> CallOpenAIParsedAsync<T>(string apiKey, object body)
    {
        using var http = new HttpClient();
        http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        http.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json");

        var json = JsonSerializer.Serialize(body);
        var resp = await http.PostAsync(ApiUrl, new StringContent(json, Encoding.UTF8, "application/json"));
        var raw = await resp.Content.ReadAsStringAsync();

        if (!resp.IsSuccessStatusCode)
        {
            Console.WriteLine($"OpenAI HTTP {(int)resp.StatusCode}: {raw}");
            return (default, raw);
        }

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;

            // 1) directly parsed
            if (root.TryGetProperty("output_parsed", out var op))
                return (JsonSerializer.Deserialize<T>(op.GetRawText()), raw);

            // 2) search output[*].content[*]
            if (root.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in output.EnumerateArray())
                {
                    if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array) continue;

                    // 2a) elements with "parsed"
                    foreach (var c in content.EnumerateArray())
                        if (c.TryGetProperty("parsed", out var parsedEl))
                            return (JsonSerializer.Deserialize<T>(parsedEl.GetRawText()), raw);

                    // 2b) "output_text" items containing JSON text
                    foreach (var c in content.EnumerateArray())
                    {
                        if (c.TryGetProperty("type", out var tEl) &&
                            tEl.ValueKind == JsonValueKind.String &&
                            tEl.GetString() == "output_text" &&
                            c.TryGetProperty("text", out var txtEl) &&
                            txtEl.ValueKind == JsonValueKind.String)
                        {
                            var maybe = TryExtractJsonObject(txtEl.GetString() ?? "");
                            if (!string.IsNullOrWhiteSpace(maybe))
                                return (JsonSerializer.Deserialize<T>(maybe), raw);
                        }
                    }
                }
            }

            // 3) rarely: root-level output_text / text
            if (root.TryGetProperty("output_text", out var outText) && outText.ValueKind == JsonValueKind.String)
            {
                var maybe = TryExtractJsonObject(outText.GetString() ?? "");
                if (!string.IsNullOrWhiteSpace(maybe)) return (JsonSerializer.Deserialize<T>(maybe), raw);
            }
            if (root.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
            {
                var maybe = TryExtractJsonObject(textEl.GetString() ?? "");
                if (!string.IsNullOrWhiteSpace(maybe)) return (JsonSerializer.Deserialize<T>(maybe), raw);
            }

            Console.WriteLine("No parsable JSON found (parsed/output_text). RAW:");
            Console.WriteLine(raw);
            return (default, raw);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Response parsing error: {ex.Message}");
            Console.WriteLine(raw);
            return (default, raw);
        }
    }

    // helper – extract first balanced JSON object from a string
    static string? TryExtractJsonObject(string s)
    {
        if (string.IsNullOrEmpty(s)) return null;
        int start = s.IndexOf('{');
        if (start < 0) return null;
        int depth = 0;
        for (int i = start; i < s.Length; i++)
        {
            char ch = s[i];
            if (ch == '{') depth++;
            else if (ch == '}')
            {
                depth--;
                if (depth == 0)
                    return s[start..(i + 1)];
            }
        }
        return null;
    }

    // ====== Screenshot (PRIMARY) -> PNG + data URLs ======
    // Always draw the current UIA focus overlay on the full screenshot.
    static (string dataUrl, string savedPath,
            string? focusDataUrl, Rectangle? focusRect,
            Rectangle? focusUiaRect, string? focusUiaDataUrl, string? focusUiaImagePath)
        ScreenshotToDataUrl(string saveDir, string commandId, int step, Rectangle? explicitFocusRect)
    {
        var (vx, vy, vw, vh) = GetPrimaryScreen();

        using var bmp = new Bitmap(vw, vh, PixelFormat.Format24bppRgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(vx, vy, 0, 0, new Size(vw, vh), CopyPixelOperation.SourceCopy);

            // - UIA overlay (white+red rounded ring)
            var focusRectAbs = GetFocusedUiaBoundingRect();
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
                using var glow = new Pen(Color.White, FocusRingThickness * 2) { Alignment = PenAlignment.Center };
                using var ring = new Pen(Color.Red, FocusRingThickness) { Alignment = PenAlignment.Center };
                g.DrawPath(glow, path);
                g.DrawPath(ring, path);
            }

            // - Grid overlay (optional)
            DrawGridOverlay(g, vw, vh);
        }

        var fileName = $"{commandId}_{step}.png";
        var fullPath = Path.Combine(saveDir, fileName);
        bmp.Save(fullPath, ImageFormat.Png);
        var fullDataUrl = EncodeBitmapToDataUrl(bmp);

        string? focusUrl = null;
        Rectangle? rect = null;

        if (SendFocusCrop && explicitFocusRect is Rectangle rWant)
        {
            var rOk = ClampRect(rWant);
            using var crop = bmp.Clone(rOk, PixelFormat.Format24bppRgb);

            var cropName = $"{commandId}_{step}_crop.png";
            var cropPath = Path.Combine(saveDir, cropName);
            crop.Save(cropPath, ImageFormat.Png);

            focusUrl = EncodeBitmapToDataUrl(crop);
            rect = rOk;
        }

        // small FOCUS_UIA crop (optional)
        string? focusUiaDataUrl = null;
        string? focusUiaPath = null;
        var focusUiaRectAbs = GetFocusedUiaBoundingRect();
        if (IncludeFocusUiaCrop && focusUiaRectAbs is Rectangle frAbs2)
        {
            var local = ClampRect(new Rectangle(frAbs2.Left - vx, frAbs2.Top - vy, frAbs2.Width, frAbs2.Height));
            using var crop = bmp.Clone(local, PixelFormat.Format24bppRgb);
            focusUiaDataUrl = EncodeBitmapToDataUrl(crop);
            focusUiaPath = Path.Combine(saveDir, $"{commandId}_{step}_focus_uia.png");
            crop.Save(focusUiaPath, ImageFormat.Png);
        }

        return (fullDataUrl, fullPath, focusUrl, rect, focusUiaRectAbs, focusUiaDataUrl, focusUiaPath);
    }

    static void SaveAimOverlay(string baseShotPath, Rectangle rect, string overlayPath)
    {
        using var bmp = new Bitmap(baseShotPath);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.None;

        using var penYellow = new Pen(Color.Yellow, 3f) { DashStyle = DashStyle.Dash };
        using var penBlue = new Pen(Color.DeepSkyBlue, 1.5f) { DashStyle = DashStyle.Dash };

        var r = Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right - 1, rect.Bottom - 1);
        g.DrawRectangle(penYellow, r);
        var r2 = Rectangle.FromLTRB(Math.Max(0, r.Left - 1), Math.Max(0, r.Top - 1),
                                    Math.Min(bmp.Width - 1, r.Right + 1), Math.Min(bmp.Height - 1, r.Bottom + 1));
        g.DrawRectangle(penBlue, r2);

        bmp.Save(overlayPath, ImageFormat.Png);
    }

    static GraphicsPath RoundedRect(Rectangle r, int radius)
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

    static void DrawGridOverlay(Graphics g, int w, int h)
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

    static string EncodeBitmapToDataUrl(Bitmap bmp)
    {
        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        return $"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}";
    }

    static Rectangle SquareAround(int cx, int cy, int size)
    {
        int half = size / 2;
        return new Rectangle(cx - half, cy - half, size, size);
    }

    static Rectangle ClampRect(Rectangle r)
    {
        var (_, _, w, h) = GetPrimaryScreen();
        int left = Math.Max(0, Math.Min(w - 1, r.Left));
        int top = Math.Max(0, Math.Min(h - 1, r.Top));
        int right = Math.Max(left + 1, Math.Min(w, r.Right));
        int bottom = Math.Max(top + 1, Math.Min(h, r.Bottom));
        return Rectangle.FromLTRB(left, top, right, bottom);
    }

    // subfolders
    static string EnsureScreensDir()
    {
        var dir = Path.Combine(AppContext.BaseDirectory, "screens");
        Directory.CreateDirectory(dir);
        return dir;
    }
    static string EnsureRequestsDir()
    {
        var dir = Path.Combine(AppContext.BaseDirectory, "requests");
        Directory.CreateDirectory(dir);
        return dir;
    }
    static string EnsureLogDir()
    {
        var dir = Path.Combine(AppContext.BaseDirectory, "logs");
        Directory.CreateDirectory(dir);
        return dir;
    }

    static void SaveJson(string path, object o) =>
        File.WriteAllText(path, JsonSerializer.Serialize(o, PrettyJson), Encoding.UTF8);
    static void SaveRaw(string path, string raw) =>
        File.WriteAllText(path, raw, Encoding.UTF8);

    // PRIMARY screen
    static (int X, int Y, int W, int H) GetPrimaryScreen()
    {
        int w = GetSystemMetrics((int)SystemMetric.SM_CXSCREEN);
        int h = GetSystemMetrics((int)SystemMetric.SM_CYSCREEN);
        return (0, 0, w, h);
    }
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int nIndex);

    // --- Cursor pos (PRIMARY) ---
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT lpPoint);
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X; public int Y; }
    static (int X, int Y, double Nx, double Ny) GetCursorPositionInPrimary()
    {
        if (!GetCursorPos(out var p)) return (0, 0, 0, 0);
        var (vx, vy, vw, vh) = GetPrimaryScreen();
        int relX = Math.Max(0, Math.Min(vw - 1, p.X - vx));
        int relY = Math.Max(0, Math.Min(vh - 1, p.Y - vy));
        double nx = vw > 1 ? (double)relX / (vw - 1) : 0.0;
        double ny = vh > 1 ? (double)relY / (vh - 1) : 0.0;
        return (relX, relY, nx, ny);
    }

    static string Describe(ActionDto a)
    {
        if (a is null) return "null";

        // request_crop / point
        if (a.Type is "request_crop" or "point")
        {
            string prefix = a.Type == "point" ? "point" : "request_crop";
            if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                return $"{prefix} bbox=({a.Crop.Left},{a.Crop.Top})–({a.Crop.Right},{a.Crop.Bottom})";
            if (a.XPx is not null && a.YPx is not null)
                return $"{prefix} center=({a.XPx},{a.YPx})";
            if (a.X is not null && a.Y is not null)
            {
                var (pxN, pyN) = NormalizedToPixels(a.X.Value, a.Y.Value);
                return $"{prefix} center≈({pxN},{pyN})";
            }
            return $"{prefix} (missing parameters)";
        }

        // aim
        if (a.Type == "aim")
        {
            if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                return $"aim bbox=({a.BBox.Left},{a.BBox.Top})–({a.BBox.Right},{a.BBox.Bottom})";
            if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
                return $"aim(crop) bbox=({a.Crop.Left},{a.Crop.Top})–({a.Crop.Right},{a.Crop.Bottom})";
            if (a.XPx is not null && a.YPx is not null)
                return $"aim center=({a.XPx},{a.YPx})";
            if (a.X is not null && a.Y is not null)
            {
                var (pxN, pyN) = NormalizedToPixels(a.X.Value, a.Y.Value);
                return $"aim center≈({pxN},{pyN})";
            }
            return "aim (missing parameters)";
        }

        // move / click / double_click
        if (a.Type is "move" or "click" or "double_click")
        {
            // by bbox
            if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
            {
                int cx = (a.BBox.Left!.Value + a.BBox.Right!.Value) / 2;
                int cy = (a.BBox.Top!.Value + a.BBox.Bottom!.Value) / 2;
                string sfxB = a.Type == "click" ? $" {a.Button ?? "left"}"
                            : a.Type == "double_click" ? " (double)" : "";
                return $"{a.Type} bbox→({cx},{cy}){sfxB}";
            }

            // by explicit coordinates
            if (a.XPx is not null && a.YPx is not null)
            {
                string sfxP = a.Type == "click" ? $" {a.Button ?? "left"}"
                            : a.Type == "double_click" ? " (double)" : "";
                return $"{a.Type} ({a.XPx},{a.YPx}){sfxP}";
            }
            if (a.X is not null && a.Y is not null)
            {
                var (px, py) = NormalizedToPixels(a.X.Value, a.Y.Value);
                string sfxN = a.Type == "click" ? $" {a.Button ?? "left"}"
                            : a.Type == "double_click" ? " (double)" : "";
                return $"{a.Type} ({px},{py}){sfxN}";
            }

            // missing coords
            string sfx = a.Type == "click" ? $" {a.Button ?? "left"}"
                       : a.Type == "double_click" ? " (double)" : "";
            return $"{a.Type} (coords: missing){sfx}";
        }

        // scroll
        if (a.Type == "scroll")
        {
            int dy = a.ScrollDy ?? 0;
            return $"scroll dy={dy}";
        }

        // keys / type_text / wait / done
        if (a.Type == "keys") return $"keys [{string.Join("+", a.Keys ?? Array.Empty<string>())}]";
        if (a.Type == "type_text") return $"type_text \"{a.Text}\"";
        if (a.Type == "wait") return $"wait {Math.Max(0, a.WaitSeconds ?? 1)}s";
        if (a.Type == "done") return "done";

        return $"unknown {a.Type}";
    }

    static string ActionSignature(ActionDto a)
    {
        if (a == null || string.IsNullOrEmpty(a.Type)) return "null";
        string t = a.Type.ToLowerInvariant();

        if (t == "keys") return $"keys:{string.Join("+", (a.Keys ?? Array.Empty<string>()).Select(k => k.ToLowerInvariant()))}";
        if (t == "type_text") return "type_text";
        if (t == "scroll") return $"scroll:{a.ScrollDy ?? 0}";
        if (t == "wait") return $"wait:{Math.Max(0, a.WaitSeconds ?? 1)}";

        if (t is "move" or "click" or "double_click")
        {
            int cx, cy;
            if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
            {
                cx = (a.BBox.Left!.Value + a.BBox.Right!.Value) / 2;
                cy = (a.BBox.Top!.Value + a.BBox.Bottom!.Value) / 2;
            }
            else
            {
                var p = ResolvePoint(a);
                cx = p.X; cy = p.Y;
            }
            cx = (cx / 16) * 16;
            cy = (cy / 16) * 16;
            return $"{t}:{cx},{cy}";
        }

        if (t == "aim")
        {
            var rect = ResolveAimRect(a);
            if (rect is Rectangle r)
            {
                int cx = (r.Left + r.Right) / 2;
                int cy = (r.Top + r.Bottom) / 2;
                cx = (cx / 16) * 16;
                cy = (cy / 16) * 16;
                return $"aim:{cx},{cy}";
            }
            return "aim";
        }

        if (t == "point")
            return "point"; // helper, do not gate clicks with this signature

        return t;
    }

    static Rectangle? ResolveAimRect(ActionDto a)
    {
        if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
            return ClampRect(RectFromBBox(a.BBox));
        if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
            return ClampRect(RectFromBBox(a.Crop));
        if (a.XPx.HasValue && a.YPx.HasValue)
            return ClampRect(SquareAround(a.XPx.Value, a.YPx.Value, FocusCropSize));
        if (a.X.HasValue && a.Y.HasValue)
        {
            var (px, py) = NormalizedToPixels(a.X.Value, a.Y.Value);
            return ClampRect(SquareAround(px, py, FocusCropSize));
        }
        return null;
    }

    static Rectangle? ResolveCropRect(ActionDto a)
    {
        if (a.Crop is { Left: { }, Top: { }, Right: { }, Bottom: { } })
            return ClampRect(RectFromBBox(a.Crop));
        if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
            return ClampRect(RectFromBBox(a.BBox));
        if (a.XPx.HasValue && a.YPx.HasValue)
            return ClampRect(SquareAround(a.XPx.Value, a.YPx.Value, FocusCropSize));
        if (a.X.HasValue && a.Y.HasValue)
        {
            var (px, py) = NormalizedToPixels(a.X.Value, a.Y.Value);
            return ClampRect(SquareAround(px, py, FocusCropSize));
        }
        return null;
    }

    static Rectangle RectFromBBox(BBox b) =>
        Rectangle.FromLTRB(b.Left!.Value, b.Top!.Value, b.Right!.Value, b.Bottom!.Value);

    // ===== Action execution =====
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    static void ExecuteAction(ActionDto action)
    {
        switch (action.Type)
        {
            case "move":
                {
                    var (x, y) = ResolvePoint(action);
                    SetCursorPos(x, y);
                    break;
                }
            case "click":
                {
                    var (x, y) = ResolvePoint(action);
                    SetCursorPos(x, y);
                    MouseClick(action.Button ?? "left");
                    break;
                }
            case "double_click":
                {
                    var (x, y) = ResolvePoint(action);
                    SetCursorPos(x, y);
                    MouseDoubleClick(action.Button ?? "left");
                    break;
                }
            case "keys":
                {
                    if (action.Keys is null || action.Keys.Length == 0)
                        throw new InvalidOperationException("Missing keys");
                    PressKeysSmart(action.Keys);
                    break;
                }
            case "type_text":
                {
                    if (action.Text is null) throw new InvalidOperationException("Missing text");
                    TypeUnicodeString(action.Text);
                    break;
                }
            case "scroll":
                {
                    int dy = action.ScrollDy ?? 0;
                    if (dy != 0) MouseScroll(dy);
                    break;
                }
            case "request_crop":
            case "point":
            case "aim":
            case "wait": // handled in loop (await Task.Delay), nothing here
            case "done":
                break;

            default:
                throw new InvalidOperationException($"Unknown action type: {action.Type}");
        }
    }

    static (int X, int Y) ResolvePoint(ActionDto a)
    {
        if (a.BBox is { Left: { }, Top: { }, Right: { }, Bottom: { } })
        {
            int cx = (a.BBox.Left!.Value + a.BBox.Right!.Value) / 2;
            int cy = (a.BBox.Top!.Value + a.BBox.Bottom!.Value) / 2;
            return (cx, cy);
        }
        if (a.XPx.HasValue && a.YPx.HasValue) return (a.XPx.Value, a.YPx.Value);
        if (a.X.HasValue && a.Y.HasValue) return NormalizedToPixels(a.X.Value, a.Y.Value);
        throw new InvalidOperationException("Missing coordinates (bbox or x_px/y_px or x/y).");
    }

    static bool HasExplicitPoint(ActionDto a) =>
        (a.BBox is { Left: not null, Top: not null, Right: not null, Bottom: not null })
        || (a.XPx is not null && a.YPx is not null)
        || (a.X is not null && a.Y is not null);

    static (int X, int Y) NormalizedToPixels(double nx, double ny)
    {
        var (vx, vy, vw, vh) = GetPrimaryScreen();
        int x = vx + (int)Math.Round(nx * (vw - 1));
        int y = vy + (int)Math.Round(ny * (vh - 1));
        return (x, y);
    }

    // ==== Mouse ====
    static void MouseClick(string button)
    {
        uint down, up;
        switch ((button ?? "left").ToLowerInvariant())
        {
            case "right": down = 0x0008; up = 0x0010; break;
            case "middle": down = 0x0020; up = 0x0040; break;
            default: down = 0x0002; up = 0x0004; break; // left
        }
        var inputs = new INPUT[]
        {
            new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = down } } },
            new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = up } } },
        };
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
    }

    static void MouseDoubleClick(string button)
    {
        MouseClick(button);
        Thread.Sleep(80);
        MouseClick(button);
    }

    static void MouseScroll(int dyLines)
    {
        const int WHEEL_DELTA = 120;
        int delta = dyLines * WHEEL_DELTA; // positive is down
        var inputs = new INPUT[]
        {
            new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = 0x0800, mouseData = (uint)delta } } } // MOUSEEVENTF_WHEEL
        };
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
    }

    // ==== Keyboard ====
    static void PressKeysSmart(string[] keys)
    {
        if (keys is null || keys.Length == 0)
            throw new InvalidOperationException("Missing keys");

        // Recognized modifiers in chords
        var modifiers = new HashSet<string>(
            new[] { "ctrl", "alt", "shift", "win", "super", "meta", "cmd" },
            StringComparer.OrdinalIgnoreCase
        );

        // 1) Back-compat: ["ctrl","esc"] → treat as one chord
        if (keys.Length >= 2 && keys.Take(keys.Length - 1).All(k => modifiers.Contains(k)))
        {
            var mods = keys.Take(keys.Length - 1).ToArray();
            var main = keys.Last();
            foreach (var m in mods) KeyDown(m);
            PressKey(main);
            foreach (var m in mods.Reverse()) KeyUp(m);
            return;
        }

        // 2) New mode: items may contain pluses, e.g., ["ctrl+shift+esc", "win+r", "tab"]
        foreach (var item in keys)
        {
            if (string.IsNullOrWhiteSpace(item))
                continue;

            var parts = item.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            // Single-string chord: modifiers + main key
            if (parts.Length >= 2 && parts.Take(parts.Length - 1).All(p => modifiers.Contains(p)))
            {
                var mods = parts.Take(parts.Length - 1).ToArray();
                var main = parts.Last();
                foreach (var m in mods) KeyDown(m);
                PressKey(main);
                foreach (var m in mods.Reverse()) KeyUp(m);
            }
            else
            {
                // Single key (e.g., "tab", "esc", "f5")
                PressKey(item);
            }
        }
    }

    static void PressKey(string key)
    {
        if (string.IsNullOrEmpty(key))
            throw new ArgumentException("Empty key");

        if (key == " " || key.Equals("space", StringComparison.OrdinalIgnoreCase))
        {
            TypeUnicodeString(" ");
            return;
        }

        if (key.Length == 1)
        {
            char ch = key[0];
            bool asciiLetter = (ch is >= 'a' and <= 'z') || (ch is >= 'A' and <= 'Z');
            bool asciiDigit = (ch is >= '0' and <= '9');
            if (!asciiLetter && !asciiDigit)
            {
                TypeUnicodeString(key);
                return;
            }
        }

        KeyDown(key);
        KeyUp(key);
    }

    static void KeyDown(string key) { SendKeyboard(KeyNameToVk(key), false); }
    static void KeyUp(string key) { SendKeyboard(KeyNameToVk(key), true); }

    static ushort KeyNameToVk(string key)
    {
        key = key.ToLowerInvariant();
        if (key is "super" or "meta" or "cmd") key = "win";
        if (key == " ") return 0x20; // VK_SPACE

        return key switch
        {
            "ctrl" => 0x11,
            "alt" => 0x12,
            "shift" => 0x10,
            "win" => 0x5B,
            "enter" or "return" => 0x0D,
            "tab" => 0x09,
            "esc" or "escape" => 0x1B,
            "space" => 0x20,
            "backspace" => 0x08,
            "delete" => 0x2E,
            "left" => 0x25,
            "up" => 0x26,
            "right" => 0x27,
            "down" => 0x28,
            "home" => 0x24,
            "end" => 0x23,
            "pageup" => 0x21,
            "pagedown" => 0x22,
            _ => MapAlnumToVk(key)
        };
    }

    static ushort MapAlnumToVk(string key)
    {
        if (key.Length == 1)
        {
            char c = key[0];
            if (char.IsLetter(c)) return (ushort)char.ToUpperInvariant(c); // 'A'..'Z'
            if (char.IsDigit(c)) return (ushort)c;                         // '0'..'9'
        }
        if (key.StartsWith("f", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(key[1..], out int f) && f is >= 1 and <= 24)
            return (ushort)(0x70 + (f - 1)); // F1..F24

        throw new ArgumentException($"Unknown key: {key}");
    }

    static void TypeUnicodeString(string s)
    {
        foreach (var ch in s)
        {
            SendUnicodeKey(ch, false);
            SendUnicodeKey(ch, true);
        }
    }

    // ===== WinAPI / SendInput P/Invoke =====
    [StructLayout(LayoutKind.Sequential)]
    struct INPUT { public uint type; public InputUnion U; }

    [StructLayout(LayoutKind.Explicit)]
    struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT
    {
        public int dx, dy;
        public uint mouseData, dwFlags, time;
        public nint dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT
    {
        public ushort wVk, wScan;
        public uint dwFlags, time;
        public nint dwExtraInfo;
    }

    static void SendKeyboard(ushort vk, bool keyUp)
    {
        var input = new INPUT
        {
            type = 1,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = vk,
                    wScan = 0,
                    dwFlags = keyUp ? 0x0002u : 0u, // KEYEVENTF_KEYUP
                    time = 0,
                    dwExtraInfo = 0
                }
            }
        };
        if (SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>()) == 0)
            throw new InvalidOperationException("SendInput VK failed.");
    }

    static void SendUnicodeKey(char ch, bool keyUp)
    {
        var input = new INPUT
        {
            type = 1,
            U = new InputUnion
            {
                ki = new KEYBDINPUT
                {
                    wVk = 0,
                    wScan = ch,
                    dwFlags = 0x0004u | (keyUp ? 0x0002u : 0u), // KEYEVENTF_UNICODE | KEYUP?
                    time = 0,
                    dwExtraInfo = 0
                }
            }
        };
        if (SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>()) == 0)
            throw new InvalidOperationException("SendInput UNICODE failed.");
    }

    // ==== Image delta (0..1) ====
    static double ComputeImageDelta(string pathA, string pathB, int w = 96, int h = 54)
    {
        using var a = new Bitmap(pathA);
        using var b = new Bitmap(pathB);
        using var ra = ResizeTo(a, w, h);
        using var rb = ResizeTo(b, w, h);

        double sum = 0;
        int n = w * h;
        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                var ca = ra.GetPixel(x, y);
                var cb = rb.GetPixel(x, y);
                double la = (0.299 * ca.R + 0.587 * ca.G + 0.114 * ca.B);
                double lb = (0.299 * cb.R + 0.587 * cb.G + 0.114 * cb.B);
                sum += Math.Abs(la - lb) / 255.0;
            }
        }
        return sum / n;
    }

    static Bitmap ResizeTo(Bitmap src, int w, int h)
    {
        var dst = new Bitmap(w, h, PixelFormat.Format24bppRgb);
        using var g = Graphics.FromImage(dst);
        g.InterpolationMode = InterpolationMode.HighQualityBilinear;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        g.SmoothingMode = SmoothingMode.None;
        g.DrawImage(src, new Rectangle(0, 0, w, h));
        return dst;
    }

    static string Tail(string s, int n)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s.Length <= n ? s : s[^n..];
    }

    // === UI Automation – bounding rect of the currently focused element ===
    // Returns pixel coords in screen space (VirtualScreen/Primary, referenced to (0,0) of primary).
    static Rectangle? GetFocusedUiaBoundingRect()
    {
        try
        {
            var el = AutomationElement.FocusedElement;
            if (el is null) return null;

            // 1) BoundingRectangle from UIA
            var rectObj = el.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
            if (rectObj is System.Windows.Rect r1 && IsUsable(r1))
                return WpfRectToGdi(r1);

            // 2) Fallback: window rect from HWND
            var hwndObj = el.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty, true);
            if (hwndObj is int hwnd && hwnd != 0 && GetWindowRect(new IntPtr(hwnd), out var wr))
            {
                var r2 = new System.Windows.Rect(wr.Left, wr.Top, wr.Right - wr.Left, wr.Bottom - wr.Top);
                if (IsUsable(r2)) return WpfRectToGdi(r2);
            }

            // 3) Ultimately: walk up ControlView parents
            var walker = TreeWalker.ControlViewWalker;
            var cur = walker.GetParent(el);
            int hops = 0;
            while (cur is not null && hops++ < 8)
            {
                var ro = cur.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
                if (ro is System.Windows.Rect r3 && IsUsable(r3)) return WpfRectToGdi(r3);
                cur = walker.GetParent(cur);
            }
        }
        catch { /* ignore */ }
        return null;

        static bool IsUsable(System.Windows.Rect r) => !r.IsEmpty && r.Width > 1 && r.Height > 1;
        static Rectangle WpfRectToGdi(System.Windows.Rect r) =>
            Rectangle.FromLTRB((int)Math.Round(r.Left), (int)Math.Round(r.Top),
                               (int)Math.Round(r.Right), (int)Math.Round(r.Bottom));
    }

    // P/Invoke GetWindowRect (UIA fallback)
    [DllImport("user32.dll", SetLastError = true)]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left, Top, Right, Bottom; }
}

// === enum outside class ===
public enum SystemMetric
{
    SM_CXSCREEN = 0,
    SM_CYSCREEN = 1,
    SM_XVIRTUALSCREEN = 76,
    SM_YVIRTUALSCREEN = 77,
    SM_CXVIRTUALSCREEN = 78,
    SM_CYVIRTUALSCREEN = 79
}

// === DTOs (control) ===
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

    [JsonPropertyName("scroll_dy")] public int? ScrollDy { get; set; } // positive = down

    [JsonPropertyName("bbox")] public BBox? BBox { get; set; }      // target/click
    [JsonPropertyName("crop")] public BBox? Crop { get; set; }      // request_crop / aim / point

    [JsonPropertyName("wait_seconds")] public int? WaitSeconds { get; set; } // wait duration for 'wait'
    [JsonPropertyName("note")] public string? Note { get; set; }    // short comment (required in schema)
}

// === DTO Q&A (text + location) ===
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

public sealed class VerifyDto
{
    [JsonPropertyName("verdict")] public string? Verdict { get; set; } // "yes"|"no"
    [JsonPropertyName("reason")] public string? Reason { get; set; }
}

public sealed class BBox
{
    [JsonPropertyName("left")] public int? Left { get; set; }
    [JsonPropertyName("top")] public int? Top { get; set; }
    [JsonPropertyName("right")] public int? Right { get; set; }
    [JsonPropertyName("bottom")] public int? Bottom { get; set; }
}

// === Tee writer (log to file + stdout) ===
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
