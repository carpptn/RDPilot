internal static partial class RDPilotApplication
{
    /// <summary>
    /// Builds control, question-answering, and verification prompts and request payloads.
    /// </summary>
    internal static class PromptAndRequestFactory
    {
            // === System rules (control) ===
            internal static string BuildSystemRules()
            {
                var sb = new StringBuilder();
                sb.AppendLine("You are an agent that controls a Windows 10/11 computer through safe local actions and, when needed, the visible UI.");
                sb.AppendLine("Return EXACTLY ONE action per round as JSON matching the provided schema.");
                sb.AppendLine("Return one JSON object only. Do not include alternative plans or additional candidate actions.");
                sb.AppendLine();
                sb.AppendLine("Important: The screenshot may contain a white+red rounded rectangle overlay – that's the element with current keyboard focus (FOCUS_UIA). Treat it as a reliable source of truth.");
                sb.AppendLine();
                sb.AppendLine("Keep 'note' short (max 120 chars). In fast mode, prefer an empty note over a long explanation.");
                sb.AppendLine("Set confidence from 0.0 to 1.0 for how certain you are that the single returned action is the right next step.");
                sb.AppendLine();
                sb.AppendLine("Guidelines:");
                sb.AppendLine("- Prefer the shortest safe path through the real UI.");
                if (AllowHighLevelActions)
                {
                    sb.AppendLine("- High-level local actions are enabled. Use 'open_url'/'launch_app' only when they directly satisfy the next UI step.");
                }
                else
                {
                    sb.AppendLine("- High-level local actions are disabled; operate via visible UI, keyboard, mouse, clipboard paste, and UIA targets.");
                }
                sb.AppendLine("- Use 'paste_text' for long text blocks; it uses the clipboard and is faster than key-by-key typing.");
                if (AllowRunCommand)
                    sb.AppendLine("- 'run_command' is allowed for deterministic local commands. Keep it explicit and minimal.");
                else
                sb.AppendLine("- 'run_command' is not available unless explicitly enabled by the operator.");
                sb.AppendLine("- For text input use 'type_text' (full UNICODE string). Use 'keys' only for shortcuts and function keys.");
                sb.AppendLine($"- Keep each 'type_text'/'paste_text' text value under {MaxActionTextChars} characters. For longer content, split it across multiple consecutive paste_text actions; do not emit oversized JSON.");
                sb.AppendLine("- If the active target window is too small, clipped, or partly off-screen, fix it yourself through the real UI before precise work. Prefer a single 'keys' action like [\"win\",\"up\"] to maximize; use Win+Left/Win+Right or Alt+Space keyboard window commands only when needed.");
                sb.AppendLine("- WINDOW_VISIBILITY_HINT is advisory metadata; RDPilot will not automatically move or maximize the target window for you.");
                sb.AppendLine();
                if (MouseEnabled)
                {
                    sb.AppendLine("- MOUSE_ALLOWED: true.");
                    sb.AppendLine("- Work in SCREEN_SIZE pixel coordinates from the screenshot image in this request (0,0 is top-left). Do not use REAL_SCREEN_SIZE for action coordinates.");
                    sb.AppendLine("- Even when a FOCUS_CROP image is shown, x_px/y_px/bbox/crop must use full screenshot SCREEN_SIZE coordinates, not crop-local coordinates.");
                    sb.AppendLine("- When clicking, aim at the center of the bbox.");
                    if (IncludeUiaTargets && MaxUiaTargets > 0)
                        sb.AppendLine("- Use 'focus_uia' or 'click_uia' when a target appears in UIA_TARGETS; choose by uia_index.");
                    sb.AppendLine("- For click/double_click, prefer a bbox for the target. If using x_px/y_px, choose a safe interior point near the visual center, never the top-left corner.");
                    sb.AppendLine("- One action per round. Decide solely from the screenshot and metadata (SCREEN_SIZE, CURSOR_POS, FOCUS_UIA/FOCUS_CROP, UIA_TARGETS, DELTA/REPEAT).");
                    sb.AppendLine("- If the target is ambiguous, prefer actions relative to FOCUS_UIA (e.g., TAB/Shift+TAB or aim at the center of FOCUS_UIA).");
                }
                else
                {
                    sb.AppendLine("- MOUSE_ALLOWED: false. Do not use move/click/double_click/scroll/focus_uia/click_uia.");
                    sb.AppendLine("- Prefer keyboard navigation, shortcuts, paste_text, type_text, TAB/Shift+TAB, Enter, Escape, and application accelerators.");
                    sb.AppendLine("- Use 'request_crop' or 'point' only when a closer visual look is needed; they do not interact with the app.");
                    sb.AppendLine("- request_crop/point coordinates use full screenshot SCREEN_SIZE coordinates, not crop-local or REAL_SCREEN_SIZE coordinates.");
                    sb.AppendLine("- One action per round. Decide solely from the screenshot and metadata (SCREEN_SIZE, CURSOR_POS, FOCUS_UIA, DELTA/REPEAT).");
                }
                sb.AppendLine();
                sb.AppendLine("- Return 'done' ONLY when the screen state clearly confirms the goal. Set high confidence only when no extra verification should be needed.");
                sb.AppendLine("- DO NOT use machine-specific taskbar/app-number shortcuts: Win+1..9, Super+1..9, etc.");
                sb.AppendLine("- Prefer deterministic strategies. If a proposed action may be nondeterministic – choose an alternative.");
                sb.AppendLine();
                if (MouseEnabled)
                {
                    if (DirectClickWithoutAim)
                        sb.AppendLine("- You may click directly when the target bbox/point is large and unambiguous. Use 'aim' first for small or uncertain targets.");
                    else
                        sb.AppendLine("- BEFORE any 'click'/'double_click' you MUST set an 'aim' (the target region). Clicks outside the active AIM are ignored.");
                    sb.AppendLine("- After setting AIM, ensure the intended target is visible within the AIM frame. If not, re-aim until it is.");
                    sb.AppendLine($"- After a large visual change (LAST_STEP_DELTA > {AimExpireDelta:0.###}) the previous AIM expires; set a new one before clicking.");
                    sb.AppendLine($"- Define 'aim' via 'bbox' (preferred) or a point (x/y or x_px/y_px); in the latter case the crop is a square of ~{FocusCropSize}px.");
                    sb.AppendLine("- 'request_crop' and 'point' are only for requesting zoom/homing; they do NOT replace 'aim'.");
                    sb.AppendLine("- If an active AIM exists, in 'click'/'double_click' you MUST PROVIDE COORDINATES inside AIM (do not rely on implicit centering).");
                    sb.AppendLine("- 'double_click' means a standard double click (e.g., launch an app). 'button:right' in 'click' = e.g., context menu.");
                    sb.AppendLine();
                }
                sb.AppendLine("- Use 'wait' when a long-running process is visible (e.g., progress bar, installer, render, upload).");
                sb.AppendLine(MaxWaitSeconds > 0
                    ? $"  Set 'wait_seconds' to a realistic duration up to {MaxWaitSeconds}s; during 'wait' no screenshots are taken. Reassess the screen afterwards."
                    : "  Set 'wait_seconds' to a realistic duration; during 'wait' no screenshots are taken. Reassess the screen afterwards.");
                if (GridStepPx > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine($"- The screenshot contains a semi-transparent grid every {GridStepPx} pixels; use it to provide precise x_px/y_px.");
                    sb.AppendLine("- The origin (0,0) is the top-left corner of the screen.");
                }
                return sb.ToString();
            }
        
            // === Q&A rules ===
            internal static string BuildQaSystemRules()
            {
                var sb = new StringBuilder();
                sb.AppendLine("You are a screen analyst. Answer strictly based on the screenshot, metadata (SCREEN_SIZE, CURSOR_POS), and the user's question.");
                sb.AppendLine("The image may include a white+red rounded rectangle overlay – that's the element with current keyboard focus (FOCUS_UIA). Treat it as a reliable focus indicator.");
                sb.AppendLine("Return BOTH: a short textual answer and location metadata for the most relevant element. Add a short 'note'.");
                sb.AppendLine("Always think in SCREEN_SIZE image pixel coordinates (0,0 top-left). If a location makes sense, choose the center of the visible bbox.");
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
            internal static Dictionary<string, object> NullableIntegerSchema(int minimum = 0, int? maximum = null)
            {
                var schema = new Dictionary<string, object>
                {
                    ["type"] = new object[] { "integer", "null" },
                    ["minimum"] = minimum
                };
                if (maximum.HasValue)
                    schema["maximum"] = maximum.Value;
                return schema;
            }
        
            internal static object BoxSchema(int screenW, int screenH) => new
            {
                type = new object[] { "object", "null" },
                additionalProperties = false,
                properties = new Dictionary<string, object>
                {
                    ["left"] = NullableIntegerSchema(0, Math.Max(0, screenW - 1)),
                    ["top"] = NullableIntegerSchema(0, Math.Max(0, screenH - 1)),
                    ["right"] = NullableIntegerSchema(0, Math.Max(0, screenW)),
                    ["bottom"] = NullableIntegerSchema(0, Math.Max(0, screenH)),
                },
                required = new[] { "left", "top", "right", "bottom" }
            };
        
            internal static object WaitSecondsSchema()
            {
                var schema = NullableIntegerSchema();
                if (MaxWaitSeconds > 0)
                    schema["maximum"] = MaxWaitSeconds;
                return schema;
            }
        
            internal static object ScrollDySchema() => new
            {
                type = new object[] { "integer", "null" },
                minimum = -20,
                maximum = 20
            };
        
            internal static string[] ControlActionTypes()
            {
                var types = new List<string>
                {
                    "paste_text",
                    "keys", "type_text",
                    "request_crop", "point", "aim", "wait", "done"
                };
                if (MouseEnabled)
                {
                    var mouseTypes = new List<string> { "move", "click", "double_click", "scroll" };
                    if (CurrentUiaTargets.Count > 0)
                        mouseTypes.InsertRange(0, new[] { "focus_uia", "click_uia" });
                    types.InsertRange(1, mouseTypes);
                }
                if (AllowHighLevelActions)
                    types.InsertRange(0, new[] { "open_url", "launch_app" });
                if (AllowRunCommand)
                    types.Insert(AllowHighLevelActions ? 2 : 0, "run_command");
                return types.ToArray();
            }
        
            // === Request build (control) ===
            internal static object BuildRequestBody(
                string model, string systemRules, string goal, string historyPlusMeta,
                string imageDataUrl, int screenW, int screenH,
                int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
                string? focusDataUrl, Rectangle? focusRect,
                Rectangle? focusUiaRect, string? focusUiaDataUrl,
                UiPromptContext promptContext,
                bool reuseUiaTargets,
                string? previousResponseId,
                bool omitFullScreenImage = false)
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
                            ["type"] = new { type = "string", @enum = ControlActionTypes() },
                            ["x"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                            ["y"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                            ["x_px"] = NullableIntegerSchema(0, Math.Max(0, screenW - 1)),
                            ["y_px"] = NullableIntegerSchema(0, Math.Max(0, screenH - 1)),
                            ["button"] = new { type = new object[] { "string", "null" }, @enum = new object[] { "left", "right", "middle", null! } },
                            ["keys"] = new { type = new object[] { "array", "null" }, items = new { type = "string", maxLength = 32 }, maxItems = 8 },
                            ["text"] = new { type = new object[] { "string", "null" }, maxLength = MaxActionTextChars },
                            ["url"] = new { type = new object[] { "string", "null" } },
                            ["app"] = new { type = new object[] { "string", "null" } },
                            ["command"] = new { type = new object[] { "string", "null" } },
                            ["uia_index"] = NullableIntegerSchema(0, Math.Max(0, CurrentUiaTargets.Count - 1)),
                            ["scroll_dy"] = ScrollDySchema(), // positive = down
                            ["bbox"] = BoxSchema(screenW, screenH),
                            ["crop"] = BoxSchema(screenW, screenH),
                            ["wait_seconds"] = WaitSecondsSchema(),
                            ["confidence"] = new { type = new object[] { "number", "null" }, minimum = 0.0, maximum = 1.0 },
                            ["note"] = new { type = new object[] { "string", "null" }, maxLength = 120 }
                        },
                        required = new[] { "type", "x", "y", "x_px", "y_px", "button", "keys", "text", "url", "app", "command", "uia_index", "scroll_dy", "bbox", "crop", "wait_seconds", "confidence", "note" }
                    }
                };
        
                var stableUserText = new StringBuilder()
                    .AppendLine($"GOAL: {goal}")
                    .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px; coordinate space for x_px/y_px/bbox/crop)")
                    .AppendLine($"MOUSE_ALLOWED: {(MouseEnabled ? "true" : "false")}");
                if (CurrentScreenMap.IsScaled)
                    stableUserText.AppendLine($"REAL_SCREEN_SIZE: width={CurrentScreenMap.ScreenW}, height={CurrentScreenMap.ScreenH} (px; RDPilot maps SCREEN_SIZE coordinates to the real screen)");
        
                var userText = new StringBuilder()
                    .AppendLine("HISTORY:")
                    .AppendLine(historyPlusMeta)
                    .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})")
                    .AppendLine($"ACTIVE_WINDOW: {promptContext.ActiveWindowTitle}")
                    .AppendLine($"ACTIVE_PROCESS: {promptContext.ActiveProcessName}");
                AppendActiveWindowGeometry(userText, promptContext);
        
                if (focusUiaRect.HasValue)
                {
                    var r = focusUiaRect.Value;
                    int cx = (r.Left + r.Right) / 2;
                    int cy = (r.Top + r.Bottom) / 2;
                    userText.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
                    userText.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
                    AppendFocusedUiaSummary(userText, promptContext.FocusedUiaSummary);
                }
                else
                {
                    userText.AppendLine("FOCUS_UIA: none");
                }
                AppendBlockingPromptHint(userText, promptContext);
        
                if (focusRect.HasValue)
                {
                    var r = focusRect.Value;
                    userText.AppendLine($"FOCUS_CROP: left={r.Left}, top={r.Top}, width={r.Width}, height={r.Height} (px). The crop image is primary for local detail; the full-screen image is a small overview.");
                }
        
                AppendUiaTargets(userText, reuseUiaTargets, reuseUiaTargets ? "screen unchanged" : null);
        
                var userContent = new List<object>
                {
                    new { type = "input_text",  text = stableUserText.ToString() },
                    new { type = "input_text",  text = userText.ToString() }
                };
                if (focusDataUrl != null)
                    userContent.Add(new { type = "input_image", image_url = focusDataUrl });
                if (IncludeFocusUiaCrop && focusUiaDataUrl != null)
                    userContent.Add(new { type = "input_image", image_url = focusUiaDataUrl });
                if (!omitFullScreenImage)
                    userContent.Add(new { type = "input_image", image_url = imageDataUrl });
        
                var req = new Dictionary<string, object>
                {
                    ["model"] = model,
                    ["input"] = new object[]
                    {
                        new { role = "system", content = new object[] { new { type = "input_text", text = systemRules } } },
                        new { role = "user",   content = userContent.ToArray() }
                    },
                    ["text"] = TextOptions(format)
                };
                if (SupportsTemperature(model))
                    req["temperature"] = 0.0;
                if (!string.IsNullOrWhiteSpace(previousResponseId))
                    req["previous_response_id"] = previousResponseId;
                AddReasoningOptions(req, model, cacheScope: "control");
        
                return req;
            }
        
            // Logging variant (no base64)
            internal static object BuildRequestBody_ForLog(
                string model, string systemRules, string goal, string historyPlusMeta,
                string? fullImagePath, int screenW, int screenH,
                int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
                string? focusImagePath, Rectangle? focusRect,
                Rectangle? focusUiaRect, string? focusUiaImagePath,
                UiPromptContext promptContext,
                string? previousResponseId,
                bool omitFullScreenImage = false)
            {
                var format = new { name = "SingleAction" };
                var stableUserText = new StringBuilder()
                    .AppendLine($"GOAL: {goal}")
                    .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px; coordinate space for x_px/y_px/bbox/crop)")
                    .AppendLine($"MOUSE_ALLOWED: {(MouseEnabled ? "true" : "false")}");
                if (CurrentScreenMap.IsScaled)
                    stableUserText.AppendLine($"REAL_SCREEN_SIZE: width={CurrentScreenMap.ScreenW}, height={CurrentScreenMap.ScreenH} (px; RDPilot maps SCREEN_SIZE coordinates to the real screen)");
        
                var userText = new StringBuilder()
                    .AppendLine("HISTORY:")
                    .AppendLine(historyPlusMeta)
                    .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})")
                    .AppendLine($"ACTIVE_WINDOW: {promptContext.ActiveWindowTitle}")
                    .AppendLine($"ACTIVE_PROCESS: {promptContext.ActiveProcessName}");
                AppendActiveWindowGeometry(userText, promptContext);
        
                if (focusUiaRect.HasValue)
                {
                    var r = focusUiaRect.Value;
                    int cx = (r.Left + r.Right) / 2;
                    int cy = (r.Top + r.Bottom) / 2;
                    userText.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
                    userText.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
                    AppendFocusedUiaSummary(userText, promptContext.FocusedUiaSummary);
                }
                else
                {
                    userText.AppendLine("FOCUS_UIA: none");
                }
                AppendBlockingPromptHint(userText, promptContext);
        
                if (focusRect.HasValue)
                {
                    var r = focusRect.Value;
                    userText.AppendLine($"FOCUS_CROP: left={r.Left}, top={r.Top}, width={r.Width}, height={r.Height} (px). The crop image is primary for local detail; the full-screen image is a small overview.");
                }
        
                AppendUiaTargets(userText, reuseExisting: true, reuseReason: "same request");
        
                var userContent = new List<object>
                {
                    new { type = "input_text",  text = stableUserText.ToString() },
                    new { type = "input_text",  text = userText.ToString() }
                };
                if (focusImagePath != null)
                    userContent.Add(new { type = "input_image", image_url = LogImageRef(focusImagePath) });
                if (IncludeFocusUiaCrop && focusUiaImagePath != null)
                    userContent.Add(new { type = "input_image", image_url = LogImageRef(focusUiaImagePath) });
                if (!omitFullScreenImage)
                    userContent.Add(new { type = "input_image", image_url = LogImageRef(fullImagePath) });
        
                var req = new Dictionary<string, object>
                {
                    ["model"] = model,
                    ["input"] = new object[]
                    {
                        new { role = "system", content = new object[] { new { type = "input_text", text = systemRules } } },
                        new { role = "user",   content = userContent.ToArray() }
                    },
                    ["text"] = TextOptions(format)
                };
                if (SupportsTemperature(model))
                    req["temperature"] = 0.0;
                if (!string.IsNullOrWhiteSpace(previousResponseId))
                    req["previous_response_id"] = previousResponseId;
                AddReasoningOptions(req, model, cacheScope: "control");
        
                return req;
            }
        
            // === Request build (Q&A) ===
            internal static object BuildQARequestBody(string model, string qaRules, string question, string imageDataUrl,
                                             int screenW, int screenH, int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
                                             Rectangle? focusUiaRect, string? focusUiaDataUrl,
                                             UiPromptContext promptContext)
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
                            ["x_px"] = NullableIntegerSchema(0, Math.Max(0, screenW - 1)),
                            ["y_px"] = NullableIntegerSchema(0, Math.Max(0, screenH - 1)),
                            ["bbox"] = BoxSchema(screenW, screenH),
                            ["note"] = new { type = new object[] { "string", "null" } },
                        },
                        required = new[] { "answer_text", "x", "y", "x_px", "y_px", "bbox", "note" }
                    }
                };
        
                var meta = new StringBuilder()
                    .AppendLine($"QUESTION: {question}")
                    .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px; coordinate space for x_px/y_px/bbox)")
                    .AppendLine($"ACTIVE_WINDOW: {promptContext.ActiveWindowTitle}")
                    .AppendLine($"ACTIVE_PROCESS: {promptContext.ActiveProcessName}")
                    .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})");
                if (CurrentScreenMap.IsScaled)
                    meta.AppendLine($"REAL_SCREEN_SIZE: width={CurrentScreenMap.ScreenW}, height={CurrentScreenMap.ScreenH} (px)");
                AppendActiveWindowGeometry(meta, promptContext);
        
                if (focusUiaRect.HasValue)
                {
                    var r = focusUiaRect.Value;
                    int cx = (r.Left + r.Right) / 2;
                    int cy = (r.Top + r.Bottom) / 2;
                    meta.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
                    meta.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
                    AppendFocusedUiaSummary(meta, promptContext.FocusedUiaSummary);
                }
                else
                {
                    meta.AppendLine("FOCUS_UIA: none");
                }
                AppendBlockingPromptHint(meta, promptContext);
        
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
                    ["text"] = TextOptions(format)
                };
                if (SupportsTemperature(model))
                    req["temperature"] = 0.0;
                AddReasoningOptions(req, model, EffectiveQaReasoningEffort(), QaMaxOutputTokens, "qa");
        
                return req;
            }
        
            internal static object BuildQARequestBody_ForLog(string model, string qaRules, string question, string? fullImagePath,
                                                    int screenW, int screenH, int cursorXPx, int cursorYPx, double cursorXN, double cursorYN,
                                                    Rectangle? focusUiaRect, string? focusUiaImagePath,
                                                    UiPromptContext promptContext)
            {
                if (question.StartsWith("/ask ", StringComparison.OrdinalIgnoreCase))
                    question = question[5..];
        
                var meta = new StringBuilder()
                    .AppendLine($"QUESTION: {question}")
                    .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px; coordinate space for x_px/y_px/bbox)")
                    .AppendLine($"ACTIVE_WINDOW: {promptContext.ActiveWindowTitle}")
                    .AppendLine($"ACTIVE_PROCESS: {promptContext.ActiveProcessName}")
                    .AppendLine($"CURSOR_POS: x={cursorXPx}, y={cursorYPx} px | normalized=({cursorXN:0.###},{cursorYN:0.###})");
                if (CurrentScreenMap.IsScaled)
                    meta.AppendLine($"REAL_SCREEN_SIZE: width={CurrentScreenMap.ScreenW}, height={CurrentScreenMap.ScreenH} (px)");
                AppendActiveWindowGeometry(meta, promptContext);
        
                if (focusUiaRect.HasValue)
                {
                    var r = focusUiaRect.Value;
                    int cx = (r.Left + r.Right) / 2;
                    int cy = (r.Top + r.Bottom) / 2;
                    meta.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
                    meta.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
                    AppendFocusedUiaSummary(meta, promptContext.FocusedUiaSummary);
                }
                else
                {
                    meta.AppendLine("FOCUS_UIA: none");
                }
                AppendBlockingPromptHint(meta, promptContext);
        
                var content = new List<object>
                {
                    new { type = "input_text",  text = meta.ToString() },
                    new { type = "input_image", image_url = LogImageRef(fullImagePath) }
                };
                if (IncludeFocusUiaCrop && focusUiaImagePath != null)
                    content.Add(new { type = "input_image", image_url = LogImageRef(focusUiaImagePath) });
        
                var req = new Dictionary<string, object>
                {
                    ["model"] = model,
                    ["input"] = new object[]
                    {
                        new { role = "system", content = new object[] { new { type = "input_text", text = qaRules } } },
                        new { role = "user",   content = content.ToArray() }
                    },
                    ["text"] = TextOptions(new { name = "QAWithLocate" })
                };
                if (SupportsTemperature(model))
                    req["temperature"] = 0.0;
                AddReasoningOptions(req, model, EffectiveQaReasoningEffort(), QaMaxOutputTokens, "qa");
        
                return req;
            }
        
            // === VerifyGoal (Q&A yes/no) + request logs ===
            internal static async Task<VerifyDto?> VerifyGoalAsync(string apiKey, string goal, string imageDataUrl, string? currentShotPath,
                                                          int screenW, int screenH,
                                                          Rectangle? focusUiaRect, UiPromptContext promptContext,
                                                          string requestsDir, string commandId, int step,
                                                          CancellationToken cancellationToken = default)
            {
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
                    .AppendLine($"SCREEN_SIZE: width={screenW}, height={screenH} (px)")
                    .AppendLine($"ACTIVE_WINDOW: {promptContext.ActiveWindowTitle}")
                    .AppendLine($"ACTIVE_PROCESS: {promptContext.ActiveProcessName}");
                if (CurrentScreenMap.IsScaled)
                    userText.AppendLine($"REAL_SCREEN_SIZE: width={CurrentScreenMap.ScreenW}, height={CurrentScreenMap.ScreenH} (px)");
                AppendActiveWindowGeometry(userText, promptContext);
        
                if (focusUiaRect.HasValue)
                {
                    var r = focusUiaRect.Value;
                    int cx = (r.Left + r.Right) / 2;
                    int cy = (r.Top + r.Bottom) / 2;
                    userText.AppendLine($"FOCUS_UIA: left={r.Left}, top={r.Top}, right={r.Right}, bottom={r.Bottom} (px)");
                    userText.AppendLine($"FOCUS_UIA_CENTER: x={cx}, y={cy} (px)");
                    AppendFocusedUiaSummary(userText, promptContext.FocusedUiaSummary);
                }
                else
                {
                    userText.AppendLine("FOCUS_UIA: none");
                }
                AppendBlockingPromptHint(userText, promptContext);
        
                var verifyModel = EffectiveVerifyModel();
                var req = new Dictionary<string, object>
                {
                    ["model"] = verifyModel,
                    ["input"] = new object[]
                    {
                        new { role = "system", content = new object[] { new { type = "input_text", text = rules } } },
                        new { role = "user",   content = new object[] {
                                new { type = "input_text", text = userText.ToString() },
                                new { type = "input_image", image_url = imageDataUrl }
                            }
                        }
                    },
                    ["text"] = TextOptions(format)
                };
                if (SupportsTemperature(verifyModel))
                    req["temperature"] = 0.0;
                AddReasoningOptions(req, verifyModel, EffectiveVerifyReasoningEffort(), VerifyMaxOutputTokens, "verify");
        
                if (LogRequests)
                {
                    SaveJson(Path.Combine(requestsDir, $"{commandId}_{step}_verify_request.json"),
                             new
                             {
                                 request = "verify",
                                 body = new
                                 {
                                     model = verifyModel,
                                     reasoning_effort = SupportsReasoningEffort(verifyModel) ? EffectiveVerifyReasoningEffort() : null,
                                     rules,
                                     meta = userText.ToString(),
                                     screenshot = LogImageRef(currentShotPath)
                                 }
                             });
                }
        
                var (parsed, raw) = await CallOpenAIParsedAsync<VerifyDto>(apiKey, req, cancellationToken);
        
                // log response
                SaveRaw(Path.Combine(requestsDir, $"{commandId}_{step}_verify_response.json"), raw);
        
                return parsed;
            }
        
            internal static bool SupportsTemperature(string modelName) =>
                !(modelName?.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase) ?? false);
        
            internal static bool SupportsReasoningEffort(string modelName) =>
                (modelName?.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase) ?? false)
                || (modelName?.StartsWith("o", StringComparison.OrdinalIgnoreCase) ?? false);
        
            internal static object TextOptions(object format) => new
            {
                format,
                verbosity = TextVerbosity
            };
        
            internal static void AddReasoningOptions(Dictionary<string, object> req, string model, string? effortOverride = null, int? maxOutputTokensOverride = null, string cacheScope = "control")
            {
                var maxTokens = maxOutputTokensOverride ?? MaxOutputTokens;
                if (maxTokens > 0)
                    req["max_output_tokens"] = maxTokens;
        
                if (UsePromptCache && !string.IsNullOrWhiteSpace(PromptCacheKey))
                {
                    req["prompt_cache_key"] = EffectivePromptCacheKey(cacheScope);
                    req["prompt_cache_retention"] = "24h";
                }
        
                var effort = effortOverride ?? RequestReasoningEffortOverride ?? ReasoningEffort;
                if (string.IsNullOrWhiteSpace(effort) || !SupportsReasoningEffort(model))
                    return;
        
                req["reasoning"] = new { effort };
            }
        
            internal static string EffectivePromptCacheKey(string scope)
            {
                var baseKey = PromptCacheKey?.Trim();
                if (string.IsNullOrWhiteSpace(baseKey))
                    baseKey = "rdpilot";
                scope = string.IsNullOrWhiteSpace(scope) ? "control" : scope.Trim().ToLowerInvariant();
                return $"{baseKey}:{scope}";
            }
    }
}

