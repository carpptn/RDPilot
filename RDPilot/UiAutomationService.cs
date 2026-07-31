internal static partial class RDPilotApplication
{
    /// <summary>
    /// Collects UI Automation context, focused-control data, and actionable desktop targets.
    /// </summary>
    internal static class UiAutomationService
    {
            internal static string Tail(string s, int n)
            {
                if (string.IsNullOrEmpty(s)) return "";
                return s.Length <= n ? s : s[^n..];
            }
        
            internal static string Tail(StringBuilder sb, int n)
            {
                if (sb.Length == 0 || n <= 0) return "";
                if (sb.Length <= n) return sb.ToString();
                return sb.ToString(sb.Length - n, n);
            }
        
            internal static UiPromptContext CaptureUiPromptContext(string? focusedUiaSummary, int screenW, int screenH)
            {
                var title = GetActiveWindowTitleForPrompt();
                var process = GetActiveProcessNameForPrompt();
                var summary = IncludeFocusUia ? focusedUiaSummary : null;
                var (geometry, visibilityHint) = DescribeActiveWindowGeometryForPrompt(screenW, screenH);
                return new(title, process, summary, DetectBlockingPromptHint(title, process, summary), geometry, visibilityHint);
            }
        
            internal static void AppendFocusedUiaSummary(StringBuilder sb, string? summary)
            {
                if (!string.IsNullOrWhiteSpace(summary))
                    sb.AppendLine($"FOCUS_UIA_INFO: {TrimForMeta(summary, UiaSummaryMaxChars)}");
            }
        
            internal static void AppendActiveWindowGeometry(StringBuilder sb, UiPromptContext promptContext)
            {
                if (!string.IsNullOrWhiteSpace(promptContext.ActiveWindowGeometry))
                    sb.AppendLine($"ACTIVE_WINDOW_GEOMETRY: {promptContext.ActiveWindowGeometry}");
                if (!string.IsNullOrWhiteSpace(promptContext.WindowVisibilityHint))
                    sb.AppendLine($"WINDOW_VISIBILITY_HINT: {promptContext.WindowVisibilityHint}");
            }
        
            internal static (string? Geometry, string? Hint) DescribeActiveWindowGeometryForPrompt(int screenW, int screenH)
            {
                try
                {
                    var hWnd = GetForegroundWindow();
                    if (hWnd == IntPtr.Zero || !GetWindowRect(hWnd, out var wr))
                        return (null, null);
        
                    var width = wr.Right - wr.Left;
                    var height = wr.Bottom - wr.Top;
                    if (width <= 0 || height <= 0)
                        return (null, null);
        
                    var (screenX, screenY, controlledW, controlledH) =
                        GetPrimaryScreen();
                    if (screenW <= 0 || screenH <= 0)
                    {
                        screenW = controlledW;
                        screenH = controlledH;
                    }
        
                    var screen = new Rectangle(
                        screenX,
                        screenY,
                        Math.Max(1, screenW),
                        Math.Max(1, screenH));
                    var rect = new Rectangle(wr.Left, wr.Top, width, height);
                    var visible = Rectangle.Intersect(rect, screen);
                    var area = Math.Max(1L, (long)width * height);
                    var visibleArea = (long)Math.Max(0, visible.Width) * Math.Max(0, visible.Height);
                    var screenArea = Math.Max(1L, (long)screen.Width * screen.Height);
                    var visibleRatio = visibleArea / (double)area;
                    var screenAreaRatio = visibleArea / (double)screenArea;
        
                    var clippedX = Math.Max(0, screen.Left - rect.Left) + Math.Max(0, rect.Right - screen.Right);
                    var clippedY = Math.Max(0, screen.Top - rect.Top) + Math.Max(0, rect.Bottom - screen.Bottom);
                    var clipped = visibleArea == 0 ||
                                  visibleRatio < 0.85 ||
                                  clippedX > Math.Max(80, width * 0.12) ||
                                  clippedY > Math.Max(80, height * 0.12);
                    var small = visibleArea > 0 &&
                                ((visible.Width < screen.Width * 0.55 && visible.Height < screen.Height * 0.55) ||
                                 screenAreaRatio < 0.25);
        
                    var state = visibleArea == 0
                        ? "offscreen"
                        : clipped && small
                            ? "small+clipped"
                            : clipped
                                ? "clipped"
                                : small
                                    ? "small"
                                    : "normal";
        
                    var promptRect = CurrentScreenMap.ScreenToImageRect(rect);
                    var promptVisible = visibleArea > 0
                        ? CurrentScreenMap.ScreenToImageRect(visible)
                        : Rectangle.Empty;
        
                    var visibleText = visibleArea > 0
                        ? FormattableString.Invariant($"({promptVisible.Left},{promptVisible.Top})-({promptVisible.Right},{promptVisible.Bottom}) size={promptVisible.Width}x{promptVisible.Height}")
                        : "none";
                    var geometry = FormattableString.Invariant(
                        $"rect=({promptRect.Left},{promptRect.Top})-({promptRect.Right},{promptRect.Bottom}) size={promptRect.Width}x{promptRect.Height}; visible={visibleText}; visible_pct={visibleRatio * 100:0.#}; screen_area_pct={screenAreaRatio * 100:0.#}; state={state}");
        
                    var hint = clipped || small
                        ? "Active window appears too small, clipped, or off-screen. If this blocks the task, choose a real UI window-management action first, usually keys [\"win\",\"up\"] to maximize."
                        : null;
                    return (geometry, hint);
                }
                catch
                {
                    return (null, null);
                }
            }

            internal static Rectangle? GetActiveWindowRectangle()
            {
                try
                {
                    var hWnd = GetForegroundWindow();
                    if (hWnd == IntPtr.Zero || !GetWindowRect(hWnd, out var rect))
                        return null;
                    var width = rect.Right - rect.Left;
                    var height = rect.Bottom - rect.Top;
                    return width > 0 && height > 0
                        ? new Rectangle(rect.Left, rect.Top, width, height)
                        : null;
                }
                catch
                {
                    return null;
                }
            }
        
            internal static string? DetectBlockingPromptHint(string title, string processName, string? focusedUiaSummary)
            {
                var text = $"{title}\n{processName}\n{focusedUiaSummary}".ToLowerInvariant();
                if (text.Contains("user account control") || text.Contains("kontrola konta użytkownika") || text.Contains("kontrola konta uzytkownika") || processName.Equals("consent", StringComparison.OrdinalIgnoreCase))
                    return "Possible UAC/security prompt is active. Use visible UI only; decide whether the goal requires approving or dismissing it.";
        
                if (text.Contains("permission") || text.Contains("uprawnien") || text.Contains("zezwol") || text.Contains("allow ") || text.Contains("deny ") || text.Contains("blocked"))
                    return "Possible permission prompt is active. Handle the visible prompt explicitly before continuing.";
        
                if (text.Contains("dialog") || text.Contains("modal") || text.Contains("ok button") || text.Contains("cancel button") || text.Contains("anuluj") || text.Contains("zapisz jako"))
                    return "Possible modal dialog is active. Resolve the dialog through its visible buttons or keyboard shortcuts before returning to the main window.";
        
                return null;
            }
        
            internal static void AppendBlockingPromptHint(StringBuilder sb, UiPromptContext promptContext)
            {
                if (!string.IsNullOrWhiteSpace(promptContext.BlockingPromptHint))
                    sb.AppendLine($"BLOCKING_PROMPT_HINT: {promptContext.BlockingPromptHint}");
            }
        
            internal static void AppendUiaTargets(StringBuilder sb, bool reuseExisting = false, string? reuseReason = null)
            {
                if (CurrentUiaTargets.Count == 0)
                    return;
        
                sb.AppendLine(reuseExisting
                    ? $"UIA_TARGETS: (reused; {reuseReason ?? "current list"}; coords in SCREEN_SIZE image space)"
                    : "UIA_TARGETS: (coords in SCREEN_SIZE image space)");
                foreach (var t in CurrentUiaTargets)
                {
                    var rect = CurrentScreenMap.ScreenToImageRect(t.Rect);
                    var center = CurrentScreenMap.ScreenToImagePoint(t.CenterX, t.CenterY);
                    sb.AppendLine($"- #{t.Index}: {t.ControlType} name=\"{TrimForMeta(t.Name, UiaTargetNameMaxChars)}\" bbox=({rect.Left},{rect.Top})-({rect.Right},{rect.Bottom}) center=({center.X},{center.Y})");
                }
            }
        
            internal static void PrepareUiaTargetsForPrompt(bool reuseExisting, int screenW, int screenH)
            {
                if (!IncludeUiaTargets || MaxUiaTargets <= 0 || !MouseEnabled)
                {
                    CurrentUiaTargets = new();
                    return;
                }
        
                if (!reuseExisting || CurrentUiaTargets.Count == 0)
                    CurrentUiaTargets = GetUiaTargets(MaxUiaTargets, screenW, screenH);
            }
        
            internal static List<UiaTarget> GetUiaTargets(int maxTargets, int screenW, int screenH)
            {
                var totalSw = Stopwatch.StartNew();
                var candidates = new List<UiaTargetCandidate>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    var root = GetFocusedWindowElement() ?? AutomationElement.RootElement;
                    var (screenX, screenY, _, _) = GetPrimaryScreen();
                    var screen = new Rectangle(screenX, screenY, screenW, screenH);
                    var walker = TreeWalker.ControlViewWalker;
                    var queue = new Queue<AutomationElement>();
                    var budgetSw = Stopwatch.StartNew();
                    var scanned = 0;
                    var candidateBudget = Math.Clamp(maxTargets * Math.Max(1, UiaCandidateMultiplier), maxTargets, 200);
        
                    EnqueueChildren(root, walker, queue, budgetSw);
        
                    while (queue.Count > 0 && candidates.Count < candidateBudget)
                    {
                        if (MaxUiaNodesScanned > 0 && scanned >= MaxUiaNodesScanned)
                            break;
                        if (UiaScanTimeBudgetMs > 0 && budgetSw.ElapsedMilliseconds >= UiaScanTimeBudgetMs)
                            break;
        
                        var el = queue.Dequeue();
                        scanned++;
        
                        TryAddTarget(el, screen, candidates, seen, candidateBudget);
        
                        if (candidates.Count < candidateBudget)
                            EnqueueChildren(el, walker, queue, budgetSw);
                    }
                }
                catch
                {
                    // UIA can throw when windows close while enumerating. Treat as no targets.
                }
                finally
                {
                    RecordUiaMetric(totalSw);
                }
        
                return candidates
                    .OrderByDescending(c => c.Score)
                    .ThenBy(c => c.Target.Rect.Top)
                    .ThenBy(c => c.Target.Rect.Left)
                    .Take(maxTargets)
                    .Select((c, i) => new UiaTarget(i, c.Target.Name, c.Target.ControlType, c.Target.Rect))
                    .ToList();
        
                static void EnqueueChildren(AutomationElement parent, TreeWalker walker, Queue<AutomationElement> queue, Stopwatch sw)
                {
                    AutomationElement? child;
                    try { child = walker.GetFirstChild(parent); }
                    catch { return; }
        
                    while (child is not null)
                    {
                        if (MaxUiaNodesScanned > 0 && queue.Count >= MaxUiaNodesScanned)
                            break;
                        if (UiaScanTimeBudgetMs > 0 && sw.ElapsedMilliseconds >= UiaScanTimeBudgetMs)
                            break;
        
                        queue.Enqueue(child);
        
                        try { child = walker.GetNextSibling(child); }
                        catch { break; }
                    }
                }
        
                static void TryAddTarget(AutomationElement el, Rectangle screen, List<UiaTargetCandidate> result, HashSet<string> seen, int maxTargets)
                {
                    if (result.Count >= maxTargets)
                        return;
        
                    try
                    {
                        if (!SafeBool(el, AutomationElement.IsControlElementProperty))
                            return;
        
                        var rectObj = el.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
                        if (rectObj is not System.Windows.Rect wr || wr.IsEmpty || wr.Width < 4 || wr.Height < 4)
                            return;
        
                        var rect = Rectangle.FromLTRB((int)Math.Round(wr.Left), (int)Math.Round(wr.Top), (int)Math.Round(wr.Right), (int)Math.Round(wr.Bottom));
                        rect.Intersect(screen);
                        if (rect.Width < 4 || rect.Height < 4)
                            return;
        
                        var name = el.GetCurrentPropertyValue(AutomationElement.NameProperty, true) as string ?? "";
                        var type = (el.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty, true) as ControlType)?.ProgrammaticName?.Replace("ControlType.", "") ?? "control";
                        var focusable = SafeBool(el, AutomationElement.IsKeyboardFocusableProperty);
                        var enabled = SafeBool(el, AutomationElement.IsEnabledProperty);
                        if (!enabled || (!focusable && string.IsNullOrWhiteSpace(name)))
                            return;
        
                        var screenArea = Math.Max(1L, (long)screen.Width * screen.Height);
                        var targetArea = (long)rect.Width * rect.Height;
                        var areaRatio = targetArea / (double)screenArea;
                        if (MaxUiaTargetAreaRatio > 0 &&
                            areaRatio > MaxUiaTargetAreaRatio &&
                            IsLowValueUiaContainer(type, focusable, name))
                        {
                            return;
                        }
        
                        var key = UiaDedupeKey(type, name, rect);
                        if (!seen.Add(key))
                            return;
        
                        var score = ScoreUiaTarget(type, name, focusable, rect, areaRatio);
                        result.Add(new UiaTargetCandidate(new UiaTarget(0, name, type, rect), score));
                    }
                    catch
                    {
                        // Individual UIA elements can become invalid while the tree is being walked.
                    }
                }
        
                static bool SafeBool(AutomationElement el, AutomationProperty prop)
                {
                    try { return el.GetCurrentPropertyValue(prop, true) is bool b && b; }
                    catch { return false; }
                }
            }
        
            internal static string UiaDedupeKey(string type, string name, Rectangle rect)
            {
                var normalizedName = TrimForMeta(name.Trim(), 80);
                return $"{type}|{normalizedName}|{rect.Left},{rect.Top},{rect.Right},{rect.Bottom}";
            }
        
            internal static int ScoreUiaTarget(string type, string name, bool focusable, Rectangle rect, double areaRatio)
            {
                var score = 0;
                if (IsActionableUiaType(type)) score += 60;
                if (focusable) score += 25;
                if (!string.IsNullOrWhiteSpace(name)) score += 15;
                if (rect.Width >= 12 && rect.Height >= 12) score += 5;
                if (areaRatio <= 0.08) score += 10;
                else if (areaRatio >= 0.30) score -= 20;
                if (IsLowValueUiaContainer(type, focusable, name)) score -= 35;
                return score;
            }
        
            internal static bool IsActionableUiaType(string type) => type is
                "Button" or "Edit" or "Hyperlink" or "MenuItem" or "TabItem" or
                "ListItem" or "TreeItem" or "CheckBox" or "RadioButton" or
                "ComboBox" or "Slider" or "Spinner" or "DataItem" or "HeaderItem" or
                "Calendar" or "Thumb";
        
            internal static bool IsLowValueUiaContainer(string type, bool focusable, string name)
            {
                if (IsActionableUiaType(type))
                    return false;
        
                if (focusable && !string.IsNullOrWhiteSpace(name) && type is "Text" or "Document")
                    return false;
        
                return type is
                    "Window" or "Pane" or "Group" or "Custom" or "Document" or
                    "Text" or "Image" or "Separator" or "ToolBar" or "StatusBar" or
                    "TitleBar" or "MenuBar";
            }
        
            internal static AutomationElement? GetFocusedWindowElement()
            {
                try
                {
                    var el = AutomationElement.FocusedElement;
                    if (el is null) return null;
        
                    var walker = TreeWalker.ControlViewWalker;
                    var cur = el;
                    for (var i = 0; i < 12 && cur is not null; i++)
                    {
                        var type = cur.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty, true) as ControlType;
                        if (type == ControlType.Window)
                            return cur;
                        cur = walker.GetParent(cur);
                    }
                }
                catch { }
        
                return null;
            }
        
            internal static FocusedUiaSnapshot CaptureFocusedUiaSnapshot()
            {
                try
                {
                    var el = AutomationElement.FocusedElement;
                    return el is null
                        ? new(null, null)
                        : new(GetFocusedUiaBoundingRect(el), GetFocusedUiaSummary(el));
                }
                catch
                {
                    return new(null, null);
                }
            }
        
            // === UI Automation – bounding rect of the currently focused element ===
            // Returns pixel coords in screen space (VirtualScreen/Primary, referenced to (0,0) of primary).
            internal static Rectangle? GetFocusedUiaBoundingRect()
            {
                try
                {
                    var el = AutomationElement.FocusedElement;
                    if (el is null) return null;
                    return GetFocusedUiaBoundingRect(el);
                }
                catch { /* ignore */ }
                return null;
            }
        
            internal static Rectangle? GetFocusedUiaBoundingRect(AutomationElement el)
            {
                try
                {
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
        
            internal static string? GetFocusedUiaSummary()
            {
                try
                {
                    var el = AutomationElement.FocusedElement;
                    if (el is null) return null;
                    return GetFocusedUiaSummary(el);
                }
                catch
                {
                    return null;
                }
            }
        
            internal static string? GetFocusedUiaSummary(AutomationElement el)
            {
                try
                {
                    var parts = new List<string>();
                    AddProp(parts, "name", el.GetCurrentPropertyValue(AutomationElement.NameProperty, true));
                    AddProp(parts, "automation_id", el.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty, true));
                    AddProp(parts, "control_type", (el.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty, true) as ControlType)?.ProgrammaticName);
                    AddProp(parts, "class", el.GetCurrentPropertyValue(AutomationElement.ClassNameProperty, true));
                    AddProp(parts, "framework", el.GetCurrentPropertyValue(AutomationElement.FrameworkIdProperty, true));
        
                    var parentWindow = FindParentWindowName(el);
                    if (!string.IsNullOrWhiteSpace(parentWindow))
                        parts.Add($"window=\"{TrimForMeta(parentWindow, 80)}\"");
        
                    return parts.Count == 0 ? null : string.Join("; ", parts);
                }
                catch
                {
                    return null;
                }
        
                static void AddProp(List<string> parts, string name, object? value)
                {
                    var s = value as string;
                    if (!string.IsNullOrWhiteSpace(s))
                        parts.Add($"{name}=\"{TrimForMeta(s, 80)}\"");
                }
            }
        
            internal static string? FindParentWindowName(AutomationElement el)
            {
                try
                {
                    var walker = TreeWalker.ControlViewWalker;
                    var cur = el;
                    for (var i = 0; i < 12 && cur is not null; i++)
                    {
                        var type = cur.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty, true) as ControlType;
                        if (type == ControlType.Window)
                        {
                            var name = cur.GetCurrentPropertyValue(AutomationElement.NameProperty, true) as string;
                            return name;
                        }
                        cur = walker.GetParent(cur);
                    }
                }
                catch { }
        
                return null;
            }
        
            internal static string TrimForMeta(string value, int max) =>
                max <= 0 ? "" : value.Length <= max ? value : value[..max];
        
            // P/Invoke GetWindowRect (UIA fallback)
            [DllImport("user32.dll", SetLastError = true)]
            private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
            [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
            [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        
            internal static string GetActiveWindowTitleForPrompt()
            {
                try
                {
                    var hWnd = GetForegroundWindow();
                    if (hWnd == IntPtr.Zero) return "unknown";
                    var sb = new StringBuilder(256);
                    var len = GetWindowText(hWnd, sb, sb.Capacity);
                    return len > 0 ? $"\"{TrimForMeta(sb.ToString(), 120)}\"" : "unknown";
                }
                catch
                {
                    return "unknown";
                }
            }
        
            internal static string GetActiveProcessNameForPrompt()
            {
                try
                {
                    var hWnd = GetForegroundWindow();
                    if (hWnd == IntPtr.Zero) return "unknown";
                    GetWindowThreadProcessId(hWnd, out var pid);
                    if (pid == 0) return "unknown";
                    using var process = Process.GetProcessById((int)pid);
                    return TrimForMeta(process.ProcessName, 80);
                }
                catch
                {
                    return "unknown";
                }
            }
        
            [StructLayout(LayoutKind.Sequential)]
            struct RECT { public int Left, Top, Right, Bottom; }
    }
}





