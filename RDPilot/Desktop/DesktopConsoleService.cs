
namespace RDPilot.Desktop;

    /// <summary>
    /// Controls the console window and the emergency cancellation hotkey.
    /// </summary>
    internal static class DesktopConsoleService
    {
            // --- HOTKEY: Ctrl+Alt+Q (works when console has focus) ---
            [DllImport("user32.dll")] internal static extern short GetAsyncKeyState(int vKey);
            internal static bool CancelRequested
            {
                get => CurrentRun.CancelRequested;
                set => CurrentRun.CancelRequested = value;
            }
            internal static bool IsPressed(int vk) => (GetAsyncKeyState(vk) & 0x8000) != 0;
        
            internal static CancellationMonitor StartCancelHotkeyListener() =>
                new(CurrentRun, () => IsPressed(0x11) && IsPressed(0x12) && IsPressed(0x51));

            // --- DPI awareness ---
            [DllImport("user32.dll")] internal static extern bool SetProcessDpiAwarenessContext(nint value);
        
            // --- Console always-on-top (Win32) ---
            [DllImport("kernel32.dll")] internal static extern IntPtr GetConsoleWindow();
            [DllImport("user32.dll")] internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
            [DllImport("user32.dll")] internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            internal static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
            internal static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
            const uint SWP_NOSIZE = 0x0001;
            const uint SWP_NOMOVE = 0x0002;
            const uint SWP_NOACTIVATE = 0x0010;
            const uint SWP_SHOWWINDOW = 0x0040;
            const int SW_HIDE = 0;
            const int SW_SHOW = 5;
            const int SW_MINIMIZE = 6;
            const int SW_RESTORE = 9;
        
            internal static bool ConcealConsoleWindow()
            {
                return Options.MinimizeConsoleDuringRun
                    ? MinimizeConsoleWindow()
                    : HideConsoleWindow();
            }
        
            internal static bool HideConsoleWindow()
            {
                var hWnd = GetConsoleWindow();
                return hWnd != IntPtr.Zero && ShowWindow(hWnd, SW_HIDE);
            }
        
            internal static bool MinimizeConsoleWindow()
            {
                var hWnd = GetConsoleWindow();
                return hWnd != IntPtr.Zero && ShowWindow(hWnd, SW_MINIMIZE);
            }
        
            internal static void RestoreConsoleWindow()
            {
                var hWnd = GetConsoleWindow();
                if (hWnd != IntPtr.Zero)
                {
                    ShowWindow(hWnd, SW_SHOW);
                    ShowWindow(hWnd, SW_RESTORE);
                }
            }
    }
