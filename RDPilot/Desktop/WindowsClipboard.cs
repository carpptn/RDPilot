namespace RDPilot.Desktop;

/// <summary>Owns a hidden clipboard window and transfers memory ownership only after a successful write.</summary>
internal static class WindowsClipboard
{
    internal interface INative
    {
        nint CreateOwner();
        void DestroyOwner(nint owner);
        bool Open(nint owner);
        bool Empty();
        bool SetUnicodeText(nint memory);
        void Close();
        nint Allocate(int bytes);
        nint Lock(nint memory);
        void Unlock(nint memory);
        void Free(nint memory);
        int LastError { get; }
    }

    internal static void SetText(string text) => SetText(text, new Native());

    internal static void SetText(string text, INative native)
    {
        if (CancelRequested) throw new OperationCanceledException("Clipboard write cancelled.");
        // The owner must be non-null when EmptyClipboard precedes SetClipboardData.
        // This message-only window is created and destroyed synchronously on the same thread.
        var owner = native.CreateOwner();
        if (owner == 0) throw Failure("Create clipboard owner window");
        nint memory = 0;
        try
        {
            var bytes = Encoding.Unicode.GetBytes(text + "\0");
            memory = native.Allocate(bytes.Length);
            if (memory == 0) throw Failure("GlobalAlloc");
            var pointer = native.Lock(memory);
            if (pointer == 0) throw Failure("GlobalLock");
            try { Marshal.Copy(bytes, 0, pointer, bytes.Length); }
            finally { native.Unlock(memory); }

            var opened = false;
            for (var attempt = 0; attempt < 6; attempt++)
            {
                if (CancelRequested) throw new OperationCanceledException("Clipboard write cancelled.");
                if (native.Open(owner)) { opened = true; break; }
                if (attempt < 5) RuntimeClock.Sleep(50);
            }
            if (!opened) throw Failure("OpenClipboard");
            try
            {
                if (CancelRequested) throw new OperationCanceledException("Clipboard write cancelled.");
                if (!native.Empty()) throw Failure("EmptyClipboard");
                if (!native.SetUnicodeText(memory)) throw Failure("SetClipboardData");
                memory = 0; // Windows now owns the non-delayed clipboard data.
            }
            finally { native.Close(); }
        }
        finally
        {
            if (memory != 0) native.Free(memory);
            native.DestroyOwner(owner);
        }

        InvalidOperationException Failure(string operation) =>
            new($"{operation} failed (Win32 error {native.LastError}).");
    }

    sealed class Native : INative
    {
        public nint CreateOwner() => CreateWindowEx(0, "STATIC", "RDPilot Clipboard", 0,
            0, 0, 0, 0, new nint(-3), 0, 0, 0); // HWND_MESSAGE; never visible or activated.
        public void DestroyOwner(nint owner) => DestroyWindow(owner);
        public bool Open(nint owner) => OpenClipboard(owner);
        public bool Empty() => EmptyClipboard();
        public bool SetUnicodeText(nint memory) => SetClipboardData(13, memory) != 0;
        public void Close() => CloseClipboard();
        public nint Allocate(int bytes) => GlobalAlloc(2, (nuint)bytes);
        public nint Lock(nint memory) => GlobalLock(memory);
        public void Unlock(nint memory) => GlobalUnlock(memory);
        public void Free(nint memory) => GlobalFree(memory);
        public int LastError => Marshal.GetLastWin32Error();

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern nint CreateWindowEx(uint extendedStyle, string className, string windowName, uint style,
            int x, int y, int width, int height, nint parent, nint menu, nint instance, nint parameter);
        [DllImport("user32.dll", SetLastError = true)] static extern bool DestroyWindow(nint window);
        [DllImport("user32.dll", SetLastError = true)] static extern bool OpenClipboard(nint owner);
        [DllImport("user32.dll", SetLastError = true)] static extern bool EmptyClipboard();
        [DllImport("user32.dll", SetLastError = true)] static extern nint SetClipboardData(uint format, nint memory);
        [DllImport("user32.dll", SetLastError = true)] static extern bool CloseClipboard();
        [DllImport("kernel32.dll", SetLastError = true)] static extern nint GlobalAlloc(uint flags, nuint bytes);
        [DllImport("kernel32.dll", SetLastError = true)] static extern nint GlobalLock(nint memory);
        [DllImport("kernel32.dll", SetLastError = true)] static extern bool GlobalUnlock(nint memory);
        [DllImport("kernel32.dll", SetLastError = true)] static extern nint GlobalFree(nint memory);
    }
}
