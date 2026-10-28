namespace RDPilot.Desktop;

/// <summary>Tracks accepted native events across calls, including interrupted drags and chords.</summary>
internal static class PressedInputState
{
    private static readonly object Gate = new();
    private static readonly Dictionary<(ushort Vk, ushort Scan, uint Flags), INPUT> Keys = [];
    private static uint mouseDown;

    internal static void RecordAccepted(ReadOnlySpan<INPUT> events)
    {
        lock (Gate)
        {
            foreach (var input in events)
            {
                if (input.type == 1)
                {
                    var key = (input.U.ki.wVk, input.U.ki.wScan, input.U.ki.dwFlags & ~2u);
                    if ((input.U.ki.dwFlags & 2) != 0) Keys.Remove(key);
                    else Keys[key] = input;
                }
                else if (input.type == 0)
                {
                    foreach (var down in new uint[] { 2, 8, 32 })
                    {
                        if ((input.U.mi.dwFlags & down) != 0) mouseDown |= down;
                        if ((input.U.mi.dwFlags & (down << 1)) != 0) mouseDown &= ~down;
                    }
                }
            }
        }
    }

    internal static INPUT[] Releases()
    {
        lock (Gate)
        {
            var releases = Keys.Values.Reverse().Select(input =>
            {
                input.U.ki.dwFlags |= 2;
                return input;
            }).ToList();
            if (mouseDown != 0)
                releases.Add(new INPUT { type = 0, U = new InputUnion { mi = new MOUSEINPUT { dwFlags = mouseDown << 1 } } });
            return releases.ToArray();
        }
    }
}
