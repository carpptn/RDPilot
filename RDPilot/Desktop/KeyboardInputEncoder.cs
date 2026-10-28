namespace RDPilot.Desktop;

/// <summary>Encodes physical shortcuts separately from Unicode text input.</summary>
internal static class KeyboardInputEncoder
{
    internal readonly record struct CharacterKey(ushort VirtualKey, byte Modifiers);
    internal static Func<char, CharacterKey?> CharacterMapper { get; set; } = MapForegroundCharacter;

    internal static INPUT[] Chord(string[] modifiers, string main)
    {
        var codes = modifiers.Select(KeyNameToVk).Distinct().ToList();
        ushort key;
        if (main.Length == 1 && !char.IsAsciiLetterOrDigit(main[0]) && main[0] != ' ')
        {
            var mapped = CharacterMapper(main[0]);
            if (mapped is not { } value || value.VirtualKey is 0 or > 254 || (value.Modifiers & ~7) != 0)
                throw new ArgumentException($"No physical shortcut key for '{main}' in the foreground keyboard layout.");
            key = value.VirtualKey;
            foreach (var (flag, code) in new (int, ushort)[] { (1, 0x10), (2, 0x11), (4, 0x12) })
                if ((value.Modifiers & flag) != 0 && !codes.Contains(code)) codes.Add(code);
        }
        else key = KeyNameToVk(main);

        var inputs = codes.Select(code => VirtualKey(code, false)).ToList();
        inputs.Add(VirtualKey(key, false));
        inputs.Add(VirtualKey(key, true));
        inputs.AddRange(codes.AsEnumerable().Reverse().Select(code => VirtualKey(code, true)));
        return inputs.ToArray();
    }

    internal static INPUT VirtualKey(ushort key, bool keyUp) => new()
    {
        type = 1,
        U = new InputUnion { ki = new KEYBDINPUT { wVk = key, dwFlags = keyUp ? 2u : 0u } }
    };

    static CharacterKey? MapForegroundCharacter(char value)
    {
        var window = GetForegroundWindow();
        var thread = window == 0 ? 0 : GetWindowThreadProcessId(window, out _);
        var encoded = VkKeyScanExW(value, GetKeyboardLayout(thread));
        return encoded == -1 ? null : new((ushort)(encoded & 255), (byte)((encoded >> 8) & 255));
    }

    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(nint window, out uint process);
    [DllImport("user32.dll")] static extern nint GetKeyboardLayout(uint thread);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    static extern short VkKeyScanExW(char value, nint layout);
}
