/// <summary>
/// Applies a consistent color palette to console output while keeping redirected and file logs plain.
/// </summary>
internal static class ConsoleTheme
{
    private static readonly object SyncRoot = new();
    private static TextWriter? plainOutput;

    public static void Enable()
    {
        plainOutput = Console.Out;
        Console.SetOut(new ColorizingTextWriter(plainOutput));
        Console.SetError(new ColorizingTextWriter(Console.Error, ConsoleColor.Red));
    }

    public static void WriteStartupBanner(
        string model,
        string profile,
        string reasoning,
        string qaModel,
        string verifyModel)
    {
        WriteLine("╭─ RDPilot ─────────────────────────────────────────", ConsoleColor.Cyan);
        WriteLine($"│ Model      {model}", ConsoleColor.White);
        WriteLine($"│ Profile    {profile} · reasoning {reasoning}", ConsoleColor.Gray);
        WriteLine($"│ Helpers    Q&A {qaModel} · verify {verifyModel}", ConsoleColor.DarkGray);
        WriteLine("│ Abort      Ctrl+Alt+Q", ConsoleColor.Yellow);
        WriteLine("╰────────────────────────────────────────────────────", ConsoleColor.Cyan);
    }

    private static void WriteLine(string value, ConsoleColor color)
    {
        lock (SyncRoot)
        {
            var output = plainOutput ?? Console.Out;
            if (Console.IsOutputRedirected)
            {
                output.WriteLine(value);
                return;
            }

            var originalColor = Console.ForegroundColor;
            try
            {
                Console.ForegroundColor = color;
                output.WriteLine(value);
            }
            finally
            {
                Console.ForegroundColor = originalColor;
            }
        }
    }

    private sealed class ColorizingTextWriter(TextWriter inner, ConsoleColor? forcedColor = null) : TextWriter
    {
        public override Encoding Encoding => inner.Encoding;

        public override void Flush() => inner.Flush();

        public override void Write(char value) => WriteWithColor(value.ToString(), lineTerminated: false);

        public override void Write(string? value) => WriteWithColor(value, lineTerminated: false);

        public override void WriteLine() => inner.WriteLine();

        public override void WriteLine(string? value) => WriteWithColor(value, lineTerminated: true);

        private void WriteWithColor(string? value, bool lineTerminated)
        {
            lock (SyncRoot)
            {
                if (Console.IsOutputRedirected)
                {
                    WritePlain(value, lineTerminated);
                    return;
                }

                var originalColor = Console.ForegroundColor;
                try
                {
                    Console.ForegroundColor = forcedColor ?? SelectColor(value);
                    WritePlain(value, lineTerminated);
                }
                finally
                {
                    Console.ForegroundColor = originalColor;
                }
            }
        }

        private void WritePlain(string? value, bool lineTerminated)
        {
            if (lineTerminated)
                inner.WriteLine(value);
            else
                inner.Write(value);
        }

        private static ConsoleColor SelectColor(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return ConsoleColor.Gray;

            var text = value.TrimStart();

            if (IsError(text))
                return ConsoleColor.Red;
            if (IsSuccess(text))
                return ConsoleColor.Green;
            if (text.StartsWith("[metrics]", StringComparison.Ordinal))
                return ConsoleColor.Gray;
            if (text.StartsWith("[guard]", StringComparison.Ordinal) ||
                text.StartsWith("[sanity]", StringComparison.Ordinal) ||
                text.Contains("warning", StringComparison.OrdinalIgnoreCase))
                return ConsoleColor.Yellow;
            if (text.StartsWith("[openai]", StringComparison.Ordinal) ||
                text.StartsWith("OpenAI", StringComparison.Ordinal))
                return ConsoleColor.Cyan;
            if (text.StartsWith("[shot]", StringComparison.Ordinal) ||
                text.StartsWith("[coords]", StringComparison.Ordinal) ||
                text.StartsWith("[crop]", StringComparison.Ordinal) ||
                text.StartsWith("[aim", StringComparison.Ordinal) ||
                text.StartsWith("[focus_uia]", StringComparison.Ordinal))
                return ConsoleColor.DarkCyan;
            if (text.StartsWith("[wait]", StringComparison.Ordinal) ||
                text.StartsWith("[settle]", StringComparison.Ordinal) ||
                text.StartsWith("[analysis]", StringComparison.Ordinal))
                return ConsoleColor.DarkGray;
            if (text.StartsWith("[Q&A]", StringComparison.Ordinal) ||
                text.StartsWith("Question:", StringComparison.Ordinal) ||
                text.StartsWith("Goal:", StringComparison.Ordinal) ||
                text.StartsWith("Command ID:", StringComparison.Ordinal))
                return ConsoleColor.Magenta;
            if (text.StartsWith("Command/Question:", StringComparison.Ordinal) ||
                text.StartsWith("Enter OPENAI_API_KEY:", StringComparison.Ordinal) ||
                text.StartsWith("Emergency abort", StringComparison.Ordinal))
                return ConsoleColor.Yellow;
            if (text.StartsWith('[') && text.Length > 1 && char.IsDigit(text[1]))
                return ConsoleColor.White;

            return ConsoleColor.Gray;
        }

        private static bool IsError(string text) =>
            text.Contains("error", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("failed", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("aborted", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("missing", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("NOT confirmed", StringComparison.Ordinal) ||
            text.Contains('❌');

        private static bool IsSuccess(string text) =>
            text.Contains('✅') ||
            text.Contains("🟩", StringComparison.Ordinal) ||
            text.StartsWith("Finished", StringComparison.Ordinal) ||
            text.StartsWith("Answer:", StringComparison.Ordinal) ||
            text.Contains("Goal confirmed", StringComparison.OrdinalIgnoreCase) ||
            text.StartsWith("[openai] ok", StringComparison.Ordinal);
    }
}
