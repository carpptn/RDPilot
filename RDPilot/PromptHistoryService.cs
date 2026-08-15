using System.Text;
using System.Text.Json;

internal static partial class RDPilotApplication
{
    internal static class PromptHistoryService
    {
        const int HistoryVersion = 1;
        const int MaxEntries = 500;

        sealed class PromptHistoryDocument
        {
            public int Version { get; set; } = HistoryVersion;
            public List<string> Prompts { get; set; } = [];
        }

        internal sealed class NavigationState(IReadOnlyList<string> entries)
        {
            int index = entries.Count;
            string draft = "";

            internal string Up(string current)
            {
                if (entries.Count == 0)
                    return current;
                if (index == entries.Count)
                    draft = current;
                if (index > 0)
                    index--;
                return entries[index];
            }

            internal string Down(string current)
            {
                if (entries.Count == 0 || index == entries.Count)
                    return current;
                index++;
                return index == entries.Count ? draft : entries[index];
            }
        }

        internal static string EffectivePath() =>
            Path.Combine(
                AppContext.BaseDirectory,
                "memory",
                "prompt-history.json");

        internal static string LegacyPath() =>
            Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "RDPilot",
                "prompt-history.json");

        internal static List<string> Load(string? path = null)
        {
            if (path is null)
            {
                path = EffectivePath();
                if (!File.Exists(path))
                    TryMigrateLegacyHistory(LegacyPath(), path);
            }
            if (!File.Exists(path))
                return [];

            try
            {
                var document = JsonSerializer.Deserialize<PromptHistoryDocument>(
                    File.ReadAllText(path),
                    PrettyJson);
                return NormalizeEntries(document?.Prompts ?? []);
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[history] could not load {path}: {ex.Message}");
                return [];
            }
        }

        internal static void Remember(
            List<string> entries,
            string prompt,
            string? path = null)
        {
            var normalized = NormalizePrompt(prompt);
            if (normalized.Length == 0)
                return;

            entries.RemoveAll(entry =>
                string.Equals(entry, normalized, StringComparison.Ordinal));
            entries.Add(normalized);
            if (entries.Count > MaxEntries)
            {
                entries.RemoveRange(0, entries.Count - MaxEntries);
            }
            Save(entries, path);
        }

        internal static bool Save(
            IReadOnlyCollection<string> entries,
            string? path = null)
        {
            path ??= EffectivePath();
            try
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);
                var document = new PromptHistoryDocument
                {
                    Prompts = NormalizeEntries(entries)
                };
                var json = JsonSerializer.Serialize(document, PrettyJson)
                    .Replace("\r\n", "\n", StringComparison.Ordinal)
                    .Replace("\n", "\r\n", StringComparison.Ordinal);
                var temporaryPath = path + ".tmp";
                File.WriteAllText(
                    temporaryPath,
                    json,
                    new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                File.Move(temporaryPath, path, overwrite: true);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[history] could not save {path}: {ex.Message}");
                return false;
            }
        }

        internal static bool TryMigrateLegacyHistory(
            string legacyPath,
            string destinationPath)
        {
            if (!File.Exists(legacyPath) || File.Exists(destinationPath))
                return false;

            try
            {
                var document = JsonSerializer.Deserialize<PromptHistoryDocument>(
                    File.ReadAllText(legacyPath),
                    PrettyJson);
                if (document is null)
                    return false;

                var entries = NormalizeEntries(document.Prompts);
                if (!Save(entries, destinationPath) || !File.Exists(destinationPath))
                    return false;

                File.Delete(legacyPath);
                var legacyDirectory = Path.GetDirectoryName(legacyPath);
                if (!string.IsNullOrWhiteSpace(legacyDirectory) &&
                    Directory.Exists(legacyDirectory) &&
                    !Directory.EnumerateFileSystemEntries(legacyDirectory).Any())
                {
                    Directory.Delete(legacyDirectory);
                }

                Console.WriteLine(
                    $"[history] migrated {entries.Count} prompt(s) to {destinationPath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[history] could not migrate {legacyPath}: {ex.Message}");
                return false;
            }
        }

        internal static string ReadLine(
            string prompt,
            IReadOnlyList<string> entries)
        {
            Console.Write(prompt);
            if (Console.IsInputRedirected)
                return Console.ReadLine() ?? "";

            try
            {
                return ReadInteractiveLine(entries);
            }
            catch (Exception ex) when (
                ex is ArgumentOutOfRangeException or IOException or
                    InvalidOperationException)
            {
                Console.WriteLine();
                Console.WriteLine(
                    $"[history] interactive redraw unavailable ({ex.GetType().Name}); falling back to standard input.");
                Console.Write("Command/Question: ");
                return Console.ReadLine() ?? "";
            }
        }

        static string ReadInteractiveLine(IReadOnlyList<string> entries)
        {
            var navigation = new NavigationState(entries);
            var buffer = new StringBuilder();
            var cursor = 0;
            var renderedLength = 0;
            var startLeft = Console.CursorLeft;
            var startTop = Console.CursorTop;

            void SetInput(string value)
            {
                buffer.Clear();
                buffer.Append(value);
                cursor = buffer.Length;
                Render();
            }

            void EnsureRenderCapacity(int characterOffset)
            {
                var width = Math.Max(1, Console.BufferWidth);
                var bufferHeight = Math.Max(1, Console.BufferHeight);
                var requiredRowOffset =
                    (startLeft + Math.Max(0, characterOffset)) / width;
                var overflow = startTop + requiredRowOffset -
                               (bufferHeight - 1);
                if (overflow <= 0)
                    return;

                Console.SetCursorPosition(0, bufferHeight - 1);
                for (var row = 0; row < overflow; row++)
                    Console.WriteLine();
                startTop = Math.Max(0, startTop - overflow);
                if (startTop + requiredRowOffset >= bufferHeight)
                {
                    throw new InvalidOperationException(
                        "the recalled prompt is taller than the console buffer");
                }
            }

            void PositionCursor(int characterOffset)
            {
                EnsureRenderCapacity(characterOffset);
                var width = Math.Max(1, Console.BufferWidth);
                var absoluteOffset = startLeft + characterOffset;
                Console.SetCursorPosition(
                    absoluteOffset % width,
                    startTop + absoluteOffset / width);
            }

            void Render()
            {
                EnsureRenderCapacity(
                    Math.Max(renderedLength, buffer.Length) + 1);
                PositionCursor(0);
                if (renderedLength > 0)
                    Console.Write(new string(' ', renderedLength));
                PositionCursor(0);
                Console.Write(buffer.ToString());
                renderedLength = Math.Max(renderedLength, buffer.Length);
                PositionCursor(cursor);
            }

            while (true)
            {
                var key = Console.ReadKey(intercept: true);
                switch (key.Key)
                {
                    case ConsoleKey.Enter:
                        PositionCursor(buffer.Length);
                        Console.WriteLine();
                        return buffer.ToString();
                    case ConsoleKey.UpArrow:
                        SetInput(navigation.Up(buffer.ToString()));
                        break;
                    case ConsoleKey.DownArrow:
                        SetInput(navigation.Down(buffer.ToString()));
                        break;
                    case ConsoleKey.LeftArrow:
                        if (cursor > 0)
                        {
                            cursor--;
                            PositionCursor(cursor);
                        }
                        break;
                    case ConsoleKey.RightArrow:
                        if (cursor < buffer.Length)
                        {
                            cursor++;
                            PositionCursor(cursor);
                        }
                        break;
                    case ConsoleKey.Home:
                        cursor = 0;
                        PositionCursor(cursor);
                        break;
                    case ConsoleKey.End:
                        cursor = buffer.Length;
                        PositionCursor(cursor);
                        break;
                    case ConsoleKey.Backspace:
                        if (cursor > 0)
                        {
                            buffer.Remove(--cursor, 1);
                            Render();
                        }
                        break;
                    case ConsoleKey.Delete:
                        if (cursor < buffer.Length)
                        {
                            buffer.Remove(cursor, 1);
                            Render();
                        }
                        break;
                    case ConsoleKey.Escape:
                        buffer.Clear();
                        cursor = 0;
                        Render();
                        break;
                    default:
                        if (!char.IsControl(key.KeyChar))
                        {
                            buffer.Insert(cursor, key.KeyChar);
                            cursor++;
                            Render();
                        }
                        break;
                }
            }
        }

        internal static List<string> NormalizeEntries(
            IEnumerable<string> entries)
        {
            var result = new List<string>();
            foreach (var entry in entries)
            {
                var normalized = NormalizePrompt(entry);
                if (normalized.Length == 0)
                    continue;
                result.RemoveAll(existing =>
                    string.Equals(existing, normalized, StringComparison.Ordinal));
                result.Add(normalized);
            }
            if (result.Count > MaxEntries)
                result.RemoveRange(0, result.Count - MaxEntries);
            return result;
        }

        static string NormalizePrompt(string prompt) =>
            prompt
                .Replace("\r\n", " ", StringComparison.Ordinal)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Trim();
    }
}
