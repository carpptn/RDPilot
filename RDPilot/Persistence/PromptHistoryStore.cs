using System.Runtime.CompilerServices;

namespace RDPilot.Persistence;

/// <summary>Validates and merges history under a path lock; console snapshots are never authoritative writes.</summary>
internal static class PromptHistoryStore
{
    const int HistoryVersion = 1;
    const int MaxEntries = 500;
    static readonly ConditionalWeakTable<List<string>, Dictionary<string, List<string>>> Pending = new();

    sealed class Document
    {
        public int Version { get; set; } = HistoryVersion;
        public List<string> Prompts { get; set; } = [];
    }

    internal static List<string> Load(string path)
    {
        try
        {
            path = Path.GetFullPath(path);
            using var lease = StorageFileLock.Acquire(path);
            return ReadExisting(path) ?? [];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[history] could not load {path}: {ex.Message}");
            // The console remains usable. Every write independently validates the file again.
            return [];
        }
    }

    internal static void Remember(List<string> entries, string prompt, string path)
    {
        var normalized = NormalizePrompt(prompt);
        if (normalized.Length == 0) return;
        var local = NormalizeEntries(entries.Append(normalized));
        entries.Clear();
        entries.AddRange(local);
        try
        {
            path = Path.GetFullPath(path);
            var byPath = Pending.GetValue(entries, _ => new(StringComparer.OrdinalIgnoreCase));
            if (!byPath.TryGetValue(path, out var pending)) byPath[path] = pending = [];
            pending.RemoveAll(value => value == normalized);
            pending.Add(normalized);
            if (pending.Count > MaxEntries) pending.RemoveAt(0);
            if (TryAppend(pending, path, out var saved))
            {
                entries.Clear();
                entries.AddRange(saved);
                byPath.Remove(path);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[history] could not remember prompt in {path}: {ex.Message}");
        }
    }

    internal static bool TryAppend(IReadOnlyCollection<string> prompts, string path, out List<string> saved)
    {
        saved = [];
        try
        {
            path = Path.GetFullPath(path);
            using var lease = StorageFileLock.Acquire(path);
            // Only genuinely missing files may start empty. Invalid/unreadable files throw.
            var existing = ReadExisting(path) ?? [];
            saved = NormalizeEntries(existing.Concat(prompts));
            Write(path, saved);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[history] could not save {path}: {ex.Message}");
            return false;
        }
    }

    internal static bool TryMigrate(string legacyPath, string destinationPath)
    {
        try
        {
            legacyPath = Path.GetFullPath(legacyPath);
            destinationPath = Path.GetFullPath(destinationPath);
            if (string.Equals(legacyPath, destinationPath, StringComparison.OrdinalIgnoreCase)) return false;
            // Stable lock ordering also protects concurrent migration and ordinary writers.
            var paths = new[] { legacyPath, destinationPath }.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
            using var first = StorageFileLock.Acquire(paths[0]);
            using var second = StorageFileLock.Acquire(paths[1]);
            if (ReadExisting(destinationPath) is not null) return false;
            var entries = ReadExisting(legacyPath);
            if (entries is null) return false;
            Write(destinationPath, entries);
            File.Delete(legacyPath);
            try
            {
                var directory = Path.GetDirectoryName(legacyPath)!;
                if (!Directory.EnumerateFileSystemEntries(directory).Any()) Directory.Delete(directory);
            }
            catch (Exception ex) { Console.WriteLine($"[history] could not remove legacy directory: {ex.Message}"); }
            Console.WriteLine($"[history] migrated {entries.Count} prompt(s) to {destinationPath}");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[history] could not migrate {legacyPath}: {ex.Message}");
            return false;
        }
    }

    static List<string>? ReadExisting(string path)
    {
        string json;
        try { json = File.ReadAllText(path); }
        catch (FileNotFoundException) { return null; }
        catch (DirectoryNotFoundException) { return null; }
        using var document = JsonDocument.Parse(json);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("Prompt history must be an object.");
        var properties = document.RootElement.EnumerateObject().ToArray();
        var versions = properties.Where(p => p.Name.Equals("version", StringComparison.OrdinalIgnoreCase)).ToArray();
        var prompts = properties.Where(p => p.Name.Equals("prompts", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (versions.Length != 1 || versions[0].Value.ValueKind != JsonValueKind.Number ||
            !versions[0].Value.TryGetInt32(out var version) || version != HistoryVersion)
            throw new InvalidDataException("Unsupported or ambiguous prompt history version; preserving the file.");
        if (prompts.Length != 1 || prompts[0].Value.ValueKind != JsonValueKind.Array ||
            prompts[0].Value.EnumerateArray().Any(value => value.ValueKind != JsonValueKind.String) ||
            properties.Any(p => p.Name.Equals("lessons", StringComparison.OrdinalIgnoreCase) ||
                                p.Name.Equals("cases", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("Invalid prompt history structure; preserving the file.");
        return NormalizeEntries(prompts[0].Value.EnumerateArray().Select(value => value.GetString()!));
    }

    static void Write(string path, List<string> prompts) =>
        AtomicTextStore.Write(path, JsonSerializer.Serialize(new Document { Prompts = prompts }, PrettyJson));

    internal static List<string> NormalizeEntries(IEnumerable<string> entries)
    {
        var result = new List<string>();
        foreach (var entry in entries)
        {
            var normalized = NormalizePrompt(entry);
            if (normalized.Length == 0) continue;
            result.RemoveAll(existing => string.Equals(existing, normalized, StringComparison.Ordinal));
            result.Add(normalized);
        }
        if (result.Count > MaxEntries) result.RemoveRange(0, result.Count - MaxEntries);
        return result;
    }

    internal static string NormalizePrompt(string prompt) => prompt
        .Replace("\r\n", " ", StringComparison.Ordinal).Replace('\r', ' ').Replace('\n', ' ').Trim();
}