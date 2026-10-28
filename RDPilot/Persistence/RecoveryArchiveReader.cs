namespace RDPilot.Persistence;

/// <summary>Rejects unrelated, ambiguous and unsupported archive documents before any mutation.</summary>
internal static class RecoveryArchiveReader
{
    internal static RecoveryLessonArchiveStore ReadExisting(string path)
    {
        string json;
        try { json = File.ReadAllText(path); }
        catch (FileNotFoundException) { return new(); }
        catch (DirectoryNotFoundException) { return new(); }

        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("Recovery archive must be an object.");
        var properties = root.EnumerateObject().ToArray();
        var versions = properties.Where(p => p.Name.Equals("version", StringComparison.OrdinalIgnoreCase)).ToArray();
        var lessons = properties.Where(p => p.Name.Equals("lessons", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (versions.Length != 1 || versions[0].Value.ValueKind != JsonValueKind.Number ||
            !versions[0].Value.TryGetInt32(out var version) || version != 1)
            throw new InvalidDataException("Recovery archive requires supported version 1; preserving the existing file.");
        if (lessons.Length != 1 || lessons[0].Value.ValueKind != JsonValueKind.Array ||
            lessons[0].Value.EnumerateArray().Any(item => item.ValueKind != JsonValueKind.Object) ||
            properties.Any(p => p.Name.Equals("cases", StringComparison.OrdinalIgnoreCase) ||
                                p.Name.Equals("prompts", StringComparison.OrdinalIgnoreCase) ||
                                p.Name.Equals("calibration", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException("The existing document is not a recovery archive; refusing to replace it.");
        return JsonSerializer.Deserialize<RecoveryLessonArchiveStore>(json, PrettyJson)
            ?? throw new InvalidDataException("Recovery archive is null.");
    }
}