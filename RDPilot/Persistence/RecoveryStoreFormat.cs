namespace RDPilot.Persistence;

/// <summary>Reads format metadata before interpreting a version-specific payload.</summary>
internal static class RecoveryStoreFormat
{
    internal static int ReadVersion(string json)
    {
        using var document = JsonDocument.Parse(json);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("Recovery store must be an object.");
        var version = 1;
        foreach (var property in document.RootElement.EnumerateObject())
        {
            if (!property.Name.Equals("version", StringComparison.OrdinalIgnoreCase)) continue;
            if (property.Value.ValueKind != JsonValueKind.Number || !property.Value.TryGetInt32(out var value))
                throw new InvalidDataException("Recovery store version must be an integer.");
            // A duplicate field must never hide a future version from the write guard.
            version = Math.Max(version, value);
        }
        return version;
    }
}