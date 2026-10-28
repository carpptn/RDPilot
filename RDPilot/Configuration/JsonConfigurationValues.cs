namespace RDPilot.Configuration;

/// <summary>Checks JSON types before handing scalar values to the shared configuration validators.</summary>
internal sealed class JsonConfigurationValues(JsonElement root, string source)
{
    internal bool TryNumber(string name, out string value)
    {
        value = "";
        if (!root.TryGetProperty(name, out var element)) return false;
        if (element.ValueKind != JsonValueKind.Number) return Invalid(name, "a number");
        value = element.GetRawText();
        return true;
    }

    internal bool TryBool(string name, out bool value)
    {
        value = false;
        if (!root.TryGetProperty(name, out var element)) return false;
        if (element.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            return Invalid(name, "a boolean");
        value = element.GetBoolean();
        return true;
    }

    internal bool TryString(string name, out string value)
    {
        value = "";
        if (!root.TryGetProperty(name, out var element)) return false;
        if (element.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(element.GetString()))
            return Invalid(name, "a non-empty string");
        value = element.GetString()!;
        return true;
    }

    internal bool TryOptionalString(string name, out string? value)
    {
        value = null;
        if (!root.TryGetProperty(name, out var element)) return false;
        if (element.ValueKind == JsonValueKind.Null) return true;
        if (!TryString(name, out var text)) return false;
        value = text;
        return true;
    }

    internal bool TryEffort(string name, out string value)
    {
        value = "default";
        if (!root.TryGetProperty(name, out var element)) return false;
        return element.ValueKind == JsonValueKind.Null || TryString(name, out value);
    }

    bool Invalid(string name, string expected)
    {
        ConfigurationService.ReportConfigurationError($"Invalid {name} from {source}; expected {expected}.");
        return false;
    }
}
