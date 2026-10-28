namespace RDPilot.OpenAI;

/// <summary>One failure policy for JSON responses, SSE final responses and HTTP errors.</summary>
internal readonly record struct ApiResponseOutcome(bool Accepted, bool Retriable, string FailureKind)
{
    internal static ApiResponseOutcome Classify(int statusCode, string raw)
    {
        if (statusCode is < 200 or >= 300)
            return new(false, IsRetriableStatus(statusCode), $"http_{statusCode}");
        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return new(false, false, "invalid_response");
            var status = root.TryGetProperty("status", out var state) && state.ValueKind == JsonValueKind.String
                ? state.GetString() : null;
            if (status == "failed" || root.TryGetProperty("error", out var error) && error.ValueKind != JsonValueKind.Null)
            {
                var code = root.TryGetProperty("error", out error) && error.ValueKind == JsonValueKind.Object &&
                    error.TryGetProperty("code", out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString() ?? "response_failed" : "response_failed";
                return new(false, code is "server_error" or "rate_limit_exceeded", code);
            }
            if (status == "cancelled") return new(false, false, "response_cancelled");
            // Incomplete responses remain available to the bounded output-budget retry policy.
            return status is "completed" or "incomplete"
                ? new(true, false, "") : new(false, false, "invalid_response_status");
        }
        catch (JsonException) { return new(false, false, "invalid_response"); }
    }
}
