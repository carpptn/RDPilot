namespace RDPilot.OpenAI;

internal static class OpenAiTransport
{
    internal static HttpClient HttpClient { get; set; } = new() { Timeout = Timeout.InfiniteTimeSpan };

    private sealed record TransportResult(bool Ok, int StatusCode, string Raw, TimeSpan Elapsed,
        int RequestBytes, int ResponseBytes, List<ActionDto>? AcceptedActions, string? ResponseId);
    private sealed record ResponseBody(string Raw, int Bytes, List<ActionDto>? Actions, string? ResponseId);

    internal static async Task<(bool ok, int statusCode, string raw, TimeSpan elapsed, int requestBytes)> SendOpenAIRequestAsync(
        string apiKey, object body, CancellationToken cancellationToken = default)
    {
        var result = await SendAsync(apiKey, body, control: false, cancellationToken);
        return (result.Ok, result.StatusCode, result.Raw, result.Elapsed, result.RequestBytes);
    }

    internal static async Task<(bool ok, int statusCode, string raw, TimeSpan elapsed, int requestBytes,
        int responseBytes, List<ActionDto>? acceptedActions, string? responseId)> SendOpenAIControlRequestAsync(
        string apiKey, object body, CancellationToken cancellationToken = default)
    {
        var result = await SendAsync(apiKey, body, control: true, cancellationToken);
        return (result.Ok, result.StatusCode, result.Raw, result.Elapsed, result.RequestBytes,
            result.ResponseBytes, result.AcceptedActions, result.ResponseId);
    }

    private static async Task<TransportResult> SendAsync(string apiKey, object body, bool control, CancellationToken cancellationToken)
    {
        var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(body);
        var timer = RuntimeClock.StartTimer();
        var raw = "";
        var statusCode = 0;
        var responseBytes = 0;
        var outcome = new ApiResponseOutcome(false, false, "");
        CurrentRun.LastOpenAiFailureWasRetriable = false;
        CurrentRun.LastOpenAiFailureKind = "";
        CurrentRun.LastOpenAiResponseId = null;
        for (var attempt = 0; attempt <= Options.OpenAiMaxRetries; attempt++)
        {
            try
            {
                using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                if (Options.OpenAiTimeoutSeconds > 0)
                    timeout.CancelAfter(TimeSpan.FromSeconds(Options.OpenAiTimeoutSeconds));
                using var request = new HttpRequestMessage(HttpMethod.Post, Options.ApiUrl)
                { Content = new ByteArrayContent(jsonBytes) };
                request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                using var response = await HttpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token);
                statusCode = (int)response.StatusCode;
                ResponseBody received;
                if (response.IsSuccessStatusCode &&
                    response.Content.Headers.ContentType?.MediaType?.Equals("text/event-stream", StringComparison.OrdinalIgnoreCase) == true)
                    received = await ReadResponseStreamAsync(response.Content, control && !Options.UsePreviousResponseState, timeout.Token);
                else
                {
                    var text = await response.Content.ReadAsStringAsync(timeout.Token);
                    received = new(text, Encoding.UTF8.GetByteCount(text), null, TryGetCompletedResponseId(text));
                }
                raw = received.Raw;
                responseBytes += received.Bytes;
                outcome = ApiResponseOutcome.Classify(statusCode, raw);
                if (outcome.Accepted)
                {
                    timer.Stop();
                    CurrentRun.LastOpenAiResponseId = IsCompletedResponseJson(raw) ? received.ResponseId : null;
                    return new(true, statusCode, raw, timer.Elapsed, jsonBytes.Length, responseBytes, received.Actions, received.ResponseId);
                }
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                raw = "cancelled";
                statusCode = 0;
                outcome = new(false, false, "cancelled");
                break;
            }
            catch (Exception ex) when (ex is HttpRequestException or OperationCanceledException or IOException)
            {
                raw = ex.Message;
                statusCode = 0;
                outcome = new(false, true, "transport_or_timeout");
            }
            if (!outcome.Retriable || attempt == Options.OpenAiMaxRetries) break;
            CurrentRun.RunOpenAiRetries++;
            Console.WriteLine($"[openai] retry {attempt + 1}/{Options.OpenAiMaxRetries} after {outcome.FailureKind}");
            try { await RuntimeClock.Delay(TimeSpan.FromMilliseconds(250 * Math.Pow(2, attempt)), cancellationToken); }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                raw = "cancelled";
                statusCode = 0;
                outcome = new(false, false, "cancelled");
                break;
            }
        }
        timer.Stop();
        CurrentRun.LastOpenAiFailureWasRetriable = outcome.Retriable;
        CurrentRun.LastOpenAiFailureKind = outcome.FailureKind;
        return new(false, statusCode, raw, timer.Elapsed, jsonBytes.Length, responseBytes, null, null);
    }

    private static async Task<ResponseBody> ReadResponseStreamAsync(HttpContent content, bool allowEarlyAccept, CancellationToken token)
    {
        using var stream = await content.ReadAsStreamAsync(token);
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var data = new StringBuilder();
        string? responseId = null;
        var bytes = 0;
        while (true)
        {
            var line = await reader.ReadLineAsync(token);
            if (line is not null) bytes += Encoding.UTF8.GetByteCount(line) + 1;
            if (line is null || line.Length == 0)
            {
                if (data.Length > 0)
                {
                    var eventData = data.ToString();
                    data.Clear();
                    if (TryHandleControlStreamEvent(eventData, allowEarlyAccept, ref responseId,
                        out var raw, out var actions, out var completed))
                        return new(raw, bytes, actions, responseId);
                    if (completed) break;
                }
                if (line is null) break;
            }
            else if (line.StartsWith("data:", StringComparison.Ordinal))
            {
                if (data.Length > 0) data.Append('\n');
                data.Append(line.AsSpan(5).TrimStart());
            }
        }
        throw new IOException("OpenAI stream ended before a final control response was received.");
    }
}
