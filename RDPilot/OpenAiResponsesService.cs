internal static partial class RDPilotApplication
{
    /// <summary>
    /// Communicates with the OpenAI Responses API and parses structured model output.
    /// </summary>
    internal static class OpenAiResponsesService
    {
            // ==== API calls ====
            internal static async Task<(bool ok, int statusCode, string raw, TimeSpan elapsed, int requestBytes)> SendOpenAIRequestAsync(string apiKey, object body, CancellationToken cancellationToken = default)
            {
                var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(body);
                var requestBytes = jsonBytes.Length;
                var sw = Stopwatch.StartNew();
                string raw = "";
                int statusCode = 0;
                LastOpenAiFailureWasRetriable = false;
                LastOpenAiFailureKind = "";
                LastOpenAiResponseId = null;

                for (var attempt = 0; attempt <= OpenAiMaxRetries; attempt++)
                {
                    try
                    {
                        using var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl)
                        {
                            Content = new ByteArrayContent(jsonBytes)
                        };
                        request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                        using var resp = await OpenAiHttp.SendAsync(request, cancellationToken);
                        raw = await resp.Content.ReadAsStringAsync(cancellationToken);
                        statusCode = (int)resp.StatusCode;

                        if (resp.IsSuccessStatusCode)
                        {
                            sw.Stop();
                            LastOpenAiFailureWasRetriable = false;
                            LastOpenAiFailureKind = "";
                            return (true, statusCode, raw, sw.Elapsed, requestBytes);
                        }

                        if (!IsRetriableStatus(statusCode) || attempt == OpenAiMaxRetries)
                            break;

                        RunOpenAiRetries++;
                        Console.WriteLine($"[openai] retry {attempt + 1}/{OpenAiMaxRetries} after HTTP {statusCode}");
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        raw = "cancelled";
                        statusCode = 0;
                        break;
                    }
                    catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
                    {
                        raw = ex.Message;
                        statusCode = 0;

                        if (attempt == OpenAiMaxRetries)
                            break;

                        RunOpenAiRetries++;
                        Console.WriteLine($"[openai] retry {attempt + 1}/{OpenAiMaxRetries} after transport error: {ex.Message}");
                    }

                    try { await Task.Delay(TimeSpan.FromMilliseconds(250 * Math.Pow(2, attempt)), cancellationToken); }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        raw = "cancelled";
                        statusCode = 0;
                        break;
                    }
                }

                sw.Stop();
                LastOpenAiFailureWasRetriable = statusCode == 0 || IsRetriableStatus(statusCode);
                LastOpenAiFailureKind = statusCode == 0
                    ? (raw.Equals("cancelled", StringComparison.OrdinalIgnoreCase) ? "cancelled" : "transport_or_timeout")
                    : $"http_{statusCode}";
                return (false, statusCode, raw, sw.Elapsed, requestBytes);
            }

            internal static async Task<(bool ok, int statusCode, string raw, TimeSpan elapsed, int requestBytes, int responseBytes, List<ActionDto>? acceptedActions, string? responseId)> SendOpenAIControlRequestAsync(
                string apiKey,
                object body,
                CancellationToken cancellationToken = default)
            {
                var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(body);
                var requestBytes = jsonBytes.Length;
                var allowEarlyAccept = !UsePreviousResponseState;

                var sw = Stopwatch.StartNew();
                string raw = "";
                string? responseId = null;
                var responseBytes = 0;
                var statusCode = 0;
                LastOpenAiFailureWasRetriable = false;
                LastOpenAiFailureKind = "";
                LastOpenAiResponseId = null;

                for (var attempt = 0; attempt <= OpenAiMaxRetries; attempt++)
                {
                    try
                    {
                        using var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl)
                        {
                            Content = new ByteArrayContent(jsonBytes)
                        };
                        request.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                        using var resp = await OpenAiHttp.SendAsync(
                            request,
                            HttpCompletionOption.ResponseHeadersRead,
                            cancellationToken);
                        statusCode = (int)resp.StatusCode;
                        if (!resp.IsSuccessStatusCode)
                        {
                            raw = await resp.Content.ReadAsStringAsync(cancellationToken);
                            responseBytes += Encoding.UTF8.GetByteCount(raw);
                            if (!IsRetriableStatus(statusCode) || attempt == OpenAiMaxRetries)
                                break;

                            RunOpenAiRetries++;
                            Console.WriteLine($"[openai] retry {attempt + 1}/{OpenAiMaxRetries} after HTTP {statusCode}");
                        }
                        else if (resp.Content.Headers.ContentType?.MediaType?.Equals(
                                     "text/event-stream",
                                     StringComparison.OrdinalIgnoreCase) != true)
                        {
                            raw = await resp.Content.ReadAsStringAsync(cancellationToken);
                            responseBytes += Encoding.UTF8.GetByteCount(raw);
                            responseId = TryGetCompletedResponseId(raw);
                            LastOpenAiResponseId = responseId;
                            sw.Stop();
                            return (true, statusCode, raw, sw.Elapsed, requestBytes, responseBytes, null, responseId);
                        }
                        else
                        {
                            using var stream = await resp.Content.ReadAsStreamAsync(cancellationToken);
                            using var reader = new StreamReader(stream, Encoding.UTF8);
                            var data = new StringBuilder();
                            while (true)
                            {
                                var line = await reader.ReadLineAsync(cancellationToken);
                                if (line is null)
                                {
                                    if (data.Length > 0 && TryHandleControlStreamEvent(
                                            data.ToString(),
                                            allowEarlyAccept,
                                            ref responseId,
                                            out raw,
                                            out var acceptedAtEnd,
                                        out var completedAtEnd))
                                    {
                                        sw.Stop();
                                        LastOpenAiResponseId = IsCompletedResponseJson(raw) ? responseId : null;
                                        return (true, statusCode, raw, sw.Elapsed, requestBytes, responseBytes, acceptedAtEnd, responseId);
                                    }
                                    break;
                                }

                                responseBytes += Encoding.UTF8.GetByteCount(line) + 1;
                                if (line.Length == 0)
                                {
                                    if (data.Length == 0)
                                        continue;

                                    var eventData = data.ToString();
                                    data.Clear();
                                    if (TryHandleControlStreamEvent(
                                        eventData,
                                        allowEarlyAccept,
                                        ref responseId,
                                        out raw,
                                        out var accepted,
                                        out var completed))
                                    {
                                        sw.Stop();
                                        LastOpenAiResponseId = IsCompletedResponseJson(raw) ? responseId : null;
                                        return (true, statusCode, raw, sw.Elapsed, requestBytes, responseBytes, accepted, responseId);
                                    }
                                    if (completed)
                                        break;
                                    continue;
                                }

                                if (line.StartsWith("data:", StringComparison.Ordinal))
                                {
                                    if (data.Length > 0)
                                        data.Append('\n');
                                    data.Append(line.AsSpan(5).TrimStart());
                                }
                            }

                            raw = "OpenAI stream ended before a complete control response was received.";
                            statusCode = 0;
                            if (attempt == OpenAiMaxRetries)
                                break;
                            RunOpenAiRetries++;
                            Console.WriteLine($"[openai] retry {attempt + 1}/{OpenAiMaxRetries} after incomplete response stream");
                        }
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        raw = "cancelled";
                        statusCode = 0;
                        break;
                    }
                    catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or IOException)
                    {
                        raw = ex.Message;
                        statusCode = 0;
                        if (attempt == OpenAiMaxRetries)
                            break;
                        RunOpenAiRetries++;
                        Console.WriteLine($"[openai] retry {attempt + 1}/{OpenAiMaxRetries} after transport error: {ex.Message}");
                    }

                    try { await Task.Delay(TimeSpan.FromMilliseconds(250 * Math.Pow(2, attempt)), cancellationToken); }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        raw = "cancelled";
                        statusCode = 0;
                        break;
                    }
                }

                sw.Stop();
                LastOpenAiFailureWasRetriable = statusCode == 0 || IsRetriableStatus(statusCode);
                LastOpenAiFailureKind = statusCode == 0
                    ? (raw.Equals("cancelled", StringComparison.OrdinalIgnoreCase) ? "cancelled" : "transport_or_timeout")
                    : $"http_{statusCode}";
                return (false, statusCode, raw, sw.Elapsed, requestBytes, responseBytes, null, responseId);
            }

            internal static bool TryHandleControlStreamEvent(
                string eventData,
                bool allowEarlyAccept,
                ref string? responseId,
                out string raw,
                out List<ActionDto>? acceptedActions,
                out bool completed)
            {
                raw = "";
                acceptedActions = null;
                completed = false;
                if (string.IsNullOrWhiteSpace(eventData) || eventData == "[DONE]")
                    return false;

                try
                {
                    using var eventDoc = JsonDocument.Parse(eventData);
                    var root = eventDoc.RootElement;
                    var eventType = root.TryGetProperty("type", out var typeElement) &&
                                    typeElement.ValueKind == JsonValueKind.String
                        ? typeElement.GetString()
                        : null;

                    if (eventType == "response.created" &&
                        root.TryGetProperty("response", out var createdResponse) &&
                        createdResponse.TryGetProperty("id", out var createdId) &&
                        createdId.ValueKind == JsonValueKind.String)
                    {
                        responseId = createdId.GetString();
                        return false;
                    }

                    if (allowEarlyAccept && eventType == "response.output_text.done" &&
                        root.TryGetProperty("text", out var textElement) &&
                        textElement.ValueKind == JsonValueKind.String &&
                        TryParseControlActionJson(textElement.GetString(), out var actions))
                    {
                        acceptedActions = actions;
                        raw = BuildAcceptedStreamResponse(responseId, textElement.GetString()!);
                        return true;
                    }

                    if (eventType is "response.completed" or "response.incomplete" or "response.failed")
                    {
                        completed = true;
                        if (root.TryGetProperty("response", out var finalResponse))
                        {
                            raw = finalResponse.GetRawText();
                            if (finalResponse.TryGetProperty("id", out var finalId) &&
                                finalId.ValueKind == JsonValueKind.String)
                            {
                                responseId = finalId.GetString();
                            }
                            return true;
                        }
                    }
                }
                catch (JsonException)
                {
                    // Ignore malformed secondary SSE data; a later completed event may still be valid.
                }

                return false;
            }

            internal static bool TryParseControlActionJson(string? json, out List<ActionDto> actions)
            {
                actions = new List<ActionDto>();
                var extracted = TryExtractJsonObject(json ?? "");
                if (extracted is null)
                    return false;

                try
                {
                    var batch = JsonSerializer.Deserialize<ActionBatchDto>(extracted);
                    var batchActions = batch is null ? Array.Empty<ActionDto>() : ApplyActionBatchMetadata(batch);
                    if (batchActions.Length == 0)
                        return false;
                    foreach (var action in batchActions)
                    {
                        if (!IsKnownAction(action))
                            break;
                        actions.Add(action);
                    }
                    return actions.Count > 0;
                }
                catch
                {
                    return false;
                }
            }

            internal static string BuildAcceptedStreamResponse(string? responseId, string outputText) =>
                JsonSerializer.Serialize(new
                {
                    id = responseId,
                    @object = "response",
                    status = "completed",
                    output = new[]
                    {
                        new
                        {
                            type = "message",
                            status = "completed",
                            role = "assistant",
                            content = new[] { new { type = "output_text", text = outputText } }
                        }
                    },
                    rdpilot_stream_early_accept = true
                });

            internal static bool IsRetriableStatus(int statusCode) =>
                statusCode == 408 || statusCode == 409 || statusCode == 429 || statusCode >= 500;

            internal static void RecordUsageMetrics(JsonElement root)
            {
                if (!root.TryGetProperty("usage", out var usage))
                    return;

                if (usage.TryGetProperty("input_tokens", out var input) && input.TryGetInt64(out var inputTokens))
                    RunInputTokens += inputTokens;
                if (usage.TryGetProperty("output_tokens", out var output) && output.TryGetInt64(out var outputTokens))
                    RunOutputTokens += outputTokens;
                if (usage.TryGetProperty("input_tokens_details", out var inputDetails) &&
                    inputDetails.TryGetProperty("cached_tokens", out var cached) &&
                    cached.TryGetInt64(out var cachedTokens))
                    RunCachedTokens += cachedTokens;
                if (usage.TryGetProperty("output_tokens_details", out var outputDetails) &&
                    outputDetails.TryGetProperty("reasoning_tokens", out var reasoning) &&
                    reasoning.TryGetInt64(out var reasoningTokens))
                    RunReasoningTokens += reasoningTokens;
            }

            internal static void RecordResponseId(JsonElement root)
            {
                if (IsCompletedResponse(root) &&
                    root.TryGetProperty("id", out var id) &&
                    id.ValueKind == JsonValueKind.String)
                {
                    LastOpenAiResponseId = id.GetString();
                }
            }

            internal static string? TryGetCompletedResponseId(string raw)
            {
                try
                {
                    using var doc = JsonDocument.Parse(raw);
                    var root = doc.RootElement;
                    if (IsCompletedResponse(root) &&
                        root.TryGetProperty("id", out var id) &&
                        id.ValueKind == JsonValueKind.String)
                    {
                        return id.GetString();
                    }
                }
                catch (JsonException)
                {
                }

                return null;
            }

            internal static bool IsCompletedResponseJson(string raw)
            {
                try
                {
                    using var doc = JsonDocument.Parse(raw);
                    return IsCompletedResponse(doc.RootElement);
                }
                catch (JsonException)
                {
                    return false;
                }
            }

            internal static bool IsCompletedResponse(JsonElement root) =>
                root.TryGetProperty("status", out var status) &&
                status.ValueKind == JsonValueKind.String &&
                status.GetString()?.Equals("completed", StringComparison.OrdinalIgnoreCase) == true;

            internal static bool ContainsCompactionItem(JsonElement root)
            {
                if (!root.TryGetProperty("output", out var output) ||
                    output.ValueKind != JsonValueKind.Array)
                {
                    return false;
                }

                foreach (var item in output.EnumerateArray())
                {
                    if (item.TryGetProperty("type", out var type) &&
                        type.ValueKind == JsonValueKind.String &&
                        type.GetString()?.Equals("compaction", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        return true;
                    }
                }

                return false;
            }

            internal static bool IsIncompleteDueToMaxOutput(JsonElement root) =>
                root.TryGetProperty("status", out var status) &&
                status.ValueKind == JsonValueKind.String &&
                status.GetString()?.Equals("incomplete", StringComparison.OrdinalIgnoreCase) == true &&
                root.TryGetProperty("incomplete_details", out var details) &&
                details.ValueKind == JsonValueKind.Object &&
                details.TryGetProperty("reason", out var reason) &&
                reason.ValueKind == JsonValueKind.String &&
                reason.GetString()?.Equals("max_output_tokens", StringComparison.OrdinalIgnoreCase) == true;

            internal static string? LowerReasoningEffortOneStep(string? effort) =>
                effort?.Trim().ToLowerInvariant() switch
                {
                    "max" => "xhigh",
                    "xhigh" => "high",
                    "high" => "medium",
                    "medium" => "low",
                    _ => null
                };

            internal static bool TryBuildMaxOutputRetryBody(
                object body,
                out object retryBody,
                out int retryMaxTokens,
                out string retryEffort,
                out bool effortChanged,
                int sameEffortRetryCount = 0)
            {
                var node = JsonSerializer.SerializeToNode(body) as JsonObject
                           ?? throw new InvalidOperationException("Could not clone OpenAI request body for max_output_tokens retry.");

                var currentMax = MaxOutputTokens;
                if (node.TryGetPropertyValue("max_output_tokens", out var maxNode) &&
                    maxNode is not null &&
                    maxNode.GetValueKind() == JsonValueKind.Number &&
                    maxNode.GetValue<int>() > 0)
                {
                    currentMax = maxNode.GetValue<int>();
                }

                var cap = Math.Max(currentMax, IncompleteMaxOutputTokenCap);
                var grownMax = currentMax < 1000
                    ? Math.Max(1200L, (long)currentMax * 4L)
                    : Math.Max((long)currentMax + 1000L, (long)currentMax * 3L);
                var grownMaxTokens = (int)Math.Min(grownMax, cap);

                JsonObject? reasoning = null;
                string? currentEffort = null;
                if (node.TryGetPropertyValue("reasoning", out var reasoningNode) &&
                    reasoningNode is JsonObject reasoningObject)
                {
                    reasoning = reasoningObject;
                    if (reasoning.TryGetPropertyValue("effort", out var effortNode) &&
                        effortNode is not null &&
                        effortNode.GetValueKind() == JsonValueKind.String)
                    {
                        currentEffort = effortNode.GetValue<string>();
                    }
                }

                var lowerEffort = LowerReasoningEffortOneStep(currentEffort);
                var shouldLowerEffort = lowerEffort is not null;
                var canGrowWithoutEffort = currentEffort is null &&
                    sameEffortRetryCount < Math.Max(0, IncompleteMaxOutputRetries) &&
                    grownMaxTokens > currentMax;
                retryMaxTokens = shouldLowerEffort
                    ? currentMax
                    : canGrowWithoutEffort
                        ? grownMaxTokens
                        : currentMax;
                var nextEffort = shouldLowerEffort ? lowerEffort : currentEffort;

                retryEffort = nextEffort ?? "unchanged";
                effortChanged = !string.Equals(nextEffort, currentEffort, StringComparison.OrdinalIgnoreCase);
                if (retryMaxTokens == currentMax && !effortChanged)
                {
                    retryBody = body;
                    return false;
                }

                node["max_output_tokens"] = retryMaxTokens;
                if (reasoning is not null && effortChanged && nextEffort is not null)
                    reasoning["effort"] = nextEffort;

                retryBody = node;
                return true;
            }

            internal static object BuildMaxOutputRetryBody(object body, out int retryMaxTokens, out string retryEffort)
            {
                if (!TryBuildMaxOutputRetryBody(body, out var retryBody, out retryMaxTokens, out retryEffort, out _))
                    throw new InvalidOperationException("No larger output budget or lower reasoning effort is available for retry.");
                return retryBody;
            }

            internal static bool IsPreviousResponseStateFailure(int statusCode, string raw)
            {
                if (statusCode is not (400 or 404) || string.IsNullOrWhiteSpace(raw))
                    return false;

                var normalized = raw.ToLowerInvariant();
                return normalized.Contains("previous_response_id", StringComparison.Ordinal) ||
                       normalized.Contains("previous response", StringComparison.Ordinal) &&
                       (normalized.Contains("not found", StringComparison.Ordinal) ||
                        normalized.Contains("invalid", StringComparison.Ordinal) ||
                        normalized.Contains("incomplete", StringComparison.Ordinal) ||
                        normalized.Contains("expired", StringComparison.Ordinal));
            }

            internal static bool IsContextManagementFailure(int statusCode, string raw)
            {
                if (statusCode != 400 || string.IsNullOrWhiteSpace(raw))
                    return false;

                var normalized = raw.ToLowerInvariant();
                return normalized.Contains("context_management", StringComparison.Ordinal) ||
                       normalized.Contains("compact_threshold", StringComparison.Ordinal);
            }

            internal static bool TryBuildRequestWithoutProperty(
                object body,
                string propertyName,
                out object fallbackBody)
            {
                var node = JsonSerializer.SerializeToNode(body) as JsonObject;
                if (node is null || !node.Remove(propertyName))
                {
                    fallbackBody = body;
                    return false;
                }

                fallbackBody = node;
                return true;
            }

            internal static async Task<(ActionDto? parsed, string raw, string? responseId, bool contextFallbackUsed, bool compactionFallbackUsed, bool compactionOccurred)> CallOpenAIAsync(
                string apiKey,
                object body,
                CancellationToken cancellationToken = default,
                bool allowDrawGestureBatch = false,
                int observedTurnBatchLimit = 0)
            {
                var requestBody = body;
                var incompleteRetryCount = 0;
                var sameEffortRetryCount = 0;
                var contextFallbackUsed = false;
                var compactionFallbackUsed = false;
                while (true)
                {
                    var (ok, statusCode, raw, elapsed, requestBytes, responseBytes, acceptedActions, responseId) =
                        await SendOpenAIControlRequestAsync(apiKey, requestBody, cancellationToken);
                    RunOpenAiCalls++;
                    RunOpenAiElapsed += elapsed;
                    RunOpenAiRequestBytes += requestBytes;
                    RunOpenAiBytes += responseBytes;
                    Console.WriteLine($"[openai] {(ok ? "ok" : "error")} in {elapsed.TotalSeconds:0.0}s; request_bytes={requestBytes}; response_bytes={responseBytes}{(acceptedActions is not null ? "; accepted=first_complete_action_batch" : "")}");

                    if (!ok)
                    {
                        if (!compactionFallbackUsed &&
                            IsContextManagementFailure(statusCode, raw) &&
                            TryBuildRequestWithoutProperty(
                                requestBody,
                                "context_management",
                                out var requestWithoutCompaction))
                        {
                            RunOpenAiRetries++;
                            RunControlCompactionFallbacks++;
                            compactionFallbackUsed = true;
                            requestBody = requestWithoutCompaction;
                            Console.WriteLine("[context] server-side compaction was rejected by the API; retrying this turn without context_management.");
                            continue;
                        }

                        if (!contextFallbackUsed &&
                            IsPreviousResponseStateFailure(statusCode, raw) &&
                            TryBuildRequestWithoutProperty(
                                requestBody,
                                "previous_response_id",
                                out var requestWithoutPreviousResponse))
                        {
                            RunOpenAiRetries++;
                            RunControlContextFallbacks++;
                            contextFallbackUsed = true;
                            requestBody = requestWithoutPreviousResponse;
                            Console.WriteLine("[context] previous_response_id was rejected; retrying this turn from the application checkpoint.");
                            continue;
                        }

                        Console.WriteLine($"OpenAI HTTP {statusCode}: {raw}");
                        return (null, raw, null, contextFallbackUsed, compactionFallbackUsed, false);
                    }

                    if (acceptedActions is { Count: > 0 })
                    {
                        RunEarlyAcceptedControlStreams++;
                        if (acceptedActions.Count > 1)
                        {
                            RunMultiCandidateResponses++;
                            Console.WriteLine($"[openai] stream accepted a {acceptedActions.Count}-action sequence; evaluating guarded follow-ups.");
                            QueueSafeCandidates(
                                acceptedActions,
                                allowDrawGestureBatch,
                                observedTurnBatchLimit);
                        }
                        return (acceptedActions[0], raw, responseId, contextFallbackUsed, compactionFallbackUsed, false);
                    }

                    try
                    {
                        using var doc = JsonDocument.Parse(raw);
                        var root = doc.RootElement;
                        RecordResponseId(root);
                        RecordUsageMetrics(root);
                        var completedResponseId = IsCompletedResponse(root)
                            ? responseId ?? TryGetCompletedResponseId(raw)
                            : null;
                        var compactionOccurred = IsCompletedResponse(root) &&
                                                 ContainsCompactionItem(root);
                        if (compactionOccurred)
                            RunControlContextCompactions++;

                        if (IsCompletedResponse(root) &&
                            TryParseControlActionSequence(
                                root,
                                out var validCandidates,
                                out var payloadCount,
                                out var legacyPayload))
                        {
                            if (payloadCount > 1)
                                Console.WriteLine($"[openai] warning: response contained {payloadCount} {(legacyPayload ? "legacy action" : "action-sequence")} payloads; using the first valid payload.");

                            if (validCandidates.Count > 1)
                            {
                                RunMultiCandidateResponses++;
                                Console.WriteLine(
                                    ExecuteMultiActionCandidates
                                        ? $"[openai] response proposed a {validCandidates.Count}-action sequence; evaluating guarded follow-ups."
                                        : $"[openai] response proposed a {validCandidates.Count}-action sequence; batching is disabled, using the first action.");
                                QueueSafeCandidates(
                                    validCandidates,
                                    allowDrawGestureBatch,
                                    observedTurnBatchLimit);
                            }
                            return (validCandidates[0], raw, completedResponseId, contextFallbackUsed, compactionFallbackUsed, compactionOccurred);
                        }

                        if (IsIncompleteDueToMaxOutput(root) &&
                            TryBuildMaxOutputRetryBody(requestBody, out var retryBody, out var retryMaxTokens, out var retryEffort, out var effortChanged, sameEffortRetryCount))
                        {
                            RunOpenAiRetries++;
                            incompleteRetryCount++;
                            sameEffortRetryCount = effortChanged ? 0 : sameEffortRetryCount + 1;
                            requestBody = retryBody;
                            Console.WriteLine($"[openai] incomplete=max_output_tokens; retry {incompleteRetryCount} with max_output_tokens={retryMaxTokens}, reasoning_effort={retryEffort}");
                            continue;
                        }

                        if (IsIncompleteDueToMaxOutput(root))
                        {
                            Console.WriteLine($"[openai] incomplete=max_output_tokens after {incompleteRetryCount + 1} attempt(s); reasoning effort reached the automatic fallback floor (low), or no configured effort ladder is available. No complete JSON action was emitted.");
                            return (null, raw, null, contextFallbackUsed, compactionFallbackUsed, false);
                        }

                        Console.WriteLine("No parsable JSON found (parsed/output_text). RAW:");
                        Console.WriteLine(raw);
                        return (null, raw, completedResponseId, contextFallbackUsed, compactionFallbackUsed, compactionOccurred);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Response parsing error: {ex.Message}");
                        Console.WriteLine(raw);
                        return (null, raw, responseId, contextFallbackUsed, compactionFallbackUsed, false);
                    }
                }

            }

            internal static async Task<(T? parsed, string raw)> CallOpenAIParsedAsync<T>(string apiKey, object body, CancellationToken cancellationToken = default)
            {
                var requestBody = body;
                var incompleteRetryCount = 0;
                var sameEffortRetryCount = 0;
                while (true)
                {
                    var (ok, statusCode, raw, elapsed, requestBytes) = await SendOpenAIRequestAsync(apiKey, requestBody, cancellationToken);
                    RunOpenAiCalls++;
                    RunOpenAiElapsed += elapsed;
                    RunOpenAiRequestBytes += requestBytes;
                    var responseBytes = Encoding.UTF8.GetByteCount(raw);
                    RunOpenAiBytes += responseBytes;
                    Console.WriteLine($"[openai] {(ok ? "ok" : "error")} in {elapsed.TotalSeconds:0.0}s; request_bytes={requestBytes}; response_bytes={responseBytes}");

                    if (!ok)
                    {
                        Console.WriteLine($"OpenAI HTTP {statusCode}: {raw}");
                        return (default, raw);
                    }

                    try
                    {
                        using var doc = JsonDocument.Parse(raw);
                        var root = doc.RootElement;
                        RecordResponseId(root);
                        RecordUsageMetrics(root);

                        if (TryParseResponsePayload<T>(root, out var parsed, out var candidates, out _))
                        {
                            if (candidates > 1)
                            {
                                RunMultiCandidateResponses++;
                                Console.WriteLine($"[openai] warning: response contained {candidates} JSON candidates; using the first.");
                            }
                            return (parsed, raw);
                        }

                        if (IsIncompleteDueToMaxOutput(root) &&
                            TryBuildMaxOutputRetryBody(requestBody, out var retryBody, out var retryMaxTokens, out var retryEffort, out var effortChanged, sameEffortRetryCount))
                        {
                            RunOpenAiRetries++;
                            incompleteRetryCount++;
                            sameEffortRetryCount = effortChanged ? 0 : sameEffortRetryCount + 1;
                            requestBody = retryBody;
                            Console.WriteLine($"[openai] incomplete=max_output_tokens; retry {incompleteRetryCount} with max_output_tokens={retryMaxTokens}, reasoning_effort={retryEffort}");
                            continue;
                        }

                        if (IsIncompleteDueToMaxOutput(root))
                        {
                            Console.WriteLine($"[openai] incomplete=max_output_tokens after {incompleteRetryCount + 1} attempt(s); reasoning effort reached the automatic fallback floor (low), or no configured effort ladder is available. No complete JSON payload was emitted.");
                            return (default, raw);
                        }

                        Console.WriteLine("No parsable JSON found (parsed/output_text). RAW:");
                        Console.WriteLine(raw);
                        return (default, raw);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Response parsing error: {ex.Message}");
                        Console.WriteLine(raw);
                        return (default, raw);
                    }
                }

            }

            internal static bool TryParseControlActionSequence(
                JsonElement root,
                out List<ActionDto> actions,
                out int payloadCount,
                out bool legacyPayload)
            {
                actions = new List<ActionDto>();
                legacyPayload = false;

                if (TryParseResponsePayload<ActionBatchDto>(
                        root,
                        out _,
                        out payloadCount,
                        out var parsedBatches))
                {
                    var batch = parsedBatches
                        .Select(ApplyActionBatchMetadata)
                        .FirstOrDefault(candidate => candidate.Length > 0 && IsKnownAction(candidate[0]));
                    if (batch is not null)
                    {
                        foreach (var action in batch)
                        {
                            if (!IsKnownAction(action))
                                break;
                            actions.Add(action);
                        }
                    }

                    if (actions.Count > 0)
                        return true;
                }

                // Compatibility with previously logged or non-schema responses.
                legacyPayload = true;
                if (TryParseResponsePayload<ActionDto>(
                        root,
                        out _,
                        out payloadCount,
                        out var legacyActions))
                {
                    var first = legacyActions.FirstOrDefault(IsKnownAction);
                    if (first is not null)
                    {
                        actions.Add(first);
                        return true;
                    }
                }

                return false;
            }

            internal static ActionDto[] ApplyActionBatchMetadata(ActionBatchDto batch)
            {
                var actions = batch.Actions ?? Array.Empty<ActionDto>();
                if (actions.Length == 0)
                    return actions;

                var first = actions[0];
                first.Confidence ??= batch.Confidence;
                first.Note ??= batch.Note;
                first.WorldStateSummary ??= batch.WorldStateSummary;
                first.MechanicsHypothesis ??= batch.MechanicsHypothesis;
                first.SalientChangeObservation ??= batch.SalientChangeObservation;
                first.ShortTermPlan ??= batch.ShortTermPlan;
                first.PlanStatus ??= batch.PlanStatus;
                first.PlanRevisionReason ??= batch.PlanRevisionReason;
                first.PlannedInputs ??= batch.PlannedInputs;
                first.PlanWaypoint ??= batch.PlanWaypoint;
                first.PlanStateId ??= batch.PlanStateId;
                first.PlanConfidence ??= batch.PlanConfidence;
                first.RecoveryStrategyId ??= batch.RecoveryStrategyId;
                first.RecoveryStrategyStep ??= batch.RecoveryStrategyStep;
                return actions;
            }

            internal static bool TryParseResponsePayload<T>(JsonElement root, out T? parsed, out int candidateCount, out List<T> parsedCandidates)
            {
                parsed = default;
                parsedCandidates = new List<T>();
                var count = 0;
                var jsonCandidates = new List<string>();
                var seenCandidates = new HashSet<string>(StringComparer.Ordinal);

                void AddCandidate(string? json)
                {
                    if (string.IsNullOrWhiteSpace(json))
                        return;

                    json = json.Trim();
                    if (!seenCandidates.Add(json))
                        return;

                    count++;
                    jsonCandidates.Add(json);
                }

                if (root.TryGetProperty("output_parsed", out var op))
                    AddCandidate(op.GetRawText());

                if (root.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in output.EnumerateArray())
                    {
                        if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array)
                            continue;

                        foreach (var c in content.EnumerateArray())
                        {
                            if (c.TryGetProperty("parsed", out var parsedEl))
                            {
                                AddCandidate(parsedEl.GetRawText());
                                continue;
                            }

                            if (c.TryGetProperty("type", out var tEl) &&
                                tEl.ValueKind == JsonValueKind.String &&
                                tEl.GetString() == "output_text" &&
                                c.TryGetProperty("text", out var txtEl) &&
                                txtEl.ValueKind == JsonValueKind.String)
                            {
                                AddCandidate(TryExtractJsonObject(txtEl.GetString() ?? ""));
                            }
                        }
                    }
                }

                if (root.TryGetProperty("output_text", out var outText) && outText.ValueKind == JsonValueKind.String)
                    AddCandidate(TryExtractJsonObject(outText.GetString() ?? ""));

                if (root.TryGetProperty("text", out var textEl) && textEl.ValueKind == JsonValueKind.String)
                    AddCandidate(TryExtractJsonObject(textEl.GetString() ?? ""));

                foreach (var json in jsonCandidates)
                {
                    try
                    {
                        var candidate = JsonSerializer.Deserialize<T>(json);
                        if (candidate is not null)
                            parsedCandidates.Add(candidate);
                    }
                    catch
                    {
                        // Ignore malformed secondary candidates; the raw response is still logged.
                    }
                }

                if (parsedCandidates.Count == 0)
                {
                    candidateCount = count;
                    return false;
                }

                parsed = parsedCandidates[0];
                candidateCount = count;
                return true;
            }

            internal static void QueueSafeCandidates(
                List<ActionDto> candidates,
                bool allowDrawGestureBatch = false,
                int observedTurnBatchLimit = 0)
            {
                if (!ExecuteMultiActionCandidates || MaxQueuedBatchActions <= 0 || candidates.Count <= 1)
                    return;

                var queuedActions = SafeBatchFollowUps(
                    candidates,
                    allowDrawGestureBatch,
                    observedTurnBatchLimit);
                foreach (var candidate in queuedActions)
                    PendingSafeActions.Enqueue(candidate);

                if (queuedActions.Length > 0)
                {
                    Console.WriteLine($"[batch] queued {queuedActions.Length} safe follow-up action(s) from the same response.");
                    if (candidates[0].Type == "drag_path")
                    {
                        var acceptedGestures = new[] { candidates[0] }.Concat(queuedActions).ToArray();
                        Console.WriteLine(
                            $"[batch] stable-canvas draw sequence accepted; strokes={acceptedGestures.Length}; points={acceptedGestures.Sum(action => action.Path?.Length ?? 0)}; duration_ms={acceptedGestures.Sum(action => action.DurationMs ?? 0)}");
                    }
                }
                if (queuedActions.Length < candidates.Count - 1)
                {
                    Console.WriteLine(
                        $"[batch] observation barrier discarded {candidates.Count - 1 - queuedActions.Length} unsafe or state-dependent follow-up action(s).");
                }
            }

            internal static ActionDto[] SafeBatchFollowUps(
                List<ActionDto> candidates,
                bool allowDrawGestureBatch = false,
                int observedTurnBatchLimit = 0)
            {
                if (MaxQueuedBatchActions <= 0 || candidates.Count <= 1)
                    return Array.Empty<ActionDto>();

                if (observedTurnBatchLimit > 0)
                {
                    var observedTurnActions = SafeObservedTurnFollowUps(
                        candidates,
                        observedTurnBatchLimit);
                    if (observedTurnActions.Length > 0)
                        return observedTurnActions;
                }

                var result = new List<ActionDto>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    ActionSignature(candidates[0])
                };
                var previous = candidates[0];

                if (!IsWellFormedSafeBatchedAction(previous))
                    return Array.Empty<ActionDto>();

                var gestureBatch = previous.Type == "drag_path";
                var gesturePoints = gestureBatch ? previous.Path?.Length ?? 0 : 0;
                var gestureDurationMs = gestureBatch ? previous.DurationMs ?? 0 : 0;
                if (gestureBatch &&
                    (!allowDrawGestureBatch || !IsSafeBatchedDrawGesture(previous)))
                    return Array.Empty<ActionDto>();

                foreach (var candidate in candidates.Skip(1))
                {
                    if (result.Count >= MaxQueuedBatchActions)
                        break;

                    if (!IsSafeBatchedAction(candidate))
                        break;

                    if (!IsWellFormedSafeBatchedAction(candidate))
                        break;

                    if (!seen.Add(ActionSignature(candidate)))
                        continue;

                    if (!CanBatchWithoutObservation(previous, candidate))
                        break;

                    if (gestureBatch)
                    {
                        if (!IsSafeBatchedDrawGesture(candidate))
                            break;
                        var nextPoints = gesturePoints + (candidate.Path?.Length ?? 0);
                        var nextDurationMs = gestureDurationMs + (candidate.DurationMs ?? 0);
                        if (nextPoints > MaxBatchedGesturePoints ||
                            nextDurationMs > MaxBatchedGestureDurationMs)
                        {
                            break;
                        }
                        gesturePoints = nextPoints;
                        gestureDurationMs = nextDurationMs;
                    }

                    result.Add(candidate);
                    previous = candidate;
                    if (candidate.Type == "wait")
                        break;
                }

                return result.ToArray();
            }

            internal static ActionDto[] SafeObservedTurnFollowUps(
                IReadOnlyList<ActionDto> candidates,
                int observedTurnBatchLimit)
            {
                if (candidates.Count <= 1 ||
                    observedTurnBatchLimit <= 0 ||
                    candidates[0].PlannedInputs is not { Length: >= 2 } plannedInputs ||
                    (candidates[0].PlanConfidence ?? candidates[0].Confidence ?? 0) <
                    ControlLoopService.TurnBasedTransitionTracker.MinimumStructuredPlanConfidence ||
                    !TryBindObservedTurnActionToPlannedInput(
                        candidates[0],
                        plannedInputs[0]))
                {
                    return Array.Empty<ActionDto>();
                }

                var limit = Math.Min(
                    observedTurnBatchLimit,
                    plannedInputs.Length - 1);
                var result = new List<ActionDto>(limit);
                var directionalClickPoints = new Dictionary<string, (int X, int Y)>(
                    StringComparer.Ordinal);
                if (!TryReconcileObservedDirectionalClick(
                        candidates[0],
                        plannedInputs[0],
                        directionalClickPoints))
                {
                    return Array.Empty<ActionDto>();
                }
                for (var index = 1; index < candidates.Count && result.Count < limit; index++)
                {
                    var candidate = candidates[index];
                    if (!TryReconcileObservedDirectionalClick(
                            candidate,
                            plannedInputs[index],
                            directionalClickPoints) ||
                        !TryBindObservedTurnActionToPlannedInput(
                            candidate,
                            plannedInputs[index]) ||
                        !CaptureResolvedAction(candidate, null).IsValid)
                    {
                        break;
                    }

                    result.Add(candidate);
                }

                return result.ToArray();
            }

            static bool TryReconcileObservedDirectionalClick(
                ActionDto action,
                string? plannedInput,
                IDictionary<string, (int X, int Y)> knownPoints)
            {
                if (action.Type is not ("click" or "double_click"))
                    return true;
                if (!TryNormalizeDirectionalLabel(plannedInput, out var plannedLabel) ||
                    !TryGetActionImagePoint(action, out var actionPoint))
                {
                    return false;
                }

                const int sameControlTolerance = 4;
                if (knownPoints.TryGetValue(plannedLabel, out var knownPoint))
                {
                    if (PointDistanceSquared(actionPoint, knownPoint) >
                        sameControlTolerance * sameControlTolerance)
                    {
                        RecenterClickAction(action, actionPoint, knownPoint);
                        Console.WriteLine(
                            $"[batch] corrected a repeated {plannedLabel} click to its previously observed control position.");
                    }
                    return true;
                }

                var conflictingPoint = knownPoints.FirstOrDefault(entry =>
                    !string.Equals(entry.Key, plannedLabel, StringComparison.Ordinal) &&
                    PointDistanceSquared(actionPoint, entry.Value) <=
                    sameControlTolerance * sameControlTolerance);
                if (!string.IsNullOrWhiteSpace(conflictingPoint.Key))
                {
                    if (!TryInferDirectionalPadPoint(
                            knownPoints,
                            plannedLabel,
                            out var inferredPoint))
                    {
                        Console.WriteLine(
                            $"[batch] rejected inconsistent directional click: {plannedLabel} reused the {conflictingPoint.Key} control position.");
                        return false;
                    }

                    RecenterClickAction(action, actionPoint, inferredPoint);
                    actionPoint = inferredPoint;
                    Console.WriteLine(
                        $"[batch] repaired inconsistent directional click geometry: {plannedLabel} had reused the {conflictingPoint.Key} control position.");
                }

                knownPoints[plannedLabel] = actionPoint;
                return true;
            }

            static bool TryGetActionImagePoint(
                ActionDto action,
                out (int X, int Y) point)
            {
                if (action.BBox is { Left: { } left, Top: { } top, Right: { } right, Bottom: { } bottom })
                {
                    point = ((left + right) / 2, (top + bottom) / 2);
                    return true;
                }
                if (action.XPx is int x && action.YPx is int y)
                {
                    point = (x, y);
                    return true;
                }
                if (action.X is double normalizedX && action.Y is double normalizedY)
                {
                    point = (
                        (int)Math.Round(normalizedX * Math.Max(1, CurrentScreenMap.ImageW - 1)),
                        (int)Math.Round(normalizedY * Math.Max(1, CurrentScreenMap.ImageH - 1)));
                    return true;
                }
                point = default;
                return false;
            }

            static long PointDistanceSquared(
                (int X, int Y) first,
                (int X, int Y) second)
            {
                var dx = (long)first.X - second.X;
                var dy = (long)first.Y - second.Y;
                return dx * dx + dy * dy;
            }

            static bool TryInferDirectionalPadPoint(
                IDictionary<string, (int X, int Y)> knownPoints,
                string targetLabel,
                out (int X, int Y) point)
            {
                var vertical = knownPoints.FirstOrDefault(entry =>
                    entry.Key is "ArrowUp" or "ArrowDown");
                var horizontal = knownPoints.FirstOrDefault(entry =>
                    entry.Key is "ArrowLeft" or "ArrowRight");
                if (string.IsNullOrWhiteSpace(vertical.Key) ||
                    string.IsNullOrWhiteSpace(horizontal.Key))
                {
                    point = default;
                    return false;
                }

                var centerX = vertical.Value.X;
                var centerY = horizontal.Value.Y;
                var verticalOffset = Math.Abs(vertical.Value.Y - centerY);
                var horizontalOffset = Math.Abs(horizontal.Value.X - centerX);
                var orientationMatches =
                    (vertical.Key == "ArrowUp" && vertical.Value.Y < centerY ||
                     vertical.Key == "ArrowDown" && vertical.Value.Y > centerY) &&
                    (horizontal.Key == "ArrowLeft" && horizontal.Value.X < centerX ||
                     horizontal.Key == "ArrowRight" && horizontal.Value.X > centerX);
                if (!orientationMatches ||
                    verticalOffset < 4 ||
                    horizontalOffset < 4 ||
                    verticalOffset > horizontalOffset * 2 ||
                    horizontalOffset > verticalOffset * 2)
                {
                    point = default;
                    return false;
                }

                point = targetLabel switch
                {
                    "ArrowUp" => (centerX, centerY - verticalOffset),
                    "ArrowDown" => (centerX, centerY + verticalOffset),
                    "ArrowLeft" => (centerX - horizontalOffset, centerY),
                    "ArrowRight" => (centerX + horizontalOffset, centerY),
                    _ => default
                };
                return targetLabel is "ArrowUp" or "ArrowDown" or "ArrowLeft" or "ArrowRight";
            }

            static void RecenterClickAction(
                ActionDto action,
                (int X, int Y) oldPoint,
                (int X, int Y) newPoint)
            {
                var dx = newPoint.X - oldPoint.X;
                var dy = newPoint.Y - oldPoint.Y;
                if (action.BBox is { Left: { } left, Top: { } top, Right: { } right, Bottom: { } bottom })
                {
                    action.BBox.Left = left + dx;
                    action.BBox.Top = top + dy;
                    action.BBox.Right = right + dx;
                    action.BBox.Bottom = bottom + dy;
                }
                action.XPx = newPoint.X;
                action.YPx = newPoint.Y;
            }

            internal static bool TryBindObservedTurnActionToPlannedInput(
                ActionDto action,
                string? plannedInput)
            {
                if (!TryNormalizeDirectionalLabel(plannedInput, out var plannedLabel))
                    return false;

                if (action.Type is "click" or "double_click")
                {
                    if (!MouseEnabled ||
                        !DirectClickWithoutAim ||
                        !HasExplicitPoint(action))
                    {
                        return false;
                    }

                    action.ResolvedTurnInputLabel = plannedLabel;
                    return true;
                }

                if (!TryGetObservedTurnInputLabel(action, out var actionLabel) ||
                    !DirectionalLabelsMatch(actionLabel, plannedLabel))
                {
                    return false;
                }

                action.ResolvedTurnInputLabel = plannedLabel;
                return true;
            }

            internal static bool TryGetObservedTurnInputLabel(
                ActionDto action,
                out string label)
            {
                label = "";
                if (TryNormalizeDirectionalLabel(
                        action.ResolvedTurnInputLabel,
                        out label))
                {
                    return true;
                }

                if (action.Type == "keys" && action.Keys is { Length: 1 })
                    return TryNormalizeDirectionalLabel(action.Keys[0], out label);

                if (action.Type is not ("click" or "double_click") ||
                    !MouseEnabled ||
                    !DirectClickWithoutAim ||
                    !HasExplicitPoint(action))
                {
                    return false;
                }

                return TryNormalizeDirectionalLabel(action.Note, out label);
            }

            internal static bool TryNormalizeDirectionalLabel(
                string? value,
                out string label)
            {
                var text = value ?? "";
                if (Regex.IsMatch(text, @"\b(arrowright|right|w prawo|praw\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    label = "ArrowRight";
                else if (Regex.IsMatch(text, @"\b(arrowleft|left|w lewo|lew\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    label = "ArrowLeft";
                else if (Regex.IsMatch(text, @"\b(arrowup|up|w gór\w*|gór\w*)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    label = "ArrowUp";
                else if (Regex.IsMatch(text, @"\b(arrowdown|down|w dół|dol\w*|dół)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    label = "ArrowDown";
                else
                {
                    label = "";
                    return false;
                }

                return true;
            }

            static bool DirectionalLabelsMatch(string left, string right) =>
                TryNormalizeDirectionalLabel(right, out var normalizedRight) &&
                string.Equals(left, normalizedRight, StringComparison.Ordinal);

            internal static bool IsWellFormedSafeBatchedAction(ActionDto action)
            {
                if (!CaptureResolvedAction(action, null).IsValid)
                    return false;

                return action.Type switch
                {
                    "keys" => action.Keys is { Length: > 0 } &&
                              action.Keys.All(key => !string.IsNullOrWhiteSpace(key)),
                    "type_text" or "paste_text" => !string.IsNullOrEmpty(action.Text),
                    "wait" => action.WaitSeconds is > 0,
                    "run_command" => !string.IsNullOrWhiteSpace(action.Command),
                    "open_url" => !string.IsNullOrWhiteSpace(action.Url),
                    "launch_app" => !string.IsNullOrWhiteSpace(action.App),
                    "drag_path" => IsSafeBatchedDrawGesture(action),
                    _ => false
                };
            }

            internal static bool IsSafeBatchedDrawGesture(ActionDto action) =>
                MouseEnabled &&
                DirectClickWithoutAim &&
                action.Type == "drag_path" &&
                string.Equals(action.GestureKind, "draw", StringComparison.OrdinalIgnoreCase) &&
                action.Path is { Length: >= 2 } &&
                action.DurationMs is >= 100 &&
                action.DurationMs <= MaxGestureDurationMs;

            internal static bool CanBatchWithoutObservation(
                ActionDto previous,
                ActionDto next)
            {
                if (!IsSafeBatchedAction(previous) || !IsSafeBatchedAction(next))
                    return false;
                if (previous.Type == "wait")
                    return false;
                if (next.Type == "wait")
                    return true;

                if (previous.Type == "drag_path" || next.Type == "drag_path")
                    return IsSafeBatchedDrawGesture(previous) &&
                           IsSafeBatchedDrawGesture(next);

                if (previous.Type is "open_url" or "launch_app" or "run_command" ||
                    IsCommitKeys(previous))
                {
                    return false;
                }

                if (next.Type is "type_text" or "paste_text")
                {
                    return previous.Type is "type_text" or "paste_text" ||
                           previous.Type == "keys" && IsTextEntryPrelude(previous);
                }

                if (previous.Type == "keys" && IsFieldAdvanceKeys(previous) &&
                    next.Type == "keys" && IsTextSelectionKeys(next))
                {
                    return true;
                }

                return next.Type == "keys" &&
                       previous.Type is "type_text" or "paste_text" &&
                       (IsCommitKeys(next) || IsFieldAdvanceKeys(next));
            }

            internal static bool IsCommitKeys(ActionDto action)
            {
                if (action.Type != "keys" || action.Keys is null)
                    return false;
                var signature = NormalizeBatchKeySignature(action.Keys);
                return signature is "enter" or "return";
            }

            internal static bool IsFieldAdvanceKeys(ActionDto action)
            {
                if (action.Type != "keys" || action.Keys is null)
                    return false;
                var signature = NormalizeBatchKeySignature(action.Keys);
                return signature is "tab" or "shift+tab";
            }

            internal static bool IsTextSelectionKeys(ActionDto action)
            {
                if (action.Type != "keys" || action.Keys is null)
                    return false;
                return NormalizeBatchKeySignature(action.Keys) == "ctrl+a";
            }

            internal static bool IsTextEntryPrelude(ActionDto action)
            {
                if (action.Type != "keys" || action.Keys is null)
                    return false;
                var signature = NormalizeBatchKeySignature(action.Keys);
                return signature is "win" or "win+r" or "ctrl+l" or "ctrl+k" or
                    "ctrl+e" or "ctrl+a" or "tab" or "shift+tab" or "f6";
            }

            static string NormalizeBatchKeySignature(IEnumerable<string> keys) =>
                string.Join("+", keys)
                    .Replace(" ", "", StringComparison.Ordinal)
                    .ToLowerInvariant();

            // helper – extract first balanced JSON object from a string
            internal static string? TryExtractJsonObject(string s)
            {
                if (string.IsNullOrEmpty(s)) return null;
                int start = s.IndexOf('{');
                if (start < 0) return null;
                int depth = 0;
                for (int i = start; i < s.Length; i++)
                {
                    char ch = s[i];
                    if (ch == '{') depth++;
                    else if (ch == '}')
                    {
                        depth--;
                        if (depth == 0)
                            return s[start..(i + 1)];
                    }
                }
                return null;
            }
    }
}

