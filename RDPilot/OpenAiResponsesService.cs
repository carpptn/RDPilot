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
                if (root.TryGetProperty("id", out var id) && id.ValueKind == JsonValueKind.String)
                    LastOpenAiResponseId = id.GetString();
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
        
            internal static object BuildMaxOutputRetryBody(object body, out int retryMaxTokens, out string retryEffort)
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
        
                retryMaxTokens = currentMax < 1000
                    ? Math.Max(1200, currentMax * 4)
                    : Math.Max(currentMax + 1000, currentMax * 3);
                retryMaxTokens = Math.Min(retryMaxTokens, Math.Max(currentMax, IncompleteMaxOutputTokenCap));
                node["max_output_tokens"] = retryMaxTokens;
        
                retryEffort = "unchanged";
                if (node.TryGetPropertyValue("reasoning", out var reasoningNode) &&
                    reasoningNode is JsonObject reasoning &&
                    reasoning.TryGetPropertyValue("effort", out var effortNode) &&
                    effortNode is not null &&
                    effortNode.GetValueKind() == JsonValueKind.String)
                {
                    var effort = effortNode.GetValue<string>();
                    if (effort is "medium" or "high" or "xhigh")
                    {
                        reasoning["effort"] = "low";
                        retryEffort = "low";
                    }
                    else
                    {
                        retryEffort = effort;
                    }
                }
        
                return node;
            }
        
            internal static async Task<(ActionDto? parsed, string raw)> CallOpenAIAsync(string apiKey, object body, CancellationToken cancellationToken = default)
            {
                var requestBody = body;
                for (var retry = 0; retry <= IncompleteMaxOutputRetries; retry++)
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
                        return (null, raw);
                    }
        
                    try
                    {
                        using var doc = JsonDocument.Parse(raw);
                        var root = doc.RootElement;
                        RecordResponseId(root);
                        RecordUsageMetrics(root);
        
                        if (TryParseResponsePayload<ActionDto>(root, out var parsed, out var candidates, out var parsedCandidates))
                        {
                            var validCandidates = parsedCandidates.Where(IsKnownAction).ToList();
                            if (validCandidates.Count == 0)
                            {
                                Console.WriteLine("No valid action JSON found (parsed/output_text). RAW:");
                                Console.WriteLine(raw);
                                return (null, raw);
                            }
        
                            if (validCandidates.Count > 1)
                            {
                                RunMultiCandidateResponses++;
                                Console.WriteLine($"[openai] warning: response contained {validCandidates.Count} valid action candidates; using the first.");
                                QueueSafeCandidates(validCandidates);
                            }
                            return (validCandidates[0], raw);
                        }
        
                        if (IsIncompleteDueToMaxOutput(root) && retry < IncompleteMaxOutputRetries)
                        {
                            RunOpenAiRetries++;
                            requestBody = BuildMaxOutputRetryBody(requestBody, out var retryMaxTokens, out var retryEffort);
                            Console.WriteLine($"[openai] incomplete=max_output_tokens; retry {retry + 1}/{IncompleteMaxOutputRetries} with max_output_tokens={retryMaxTokens}, reasoning_effort={retryEffort}");
                            continue;
                        }
        
                        if (IsIncompleteDueToMaxOutput(root))
                        {
                            Console.WriteLine($"[openai] incomplete=max_output_tokens after {retry + 1} attempt(s); no complete JSON action was emitted. Increase --incomplete-max-output-token-cap or lower --max-action-text-chars if this repeats.");
                            return (null, raw);
                        }
        
                        Console.WriteLine("No parsable JSON found (parsed/output_text). RAW:");
                        Console.WriteLine(raw);
                        return (null, raw);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Response parsing error: {ex.Message}");
                        Console.WriteLine(raw);
                        return (null, raw);
                    }
                }
        
                return (null, "");
            }
        
            internal static async Task<(T? parsed, string raw)> CallOpenAIParsedAsync<T>(string apiKey, object body, CancellationToken cancellationToken = default)
            {
                var requestBody = body;
                for (var retry = 0; retry <= IncompleteMaxOutputRetries; retry++)
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
        
                        if (IsIncompleteDueToMaxOutput(root) && retry < IncompleteMaxOutputRetries)
                        {
                            RunOpenAiRetries++;
                            requestBody = BuildMaxOutputRetryBody(requestBody, out var retryMaxTokens, out var retryEffort);
                            Console.WriteLine($"[openai] incomplete=max_output_tokens; retry {retry + 1}/{IncompleteMaxOutputRetries} with max_output_tokens={retryMaxTokens}, reasoning_effort={retryEffort}");
                            continue;
                        }
        
                        if (IsIncompleteDueToMaxOutput(root))
                        {
                            Console.WriteLine($"[openai] incomplete=max_output_tokens after {retry + 1} attempt(s); no complete JSON payload was emitted. Increase --incomplete-max-output-token-cap or lower --max-action-text-chars if this repeats.");
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
        
                return (default, "");
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
        
            internal static void QueueSafeCandidates(List<ActionDto> candidates)
            {
                if (!ExecuteMultiActionCandidates || MaxQueuedBatchActions <= 0 || candidates.Count <= 1)
                    return;
        
                var queuedActions = SafeBatchFollowUps(candidates);
                foreach (var candidate in queuedActions)
                    PendingSafeActions.Enqueue(candidate);
        
                if (queuedActions.Length > 0)
                    Console.WriteLine($"[batch] queued {queuedActions.Length} safe follow-up action(s) from the same response.");
            }
        
            internal static ActionDto[] SafeBatchFollowUps(List<ActionDto> candidates)
            {
                if (MaxQueuedBatchActions <= 0 || candidates.Count <= 1)
                    return Array.Empty<ActionDto>();
        
                var result = new List<ActionDto>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    ActionSignature(candidates[0])
                };
        
                foreach (var candidate in candidates.Skip(1))
                {
                    if (result.Count >= MaxQueuedBatchActions)
                        break;
        
                    if (!IsSafeBatchedAction(candidate))
                        break;
        
                    if (!seen.Add(ActionSignature(candidate)))
                        continue;
        
                    result.Add(candidate);
                }
        
                return result.ToArray();
            }
        
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

