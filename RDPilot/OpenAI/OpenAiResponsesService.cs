
namespace RDPilot.OpenAI;

    /// <summary>
    /// Communicates with the OpenAI Responses API and parses structured model output.
    /// </summary>
    internal static class OpenAiResponsesService
    {
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
                    if (root.ValueKind != JsonValueKind.Object) return false;
                    var eventType = root.TryGetProperty("type", out var typeElement) &&
                                    typeElement.ValueKind == JsonValueKind.String
                        ? typeElement.GetString()
                        : null;

                    if (eventType == "response.created" &&
                        root.TryGetProperty("response", out var createdResponse) &&
                        createdResponse.ValueKind == JsonValueKind.Object &&
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

                    if (eventType == "response.web_search_call.searching")
                    {
                        var itemId = root.TryGetProperty("item_id", out var itemIdElement) &&
                                     itemIdElement.ValueKind == JsonValueKind.String
                            ? itemIdElement.GetString()
                            : null;
                        var firstEventForCall = itemId is null || CurrentRun.ReportedWebSearchCallIds.Add(itemId);
                        if (firstEventForCall)
                        {
                            CurrentRun.RunWebSearchCalls++;
                            Console.WriteLine("[web] search started");
                        }
                        return false;
                    }

                    if (eventType == "error")
                    {
                        completed = true;
                        raw = JsonSerializer.Serialize(new { status = "failed", error = root });
                        return true;
                    }
                    if (eventType is "response.completed" or "response.incomplete" or "response.failed")
                    {
                        completed = true;
                        if (root.TryGetProperty("response", out var finalResponse))
                        {
                            raw = finalResponse.GetRawText();
                            if (finalResponse.ValueKind == JsonValueKind.Object && finalResponse.TryGetProperty("id", out var finalId) &&
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
                    CurrentRun.RunInputTokens += inputTokens;
                if (usage.TryGetProperty("output_tokens", out var output) && output.TryGetInt64(out var outputTokens))
                    CurrentRun.RunOutputTokens += outputTokens;
                if (usage.TryGetProperty("input_tokens_details", out var inputDetails) &&
                    inputDetails.TryGetProperty("cached_tokens", out var cached) &&
                    cached.TryGetInt64(out var cachedTokens))
                    CurrentRun.RunCachedTokens += cachedTokens;
                if (usage.TryGetProperty("output_tokens_details", out var outputDetails) &&
                    outputDetails.TryGetProperty("reasoning_tokens", out var reasoning) &&
                    reasoning.TryGetInt64(out var reasoningTokens))
                    CurrentRun.RunReasoningTokens += reasoningTokens;
            }

            internal static void RecordWebSearchMetrics(JsonElement root)
            {
                if (!root.TryGetProperty("output", out var output) ||
                    output.ValueKind != JsonValueKind.Array)
                {
                    return;
                }

                foreach (var item in output.EnumerateArray())
                {
                    if (!item.TryGetProperty("type", out var typeElement) ||
                        typeElement.ValueKind != JsonValueKind.String ||
                        !string.Equals(typeElement.GetString(), "web_search_call", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var itemId = item.TryGetProperty("id", out var idElement) &&
                                 idElement.ValueKind == JsonValueKind.String
                        ? idElement.GetString()
                        : null;
                    if (itemId is null || CurrentRun.ReportedWebSearchCallIds.Add(itemId))
                        CurrentRun.RunWebSearchCalls++;

                    var detail = "completed";
                    if (item.TryGetProperty("action", out var action) &&
                        action.ValueKind == JsonValueKind.Object)
                    {
                        if (action.TryGetProperty("query", out var query) &&
                            query.ValueKind == JsonValueKind.String)
                        {
                            detail = $"query={query.GetString()}";
                        }
                        else if (action.TryGetProperty("queries", out var queries) &&
                                 queries.ValueKind == JsonValueKind.Array)
                        {
                            var queryList = queries.EnumerateArray()
                                .Where(value => value.ValueKind == JsonValueKind.String)
                                .Select(value => value.GetString())
                                .Where(value => !string.IsNullOrWhiteSpace(value));
                            detail = $"queries={string.Join(" | ", queryList)}";
                        }
                        else if (action.TryGetProperty("type", out var actionType) &&
                                 actionType.ValueKind == JsonValueKind.String)
                        {
                            detail = $"action={actionType.GetString()}";
                        }
                    }
                    Console.WriteLine($"[web] {detail}");
                }
            }

            internal static void RecordResponseId(JsonElement root)
            {
                if (IsCompletedResponse(root) &&
                    root.TryGetProperty("id", out var id) &&
                    id.ValueKind == JsonValueKind.String)
                {
                    CurrentRun.LastOpenAiResponseId = id.GetString();
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
                root.ValueKind == JsonValueKind.Object &&
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

                var currentMax = Options.MaxOutputTokens;
                if (node.TryGetPropertyValue("max_output_tokens", out var maxNode) &&
                    maxNode is not null &&
                    maxNode.GetValueKind() == JsonValueKind.Number &&
                    maxNode.GetValue<int>() > 0)
                {
                    currentMax = maxNode.GetValue<int>();
                }

                var cap = Math.Max(currentMax, Options.IncompleteMaxOutputTokenCap);
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
                    sameEffortRetryCount < Math.Max(0, Options.IncompleteMaxOutputRetries) &&
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

            internal static bool TryBuildStandaloneRetry(object request, object? standalone, out object retry)
            {
                var node = JsonSerializer.SerializeToNode(request) as JsonObject;
                retry = request;
                if (node is null || !node.Remove("previous_response_id")) return false;
                if (standalone is not null)
                {
                    var checkpoint = JsonSerializer.SerializeToNode(standalone) as JsonObject;
                    node["input"] = checkpoint?["input"]?.DeepClone();
                }
                // A visual control retry must not lose both the server context and its screen evidence.
                if (!ContainsInputImage(node["input"])) return false;
                retry = node;
                return true;
            }

            private static bool ContainsInputImage(JsonNode? node) => node switch
            {
                JsonObject obj => (obj["type"] is JsonValue type && type.TryGetValue<string>(out var kind) && kind == "input_image" &&
                    obj["image_url"] is JsonValue url && url.TryGetValue<string>(out var image) && !string.IsNullOrWhiteSpace(image)) ||
                    obj.Any(pair => ContainsInputImage(pair.Value)),
                JsonArray array => array.Any(ContainsInputImage),
                _ => false
            };

            internal static async Task<(List<ActionDto>? parsed, string raw, string? responseId, bool contextFallbackUsed, bool compactionFallbackUsed, bool compactionOccurred)> CallOpenAIAsync(
                string apiKey,
                object body,
                CancellationToken cancellationToken = default,
                object? standaloneRequest = null)
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
                    CurrentRun.RunOpenAiCalls++;
                    CurrentRun.RunOpenAiElapsed += elapsed;
                    CurrentRun.RunOpenAiRequestBytes += requestBytes;
                    CurrentRun.RunOpenAiBytes += responseBytes;
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
                            CurrentRun.RunOpenAiRetries++;
                            CurrentRun.RunControlCompactionFallbacks++;
                            compactionFallbackUsed = true;
                            requestBody = requestWithoutCompaction;
                            Console.WriteLine("[context] server-side compaction was rejected by the API; retrying this turn without context_management.");
                            continue;
                        }

                        if (!contextFallbackUsed &&
                            IsPreviousResponseStateFailure(statusCode, raw) &&
                            TryBuildStandaloneRetry(requestBody, standaloneRequest, out var requestWithoutPreviousResponse))
                        {
                            CurrentRun.RunOpenAiRetries++;
                            CurrentRun.RunControlContextFallbacks++;
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
                        CurrentRun.RunEarlyAcceptedControlStreams++;
                        if (acceptedActions.Count > 1)
                        {
                            CurrentRun.RunMultiCandidateResponses++;
                            Console.WriteLine($"[openai] stream accepted a {acceptedActions.Count}-action sequence; evaluating guarded follow-ups.");
                        }
                        return (acceptedActions, raw, responseId, contextFallbackUsed, compactionFallbackUsed, false);
                    }

                    try
                    {
                        using var doc = JsonDocument.Parse(raw);
                        var root = doc.RootElement;
                        RecordResponseId(root);
                        RecordUsageMetrics(root);
                        RecordWebSearchMetrics(root);
                        var completedResponseId = IsCompletedResponse(root)
                            ? responseId ?? TryGetCompletedResponseId(raw)
                            : null;
                        var compactionOccurred = IsCompletedResponse(root) &&
                                                 ContainsCompactionItem(root);
                        if (compactionOccurred)
                            CurrentRun.RunControlContextCompactions++;

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
                                CurrentRun.RunMultiCandidateResponses++;
                                Console.WriteLine(
                                    Options.ExecuteMultiActionCandidates
                                        ? $"[openai] response proposed a {validCandidates.Count}-action sequence; evaluating guarded follow-ups."
                                        : $"[openai] response proposed a {validCandidates.Count}-action sequence; batching is disabled, using the first action.");
                            }
                            return (validCandidates, raw, completedResponseId, contextFallbackUsed, compactionFallbackUsed, compactionOccurred);
                        }

                        if (IsIncompleteDueToMaxOutput(root) &&
                            TryBuildMaxOutputRetryBody(requestBody, out var retryBody, out var retryMaxTokens, out var retryEffort, out var effortChanged, sameEffortRetryCount))
                        {
                            CurrentRun.RunOpenAiRetries++;
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
                    CurrentRun.RunOpenAiCalls++;
                    CurrentRun.RunOpenAiElapsed += elapsed;
                    CurrentRun.RunOpenAiRequestBytes += requestBytes;
                    var responseBytes = Encoding.UTF8.GetByteCount(raw);
                    CurrentRun.RunOpenAiBytes += responseBytes;
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
                        RecordWebSearchMetrics(root);

                        if (IsCompletedResponse(root) && TryParseResponsePayload<T>(root, out var parsed, out var candidates, out _))
                        {
                            if (candidates > 1)
                            {
                                CurrentRun.RunMultiCandidateResponses++;
                                Console.WriteLine($"[openai] warning: response contained {candidates} JSON candidates; using the first.");
                            }
                            return (parsed, raw);
                        }

                        if (IsIncompleteDueToMaxOutput(root) &&
                            TryBuildMaxOutputRetryBody(requestBody, out var retryBody, out var retryMaxTokens, out var retryEffort, out var effortChanged, sameEffortRetryCount))
                        {
                            CurrentRun.RunOpenAiRetries++;
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

            // helper – extract first balanced JSON object from a string
            internal static string? TryExtractJsonObject(string s)
            {
                if (string.IsNullOrEmpty(s)) return null;
                int start = s.IndexOf('{');
                if (start < 0) return null;
                // JsonDocument observes strings, escapes and nested arrays while consuming exactly one value.
                try
                {
                    var reader = new Utf8JsonReader(Encoding.UTF8.GetBytes(s[start..]));
                    using var document = JsonDocument.ParseValue(ref reader);
                    return document.RootElement.GetRawText();
                }
                catch (JsonException) { return null; }
            }
    }
