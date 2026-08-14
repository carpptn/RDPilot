internal static partial class RDPilotApplication
{
    /// <summary>
    /// Analyzes recorded request, response, screenshot, and run-log artifacts.
    /// </summary>
    internal static partial class ArtifactAnalysisService
    {
        private static readonly string[] LogLineSeparators = ["\r\n", "\n"];

        [GeneratedRegex(@"^\[\d+\] ", RegexOptions.Multiline)]
        private static partial Regex StepEntryRegex();

        [GeneratedRegex("Goal NOT confirmed")]
        private static partial Regex RejectedGoalRegex();

        [GeneratedRegex("OpenAI HTTP")]
        private static partial Regex OpenAiHttpErrorRegex();

            internal static void AnalyzeArtifacts()
            {
                var baseDir = AppContext.BaseDirectory;
                var requestsDir = Path.Combine(baseDir, "requests");
                var screensDir = Path.Combine(baseDir, "screens");
                var logsDir = Path.Combine(baseDir, "logs");
        
                Console.WriteLine($"Artifact root: {baseDir}");
                AnalyzeResponses(requestsDir);
                AnalyzeRequestPayloads(requestsDir);
                AnalyzeScreens(screensDir);
                AnalyzeLogs(logsDir);
            }
        
            internal static void ReplayResponse(string responsePath)
            {
                var path = Path.GetFullPath(responsePath);
                if (!File.Exists(path))
                {
                    Console.Error.WriteLine($"Response file not found: {path}");
                    return;
                }
        
                var raw = File.ReadAllText(path);
                if (!raw.TrimStart().StartsWith('{'))
                {
                    Console.WriteLine($"Replay: non-JSON response file ({raw.Length} chars).");
                    Console.WriteLine(raw);
                    return;
                }
        
                using var doc = JsonDocument.Parse(raw);
                if (!TryParseControlActionSequence(
                        doc.RootElement,
                        out var candidates,
                        out var payloadCount,
                        out var legacyPayload))
                {
                    Console.WriteLine("Replay: no valid control-action sequence.");
                    return;
                }
                var first = candidates[0];
        
                Console.WriteLine($"Replay response: {path}");
                Console.WriteLine($"Payloads: {payloadCount}; legacy={legacyPayload}; proposed_actions={candidates.Count}");
                Console.WriteLine($"First: {Describe(first)}");
        
                var queued = SafeBatchFollowUps(candidates);
                Console.WriteLine($"Safe batch follow-ups: {queued.Length}");
                for (var i = 0; i < queued.Length; i++)
                    Console.WriteLine($"  {i + 1}. {Describe(queued[i])}");
            }
        
            internal static async Task ReplayRequestAsync(string? apiKey, string requestPath, bool dryRun)
            {
                var path = Path.GetFullPath(requestPath);
                if (!File.Exists(path))
                {
                    Console.Error.WriteLine($"Request file not found: {path}");
                    return;
                }
        
                JsonNode? root;
                try
                {
                    root = JsonNode.Parse(File.ReadAllText(path));
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Could not parse request JSON: {ex.Message}");
                    return;
                }
        
                if (root is not JsonObject rootObj)
                {
                    Console.Error.WriteLine("Replay request expects a JSON object.");
                    return;
                }
        
                var body = LoggedReplayBody(rootObj);
                if (body is null)
                {
                    Console.Error.WriteLine("Replay request supports logged request bodies with model/input/text fields. Older verifier wrapper logs are not replayable.");
                    return;
                }
        
                try
                {
                    var stats = HydrateReplayImages(body);
                    var bodyBytes = JsonSerializer.SerializeToUtf8Bytes(body);
                    Console.WriteLine($"Replay request: {path}");
                    Console.WriteLine($"Hydrated request: bytes={bodyBytes.Length}; images={stats.Images}; image_mb={stats.ImageBytes / 1024.0 / 1024.0:0.0}; model={body["model"]?.GetValue<string>() ?? "unknown"}");
        
                    if (dryRun)
                        return;
        
                    if (string.IsNullOrWhiteSpace(apiKey))
                    {
                        Console.Error.WriteLine("Missing API key for replay request.");
                        return;
                    }
        
                    using var cancelCts = StartCancelHotkeyListener();
                    ResetRunMetrics();
                    var (ok, statusCode, raw, elapsed, requestBytes) = await SendOpenAIRequestAsync(apiKey, body, cancelCts.Token);
                    RunOpenAiCalls++;
                    RunOpenAiElapsed += elapsed;
                    RunOpenAiRequestBytes += requestBytes;
                    RunOpenAiBytes += Encoding.UTF8.GetByteCount(raw);
                    if (ok)
                    {
                        try
                        {
                            using var responseDoc = JsonDocument.Parse(raw);
                            RecordUsageMetrics(responseDoc.RootElement);
                        }
                        catch { }
                    }
        
                    var outPath = Path.Combine(Path.GetDirectoryName(path) ?? Environment.CurrentDirectory,
                                               Path.GetFileNameWithoutExtension(path) + "_replay_response.json");
                    SaveRaw(outPath, raw);
                    Console.WriteLine($"Replay OpenAI: {(ok ? "ok" : "error")} status={statusCode} elapsed={elapsed.TotalSeconds:0.0}s response={outPath}");
                    PrintRunMetrics();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Replay request failed: {ex.Message}");
                }
            }
        
            internal static JsonObject? LoggedReplayBody(JsonObject root)
            {
                if (root.ContainsKey("model") && root.ContainsKey("input") && root.ContainsKey("text"))
                    return root;
        
                if (root["body"] is JsonObject body &&
                    body.ContainsKey("model") &&
                    body.ContainsKey("input") &&
                    body.ContainsKey("text"))
                    return body;
        
                return null;
            }
        
            internal static ReplayImageStats HydrateReplayImages(JsonNode node)
            {
                var stats = new ReplayImageStats();
                HydrateReplayImages(node, stats);
                return stats;
            }
        
            internal static void HydrateReplayImages(JsonNode? node, ReplayImageStats stats)
            {
                if (node is JsonObject obj)
                {
                    foreach (var kv in obj.ToList())
                    {
                        if (kv.Value is JsonValue value &&
                            kv.Key.Equals("image_url", StringComparison.OrdinalIgnoreCase) &&
                            value.TryGetValue<string>(out var imageRef) &&
                            imageRef.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                        {
                            var imagePath = imageRef["file://".Length..];
                            var dataUrl = ImageFileToDataUrl(imagePath, out var bytes);
                            obj[kv.Key] = dataUrl;
                            stats.Images++;
                            stats.ImageBytes += bytes;
                            continue;
                        }
        
                        HydrateReplayImages(kv.Value, stats);
                    }
                }
                else if (node is JsonArray arr)
                {
                    foreach (var item in arr)
                        HydrateReplayImages(item, stats);
                }
            }
        
            internal static string ImageFileToDataUrl(string path, out long bytes)
            {
                path = path.Trim();
                if (!File.Exists(path))
                    throw new FileNotFoundException($"Replay image not found: {path}", path);
        
                var raw = File.ReadAllBytes(path);
                bytes = raw.LongLength;
                var ext = Path.GetExtension(path).ToLowerInvariant();
                var mime = ext is ".jpg" or ".jpeg" ? "image/jpeg" : "image/png";
                return $"data:{mime};base64,{Convert.ToBase64String(raw)}";
            }
        
            internal static void AnalyzeResponses(string requestsDir)
            {
                if (!Directory.Exists(requestsDir))
                {
                    Console.WriteLine("[analysis] no requests directory");
                    return;
                }
        
                var files = Directory.GetFiles(requestsDir, "*_response.json");
                var calls = 0;
                var multi = 0;
                var inputTokens = 0L;
                var cachedTokens = 0L;
                var totalTokens = 0L;
                var reasoningTokens = 0L;
                var apiSeconds = 0L;
                var errors = 0;
                var slowCalls = new List<ResponseAnalysisRow>();
                var modelCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var kindCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        
                foreach (var file in files)
                {
                    try
                    {
                        var text = File.ReadAllText(file);
                        if (!text.TrimStart().StartsWith('{'))
                        {
                            errors++;
                            continue;
                        }
        
                        using var doc = JsonDocument.Parse(text);
                        var root = doc.RootElement;
                        if (!root.TryGetProperty("id", out _))
                            continue;
        
                        calls++;
                        var kind = ArtifactKindFromFileName(file);
                        kindCounts[kind] = kindCounts.TryGetValue(kind, out var kindCount) ? kindCount + 1 : 1;
                        var model = root.TryGetProperty("model", out var modelEl) ? modelEl.GetString() ?? "unknown" : "unknown";
                        modelCounts[model] = modelCounts.TryGetValue(model, out var modelCount) ? modelCount + 1 : 1;
        
                        var seconds = 0L;
                        if (root.TryGetProperty("created_at", out var created) && root.TryGetProperty("completed_at", out var completed))
                        {
                            seconds = completed.GetInt64() - created.GetInt64();
                            apiSeconds += seconds;
                        }
        
                        var responseTokens = 0L;
                        var responseInputTokens = 0L;
                        var responseCachedTokens = 0L;
                        var responseReasoningTokens = 0L;
                        if (root.TryGetProperty("usage", out var usage))
                        {
                            if (usage.TryGetProperty("input_tokens", out var input))
                            {
                                responseInputTokens = input.GetInt64();
                                inputTokens += responseInputTokens;
                            }
                            if (usage.TryGetProperty("input_tokens_details", out var inputDetails) &&
                                inputDetails.TryGetProperty("cached_tokens", out var cached))
                            {
                                responseCachedTokens = cached.GetInt64();
                                cachedTokens += responseCachedTokens;
                            }
                            if (usage.TryGetProperty("total_tokens", out var total))
                            {
                                responseTokens = total.GetInt64();
                                totalTokens += responseTokens;
                            }
                            if (usage.TryGetProperty("output_tokens_details", out var details) &&
                                details.TryGetProperty("reasoning_tokens", out var reasoning))
                            {
                                responseReasoningTokens = reasoning.GetInt64();
                                reasoningTokens += responseReasoningTokens;
                            }
                        }
        
                        var candidateCount = CountParsedActionCandidates(root);
                        if (candidateCount > 1)
                            multi++;
        
                        slowCalls.Add(new ResponseAnalysisRow(Path.GetFileName(file), kind, model, seconds, responseInputTokens, responseCachedTokens, responseTokens, responseReasoningTokens, candidateCount));
                    }
                    catch
                    {
                        errors++;
                    }
                }
        
                var avgSeconds = calls > 0 ? apiSeconds / (double)calls : 0;
                var avgTokens = calls > 0 ? totalTokens / (double)calls : 0;
                var cachedPct = inputTokens > 0 ? cachedTokens * 100.0 / inputTokens : 0.0;
                Console.WriteLine($"[analysis] responses={calls}; api_seconds={apiSeconds}; avg_seconds={avgSeconds:0.0}; input_tokens={inputTokens}; cached_tokens={cachedTokens}; cached_pct={cachedPct:0.0}; total_tokens={totalTokens}; avg_tokens={avgTokens:0}; reasoning_tokens={reasoningTokens}; multi_action_responses={multi}; error_files={errors}");
                if (kindCounts.Count > 0)
                    Console.WriteLine($"[analysis] response_kinds={string.Join(", ", kindCounts.OrderByDescending(kv => kv.Value).Select(kv => $"{kv.Key}:{kv.Value}"))}");
                if (modelCounts.Count > 0)
                    Console.WriteLine($"[analysis] models={string.Join(", ", modelCounts.OrderByDescending(kv => kv.Value).Select(kv => $"{kv.Key}:{kv.Value}"))}");
        
                foreach (var row in slowCalls.OrderByDescending(r => r.Seconds).ThenByDescending(r => r.TotalTokens).Take(5))
                    Console.WriteLine($"[analysis] slow_response {row.FileName}: {row.Seconds}s; kind={row.Kind}; model={row.Model}; input={row.InputTokens}; cached={row.CachedTokens}; tokens={row.TotalTokens}; reasoning={row.ReasoningTokens}; candidates={row.CandidateCount}");
            }
        
            internal static void AnalyzeRequestPayloads(string requestsDir)
            {
                if (!Directory.Exists(requestsDir))
                    return;
        
                var files = Directory.GetFiles(requestsDir, "*_request.json");
                if (files.Length == 0)
                    return;
        
                var count = 0;
                var errors = 0;
                var totalBytes = 0L;
                var totalTextChars = 0L;
                var totalImageRefs = 0;
                var rows = new List<RequestAnalysisRow>();
                var kindCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var modelCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var cacheKeyCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        
                foreach (var file in files)
                {
                    try
                    {
                        var info = new FileInfo(file);
                        var raw = File.ReadAllText(file);
                        if (!raw.TrimStart().StartsWith('{'))
                        {
                            errors++;
                            continue;
                        }
        
                        using var doc = JsonDocument.Parse(raw);
                        var body = UnwrapLoggedRequestBody(doc.RootElement);
                        var kind = ArtifactKindFromFileName(file);
                        var model = JsonString(body, "model") ?? "unknown";
                        var cacheKey = JsonString(body, "prompt_cache_key") ?? "none";
                        var effort = JsonString(body, "reasoning_effort") ?? JsonString(body, "reasoning", "effort") ?? "default";
                        var maxOutput = JsonInt(body, "max_output_tokens");
                        var textChars = 0L;
                        var imageRefs = 0;
                        CountPromptPayload(body, ref textChars, ref imageRefs);
        
                        count++;
                        totalBytes += info.Length;
                        totalTextChars += textChars;
                        totalImageRefs += imageRefs;
                        kindCounts[kind] = kindCounts.TryGetValue(kind, out var kindCount) ? kindCount + 1 : 1;
                        modelCounts[model] = modelCounts.TryGetValue(model, out var modelCount) ? modelCount + 1 : 1;
                        cacheKeyCounts[cacheKey] = cacheKeyCounts.TryGetValue(cacheKey, out var cacheCount) ? cacheCount + 1 : 1;
                        rows.Add(new RequestAnalysisRow(Path.GetFileName(file), kind, info.Length, model, textChars, imageRefs, cacheKey, maxOutput, effort));
                    }
                    catch
                    {
                        errors++;
                    }
                }
        
                var avgKb = count > 0 ? totalBytes / 1024.0 / count : 0.0;
                var avgTextChars = count > 0 ? totalTextChars / (double)count : 0.0;
                var avgImages = count > 0 ? totalImageRefs / (double)count : 0.0;
                Console.WriteLine($"[analysis] requests={count}; request_mb={totalBytes / 1024.0 / 1024.0:0.0}; avg_request_kb={avgKb:0.0}; text_chars={totalTextChars}; avg_text_chars={avgTextChars:0}; image_refs={totalImageRefs}; avg_images={avgImages:0.0}; error_files={errors}");
                if (kindCounts.Count > 0)
                    Console.WriteLine($"[analysis] request_kinds={string.Join(", ", kindCounts.OrderByDescending(kv => kv.Value).Select(kv => $"{kv.Key}:{kv.Value}"))}");
                if (modelCounts.Count > 0)
                    Console.WriteLine($"[analysis] request_models={string.Join(", ", modelCounts.OrderByDescending(kv => kv.Value).Select(kv => $"{kv.Key}:{kv.Value}"))}");
                if (cacheKeyCounts.Count > 0)
                    Console.WriteLine($"[analysis] request_cache_keys={string.Join(", ", cacheKeyCounts.OrderByDescending(kv => kv.Value).Take(5).Select(kv => $"{kv.Key}:{kv.Value}"))}");
        
                foreach (var row in rows.OrderByDescending(r => r.Bytes).Take(5))
                    Console.WriteLine($"[analysis] large_request {row.FileName}: {row.Bytes / 1024.0:0.0} KB; kind={row.Kind}; model={row.Model}; text_chars={row.TextChars}; images={row.ImageRefs}; max_output={row.MaxOutputTokens}; effort={row.Effort}; cache={row.CacheKey}");
            }
        
            internal static JsonElement UnwrapLoggedRequestBody(JsonElement root)
            {
                if (root.ValueKind == JsonValueKind.Object &&
                    root.TryGetProperty("body", out var body) &&
                    body.ValueKind == JsonValueKind.Object)
                    return body;
                return root;
            }
        
            internal static string ArtifactKindFromFileName(string path)
            {
                var name = Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
                if (name.EndsWith("_qa_request", StringComparison.Ordinal) || name.EndsWith("_qa_response", StringComparison.Ordinal))
                    return "qa";
                if (name.EndsWith("_verify_request", StringComparison.Ordinal) || name.EndsWith("_verify_response", StringComparison.Ordinal))
                    return "verify";
                return "control";
            }
        
            internal static string? JsonString(JsonElement obj, string name)
            {
                return obj.ValueKind == JsonValueKind.Object &&
                       obj.TryGetProperty(name, out var value) &&
                       value.ValueKind == JsonValueKind.String
                    ? value.GetString()
                    : null;
            }
        
            internal static string? JsonString(JsonElement obj, string parentName, string childName)
            {
                return obj.ValueKind == JsonValueKind.Object &&
                       obj.TryGetProperty(parentName, out var parent) &&
                       parent.ValueKind == JsonValueKind.Object
                    ? JsonString(parent, childName)
                    : null;
            }
        
            internal static int JsonInt(JsonElement obj, string name)
            {
                return obj.ValueKind == JsonValueKind.Object &&
                       obj.TryGetProperty(name, out var value) &&
                       value.ValueKind == JsonValueKind.Number &&
                       value.TryGetInt32(out var parsed)
                    ? parsed
                    : 0;
            }
        
            internal static void CountPromptPayload(JsonElement element, ref long textChars, ref int imageRefs)
            {
                switch (element.ValueKind)
                {
                    case JsonValueKind.Object:
                        foreach (var prop in element.EnumerateObject())
                        {
                            if ((prop.NameEquals("text") || prop.NameEquals("rules") || prop.NameEquals("meta")) &&
                                prop.Value.ValueKind == JsonValueKind.String)
                            {
                                textChars += prop.Value.GetString()?.Length ?? 0;
                                continue;
                            }
        
                            if ((prop.NameEquals("image_url") || prop.NameEquals("screenshot")) &&
                                prop.Value.ValueKind == JsonValueKind.String)
                            {
                                imageRefs++;
                                continue;
                            }
        
                            CountPromptPayload(prop.Value, ref textChars, ref imageRefs);
                        }
                        break;
        
                    case JsonValueKind.Array:
                        foreach (var item in element.EnumerateArray())
                            CountPromptPayload(item, ref textChars, ref imageRefs);
                        break;
                }
            }
        
            internal static int CountOutputTextItems(JsonElement root)
            {
                var count = 0;
                if (!root.TryGetProperty("output", out var output) || output.ValueKind != JsonValueKind.Array)
                    return count;
        
                foreach (var item in output.EnumerateArray())
                    if (item.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array)
                        count += content.EnumerateArray().Count(c => c.TryGetProperty("type", out var t) && t.GetString() == "output_text");
        
                return count;
            }
        
            internal static int CountParsedActionCandidates(JsonElement root)
            {
                return TryParseControlActionSequence(root, out var actions, out _, out _)
                    ? actions.Count
                    : CountOutputTextItems(root);
            }
        
            internal static void AnalyzeScreens(string screensDir)
            {
                if (!Directory.Exists(screensDir))
                {
                    Console.WriteLine("[analysis] no screens directory");
                    return;
                }
        
                var files = Directory.GetFiles(screensDir, "*.*")
                    .Where(IsScreenImageArtifact)
                    .ToArray();
                var totalBytes = files.Sum(f => new FileInfo(f).Length);
                var full = files.Count(f => !Path.GetFileNameWithoutExtension(f).Contains("_crop") &&
                                            !Path.GetFileNameWithoutExtension(f).Contains("_focus_uia") &&
                                            !Path.GetFileNameWithoutExtension(f).Contains("_aim_overlay"));
                var overlays = files.Count(f => Path.GetFileNameWithoutExtension(f).Contains("_aim_overlay"));
                var focus = files.Count(f => Path.GetFileNameWithoutExtension(f).Contains("_focus_uia"));
                Console.WriteLine($"[analysis] screens={files.Length}; full={full}; focus_uia={focus}; aim_overlays={overlays}; total_mb={totalBytes / 1024.0 / 1024.0:0.0}");
        
                foreach (var file in files.Select(f => new FileInfo(f)).OrderByDescending(f => f.Length).Take(5))
                    Console.WriteLine($"[analysis] large_screen {file.Name}: {file.Length / 1024.0 / 1024.0:0.00} MB");
            }
        
            internal static void AnalyzeLogs(string logsDir)
            {
                if (!Directory.Exists(logsDir))
                {
                    Console.WriteLine("[analysis] no logs directory");
                    return;
                }
        
                var logs = Directory.GetFiles(logsDir, "*.log");
                var steps = 0;
                var rejected = 0;
                var httpErrors = 0;
                var metricRuns = 0;
                var metricOpenAiCalls = 0L;
                var metricOpenAiRetries = 0L;
                var metricOpenAiSeconds = 0.0;
                var metricRequestBytes = 0L;
                var metricResponseBytes = 0L;
                var metricInputTokens = 0L;
                var metricCachedTokens = 0L;
                var metricReasoningTokens = 0L;
                var metricScreenshots = 0L;
                var metricScreenshotSeconds = 0.0;
                var metricScreenProbes = 0L;
                var metricScreenProbeSeconds = 0.0;
                var metricScreenSanityWarnings = 0L;
                var metricImageEncodes = 0L;
                var metricImageEncodeSeconds = 0.0;
                var metricScreenLogs = 0L;
                var metricScreenLogSeconds = 0.0;
                var metricArtifactLogs = 0L;
                var metricArtifactLogSeconds = 0.0;
                var metricUiaCalls = 0L;
                var metricUiaSeconds = 0.0;
                var metricLocalActions = 0L;
                var metricLocalActionSeconds = 0.0;
        
                foreach (var file in logs)
                {
                    var text = File.ReadAllText(file);
                    steps += StepEntryRegex().Matches(text).Count;
                    rejected += RejectedGoalRegex().Matches(text).Count;
                    httpErrors += OpenAiHttpErrorRegex().Matches(text).Count;

                    foreach (var line in text.Split(LogLineSeparators, StringSplitOptions.None).Where(l => l.StartsWith("[metrics]", StringComparison.Ordinal)))
                    {
                        var metrics = ParseMetricsLine(line);
                        metricRuns++;
                        metricOpenAiCalls += MetricLong(metrics, "openai_calls");
                        metricOpenAiRetries += MetricLong(metrics, "openai_retries");
                        metricOpenAiSeconds += MetricSeconds(metrics, "openai_time");
                        metricRequestBytes += MetricLong(metrics, "request_bytes");
                        metricResponseBytes += MetricLong(metrics, "response_bytes");
                        metricInputTokens += MetricLong(metrics, "input_tokens");
                        metricCachedTokens += MetricLong(metrics, "cached_tokens");
                        metricReasoningTokens += MetricLong(metrics, "reasoning_tokens");
                        metricScreenshots += MetricLong(metrics, "screenshots");
                        metricScreenshotSeconds += MetricSeconds(metrics, "screenshot_time");
                        metricScreenProbes += MetricLong(metrics, "screen_probes");
                        metricScreenProbeSeconds += MetricSeconds(metrics, "screen_probe_time");
                        metricScreenSanityWarnings += MetricLong(metrics, "screen_sanity_warnings");
                        metricImageEncodes += MetricLong(metrics, "image_encodes");
                        metricImageEncodeSeconds += MetricSeconds(metrics, "image_encode_time");
                        metricScreenLogs += MetricLong(metrics, "screen_logs");
                        metricScreenLogSeconds += MetricSeconds(metrics, "screen_log_time");
                        metricArtifactLogs += MetricLong(metrics, "artifact_logs");
                        metricArtifactLogSeconds += MetricSeconds(metrics, "artifact_log_time");
                        metricUiaCalls += MetricLong(metrics, "uia_calls");
                        metricUiaSeconds += MetricSeconds(metrics, "uia_time");
                        metricLocalActions += MetricLong(metrics, "local_actions");
                        metricLocalActionSeconds += MetricSeconds(metrics, "local_action_time");
                    }
                }
                Console.WriteLine($"[analysis] logs={logs.Length}; steps={steps}; rejected_done={rejected}; http_errors={httpErrors}");
                if (metricRuns > 0)
                {
                    var metricCachedPct = metricInputTokens > 0 ? metricCachedTokens * 100.0 / metricInputTokens : 0.0;
                    Console.WriteLine($"[analysis] runtime_metrics runs={metricRuns}; openai_calls={metricOpenAiCalls}; openai_retries={metricOpenAiRetries}; openai_time={metricOpenAiSeconds:0.0}s; request_mb={metricRequestBytes / 1024.0 / 1024.0:0.0}; response_mb={metricResponseBytes / 1024.0 / 1024.0:0.0}; input_tokens={metricInputTokens}; cached_tokens={metricCachedTokens}; cached_pct={metricCachedPct:0.0}; reasoning_tokens={metricReasoningTokens}; screenshots={metricScreenshots}; screenshot_time={metricScreenshotSeconds:0.0}s; screen_probes={metricScreenProbes}; screen_probe_time={metricScreenProbeSeconds:0.0}s; screen_sanity_warnings={metricScreenSanityWarnings}; image_encodes={metricImageEncodes}; image_encode_time={metricImageEncodeSeconds:0.0}s; screen_logs={metricScreenLogs}; screen_log_time={metricScreenLogSeconds:0.0}s; artifact_logs={metricArtifactLogs}; artifact_log_time={metricArtifactLogSeconds:0.0}s; uia_calls={metricUiaCalls}; uia_time={metricUiaSeconds:0.0}s; local_actions={metricLocalActions}; local_action_time={metricLocalActionSeconds:0.0}s");
                }
            }
        
            internal static Dictionary<string, string> ParseMetricsLine(string line)
            {
                var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var close = line.IndexOf(']');
                var payload = close >= 0 ? line[(close + 1)..] : line;
                foreach (var part in payload.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    var eq = part.IndexOf('=');
                    if (eq <= 0)
                        continue;
                    result[part[..eq].Trim()] = part[(eq + 1)..].Trim();
                }
                return result;
            }
        
            internal static long MetricLong(Dictionary<string, string> metrics, string key)
                => metrics.TryGetValue(key, out var value) && long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                    ? parsed
                    : 0L;
        
            internal static double MetricSeconds(Dictionary<string, string> metrics, string key)
            {
                if (!metrics.TryGetValue(key, out var value))
                    return 0.0;
                value = value.EndsWith("s", StringComparison.OrdinalIgnoreCase) ? value[..^1] : value;
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
                    return parsed;
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
                    return parsed;
                return double.TryParse(value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) ? parsed : 0.0;
            }
    }
}

