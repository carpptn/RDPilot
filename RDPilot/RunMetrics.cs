internal static partial class RDPilotApplication
{
    /// <summary>
    /// Collects and reports metrics for a single agent run.
    /// </summary>
    internal static class RunMetrics
    {
            internal static void ResetRunMetrics()
            {
                RunOpenAiCalls = 0;
                RunOpenAiRetries = 0;
                RunOpenAiElapsed = TimeSpan.Zero;
                RunOpenAiRequestBytes = 0;
                RunOpenAiBytes = 0;
                RunMultiCandidateResponses = 0;
                RunWebSearchCalls = 0;
                RunInputTokens = 0;
                RunCachedTokens = 0;
                RunOutputTokens = 0;
                RunReasoningTokens = 0;
                RunEarlyAcceptedControlStreams = 0;
                RunControlContextTurns = 0;
                RunControlContextRestarts = 0;
                RunControlContextFallbacks = 0;
                RunControlContextCompactions = 0;
                RunControlCompactionFallbacks = 0;
                RunScreenshotCount = 0;
                RunScreenshotElapsed = TimeSpan.Zero;
                RunScreenProbeCount = 0;
                RunScreenProbeElapsed = TimeSpan.Zero;
                RunScreenSanityWarnings = 0;
                RunImageEncodeCount = 0;
                RunImageEncodeElapsed = TimeSpan.Zero;
                RunScreenLogCount = 0;
                RunScreenLogElapsed = TimeSpan.Zero;
                RunArtifactLogWrites = 0;
                RunArtifactLogElapsed = TimeSpan.Zero;
                RunUiaCalls = 0;
                RunUiaElapsed = TimeSpan.Zero;
                RunLocalActions = 0;
                RunLocalActionElapsed = TimeSpan.Zero;
                ReportedSanityWarnings.Clear();
                ReportedWebSearchCallIds.Clear();
            }
        
            internal static void PrintRunMetrics()
            {
                if (RunOpenAiCalls <= 0)
                    return;
        
                var cachedPct = RunInputTokens > 0 ? RunCachedTokens * 100.0 / RunInputTokens : 0.0;
                Console.WriteLine(FormattableString.Invariant($"[metrics] openai_calls={RunOpenAiCalls}; openai_retries={RunOpenAiRetries}; openai_time={RunOpenAiElapsed.TotalSeconds:0.0}s; request_bytes={RunOpenAiRequestBytes}; response_bytes={RunOpenAiBytes}; input_tokens={RunInputTokens}; cached_tokens={RunCachedTokens}; cached_pct={cachedPct:0.0}; output_tokens={RunOutputTokens}; reasoning_tokens={RunReasoningTokens}; web_search_calls={RunWebSearchCalls}; early_control_streams={RunEarlyAcceptedControlStreams}; context_turns={RunControlContextTurns}; context_restarts={RunControlContextRestarts}; context_fallbacks={RunControlContextFallbacks}; context_compactions={RunControlContextCompactions}; compaction_fallbacks={RunControlCompactionFallbacks}; multi_candidate_responses={RunMultiCandidateResponses}; screenshots={RunScreenshotCount}; screenshot_time={RunScreenshotElapsed.TotalSeconds:0.0}s; screen_probes={RunScreenProbeCount}; screen_probe_time={RunScreenProbeElapsed.TotalSeconds:0.0}s; screen_sanity_warnings={RunScreenSanityWarnings}; image_encodes={RunImageEncodeCount}; image_encode_time={RunImageEncodeElapsed.TotalSeconds:0.0}s; screen_logs={RunScreenLogCount}; screen_log_time={RunScreenLogElapsed.TotalSeconds:0.0}s; artifact_logs={RunArtifactLogWrites}; artifact_log_time={RunArtifactLogElapsed.TotalSeconds:0.0}s; uia_calls={RunUiaCalls}; uia_time={RunUiaElapsed.TotalSeconds:0.0}s; local_actions={RunLocalActions}; local_action_time={RunLocalActionElapsed.TotalSeconds:0.0}s"));
                if (RunEarlyAcceptedControlStreams > 0)
                    Console.WriteLine($"[metrics] token_usage_partial=true; {RunEarlyAcceptedControlStreams} control response(s) were accepted before the final usage event");
                if (ShouldWarnPromptCache(UsePromptCache, RunOpenAiCalls, RunInputTokens, RunCachedTokens, RunEarlyAcceptedControlStreams))
                    Console.WriteLine("[metrics] prompt_cache_warning=cached_tokens_zero; check model support, prompt stability, and prompt_cache_key scope");
            }

            internal static bool ShouldWarnPromptCache(bool enabled, int calls, long inputTokens, long cachedTokens, int earlyAcceptedControlStreams)
            {
                return enabled && calls > 1 && inputTokens > 0 && cachedTokens == 0 && earlyAcceptedControlStreams == 0;
            }
        
            internal static void RecordUiaMetric(Stopwatch sw)
            {
                sw.Stop();
                RunUiaCalls++;
                RunUiaElapsed += sw.Elapsed;
            }
    }
}

