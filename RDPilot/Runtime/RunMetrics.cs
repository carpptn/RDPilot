
namespace RDPilot.Runtime;

    /// <summary>
    /// Collects and reports metrics for a single agent run.
    /// </summary>
    internal static class RunMetrics
    {
            internal static void ResetRunMetrics()
            {
                CurrentRun.RunOpenAiCalls = 0;
                CurrentRun.RunOpenAiRetries = 0;
                CurrentRun.RunOpenAiElapsed = TimeSpan.Zero;
                CurrentRun.RunOpenAiRequestBytes = 0;
                CurrentRun.RunOpenAiBytes = 0;
                CurrentRun.RunMultiCandidateResponses = 0;
                CurrentRun.RunWebSearchCalls = 0;
                CurrentRun.RunInputTokens = 0;
                CurrentRun.RunCachedTokens = 0;
                CurrentRun.RunOutputTokens = 0;
                CurrentRun.RunReasoningTokens = 0;
                CurrentRun.RunEarlyAcceptedControlStreams = 0;
                CurrentRun.RunControlContextTurns = 0;
                CurrentRun.RunControlContextRestarts = 0;
                CurrentRun.RunControlContextFallbacks = 0;
                CurrentRun.RunControlContextCompactions = 0;
                CurrentRun.RunControlCompactionFallbacks = 0;
                CurrentRun.RunScreenshotCount = 0;
                CurrentRun.RunScreenshotElapsed = TimeSpan.Zero;
                CurrentRun.RunScreenProbeCount = 0;
                CurrentRun.RunScreenProbeElapsed = TimeSpan.Zero;
                CurrentRun.RunScreenSanityWarnings = 0;
                CurrentRun.RunImageEncodeCount = 0;
                CurrentRun.RunImageEncodeElapsed = TimeSpan.Zero;
                CurrentRun.RunScreenLogCount = 0;
                CurrentRun.RunScreenLogElapsed = TimeSpan.Zero;
                CurrentRun.RunArtifactLogWrites = 0;
                CurrentRun.RunArtifactLogElapsed = TimeSpan.Zero;
                CurrentRun.RunUiaCalls = 0;
                CurrentRun.RunUiaElapsed = TimeSpan.Zero;
                CurrentRun.RunLocalActions = 0;
                CurrentRun.RunLocalActionElapsed = TimeSpan.Zero;
                CurrentRun.ReportedSanityWarnings.Clear();
                CurrentRun.ReportedWebSearchCallIds.Clear();
            }
        
            internal static void PrintRunMetrics()
            {
                if (CurrentRun.RunOpenAiCalls <= 0)
                    return;
        
                var cachedPct = CurrentRun.RunInputTokens > 0 ? CurrentRun.RunCachedTokens * 100.0 / CurrentRun.RunInputTokens : 0.0;
                Console.WriteLine(FormattableString.Invariant($"[metrics] openai_calls={CurrentRun.RunOpenAiCalls}; openai_retries={CurrentRun.RunOpenAiRetries}; openai_time={CurrentRun.RunOpenAiElapsed.TotalSeconds:0.0}s; request_bytes={CurrentRun.RunOpenAiRequestBytes}; response_bytes={CurrentRun.RunOpenAiBytes}; input_tokens={CurrentRun.RunInputTokens}; cached_tokens={CurrentRun.RunCachedTokens}; cached_pct={cachedPct:0.0}; output_tokens={CurrentRun.RunOutputTokens}; reasoning_tokens={CurrentRun.RunReasoningTokens}; web_search_calls={CurrentRun.RunWebSearchCalls}; early_control_streams={CurrentRun.RunEarlyAcceptedControlStreams}; context_turns={CurrentRun.RunControlContextTurns}; context_restarts={CurrentRun.RunControlContextRestarts}; context_fallbacks={CurrentRun.RunControlContextFallbacks}; context_compactions={CurrentRun.RunControlContextCompactions}; compaction_fallbacks={CurrentRun.RunControlCompactionFallbacks}; multi_candidate_responses={CurrentRun.RunMultiCandidateResponses}; screenshots={CurrentRun.RunScreenshotCount}; screenshot_time={CurrentRun.RunScreenshotElapsed.TotalSeconds:0.0}s; screen_probes={CurrentRun.RunScreenProbeCount}; screen_probe_time={CurrentRun.RunScreenProbeElapsed.TotalSeconds:0.0}s; screen_sanity_warnings={CurrentRun.RunScreenSanityWarnings}; image_encodes={CurrentRun.RunImageEncodeCount}; image_encode_time={CurrentRun.RunImageEncodeElapsed.TotalSeconds:0.0}s; screen_logs={CurrentRun.RunScreenLogCount}; screen_log_time={CurrentRun.RunScreenLogElapsed.TotalSeconds:0.0}s; artifact_logs={CurrentRun.RunArtifactLogWrites}; artifact_log_time={CurrentRun.RunArtifactLogElapsed.TotalSeconds:0.0}s; uia_calls={CurrentRun.RunUiaCalls}; uia_time={CurrentRun.RunUiaElapsed.TotalSeconds:0.0}s; local_actions={CurrentRun.RunLocalActions}; local_action_time={CurrentRun.RunLocalActionElapsed.TotalSeconds:0.0}s"));
                if (CurrentRun.RunEarlyAcceptedControlStreams > 0)
                    Console.WriteLine($"[metrics] token_usage_partial=true; {CurrentRun.RunEarlyAcceptedControlStreams} control response(s) were accepted before the final usage event");
                if (ShouldWarnPromptCache(Options.UsePromptCache, CurrentRun.RunOpenAiCalls, CurrentRun.RunInputTokens, CurrentRun.RunCachedTokens, CurrentRun.RunEarlyAcceptedControlStreams))
                    Console.WriteLine("[metrics] prompt_cache_warning=cached_tokens_zero; check model support, prompt stability, and prompt_cache_key scope");
            }

            internal static bool ShouldWarnPromptCache(bool enabled, int calls, long inputTokens, long cachedTokens, int earlyAcceptedControlStreams)
            {
                return enabled && calls > 1 && inputTokens > 0 && cachedTokens == 0 && earlyAcceptedControlStreams == 0;
            }
        
            internal static void RecordUiaMetric(RuntimeClock.ElapsedTimer sw)
            {
                sw.Stop();
                CurrentRun.RunUiaCalls++;
                CurrentRun.RunUiaElapsed += sw.Elapsed;
            }
    }
