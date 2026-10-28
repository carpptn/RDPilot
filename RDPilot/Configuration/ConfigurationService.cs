
namespace RDPilot.Configuration;

    /// <summary>
    /// Coordinates configuration sources and owns shared validation, profiles, and normalization.
    /// </summary>
    internal static class ConfigurationService
    {
            internal static void PrintEffectiveConfig()
            {
                var uiMode = Options.ForceRealUiOnly
                    ? "real-ui-only"
                    : (Options.AllowHighLevelActions || Options.AllowRunCommand ? "real UI + enabled local adapters" : "real UI");
        
                Console.WriteLine($"Profile: {Options.RunProfile}");
                Console.WriteLine($"Execution: {(Options.BatchMode ? "batch" : "interactive")}");
                Console.WriteLine($"Observation: mode={Options.ObservationMode}; initial={(Options.ObservationMode == "auto" ? "general" : Options.ObservationMode)}; verbose_log={(Options.ObservationLogVerbose ? "on" : "off")}");
                Console.WriteLine($"Model: {Options.Model}");
                Console.WriteLine($"QA model: {EffectiveQaModel()}; verify model: {EffectiveVerifyModel()}");
                Console.WriteLine($"Reasoning effort: control={ReasoningEffortDisplay(Options.Model, Options.ReasoningEffort)}; qa={ReasoningEffortDisplay(EffectiveQaModel(), EffectiveQaReasoningEffort())}; verify={ReasoningEffortDisplay(EffectiveVerifyModel(), EffectiveVerifyReasoningEffort())}; adaptive={(Options.AdaptiveReasoningEffort ? "on" : "off")}");
                Console.WriteLine($"UI mode: {uiMode}; mouse={(Options.MouseEnabled ? "enabled" : "disabled")}; desktop={(Options.MultiMonitorEnabled ? "virtual multi-monitor" : "primary monitor only")}; post-action delay={Options.UiSettleDelayMs} ms; grid={(Options.GridStepPx > 0 ? $"{Options.GridStepPx}px" : "off")}");
                Console.WriteLine($"Images: send={Options.ScreenshotSendFormat} max-width={ScreenshotMaxWidthDisplay()} focused-overview={FocusedOverviewMaxWidthDisplay()} qa={QaScreenshotMaxWidthDisplay()} verify={VerifyScreenshotMaxWidthDisplay()} quality={Options.ScreenshotJpegQuality}; crop={Options.CropSendFormat} max-width={CropMaxWidthDisplay()} size={Options.FocusCropSize}px; screen-log={Options.ScreenLogFormat} max-width={ScreenLogMaxWidthDisplay()}; focus_uia={(Options.IncludeFocusUia ? "on" : "off")}; focus crop={(Options.IncludeFocusUiaCrop ? "on" : "off")}; debug images={(Options.DebugImages ? "on" : "off")}");
                Console.WriteLine($"Output: max_tokens={Options.MaxOutputTokens}; qa_max_tokens={Options.QaMaxOutputTokens}; verify_max_tokens={Options.VerifyMaxOutputTokens}; turn_reanalysis_max_tokens={Options.TurnReanalysisMaxOutputTokens}; incomplete_effort_fallback=max->low; no_effort_cap_retries={Options.IncompleteMaxOutputRetries}/{Options.IncompleteMaxOutputTokenCap}; verbosity={Options.TextVerbosity}; action_text_chars={Options.MaxActionTextChars}; history_chars={Options.HistoryTailChars}; history_lines={Options.HistoryTailLines}; verify={Options.VerifyMode}; verify_early_steps={Options.VerifyEarlySteps}; verify_low_confidence={Options.VerifyLowConfidenceThreshold:0.##}; skip_verify_confidence={Options.SkipVerifyConfidenceThreshold:0.##}; verify-refresh={(Options.RefreshScreenshotBeforeVerify ? "on" : "off")}; long text paste threshold={Options.ClipboardPasteThreshold}; prompt_cache={(Options.UsePromptCache ? Options.PromptCacheKey ?? "on" : "off")}; previous_response_state={(Options.UsePreviousResponseState ? "on" : "off")}; reasoning_context={(Options.UsePreviousResponseState ? Options.ControlReasoningContext : "current_turn")}; context_compaction={(Options.UsePreviousResponseState && Options.ControlContextCompactionEnabled ? Options.ControlContextCompactThreshold.ToString() : "off")}; context_fallback_limit={Options.ControlContextFallbackLimit}; omit_unchanged_screen={(Options.OmitUnchangedScreenImageWithState ? "on" : "off")}");
                Console.WriteLine($"Logs: requests={(Options.LogRequests ? (Options.PrettyRequestLogs ? "pretty" : "compact") : "off")}; screens={(Options.LogScreens ? "on" : "state-only")}; retries={Options.OpenAiMaxRetries}; timeout={(Options.OpenAiTimeoutSeconds > 0 ? $"{Options.OpenAiTimeoutSeconds}s" : "infinite")}; batch candidates={(Options.ExecuteMultiActionCandidates ? $"on/{Options.MaxQueuedBatchActions}" : "off")}; turn_batch_inputs={Options.TurnBasedMaxBatchInputs}");
                Console.WriteLine($"Loop guards: goal_mode={Options.GoalMode}; max_steps={(Options.MaxSteps > 0 ? Options.MaxSteps.ToString() : "unlimited")}; max_wait={(Options.MaxWaitSeconds > 0 ? $"{Options.MaxWaitSeconds}s" : "off")}; stagnation={(Options.MaxStagnationStepsBeforeAbort > 0 ? Options.MaxStagnationStepsBeforeAbort.ToString() : "off")}; repeated_action={(Options.MaxRepeatedActionBeforeAbort > 0 ? Options.MaxRepeatedActionBeforeAbort.ToString() : "off")}; rejected_proposals={(Options.MaxRejectedProposalRepeatsBeforeAbort > 0 ? Options.MaxRejectedProposalRepeatsBeforeAbort.ToString() : "off")}; inspection_actions={(Options.MaxConsecutiveInspectionActions > 0 ? Options.MaxConsecutiveInspectionActions.ToString() : "unlimited")}; repeat_cooldown={(Options.ActionRepeatCooldownSteps > 0 ? Options.ActionRepeatCooldownSteps.ToString() : "off")}; proactive_confidence={Options.ProactiveLoopConfidenceThreshold:0.00}; model_failures={(Options.MaxModelFailuresBeforeAbort > 0 ? Options.MaxModelFailuresBeforeAbort.ToString() : "off")}; action_failures={(Options.MaxActionFailuresBeforeAbort > 0 ? Options.MaxActionFailuresBeforeAbort.ToString() : "off")}");
                Console.WriteLine($"Recovery memory: {(Options.RecoveryMemoryEnabled ? $"on; trigger={Options.RecoveryMemoryTriggerSteps}; validate={Options.RecoveryMemoryValidationSteps}; failure_limit={Options.RecoveryMemoryFailureLimit}; active_max={Options.RecoveryMemoryMaxLessons}; quarantine_max={Options.RecoveryMemoryMaxQuarantinedLessons}; context={Options.RecoveryMemoryReservedLessonsPerContext}/{Options.RecoveryMemorySoftMaxLessonsPerContext}; file_max={Options.RecoveryMemoryMaxFileBytes}B; archive={EffectiveRecoveryMemoryArchivePath()} ({Options.RecoveryMemoryArchiveMaxBytes}B x {Options.RecoveryMemoryArchiveRetainedFiles}); prompt={Options.RecoveryMemoryPromptMaxLessons}; progress_verify={(Options.RecoveryProgressVerificationEnabled ? $"on/{Options.RecoveryProgressConfidenceThreshold:0.00}" : "off")}; telemetry={Options.RecoveryTelemetryMaxBytes}B/{Options.RecoveryTelemetryRetainedFiles}; replay_auto_export={(Options.LoopReplayAutoExportEnabled ? EffectiveLoopReplayCorpusPath() : "off")}; path={EffectiveRecoveryMemoryPath()}" : "off")}");
                Console.WriteLine($"Runtime bounds: semantic_states={Options.RuntimeSemanticStateLimit}; graph_edges={Options.RuntimeGraphEdgeLimit}; recovery_actions={Options.RuntimeRecoveryActionLimit}; cooldowns={Options.RuntimeCooldownEntryLimit}; graph_candidate_ttl={Options.GraphCandidateTtlSteps} steps");
                Console.WriteLine($"Screen settle: polling={(Options.ScreenPollingEnabled ? $"on initial={Options.ScreenPollInitialDelayMs}ms interval={Options.ScreenPollIntervalMs}ms timeout={Options.ScreenPollTimeoutMs}ms wait_extra={Options.WaitNoChangeExtraMs}ms" : "off")}; sanity={(Options.ScreenSanityChecks ? "on" : "off")}");
                Console.WriteLine($"Input retry: sendinput_retries={Options.SendInputMaxRetries}; retry_delay={Options.SendInputRetryDelayMs}ms");
                Console.WriteLine($"Console: auto_hide={(Options.AutoHideConsoleDuringRun ? "on" : "off")}; minimize_flag={(Options.MinimizeConsoleDuringRun ? "on" : "off")}; restore_after_run={(Options.RestoreConsoleAfterRun ? "on" : "off")}");
                Console.WriteLine($"UIA targets: {(Options.IncludeUiaTargets ? $"on/{Options.MaxUiaTargets}; name_chars={Options.UiaTargetNameMaxChars}; summary_chars={Options.UiaSummaryMaxChars}; scan={Options.UiaScanTimeBudgetMs}ms/{Options.MaxUiaNodesScanned} nodes; candidates={Options.UiaCandidateMultiplier}x; max_area={Options.MaxUiaTargetAreaRatio:0.##}; reuse={(Options.ReuseUiaTargetsWhenScreenUnchanged ? "on" : "off")}" : "off")}");
                Console.WriteLine($"Local high-level actions: {(Options.AllowHighLevelActions ? "enabled" : "disabled")}; run_command={(Options.AllowRunCommand ? "enabled" : "disabled")}; real_ui_only={(Options.ForceRealUiOnly ? "on" : "off")}");
                Console.WriteLine($"Internet tools: web_search={(Options.AllowWebSearch ? "enabled" : "disabled")}");
            }

            internal static void PrintStartupSummary()
            {
                ConsoleTheme.WriteStartupBanner(
                    Options.Model,
                    Options.RunProfile,
                    ReasoningEffortDisplay(Options.Model, Options.ReasoningEffort),
                    EffectiveQaModel(),
                    EffectiveVerifyModel());
            }
        
            internal static void ApplyEnvironmentConfig() => ConfigurationLoader.ApplyEnvironmentConfig();
        
            internal static void ApplyEnvironmentProfileOverrides() => ConfigurationLoader.ApplyEnvironmentProfileOverrides();
        
            internal static bool IsTruthyEnv(string name)
            {
                var value = Environment.GetEnvironmentVariable(name);
                return !string.IsNullOrWhiteSpace(value) && IsTruthy(value);
            }
        
            internal static string? ApplyCliArgs(string[] args) => ParseCliArgs(args).Task;

            internal static CliValidationResult ParseCliArgs(string[] args) =>
                CliConfigurationParser.ParseCliArgs(args);
        
            internal static void NormalizeConfig()
            {
                if (Options.ForceRealUiOnly)
                {
                    Options.AllowHighLevelActions = false;
                    Options.AllowRunCommand = false;
                }
                if (!Options.IncludeFocusUia)
                    Options.IncludeFocusUiaCrop = false;
                if (Options.MaxUiaTargets <= 0)
                    Options.IncludeUiaTargets = false;
                Options.FocusRingPadding = Math.Clamp(Options.FocusRingPadding, 0, 100);
                Options.FocusRingThickness = Math.Clamp(Options.FocusRingThickness, 1, 100);
                Options.FocusGlowThickness = Math.Clamp(Options.FocusGlowThickness, 1, 100);
                Options.FocusCornerRadius = Math.Clamp(Options.FocusCornerRadius, 0, 500);
                Options.LargeFocusOverlayAreaRatio = Math.Clamp(Options.LargeFocusOverlayAreaRatio, 0.0, 1.0);
                Options.ClickAimEdgeAdjustMaxAreaRatio = Math.Clamp(Options.ClickAimEdgeAdjustMaxAreaRatio, 0.0, 1.0);
                Options.ClickAimEdgeMarginRatio = Math.Clamp(Options.ClickAimEdgeMarginRatio, 0.0, 1.0);
                Options.ClickAimEdgeMinMarginPx = Math.Clamp(Options.ClickAimEdgeMinMarginPx, 0, 2000);
                Options.ClickAimEdgeMaxMarginPx = Math.Clamp(Options.ClickAimEdgeMaxMarginPx, Options.ClickAimEdgeMinMarginPx, 2000);
                Options.GridStepPx = Math.Max(0, Options.GridStepPx);
                Options.GridLabelEveryPx = Math.Max(0, Options.GridLabelEveryPx);
                Options.GridMajorEveryPx = Math.Max(0, Options.GridMajorEveryPx);
                Options.AimExpireDelta = Math.Clamp(Options.AimExpireDelta, 0.0, 1.0);
                Options.NoChangeThreshold = Math.Clamp(Options.NoChangeThreshold, 0.0, 1.0);
                Options.MaxGesturePathPoints = Math.Clamp(Options.MaxGesturePathPoints, 2, 1024);
                Options.MaxGestureDurationMs = Math.Clamp(Options.MaxGestureDurationMs, 100, 60000);
                Options.MaxHeldKeys = Math.Clamp(Options.MaxHeldKeys, 1, 16);
                Options.MaxKeyHoldDurationMs = Math.Clamp(Options.MaxKeyHoldDurationMs, 100, 60000);
                Options.IneffectiveMouseClusterPx = Math.Clamp(Options.IneffectiveMouseClusterPx, 8, 2000);
                Options.MaxFocusUiaCropPixels = Math.Max(0, Options.MaxFocusUiaCropPixels);
                Options.MaxArtifactsPerDir = Math.Max(0, Options.MaxArtifactsPerDir);
                Options.MaxBatchedGesturePoints = Math.Clamp(Options.MaxBatchedGesturePoints, 2, 10000);
                Options.MaxBatchedGestureDurationMs = Math.Clamp(Options.MaxBatchedGestureDurationMs, 100, 60000);
                Options.MaxSteps = Math.Max(0, Options.MaxSteps);
                Options.MaxActionTextChars = Math.Max(256, Options.MaxActionTextChars);
                Options.TurnReanalysisMaxOutputTokens = Math.Max(Options.MaxOutputTokens, Options.TurnReanalysisMaxOutputTokens);
                Options.IncompleteMaxOutputTokenCap = Math.Max(Options.TurnReanalysisMaxOutputTokens, Options.IncompleteMaxOutputTokenCap);
                Options.RecoveryMemoryTriggerSteps = Math.Max(1, Options.RecoveryMemoryTriggerSteps);
                Options.RecoveryMemoryValidationSteps = Math.Max(1, Options.RecoveryMemoryValidationSteps);
                Options.RecoveryMemoryMaxLessons = Math.Max(1, Options.RecoveryMemoryMaxLessons);
                Options.RecoveryMemoryMaxQuarantinedLessons = Math.Max(1, Options.RecoveryMemoryMaxQuarantinedLessons);
                Options.RecoveryMemoryReservedLessonsPerContext = Math.Max(0, Options.RecoveryMemoryReservedLessonsPerContext);
                Options.RecoveryMemorySoftMaxLessonsPerContext = Math.Max(
                    Math.Max(1, Options.RecoveryMemoryReservedLessonsPerContext),
                    Options.RecoveryMemorySoftMaxLessonsPerContext);
                Options.RecoveryMemoryMaxFileBytes = Math.Max(1024 * 1024, Options.RecoveryMemoryMaxFileBytes);
                Options.RecoveryMemoryArchiveMaxBytes = Math.Max(1024 * 1024, Options.RecoveryMemoryArchiveMaxBytes);
                Options.RecoveryMemoryArchiveRetainedFiles = Math.Clamp(Options.RecoveryMemoryArchiveRetainedFiles, 1, 20);
                Options.RecoveryMemoryPromptMaxLessons = Math.Clamp(Options.RecoveryMemoryPromptMaxLessons, 0, 10);
                Options.RecoveryMemoryFailureLimit = Math.Max(1, Options.RecoveryMemoryFailureLimit);
                Options.MaxRejectedProposalRepeatsBeforeAbort = Math.Max(0, Options.MaxRejectedProposalRepeatsBeforeAbort);
                Options.RuntimeSemanticStateLimit = Math.Clamp(Options.RuntimeSemanticStateLimit, 32, 100000);
                Options.RuntimeGraphEdgeLimit = Math.Clamp(Options.RuntimeGraphEdgeLimit, 32, 100000);
                Options.RuntimeRecoveryActionLimit = Math.Clamp(Options.RuntimeRecoveryActionLimit, 8, 10000);
                Options.RuntimeCooldownEntryLimit = Math.Clamp(Options.RuntimeCooldownEntryLimit, 16, 100000);
                Options.GraphCandidateTtlSteps = Math.Clamp(Options.GraphCandidateTtlSteps, 2, 10000);
                Options.ProactiveLoopConfidenceThreshold = Math.Clamp(Options.ProactiveLoopConfidenceThreshold, 0.5, 1.0);
                Options.RecoveryProgressConfidenceThreshold = Math.Clamp(Options.RecoveryProgressConfidenceThreshold, 0.5, 1.0);
                Options.RecoveryTelemetryMaxBytes = Math.Max(65536, Options.RecoveryTelemetryMaxBytes);
                Options.RecoveryTelemetryRetainedFiles = Math.Clamp(Options.RecoveryTelemetryRetainedFiles, 1, 20);
                Options.GoalMode = NormalizeGoalMode(Options.GoalMode) is var normalizedGoalMode &&
                           normalizedGoalMode is "auto" or "finite" or "continuous"
                    ? normalizedGoalMode
                    : "auto";
                Options.ObservationMode = NormalizeObservationMode(Options.ObservationMode) is var normalizedObservationMode &&
                                  IsAllowedObservationMode(normalizedObservationMode)
                    ? normalizedObservationMode
                    : "auto";
                Options.ControlReasoningContext = NormalizeReasoningContext(Options.ControlReasoningContext);
                Options.ControlContextCompactThreshold = Math.Max(1, Options.ControlContextCompactThreshold);
                Options.ControlContextFallbackLimit = Math.Clamp(Options.ControlContextFallbackLimit, 1, 20);
                if (Options.RecoveryMemoryEnabled || !string.IsNullOrWhiteSpace(Options.LoopReplayImportPath) ||
                    !string.IsNullOrWhiteSpace(Options.LoopReplayExportPath))
                {
                    try
                    {
                        var paths = RecoveryStoragePaths.ResolveValidated(Options);
                        if (!string.IsNullOrWhiteSpace(Options.LoopReplayExportPath))
                            _ = paths.ValidateCorpusDestination(Options.LoopReplayExportPath);
                        if (!string.IsNullOrWhiteSpace(Options.RecoveryMemoryExportPath))
                            _ = paths.ValidateMemoryExportDestination(Options.RecoveryMemoryExportPath);
                    }
                    catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NotSupportedException or PathTooLongException)
                    {
                        ReportConfigurationError($"Invalid recovery storage configuration: {ex.Message}");
                    }
                }
            }
        

            internal static void ReportConfigurationError(string message)
            {
                CurrentRun.ConfigurationErrors.Add(message);
                CurrentRun.CliArgumentError = true;
                Console.Error.WriteLine(message);
            }

            internal static string? ResolveAutomaticConfigPath(string baseDirectory, string currentDirectory) =>
                ConfigurationLoader.ResolveAutomaticConfigPath(baseDirectory, currentDirectory);

            internal static void ApplyAutomaticConfigFile(string? baseDirectory = null, string? currentDirectory = null) =>
                ConfigurationLoader.ApplyAutomaticConfigFile(baseDirectory, currentDirectory);

            internal static void ApplyConfigFile(string? path) => ConfigurationLoader.ApplyConfigFile(path);
        
            static readonly string[] ProfileControlledSettingNames =
            [
                nameof(Options.ReasoningEffort),
                nameof(Options.QaReasoningEffort),
                nameof(Options.VerifyReasoningEffort),
                nameof(Options.UiSettleDelayMs),
                nameof(Options.MaxOutputTokens),
                nameof(Options.QaMaxOutputTokens),
                nameof(Options.VerifyMaxOutputTokens),
                nameof(Options.TurnReanalysisMaxOutputTokens),
                nameof(Options.QaScreenshotMaxWidth),
                nameof(Options.VerifyScreenshotMaxWidth),
                nameof(Options.TextVerbosity),
                nameof(Options.HistoryTailChars),
                nameof(Options.HistoryTailLines),
                nameof(Options.MaxStagnationStepsBeforeAbort),
                nameof(Options.MaxRepeatedActionBeforeAbort),
                nameof(Options.ActionRepeatCooldownSteps),
                nameof(Options.SkipVerifyConfidenceThreshold),
                nameof(Options.MaxModelFailuresBeforeAbort),
                nameof(Options.MaxActionFailuresBeforeAbort),
                nameof(Options.ScreenPollingEnabled),
                nameof(Options.ScreenPollInitialDelayMs),
                nameof(Options.ScreenPollIntervalMs),
                nameof(Options.ScreenPollTimeoutMs),
                nameof(Options.WaitNoChangeExtraMs),
                nameof(Options.MaxWaitSeconds),
                nameof(Options.MaxScreenshotSendWidth),
                nameof(Options.ScreenshotSendFormat),
                nameof(Options.FocusedOverviewMaxWidth),
                nameof(Options.MaxCropSendWidth),
                nameof(Options.CropSendFormat),
                nameof(Options.ScreenshotJpegQuality),
                nameof(Options.ScreenLogFormat),
                nameof(Options.MaxScreenLogWidth),
                nameof(Options.PrettyRequestLogs),
                nameof(Options.IncludeFocusUiaCrop),
                nameof(Options.DebugImages),
                nameof(Options.VerifyMode),
                nameof(Options.UiaTargetNameMaxChars),
                nameof(Options.UiaSummaryMaxChars),
                nameof(Options.UiaScanTimeBudgetMs),
                nameof(Options.MaxUiaNodesScanned),
                nameof(Options.UiaCandidateMultiplier),
                nameof(Options.MaxUiaTargetAreaRatio)
            ];
            static Dictionary<string, object?>? CodeProfileDefaults;

            static int ReasoningEffortRank(string? effort) =>
                effort?.Trim().ToLowerInvariant() switch
                {
                    "none" => 0,
                    "minimal" => 1,
                    "low" => 2,
                    "medium" => 3,
                    "high" => 4,
                    "xhigh" => 5,
                    "max" => 6,
                    _ => -1
                };

            static bool IsAtLeastReasoningEffort(string? effort, string minimum) =>
                ReasoningEffortRank(effort) >= ReasoningEffortRank(minimum);

            static string? RaiseReasoningEffort(string? current, string minimum) =>
                IsAtLeastReasoningEffort(current, minimum) ? current : minimum;

            internal static void ApplyProfile(string? profile)
            {
                if (string.IsNullOrWhiteSpace(profile))
                    return;
        
                var normalized = profile.Trim().ToLowerInvariant();
                if (normalized is not ("custom" or "fast" or "balanced" or "quality"))
                {
                    ReportConfigurationError($"Unknown profile '{profile}'. Allowed: custom, fast, balanced, quality.");
                    return;
                }

                // Profiles are applied from a known baseline. This makes
                // `custom` a real reset to code defaults and prevents results
                // from depending on which profile happened to run earlier.
                ApplyCustomProfileDefaults();
                switch (normalized)
                {
                    case "custom":
                        Options.RunProfile = "custom";
                        break;

                    case "fast":
                        Options.RunProfile = "fast";
                        if (!Options.ReasoningEffortExplicit)
                            Options.ReasoningEffort = RaiseReasoningEffort(Options.ReasoningEffort, "low");
                        if (!Options.QaReasoningEffortExplicit &&
                            !IsAtLeastReasoningEffort(EffectiveQaReasoningEffort(), "low"))
                            Options.QaReasoningEffort = "low";
                        if (!Options.VerifyReasoningEffortExplicit &&
                            !IsAtLeastReasoningEffort(EffectiveVerifyReasoningEffort(), "low"))
                            Options.VerifyReasoningEffort = "low";
                        Options.UiSettleDelayMs = Math.Min(Options.UiSettleDelayMs, 300);
                        if (!IsAtLeastReasoningEffort(Options.ReasoningEffort, "low"))
                            Options.MaxOutputTokens = Math.Min(Options.MaxOutputTokens, 300);
                        if (!IsAtLeastReasoningEffort(EffectiveQaReasoningEffort(), "low"))
                            Options.QaMaxOutputTokens = Math.Min(Options.QaMaxOutputTokens, 300);
                        if (!IsAtLeastReasoningEffort(EffectiveVerifyReasoningEffort(), "low"))
                            Options.VerifyMaxOutputTokens = Math.Min(Options.VerifyMaxOutputTokens, 120);
                        Options.QaScreenshotMaxWidth = Math.Min(Options.QaScreenshotMaxWidth, 1024);
                        Options.VerifyScreenshotMaxWidth = Math.Min(Options.VerifyScreenshotMaxWidth, 1024);
                        Options.TextVerbosity = "low";
                        Options.HistoryTailChars = Math.Min(Options.HistoryTailChars, 1200);
                        Options.HistoryTailLines = Math.Min(Options.HistoryTailLines, 12);
                        Options.MaxStagnationStepsBeforeAbort = Math.Min(Options.MaxStagnationStepsBeforeAbort, 8);
                        Options.MaxRepeatedActionBeforeAbort = Math.Min(Options.MaxRepeatedActionBeforeAbort, 5);
                        Options.ActionRepeatCooldownSteps = Math.Min(Options.ActionRepeatCooldownSteps, 2);
                        Options.SkipVerifyConfidenceThreshold = Math.Min(Options.SkipVerifyConfidenceThreshold, 0.92);
                        Options.MaxModelFailuresBeforeAbort = Math.Min(Options.MaxModelFailuresBeforeAbort, 2);
                        Options.MaxActionFailuresBeforeAbort = Math.Min(Options.MaxActionFailuresBeforeAbort, 2);
                        Options.ScreenPollingEnabled = true;
                        Options.ScreenPollInitialDelayMs = Math.Min(Options.ScreenPollInitialDelayMs, 120);
                        Options.ScreenPollIntervalMs = Math.Min(Options.ScreenPollIntervalMs, 150);
                        Options.ScreenPollTimeoutMs = Math.Min(Options.ScreenPollTimeoutMs, 1200);
                        Options.WaitNoChangeExtraMs = Math.Min(Options.WaitNoChangeExtraMs, 750);
                        Options.MaxWaitSeconds = Math.Min(Options.MaxWaitSeconds, 30);
                        Options.MaxScreenshotSendWidth = 1280;
                        Options.ScreenshotSendFormat = "jpeg";
                        Options.FocusedOverviewMaxWidth = 640;
                        Options.MaxCropSendWidth = 768;
                        Options.CropSendFormat = "jpeg";
                        Options.ScreenshotJpegQuality = 80;
                        Options.ScreenLogFormat = "jpeg";
                        Options.MaxScreenLogWidth = 1280;
                        Options.PrettyRequestLogs = false;
                        Options.IncludeFocusUiaCrop = false;
                        Options.DebugImages = false;
                        Options.VerifyMode = "auto";
                        Options.UiaTargetNameMaxChars = Math.Min(Options.UiaTargetNameMaxChars, 48);
                        Options.UiaSummaryMaxChars = Math.Min(Options.UiaSummaryMaxChars, 320);
                        Options.UiaScanTimeBudgetMs = Math.Min(Options.UiaScanTimeBudgetMs, 60);
                        Options.MaxUiaNodesScanned = Math.Min(Options.MaxUiaNodesScanned, 400);
                        Options.UiaCandidateMultiplier = Math.Min(Options.UiaCandidateMultiplier, 4);
                        Options.MaxUiaTargetAreaRatio = Math.Min(Options.MaxUiaTargetAreaRatio, 0.45);
                        break;
        
                    case "balanced":
                        Options.RunProfile = "balanced";
                        if (!Options.ReasoningEffortExplicit)
                            Options.ReasoningEffort = RaiseReasoningEffort(Options.ReasoningEffort, "low");
                        if (!Options.QaReasoningEffortExplicit &&
                            !IsAtLeastReasoningEffort(EffectiveQaReasoningEffort(), "low"))
                            Options.QaReasoningEffort = "low";
                        if (!Options.VerifyReasoningEffortExplicit &&
                            !IsAtLeastReasoningEffort(EffectiveVerifyReasoningEffort(), "low"))
                            Options.VerifyReasoningEffort = "low";
                        Options.UiSettleDelayMs = Math.Max(Options.UiSettleDelayMs, 500);
                        if (!IsAtLeastReasoningEffort(Options.ReasoningEffort, "low"))
                            Options.MaxOutputTokens = Math.Max(Options.MaxOutputTokens, 450);
                        if (!IsAtLeastReasoningEffort(EffectiveQaReasoningEffort(), "low"))
                            Options.QaMaxOutputTokens = Math.Max(Options.QaMaxOutputTokens, 450);
                        if (!IsAtLeastReasoningEffort(EffectiveVerifyReasoningEffort(), "low"))
                            Options.VerifyMaxOutputTokens = Math.Max(Options.VerifyMaxOutputTokens, 160);
                        Options.QaScreenshotMaxWidth = Math.Max(Options.QaScreenshotMaxWidth, 1280);
                        Options.VerifyScreenshotMaxWidth = Math.Max(Options.VerifyScreenshotMaxWidth, 1280);
                        Options.TextVerbosity = "low";
                        Options.HistoryTailChars = Math.Max(Options.HistoryTailChars, 2000);
                        Options.HistoryTailLines = Math.Max(Options.HistoryTailLines, 20);
                        Options.MaxStagnationStepsBeforeAbort = Math.Max(Options.MaxStagnationStepsBeforeAbort, 12);
                        Options.MaxRepeatedActionBeforeAbort = Math.Max(Options.MaxRepeatedActionBeforeAbort, 7);
                        Options.ActionRepeatCooldownSteps = Math.Max(Options.ActionRepeatCooldownSteps, 3);
                        Options.SkipVerifyConfidenceThreshold = Math.Min(Options.SkipVerifyConfidenceThreshold, 0.95);
                        Options.MaxModelFailuresBeforeAbort = Math.Max(Options.MaxModelFailuresBeforeAbort, 3);
                        Options.MaxActionFailuresBeforeAbort = Math.Max(Options.MaxActionFailuresBeforeAbort, 3);
                        Options.ScreenPollingEnabled = true;
                        Options.ScreenPollInitialDelayMs = Math.Max(Options.ScreenPollInitialDelayMs, 150);
                        Options.ScreenPollIntervalMs = Math.Max(Options.ScreenPollIntervalMs, 180);
                        Options.ScreenPollTimeoutMs = Math.Max(Options.ScreenPollTimeoutMs, 1800);
                        Options.WaitNoChangeExtraMs = Math.Max(Options.WaitNoChangeExtraMs, 1000);
                        Options.MaxWaitSeconds = Math.Max(Options.MaxWaitSeconds, 60);
                        Options.MaxScreenshotSendWidth = 1600;
                        Options.ScreenshotSendFormat = "jpeg";
                        Options.FocusedOverviewMaxWidth = Math.Max(Options.FocusedOverviewMaxWidth, 960);
                        Options.MaxCropSendWidth = 1024;
                        Options.CropSendFormat = "jpeg";
                        Options.ScreenshotJpegQuality = 88;
                        Options.ScreenLogFormat = "jpeg";
                        Options.MaxScreenLogWidth = 1600;
                        Options.PrettyRequestLogs = false;
                        Options.IncludeFocusUiaCrop = true;
                        Options.DebugImages = false;
                        Options.VerifyMode = "auto";
                        Options.UiaTargetNameMaxChars = Math.Max(Options.UiaTargetNameMaxChars, 64);
                        Options.UiaSummaryMaxChars = Math.Max(Options.UiaSummaryMaxChars, 480);
                        Options.UiaScanTimeBudgetMs = Math.Max(Options.UiaScanTimeBudgetMs, 90);
                        Options.MaxUiaNodesScanned = Math.Max(Options.MaxUiaNodesScanned, 700);
                        Options.UiaCandidateMultiplier = Math.Max(Options.UiaCandidateMultiplier, 5);
                        Options.MaxUiaTargetAreaRatio = Math.Max(Options.MaxUiaTargetAreaRatio, 0.55);
                        break;
        
                    case "quality":
                        Options.RunProfile = "quality";
                        if (!Options.ReasoningEffortExplicit)
                            Options.ReasoningEffort = RaiseReasoningEffort(Options.ReasoningEffort, "medium");
                        if (!Options.QaReasoningEffortExplicit &&
                            !IsAtLeastReasoningEffort(EffectiveQaReasoningEffort(), "low"))
                            Options.QaReasoningEffort = "low";
                        if (!Options.VerifyReasoningEffortExplicit &&
                            !IsAtLeastReasoningEffort(EffectiveVerifyReasoningEffort(), "low"))
                            Options.VerifyReasoningEffort = "low";
                        Options.UiSettleDelayMs = Math.Max(Options.UiSettleDelayMs, 1000);
                        if (!IsAtLeastReasoningEffort(Options.ReasoningEffort, "medium"))
                            Options.MaxOutputTokens = Math.Max(Options.MaxOutputTokens, 800);
                        if (!IsAtLeastReasoningEffort(EffectiveQaReasoningEffort(), "low"))
                            Options.QaMaxOutputTokens = Math.Max(Options.QaMaxOutputTokens, 800);
                        if (!IsAtLeastReasoningEffort(EffectiveVerifyReasoningEffort(), "low"))
                            Options.VerifyMaxOutputTokens = Math.Max(Options.VerifyMaxOutputTokens, 240);
                        Options.QaScreenshotMaxWidth = 0;
                        Options.VerifyScreenshotMaxWidth = 0;
                        Options.TextVerbosity = "medium";
                        Options.HistoryTailChars = Math.Max(Options.HistoryTailChars, 4000);
                        Options.HistoryTailLines = Math.Max(Options.HistoryTailLines, 40);
                        Options.MaxStagnationStepsBeforeAbort = Math.Max(Options.MaxStagnationStepsBeforeAbort, 20);
                        Options.MaxRepeatedActionBeforeAbort = Math.Max(Options.MaxRepeatedActionBeforeAbort, 10);
                        Options.ActionRepeatCooldownSteps = Math.Max(Options.ActionRepeatCooldownSteps, 4);
                        Options.SkipVerifyConfidenceThreshold = Math.Max(Options.SkipVerifyConfidenceThreshold, 0.98);
                        Options.MaxModelFailuresBeforeAbort = Math.Max(Options.MaxModelFailuresBeforeAbort, 4);
                        Options.MaxActionFailuresBeforeAbort = Math.Max(Options.MaxActionFailuresBeforeAbort, 4);
                        Options.ScreenPollingEnabled = true;
                        Options.ScreenPollInitialDelayMs = Math.Max(Options.ScreenPollInitialDelayMs, 250);
                        Options.ScreenPollIntervalMs = Math.Max(Options.ScreenPollIntervalMs, 250);
                        Options.ScreenPollTimeoutMs = Math.Max(Options.ScreenPollTimeoutMs, 3000);
                        Options.WaitNoChangeExtraMs = Math.Max(Options.WaitNoChangeExtraMs, 1500);
                        Options.MaxWaitSeconds = Math.Max(Options.MaxWaitSeconds, 120);
                        Options.MaxScreenshotSendWidth = 0;
                        Options.ScreenshotSendFormat = "png";
                        Options.FocusedOverviewMaxWidth = 0;
                        Options.MaxCropSendWidth = 0;
                        Options.CropSendFormat = "png";
                        Options.ScreenshotJpegQuality = 95;
                        Options.ScreenLogFormat = "png";
                        Options.MaxScreenLogWidth = 0;
                        Options.PrettyRequestLogs = true;
                        Options.IncludeFocusUiaCrop = true;
                        Options.DebugImages = true;
                        Options.VerifyMode = "always";
                        Options.UiaTargetNameMaxChars = Math.Max(Options.UiaTargetNameMaxChars, 96);
                        Options.UiaSummaryMaxChars = Math.Max(Options.UiaSummaryMaxChars, 800);
                        Options.UiaScanTimeBudgetMs = Math.Max(Options.UiaScanTimeBudgetMs, 180);
                        Options.MaxUiaNodesScanned = Math.Max(Options.MaxUiaNodesScanned, 1500);
                        Options.UiaCandidateMultiplier = Math.Max(Options.UiaCandidateMultiplier, 8);
                        Options.MaxUiaTargetAreaRatio = Math.Max(Options.MaxUiaTargetAreaRatio, 0.75);
                        break;
        
                }
            }

            static void ApplyCustomProfileDefaults()
            {
                CodeProfileDefaults ??= ProfileControlledSettingNames.ToDictionary(
                    name => name,
                    name => typeof(AppOptions)
                        .GetField(
                            name,
                            System.Reflection.BindingFlags.Instance |
                            System.Reflection.BindingFlags.NonPublic)
                        ?.GetValue(new AppOptions()),
                    StringComparer.Ordinal);

                Options.RunProfile = "custom";
                foreach (var (name, value) in CodeProfileDefaults)
                {
                    if (name == nameof(Options.ReasoningEffort) &&
                        Options.ReasoningEffortExplicit)
                    {
                        continue;
                    }
                    if (name == nameof(Options.QaReasoningEffort) &&
                        Options.QaReasoningEffortExplicit)
                    {
                        continue;
                    }
                    if (name == nameof(Options.VerifyReasoningEffort) &&
                        Options.VerifyReasoningEffortExplicit)
                    {
                        continue;
                    }
                    var field = typeof(AppOptions).GetField(
                        name,
                        System.Reflection.BindingFlags.Instance |
                        System.Reflection.BindingFlags.NonPublic);
                    field?.SetValue(Options, value);
                }
            }
        
            internal static void ApplyReasoningEffort(string? value, string source)
            {
                ApplyReasoningEffort(
                    value,
                    source,
                    v => Options.ReasoningEffort = v,
                    () => Options.ReasoningEffortExplicit = true);
            }

            internal static void ApplyApiUrl(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;

                if (Uri.TryCreate(value.Trim(), UriKind.Absolute, out var uri) &&
                    uri.Scheme is "http" or "https")
                {
                    Options.ApiUrl = uri.AbsoluteUri;
                    return;
                }

                ReportConfigurationError($"Invalid API URL '{value}' from {source}; expected an absolute HTTP or HTTPS URL.");
            }
        
            internal static void ApplyReasoningEffort(string? value, string source, Action<string?> setter, Action? markExplicit = null)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                markExplicit?.Invoke();
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "default" or "auto" or "off" or "null")
                {
                    setter(null);
                    return;
                }
        
                if (!AllowedReasoningEfforts.Contains(normalized))
                {
                    ReportConfigurationError($"Invalid reasoning effort '{value}' from {source}. Allowed: default, {string.Join(", ", AllowedReasoningEfforts.OrderBy(x => x))}.");
                    return;
                }
        
                setter(normalized);
            }
        
            internal static void ApplyDelayMs(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (int.TryParse(value, out var ms) && ms >= 0)
                    Options.UiSettleDelayMs = ms;
                else
                    ReportConfigurationError($"Invalid post-action delay '{value}' from {source}; expected a non-negative integer.");
            }
        
            internal static void ApplyGridStep(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (value.Equals("off", StringComparison.OrdinalIgnoreCase))
                {
                    Options.GridStepPx = 0;
                    return;
                }
        
                if (int.TryParse(value, out var px) && px >= 0)
                    Options.GridStepPx = px;
                else
                    ReportConfigurationError($"Invalid grid step '{value}' from {source}; expected a non-negative integer or 'off'.");
            }
        
            internal static void ApplyInt(string? value, string source, int min, int max, Action<int> setter)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (int.TryParse(value, out var parsed) && parsed >= min && parsed <= max)
                    setter(parsed);
                else
                    ReportConfigurationError($"Invalid integer '{value}' from {source}; expected {min}..{max}.");
            }
        
            internal static void ApplyLong(string? value, string source, long min, long max, Action<long> setter)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (long.TryParse(value, out var parsed) && parsed >= min && parsed <= max)
                    setter(parsed);
                else
                    ReportConfigurationError($"Invalid integer '{value}' from {source}; expected {min}..{max}.");
            }
        
            internal static void ApplyDouble(string? value, string source, double min, double max, Action<double> setter)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                value = value.Trim().Replace(',', '.');
                if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) &&
                    parsed >= min &&
                    parsed <= max)
                {
                    setter(parsed);
                }
                else
                {
                    ReportConfigurationError($"Invalid number '{value}' from {source}; expected {min:0.###}..{max:0.###}.");
                }
            }
        
            internal static void ApplyBool(string? value, string source, Action<bool> setter)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (IsTruthy(value))
                    setter(true);
                else if (IsFalsy(value))
                    setter(false);
                else
                    ReportConfigurationError($"Invalid boolean '{value}' from {source}; expected 1/true/yes/on or 0/false/no/off.");
            }
        
            internal static void ApplyTextVerbosity(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "low" or "medium" or "high")
                    Options.TextVerbosity = normalized;
                else
                    ReportConfigurationError($"Invalid text verbosity '{value}' from {source}; expected low, medium, or high.");
            }

            internal static void ApplyReasoningContext(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;

                var normalized = value.Trim().ToLowerInvariant() switch
                {
                    "current-turn" or "currentturn" or "current_turn" => "current_turn",
                    "all-turns" or "allturns" or "all_turns" => "all_turns",
                    "auto" => "auto",
                    _ => null
                };
                if (normalized is null)
                    ReportConfigurationError($"Invalid control reasoning context '{value}' from {source}; expected auto, current_turn, or all_turns.");
                else
                    Options.ControlReasoningContext = normalized;
            }

            internal static string NormalizeReasoningContext(string? value) =>
                value?.Trim().ToLowerInvariant() switch
                {
                    "current-turn" or "currentturn" => "current_turn",
                    "all-turns" or "allturns" => "all_turns",
                    "auto" => "auto",
                    "current_turn" => "current_turn",
                    "all_turns" => "all_turns",
                    _ => "all_turns"
                };

            internal static void ApplyGoalMode(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;

                var normalized = NormalizeGoalMode(value);
                if (normalized is "auto" or "finite" or "continuous")
                    Options.GoalMode = normalized;
                else
                    ReportConfigurationError($"Invalid goal mode '{value}' from {source}; expected auto, finite, or continuous.");
            }

            static string NormalizeGoalMode(string? value) =>
                value?.Trim().ToLowerInvariant() ?? "auto";

            internal static void ApplyObservationMode(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;

                var normalized = NormalizeObservationMode(value);
                if (IsAllowedObservationMode(normalized))
                    Options.ObservationMode = normalized;
                else
                    ReportConfigurationError($"Invalid observation profile '{value}' from {source}; expected auto, general, static_ui, local_editing, event_driven, streaming_output, turn_based_interaction, or realtime_interaction.");
            }

            static string NormalizeObservationMode(string? value) =>
                value?.Trim().ToLowerInvariant().Replace('-', '_') ?? "auto";

            static bool IsAllowedObservationMode(string value) =>
                value is "auto" or "general" or "static_ui" or "local_editing" or
                    "event_driven" or "streaming_output" or "turn_based_interaction" or "realtime_interaction";
        
            internal static void ApplyImageFormat(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "jpeg" or "jpg")
                    Options.ScreenshotSendFormat = "jpeg";
                else if (normalized == "png")
                    Options.ScreenshotSendFormat = "png";
                else
                    ReportConfigurationError($"Invalid screenshot format '{value}' from {source}; expected jpeg or png.");
            }
        
            internal static void ApplyCropFormat(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "jpeg" or "jpg")
                    Options.CropSendFormat = "jpeg";
                else if (normalized == "png")
                    Options.CropSendFormat = "png";
                else
                    ReportConfigurationError($"Invalid crop format '{value}' from {source}; expected jpeg or png.");
            }
        
            internal static void ApplyScreenLogFormat(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "jpeg" or "jpg")
                    Options.ScreenLogFormat = "jpeg";
                else if (normalized == "png")
                    Options.ScreenLogFormat = "png";
                else
                    ReportConfigurationError($"Invalid screen log format '{value}' from {source}; expected jpeg or png.");
            }
        
            internal static void ApplyVerifyMode(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "auto" or "always" or "off")
                    Options.VerifyMode = normalized;
                else
                    ReportConfigurationError($"Invalid verify mode '{value}' from {source}; expected auto, always, or off.");
            }
        
            internal static bool IsTruthy(string value) =>
                value.Equals("1", StringComparison.OrdinalIgnoreCase)
                || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                || value.Equals("yes", StringComparison.OrdinalIgnoreCase)
                || value.Equals("on", StringComparison.OrdinalIgnoreCase);
        
            internal static bool IsFalsy(string value) =>
                value.Equals("0", StringComparison.OrdinalIgnoreCase)
                || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                || value.Equals("no", StringComparison.OrdinalIgnoreCase)
                || value.Equals("off", StringComparison.OrdinalIgnoreCase);
        
            internal static string ReasoningEffortDisplay(string model, string? effort)
            {
                if (string.IsNullOrWhiteSpace(effort))
                    return "default";
        
                return SupportsReasoningEffort(model)
                    ? effort
                    : $"{effort} (not sent for model '{model}')";
            }
        
            internal static string ScreenshotMaxWidthDisplay() =>
                Options.MaxScreenshotSendWidth > 0 ? Options.MaxScreenshotSendWidth.ToString() : "original";
        
            internal static string FocusedOverviewMaxWidthDisplay() =>
                Options.FocusedOverviewMaxWidth > 0 ? Options.FocusedOverviewMaxWidth.ToString() : "normal";
        
            internal static int EffectiveFocusedOverviewMaxWidth() =>
                Options.FocusedOverviewMaxWidth > 0 ? Options.FocusedOverviewMaxWidth : Options.MaxScreenshotSendWidth;
        
            internal static string QaScreenshotMaxWidthDisplay() =>
                Options.QaScreenshotMaxWidth > 0 ? Options.QaScreenshotMaxWidth.ToString() : "normal";
        
            internal static string VerifyScreenshotMaxWidthDisplay() =>
                Options.VerifyScreenshotMaxWidth > 0 ? Options.VerifyScreenshotMaxWidth.ToString() : "normal";
        
            internal static string CropMaxWidthDisplay() =>
                Options.MaxCropSendWidth > 0 ? Options.MaxCropSendWidth.ToString() : "original";
        
            internal static string ScreenLogMaxWidthDisplay() =>
                Options.MaxScreenLogWidth > 0 ? Options.MaxScreenLogWidth.ToString() : "original";
        
            internal static string EffectiveQaModel() => string.IsNullOrWhiteSpace(Options.QaModel) ? Options.Model : Options.QaModel!;
            internal static string EffectiveVerifyModel() => string.IsNullOrWhiteSpace(Options.VerifyModel) ? Options.Model : Options.VerifyModel!;
            internal static string? EffectiveQaReasoningEffort() => Options.QaReasoningEffortExplicit ? Options.QaReasoningEffort : (Options.QaReasoningEffort ?? Options.ReasoningEffort);
            internal static string? EffectiveVerifyReasoningEffort() => Options.VerifyReasoningEffortExplicit ? Options.VerifyReasoningEffort : (Options.VerifyReasoningEffort ?? Options.ReasoningEffort);
        
            internal static void ConfigureOpenAiHttpClient()
            {
                // Each transport attempt owns a linked timeout token through body consumption.
                OpenAiTransport.HttpClient.Timeout = Timeout.InfiniteTimeSpan;
            }
    }
