internal static partial class RDPilotApplication
{
    /// <summary>
    /// Loads, validates, normalizes, and displays application configuration.
    /// </summary>
    internal static class ConfigurationService
    {
            internal static void PrintEffectiveConfig()
            {
                var uiMode = ForceRealUiOnly
                    ? "real-ui-only"
                    : (AllowHighLevelActions || AllowRunCommand ? "real UI + enabled local adapters" : "real UI");
        
                Console.WriteLine($"Profile: {RunProfile}");
                Console.WriteLine($"Model: {Model}");
                Console.WriteLine($"QA model: {EffectiveQaModel()}; verify model: {EffectiveVerifyModel()}");
                Console.WriteLine($"Reasoning effort: control={ReasoningEffortDisplay(Model, ReasoningEffort)}; qa={ReasoningEffortDisplay(EffectiveQaModel(), EffectiveQaReasoningEffort())}; verify={ReasoningEffortDisplay(EffectiveVerifyModel(), EffectiveVerifyReasoningEffort())}; adaptive={(AdaptiveReasoningEffort ? "on" : "off")}");
                Console.WriteLine($"UI mode: {uiMode}; mouse={(MouseEnabled ? "enabled" : "disabled")}; post-action delay={UiSettleDelayMs} ms; grid={(GridStepPx > 0 ? $"{GridStepPx}px" : "off")}");
                Console.WriteLine($"Images: send={ScreenshotSendFormat} max-width={ScreenshotMaxWidthDisplay()} focused-overview={FocusedOverviewMaxWidthDisplay()} qa={QaScreenshotMaxWidthDisplay()} verify={VerifyScreenshotMaxWidthDisplay()} quality={ScreenshotJpegQuality}; crop={CropSendFormat} max-width={CropMaxWidthDisplay()} size={FocusCropSize}px; screen-log={ScreenLogFormat} max-width={ScreenLogMaxWidthDisplay()}; focus_uia={(IncludeFocusUia ? "on" : "off")}; focus crop={(IncludeFocusUiaCrop ? "on" : "off")}; debug images={(DebugImages ? "on" : "off")}");
                Console.WriteLine($"Output: max_tokens={MaxOutputTokens}; qa_max_tokens={QaMaxOutputTokens}; verify_max_tokens={VerifyMaxOutputTokens}; incomplete_retries={IncompleteMaxOutputRetries}/{IncompleteMaxOutputTokenCap}; verbosity={TextVerbosity}; action_text_chars={MaxActionTextChars}; history_chars={HistoryTailChars}; history_lines={HistoryTailLines}; verify={VerifyMode}; verify_early_steps={VerifyEarlySteps}; verify_low_confidence={VerifyLowConfidenceThreshold:0.##}; skip_verify_confidence={SkipVerifyConfidenceThreshold:0.##}; verify-refresh={(RefreshScreenshotBeforeVerify ? "on" : "off")}; long text paste threshold={ClipboardPasteThreshold}; prompt_cache={(UsePromptCache ? PromptCacheKey ?? "on" : "off")}; previous_response_state={(UsePreviousResponseState ? "on" : "off")}; omit_unchanged_screen={(OmitUnchangedScreenImageWithState ? "on" : "off")}");
                Console.WriteLine($"Logs: requests={(LogRequests ? (PrettyRequestLogs ? "pretty" : "compact") : "off")}; screens={(LogScreens ? "on" : "state-only")}; retries={OpenAiMaxRetries}; timeout={(OpenAiTimeoutSeconds > 0 ? $"{OpenAiTimeoutSeconds}s" : "infinite")}; batch candidates={(ExecuteMultiActionCandidates ? $"on/{MaxQueuedBatchActions}" : "off")}");
                Console.WriteLine($"Loop guards: max_steps={MaxSteps}; max_wait={(MaxWaitSeconds > 0 ? $"{MaxWaitSeconds}s" : "off")}; stagnation={(MaxStagnationStepsBeforeAbort > 0 ? MaxStagnationStepsBeforeAbort.ToString() : "off")}; repeated_action={(MaxRepeatedActionBeforeAbort > 0 ? MaxRepeatedActionBeforeAbort.ToString() : "off")}; repeat_cooldown={(ActionRepeatCooldownSteps > 0 ? ActionRepeatCooldownSteps.ToString() : "off")}; model_failures={(MaxModelFailuresBeforeAbort > 0 ? MaxModelFailuresBeforeAbort.ToString() : "off")}; action_failures={(MaxActionFailuresBeforeAbort > 0 ? MaxActionFailuresBeforeAbort.ToString() : "off")}");
                Console.WriteLine($"Screen settle: polling={(ScreenPollingEnabled ? $"on initial={ScreenPollInitialDelayMs}ms interval={ScreenPollIntervalMs}ms timeout={ScreenPollTimeoutMs}ms wait_extra={WaitNoChangeExtraMs}ms" : "off")}; sanity={(ScreenSanityChecks ? "on" : "off")}");
                Console.WriteLine($"Input retry: sendinput_retries={SendInputMaxRetries}; retry_delay={SendInputRetryDelayMs}ms");
                Console.WriteLine($"Console: auto_hide={(AutoHideConsoleDuringRun ? "on" : "off")}; minimize_flag={(MinimizeConsoleDuringRun ? "on" : "off")}; restore_after_run={(RestoreConsoleAfterRun ? "on" : "off")}");
                Console.WriteLine($"UIA targets: {(IncludeUiaTargets ? $"on/{MaxUiaTargets}; name_chars={UiaTargetNameMaxChars}; summary_chars={UiaSummaryMaxChars}; scan={UiaScanTimeBudgetMs}ms/{MaxUiaNodesScanned} nodes; candidates={UiaCandidateMultiplier}x; max_area={MaxUiaTargetAreaRatio:0.##}; reuse={(ReuseUiaTargetsWhenScreenUnchanged ? "on" : "off")}" : "off")}");
                Console.WriteLine($"Local high-level actions: {(AllowHighLevelActions ? "enabled" : "disabled")}; run_command={(AllowRunCommand ? "enabled" : "disabled")}; real_ui_only={(ForceRealUiOnly ? "on" : "off")}");
            }

            internal static void PrintStartupSummary()
            {
                ConsoleTheme.WriteStartupBanner(
                    Model,
                    RunProfile,
                    ReasoningEffortDisplay(Model, ReasoningEffort),
                    EffectiveQaModel(),
                    EffectiveVerifyModel());
            }
        
            internal static void ApplyEnvironmentConfig()
            {
                ApplyProfile(RunProfile);
                ApplyConfigFile(Environment.GetEnvironmentVariable("RDPILOT_CONFIG"));
                ApplyEnvironmentProfileOverrides();
        
                var envModel = Environment.GetEnvironmentVariable("OPENAI_MODEL");
                if (!string.IsNullOrWhiteSpace(envModel))
                    Model = envModel.Trim();
                var envQaModel = Environment.GetEnvironmentVariable("OPENAI_QA_MODEL");
                if (!string.IsNullOrWhiteSpace(envQaModel))
                    QaModel = envQaModel.Trim();
                var envVerifyModel = Environment.GetEnvironmentVariable("OPENAI_VERIFY_MODEL");
                if (!string.IsNullOrWhiteSpace(envVerifyModel))
                    VerifyModel = envVerifyModel.Trim();
        
                var envEffort = Environment.GetEnvironmentVariable("OPENAI_REASONING_EFFORT")
                             ?? Environment.GetEnvironmentVariable("REASONING_EFFORT");
                ApplyReasoningEffort(envEffort, "OPENAI_REASONING_EFFORT/REASONING_EFFORT");
                ApplyReasoningEffort(Environment.GetEnvironmentVariable("OPENAI_QA_REASONING_EFFORT"), "OPENAI_QA_REASONING_EFFORT", v => QaReasoningEffort = v, () => QaReasoningEffortExplicit = true);
                ApplyReasoningEffort(Environment.GetEnvironmentVariable("OPENAI_VERIFY_REASONING_EFFORT"), "OPENAI_VERIFY_REASONING_EFFORT", v => VerifyReasoningEffort = v, () => VerifyReasoningEffortExplicit = true);
        
                var envME = Environment.GetEnvironmentVariable("MOUSE_ENABLED");
                if (!string.IsNullOrWhiteSpace(envME))
                    MouseEnabled = IsTruthy(envME);
        
                ApplyDelayMs(Environment.GetEnvironmentVariable("POST_ACTION_DELAY_MS"), "POST_ACTION_DELAY_MS");
                ApplyGridStep(Environment.GetEnvironmentVariable("GRID_STEP_PX"), "GRID_STEP_PX");
                ApplyInt(Environment.GetEnvironmentVariable("MAX_STEPS"), "MAX_STEPS", 1, int.MaxValue, v => MaxSteps = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_WAIT_SECONDS"), "MAX_WAIT_SECONDS", 0, int.MaxValue, v => MaxWaitSeconds = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_OUTPUT_TOKENS"), "MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => MaxOutputTokens = v);
                ApplyInt(Environment.GetEnvironmentVariable("QA_MAX_OUTPUT_TOKENS"), "QA_MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => QaMaxOutputTokens = v);
                ApplyInt(Environment.GetEnvironmentVariable("VERIFY_MAX_OUTPUT_TOKENS"), "VERIFY_MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => VerifyMaxOutputTokens = v);
                ApplyInt(Environment.GetEnvironmentVariable("INCOMPLETE_MAX_OUTPUT_RETRIES"), "INCOMPLETE_MAX_OUTPUT_RETRIES", 0, 5, v => IncompleteMaxOutputRetries = v);
                ApplyInt(Environment.GetEnvironmentVariable("INCOMPLETE_MAX_OUTPUT_TOKEN_CAP"), "INCOMPLETE_MAX_OUTPUT_TOKEN_CAP", 1, int.MaxValue, v => IncompleteMaxOutputTokenCap = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_ACTION_TEXT_CHARS"), "MAX_ACTION_TEXT_CHARS", 256, int.MaxValue, v => MaxActionTextChars = v);
                ApplyInt(Environment.GetEnvironmentVariable("QA_SCREENSHOT_MAX_WIDTH"), "QA_SCREENSHOT_MAX_WIDTH", 0, 10000, v => QaScreenshotMaxWidth = v);
                ApplyInt(Environment.GetEnvironmentVariable("VERIFY_SCREENSHOT_MAX_WIDTH"), "VERIFY_SCREENSHOT_MAX_WIDTH", 0, 10000, v => VerifyScreenshotMaxWidth = v);
                ApplyTextVerbosity(Environment.GetEnvironmentVariable("TEXT_VERBOSITY"), "TEXT_VERBOSITY");
                ApplyInt(Environment.GetEnvironmentVariable("HISTORY_TAIL_CHARS"), "HISTORY_TAIL_CHARS", 0, int.MaxValue, v => HistoryTailChars = v);
                ApplyInt(Environment.GetEnvironmentVariable("HISTORY_TAIL_LINES"), "HISTORY_TAIL_LINES", 0, int.MaxValue, v => HistoryTailLines = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_STAGNATION_STEPS"), "MAX_STAGNATION_STEPS", 0, int.MaxValue, v => MaxStagnationStepsBeforeAbort = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_REPEATED_ACTIONS"), "MAX_REPEATED_ACTIONS", 0, int.MaxValue, v => MaxRepeatedActionBeforeAbort = v);
                ApplyInt(Environment.GetEnvironmentVariable("ACTION_REPEAT_COOLDOWN_STEPS"), "ACTION_REPEAT_COOLDOWN_STEPS", 0, int.MaxValue, v => ActionRepeatCooldownSteps = v);
                ApplyDouble(Environment.GetEnvironmentVariable("SKIP_VERIFY_CONFIDENCE_THRESHOLD"), "SKIP_VERIFY_CONFIDENCE_THRESHOLD", 0.0, 1.0, v => SkipVerifyConfidenceThreshold = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_MODEL_FAILURES"), "MAX_MODEL_FAILURES", 0, int.MaxValue, v => MaxModelFailuresBeforeAbort = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_ACTION_FAILURES"), "MAX_ACTION_FAILURES", 0, int.MaxValue, v => MaxActionFailuresBeforeAbort = v);
                ApplyBool(Environment.GetEnvironmentVariable("SCREEN_POLLING"), "SCREEN_POLLING", v => ScreenPollingEnabled = v);
                ApplyInt(Environment.GetEnvironmentVariable("SCREEN_POLL_INITIAL_DELAY_MS"), "SCREEN_POLL_INITIAL_DELAY_MS", 0, 10000, v => ScreenPollInitialDelayMs = v);
                ApplyInt(Environment.GetEnvironmentVariable("SCREEN_POLL_INTERVAL_MS"), "SCREEN_POLL_INTERVAL_MS", 10, 10000, v => ScreenPollIntervalMs = v);
                ApplyInt(Environment.GetEnvironmentVariable("SCREEN_POLL_TIMEOUT_MS"), "SCREEN_POLL_TIMEOUT_MS", 0, 60000, v => ScreenPollTimeoutMs = v);
                ApplyInt(Environment.GetEnvironmentVariable("WAIT_NO_CHANGE_EXTRA_MS"), "WAIT_NO_CHANGE_EXTRA_MS", 0, 60000, v => WaitNoChangeExtraMs = v);
                ApplyBool(Environment.GetEnvironmentVariable("SCREEN_SANITY_CHECKS"), "SCREEN_SANITY_CHECKS", v => ScreenSanityChecks = v);
                ApplyInt(Environment.GetEnvironmentVariable("SENDINPUT_MAX_RETRIES"), "SENDINPUT_MAX_RETRIES", 0, 10, v => SendInputMaxRetries = v);
                ApplyInt(Environment.GetEnvironmentVariable("SENDINPUT_RETRY_DELAY_MS"), "SENDINPUT_RETRY_DELAY_MS", 0, 1000, v => SendInputRetryDelayMs = v);
                ApplyBool(Environment.GetEnvironmentVariable("ADAPTIVE_REASONING_EFFORT"), "ADAPTIVE_REASONING_EFFORT", v => AdaptiveReasoningEffort = v);
                ApplyInt(Environment.GetEnvironmentVariable("SCREENSHOT_MAX_WIDTH"), "SCREENSHOT_MAX_WIDTH", 0, 10000, v => MaxScreenshotSendWidth = v);
                ApplyImageFormat(Environment.GetEnvironmentVariable("SCREENSHOT_FORMAT"), "SCREENSHOT_FORMAT");
                ApplyInt(Environment.GetEnvironmentVariable("FOCUSED_OVERVIEW_MAX_WIDTH"), "FOCUSED_OVERVIEW_MAX_WIDTH", 0, 10000, v => FocusedOverviewMaxWidth = v);
                ApplyInt(Environment.GetEnvironmentVariable("CROP_MAX_WIDTH"), "CROP_MAX_WIDTH", 0, 10000, v => MaxCropSendWidth = v);
                ApplyCropFormat(Environment.GetEnvironmentVariable("CROP_FORMAT"), "CROP_FORMAT");
                ApplyLong(Environment.GetEnvironmentVariable("SCREENSHOT_JPEG_QUALITY"), "SCREENSHOT_JPEG_QUALITY", 1, 100, v => ScreenshotJpegQuality = v);
                ApplyScreenLogFormat(Environment.GetEnvironmentVariable("SCREEN_LOG_FORMAT"), "SCREEN_LOG_FORMAT");
                ApplyInt(Environment.GetEnvironmentVariable("SCREEN_LOG_MAX_WIDTH"), "SCREEN_LOG_MAX_WIDTH", 0, 10000, v => MaxScreenLogWidth = v);
                ApplyBool(Environment.GetEnvironmentVariable("DEBUG_IMAGES"), "DEBUG_IMAGES", v => DebugImages = v);
                ApplyBool(Environment.GetEnvironmentVariable("LOG_REQUESTS"), "LOG_REQUESTS", v => LogRequests = v);
                ApplyBool(Environment.GetEnvironmentVariable("PRETTY_REQUEST_LOGS"), "PRETTY_REQUEST_LOGS", v => PrettyRequestLogs = v);
                ApplyBool(Environment.GetEnvironmentVariable("LOG_SCREENS"), "LOG_SCREENS", v => LogScreens = v);
                ApplyBool(Environment.GetEnvironmentVariable("INCLUDE_FOCUS_UIA"), "INCLUDE_FOCUS_UIA", v => IncludeFocusUia = v);
                ApplyBool(Environment.GetEnvironmentVariable("INCLUDE_FOCUS_UIA_CROP"), "INCLUDE_FOCUS_UIA_CROP", v => IncludeFocusUiaCrop = v);
                ApplyVerifyMode(Environment.GetEnvironmentVariable("VERIFY_MODE"), "VERIFY_MODE");
                ApplyInt(Environment.GetEnvironmentVariable("VERIFY_EARLY_STEPS"), "VERIFY_EARLY_STEPS", 0, int.MaxValue, v => VerifyEarlySteps = v);
                ApplyDouble(Environment.GetEnvironmentVariable("VERIFY_LOW_CONFIDENCE_THRESHOLD"), "VERIFY_LOW_CONFIDENCE_THRESHOLD", 0.0, 1.0, v => VerifyLowConfidenceThreshold = v);
                ApplyBool(Environment.GetEnvironmentVariable("REFRESH_SCREENSHOT_BEFORE_VERIFY"), "REFRESH_SCREENSHOT_BEFORE_VERIFY", v => RefreshScreenshotBeforeVerify = v);
                ApplyInt(Environment.GetEnvironmentVariable("CLIPBOARD_PASTE_THRESHOLD"), "CLIPBOARD_PASTE_THRESHOLD", 0, int.MaxValue, v => ClipboardPasteThreshold = v);
                ApplyInt(Environment.GetEnvironmentVariable("FOCUS_CROP_SIZE"), "FOCUS_CROP_SIZE", 64, 2000, v => FocusCropSize = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_FOCUS_UIA_CROP_PIXELS"), "MAX_FOCUS_UIA_CROP_PIXELS", 0, int.MaxValue, v => MaxFocusUiaCropPixels = v);
                ApplyInt(Environment.GetEnvironmentVariable("OPENAI_MAX_RETRIES"), "OPENAI_MAX_RETRIES", 0, 10, v => OpenAiMaxRetries = v);
                ApplyInt(Environment.GetEnvironmentVariable("OPENAI_TIMEOUT_SECONDS"), "OPENAI_TIMEOUT_SECONDS", 0, int.MaxValue, v => OpenAiTimeoutSeconds = v);
                ApplyBool(Environment.GetEnvironmentVariable("REAL_UI_ONLY"), "REAL_UI_ONLY", v => ForceRealUiOnly = v);
                ApplyBool(Environment.GetEnvironmentVariable("ALLOW_HIGH_LEVEL_ACTIONS"), "ALLOW_HIGH_LEVEL_ACTIONS", v => AllowHighLevelActions = v);
                ApplyBool(Environment.GetEnvironmentVariable("ALLOW_RUN_COMMAND"), "ALLOW_RUN_COMMAND", v => AllowRunCommand = v);
                ApplyBool(Environment.GetEnvironmentVariable("DIRECT_CLICK_WITHOUT_AIM"), "DIRECT_CLICK_WITHOUT_AIM", v => DirectClickWithoutAim = v);
                ApplyBool(Environment.GetEnvironmentVariable("AUTO_HIDE_CONSOLE"), "AUTO_HIDE_CONSOLE", v => AutoHideConsoleDuringRun = v);
                ApplyBool(Environment.GetEnvironmentVariable("MINIMIZE_CONSOLE_DURING_RUN"), "MINIMIZE_CONSOLE_DURING_RUN", v => MinimizeConsoleDuringRun = v);
                ApplyBool(Environment.GetEnvironmentVariable("RESTORE_CONSOLE_AFTER_RUN"), "RESTORE_CONSOLE_AFTER_RUN", v => RestoreConsoleAfterRun = v);
                ApplyBool(Environment.GetEnvironmentVariable("PROMPT_CACHE"), "PROMPT_CACHE", v => UsePromptCache = v);
                ApplyBool(Environment.GetEnvironmentVariable("USE_PREVIOUS_RESPONSE_ID"), "USE_PREVIOUS_RESPONSE_ID", v => UsePreviousResponseState = v);
                ApplyBool(Environment.GetEnvironmentVariable("OMIT_UNCHANGED_SCREEN_IMAGE"), "OMIT_UNCHANGED_SCREEN_IMAGE", v => OmitUnchangedScreenImageWithState = v);
                ApplyBool(Environment.GetEnvironmentVariable("INCLUDE_UIA_TARGETS"), "INCLUDE_UIA_TARGETS", v => IncludeUiaTargets = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_UIA_TARGETS"), "MAX_UIA_TARGETS", 0, 100, v => MaxUiaTargets = v);
                ApplyInt(Environment.GetEnvironmentVariable("UIA_TARGET_NAME_MAX_CHARS"), "UIA_TARGET_NAME_MAX_CHARS", 0, 500, v => UiaTargetNameMaxChars = v);
                ApplyInt(Environment.GetEnvironmentVariable("UIA_SUMMARY_MAX_CHARS"), "UIA_SUMMARY_MAX_CHARS", 0, 2000, v => UiaSummaryMaxChars = v);
                ApplyInt(Environment.GetEnvironmentVariable("UIA_SCAN_TIME_BUDGET_MS"), "UIA_SCAN_TIME_BUDGET_MS", 0, 5000, v => UiaScanTimeBudgetMs = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_UIA_NODES_SCANNED"), "MAX_UIA_NODES_SCANNED", 0, 10000, v => MaxUiaNodesScanned = v);
                ApplyInt(Environment.GetEnvironmentVariable("UIA_CANDIDATE_MULTIPLIER"), "UIA_CANDIDATE_MULTIPLIER", 1, 20, v => UiaCandidateMultiplier = v);
                ApplyDouble(Environment.GetEnvironmentVariable("UIA_MAX_AREA_RATIO"), "UIA_MAX_AREA_RATIO", 0.0, 1.0, v => MaxUiaTargetAreaRatio = v);
                ApplyBool(Environment.GetEnvironmentVariable("REUSE_UIA_TARGETS_ON_NO_CHANGE"), "REUSE_UIA_TARGETS_ON_NO_CHANGE", v => ReuseUiaTargetsWhenScreenUnchanged = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_ARTIFACTS_PER_DIR"), "MAX_ARTIFACTS_PER_DIR", 0, int.MaxValue, v => MaxArtifactsPerDir = v);
                ApplyBool(Environment.GetEnvironmentVariable("EXECUTE_MULTI_ACTION_CANDIDATES"), "EXECUTE_MULTI_ACTION_CANDIDATES", v => ExecuteMultiActionCandidates = v);
                ApplyInt(Environment.GetEnvironmentVariable("MAX_QUEUED_BATCH_ACTIONS"), "MAX_QUEUED_BATCH_ACTIONS", 0, 20, v => MaxQueuedBatchActions = v);
                var envPromptCacheKey = Environment.GetEnvironmentVariable("PROMPT_CACHE_KEY");
                if (!string.IsNullOrWhiteSpace(envPromptCacheKey))
                    PromptCacheKey = envPromptCacheKey.Trim();
            }
        
            internal static void ApplyEnvironmentProfileOverrides()
            {
                var profile = Environment.GetEnvironmentVariable("RDPILOT_PROFILE");
                var fastMode = IsTruthyEnv("FAST_MODE");
                var balancedMode = IsTruthyEnv("BALANCED_MODE");
                var qualityMode = IsTruthyEnv("QUALITY_MODE");
        
                if (string.IsNullOrWhiteSpace(profile))
                {
                    if (qualityMode) profile = "quality";
                    else if (balancedMode) profile = "balanced";
                    else if (fastMode) profile = "fast";
                }
                else if (fastMode || balancedMode || qualityMode)
                {
                    Console.Error.WriteLine("RDPILOT_PROFILE overrides FAST_MODE/BALANCED_MODE/QUALITY_MODE.");
                }
        
                if (!string.IsNullOrWhiteSpace(profile))
                    ApplyProfile(profile);
            }
        
            internal static bool IsTruthyEnv(string name)
            {
                var value = Environment.GetEnvironmentVariable(name);
                return !string.IsNullOrWhiteSpace(value) && IsTruthy(value);
            }
        
            internal static string? ApplyCliArgs(string[] args)
            {
                ApplyCliProfileArgs(args);
                var positional = new List<string>();
        
                for (var i = 0; i < args.Length; i++)
                {
                    var arg = args[i];
                    if (string.IsNullOrWhiteSpace(arg))
                        continue;
        
                    if (arg.Equals("--mouse", StringComparison.OrdinalIgnoreCase))
                    {
                        MouseEnabled = true;
                        continue;
                    }
                    if (arg.Equals("--no-mouse", StringComparison.OrdinalIgnoreCase))
                    {
                        MouseEnabled = false;
                        continue;
                    }
                    if (arg.Equals("--fast", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (arg.Equals("--balanced", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (arg.Equals("--quality", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (arg.Equals("--debug-images", StringComparison.OrdinalIgnoreCase))
                    {
                        DebugImages = true;
                        IncludeFocusUia = true;
                        IncludeFocusUiaCrop = true;
                        continue;
                    }
                    if (arg.Equals("--focus-uia", StringComparison.OrdinalIgnoreCase))
                    {
                        IncludeFocusUia = true;
                        continue;
                    }
                    if (arg.Equals("--no-focus-uia", StringComparison.OrdinalIgnoreCase))
                    {
                        IncludeFocusUia = false;
                        IncludeFocusUiaCrop = false;
                        continue;
                    }
                    if (arg.Equals("--no-request-logs", StringComparison.OrdinalIgnoreCase))
                    {
                        LogRequests = false;
                        continue;
                    }
                    if (arg.Equals("--pretty-request-logs", StringComparison.OrdinalIgnoreCase))
                    {
                        PrettyRequestLogs = true;
                        continue;
                    }
                    if (arg.Equals("--compact-request-logs", StringComparison.OrdinalIgnoreCase))
                    {
                        PrettyRequestLogs = false;
                        continue;
                    }
                    if (arg.Equals("--no-screen-logs", StringComparison.OrdinalIgnoreCase))
                    {
                        LogScreens = false;
                        continue;
                    }
                    if (arg.Equals("--real-ui-only", StringComparison.OrdinalIgnoreCase))
                    {
                        ForceRealUiOnly = true;
                        continue;
                    }
                    if (arg.Equals("--allow-run-command", StringComparison.OrdinalIgnoreCase))
                    {
                        ForceRealUiOnly = false;
                        AllowRunCommand = true;
                        continue;
                    }
                    if (arg.Equals("--allow-high-level-actions", StringComparison.OrdinalIgnoreCase))
                    {
                        ForceRealUiOnly = false;
                        AllowHighLevelActions = true;
                        continue;
                    }
                    if (arg.Equals("--no-direct-click", StringComparison.OrdinalIgnoreCase))
                    {
                        DirectClickWithoutAim = false;
                        continue;
                    }
                    if (arg.Equals("--minimize-console", StringComparison.OrdinalIgnoreCase))
                    {
                        MinimizeConsoleDuringRun = true;
                        continue;
                    }
                    if (arg.Equals("--auto-hide-console", StringComparison.OrdinalIgnoreCase))
                    {
                        AutoHideConsoleDuringRun = true;
                        continue;
                    }
                    if (arg.Equals("--no-auto-hide-console", StringComparison.OrdinalIgnoreCase))
                    {
                        AutoHideConsoleDuringRun = false;
                        continue;
                    }
                    if (arg.Equals("--restore-console", StringComparison.OrdinalIgnoreCase))
                    {
                        RestoreConsoleAfterRun = true;
                        continue;
                    }
                    if (arg.Equals("--no-restore-console", StringComparison.OrdinalIgnoreCase))
                    {
                        RestoreConsoleAfterRun = false;
                        continue;
                    }
                    if (arg.Equals("--analyze-logs", StringComparison.OrdinalIgnoreCase))
                    {
                        AnalyzeLogsOnly = true;
                        continue;
                    }
                    if (arg.Equals("--print-config", StringComparison.OrdinalIgnoreCase))
                    {
                        PrintConfigOnly = true;
                        continue;
                    }
                    if (arg.Equals("--batch-candidates", StringComparison.OrdinalIgnoreCase))
                    {
                        ExecuteMultiActionCandidates = true;
                        continue;
                    }
                    if (arg.Equals("--no-batch-candidates", StringComparison.OrdinalIgnoreCase))
                    {
                        ExecuteMultiActionCandidates = false;
                        continue;
                    }
                    if (arg.Equals("--refresh-before-verify", StringComparison.OrdinalIgnoreCase))
                    {
                        RefreshScreenshotBeforeVerify = true;
                        continue;
                    }
                    if (arg.Equals("--no-refresh-before-verify", StringComparison.OrdinalIgnoreCase))
                    {
                        RefreshScreenshotBeforeVerify = false;
                        continue;
                    }
                    if (arg.Equals("--adaptive-effort", StringComparison.OrdinalIgnoreCase))
                    {
                        AdaptiveReasoningEffort = true;
                        continue;
                    }
                    if (arg.Equals("--no-adaptive-effort", StringComparison.OrdinalIgnoreCase))
                    {
                        AdaptiveReasoningEffort = false;
                        continue;
                    }
                    if (arg.Equals("--screen-polling", StringComparison.OrdinalIgnoreCase))
                    {
                        ScreenPollingEnabled = true;
                        continue;
                    }
                    if (arg.Equals("--no-screen-polling", StringComparison.OrdinalIgnoreCase))
                    {
                        ScreenPollingEnabled = false;
                        continue;
                    }
                    if (arg.Equals("--screen-sanity", StringComparison.OrdinalIgnoreCase))
                    {
                        ScreenSanityChecks = true;
                        continue;
                    }
                    if (arg.Equals("--no-screen-sanity", StringComparison.OrdinalIgnoreCase))
                    {
                        ScreenSanityChecks = false;
                        continue;
                    }
                    if (arg.Equals("--prompt-cache", StringComparison.OrdinalIgnoreCase))
                    {
                        UsePromptCache = true;
                        continue;
                    }
                    if (arg.Equals("--no-prompt-cache", StringComparison.OrdinalIgnoreCase))
                    {
                        UsePromptCache = false;
                        continue;
                    }
                    if (arg.Equals("--previous-response-state", StringComparison.OrdinalIgnoreCase))
                    {
                        UsePreviousResponseState = true;
                        continue;
                    }
                    if (arg.Equals("--no-previous-response-state", StringComparison.OrdinalIgnoreCase))
                    {
                        UsePreviousResponseState = false;
                        continue;
                    }
                    if (arg.Equals("--omit-unchanged-screen", StringComparison.OrdinalIgnoreCase))
                    {
                        OmitUnchangedScreenImageWithState = true;
                        continue;
                    }
                    if (arg.Equals("--no-omit-unchanged-screen", StringComparison.OrdinalIgnoreCase))
                    {
                        OmitUnchangedScreenImageWithState = false;
                        continue;
                    }
                    if (arg.Equals("--reuse-uia-targets", StringComparison.OrdinalIgnoreCase))
                    {
                        ReuseUiaTargetsWhenScreenUnchanged = true;
                        continue;
                    }
                    if (arg.Equals("--no-reuse-uia-targets", StringComparison.OrdinalIgnoreCase))
                    {
                        ReuseUiaTargetsWhenScreenUnchanged = false;
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--profile", out var profile))
                    {
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--model", out var model))
                    {
                        if (!string.IsNullOrWhiteSpace(model))
                            Model = model.Trim();
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--qa-model", out var qaModel))
                    {
                        if (!string.IsNullOrWhiteSpace(qaModel))
                            QaModel = qaModel.Trim();
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify-model", out var verifyModel))
                    {
                        if (!string.IsNullOrWhiteSpace(verifyModel))
                            VerifyModel = verifyModel.Trim();
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--effort", out var effort) ||
                        TryReadOption(args, ref i, "--reasoning-effort", out effort))
                    {
                        ApplyReasoningEffort(effort, "--effort");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--qa-effort", out var qaEffort))
                    {
                        ApplyReasoningEffort(qaEffort, "--qa-effort", v => QaReasoningEffort = v, () => QaReasoningEffortExplicit = true);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify-effort", out var verifyEffort))
                    {
                        ApplyReasoningEffort(verifyEffort, "--verify-effort", v => VerifyReasoningEffort = v, () => VerifyReasoningEffortExplicit = true);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--delay", out var delay))
                    {
                        ApplyDelayMs(delay, "--delay");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--grid", out var grid))
                    {
                        ApplyGridStep(grid, "--grid");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-steps", out var maxSteps))
                    {
                        ApplyInt(maxSteps, "--max-steps", 1, int.MaxValue, v => MaxSteps = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-wait", out var maxWait))
                    {
                        ApplyInt(maxWait, "--max-wait", 0, int.MaxValue, v => MaxWaitSeconds = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-output-tokens", out var maxOutput))
                    {
                        ApplyInt(maxOutput, "--max-output-tokens", 1, int.MaxValue, v => MaxOutputTokens = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--qa-max-output-tokens", out var qaMaxOutput))
                    {
                        ApplyInt(qaMaxOutput, "--qa-max-output-tokens", 1, int.MaxValue, v => QaMaxOutputTokens = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify-max-output-tokens", out var verifyMaxOutput))
                    {
                        ApplyInt(verifyMaxOutput, "--verify-max-output-tokens", 1, int.MaxValue, v => VerifyMaxOutputTokens = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--incomplete-max-output-retries", out var incompleteRetries))
                    {
                        ApplyInt(incompleteRetries, "--incomplete-max-output-retries", 0, 5, v => IncompleteMaxOutputRetries = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--incomplete-max-output-token-cap", out var incompleteTokenCap))
                    {
                        ApplyInt(incompleteTokenCap, "--incomplete-max-output-token-cap", 1, int.MaxValue, v => IncompleteMaxOutputTokenCap = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-action-text-chars", out var maxActionTextChars))
                    {
                        ApplyInt(maxActionTextChars, "--max-action-text-chars", 256, int.MaxValue, v => MaxActionTextChars = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--qa-screenshot-max-width", out var qaScreenshotMaxWidth))
                    {
                        ApplyInt(qaScreenshotMaxWidth, "--qa-screenshot-max-width", 0, 10000, v => QaScreenshotMaxWidth = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify-screenshot-max-width", out var verifyScreenshotMaxWidth))
                    {
                        ApplyInt(verifyScreenshotMaxWidth, "--verify-screenshot-max-width", 0, 10000, v => VerifyScreenshotMaxWidth = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verbosity", out var verbosity))
                    {
                        ApplyTextVerbosity(verbosity, "--verbosity");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--history-chars", out var historyChars))
                    {
                        ApplyInt(historyChars, "--history-chars", 0, int.MaxValue, v => HistoryTailChars = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--history-lines", out var historyLines))
                    {
                        ApplyInt(historyLines, "--history-lines", 0, int.MaxValue, v => HistoryTailLines = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-stagnation", out var maxStagnation))
                    {
                        ApplyInt(maxStagnation, "--max-stagnation", 0, int.MaxValue, v => MaxStagnationStepsBeforeAbort = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-repeated-actions", out var maxRepeated))
                    {
                        ApplyInt(maxRepeated, "--max-repeated-actions", 0, int.MaxValue, v => MaxRepeatedActionBeforeAbort = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--repeat-cooldown", out var repeatCooldown))
                    {
                        ApplyInt(repeatCooldown, "--repeat-cooldown", 0, int.MaxValue, v => ActionRepeatCooldownSteps = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--skip-verify-confidence", out var skipVerifyConfidence))
                    {
                        ApplyDouble(skipVerifyConfidence, "--skip-verify-confidence", 0.0, 1.0, v => SkipVerifyConfidenceThreshold = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-model-failures", out var maxModelFailures))
                    {
                        ApplyInt(maxModelFailures, "--max-model-failures", 0, int.MaxValue, v => MaxModelFailuresBeforeAbort = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-action-failures", out var maxActionFailures))
                    {
                        ApplyInt(maxActionFailures, "--max-action-failures", 0, int.MaxValue, v => MaxActionFailuresBeforeAbort = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screen-poll-initial-delay", out var pollInitialDelay))
                    {
                        ApplyInt(pollInitialDelay, "--screen-poll-initial-delay", 0, 10000, v => ScreenPollInitialDelayMs = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screen-poll-interval", out var pollInterval))
                    {
                        ApplyInt(pollInterval, "--screen-poll-interval", 10, 10000, v => ScreenPollIntervalMs = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screen-poll-timeout", out var pollTimeout))
                    {
                        ApplyInt(pollTimeout, "--screen-poll-timeout", 0, 60000, v => ScreenPollTimeoutMs = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--wait-no-change-extra", out var waitNoChangeExtra))
                    {
                        ApplyInt(waitNoChangeExtra, "--wait-no-change-extra", 0, 60000, v => WaitNoChangeExtraMs = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--sendinput-retries", out var sendInputRetries))
                    {
                        ApplyInt(sendInputRetries, "--sendinput-retries", 0, 10, v => SendInputMaxRetries = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--sendinput-retry-delay", out var sendInputRetryDelay))
                    {
                        ApplyInt(sendInputRetryDelay, "--sendinput-retry-delay", 0, 1000, v => SendInputRetryDelayMs = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screenshot-max-width", out var maxWidth))
                    {
                        ApplyInt(maxWidth, "--screenshot-max-width", 0, 10000, v => MaxScreenshotSendWidth = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screenshot-format", out var imageFormat))
                    {
                        ApplyImageFormat(imageFormat, "--screenshot-format");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--focused-overview-max-width", out var focusedOverviewMaxWidth))
                    {
                        ApplyInt(focusedOverviewMaxWidth, "--focused-overview-max-width", 0, 10000, v => FocusedOverviewMaxWidth = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--crop-max-width", out var cropMaxWidth))
                    {
                        ApplyInt(cropMaxWidth, "--crop-max-width", 0, 10000, v => MaxCropSendWidth = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--crop-format", out var cropFormat))
                    {
                        ApplyCropFormat(cropFormat, "--crop-format");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--jpeg-quality", out var jpegQuality))
                    {
                        ApplyLong(jpegQuality, "--jpeg-quality", 1, 100, v => ScreenshotJpegQuality = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screen-log-format", out var screenLogFormat))
                    {
                        ApplyScreenLogFormat(screenLogFormat, "--screen-log-format");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--screen-log-max-width", out var screenLogMaxWidth))
                    {
                        ApplyInt(screenLogMaxWidth, "--screen-log-max-width", 0, 10000, v => MaxScreenLogWidth = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify", out var verifyMode))
                    {
                        ApplyVerifyMode(verifyMode, "--verify");
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify-early-steps", out var verifyEarlySteps))
                    {
                        ApplyInt(verifyEarlySteps, "--verify-early-steps", 0, int.MaxValue, v => VerifyEarlySteps = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--verify-low-confidence", out var verifyLowConfidence))
                    {
                        ApplyDouble(verifyLowConfidence, "--verify-low-confidence", 0.0, 1.0, v => VerifyLowConfidenceThreshold = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--paste-threshold", out var pasteThreshold))
                    {
                        ApplyInt(pasteThreshold, "--paste-threshold", 0, int.MaxValue, v => ClipboardPasteThreshold = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--focus-crop-size", out var focusCropSize))
                    {
                        ApplyInt(focusCropSize, "--focus-crop-size", 64, 2000, v => FocusCropSize = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-focus-crop-pixels", out var maxFocusPixels))
                    {
                        ApplyInt(maxFocusPixels, "--max-focus-crop-pixels", 0, int.MaxValue, v => MaxFocusUiaCropPixels = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-uia-targets", out var maxUiaTargets))
                    {
                        ApplyInt(maxUiaTargets, "--max-uia-targets", 0, 100, v => MaxUiaTargets = v);
                        IncludeUiaTargets = MaxUiaTargets > 0;
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--uia-name-chars", out var uiaNameChars))
                    {
                        ApplyInt(uiaNameChars, "--uia-name-chars", 0, 500, v => UiaTargetNameMaxChars = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--uia-summary-chars", out var uiaSummaryChars))
                    {
                        ApplyInt(uiaSummaryChars, "--uia-summary-chars", 0, 2000, v => UiaSummaryMaxChars = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--uia-scan-ms", out var uiaScanMs))
                    {
                        ApplyInt(uiaScanMs, "--uia-scan-ms", 0, 5000, v => UiaScanTimeBudgetMs = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-uia-nodes", out var maxUiaNodes))
                    {
                        ApplyInt(maxUiaNodes, "--max-uia-nodes", 0, 10000, v => MaxUiaNodesScanned = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--uia-candidate-multiplier", out var uiaCandidateMultiplier))
                    {
                        ApplyInt(uiaCandidateMultiplier, "--uia-candidate-multiplier", 1, 20, v => UiaCandidateMultiplier = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--uia-max-area-ratio", out var uiaMaxAreaRatio))
                    {
                        ApplyDouble(uiaMaxAreaRatio, "--uia-max-area-ratio", 0.0, 1.0, v => MaxUiaTargetAreaRatio = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-artifacts", out var maxArtifacts))
                    {
                        ApplyInt(maxArtifacts, "--max-artifacts", 0, int.MaxValue, v => MaxArtifactsPerDir = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--max-batch-actions", out var maxBatchActions))
                    {
                        ApplyInt(maxBatchActions, "--max-batch-actions", 0, 20, v => MaxQueuedBatchActions = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--replay-response", out var replayPath))
                    {
                        ReplayResponsePath = replayPath;
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--replay-request", out var replayRequestPath))
                    {
                        ReplayRequestPath = replayRequestPath;
                        continue;
                    }
                    if (arg.Equals("--replay-request-dry-run", StringComparison.OrdinalIgnoreCase))
                    {
                        ReplayRequestDryRun = true;
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--retries", out var retries))
                    {
                        ApplyInt(retries, "--retries", 0, 10, v => OpenAiMaxRetries = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--openai-timeout", out var openAiTimeout))
                    {
                        ApplyInt(openAiTimeout, "--openai-timeout", 0, int.MaxValue, v => OpenAiTimeoutSeconds = v);
                        continue;
                    }
                    if (TryReadOption(args, ref i, "--prompt-cache-key", out var promptCacheKey))
                    {
                        PromptCacheKey = string.IsNullOrWhiteSpace(promptCacheKey) ? null : promptCacheKey.Trim();
                        continue;
                    }
        
                    if (arg.StartsWith("--", StringComparison.Ordinal))
                    {
                        Console.Error.WriteLine($"Unknown option ignored: {arg}");
                        continue;
                    }
        
                    positional.Add(arg);
                }
        
                return positional.Count == 0 ? null : string.Join(" ", positional);
            }
        
            internal static void NormalizeConfig()
            {
                if (ForceRealUiOnly)
                {
                    AllowHighLevelActions = false;
                    AllowRunCommand = false;
                }
                if (!IncludeFocusUia)
                    IncludeFocusUiaCrop = false;
                if (MaxUiaTargets <= 0)
                    IncludeUiaTargets = false;
                MaxActionTextChars = Math.Max(256, MaxActionTextChars);
                IncompleteMaxOutputTokenCap = Math.Max(MaxOutputTokens, IncompleteMaxOutputTokenCap);
            }
        
            internal static void ApplyCliProfileArgs(string[] args)
            {
                string? selected = null;
                for (var i = 0; i < args.Length; i++)
                {
                    var arg = args[i];
                    if (arg.Equals("--fast", StringComparison.OrdinalIgnoreCase))
                        selected = "fast";
                    else if (arg.Equals("--balanced", StringComparison.OrdinalIgnoreCase))
                        selected = "balanced";
                    else if (arg.Equals("--quality", StringComparison.OrdinalIgnoreCase))
                        selected = "quality";
                    else if (arg.Equals("--profile", StringComparison.OrdinalIgnoreCase) &&
                             i + 1 < args.Length &&
                             !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                        selected = args[++i];
                }
        
                if (!string.IsNullOrWhiteSpace(selected))
                    ApplyProfile(selected);
            }
        
            internal static bool TryReadOption(string[] args, ref int index, string optionName, out string value)
            {
                value = "";
                var arg = args[index];
        
                if (arg.Equals(optionName, StringComparison.OrdinalIgnoreCase))
                {
                    if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
                    {
                        Console.Error.WriteLine($"Missing value for {optionName}.");
                        return true;
                    }
        
                    value = args[++index];
                    return true;
                }
        
                var prefix = optionName + "=";
                if (arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    value = arg[prefix.Length..];
                    return true;
                }
        
                return false;
            }
        
            internal static void ApplyConfigFile(string? explicitPath)
            {
                var path = explicitPath;
                if (string.IsNullOrWhiteSpace(path))
                {
                    var local = Path.Combine(AppContext.BaseDirectory, "rdpilot.json");
                    var cwd = Path.Combine(Environment.CurrentDirectory, "rdpilot.json");
                    path = File.Exists(cwd) ? cwd : File.Exists(local) ? local : null;
                }
        
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    return;
        
                try
                {
                    using var doc = JsonDocument.Parse(File.ReadAllText(path));
                    var root = doc.RootElement;
        
                    if (TryGetString(root, "profile", out var profile)) ApplyProfile(profile);
                    if (TryGetString(root, "model", out var model)) Model = model;
                    if (TryGetString(root, "qaModel", out var qaModel)) QaModel = qaModel;
                    if (TryGetString(root, "verifyModel", out var verifyModel)) VerifyModel = verifyModel;
                    if (TryGetString(root, "reasoningEffort", out var effort)) ApplyReasoningEffort(effort, path);
                    if (TryGetString(root, "qaReasoningEffort", out var qaEffort)) ApplyReasoningEffort(qaEffort, path, v => QaReasoningEffort = v, () => QaReasoningEffortExplicit = true);
                    if (TryGetString(root, "verifyReasoningEffort", out var verifyEffort)) ApplyReasoningEffort(verifyEffort, path, v => VerifyReasoningEffort = v, () => VerifyReasoningEffortExplicit = true);
                    if (TryGetBool(root, "mouseEnabled", out var mouseEnabled)) MouseEnabled = mouseEnabled;
                    if (TryGetInt(root, "postActionDelayMs", out var delay)) UiSettleDelayMs = Math.Max(0, delay);
                    if (TryGetInt(root, "gridStepPx", out var grid)) GridStepPx = Math.Max(0, grid);
                    if (TryGetInt(root, "maxSteps", out var maxSteps)) MaxSteps = Math.Max(1, maxSteps);
                    if (TryGetInt(root, "maxWaitSeconds", out var maxWait)) MaxWaitSeconds = Math.Max(0, maxWait);
                    if (TryGetInt(root, "maxOutputTokens", out var maxOutput)) MaxOutputTokens = Math.Max(1, maxOutput);
                    if (TryGetInt(root, "qaMaxOutputTokens", out var qaMaxOutput)) QaMaxOutputTokens = Math.Max(1, qaMaxOutput);
                    if (TryGetInt(root, "verifyMaxOutputTokens", out var verifyMaxOutput)) VerifyMaxOutputTokens = Math.Max(1, verifyMaxOutput);
                    if (TryGetInt(root, "incompleteMaxOutputRetries", out var incompleteRetries)) IncompleteMaxOutputRetries = Math.Clamp(incompleteRetries, 0, 5);
                    if (TryGetInt(root, "incompleteMaxOutputTokenCap", out var incompleteTokenCap)) IncompleteMaxOutputTokenCap = Math.Max(1, incompleteTokenCap);
                    if (TryGetInt(root, "maxActionTextChars", out var maxActionTextChars)) MaxActionTextChars = Math.Max(256, maxActionTextChars);
                    if (TryGetInt(root, "qaScreenshotMaxWidth", out var qaScreenshotMaxWidth)) QaScreenshotMaxWidth = Math.Clamp(qaScreenshotMaxWidth, 0, 10000);
                    if (TryGetInt(root, "verifyScreenshotMaxWidth", out var verifyScreenshotMaxWidth)) VerifyScreenshotMaxWidth = Math.Clamp(verifyScreenshotMaxWidth, 0, 10000);
                    if (TryGetString(root, "textVerbosity", out var verbosity)) ApplyTextVerbosity(verbosity, path);
                    if (TryGetInt(root, "historyTailChars", out var historyChars)) HistoryTailChars = Math.Max(0, historyChars);
                    if (TryGetInt(root, "historyTailLines", out var historyLines)) HistoryTailLines = Math.Max(0, historyLines);
                    if (TryGetInt(root, "maxStagnationSteps", out var maxStagnation)) MaxStagnationStepsBeforeAbort = Math.Max(0, maxStagnation);
                    if (TryGetInt(root, "maxRepeatedActions", out var maxRepeated)) MaxRepeatedActionBeforeAbort = Math.Max(0, maxRepeated);
                    if (TryGetInt(root, "actionRepeatCooldownSteps", out var repeatCooldown)) ActionRepeatCooldownSteps = Math.Max(0, repeatCooldown);
                    if (TryGetDouble(root, "skipVerifyConfidenceThreshold", out var skipVerifyConfidence)) SkipVerifyConfidenceThreshold = Math.Clamp(skipVerifyConfidence, 0.0, 1.0);
                    if (TryGetInt(root, "maxModelFailures", out var maxModelFailures)) MaxModelFailuresBeforeAbort = Math.Max(0, maxModelFailures);
                    if (TryGetInt(root, "maxActionFailures", out var maxActionFailures)) MaxActionFailuresBeforeAbort = Math.Max(0, maxActionFailures);
                    if (TryGetBool(root, "screenPolling", out var screenPolling)) ScreenPollingEnabled = screenPolling;
                    if (TryGetInt(root, "screenPollInitialDelayMs", out var pollInitialDelay)) ScreenPollInitialDelayMs = Math.Clamp(pollInitialDelay, 0, 10000);
                    if (TryGetInt(root, "screenPollIntervalMs", out var pollInterval)) ScreenPollIntervalMs = Math.Clamp(pollInterval, 10, 10000);
                    if (TryGetInt(root, "screenPollTimeoutMs", out var pollTimeout)) ScreenPollTimeoutMs = Math.Clamp(pollTimeout, 0, 60000);
                    if (TryGetInt(root, "waitNoChangeExtraMs", out var waitNoChangeExtra)) WaitNoChangeExtraMs = Math.Clamp(waitNoChangeExtra, 0, 60000);
                    if (TryGetBool(root, "screenSanityChecks", out var screenSanity)) ScreenSanityChecks = screenSanity;
                    if (TryGetInt(root, "sendInputMaxRetries", out var sendInputRetries)) SendInputMaxRetries = Math.Clamp(sendInputRetries, 0, 10);
                    if (TryGetInt(root, "sendInputRetryDelayMs", out var sendInputRetryDelay)) SendInputRetryDelayMs = Math.Clamp(sendInputRetryDelay, 0, 1000);
                    if (TryGetBool(root, "adaptiveReasoningEffort", out var adaptiveEffort)) AdaptiveReasoningEffort = adaptiveEffort;
                    if (TryGetInt(root, "screenshotMaxWidth", out var maxWidth)) MaxScreenshotSendWidth = Math.Max(0, maxWidth);
                    if (TryGetString(root, "screenshotFormat", out var format)) ApplyImageFormat(format, path);
                    if (TryGetInt(root, "focusedOverviewMaxWidth", out var focusedOverviewMaxWidth)) FocusedOverviewMaxWidth = Math.Clamp(focusedOverviewMaxWidth, 0, 10000);
                    if (TryGetInt(root, "cropMaxWidth", out var cropMaxWidth)) MaxCropSendWidth = Math.Max(0, cropMaxWidth);
                    if (TryGetString(root, "cropFormat", out var cropFormat)) ApplyCropFormat(cropFormat, path);
                    if (TryGetInt(root, "screenshotJpegQuality", out var quality)) ScreenshotJpegQuality = Math.Clamp(quality, 1, 100);
                    if (TryGetString(root, "screenLogFormat", out var screenLogFormat)) ApplyScreenLogFormat(screenLogFormat, path);
                    if (TryGetInt(root, "screenLogMaxWidth", out var screenLogMaxWidth)) MaxScreenLogWidth = Math.Clamp(screenLogMaxWidth, 0, 10000);
                    if (TryGetBool(root, "includeFocusUia", out var focusUia)) IncludeFocusUia = focusUia;
                    if (TryGetBool(root, "includeFocusUiaCrop", out var focusCrop)) IncludeFocusUiaCrop = focusCrop;
                    if (TryGetBool(root, "debugImages", out var debugImages)) DebugImages = debugImages;
                    if (TryGetString(root, "verifyMode", out var verifyMode)) ApplyVerifyMode(verifyMode, path);
                    if (TryGetInt(root, "verifyEarlySteps", out var verifyEarlySteps)) VerifyEarlySteps = Math.Max(0, verifyEarlySteps);
                    if (TryGetDouble(root, "verifyLowConfidenceThreshold", out var verifyLowConfidence)) VerifyLowConfidenceThreshold = Math.Clamp(verifyLowConfidence, 0.0, 1.0);
                    if (TryGetBool(root, "refreshScreenshotBeforeVerify", out var refreshBeforeVerify)) RefreshScreenshotBeforeVerify = refreshBeforeVerify;
                    if (TryGetInt(root, "clipboardPasteThreshold", out var pasteThreshold)) ClipboardPasteThreshold = Math.Max(0, pasteThreshold);
                    if (TryGetInt(root, "focusCropSize", out var focusCropSize)) FocusCropSize = Math.Clamp(focusCropSize, 64, 2000);
                    if (TryGetInt(root, "openAiMaxRetries", out var retries)) OpenAiMaxRetries = Math.Clamp(retries, 0, 10);
                    if (TryGetInt(root, "openAiTimeoutSeconds", out var timeoutSeconds)) OpenAiTimeoutSeconds = Math.Max(0, timeoutSeconds);
                    if (TryGetBool(root, "realUiOnly", out var realUiOnly)) ForceRealUiOnly = realUiOnly;
                    if (TryGetBool(root, "allowHighLevelActions", out var allowHighLevel)) AllowHighLevelActions = allowHighLevel;
                    if (TryGetBool(root, "allowRunCommand", out var allowRun)) AllowRunCommand = allowRun;
                    if (TryGetBool(root, "directClickWithoutAim", out var directClick)) DirectClickWithoutAim = directClick;
                    if (TryGetBool(root, "autoHideConsoleDuringRun", out var autoHideConsole)) AutoHideConsoleDuringRun = autoHideConsole;
                    if (TryGetBool(root, "minimizeConsoleDuringRun", out var minimize)) MinimizeConsoleDuringRun = minimize;
                    if (TryGetBool(root, "restoreConsoleAfterRun", out var restoreConsole)) RestoreConsoleAfterRun = restoreConsole;
                    if (TryGetBool(root, "promptCache", out var promptCache)) UsePromptCache = promptCache;
                    if (TryGetString(root, "promptCacheKey", out var promptCacheKey)) PromptCacheKey = promptCacheKey;
                    if (TryGetBool(root, "usePreviousResponseId", out var previousResponseState)) UsePreviousResponseState = previousResponseState;
                    if (TryGetBool(root, "omitUnchangedScreenImage", out var omitUnchangedScreen)) OmitUnchangedScreenImageWithState = omitUnchangedScreen;
                    if (TryGetBool(root, "includeUiaTargets", out var includeUia)) IncludeUiaTargets = includeUia;
                    if (TryGetInt(root, "maxUiaTargets", out var maxUia)) MaxUiaTargets = Math.Clamp(maxUia, 0, 100);
                    if (TryGetInt(root, "uiaTargetNameMaxChars", out var uiaNameChars)) UiaTargetNameMaxChars = Math.Clamp(uiaNameChars, 0, 500);
                    if (TryGetInt(root, "uiaSummaryMaxChars", out var uiaSummaryChars)) UiaSummaryMaxChars = Math.Clamp(uiaSummaryChars, 0, 2000);
                    if (TryGetInt(root, "uiaScanTimeBudgetMs", out var uiaScanMs)) UiaScanTimeBudgetMs = Math.Clamp(uiaScanMs, 0, 5000);
                    if (TryGetInt(root, "maxUiaNodesScanned", out var maxUiaNodes)) MaxUiaNodesScanned = Math.Clamp(maxUiaNodes, 0, 10000);
                    if (TryGetInt(root, "uiaCandidateMultiplier", out var candidateMultiplier)) UiaCandidateMultiplier = Math.Clamp(candidateMultiplier, 1, 20);
                    if (TryGetDouble(root, "uiaMaxAreaRatio", out var maxAreaRatio)) MaxUiaTargetAreaRatio = Math.Clamp(maxAreaRatio, 0.0, 1.0);
                    if (TryGetBool(root, "reuseUiaTargetsWhenScreenUnchanged", out var reuseUia)) ReuseUiaTargetsWhenScreenUnchanged = reuseUia;
                    if (TryGetBool(root, "executeMultiActionCandidates", out var executeCandidates)) ExecuteMultiActionCandidates = executeCandidates;
                    if (TryGetInt(root, "maxQueuedBatchActions", out var maxQueued)) MaxQueuedBatchActions = Math.Clamp(maxQueued, 0, 20);
                    if (TryGetBool(root, "logRequests", out var logRequests)) LogRequests = logRequests;
                    if (TryGetBool(root, "prettyRequestLogs", out var prettyRequestLogs)) PrettyRequestLogs = prettyRequestLogs;
                    if (TryGetBool(root, "logScreens", out var logScreens)) LogScreens = logScreens;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Could not read config file '{path}': {ex.Message}");
                }
        
                static bool TryGetString(JsonElement root, string name, out string value)
                {
                    value = "";
                    if (root.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.String)
                    {
                        value = el.GetString() ?? "";
                        return !string.IsNullOrWhiteSpace(value);
                    }
                    return false;
                }
        
                static bool TryGetInt(JsonElement root, string name, out int value)
                {
                    value = 0;
                    return root.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.Number && el.TryGetInt32(out value);
                }
        
                static bool TryGetBool(JsonElement root, string name, out bool value)
                {
                    value = false;
                    if (root.TryGetProperty(name, out var el) &&
                        (el.ValueKind == JsonValueKind.True || el.ValueKind == JsonValueKind.False))
                    {
                        value = el.GetBoolean();
                        return true;
                    }
        
                    return false;
                }
        
                static bool TryGetDouble(JsonElement root, string name, out double value)
                {
                    value = 0;
                    return root.TryGetProperty(name, out var el) && el.ValueKind == JsonValueKind.Number && el.TryGetDouble(out value);
                }
            }
        
            internal static void ApplyProfile(string? profile)
            {
                if (string.IsNullOrWhiteSpace(profile))
                    return;
        
                var normalized = profile.Trim().ToLowerInvariant();
                switch (normalized)
                {
                    case "fast":
                        RunProfile = "fast";
                        ReasoningEffort = "low";
                        if (!QaReasoningEffortExplicit) QaReasoningEffort = "low";
                        if (!VerifyReasoningEffortExplicit) VerifyReasoningEffort = "low";
                        UiSettleDelayMs = Math.Min(UiSettleDelayMs, 300);
                        MaxOutputTokens = 300;
                        QaMaxOutputTokens = Math.Min(QaMaxOutputTokens, 300);
                        VerifyMaxOutputTokens = Math.Min(VerifyMaxOutputTokens, 120);
                        QaScreenshotMaxWidth = Math.Min(QaScreenshotMaxWidth, 1024);
                        VerifyScreenshotMaxWidth = Math.Min(VerifyScreenshotMaxWidth, 1024);
                        TextVerbosity = "low";
                        HistoryTailChars = Math.Min(HistoryTailChars, 1200);
                        HistoryTailLines = Math.Min(HistoryTailLines, 12);
                        MaxStagnationStepsBeforeAbort = Math.Min(MaxStagnationStepsBeforeAbort, 8);
                        MaxRepeatedActionBeforeAbort = Math.Min(MaxRepeatedActionBeforeAbort, 5);
                        ActionRepeatCooldownSteps = Math.Min(ActionRepeatCooldownSteps, 2);
                        SkipVerifyConfidenceThreshold = Math.Min(SkipVerifyConfidenceThreshold, 0.92);
                        MaxModelFailuresBeforeAbort = Math.Min(MaxModelFailuresBeforeAbort, 2);
                        MaxActionFailuresBeforeAbort = Math.Min(MaxActionFailuresBeforeAbort, 2);
                        ScreenPollingEnabled = true;
                        ScreenPollInitialDelayMs = Math.Min(ScreenPollInitialDelayMs, 120);
                        ScreenPollIntervalMs = Math.Min(ScreenPollIntervalMs, 150);
                        ScreenPollTimeoutMs = Math.Min(ScreenPollTimeoutMs, 1200);
                        WaitNoChangeExtraMs = Math.Min(WaitNoChangeExtraMs, 750);
                        MaxWaitSeconds = Math.Min(MaxWaitSeconds, 30);
                        MaxScreenshotSendWidth = 1280;
                        ScreenshotSendFormat = "jpeg";
                        FocusedOverviewMaxWidth = 640;
                        MaxCropSendWidth = 768;
                        CropSendFormat = "jpeg";
                        ScreenshotJpegQuality = 80;
                        ScreenLogFormat = "jpeg";
                        MaxScreenLogWidth = 1280;
                        PrettyRequestLogs = false;
                        IncludeFocusUiaCrop = false;
                        DebugImages = false;
                        VerifyMode = "auto";
                        UiaTargetNameMaxChars = Math.Min(UiaTargetNameMaxChars, 48);
                        UiaSummaryMaxChars = Math.Min(UiaSummaryMaxChars, 320);
                        UiaScanTimeBudgetMs = Math.Min(UiaScanTimeBudgetMs, 60);
                        MaxUiaNodesScanned = Math.Min(MaxUiaNodesScanned, 400);
                        UiaCandidateMultiplier = Math.Min(UiaCandidateMultiplier, 4);
                        MaxUiaTargetAreaRatio = Math.Min(MaxUiaTargetAreaRatio, 0.45);
                        break;
        
                    case "balanced":
                        RunProfile = "balanced";
                        ReasoningEffort = "low";
                        if (!QaReasoningEffortExplicit) QaReasoningEffort = "low";
                        if (!VerifyReasoningEffortExplicit) VerifyReasoningEffort = "low";
                        UiSettleDelayMs = Math.Max(UiSettleDelayMs, 500);
                        MaxOutputTokens = 450;
                        QaMaxOutputTokens = Math.Max(QaMaxOutputTokens, 450);
                        VerifyMaxOutputTokens = Math.Max(VerifyMaxOutputTokens, 160);
                        QaScreenshotMaxWidth = Math.Max(QaScreenshotMaxWidth, 1280);
                        VerifyScreenshotMaxWidth = Math.Max(VerifyScreenshotMaxWidth, 1280);
                        TextVerbosity = "low";
                        HistoryTailChars = Math.Max(HistoryTailChars, 2000);
                        HistoryTailLines = Math.Max(HistoryTailLines, 20);
                        MaxStagnationStepsBeforeAbort = Math.Max(MaxStagnationStepsBeforeAbort, 12);
                        MaxRepeatedActionBeforeAbort = Math.Max(MaxRepeatedActionBeforeAbort, 7);
                        ActionRepeatCooldownSteps = Math.Max(ActionRepeatCooldownSteps, 3);
                        SkipVerifyConfidenceThreshold = Math.Min(SkipVerifyConfidenceThreshold, 0.95);
                        MaxModelFailuresBeforeAbort = Math.Max(MaxModelFailuresBeforeAbort, 3);
                        MaxActionFailuresBeforeAbort = Math.Max(MaxActionFailuresBeforeAbort, 3);
                        ScreenPollingEnabled = true;
                        ScreenPollInitialDelayMs = Math.Max(ScreenPollInitialDelayMs, 150);
                        ScreenPollIntervalMs = Math.Max(ScreenPollIntervalMs, 180);
                        ScreenPollTimeoutMs = Math.Max(ScreenPollTimeoutMs, 1800);
                        WaitNoChangeExtraMs = Math.Max(WaitNoChangeExtraMs, 1000);
                        MaxWaitSeconds = Math.Max(MaxWaitSeconds, 60);
                        MaxScreenshotSendWidth = 1600;
                        ScreenshotSendFormat = "jpeg";
                        FocusedOverviewMaxWidth = Math.Max(FocusedOverviewMaxWidth, 960);
                        MaxCropSendWidth = 1024;
                        CropSendFormat = "jpeg";
                        ScreenshotJpegQuality = 88;
                        ScreenLogFormat = "jpeg";
                        MaxScreenLogWidth = 1600;
                        PrettyRequestLogs = false;
                        IncludeFocusUiaCrop = true;
                        DebugImages = false;
                        VerifyMode = "auto";
                        UiaTargetNameMaxChars = Math.Max(UiaTargetNameMaxChars, 64);
                        UiaSummaryMaxChars = Math.Max(UiaSummaryMaxChars, 480);
                        UiaScanTimeBudgetMs = Math.Max(UiaScanTimeBudgetMs, 90);
                        MaxUiaNodesScanned = Math.Max(MaxUiaNodesScanned, 700);
                        UiaCandidateMultiplier = Math.Max(UiaCandidateMultiplier, 5);
                        MaxUiaTargetAreaRatio = Math.Max(MaxUiaTargetAreaRatio, 0.55);
                        break;
        
                    case "quality":
                        RunProfile = "quality";
                        ReasoningEffort = "medium";
                        if (!QaReasoningEffortExplicit) QaReasoningEffort = "low";
                        if (!VerifyReasoningEffortExplicit) VerifyReasoningEffort = "low";
                        UiSettleDelayMs = Math.Max(UiSettleDelayMs, 1000);
                        MaxOutputTokens = 800;
                        QaMaxOutputTokens = Math.Max(QaMaxOutputTokens, 800);
                        VerifyMaxOutputTokens = Math.Max(VerifyMaxOutputTokens, 240);
                        QaScreenshotMaxWidth = 0;
                        VerifyScreenshotMaxWidth = 0;
                        TextVerbosity = "medium";
                        HistoryTailChars = Math.Max(HistoryTailChars, 4000);
                        HistoryTailLines = Math.Max(HistoryTailLines, 40);
                        MaxStagnationStepsBeforeAbort = Math.Max(MaxStagnationStepsBeforeAbort, 20);
                        MaxRepeatedActionBeforeAbort = Math.Max(MaxRepeatedActionBeforeAbort, 10);
                        ActionRepeatCooldownSteps = Math.Max(ActionRepeatCooldownSteps, 4);
                        SkipVerifyConfidenceThreshold = Math.Max(SkipVerifyConfidenceThreshold, 0.98);
                        MaxModelFailuresBeforeAbort = Math.Max(MaxModelFailuresBeforeAbort, 4);
                        MaxActionFailuresBeforeAbort = Math.Max(MaxActionFailuresBeforeAbort, 4);
                        ScreenPollingEnabled = true;
                        ScreenPollInitialDelayMs = Math.Max(ScreenPollInitialDelayMs, 250);
                        ScreenPollIntervalMs = Math.Max(ScreenPollIntervalMs, 250);
                        ScreenPollTimeoutMs = Math.Max(ScreenPollTimeoutMs, 3000);
                        WaitNoChangeExtraMs = Math.Max(WaitNoChangeExtraMs, 1500);
                        MaxWaitSeconds = Math.Max(MaxWaitSeconds, 120);
                        MaxScreenshotSendWidth = 0;
                        ScreenshotSendFormat = "png";
                        FocusedOverviewMaxWidth = 0;
                        MaxCropSendWidth = 0;
                        CropSendFormat = "png";
                        ScreenshotJpegQuality = 95;
                        ScreenLogFormat = "png";
                        MaxScreenLogWidth = 0;
                        PrettyRequestLogs = true;
                        IncludeFocusUiaCrop = true;
                        DebugImages = true;
                        VerifyMode = "always";
                        UiaTargetNameMaxChars = Math.Max(UiaTargetNameMaxChars, 96);
                        UiaSummaryMaxChars = Math.Max(UiaSummaryMaxChars, 800);
                        UiaScanTimeBudgetMs = Math.Max(UiaScanTimeBudgetMs, 180);
                        MaxUiaNodesScanned = Math.Max(MaxUiaNodesScanned, 1500);
                        UiaCandidateMultiplier = Math.Max(UiaCandidateMultiplier, 8);
                        MaxUiaTargetAreaRatio = Math.Max(MaxUiaTargetAreaRatio, 0.75);
                        break;
        
                    default:
                        Console.Error.WriteLine($"Unknown profile '{profile}'. Allowed: fast, balanced, quality.");
                        break;
                }
            }
        
            internal static void ApplyReasoningEffort(string? value, string source)
            {
                ApplyReasoningEffort(value, source, v => ReasoningEffort = v);
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
                    Console.Error.WriteLine($"Invalid reasoning effort '{value}' from {source}. Allowed: default, {string.Join(", ", AllowedReasoningEfforts.OrderBy(x => x))}.");
                    return;
                }
        
                setter(normalized);
            }
        
            internal static void ApplyDelayMs(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (int.TryParse(value, out var ms) && ms >= 0)
                    UiSettleDelayMs = ms;
                else
                    Console.Error.WriteLine($"Invalid post-action delay '{value}' from {source}; expected a non-negative integer.");
            }
        
            internal static void ApplyGridStep(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (value.Equals("off", StringComparison.OrdinalIgnoreCase))
                {
                    GridStepPx = 0;
                    return;
                }
        
                if (int.TryParse(value, out var px) && px >= 0)
                    GridStepPx = px;
                else
                    Console.Error.WriteLine($"Invalid grid step '{value}' from {source}; expected a non-negative integer or 'off'.");
            }
        
            internal static void ApplyInt(string? value, string source, int min, int max, Action<int> setter)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (int.TryParse(value, out var parsed) && parsed >= min && parsed <= max)
                    setter(parsed);
                else
                    Console.Error.WriteLine($"Invalid integer '{value}' from {source}; expected {min}..{max}.");
            }
        
            internal static void ApplyLong(string? value, string source, long min, long max, Action<long> setter)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                if (long.TryParse(value, out var parsed) && parsed >= min && parsed <= max)
                    setter(parsed);
                else
                    Console.Error.WriteLine($"Invalid integer '{value}' from {source}; expected {min}..{max}.");
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
                    Console.Error.WriteLine($"Invalid number '{value}' from {source}; expected {min:0.###}..{max:0.###}.");
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
                    Console.Error.WriteLine($"Invalid boolean '{value}' from {source}; expected 1/true/yes/on or 0/false/no/off.");
            }
        
            internal static void ApplyTextVerbosity(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "low" or "medium" or "high")
                    TextVerbosity = normalized;
                else
                    Console.Error.WriteLine($"Invalid text verbosity '{value}' from {source}; expected low, medium, or high.");
            }
        
            internal static void ApplyImageFormat(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "jpeg" or "jpg")
                    ScreenshotSendFormat = "jpeg";
                else if (normalized == "png")
                    ScreenshotSendFormat = "png";
                else
                    Console.Error.WriteLine($"Invalid screenshot format '{value}' from {source}; expected jpeg or png.");
            }
        
            internal static void ApplyCropFormat(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "jpeg" or "jpg")
                    CropSendFormat = "jpeg";
                else if (normalized == "png")
                    CropSendFormat = "png";
                else
                    Console.Error.WriteLine($"Invalid crop format '{value}' from {source}; expected jpeg or png.");
            }
        
            internal static void ApplyScreenLogFormat(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "jpeg" or "jpg")
                    ScreenLogFormat = "jpeg";
                else if (normalized == "png")
                    ScreenLogFormat = "png";
                else
                    Console.Error.WriteLine($"Invalid screen log format '{value}' from {source}; expected jpeg or png.");
            }
        
            internal static void ApplyVerifyMode(string? value, string source)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return;
        
                var normalized = value.Trim().ToLowerInvariant();
                if (normalized is "auto" or "always" or "off")
                    VerifyMode = normalized;
                else
                    Console.Error.WriteLine($"Invalid verify mode '{value}' from {source}; expected auto, always, or off.");
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
                MaxScreenshotSendWidth > 0 ? MaxScreenshotSendWidth.ToString() : "original";
        
            internal static string FocusedOverviewMaxWidthDisplay() =>
                FocusedOverviewMaxWidth > 0 ? FocusedOverviewMaxWidth.ToString() : "normal";
        
            internal static int EffectiveFocusedOverviewMaxWidth() =>
                FocusedOverviewMaxWidth > 0 ? FocusedOverviewMaxWidth : MaxScreenshotSendWidth;
        
            internal static string QaScreenshotMaxWidthDisplay() =>
                QaScreenshotMaxWidth > 0 ? QaScreenshotMaxWidth.ToString() : "normal";
        
            internal static string VerifyScreenshotMaxWidthDisplay() =>
                VerifyScreenshotMaxWidth > 0 ? VerifyScreenshotMaxWidth.ToString() : "normal";
        
            internal static string CropMaxWidthDisplay() =>
                MaxCropSendWidth > 0 ? MaxCropSendWidth.ToString() : "original";
        
            internal static string ScreenLogMaxWidthDisplay() =>
                MaxScreenLogWidth > 0 ? MaxScreenLogWidth.ToString() : "original";
        
            internal static string EffectiveQaModel() => string.IsNullOrWhiteSpace(QaModel) ? Model : QaModel!;
            internal static string EffectiveVerifyModel() => string.IsNullOrWhiteSpace(VerifyModel) ? Model : VerifyModel!;
            internal static string? EffectiveQaReasoningEffort() => QaReasoningEffortExplicit ? QaReasoningEffort : (QaReasoningEffort ?? ReasoningEffort);
            internal static string? EffectiveVerifyReasoningEffort() => VerifyReasoningEffortExplicit ? VerifyReasoningEffort : (VerifyReasoningEffort ?? ReasoningEffort);
        
            internal static void ConfigureOpenAiHttpClient()
            {
                OpenAiHttp.Timeout = OpenAiTimeoutSeconds > 0
                    ? TimeSpan.FromSeconds(OpenAiTimeoutSeconds)
                    : Timeout.InfiniteTimeSpan;
            }
    }
}



