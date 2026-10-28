namespace RDPilot.Configuration;

/// <summary>Loads file and environment settings in application precedence order.</summary>
internal static class ConfigurationLoader
{
    internal static void ApplyEnvironmentConfig()
    {
        ApplyProfile(Options.RunProfile);
        ApplyAutomaticConfigFile();
        ApplyConfigFile(Environment.GetEnvironmentVariable("RDPILOT_CONFIG"));
        ApplyEnvironmentProfileOverrides();

        var envModel = Environment.GetEnvironmentVariable("OPENAI_MODEL");
        if (!string.IsNullOrWhiteSpace(envModel))
            Options.Model = envModel.Trim();
        var envQaModel = Environment.GetEnvironmentVariable("OPENAI_QA_MODEL");
        if (!string.IsNullOrWhiteSpace(envQaModel))
            Options.QaModel = envQaModel.Trim();
        var envVerifyModel = Environment.GetEnvironmentVariable("OPENAI_VERIFY_MODEL");
        if (!string.IsNullOrWhiteSpace(envVerifyModel))
            Options.VerifyModel = envVerifyModel.Trim();

        var envEffort = Environment.GetEnvironmentVariable("OPENAI_REASONING_EFFORT")
                     ?? Environment.GetEnvironmentVariable("REASONING_EFFORT");
        ApplyReasoningEffort(envEffort, "OPENAI_REASONING_EFFORT/REASONING_EFFORT");
        ApplyReasoningEffort(Environment.GetEnvironmentVariable("OPENAI_QA_REASONING_EFFORT"), "OPENAI_QA_REASONING_EFFORT", v => Options.QaReasoningEffort = v, () => Options.QaReasoningEffortExplicit = true);
        ApplyReasoningEffort(Environment.GetEnvironmentVariable("OPENAI_VERIFY_REASONING_EFFORT"), "OPENAI_VERIFY_REASONING_EFFORT", v => Options.VerifyReasoningEffort = v, () => Options.VerifyReasoningEffortExplicit = true);

        var envME = Environment.GetEnvironmentVariable("MOUSE_ENABLED");
        if (!string.IsNullOrWhiteSpace(envME))
            Options.MouseEnabled = IsTruthy(envME);

        ApplyDelayMs(Environment.GetEnvironmentVariable("POST_ACTION_DELAY_MS"), "POST_ACTION_DELAY_MS");
        ApplyGridStep(Environment.GetEnvironmentVariable("GRID_STEP_PX"), "GRID_STEP_PX");
        ApplyInt(Environment.GetEnvironmentVariable("MAX_STEPS"), "MAX_STEPS", 0, int.MaxValue, v => Options.MaxSteps = v);
        ApplyBool(Environment.GetEnvironmentVariable("MULTI_MONITOR"), "MULTI_MONITOR", v => Options.MultiMonitorEnabled = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_WAIT_SECONDS"), "MAX_WAIT_SECONDS", 0, int.MaxValue, v => Options.MaxWaitSeconds = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_OUTPUT_TOKENS"), "MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => Options.MaxOutputTokens = v);
        ApplyInt(Environment.GetEnvironmentVariable("QA_MAX_OUTPUT_TOKENS"), "QA_MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => Options.QaMaxOutputTokens = v);
        ApplyInt(Environment.GetEnvironmentVariable("VERIFY_MAX_OUTPUT_TOKENS"), "VERIFY_MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => Options.VerifyMaxOutputTokens = v);
        ApplyInt(Environment.GetEnvironmentVariable("TURN_REANALYSIS_MAX_OUTPUT_TOKENS"), "TURN_REANALYSIS_MAX_OUTPUT_TOKENS", 1, int.MaxValue, v => Options.TurnReanalysisMaxOutputTokens = v);
        ApplyInt(Environment.GetEnvironmentVariable("INCOMPLETE_MAX_OUTPUT_RETRIES"), "INCOMPLETE_MAX_OUTPUT_RETRIES", 0, 5, v => Options.IncompleteMaxOutputRetries = v);
        ApplyInt(Environment.GetEnvironmentVariable("INCOMPLETE_MAX_OUTPUT_TOKEN_CAP"), "INCOMPLETE_MAX_OUTPUT_TOKEN_CAP", 1, int.MaxValue, v => Options.IncompleteMaxOutputTokenCap = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_ACTION_TEXT_CHARS"), "MAX_ACTION_TEXT_CHARS", 256, int.MaxValue, v => Options.MaxActionTextChars = v);
        ApplyInt(Environment.GetEnvironmentVariable("QA_SCREENSHOT_MAX_WIDTH"), "QA_SCREENSHOT_MAX_WIDTH", 0, 10000, v => Options.QaScreenshotMaxWidth = v);
        ApplyInt(Environment.GetEnvironmentVariable("VERIFY_SCREENSHOT_MAX_WIDTH"), "VERIFY_SCREENSHOT_MAX_WIDTH", 0, 10000, v => Options.VerifyScreenshotMaxWidth = v);
        ApplyTextVerbosity(Environment.GetEnvironmentVariable("TEXT_VERBOSITY"), "TEXT_VERBOSITY");
        ApplyInt(Environment.GetEnvironmentVariable("HISTORY_TAIL_CHARS"), "HISTORY_TAIL_CHARS", 0, int.MaxValue, v => Options.HistoryTailChars = v);
        ApplyInt(Environment.GetEnvironmentVariable("HISTORY_TAIL_LINES"), "HISTORY_TAIL_LINES", 0, int.MaxValue, v => Options.HistoryTailLines = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_STAGNATION_STEPS"), "MAX_STAGNATION_STEPS", 0, int.MaxValue, v => Options.MaxStagnationStepsBeforeAbort = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_REPEATED_ACTIONS"), "MAX_REPEATED_ACTIONS", 0, int.MaxValue, v => Options.MaxRepeatedActionBeforeAbort = v);
        ApplyInt(Environment.GetEnvironmentVariable("ACTION_REPEAT_COOLDOWN_STEPS"), "ACTION_REPEAT_COOLDOWN_STEPS", 0, int.MaxValue, v => Options.ActionRepeatCooldownSteps = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_REJECTED_PROPOSAL_REPEATS"), "MAX_REJECTED_PROPOSAL_REPEATS", 0, int.MaxValue, v => Options.MaxRejectedProposalRepeatsBeforeAbort = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_CONSECUTIVE_INSPECTION_ACTIONS"), "MAX_CONSECUTIVE_INSPECTION_ACTIONS", 0, 20, v => Options.MaxConsecutiveInspectionActions = v);
        ApplyGoalMode(Environment.GetEnvironmentVariable("GOAL_MODE"), "GOAL_MODE");
        ApplyObservationMode(Environment.GetEnvironmentVariable("OBSERVATION_PROFILE"), "OBSERVATION_PROFILE");
        ApplyBool(Environment.GetEnvironmentVariable("OBSERVATION_LOG_VERBOSE"), "OBSERVATION_LOG_VERBOSE", v => Options.ObservationLogVerbose = v);
        ApplyDouble(Environment.GetEnvironmentVariable("SKIP_VERIFY_CONFIDENCE_THRESHOLD"), "SKIP_VERIFY_CONFIDENCE_THRESHOLD", 0.0, 1.0, v => Options.SkipVerifyConfidenceThreshold = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_MODEL_FAILURES"), "MAX_MODEL_FAILURES", 0, int.MaxValue, v => Options.MaxModelFailuresBeforeAbort = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_ACTION_FAILURES"), "MAX_ACTION_FAILURES", 0, int.MaxValue, v => Options.MaxActionFailuresBeforeAbort = v);
        ApplyBool(Environment.GetEnvironmentVariable("RECOVERY_MEMORY"), "RECOVERY_MEMORY", v => Options.RecoveryMemoryEnabled = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_TRIGGER_STEPS"), "RECOVERY_MEMORY_TRIGGER_STEPS", 1, int.MaxValue, v => Options.RecoveryMemoryTriggerSteps = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_VALIDATION_STEPS"), "RECOVERY_MEMORY_VALIDATION_STEPS", 1, int.MaxValue, v => Options.RecoveryMemoryValidationSteps = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_MAX_LESSONS"), "RECOVERY_MEMORY_MAX_LESSONS", 1, int.MaxValue, v => Options.RecoveryMemoryMaxLessons = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_MAX_QUARANTINED_LESSONS"), "RECOVERY_MEMORY_MAX_QUARANTINED_LESSONS", 1, int.MaxValue, v => Options.RecoveryMemoryMaxQuarantinedLessons = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_RESERVED_LESSONS_PER_CONTEXT"), "RECOVERY_MEMORY_RESERVED_LESSONS_PER_CONTEXT", 0, int.MaxValue, v => Options.RecoveryMemoryReservedLessonsPerContext = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_SOFT_MAX_LESSONS_PER_CONTEXT"), "RECOVERY_MEMORY_SOFT_MAX_LESSONS_PER_CONTEXT", 1, int.MaxValue, v => Options.RecoveryMemorySoftMaxLessonsPerContext = v);
        ApplyLong(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_MAX_FILE_BYTES"), "RECOVERY_MEMORY_MAX_FILE_BYTES", 1024 * 1024, int.MaxValue, v => Options.RecoveryMemoryMaxFileBytes = v);
        ApplyLong(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_ARCHIVE_MAX_BYTES"), "RECOVERY_MEMORY_ARCHIVE_MAX_BYTES", 1024 * 1024, int.MaxValue, v => Options.RecoveryMemoryArchiveMaxBytes = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_ARCHIVE_RETAINED_FILES"), "RECOVERY_MEMORY_ARCHIVE_RETAINED_FILES", 1, 20, v => Options.RecoveryMemoryArchiveRetainedFiles = v);
        var envRecoveryMemoryArchivePath = Environment.GetEnvironmentVariable("RECOVERY_MEMORY_ARCHIVE_PATH");
        if (!string.IsNullOrWhiteSpace(envRecoveryMemoryArchivePath))
            Options.RecoveryMemoryArchivePath = envRecoveryMemoryArchivePath.Trim();
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_PROMPT_LESSONS"), "RECOVERY_MEMORY_PROMPT_LESSONS", 0, 10, v => Options.RecoveryMemoryPromptMaxLessons = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_MEMORY_FAILURE_LIMIT"), "RECOVERY_MEMORY_FAILURE_LIMIT", 1, int.MaxValue, v => Options.RecoveryMemoryFailureLimit = v);
        ApplyInt(Environment.GetEnvironmentVariable("RUNTIME_SEMANTIC_STATE_LIMIT"), "RUNTIME_SEMANTIC_STATE_LIMIT", 32, 100000, v => Options.RuntimeSemanticStateLimit = v);
        ApplyInt(Environment.GetEnvironmentVariable("RUNTIME_GRAPH_EDGE_LIMIT"), "RUNTIME_GRAPH_EDGE_LIMIT", 32, 100000, v => Options.RuntimeGraphEdgeLimit = v);
        ApplyInt(Environment.GetEnvironmentVariable("RUNTIME_RECOVERY_ACTION_LIMIT"), "RUNTIME_RECOVERY_ACTION_LIMIT", 8, 10000, v => Options.RuntimeRecoveryActionLimit = v);
        ApplyInt(Environment.GetEnvironmentVariable("RUNTIME_COOLDOWN_ENTRY_LIMIT"), "RUNTIME_COOLDOWN_ENTRY_LIMIT", 16, 100000, v => Options.RuntimeCooldownEntryLimit = v);
        ApplyInt(Environment.GetEnvironmentVariable("GRAPH_CANDIDATE_TTL_STEPS"), "GRAPH_CANDIDATE_TTL_STEPS", 2, 10000, v => Options.GraphCandidateTtlSteps = v);
        ApplyDouble(Environment.GetEnvironmentVariable("PROACTIVE_LOOP_CONFIDENCE_THRESHOLD"), "PROACTIVE_LOOP_CONFIDENCE_THRESHOLD", 0.5, 1.0, v => Options.ProactiveLoopConfidenceThreshold = v);
        ApplyBool(Environment.GetEnvironmentVariable("RECOVERY_PROGRESS_VERIFICATION"), "RECOVERY_PROGRESS_VERIFICATION", v => Options.RecoveryProgressVerificationEnabled = v);
        ApplyDouble(Environment.GetEnvironmentVariable("RECOVERY_PROGRESS_CONFIDENCE_THRESHOLD"), "RECOVERY_PROGRESS_CONFIDENCE_THRESHOLD", 0.5, 1.0, v => Options.RecoveryProgressConfidenceThreshold = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_TELEMETRY_MAX_BYTES"), "RECOVERY_TELEMETRY_MAX_BYTES", 65536, int.MaxValue, v => Options.RecoveryTelemetryMaxBytes = v);
        ApplyInt(Environment.GetEnvironmentVariable("RECOVERY_TELEMETRY_RETAINED_FILES"), "RECOVERY_TELEMETRY_RETAINED_FILES", 1, 20, v => Options.RecoveryTelemetryRetainedFiles = v);
        ApplyBool(Environment.GetEnvironmentVariable("LOOP_REPLAY_AUTO_EXPORT"), "LOOP_REPLAY_AUTO_EXPORT", v => Options.LoopReplayAutoExportEnabled = v);
        var envLoopReplayCorpusPath = Environment.GetEnvironmentVariable("LOOP_REPLAY_CORPUS_PATH");
        if (!string.IsNullOrWhiteSpace(envLoopReplayCorpusPath))
            Options.LoopReplayCorpusPath = envLoopReplayCorpusPath.Trim();
        var envRecoveryMemoryPath = Environment.GetEnvironmentVariable("RECOVERY_MEMORY_PATH");
        if (!string.IsNullOrWhiteSpace(envRecoveryMemoryPath))
            Options.RecoveryMemoryPath = envRecoveryMemoryPath.Trim();
        ApplyBool(Environment.GetEnvironmentVariable("SCREEN_POLLING"), "SCREEN_POLLING", v => Options.ScreenPollingEnabled = v);
        ApplyInt(Environment.GetEnvironmentVariable("SCREEN_POLL_INITIAL_DELAY_MS"), "SCREEN_POLL_INITIAL_DELAY_MS", 0, 10000, v => Options.ScreenPollInitialDelayMs = v);
        ApplyInt(Environment.GetEnvironmentVariable("SCREEN_POLL_INTERVAL_MS"), "SCREEN_POLL_INTERVAL_MS", 10, 10000, v => Options.ScreenPollIntervalMs = v);
        ApplyInt(Environment.GetEnvironmentVariable("SCREEN_POLL_TIMEOUT_MS"), "SCREEN_POLL_TIMEOUT_MS", 0, 60000, v => Options.ScreenPollTimeoutMs = v);
        ApplyInt(Environment.GetEnvironmentVariable("WAIT_NO_CHANGE_EXTRA_MS"), "WAIT_NO_CHANGE_EXTRA_MS", 0, 60000, v => Options.WaitNoChangeExtraMs = v);
        ApplyBool(Environment.GetEnvironmentVariable("SCREEN_SANITY_CHECKS"), "SCREEN_SANITY_CHECKS", v => Options.ScreenSanityChecks = v);
        ApplyInt(Environment.GetEnvironmentVariable("SENDINPUT_MAX_RETRIES"), "SENDINPUT_MAX_RETRIES", 0, 10, v => Options.SendInputMaxRetries = v);
        ApplyInt(Environment.GetEnvironmentVariable("SENDINPUT_RETRY_DELAY_MS"), "SENDINPUT_RETRY_DELAY_MS", 0, 1000, v => Options.SendInputRetryDelayMs = v);
        ApplyBool(Environment.GetEnvironmentVariable("ADAPTIVE_REASONING_EFFORT"), "ADAPTIVE_REASONING_EFFORT", v => Options.AdaptiveReasoningEffort = v);
        ApplyInt(Environment.GetEnvironmentVariable("SCREENSHOT_MAX_WIDTH"), "SCREENSHOT_MAX_WIDTH", 0, 10000, v => Options.MaxScreenshotSendWidth = v);
        ApplyImageFormat(Environment.GetEnvironmentVariable("SCREENSHOT_FORMAT"), "SCREENSHOT_FORMAT");
        ApplyInt(Environment.GetEnvironmentVariable("FOCUSED_OVERVIEW_MAX_WIDTH"), "FOCUSED_OVERVIEW_MAX_WIDTH", 0, 10000, v => Options.FocusedOverviewMaxWidth = v);
        ApplyInt(Environment.GetEnvironmentVariable("CROP_MAX_WIDTH"), "CROP_MAX_WIDTH", 0, 10000, v => Options.MaxCropSendWidth = v);
        ApplyCropFormat(Environment.GetEnvironmentVariable("CROP_FORMAT"), "CROP_FORMAT");
        ApplyLong(Environment.GetEnvironmentVariable("SCREENSHOT_JPEG_QUALITY"), "SCREENSHOT_JPEG_QUALITY", 1, 100, v => Options.ScreenshotJpegQuality = v);
        ApplyScreenLogFormat(Environment.GetEnvironmentVariable("SCREEN_LOG_FORMAT"), "SCREEN_LOG_FORMAT");
        ApplyInt(Environment.GetEnvironmentVariable("SCREEN_LOG_MAX_WIDTH"), "SCREEN_LOG_MAX_WIDTH", 0, 10000, v => Options.MaxScreenLogWidth = v);
        ApplyBool(Environment.GetEnvironmentVariable("DEBUG_IMAGES"), "DEBUG_IMAGES", v => Options.DebugImages = v);
        ApplyBool(Environment.GetEnvironmentVariable("LOG_REQUESTS"), "LOG_REQUESTS", v => Options.LogRequests = v);
        ApplyBool(Environment.GetEnvironmentVariable("PRETTY_REQUEST_LOGS"), "PRETTY_REQUEST_LOGS", v => Options.PrettyRequestLogs = v);
        ApplyBool(Environment.GetEnvironmentVariable("LOG_SCREENS"), "LOG_SCREENS", v => Options.LogScreens = v);
        ApplyBool(Environment.GetEnvironmentVariable("INCLUDE_FOCUS_UIA"), "INCLUDE_FOCUS_UIA", v => Options.IncludeFocusUia = v);
        ApplyBool(Environment.GetEnvironmentVariable("INCLUDE_FOCUS_UIA_CROP"), "INCLUDE_FOCUS_UIA_CROP", v => Options.IncludeFocusUiaCrop = v);
        ApplyVerifyMode(Environment.GetEnvironmentVariable("VERIFY_MODE"), "VERIFY_MODE");
        ApplyInt(Environment.GetEnvironmentVariable("VERIFY_EARLY_STEPS"), "VERIFY_EARLY_STEPS", 0, int.MaxValue, v => Options.VerifyEarlySteps = v);
        ApplyDouble(Environment.GetEnvironmentVariable("VERIFY_LOW_CONFIDENCE_THRESHOLD"), "VERIFY_LOW_CONFIDENCE_THRESHOLD", 0.0, 1.0, v => Options.VerifyLowConfidenceThreshold = v);
        ApplyBool(Environment.GetEnvironmentVariable("REFRESH_SCREENSHOT_BEFORE_VERIFY"), "REFRESH_SCREENSHOT_BEFORE_VERIFY", v => Options.RefreshScreenshotBeforeVerify = v);
        ApplyInt(Environment.GetEnvironmentVariable("CLIPBOARD_PASTE_THRESHOLD"), "CLIPBOARD_PASTE_THRESHOLD", 0, int.MaxValue, v => Options.ClipboardPasteThreshold = v);
        ApplyInt(Environment.GetEnvironmentVariable("FOCUS_CROP_SIZE"), "FOCUS_CROP_SIZE", 64, 2000, v => Options.FocusCropSize = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_FOCUS_UIA_CROP_PIXELS"), "MAX_FOCUS_UIA_CROP_PIXELS", 0, int.MaxValue, v => Options.MaxFocusUiaCropPixels = v);
        ApplyInt(Environment.GetEnvironmentVariable("OPENAI_MAX_RETRIES"), "OPENAI_MAX_RETRIES", 0, 10, v => Options.OpenAiMaxRetries = v);
        ApplyInt(Environment.GetEnvironmentVariable("OPENAI_TIMEOUT_SECONDS"), "OPENAI_TIMEOUT_SECONDS", 0, int.MaxValue, v => Options.OpenAiTimeoutSeconds = v);
        ApplyBool(Environment.GetEnvironmentVariable("REAL_UI_ONLY"), "REAL_UI_ONLY", v => Options.ForceRealUiOnly = v);
        ApplyBool(Environment.GetEnvironmentVariable("ALLOW_HIGH_LEVEL_ACTIONS"), "ALLOW_HIGH_LEVEL_ACTIONS", v => Options.AllowHighLevelActions = v);
        ApplyBool(Environment.GetEnvironmentVariable("ALLOW_RUN_COMMAND"), "ALLOW_RUN_COMMAND", v => Options.AllowRunCommand = v);
        ApplyBool(Environment.GetEnvironmentVariable("ALLOW_WEB_SEARCH"), "ALLOW_WEB_SEARCH", v => Options.AllowWebSearch = v);
        ApplyBool(Environment.GetEnvironmentVariable("DIRECT_CLICK_WITHOUT_AIM"), "DIRECT_CLICK_WITHOUT_AIM", v => Options.DirectClickWithoutAim = v);
        ApplyBool(Environment.GetEnvironmentVariable("AUTO_HIDE_CONSOLE"), "AUTO_HIDE_CONSOLE", v => Options.AutoHideConsoleDuringRun = v);
        ApplyBool(Environment.GetEnvironmentVariable("MINIMIZE_CONSOLE_DURING_RUN"), "MINIMIZE_CONSOLE_DURING_RUN", v => Options.MinimizeConsoleDuringRun = v);
        ApplyBool(Environment.GetEnvironmentVariable("RESTORE_CONSOLE_AFTER_RUN"), "RESTORE_CONSOLE_AFTER_RUN", v => Options.RestoreConsoleAfterRun = v);
        ApplyBool(Environment.GetEnvironmentVariable("PROMPT_CACHE"), "PROMPT_CACHE", v => Options.UsePromptCache = v);
        ApplyBool(Environment.GetEnvironmentVariable("USE_PREVIOUS_RESPONSE_ID"), "USE_PREVIOUS_RESPONSE_ID", v => Options.UsePreviousResponseState = v);
        ApplyReasoningContext(Environment.GetEnvironmentVariable("CONTROL_REASONING_CONTEXT"), "CONTROL_REASONING_CONTEXT");
        ApplyBool(Environment.GetEnvironmentVariable("CONTROL_CONTEXT_COMPACTION"), "CONTROL_CONTEXT_COMPACTION", v => Options.ControlContextCompactionEnabled = v);
        ApplyInt(Environment.GetEnvironmentVariable("CONTROL_CONTEXT_COMPACT_THRESHOLD"), "CONTROL_CONTEXT_COMPACT_THRESHOLD", 1, int.MaxValue, v => Options.ControlContextCompactThreshold = v);
        ApplyInt(Environment.GetEnvironmentVariable("CONTROL_CONTEXT_FALLBACK_LIMIT"), "CONTROL_CONTEXT_FALLBACK_LIMIT", 1, 20, v => Options.ControlContextFallbackLimit = v);
        ApplyBool(Environment.GetEnvironmentVariable("OMIT_UNCHANGED_SCREEN_IMAGE"), "OMIT_UNCHANGED_SCREEN_IMAGE", v => Options.OmitUnchangedScreenImageWithState = v);
        ApplyBool(Environment.GetEnvironmentVariable("INCLUDE_UIA_TARGETS"), "INCLUDE_UIA_TARGETS", v => Options.IncludeUiaTargets = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_UIA_TARGETS"), "MAX_UIA_TARGETS", 0, 100, v => Options.MaxUiaTargets = v);
        ApplyInt(Environment.GetEnvironmentVariable("UIA_TARGET_NAME_MAX_CHARS"), "UIA_TARGET_NAME_MAX_CHARS", 0, 500, v => Options.UiaTargetNameMaxChars = v);
        ApplyInt(Environment.GetEnvironmentVariable("UIA_SUMMARY_MAX_CHARS"), "UIA_SUMMARY_MAX_CHARS", 0, 2000, v => Options.UiaSummaryMaxChars = v);
        ApplyInt(Environment.GetEnvironmentVariable("UIA_SCAN_TIME_BUDGET_MS"), "UIA_SCAN_TIME_BUDGET_MS", 0, 5000, v => Options.UiaScanTimeBudgetMs = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_UIA_NODES_SCANNED"), "MAX_UIA_NODES_SCANNED", 0, 10000, v => Options.MaxUiaNodesScanned = v);
        ApplyInt(Environment.GetEnvironmentVariable("UIA_CANDIDATE_MULTIPLIER"), "UIA_CANDIDATE_MULTIPLIER", 1, 20, v => Options.UiaCandidateMultiplier = v);
        ApplyDouble(Environment.GetEnvironmentVariable("UIA_MAX_AREA_RATIO"), "UIA_MAX_AREA_RATIO", 0.0, 1.0, v => Options.MaxUiaTargetAreaRatio = v);
        ApplyBool(Environment.GetEnvironmentVariable("REUSE_UIA_TARGETS_ON_NO_CHANGE"), "REUSE_UIA_TARGETS_ON_NO_CHANGE", v => Options.ReuseUiaTargetsWhenScreenUnchanged = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_ARTIFACTS_PER_DIR"), "MAX_ARTIFACTS_PER_DIR", 0, int.MaxValue, v => Options.MaxArtifactsPerDir = v);
        ApplyBool(Environment.GetEnvironmentVariable("EXECUTE_MULTI_ACTION_CANDIDATES"), "EXECUTE_MULTI_ACTION_CANDIDATES", v => Options.ExecuteMultiActionCandidates = v);
        ApplyInt(Environment.GetEnvironmentVariable("MAX_QUEUED_BATCH_ACTIONS"), "MAX_QUEUED_BATCH_ACTIONS", 0, 20, v => Options.MaxQueuedBatchActions = v);
        ApplyInt(Environment.GetEnvironmentVariable("TURN_BASED_MAX_BATCH_INPUTS"), "TURN_BASED_MAX_BATCH_INPUTS", 2, 64, v => Options.TurnBasedMaxBatchInputs = v);
        var envPromptCacheKey = Environment.GetEnvironmentVariable("PROMPT_CACHE_KEY");
        if (!string.IsNullOrWhiteSpace(envPromptCacheKey))
            Options.PromptCacheKey = envPromptCacheKey.Trim();
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

    internal static string? ResolveAutomaticConfigPath(
        string baseDirectory,
        string currentDirectory)
    {
        var cwdPath = Path.Combine(currentDirectory, "appsettings.json");
        var localPath = Path.Combine(baseDirectory, "appsettings.json");
        return File.Exists(cwdPath)
            ? cwdPath
            : File.Exists(localPath)
                ? localPath
                : null;
    }

    internal static void ApplyAutomaticConfigFile(
        string? baseDirectory = null,
        string? currentDirectory = null)
    {
        var path = ResolveAutomaticConfigPath(
            baseDirectory ?? AppContext.BaseDirectory,
            currentDirectory ?? Environment.CurrentDirectory);
        ApplyConfigFile(path);
    }

    internal static void ApplyConfigFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        if (!File.Exists(path))
        {
            ReportConfigurationError($"Config file was not found: '{path}'.");
            return;
        }

        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(path));
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                throw new JsonException("Expected a configuration object.");
            var values = new JsonConfigurationValues(root, path);

            if (values.TryString("profile", out var profile)) ApplyProfile(profile);
            if (values.TryString("model", out var model)) Options.Model = model;
            if (values.TryString("apiUrl", out var apiUrl)) ApplyApiUrl(apiUrl, path);
            if (values.TryOptionalString("qaModel", out var qaModel)) Options.QaModel = qaModel;
            if (values.TryOptionalString("verifyModel", out var verifyModel)) Options.VerifyModel = verifyModel;
            if (values.TryEffort("reasoningEffort", out var effort)) ApplyReasoningEffort(effort, path);
            if (values.TryEffort("qaReasoningEffort", out var qaEffort)) ApplyReasoningEffort(qaEffort, path, v => Options.QaReasoningEffort = v, () => Options.QaReasoningEffortExplicit = root.GetProperty("qaReasoningEffort").ValueKind != JsonValueKind.Null);
            if (values.TryEffort("verifyReasoningEffort", out var verifyEffort)) ApplyReasoningEffort(verifyEffort, path, v => Options.VerifyReasoningEffort = v, () => Options.VerifyReasoningEffortExplicit = root.GetProperty("verifyReasoningEffort").ValueKind != JsonValueKind.Null);
            if (values.TryBool("mouseEnabled", out var mouseEnabled)) Options.MouseEnabled = mouseEnabled;
            if (values.TryBool("multiMonitorEnabled", out var multiMonitorEnabled)) Options.MultiMonitorEnabled = multiMonitorEnabled;
            if (values.TryNumber("postActionDelayMs", out var delay)) ApplyInt(delay, $"{path}:postActionDelayMs", 0, int.MaxValue, v => Options.UiSettleDelayMs = v);
            if (values.TryNumber("gridStepPx", out var grid)) ApplyInt(grid, $"{path}:gridStepPx", 0, int.MaxValue, v => Options.GridStepPx = v);
            if (values.TryNumber("gridLabelEveryPx", out var gridLabelEvery)) ApplyInt(gridLabelEvery, $"{path}:gridLabelEveryPx", 0, int.MaxValue, v => Options.GridLabelEveryPx = v);
            if (values.TryNumber("gridMajorEveryPx", out var gridMajorEvery)) ApplyInt(gridMajorEvery, $"{path}:gridMajorEveryPx", 0, int.MaxValue, v => Options.GridMajorEveryPx = v);
            if (values.TryBool("sendFocusCrop", out var sendFocusCrop)) Options.SendFocusCrop = sendFocusCrop;
            if (values.TryNumber("focusRingPadding", out var focusRingPadding)) ApplyInt(focusRingPadding, $"{path}:focusRingPadding", 0, 100, v => Options.FocusRingPadding = v);
            if (values.TryNumber("focusRingThickness", out var focusRingThickness)) ApplyInt(focusRingThickness, $"{path}:focusRingThickness", 1, 100, v => Options.FocusRingThickness = v);
            if (values.TryNumber("focusGlowThickness", out var focusGlowThickness)) ApplyInt(focusGlowThickness, $"{path}:focusGlowThickness", 1, 100, v => Options.FocusGlowThickness = v);
            if (values.TryNumber("focusCornerRadius", out var focusCornerRadius)) ApplyInt(focusCornerRadius, $"{path}:focusCornerRadius", 0, 500, v => Options.FocusCornerRadius = v);
            if (values.TryNumber("largeFocusOverlayAreaRatio", out var largeFocusAreaRatio)) ApplyDouble(largeFocusAreaRatio, $"{path}:largeFocusOverlayAreaRatio", 0.0, 1.0, v => Options.LargeFocusOverlayAreaRatio = v);
            if (values.TryNumber("clickAimEdgeAdjustMaxAreaRatio", out var clickAimMaxAreaRatio)) ApplyDouble(clickAimMaxAreaRatio, $"{path}:clickAimEdgeAdjustMaxAreaRatio", 0.0, 1.0, v => Options.ClickAimEdgeAdjustMaxAreaRatio = v);
            if (values.TryNumber("clickAimEdgeMarginRatio", out var clickAimMarginRatio)) ApplyDouble(clickAimMarginRatio, $"{path}:clickAimEdgeMarginRatio", 0.0, 1.0, v => Options.ClickAimEdgeMarginRatio = v);
            if (values.TryNumber("clickAimEdgeMinMarginPx", out var clickAimMinMargin)) ApplyInt(clickAimMinMargin, $"{path}:clickAimEdgeMinMarginPx", 0, 2000, v => Options.ClickAimEdgeMinMarginPx = v);
            if (values.TryNumber("clickAimEdgeMaxMarginPx", out var clickAimMaxMargin)) ApplyInt(clickAimMaxMargin, $"{path}:clickAimEdgeMaxMarginPx", 0, 2000, v => Options.ClickAimEdgeMaxMarginPx = v);
            if (values.TryNumber("aimExpireDelta", out var aimExpireDelta)) ApplyDouble(aimExpireDelta, $"{path}:aimExpireDelta", 0.0, 1.0, v => Options.AimExpireDelta = v);
            if (values.TryNumber("noChangeThreshold", out var noChangeThreshold)) ApplyDouble(noChangeThreshold, $"{path}:noChangeThreshold", 0.0, 1.0, v => Options.NoChangeThreshold = v);
            if (values.TryNumber("maxGesturePathPoints", out var maxGesturePoints)) ApplyInt(maxGesturePoints, $"{path}:maxGesturePathPoints", 2, 1024, v => Options.MaxGesturePathPoints = v);
            if (values.TryNumber("maxGestureDurationMs", out var maxGestureDuration)) ApplyInt(maxGestureDuration, $"{path}:maxGestureDurationMs", 100, 60000, v => Options.MaxGestureDurationMs = v);
            if (values.TryNumber("maxHeldKeys", out var maxHeldKeys)) ApplyInt(maxHeldKeys, $"{path}:maxHeldKeys", 1, 16, v => Options.MaxHeldKeys = v);
            if (values.TryNumber("maxKeyHoldDurationMs", out var maxKeyHoldDuration)) ApplyInt(maxKeyHoldDuration, $"{path}:maxKeyHoldDurationMs", 100, 60000, v => Options.MaxKeyHoldDurationMs = v);
            if (values.TryNumber("maxSteps", out var maxSteps)) ApplyInt(maxSteps, $"{path}:maxSteps", 0, int.MaxValue, v => Options.MaxSteps = v);
            if (values.TryNumber("maxWaitSeconds", out var maxWait)) ApplyInt(maxWait, $"{path}:maxWaitSeconds", 0, int.MaxValue, v => Options.MaxWaitSeconds = v);
            if (values.TryNumber("maxOutputTokens", out var maxOutput)) ApplyInt(maxOutput, $"{path}:maxOutputTokens", 1, int.MaxValue, v => Options.MaxOutputTokens = v);
            if (values.TryNumber("qaMaxOutputTokens", out var qaMaxOutput)) ApplyInt(qaMaxOutput, $"{path}:qaMaxOutputTokens", 1, int.MaxValue, v => Options.QaMaxOutputTokens = v);
            if (values.TryNumber("verifyMaxOutputTokens", out var verifyMaxOutput)) ApplyInt(verifyMaxOutput, $"{path}:verifyMaxOutputTokens", 1, int.MaxValue, v => Options.VerifyMaxOutputTokens = v);
            if (values.TryNumber("turnReanalysisMaxOutputTokens", out var turnReanalysisMaxOutput)) ApplyInt(turnReanalysisMaxOutput, $"{path}:turnReanalysisMaxOutputTokens", 1, int.MaxValue, v => Options.TurnReanalysisMaxOutputTokens = v);
            if (values.TryNumber("incompleteMaxOutputRetries", out var incompleteRetries)) ApplyInt(incompleteRetries, $"{path}:incompleteMaxOutputRetries", 0, 5, v => Options.IncompleteMaxOutputRetries = v);
            if (values.TryNumber("incompleteMaxOutputTokenCap", out var incompleteTokenCap)) ApplyInt(incompleteTokenCap, $"{path}:incompleteMaxOutputTokenCap", 1, int.MaxValue, v => Options.IncompleteMaxOutputTokenCap = v);
            if (values.TryNumber("maxActionTextChars", out var maxActionTextChars)) ApplyInt(maxActionTextChars, $"{path}:maxActionTextChars", 256, int.MaxValue, v => Options.MaxActionTextChars = v);
            if (values.TryNumber("qaScreenshotMaxWidth", out var qaScreenshotMaxWidth)) ApplyInt(qaScreenshotMaxWidth, $"{path}:qaScreenshotMaxWidth", 0, 10000, v => Options.QaScreenshotMaxWidth = v);
            if (values.TryNumber("verifyScreenshotMaxWidth", out var verifyScreenshotMaxWidth)) ApplyInt(verifyScreenshotMaxWidth, $"{path}:verifyScreenshotMaxWidth", 0, 10000, v => Options.VerifyScreenshotMaxWidth = v);
            if (values.TryString("textVerbosity", out var verbosity)) ApplyTextVerbosity(verbosity, path);
            if (values.TryNumber("historyTailChars", out var historyChars)) ApplyInt(historyChars, $"{path}:historyTailChars", 0, int.MaxValue, v => Options.HistoryTailChars = v);
            if (values.TryNumber("historyTailLines", out var historyLines)) ApplyInt(historyLines, $"{path}:historyTailLines", 0, int.MaxValue, v => Options.HistoryTailLines = v);
            if (values.TryNumber("maxStagnationSteps", out var maxStagnation)) ApplyInt(maxStagnation, $"{path}:maxStagnationSteps", 0, int.MaxValue, v => Options.MaxStagnationStepsBeforeAbort = v);
            if (values.TryNumber("maxRepeatedActions", out var maxRepeated)) ApplyInt(maxRepeated, $"{path}:maxRepeatedActions", 0, int.MaxValue, v => Options.MaxRepeatedActionBeforeAbort = v);
            if (values.TryNumber("actionRepeatCooldownSteps", out var repeatCooldown)) ApplyInt(repeatCooldown, $"{path}:actionRepeatCooldownSteps", 0, int.MaxValue, v => Options.ActionRepeatCooldownSteps = v);
            if (values.TryNumber("maxRejectedProposalRepeats", out var maxRejectedProposals)) ApplyInt(maxRejectedProposals, $"{path}:maxRejectedProposalRepeats", 0, int.MaxValue, v => Options.MaxRejectedProposalRepeatsBeforeAbort = v);
            if (values.TryNumber("maxConsecutiveInspectionActions", out var maxInspectionActions)) ApplyInt(maxInspectionActions, $"{path}:maxConsecutiveInspectionActions", 0, 20, v => Options.MaxConsecutiveInspectionActions = v);
            if (values.TryNumber("ineffectiveMouseClusterPx", out var ineffectiveMouseCluster)) ApplyInt(ineffectiveMouseCluster, $"{path}:ineffectiveMouseClusterPx", 8, 2000, v => Options.IneffectiveMouseClusterPx = v);
            if (values.TryNumber("skipVerifyConfidenceThreshold", out var skipVerifyConfidence)) ApplyDouble(skipVerifyConfidence, $"{path}:skipVerifyConfidenceThreshold", 0.0, 1.0, v => Options.SkipVerifyConfidenceThreshold = v);
            if (values.TryNumber("maxModelFailures", out var maxModelFailures)) ApplyInt(maxModelFailures, $"{path}:maxModelFailures", 0, int.MaxValue, v => Options.MaxModelFailuresBeforeAbort = v);
            if (values.TryNumber("maxActionFailures", out var maxActionFailures)) ApplyInt(maxActionFailures, $"{path}:maxActionFailures", 0, int.MaxValue, v => Options.MaxActionFailuresBeforeAbort = v);
            if (values.TryString("goalMode", out var goalMode)) ApplyGoalMode(goalMode, path);
            if (values.TryString("observationProfile", out var observationProfile)) ApplyObservationMode(observationProfile, path);
            if (values.TryBool("observationLogVerbose", out var observationLogVerbose)) Options.ObservationLogVerbose = observationLogVerbose;
            if (values.TryBool("recoveryMemory", out var recoveryMemory)) Options.RecoveryMemoryEnabled = recoveryMemory;
            if (values.TryOptionalString("recoveryMemoryPath", out var recoveryMemoryPath)) Options.RecoveryMemoryPath = recoveryMemoryPath;
            if (values.TryNumber("recoveryMemoryTriggerSteps", out var recoveryTrigger)) ApplyInt(recoveryTrigger, $"{path}:recoveryMemoryTriggerSteps", 1, int.MaxValue, v => Options.RecoveryMemoryTriggerSteps = v);
            if (values.TryNumber("recoveryMemoryValidationSteps", out var recoveryValidation)) ApplyInt(recoveryValidation, $"{path}:recoveryMemoryValidationSteps", 1, int.MaxValue, v => Options.RecoveryMemoryValidationSteps = v);
            if (values.TryNumber("recoveryMemoryMaxLessons", out var recoveryMaxLessons)) ApplyInt(recoveryMaxLessons, $"{path}:recoveryMemoryMaxLessons", 1, int.MaxValue, v => Options.RecoveryMemoryMaxLessons = v);
            if (values.TryNumber("recoveryMemoryMaxQuarantinedLessons", out var recoveryMaxQuarantined)) ApplyInt(recoveryMaxQuarantined, $"{path}:recoveryMemoryMaxQuarantinedLessons", 1, int.MaxValue, v => Options.RecoveryMemoryMaxQuarantinedLessons = v);
            if (values.TryNumber("recoveryMemoryReservedLessonsPerContext", out var recoveryReservedPerContext)) ApplyInt(recoveryReservedPerContext, $"{path}:recoveryMemoryReservedLessonsPerContext", 0, int.MaxValue, v => Options.RecoveryMemoryReservedLessonsPerContext = v);
            if (values.TryNumber("recoveryMemorySoftMaxLessonsPerContext", out var recoverySoftMaxPerContext)) ApplyInt(recoverySoftMaxPerContext, $"{path}:recoveryMemorySoftMaxLessonsPerContext", 1, int.MaxValue, v => Options.RecoveryMemorySoftMaxLessonsPerContext = v);
            if (values.TryNumber("recoveryMemoryMaxFileBytes", out var recoveryMaxFileBytes)) ApplyLong(recoveryMaxFileBytes, $"{path}:recoveryMemoryMaxFileBytes", 1024 * 1024, int.MaxValue, v => Options.RecoveryMemoryMaxFileBytes = v);
            if (values.TryOptionalString("recoveryMemoryArchivePath", out var recoveryArchivePath)) Options.RecoveryMemoryArchivePath = recoveryArchivePath;
            if (values.TryNumber("recoveryMemoryArchiveMaxBytes", out var recoveryArchiveMaxBytes)) ApplyLong(recoveryArchiveMaxBytes, $"{path}:recoveryMemoryArchiveMaxBytes", 1024 * 1024, int.MaxValue, v => Options.RecoveryMemoryArchiveMaxBytes = v);
            if (values.TryNumber("recoveryMemoryArchiveRetainedFiles", out var recoveryArchiveRetainedFiles)) ApplyInt(recoveryArchiveRetainedFiles, $"{path}:recoveryMemoryArchiveRetainedFiles", 1, 20, v => Options.RecoveryMemoryArchiveRetainedFiles = v);
            if (values.TryNumber("recoveryMemoryPromptLessons", out var recoveryPromptLessons)) ApplyInt(recoveryPromptLessons, $"{path}:recoveryMemoryPromptLessons", 0, 10, v => Options.RecoveryMemoryPromptMaxLessons = v);
            if (values.TryNumber("recoveryMemoryFailureLimit", out var recoveryFailureLimit)) ApplyInt(recoveryFailureLimit, $"{path}:recoveryMemoryFailureLimit", 1, int.MaxValue, v => Options.RecoveryMemoryFailureLimit = v);
            if (values.TryNumber("runtimeSemanticStateLimit", out var semanticStateLimit)) ApplyInt(semanticStateLimit, $"{path}:runtimeSemanticStateLimit", 32, 100000, v => Options.RuntimeSemanticStateLimit = v);
            if (values.TryNumber("runtimeGraphEdgeLimit", out var graphEdgeLimit)) ApplyInt(graphEdgeLimit, $"{path}:runtimeGraphEdgeLimit", 32, 100000, v => Options.RuntimeGraphEdgeLimit = v);
            if (values.TryNumber("runtimeRecoveryActionLimit", out var recoveryActionLimit)) ApplyInt(recoveryActionLimit, $"{path}:runtimeRecoveryActionLimit", 8, 10000, v => Options.RuntimeRecoveryActionLimit = v);
            if (values.TryNumber("runtimeCooldownEntryLimit", out var cooldownLimit)) ApplyInt(cooldownLimit, $"{path}:runtimeCooldownEntryLimit", 16, 100000, v => Options.RuntimeCooldownEntryLimit = v);
            if (values.TryNumber("graphCandidateTtlSteps", out var graphCandidateTtl)) ApplyInt(graphCandidateTtl, $"{path}:graphCandidateTtlSteps", 2, 10000, v => Options.GraphCandidateTtlSteps = v);
            if (values.TryNumber("proactiveLoopConfidenceThreshold", out var loopConfidenceThreshold)) ApplyDouble(loopConfidenceThreshold, $"{path}:proactiveLoopConfidenceThreshold", 0.5, 1.0, v => Options.ProactiveLoopConfidenceThreshold = v);
            if (values.TryBool("recoveryProgressVerification", out var recoveryProgressVerification)) Options.RecoveryProgressVerificationEnabled = recoveryProgressVerification;
            if (values.TryNumber("recoveryProgressConfidenceThreshold", out var recoveryProgressConfidence)) ApplyDouble(recoveryProgressConfidence, $"{path}:recoveryProgressConfidenceThreshold", 0.5, 1.0, v => Options.RecoveryProgressConfidenceThreshold = v);
            if (values.TryNumber("recoveryTelemetryMaxBytes", out var recoveryTelemetryMaxBytes)) ApplyInt(recoveryTelemetryMaxBytes, $"{path}:recoveryTelemetryMaxBytes", 65536, int.MaxValue, v => Options.RecoveryTelemetryMaxBytes = v);
            if (values.TryNumber("recoveryTelemetryRetainedFiles", out var recoveryTelemetryRetainedFiles)) ApplyInt(recoveryTelemetryRetainedFiles, $"{path}:recoveryTelemetryRetainedFiles", 1, 20, v => Options.RecoveryTelemetryRetainedFiles = v);
            if (values.TryBool("loopReplayAutoExport", out var loopReplayAutoExport)) Options.LoopReplayAutoExportEnabled = loopReplayAutoExport;
            if (values.TryOptionalString("loopReplayCorpusPath", out var loopReplayCorpusPath)) Options.LoopReplayCorpusPath = loopReplayCorpusPath;
            if (values.TryBool("screenPolling", out var screenPolling)) Options.ScreenPollingEnabled = screenPolling;
            if (values.TryNumber("screenPollInitialDelayMs", out var pollInitialDelay)) ApplyInt(pollInitialDelay, $"{path}:screenPollInitialDelayMs", 0, 10000, v => Options.ScreenPollInitialDelayMs = v);
            if (values.TryNumber("screenPollIntervalMs", out var pollInterval)) ApplyInt(pollInterval, $"{path}:screenPollIntervalMs", 10, 10000, v => Options.ScreenPollIntervalMs = v);
            if (values.TryNumber("screenPollTimeoutMs", out var pollTimeout)) ApplyInt(pollTimeout, $"{path}:screenPollTimeoutMs", 0, 60000, v => Options.ScreenPollTimeoutMs = v);
            if (values.TryNumber("waitNoChangeExtraMs", out var waitNoChangeExtra)) ApplyInt(waitNoChangeExtra, $"{path}:waitNoChangeExtraMs", 0, 60000, v => Options.WaitNoChangeExtraMs = v);
            if (values.TryBool("screenSanityChecks", out var screenSanity)) Options.ScreenSanityChecks = screenSanity;
            if (values.TryNumber("sendInputMaxRetries", out var sendInputRetries)) ApplyInt(sendInputRetries, $"{path}:sendInputMaxRetries", 0, 10, v => Options.SendInputMaxRetries = v);
            if (values.TryNumber("sendInputRetryDelayMs", out var sendInputRetryDelay)) ApplyInt(sendInputRetryDelay, $"{path}:sendInputRetryDelayMs", 0, 1000, v => Options.SendInputRetryDelayMs = v);
            if (values.TryBool("adaptiveReasoningEffort", out var adaptiveEffort)) Options.AdaptiveReasoningEffort = adaptiveEffort;
            if (values.TryNumber("screenshotMaxWidth", out var maxWidth)) ApplyInt(maxWidth, $"{path}:screenshotMaxWidth", 0, 10000, v => Options.MaxScreenshotSendWidth = v);
            if (values.TryString("screenshotFormat", out var format)) ApplyImageFormat(format, path);
            if (values.TryNumber("focusedOverviewMaxWidth", out var focusedOverviewMaxWidth)) ApplyInt(focusedOverviewMaxWidth, $"{path}:focusedOverviewMaxWidth", 0, 10000, v => Options.FocusedOverviewMaxWidth = v);
            if (values.TryNumber("cropMaxWidth", out var cropMaxWidth)) ApplyInt(cropMaxWidth, $"{path}:cropMaxWidth", 0, 10000, v => Options.MaxCropSendWidth = v);
            if (values.TryString("cropFormat", out var cropFormat)) ApplyCropFormat(cropFormat, path);
            if (values.TryNumber("screenshotJpegQuality", out var quality)) ApplyLong(quality, $"{path}:screenshotJpegQuality", 1, 100, v => Options.ScreenshotJpegQuality = v);
            if (values.TryString("screenLogFormat", out var screenLogFormat)) ApplyScreenLogFormat(screenLogFormat, path);
            if (values.TryNumber("screenLogMaxWidth", out var screenLogMaxWidth)) ApplyInt(screenLogMaxWidth, $"{path}:screenLogMaxWidth", 0, 10000, v => Options.MaxScreenLogWidth = v);
            if (values.TryBool("includeFocusUia", out var focusUia)) Options.IncludeFocusUia = focusUia;
            if (values.TryBool("includeFocusUiaCrop", out var focusCrop)) Options.IncludeFocusUiaCrop = focusCrop;
            if (values.TryBool("debugImages", out var debugImages)) Options.DebugImages = debugImages;
            if (values.TryString("verifyMode", out var verifyMode)) ApplyVerifyMode(verifyMode, path);
            if (values.TryNumber("verifyEarlySteps", out var verifyEarlySteps)) ApplyInt(verifyEarlySteps, $"{path}:verifyEarlySteps", 0, int.MaxValue, v => Options.VerifyEarlySteps = v);
            if (values.TryNumber("verifyLowConfidenceThreshold", out var verifyLowConfidence)) ApplyDouble(verifyLowConfidence, $"{path}:verifyLowConfidenceThreshold", 0.0, 1.0, v => Options.VerifyLowConfidenceThreshold = v);
            if (values.TryBool("refreshScreenshotBeforeVerify", out var refreshBeforeVerify)) Options.RefreshScreenshotBeforeVerify = refreshBeforeVerify;
            if (values.TryNumber("clipboardPasteThreshold", out var pasteThreshold)) ApplyInt(pasteThreshold, $"{path}:clipboardPasteThreshold", 0, int.MaxValue, v => Options.ClipboardPasteThreshold = v);
            if (values.TryNumber("focusCropSize", out var focusCropSize)) ApplyInt(focusCropSize, $"{path}:focusCropSize", 64, 2000, v => Options.FocusCropSize = v);
            if (values.TryNumber("maxFocusUiaCropPixels", out var maxFocusCropPixels)) ApplyInt(maxFocusCropPixels, $"{path}:maxFocusUiaCropPixels", 0, int.MaxValue, v => Options.MaxFocusUiaCropPixels = v);
            if (values.TryNumber("openAiMaxRetries", out var retries)) ApplyInt(retries, $"{path}:openAiMaxRetries", 0, 10, v => Options.OpenAiMaxRetries = v);
            if (values.TryNumber("openAiTimeoutSeconds", out var timeoutSeconds)) ApplyInt(timeoutSeconds, $"{path}:openAiTimeoutSeconds", 0, int.MaxValue, v => Options.OpenAiTimeoutSeconds = v);
            if (values.TryBool("realUiOnly", out var realUiOnly)) Options.ForceRealUiOnly = realUiOnly;
            if (values.TryBool("allowHighLevelActions", out var allowHighLevel)) Options.AllowHighLevelActions = allowHighLevel;
            if (values.TryBool("allowRunCommand", out var allowRun)) Options.AllowRunCommand = allowRun;
            if (values.TryBool("allowWebSearch", out var allowWebSearch)) Options.AllowWebSearch = allowWebSearch;
            if (values.TryBool("directClickWithoutAim", out var directClick)) Options.DirectClickWithoutAim = directClick;
            if (values.TryBool("autoHideConsoleDuringRun", out var autoHideConsole)) Options.AutoHideConsoleDuringRun = autoHideConsole;
            if (values.TryBool("minimizeConsoleDuringRun", out var minimize)) Options.MinimizeConsoleDuringRun = minimize;
            if (values.TryBool("restoreConsoleAfterRun", out var restoreConsole)) Options.RestoreConsoleAfterRun = restoreConsole;
            if (values.TryBool("promptCache", out var promptCache)) Options.UsePromptCache = promptCache;
            if (values.TryOptionalString("promptCacheKey", out var promptCacheKey)) Options.PromptCacheKey = promptCacheKey;
            if (values.TryBool("usePreviousResponseId", out var previousResponseState)) Options.UsePreviousResponseState = previousResponseState;
            if (values.TryString("controlReasoningContext", out var controlReasoningContext)) ApplyReasoningContext(controlReasoningContext, path);
            if (values.TryBool("controlContextCompaction", out var contextCompaction)) Options.ControlContextCompactionEnabled = contextCompaction;
            if (values.TryNumber("controlContextCompactThreshold", out var contextCompactThreshold)) ApplyInt(contextCompactThreshold, $"{path}:controlContextCompactThreshold", 1, int.MaxValue, v => Options.ControlContextCompactThreshold = v);
            if (values.TryNumber("controlContextFallbackLimit", out var contextFallbackLimit)) ApplyInt(contextFallbackLimit, $"{path}:controlContextFallbackLimit", 1, 20, v => Options.ControlContextFallbackLimit = v);
            if (values.TryBool("omitUnchangedScreenImage", out var omitUnchangedScreen)) Options.OmitUnchangedScreenImageWithState = omitUnchangedScreen;
            if (values.TryBool("includeUiaTargets", out var includeUia)) Options.IncludeUiaTargets = includeUia;
            if (values.TryNumber("maxUiaTargets", out var maxUia)) ApplyInt(maxUia, $"{path}:maxUiaTargets", 0, 100, v => Options.MaxUiaTargets = v);
            if (values.TryNumber("uiaTargetNameMaxChars", out var uiaNameChars)) ApplyInt(uiaNameChars, $"{path}:uiaTargetNameMaxChars", 0, 500, v => Options.UiaTargetNameMaxChars = v);
            if (values.TryNumber("uiaSummaryMaxChars", out var uiaSummaryChars)) ApplyInt(uiaSummaryChars, $"{path}:uiaSummaryMaxChars", 0, 2000, v => Options.UiaSummaryMaxChars = v);
            if (values.TryNumber("uiaScanTimeBudgetMs", out var uiaScanMs)) ApplyInt(uiaScanMs, $"{path}:uiaScanTimeBudgetMs", 0, 5000, v => Options.UiaScanTimeBudgetMs = v);
            if (values.TryNumber("maxUiaNodesScanned", out var maxUiaNodes)) ApplyInt(maxUiaNodes, $"{path}:maxUiaNodesScanned", 0, 10000, v => Options.MaxUiaNodesScanned = v);
            if (values.TryNumber("uiaCandidateMultiplier", out var candidateMultiplier)) ApplyInt(candidateMultiplier, $"{path}:uiaCandidateMultiplier", 1, 20, v => Options.UiaCandidateMultiplier = v);
            if (values.TryNumber("uiaMaxAreaRatio", out var maxAreaRatio)) ApplyDouble(maxAreaRatio, $"{path}:uiaMaxAreaRatio", 0.0, 1.0, v => Options.MaxUiaTargetAreaRatio = v);
            if (values.TryBool("reuseUiaTargetsWhenScreenUnchanged", out var reuseUia)) Options.ReuseUiaTargetsWhenScreenUnchanged = reuseUia;
            if (values.TryBool("executeMultiActionCandidates", out var executeCandidates)) Options.ExecuteMultiActionCandidates = executeCandidates;
            if (values.TryNumber("maxQueuedBatchActions", out var maxQueued)) ApplyInt(maxQueued, $"{path}:maxQueuedBatchActions", 0, 20, v => Options.MaxQueuedBatchActions = v);
            if (values.TryNumber("turnBasedMaxBatchInputs", out var turnBatchInputs)) ApplyInt(turnBatchInputs, $"{path}:turnBasedMaxBatchInputs", 2, 64, v => Options.TurnBasedMaxBatchInputs = v);
            if (values.TryNumber("maxBatchedGesturePoints", out var maxBatchedGesturePoints)) ApplyInt(maxBatchedGesturePoints, $"{path}:maxBatchedGesturePoints", 2, 10000, v => Options.MaxBatchedGesturePoints = v);
            if (values.TryNumber("maxBatchedGestureDurationMs", out var maxBatchedGestureDuration)) ApplyInt(maxBatchedGestureDuration, $"{path}:maxBatchedGestureDurationMs", 100, 60000, v => Options.MaxBatchedGestureDurationMs = v);
            if (values.TryNumber("maxArtifactsPerDir", out var maxArtifacts)) ApplyInt(maxArtifacts, $"{path}:maxArtifactsPerDir", 0, int.MaxValue, v => Options.MaxArtifactsPerDir = v);
            if (values.TryBool("logRequests", out var logRequests)) Options.LogRequests = logRequests;
            if (values.TryBool("prettyRequestLogs", out var prettyRequestLogs)) Options.PrettyRequestLogs = prettyRequestLogs;
            if (values.TryBool("logScreens", out var logScreens)) Options.LogScreens = logScreens;
        }
        catch (Exception ex)
        {
            ReportConfigurationError($"Could not read config file '{path}': {ex.Message}");
        }

    }
}
