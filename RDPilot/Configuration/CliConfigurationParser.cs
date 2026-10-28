namespace RDPilot.Configuration;

/// <summary>Applies CLI profiles first, then explicit arguments in their original order.</summary>
internal static class CliConfigurationParser
{
    internal static CliValidationResult ParseCliArgs(string[] args)
    {
        var errorStart = CurrentRun.ConfigurationErrors.Count;
        CurrentRun.CliArgumentError = false;
        var arguments = CliArguments.Parse(args, ReportConfigurationError);
        ApplyCliProfileArgs(arguments);
        var positional = new List<string>();
        string? explicitTask = null;

        void SetExplicitTask(string value, string source)
        {
            Options.BatchMode = true;
            if (string.IsNullOrWhiteSpace(value))
            {
                ReportConfigurationError($"Empty task from {source}.");
                return;
            }
            if (explicitTask is not null)
            {
                ReportConfigurationError($"Task was specified more than once; duplicate source: {source}.");
                return;
            }
            explicitTask = value;
        }

        foreach (var argument in arguments)
        {
            var arg = argument.Name;

            if (arg.Equals("--batch", StringComparison.OrdinalIgnoreCase))
            {
                Options.BatchMode = true;
                continue;
            }
            if (argument.TryRead("--task", out var task) ||
                argument.TryRead("--goal", out task))
            {
                SetExplicitTask(task, arg);
                continue;
            }
            if (argument.TryRead("--task-file", out var taskFile))
            {
                Options.BatchMode = true;
                if (string.IsNullOrWhiteSpace(taskFile))
                {
                    CurrentRun.CliArgumentError = true;
                    continue;
                }
                try
                {
                    var taskText = taskFile == "-"
                        ? Console.In.ReadToEnd()
                        : File.ReadAllText(Path.GetFullPath(taskFile));
                    SetExplicitTask(taskText, arg);
                }
                catch (Exception ex)
                {
                    ReportConfigurationError($"Could not read task file '{taskFile}': {ex.Message}");
                }
                continue;
            }

            if (arg.Equals("--mouse", StringComparison.OrdinalIgnoreCase))
            {
                Options.MouseEnabled = true;
                continue;
            }
            if (arg.Equals("--no-mouse", StringComparison.OrdinalIgnoreCase))
            {
                Options.MouseEnabled = false;
                continue;
            }
            if (arg.Equals("--multi-monitor", StringComparison.OrdinalIgnoreCase))
            {
                Options.MultiMonitorEnabled = true;
                continue;
            }
            if (arg.Equals("--primary-monitor-only", StringComparison.OrdinalIgnoreCase))
            {
                Options.MultiMonitorEnabled = false;
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
                Options.DebugImages = true;
                Options.IncludeFocusUia = true;
                Options.IncludeFocusUiaCrop = true;
                continue;
            }
            if (arg.Equals("--focus-uia", StringComparison.OrdinalIgnoreCase))
            {
                Options.IncludeFocusUia = true;
                continue;
            }
            if (arg.Equals("--no-focus-uia", StringComparison.OrdinalIgnoreCase))
            {
                Options.IncludeFocusUia = false;
                Options.IncludeFocusUiaCrop = false;
                continue;
            }
            if (arg.Equals("--no-request-logs", StringComparison.OrdinalIgnoreCase))
            {
                Options.LogRequests = false;
                continue;
            }
            if (arg.Equals("--pretty-request-logs", StringComparison.OrdinalIgnoreCase))
            {
                Options.PrettyRequestLogs = true;
                continue;
            }
            if (arg.Equals("--compact-request-logs", StringComparison.OrdinalIgnoreCase))
            {
                Options.PrettyRequestLogs = false;
                continue;
            }
            if (arg.Equals("--no-screen-logs", StringComparison.OrdinalIgnoreCase))
            {
                Options.LogScreens = false;
                continue;
            }
            if (arg.Equals("--real-ui-only", StringComparison.OrdinalIgnoreCase))
            {
                Options.ForceRealUiOnly = true;
                continue;
            }
            if (arg.Equals("--allow-run-command", StringComparison.OrdinalIgnoreCase))
            {
                Options.ForceRealUiOnly = false;
                Options.AllowRunCommand = true;
                continue;
            }
            if (arg.Equals("--allow-high-level-actions", StringComparison.OrdinalIgnoreCase))
            {
                Options.ForceRealUiOnly = false;
                Options.AllowHighLevelActions = true;
                continue;
            }
            if (arg.Equals("--allow-web-search", StringComparison.OrdinalIgnoreCase))
            {
                Options.AllowWebSearch = true;
                continue;
            }
            if (arg.Equals("--no-web-search", StringComparison.OrdinalIgnoreCase))
            {
                Options.AllowWebSearch = false;
                continue;
            }
            if (arg.Equals("--no-direct-click", StringComparison.OrdinalIgnoreCase))
            {
                Options.DirectClickWithoutAim = false;
                continue;
            }
            if (arg.Equals("--minimize-console", StringComparison.OrdinalIgnoreCase))
            {
                Options.MinimizeConsoleDuringRun = true;
                continue;
            }
            if (arg.Equals("--auto-hide-console", StringComparison.OrdinalIgnoreCase))
            {
                Options.AutoHideConsoleDuringRun = true;
                continue;
            }
            if (arg.Equals("--no-auto-hide-console", StringComparison.OrdinalIgnoreCase))
            {
                Options.AutoHideConsoleDuringRun = false;
                continue;
            }
            if (arg.Equals("--restore-console", StringComparison.OrdinalIgnoreCase))
            {
                Options.RestoreConsoleAfterRun = true;
                continue;
            }
            if (arg.Equals("--no-restore-console", StringComparison.OrdinalIgnoreCase))
            {
                Options.RestoreConsoleAfterRun = false;
                continue;
            }
            if (arg.Equals("--analyze-logs", StringComparison.OrdinalIgnoreCase))
            {
                Options.AnalyzeLogsOnly = true;
                continue;
            }
            if (arg.Equals("--print-config", StringComparison.OrdinalIgnoreCase))
            {
                Options.PrintConfigOnly = true;
                continue;
            }
            if (arg.Equals("--batch-candidates", StringComparison.OrdinalIgnoreCase))
            {
                Options.ExecuteMultiActionCandidates = true;
                continue;
            }
            if (arg.Equals("--no-batch-candidates", StringComparison.OrdinalIgnoreCase))
            {
                Options.ExecuteMultiActionCandidates = false;
                continue;
            }
            if (arg.Equals("--recovery-memory", StringComparison.OrdinalIgnoreCase))
            {
                Options.RecoveryMemoryEnabled = true;
                continue;
            }
            if (arg.Equals("--no-recovery-memory", StringComparison.OrdinalIgnoreCase))
            {
                Options.RecoveryMemoryEnabled = false;
                continue;
            }
            if (arg.Equals("--memory-list", StringComparison.OrdinalIgnoreCase))
            {
                Options.RecoveryMemoryCommand = "list";
                continue;
            }
            if (arg.Equals("--memory-prune", StringComparison.OrdinalIgnoreCase))
            {
                Options.RecoveryMemoryCommand = "prune";
                continue;
            }
            if (arg.Equals("--recovery-progress-verification", StringComparison.OrdinalIgnoreCase))
            {
                Options.RecoveryProgressVerificationEnabled = true;
                continue;
            }
            if (arg.Equals("--no-recovery-progress-verification", StringComparison.OrdinalIgnoreCase))
            {
                Options.RecoveryProgressVerificationEnabled = false;
                continue;
            }
            if (arg.Equals("--loop-replay-auto-export", StringComparison.OrdinalIgnoreCase))
            {
                Options.LoopReplayAutoExportEnabled = true;
                continue;
            }
            if (arg.Equals("--no-loop-replay-auto-export", StringComparison.OrdinalIgnoreCase))
            {
                Options.LoopReplayAutoExportEnabled = false;
                continue;
            }
            if (arg.Equals("--refresh-before-verify", StringComparison.OrdinalIgnoreCase))
            {
                Options.RefreshScreenshotBeforeVerify = true;
                continue;
            }
            if (arg.Equals("--no-refresh-before-verify", StringComparison.OrdinalIgnoreCase))
            {
                Options.RefreshScreenshotBeforeVerify = false;
                continue;
            }
            if (arg.Equals("--adaptive-effort", StringComparison.OrdinalIgnoreCase))
            {
                Options.AdaptiveReasoningEffort = true;
                continue;
            }
            if (arg.Equals("--no-adaptive-effort", StringComparison.OrdinalIgnoreCase))
            {
                Options.AdaptiveReasoningEffort = false;
                continue;
            }
            if (arg.Equals("--screen-polling", StringComparison.OrdinalIgnoreCase))
            {
                Options.ScreenPollingEnabled = true;
                continue;
            }
            if (arg.Equals("--no-screen-polling", StringComparison.OrdinalIgnoreCase))
            {
                Options.ScreenPollingEnabled = false;
                continue;
            }
            if (arg.Equals("--screen-sanity", StringComparison.OrdinalIgnoreCase))
            {
                Options.ScreenSanityChecks = true;
                continue;
            }
            if (arg.Equals("--no-screen-sanity", StringComparison.OrdinalIgnoreCase))
            {
                Options.ScreenSanityChecks = false;
                continue;
            }
            if (arg.Equals("--prompt-cache", StringComparison.OrdinalIgnoreCase))
            {
                Options.UsePromptCache = true;
                continue;
            }
            if (arg.Equals("--no-prompt-cache", StringComparison.OrdinalIgnoreCase))
            {
                Options.UsePromptCache = false;
                continue;
            }
            if (arg.Equals("--previous-response-state", StringComparison.OrdinalIgnoreCase))
            {
                Options.UsePreviousResponseState = true;
                continue;
            }
            if (arg.Equals("--no-previous-response-state", StringComparison.OrdinalIgnoreCase))
            {
                Options.UsePreviousResponseState = false;
                continue;
            }
            if (arg.Equals("--context-compaction", StringComparison.OrdinalIgnoreCase))
            {
                Options.ControlContextCompactionEnabled = true;
                continue;
            }
            if (arg.Equals("--no-context-compaction", StringComparison.OrdinalIgnoreCase))
            {
                Options.ControlContextCompactionEnabled = false;
                continue;
            }
            if (arg.Equals("--omit-unchanged-screen", StringComparison.OrdinalIgnoreCase))
            {
                Options.OmitUnchangedScreenImageWithState = true;
                continue;
            }
            if (arg.Equals("--no-omit-unchanged-screen", StringComparison.OrdinalIgnoreCase))
            {
                Options.OmitUnchangedScreenImageWithState = false;
                continue;
            }
            if (arg.Equals("--reuse-uia-targets", StringComparison.OrdinalIgnoreCase))
            {
                Options.ReuseUiaTargetsWhenScreenUnchanged = true;
                continue;
            }
            if (arg.Equals("--no-reuse-uia-targets", StringComparison.OrdinalIgnoreCase))
            {
                Options.ReuseUiaTargetsWhenScreenUnchanged = false;
                continue;
            }
            if (argument.TryRead("--profile", out var profile))
            {
                continue;
            }
            if (argument.TryRead("--model", out var model))
            {
                if (!string.IsNullOrWhiteSpace(model))
                    Options.Model = model.Trim();
                continue;
            }
            if (argument.TryRead("--qa-model", out var qaModel))
            {
                if (!string.IsNullOrWhiteSpace(qaModel))
                    Options.QaModel = qaModel.Trim();
                continue;
            }
            if (argument.TryRead("--verify-model", out var verifyModel))
            {
                if (!string.IsNullOrWhiteSpace(verifyModel))
                    Options.VerifyModel = verifyModel.Trim();
                continue;
            }
            if (argument.TryRead("--effort", out var effort) ||
                argument.TryRead("--reasoning-effort", out effort))
            {
                ApplyReasoningEffort(effort, "--effort");
                continue;
            }
            if (argument.TryRead("--qa-effort", out var qaEffort))
            {
                ApplyReasoningEffort(qaEffort, "--qa-effort", v => Options.QaReasoningEffort = v, () => Options.QaReasoningEffortExplicit = true);
                continue;
            }
            if (argument.TryRead("--verify-effort", out var verifyEffort))
            {
                ApplyReasoningEffort(verifyEffort, "--verify-effort", v => Options.VerifyReasoningEffort = v, () => Options.VerifyReasoningEffortExplicit = true);
                continue;
            }
            if (argument.TryRead("--delay", out var delay))
            {
                ApplyDelayMs(delay, "--delay");
                continue;
            }
            if (argument.TryRead("--grid", out var grid))
            {
                ApplyGridStep(grid, "--grid");
                continue;
            }
            if (argument.TryRead("--max-steps", out var maxSteps))
            {
                ApplyInt(maxSteps, "--max-steps", 0, int.MaxValue, v => Options.MaxSteps = v);
                continue;
            }
            if (argument.TryRead("--goal-mode", out var goalMode))
            {
                ApplyGoalMode(goalMode, "--goal-mode");
                continue;
            }
            if (argument.TryRead("--observation-profile", out var observationProfile))
            {
                ApplyObservationMode(observationProfile, "--observation-profile");
                continue;
            }
            if (arg.Equals("--observation-log-verbose", StringComparison.OrdinalIgnoreCase))
            {
                Options.ObservationLogVerbose = true;
                continue;
            }
            if (arg.Equals("--no-observation-log-verbose", StringComparison.OrdinalIgnoreCase))
            {
                Options.ObservationLogVerbose = false;
                continue;
            }
            if (argument.TryRead("--max-wait", out var maxWait))
            {
                ApplyInt(maxWait, "--max-wait", 0, int.MaxValue, v => Options.MaxWaitSeconds = v);
                continue;
            }
            if (argument.TryRead("--max-output-tokens", out var maxOutput))
            {
                ApplyInt(maxOutput, "--max-output-tokens", 1, int.MaxValue, v => Options.MaxOutputTokens = v);
                continue;
            }
            if (argument.TryRead("--qa-max-output-tokens", out var qaMaxOutput))
            {
                ApplyInt(qaMaxOutput, "--qa-max-output-tokens", 1, int.MaxValue, v => Options.QaMaxOutputTokens = v);
                continue;
            }
            if (argument.TryRead("--verify-max-output-tokens", out var verifyMaxOutput))
            {
                ApplyInt(verifyMaxOutput, "--verify-max-output-tokens", 1, int.MaxValue, v => Options.VerifyMaxOutputTokens = v);
                continue;
            }
            if (argument.TryRead("--turn-reanalysis-max-output-tokens", out var turnReanalysisMaxOutput))
            {
                ApplyInt(turnReanalysisMaxOutput, "--turn-reanalysis-max-output-tokens", 1, int.MaxValue, v => Options.TurnReanalysisMaxOutputTokens = v);
                continue;
            }
            if (argument.TryRead("--incomplete-max-output-retries", out var incompleteRetries))
            {
                ApplyInt(incompleteRetries, "--incomplete-max-output-retries", 0, 5, v => Options.IncompleteMaxOutputRetries = v);
                continue;
            }
            if (argument.TryRead("--incomplete-max-output-token-cap", out var incompleteTokenCap))
            {
                ApplyInt(incompleteTokenCap, "--incomplete-max-output-token-cap", 1, int.MaxValue, v => Options.IncompleteMaxOutputTokenCap = v);
                continue;
            }
            if (argument.TryRead("--max-action-text-chars", out var maxActionTextChars))
            {
                ApplyInt(maxActionTextChars, "--max-action-text-chars", 256, int.MaxValue, v => Options.MaxActionTextChars = v);
                continue;
            }
            if (argument.TryRead("--qa-screenshot-max-width", out var qaScreenshotMaxWidth))
            {
                ApplyInt(qaScreenshotMaxWidth, "--qa-screenshot-max-width", 0, 10000, v => Options.QaScreenshotMaxWidth = v);
                continue;
            }
            if (argument.TryRead("--verify-screenshot-max-width", out var verifyScreenshotMaxWidth))
            {
                ApplyInt(verifyScreenshotMaxWidth, "--verify-screenshot-max-width", 0, 10000, v => Options.VerifyScreenshotMaxWidth = v);
                continue;
            }
            if (argument.TryRead("--verbosity", out var verbosity))
            {
                ApplyTextVerbosity(verbosity, "--verbosity");
                continue;
            }
            if (argument.TryRead("--history-chars", out var historyChars))
            {
                ApplyInt(historyChars, "--history-chars", 0, int.MaxValue, v => Options.HistoryTailChars = v);
                continue;
            }
            if (argument.TryRead("--history-lines", out var historyLines))
            {
                ApplyInt(historyLines, "--history-lines", 0, int.MaxValue, v => Options.HistoryTailLines = v);
                continue;
            }
            if (argument.TryRead("--max-stagnation", out var maxStagnation))
            {
                ApplyInt(maxStagnation, "--max-stagnation", 0, int.MaxValue, v => Options.MaxStagnationStepsBeforeAbort = v);
                continue;
            }
            if (argument.TryRead("--max-repeated-actions", out var maxRepeated))
            {
                ApplyInt(maxRepeated, "--max-repeated-actions", 0, int.MaxValue, v => Options.MaxRepeatedActionBeforeAbort = v);
                continue;
            }
            if (argument.TryRead("--repeat-cooldown", out var repeatCooldown))
            {
                ApplyInt(repeatCooldown, "--repeat-cooldown", 0, int.MaxValue, v => Options.ActionRepeatCooldownSteps = v);
                continue;
            }
            if (argument.TryRead("--max-rejected-proposals", out var maxRejectedProposals))
            {
                ApplyInt(maxRejectedProposals, "--max-rejected-proposals", 0, int.MaxValue, v => Options.MaxRejectedProposalRepeatsBeforeAbort = v);
                continue;
            }
            if (argument.TryRead("--max-inspection-actions", out var maxInspectionActions))
            {
                ApplyInt(maxInspectionActions, "--max-inspection-actions", 0, 20, v => Options.MaxConsecutiveInspectionActions = v);
                continue;
            }
            if (argument.TryRead("--skip-verify-confidence", out var skipVerifyConfidence))
            {
                ApplyDouble(skipVerifyConfidence, "--skip-verify-confidence", 0.0, 1.0, v => Options.SkipVerifyConfidenceThreshold = v);
                continue;
            }
            if (argument.TryRead("--max-model-failures", out var maxModelFailures))
            {
                ApplyInt(maxModelFailures, "--max-model-failures", 0, int.MaxValue, v => Options.MaxModelFailuresBeforeAbort = v);
                continue;
            }
            if (argument.TryRead("--max-action-failures", out var maxActionFailures))
            {
                ApplyInt(maxActionFailures, "--max-action-failures", 0, int.MaxValue, v => Options.MaxActionFailuresBeforeAbort = v);
                continue;
            }
            if (argument.TryRead("--recovery-memory-path", out var recoveryMemoryPath))
            {
                if (!string.IsNullOrWhiteSpace(recoveryMemoryPath))
                    Options.RecoveryMemoryPath = recoveryMemoryPath.Trim();
                continue;
            }
            if (argument.TryRead("--memory-export", out var memoryExportPath))
            {
                Options.RecoveryMemoryCommand = "export";
                Options.RecoveryMemoryExportPath = memoryExportPath;
                continue;
            }
            if (argument.TryRead("--recovery-trigger", out var recoveryTrigger))
            {
                ApplyInt(recoveryTrigger, "--recovery-trigger", 1, int.MaxValue, v => Options.RecoveryMemoryTriggerSteps = v);
                continue;
            }
            if (argument.TryRead("--recovery-validation", out var recoveryValidation))
            {
                ApplyInt(recoveryValidation, "--recovery-validation", 1, int.MaxValue, v => Options.RecoveryMemoryValidationSteps = v);
                continue;
            }
            if (argument.TryRead("--recovery-max-lessons", out var recoveryMaxLessons))
            {
                ApplyInt(recoveryMaxLessons, "--recovery-max-lessons", 1, int.MaxValue, v => Options.RecoveryMemoryMaxLessons = v);
                continue;
            }
            if (argument.TryRead("--recovery-max-quarantined", out var recoveryMaxQuarantined))
            {
                ApplyInt(recoveryMaxQuarantined, "--recovery-max-quarantined", 1, int.MaxValue, v => Options.RecoveryMemoryMaxQuarantinedLessons = v);
                continue;
            }
            if (argument.TryRead("--recovery-reserved-per-context", out var recoveryReservedPerContext))
            {
                ApplyInt(recoveryReservedPerContext, "--recovery-reserved-per-context", 0, int.MaxValue, v => Options.RecoveryMemoryReservedLessonsPerContext = v);
                continue;
            }
            if (argument.TryRead("--recovery-soft-max-per-context", out var recoverySoftMaxPerContext))
            {
                ApplyInt(recoverySoftMaxPerContext, "--recovery-soft-max-per-context", 1, int.MaxValue, v => Options.RecoveryMemorySoftMaxLessonsPerContext = v);
                continue;
            }
            if (argument.TryRead("--recovery-max-file-bytes", out var recoveryMaxFileBytes))
            {
                ApplyLong(recoveryMaxFileBytes, "--recovery-max-file-bytes", 1024 * 1024, int.MaxValue, v => Options.RecoveryMemoryMaxFileBytes = v);
                continue;
            }
            if (argument.TryRead("--recovery-archive-max-bytes", out var recoveryArchiveMaxBytes))
            {
                ApplyLong(recoveryArchiveMaxBytes, "--recovery-archive-max-bytes", 1024 * 1024, int.MaxValue, v => Options.RecoveryMemoryArchiveMaxBytes = v);
                continue;
            }
            if (argument.TryRead("--recovery-archive-retained-files", out var recoveryArchiveRetainedFiles))
            {
                ApplyInt(recoveryArchiveRetainedFiles, "--recovery-archive-retained-files", 1, 20, v => Options.RecoveryMemoryArchiveRetainedFiles = v);
                continue;
            }
            if (argument.TryRead("--recovery-archive-path", out var recoveryArchivePath))
            {
                Options.RecoveryMemoryArchivePath = recoveryArchivePath;
                continue;
            }
            if (argument.TryRead("--recovery-prompt-lessons", out var recoveryPromptLessons))
            {
                ApplyInt(recoveryPromptLessons, "--recovery-prompt-lessons", 0, 10, v => Options.RecoveryMemoryPromptMaxLessons = v);
                continue;
            }
            if (argument.TryRead("--recovery-failure-limit", out var recoveryFailureLimit))
            {
                ApplyInt(recoveryFailureLimit, "--recovery-failure-limit", 1, int.MaxValue, v => Options.RecoveryMemoryFailureLimit = v);
                continue;
            }
            if (argument.TryRead("--runtime-semantic-states", out var semanticStateLimit))
            {
                ApplyInt(semanticStateLimit, "--runtime-semantic-states", 32, 100000, v => Options.RuntimeSemanticStateLimit = v);
                continue;
            }
            if (argument.TryRead("--runtime-graph-edges", out var graphEdgeLimit))
            {
                ApplyInt(graphEdgeLimit, "--runtime-graph-edges", 32, 100000, v => Options.RuntimeGraphEdgeLimit = v);
                continue;
            }
            if (argument.TryRead("--runtime-recovery-actions", out var recoveryActionLimit))
            {
                ApplyInt(recoveryActionLimit, "--runtime-recovery-actions", 8, 10000, v => Options.RuntimeRecoveryActionLimit = v);
                continue;
            }
            if (argument.TryRead("--runtime-cooldowns", out var cooldownLimit))
            {
                ApplyInt(cooldownLimit, "--runtime-cooldowns", 16, 100000, v => Options.RuntimeCooldownEntryLimit = v);
                continue;
            }
            if (argument.TryRead("--graph-candidate-ttl", out var graphCandidateTtl))
            {
                ApplyInt(graphCandidateTtl, "--graph-candidate-ttl", 2, 10000, v => Options.GraphCandidateTtlSteps = v);
                continue;
            }
            if (argument.TryRead("--loop-confidence-threshold", out var loopConfidenceThreshold))
            {
                ApplyDouble(loopConfidenceThreshold, "--loop-confidence-threshold", 0.5, 1.0, v => Options.ProactiveLoopConfidenceThreshold = v);
                continue;
            }
            if (argument.TryRead("--recovery-progress-confidence", out var progressConfidenceThreshold))
            {
                ApplyDouble(progressConfidenceThreshold, "--recovery-progress-confidence", 0.5, 1.0, v => Options.RecoveryProgressConfidenceThreshold = v);
                continue;
            }
            if (argument.TryRead("--recovery-telemetry-max-bytes", out var telemetryMaxBytes))
            {
                ApplyInt(telemetryMaxBytes, "--recovery-telemetry-max-bytes", 65536, int.MaxValue, v => Options.RecoveryTelemetryMaxBytes = v);
                continue;
            }
            if (argument.TryRead("--recovery-telemetry-retained-files", out var telemetryRetainedFiles))
            {
                ApplyInt(telemetryRetainedFiles, "--recovery-telemetry-retained-files", 1, 20, v => Options.RecoveryTelemetryRetainedFiles = v);
                continue;
            }
            if (argument.TryRead("--loop-replay", out var loopReplayPath))
            {
                Options.LoopReplayPath = loopReplayPath;
                continue;
            }
            if (argument.TryRead("--loop-replay-import", out var loopReplayImportPath))
            {
                Options.LoopReplayImportPath = loopReplayImportPath;
                continue;
            }
            if (argument.TryRead("--loop-replay-export", out var loopReplayExportPath))
            {
                Options.LoopReplayExportPath = loopReplayExportPath;
                continue;
            }
            if (argument.TryRead("--loop-replay-corpus", out var loopReplayCorpusPath))
            {
                Options.LoopReplayCorpusPath = loopReplayCorpusPath;
                continue;
            }
            if (argument.TryRead("--screen-poll-initial-delay", out var pollInitialDelay))
            {
                ApplyInt(pollInitialDelay, "--screen-poll-initial-delay", 0, 10000, v => Options.ScreenPollInitialDelayMs = v);
                continue;
            }
            if (argument.TryRead("--screen-poll-interval", out var pollInterval))
            {
                ApplyInt(pollInterval, "--screen-poll-interval", 10, 10000, v => Options.ScreenPollIntervalMs = v);
                continue;
            }
            if (argument.TryRead("--screen-poll-timeout", out var pollTimeout))
            {
                ApplyInt(pollTimeout, "--screen-poll-timeout", 0, 60000, v => Options.ScreenPollTimeoutMs = v);
                continue;
            }
            if (argument.TryRead("--wait-no-change-extra", out var waitNoChangeExtra))
            {
                ApplyInt(waitNoChangeExtra, "--wait-no-change-extra", 0, 60000, v => Options.WaitNoChangeExtraMs = v);
                continue;
            }
            if (argument.TryRead("--sendinput-retries", out var sendInputRetries))
            {
                ApplyInt(sendInputRetries, "--sendinput-retries", 0, 10, v => Options.SendInputMaxRetries = v);
                continue;
            }
            if (argument.TryRead("--sendinput-retry-delay", out var sendInputRetryDelay))
            {
                ApplyInt(sendInputRetryDelay, "--sendinput-retry-delay", 0, 1000, v => Options.SendInputRetryDelayMs = v);
                continue;
            }
            if (argument.TryRead("--screenshot-max-width", out var maxWidth))
            {
                ApplyInt(maxWidth, "--screenshot-max-width", 0, 10000, v => Options.MaxScreenshotSendWidth = v);
                continue;
            }
            if (argument.TryRead("--screenshot-format", out var imageFormat))
            {
                ApplyImageFormat(imageFormat, "--screenshot-format");
                continue;
            }
            if (argument.TryRead("--focused-overview-max-width", out var focusedOverviewMaxWidth))
            {
                ApplyInt(focusedOverviewMaxWidth, "--focused-overview-max-width", 0, 10000, v => Options.FocusedOverviewMaxWidth = v);
                continue;
            }
            if (argument.TryRead("--crop-max-width", out var cropMaxWidth))
            {
                ApplyInt(cropMaxWidth, "--crop-max-width", 0, 10000, v => Options.MaxCropSendWidth = v);
                continue;
            }
            if (argument.TryRead("--crop-format", out var cropFormat))
            {
                ApplyCropFormat(cropFormat, "--crop-format");
                continue;
            }
            if (argument.TryRead("--jpeg-quality", out var jpegQuality))
            {
                ApplyLong(jpegQuality, "--jpeg-quality", 1, 100, v => Options.ScreenshotJpegQuality = v);
                continue;
            }
            if (argument.TryRead("--screen-log-format", out var screenLogFormat))
            {
                ApplyScreenLogFormat(screenLogFormat, "--screen-log-format");
                continue;
            }
            if (argument.TryRead("--screen-log-max-width", out var screenLogMaxWidth))
            {
                ApplyInt(screenLogMaxWidth, "--screen-log-max-width", 0, 10000, v => Options.MaxScreenLogWidth = v);
                continue;
            }
            if (argument.TryRead("--verify", out var verifyMode))
            {
                ApplyVerifyMode(verifyMode, "--verify");
                continue;
            }
            if (argument.TryRead("--verify-early-steps", out var verifyEarlySteps))
            {
                ApplyInt(verifyEarlySteps, "--verify-early-steps", 0, int.MaxValue, v => Options.VerifyEarlySteps = v);
                continue;
            }
            if (argument.TryRead("--verify-low-confidence", out var verifyLowConfidence))
            {
                ApplyDouble(verifyLowConfidence, "--verify-low-confidence", 0.0, 1.0, v => Options.VerifyLowConfidenceThreshold = v);
                continue;
            }
            if (argument.TryRead("--paste-threshold", out var pasteThreshold))
            {
                ApplyInt(pasteThreshold, "--paste-threshold", 0, int.MaxValue, v => Options.ClipboardPasteThreshold = v);
                continue;
            }
            if (argument.TryRead("--focus-crop-size", out var focusCropSize))
            {
                ApplyInt(focusCropSize, "--focus-crop-size", 64, 2000, v => Options.FocusCropSize = v);
                continue;
            }
            if (argument.TryRead("--max-focus-crop-pixels", out var maxFocusPixels))
            {
                ApplyInt(maxFocusPixels, "--max-focus-crop-pixels", 0, int.MaxValue, v => Options.MaxFocusUiaCropPixels = v);
                continue;
            }
            if (argument.TryRead("--max-uia-targets", out var maxUiaTargets))
            {
                ApplyInt(maxUiaTargets, "--max-uia-targets", 0, 100, v => Options.MaxUiaTargets = v);
                Options.IncludeUiaTargets = Options.MaxUiaTargets > 0;
                continue;
            }
            if (argument.TryRead("--uia-name-chars", out var uiaNameChars))
            {
                ApplyInt(uiaNameChars, "--uia-name-chars", 0, 500, v => Options.UiaTargetNameMaxChars = v);
                continue;
            }
            if (argument.TryRead("--uia-summary-chars", out var uiaSummaryChars))
            {
                ApplyInt(uiaSummaryChars, "--uia-summary-chars", 0, 2000, v => Options.UiaSummaryMaxChars = v);
                continue;
            }
            if (argument.TryRead("--uia-scan-ms", out var uiaScanMs))
            {
                ApplyInt(uiaScanMs, "--uia-scan-ms", 0, 5000, v => Options.UiaScanTimeBudgetMs = v);
                continue;
            }
            if (argument.TryRead("--max-uia-nodes", out var maxUiaNodes))
            {
                ApplyInt(maxUiaNodes, "--max-uia-nodes", 0, 10000, v => Options.MaxUiaNodesScanned = v);
                continue;
            }
            if (argument.TryRead("--uia-candidate-multiplier", out var uiaCandidateMultiplier))
            {
                ApplyInt(uiaCandidateMultiplier, "--uia-candidate-multiplier", 1, 20, v => Options.UiaCandidateMultiplier = v);
                continue;
            }
            if (argument.TryRead("--uia-max-area-ratio", out var uiaMaxAreaRatio))
            {
                ApplyDouble(uiaMaxAreaRatio, "--uia-max-area-ratio", 0.0, 1.0, v => Options.MaxUiaTargetAreaRatio = v);
                continue;
            }
            if (argument.TryRead("--max-artifacts", out var maxArtifacts))
            {
                ApplyInt(maxArtifacts, "--max-artifacts", 0, int.MaxValue, v => Options.MaxArtifactsPerDir = v);
                continue;
            }
            if (argument.TryRead("--max-batch-actions", out var maxBatchActions))
            {
                ApplyInt(maxBatchActions, "--max-batch-actions", 0, 20, v => Options.MaxQueuedBatchActions = v);
                continue;
            }
            if (argument.TryRead("--turn-batch-inputs", out var turnBatchInputs))
            {
                ApplyInt(turnBatchInputs, "--turn-batch-inputs", 2, 64, v => Options.TurnBasedMaxBatchInputs = v);
                continue;
            }
            if (argument.TryRead("--replay-response", out var replayPath))
            {
                Options.ReplayResponsePath = replayPath;
                continue;
            }
            if (argument.TryRead("--replay-request", out var replayRequestPath))
            {
                Options.ReplayRequestPath = replayRequestPath;
                continue;
            }
            if (arg.Equals("--replay-request-dry-run", StringComparison.OrdinalIgnoreCase))
            {
                Options.ReplayRequestDryRun = true;
                continue;
            }
            if (argument.TryRead("--retries", out var retries))
            {
                ApplyInt(retries, "--retries", 0, 10, v => Options.OpenAiMaxRetries = v);
                continue;
            }
            if (argument.TryRead("--openai-timeout", out var openAiTimeout))
            {
                ApplyInt(openAiTimeout, "--openai-timeout", 0, int.MaxValue, v => Options.OpenAiTimeoutSeconds = v);
                continue;
            }
            if (argument.TryRead("--prompt-cache-key", out var promptCacheKey))
            {
                Options.PromptCacheKey = string.IsNullOrWhiteSpace(promptCacheKey) ? null : promptCacheKey.Trim();
                continue;
            }
            if (argument.TryRead("--control-reasoning-context", out var reasoningContext))
            {
                ApplyReasoningContext(reasoningContext, "--control-reasoning-context");
                continue;
            }
            if (argument.TryRead("--context-compact-threshold", out var compactThreshold))
            {
                ApplyInt(compactThreshold, "--context-compact-threshold", 1, int.MaxValue, v => Options.ControlContextCompactThreshold = v);
                continue;
            }
            if (argument.TryRead("--context-fallback-limit", out var fallbackLimit))
            {
                ApplyInt(fallbackLimit, "--context-fallback-limit", 1, 20, v => Options.ControlContextFallbackLimit = v);
                continue;
            }

            if (arg.StartsWith("--", StringComparison.Ordinal))
            {
                ReportConfigurationError($"Unknown option: {arg}");
                continue;
            }

            positional.Add(arg);
        }

        if (explicitTask is not null && positional.Count > 0)
        {
            ReportConfigurationError("Positional task text cannot be combined with --task, --goal, or --task-file.");
        }

        return new CliValidationResult(
            explicitTask ?? (positional.Count == 0 ? null : string.Join(" ", positional)),
            CurrentRun.ConfigurationErrors.Skip(errorStart).ToArray());
    }

    private static void ApplyCliProfileArgs(IReadOnlyList<CliArgument> arguments)
    {
        string? selected = null;
        foreach (var argument in arguments)
        {
            if (argument.TryRead("--profile", out var value))
            {
                if (value.Trim().ToLowerInvariant() is not ("custom" or "fast" or "balanced" or "quality"))
                    ReportConfigurationError($"Unknown profile '{value}'. Allowed: custom, fast, balanced, quality.");
                else
                    selected = value;
            }
            else if (argument.Name.ToLowerInvariant() is "--fast" or "--balanced" or "--quality")
                selected = argument.Name[2..];
        }
        if (selected is not null) ApplyProfile(selected);
    }
}
