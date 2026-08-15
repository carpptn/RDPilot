/// <summary>
/// Coordinates the interactive desktop-agent application and its shared runtime state.
/// </summary>
internal static partial class RDPilotApplication
{
    // === CONFIG ===
    const string DefaultModel = "gpt-5.6-luna";                 // e.g., gpt-4o-mini, gpt-4o, gpt-5
    static string Model = DefaultModel;
    static string? ReasoningEffort = "max";                  // null = API/model default; e.g., low, medium, high, xhigh, max
    static bool ReasoningEffortExplicit = false;
    static string? QaReasoningEffort = null;
    static bool QaReasoningEffortExplicit = false;
    static string? VerifyReasoningEffort = "max";
    static bool VerifyReasoningEffortExplicit = true;
    static string? QaModel = null;
    static string? VerifyModel = null;
    static string ApiUrl = "https://api.openai.com/v1/responses";
    
    const int MaxStepsDefault = 10000;
    static int MaxSteps = MaxStepsDefault;
    static bool MultiMonitorEnabled = false;

    // Mouse: enable/disable (default enabled for GPT-5.6 vision/control quality)
    static bool MouseEnabled = true;

    // Global, configurable time to let UI "settle" after an action before taking the next screenshot
    static int UiSettleDelayMs = 300;                   // can override via POST_ACTION_DELAY_MS or --delay <ms>

    // Include a small crop of the current UIA focus area as an extra image?
    static bool IncludeFocusUia = true;                 // focused UIA rect/summary/overlay
    static bool IncludeFocusUiaCrop = false;
    static bool SendFocusCrop = true;                   // attach crop from AIM/request_crop
    static int FocusCropSize = 320;                     // px

    // Focus overlay (look)
    static int FocusRingPadding = 3;
    static int FocusRingThickness = 2;
    static int FocusGlowThickness = 3;
    static int FocusCornerRadius = 6;
    static double LargeFocusOverlayAreaRatio = 0.22;
    static double ClickAimEdgeAdjustMaxAreaRatio = 0.22;
    static double ClickAimEdgeMarginRatio = 0.18;
    static int ClickAimEdgeMinMarginPx = 8;
    static int ClickAimEdgeMaxMarginPx = 32;

    // === Grid overlay on screenshots ===
    static int GridStepPx = 0;          // 0 = off; e.g., 100 = lines every 100 px
    static int GridLabelEveryPx = 100;  // label every N px
    static int GridMajorEveryPx = 100;  // thicker line every N px

    // AIM expiration after a large visual change
    static double AimExpireDelta = 0.08;

    // Stagnation / verification
    static double NoChangeThreshold = 0.005;            // 0..1 (avg pixel diff after downsampling)
    static string ObservationMode = "auto";             // auto | general | static_ui | local_editing | event_driven | streaming_output | realtime_interaction
    static bool ObservationLogVerbose = false;
    static int MaxGesturePathPoints = 128;
    static int MaxGestureDurationMs = 5_000;
    static int MaxHeldKeys = 4;
    static int MaxKeyHoldDurationMs = 5_000;

    // Speed/quality profile
    static string RunProfile = "custom";                // custom | fast | balanced | quality
    static int MaxOutputTokens = 10000;
    static int QaMaxOutputTokens = 4000;
    static int VerifyMaxOutputTokens = 6000;
    static int TurnReanalysisMaxOutputTokens = 10000;
    static int IncompleteMaxOutputRetries = 2;
    static int IncompleteMaxOutputTokenCap = 16000;
    static int MaxActionTextChars = 3000;
    static int QaScreenshotMaxWidth = 1024;
    static int VerifyScreenshotMaxWidth = 1024;
    static string TextVerbosity = "low";
    static int HistoryTailChars = 1200;
    static int HistoryTailLines = 12;
    static int ActionRepeatCooldownSteps = 2;
    static int MaxRejectedProposalRepeatsBeforeAbort = 5;
    static int IneffectiveMouseClusterPx = 96;
    static double SkipVerifyConfidenceThreshold = 0.92;
    static int MaxStagnationStepsBeforeAbort = 20;
    static int MaxRepeatedActionBeforeAbort = 5;
    static int MaxModelFailuresBeforeAbort = 2;
    static int MaxActionFailuresBeforeAbort = 2;
    static string GoalMode = "auto";
    static bool RecoveryMemoryEnabled = true;
    static string? RecoveryMemoryPath = null;
    static int RecoveryMemoryTriggerSteps = 2;
    static int RecoveryMemoryValidationSteps = 2;
    static int RecoveryMemoryMaxLessons = 500;
    static int RecoveryMemoryMaxQuarantinedLessons = 500;
    static int RecoveryMemoryReservedLessonsPerContext = 5;
    static int RecoveryMemorySoftMaxLessonsPerContext = 100;
    static long RecoveryMemoryMaxFileBytes = 32L * 1024 * 1024;
    static string? RecoveryMemoryArchivePath = null;
    static long RecoveryMemoryArchiveMaxBytes = 32L * 1024 * 1024;
    static int RecoveryMemoryArchiveRetainedFiles = 3;
    static int RecoveryMemoryPromptMaxLessons = 2;
    static int RecoveryMemoryFailureLimit = 3;
    static int RuntimeSemanticStateLimit = 256;
    static int RuntimeGraphEdgeLimit = 512;
    static int RuntimeRecoveryActionLimit = 64;
    static int RuntimeCooldownEntryLimit = 256;
    static int GraphCandidateTtlSteps = 24;
    static double ProactiveLoopConfidenceThreshold = 0.75;
    static bool RecoveryProgressVerificationEnabled = true;
    static double RecoveryProgressConfidenceThreshold = 0.68;
    static int RecoveryTelemetryMaxBytes = 5 * 1024 * 1024;
    static int RecoveryTelemetryRetainedFiles = 3;
    static string? RecoveryMemoryCommand = null;
    static string? RecoveryMemoryExportPath = null;
    static string? LoopReplayPath = null;
    static string? LoopReplayImportPath = null;
    static string? LoopReplayExportPath = null;
    static bool LoopReplayAutoExportEnabled = true;
    static string? LoopReplayCorpusPath = null;
    static int MaxWaitSeconds = 30;
    static bool ScreenPollingEnabled = true;
    static int ScreenPollInitialDelayMs = 120;
    static int ScreenPollIntervalMs = 150;
    static int ScreenPollTimeoutMs = 1200;
    static int WaitNoChangeExtraMs = 750;
    static bool ScreenSanityChecks = true;
    static int SendInputMaxRetries = 2;
    static int SendInputRetryDelayMs = 30;
    static bool AdaptiveReasoningEffort = true;
    static int MaxScreenshotSendWidth = 1280;           // 0 = original
    static string ScreenshotSendFormat = "jpeg";        // jpeg | png
    static int FocusedOverviewMaxWidth = 640;           // 0 = use normal screenshot width when a focus crop is present
    static int MaxCropSendWidth = 768;                  // 0 = original
    static string CropSendFormat = "jpeg";              // jpeg | png
    static long ScreenshotJpegQuality = 80;
    static string ScreenLogFormat = "jpeg";             // jpeg | png
    static int MaxScreenLogWidth = 1280;                // 0 = original
    static bool DebugImages = false;
    static bool LogRequests = true;
    static bool PrettyRequestLogs = false;
    static bool LogScreens = true;
    static string VerifyMode = "auto";                  // auto | always | off
    static int VerifyEarlySteps = 2;
    static double VerifyLowConfidenceThreshold = 0.75;
    static bool RefreshScreenshotBeforeVerify = false;
    static int ClipboardPasteThreshold = 120;
    static int MaxFocusUiaCropPixels = 160_000;
    static int OpenAiMaxRetries = 2;
    static int OpenAiTimeoutSeconds = 600;
    static bool ForceRealUiOnly = false;
    static bool AllowHighLevelActions = false;
    static bool AllowRunCommand = false;
    static bool AllowWebSearch = false;
    static bool BatchMode = false;
    static bool CliArgumentError = false;
    static bool DirectClickWithoutAim = true;
    static bool AutoHideConsoleDuringRun = true;
    static bool MinimizeConsoleDuringRun = false;
    static bool RestoreConsoleAfterRun = true;
    static string? RequestReasoningEffortOverride = null;
    static bool UsePromptCache = true;
    static string? PromptCacheKey = "rdpilot-control-v1";
    static bool UsePreviousResponseState = true;
    static string ControlReasoningContext = "all_turns";
    static bool ControlContextCompactionEnabled = true;
    static int ControlContextCompactThreshold = 700_000;
    static int ControlContextFallbackLimit = 3;
    static bool OmitUnchangedScreenImageWithState = false;
    static bool IncludeUiaTargets = true;
    static int MaxUiaTargets = 20;
    static int UiaTargetNameMaxChars = 48;
    static int UiaSummaryMaxChars = 320;
    static int UiaScanTimeBudgetMs = 60;
    static int MaxUiaNodesScanned = 400;
    static int UiaCandidateMultiplier = 4;
    static double MaxUiaTargetAreaRatio = 0.45;
    static bool ReuseUiaTargetsWhenScreenUnchanged = true;
    static bool AnalyzeLogsOnly = false;
    static bool PrintConfigOnly = false;
    static int MaxArtifactsPerDir = 500;
    static List<UiaTarget> CurrentUiaTargets = [];
    static ScreenCoordinateMapper CurrentScreenMap = ScreenCoordinateMapper.Create(1, 1, 1, 1);
    static bool ExecuteMultiActionCandidates = true;
    static int MaxQueuedBatchActions = 4;
    static int TurnBasedMaxBatchInputs = 32;
    static int MaxBatchedGesturePoints = 180;
    static int MaxBatchedGestureDurationMs = 12_000;
    static int MaxConsecutiveInspectionActions = 2;
    static readonly Queue<ActionDto> PendingSafeActions = new();
    static string? ReplayResponsePath = null;
    static string? ReplayRequestPath = null;
    static bool ReplayRequestDryRun = false;
    static int RunOpenAiCalls = 0;
    static int RunOpenAiRetries = 0;
    static TimeSpan RunOpenAiElapsed = TimeSpan.Zero;
    static long RunOpenAiRequestBytes = 0;
    static long RunOpenAiBytes = 0;
    static int RunMultiCandidateResponses = 0;
    static int RunWebSearchCalls = 0;
    static long RunInputTokens = 0;
    static long RunCachedTokens = 0;
    static long RunOutputTokens = 0;
    static long RunReasoningTokens = 0;
    static int RunEarlyAcceptedControlStreams = 0;
    static int RunControlContextTurns = 0;
    static int RunControlContextRestarts = 0;
    static int RunControlContextFallbacks = 0;
    static int RunControlContextCompactions = 0;
    static int RunControlCompactionFallbacks = 0;
    static int RunScreenshotCount = 0;
    static TimeSpan RunScreenshotElapsed = TimeSpan.Zero;
    static int RunScreenProbeCount = 0;
    static TimeSpan RunScreenProbeElapsed = TimeSpan.Zero;
    static int RunScreenSanityWarnings = 0;
    static int RunImageEncodeCount = 0;
    static TimeSpan RunImageEncodeElapsed = TimeSpan.Zero;
    static int RunScreenLogCount = 0;
    static TimeSpan RunScreenLogElapsed = TimeSpan.Zero;
    static int RunArtifactLogWrites = 0;
    static TimeSpan RunArtifactLogElapsed = TimeSpan.Zero;
    static int RunUiaCalls = 0;
    static TimeSpan RunUiaElapsed = TimeSpan.Zero;
    static int RunLocalActions = 0;
    static TimeSpan RunLocalActionElapsed = TimeSpan.Zero;

    static readonly JsonSerializerOptions PrettyJson = new(JsonSerializerDefaults.Web) { WriteIndented = true };
    static readonly HttpClient OpenAiHttp = new();
    static bool LastOpenAiFailureWasRetriable = false;
    static string LastOpenAiFailureKind = "";
    static string? LastOpenAiResponseId = null;
    static readonly HashSet<string> AllowedReasoningEfforts = new(StringComparer.OrdinalIgnoreCase)
    {
        "none", "minimal", "low", "medium", "high", "xhigh", "max"
    };
    static readonly HashSet<string> KnownActionTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "open_url", "launch_app", "run_command",
        "paste_text", "focus_uia", "click_uia",
        "move", "click", "double_click", "drag_drop", "drag_path", "keys", "hold_keys", "type_text", "scroll",
        "request_crop", "point", "aim", "wait", "done"
    };
    static readonly HashSet<string> ReportedSanityWarnings = new(StringComparer.OrdinalIgnoreCase);
    static readonly HashSet<string> ReportedWebSearchCallIds = new(StringComparer.Ordinal);

    static async Task<int> Main(string[] args)
    {
        try
        {
            Console.OutputEncoding = Encoding.UTF8;
            ConsoleTheme.Enable();
            AppDomain.CurrentDomain.ProcessExit += (_, _) => ReleaseAllHeldKeys();

            ApplyEnvironmentConfig();
            string? pending = ApplyCliArgs(args);
            NormalizeConfig();
            ConfigureOpenAiHttpClient();

            if (BatchMode && CliArgumentError)
                return 2;

            if (PrintConfigOnly)
            {
                PrintEffectiveConfig();
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(RecoveryMemoryCommand))
            {
                ExecuteRecoveryMemoryMaintenance(
                    RecoveryMemoryCommand,
                    RecoveryMemoryExportPath);
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(LoopReplayPath))
            {
                ExecuteLoopReplay(LoopReplayPath);
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(LoopReplayImportPath))
            {
                ImportIndependentLoopReplayCorpus(LoopReplayImportPath);
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(LoopReplayExportPath))
            {
                ExportLoopTelemetryToReplayCorpus(LoopReplayExportPath);
                return 0;
            }
            if (AnalyzeLogsOnly)
            {
                AnalyzeArtifacts();
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(ReplayResponsePath))
            {
                ReplayResponse(ReplayResponsePath);
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(ReplayRequestPath) && ReplayRequestDryRun)
            {
                await ReplayRequestAsync(null, ReplayRequestPath, dryRun: true);
                return 0;
            }

            var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                if (BatchMode || Console.IsInputRedirected)
                {
                    Console.Error.WriteLine("Missing OPENAI_API_KEY.");
                    return 2;
                }

                Console.Write("Enter OPENAI_API_KEY: ");
                apiKey = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    Console.Error.WriteLine("Missing API key.");
                    return 2;
                }
            }

            if (!string.IsNullOrWhiteSpace(ReplayRequestPath))
            {
                await ReplayRequestAsync(apiKey!, ReplayRequestPath, dryRun: false);
                return 0;
            }

            try { SetProcessDpiAwarenessContext((nint)(-4)); } catch { /* best effort */ }

            PrintStartupSummary();
            if (BatchMode)
                return await RunBatchTaskAsync(apiKey!, pending);

            var promptHistory = PromptHistoryService.Load();
            Console.WriteLine(
                "Enter a task or question. Use Up/Down for prompt history, '/ask ' for Q&A, and '/exit' to quit.");
            Console.WriteLine(
                $"Prompt history: {PromptHistoryService.EffectivePath()} ({promptHistory.Count} entries)\n");

            while (true)
            {
                string goal;
                if (!string.IsNullOrEmpty(pending))
                {
                    goal = pending!;
                    pending = null;
                    Console.WriteLine($"Command (from args): {goal}");
                }
                else
                {
                    goal = PromptHistoryService.ReadLine(
                        "Command/Question: ",
                        promptHistory);
                }

                if (string.IsNullOrWhiteSpace(goal) || goal.Trim().Equals("/exit", StringComparison.OrdinalIgnoreCase))
                    break;

                PromptHistoryService.Remember(promptHistory, goal);

                if (IsQuestion(goal))
                {
                    await RunAskOnce(apiKey!, goal);
                    Console.WriteLine();
                    Console.WriteLine("✅ Answer completed. Enter next (ENTER = exit).");
                }
                else
                {
                    var result = await RunOnce(apiKey!, goal);
                    Console.WriteLine();
                    Console.WriteLine(
                        $"{ControlRunMarker(result.Outcome)} {result.Outcome}: {result.Message} " +
                        $"(step={result.Step}). Enter next (ENTER = exit).");
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"RDPilot failed: {ex.Message}");
            return 1;
        }
    }

    internal static async Task<int> RunBatchTaskAsync(string apiKey, string? task)
    {
        if (string.IsNullOrWhiteSpace(task) ||
            task.Trim().Equals("/exit", StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine("Batch mode requires a non-empty task. Use --task <text>, --task-file <path>, or --batch with a positional task.");
            return 2;
        }

        var goal = task.Trim();
        Console.WriteLine($"[batch] task={goal}");
        if (IsQuestion(goal))
        {
            var answered = await RunAskOnce(apiKey, goal);
            if (answered)
            {
                Console.WriteLine("[batch] completed");
                return 0;
            }

            Console.Error.WriteLine(CancelRequested
                ? "[batch] cancelled"
                : "[batch] failed: no answer was returned");
            return CancelRequested ? 130 : 1;
        }

        var result = await RunOnce(apiKey, goal);
        var summary = $"[batch] {result.Outcome}: {result.Message} (step={result.Step})";
        if (result.Completed)
            Console.WriteLine(summary);
        else
            Console.Error.WriteLine(summary);
        return ControlRunExitCode(result.Outcome);
    }

    internal static int ControlRunExitCode(ControlRunOutcome outcome) => outcome switch
    {
        ControlRunOutcome.Completed => 0,
        ControlRunOutcome.Cancelled => 130,
        ControlRunOutcome.GuardStopped => 3,
        ControlRunOutcome.StepLimitReached => 4,
        _ => 1
    };

    internal static string ControlRunMarker(ControlRunOutcome outcome) => outcome switch
    {
        ControlRunOutcome.Completed => "✅",
        ControlRunOutcome.Cancelled => "⏹",
        ControlRunOutcome.GuardStopped => "🛑",
        ControlRunOutcome.StepLimitReached => "⏱",
        _ => "❌"
    };
}
