/// <summary>
/// Coordinates the interactive desktop-agent application and its shared runtime state.
/// </summary>
internal static partial class RDPilotApplication
{
    // === CONFIG ===
    const string DefaultModel = "gpt-5.6-terra";                // e.g., gpt-4o-mini, gpt-4o, gpt-5
    static string Model = DefaultModel;
    static string? ReasoningEffort = "medium";               // null = API/model default; e.g., low, medium, high, xhigh
    static string? QaReasoningEffort = null;
    static bool QaReasoningEffortExplicit = false;
    static string? VerifyReasoningEffort = "medium";
    static bool VerifyReasoningEffortExplicit = true;
    static string? QaModel = null;
    static string? VerifyModel = null;
    const string ApiUrl = "https://api.openai.com/v1/responses";
    
    const int MaxStepsDefault = 10000;
    static int MaxSteps = MaxStepsDefault;
    static bool MultiMonitorEnabled = false;

    // Mouse: enable/disable (default enabled for GPT-5.6-terra vision/control quality)
    static bool MouseEnabled = true;

    // Global, configurable time to let UI "settle" after an action before taking the next screenshot
    static int UiSettleDelayMs = 300;                   // can override via POST_ACTION_DELAY_MS or --delay <ms>

    // Include a small crop of the current UIA focus area as an extra image?
    static bool IncludeFocusUia = true;                 // focused UIA rect/summary/overlay
    static bool IncludeFocusUiaCrop = false;
    static bool SendFocusCrop = true;                   // attach crop from AIM/request_crop
    static int FocusCropSize = 320;                     // px

    // Focus overlay (look)
    const int FocusRingPadding = 3;
    const int FocusRingThickness = 2;
    const int FocusGlowThickness = 3;
    const int FocusCornerRadius = 6;
    const double LargeFocusOverlayAreaRatio = 0.22;
    const double ClickAimEdgeAdjustMaxAreaRatio = 0.22;
    const double ClickAimEdgeMarginRatio = 0.18;
    const int ClickAimEdgeMinMarginPx = 8;
    const int ClickAimEdgeMaxMarginPx = 32;

    // === Grid overlay on screenshots ===
    static int GridStepPx = 0;          // 0 = off; e.g., 100 = lines every 100 px
    static int GridLabelEveryPx = 100;  // label every N px
    static int GridMajorEveryPx = 100;  // thicker line every N px

    // AIM expiration after a large visual change
    const double AimExpireDelta = 0.08;

    // Stagnation / verification
    const double NoChangeThreshold = 0.005;             // 0..1 (avg pixel diff after downsampling)

    // Speed/quality profile
    static string RunProfile = "custom";                // custom | fast | balanced | quality
    static int MaxOutputTokens = 600;
    static int QaMaxOutputTokens = 600;
    static int VerifyMaxOutputTokens = 120;
    static int IncompleteMaxOutputRetries = 2;
    static int IncompleteMaxOutputTokenCap = 6000;
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
    static bool DirectClickWithoutAim = true;
    static bool AutoHideConsoleDuringRun = true;
    static bool MinimizeConsoleDuringRun = false;
    static bool RestoreConsoleAfterRun = true;
    static string? RequestReasoningEffortOverride = null;
    static bool UsePromptCache = true;
    static string? PromptCacheKey = "rdpilot-control-v1";
    static bool UsePreviousResponseState = false;
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
    static bool ExecuteMultiActionCandidates = false;
    static int MaxQueuedBatchActions = 4;
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
    static long RunInputTokens = 0;
    static long RunCachedTokens = 0;
    static long RunOutputTokens = 0;
    static long RunReasoningTokens = 0;
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
        "none", "minimal", "low", "medium", "high", "xhigh"
    };
    static readonly HashSet<string> KnownActionTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "open_url", "launch_app", "run_command",
        "paste_text", "focus_uia", "click_uia",
        "move", "click", "double_click", "drag_drop", "keys", "type_text", "scroll",
        "request_crop", "point", "aim", "wait", "done"
    };
    static readonly HashSet<string> ReportedSanityWarnings = new(StringComparer.OrdinalIgnoreCase);

    static async Task Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        ConsoleTheme.Enable();

        ApplyEnvironmentConfig();
        string? pending = ApplyCliArgs(args);
        NormalizeConfig();
        ConfigureOpenAiHttpClient();

        if (PrintConfigOnly)
        {
            PrintEffectiveConfig();
            return;
        }
        if (!string.IsNullOrWhiteSpace(RecoveryMemoryCommand))
        {
            ExecuteRecoveryMemoryMaintenance(
                RecoveryMemoryCommand,
                RecoveryMemoryExportPath);
            return;
        }
        if (!string.IsNullOrWhiteSpace(LoopReplayPath))
        {
            ExecuteLoopReplay(LoopReplayPath);
            return;
        }
        if (!string.IsNullOrWhiteSpace(LoopReplayImportPath))
        {
            ImportIndependentLoopReplayCorpus(LoopReplayImportPath);
            return;
        }
        if (!string.IsNullOrWhiteSpace(LoopReplayExportPath))
        {
            ExportLoopTelemetryToReplayCorpus(LoopReplayExportPath);
            return;
        }
        if (AnalyzeLogsOnly)
        {
            AnalyzeArtifacts();
            return;
        }
        if (!string.IsNullOrWhiteSpace(ReplayResponsePath))
        {
            ReplayResponse(ReplayResponsePath);
            return;
        }
        if (!string.IsNullOrWhiteSpace(ReplayRequestPath) && ReplayRequestDryRun)
        {
            await ReplayRequestAsync(null, ReplayRequestPath, dryRun: true);
            return;
        }

        var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            Console.Write("Enter OPENAI_API_KEY: ");
            apiKey = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                Console.Error.WriteLine("Missing API key.");
                return;
            }
        }

        if (!string.IsNullOrWhiteSpace(ReplayRequestPath))
        {
            await ReplayRequestAsync(apiKey!, ReplayRequestPath, dryRun: false);
            return;
        }

        try { SetProcessDpiAwarenessContext((nint)(-4)); } catch { /* best effort */ }

        PrintStartupSummary();
        Console.WriteLine("Enter a task or question. Use '/ask ' for Q&A and '/exit' to quit.\n");

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
                Console.Write("Command/Question: ");
                goal = Console.ReadLine() ?? "";
            }

            if (string.IsNullOrWhiteSpace(goal) || goal.Trim().Equals("/exit", StringComparison.OrdinalIgnoreCase))
                break;

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
                var marker = result.Outcome switch
                {
                    ControlRunOutcome.Completed => "✅",
                    ControlRunOutcome.Cancelled => "⏹",
                    ControlRunOutcome.GuardStopped => "🛑",
                    ControlRunOutcome.StepLimitReached => "⏱",
                    _ => "❌"
                };
                Console.WriteLine(
                    $"{marker} {result.Outcome}: {result.Message} " +
                    $"(step={result.Step}). Enter next (ENTER = exit).");
            }
        }
    }
}
