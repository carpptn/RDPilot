
namespace RDPilot;

/// <summary>
/// Coordinates the interactive desktop-agent application and its shared runtime state.
/// </summary>
internal static class RDPilotApplication
{
    static readonly AppOptions ConfiguredOptions = new();
    internal static AppOptions Options => CurrentRun.Configuration ?? ConfiguredOptions;
    internal static RunContext CurrentRun { get; private set; } = new();
    static int runActive;

    internal static IDisposable BeginRun()
    {
        // Console redirection and desktop input are process-wide resources.
        if (Interlocked.CompareExchange(ref runActive, 1, 0) != 0)
            throw new InvalidOperationException("A desktop run is already active.");
        try
        {
            var previous = CurrentRun;
            CurrentRun = new RunContext { Configuration = Options.Snapshot() };
            return new RunScope(previous);
        }
        catch
        {
            Volatile.Write(ref runActive, 0);
            throw;
        }
    }

    sealed class RunScope(RunContext previous) : IDisposable
    {
        int disposed;
        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
            CurrentRun = previous;
            Volatile.Write(ref runActive, 0);
        }
    }

    internal const int MaxStepsDefault = AppOptions.MaxStepsDefault;

    internal static readonly JsonSerializerOptions PrettyJson = new(JsonSerializerDefaults.Web) { WriteIndented = true };
    internal static readonly HashSet<string> AllowedReasoningEfforts = new(StringComparer.OrdinalIgnoreCase)
    {
        "none", "minimal", "low", "medium", "high", "xhigh", "max"
    };
    internal static readonly HashSet<string> KnownActionTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "open_url", "launch_app", "run_command",
        "paste_text", "focus_uia", "click_uia",
        "move", "click", "double_click", "drag_drop", "drag_path", "keys", "hold_keys", "type_text", "scroll",
        "request_crop", "point", "aim", "wait", "done"
    };

    static async Task<int> Main(string[] args)
    {
        try
        {
            Console.OutputEncoding = Encoding.UTF8;
            ConsoleTheme.Enable();
            AppDomain.CurrentDomain.ProcessExit += (_, _) => ReleaseAllHeldInputs();

            ApplyEnvironmentConfig();
            var cli = ParseCliArgs(args);
            string? pending = cli.Task;
            NormalizeConfig();
            ConfigureOpenAiHttpClient();

            if (!cli.Success || CurrentRun.ConfigurationErrors.Count > 0)
                return 2;

            if (Options.PrintConfigOnly)
            {
                PrintEffectiveConfig();
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(Options.RecoveryMemoryCommand))
            {
                return (int)ExecuteRecoveryMemoryMaintenance(
                    Options.RecoveryMemoryCommand,
                    Options.RecoveryMemoryExportPath);
            }
            if (!string.IsNullOrWhiteSpace(Options.LoopReplayPath))
            {
                return (int)ExecuteLoopReplay(Options.LoopReplayPath).Outcome;
            }
            if (!string.IsNullOrWhiteSpace(Options.LoopReplayImportPath))
            {
                ImportIndependentLoopReplayCorpus(Options.LoopReplayImportPath);
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(Options.LoopReplayExportPath))
            {
                ExportLoopTelemetryToReplayCorpus(Options.LoopReplayExportPath);
                return 0;
            }
            if (Options.AnalyzeLogsOnly)
            {
                AnalyzeArtifacts();
                return 0;
            }
            if (!string.IsNullOrWhiteSpace(Options.ReplayResponsePath))
            {
                return (int)ReplayResponse(Options.ReplayResponsePath);
            }
            if (!string.IsNullOrWhiteSpace(Options.ReplayRequestPath) && Options.ReplayRequestDryRun)
            {
                return (int)await ReplayRequestAsync(null, Options.ReplayRequestPath, dryRun: true);
            }

            var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                if (Options.BatchMode || Console.IsInputRedirected)
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

            if (!string.IsNullOrWhiteSpace(Options.ReplayRequestPath))
            {
                return (int)await ReplayRequestAsync(apiKey!, Options.ReplayRequestPath, dryRun: false);
            }

            try { SetProcessDpiAwarenessContext((nint)(-4)); } catch { /* best effort */ }

            PrintStartupSummary();
            if (Options.BatchMode)
                return await RunBatchTaskAsync(apiKey!, pending);

            var promptHistory = PromptHistoryService.Load();
            Console.WriteLine(
                "Enter a task or question. Use Up/Down for prompt history, '/ask ' for Q&A, and '/exit' to quit.");
            Console.WriteLine();

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
                    await RunInteractiveQuestionAsync(apiKey!, goal);
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

    internal static async Task<CommandOutcome> RunInteractiveQuestionAsync(string apiKey, string goal)
    {
        var outcome = await RunAskOnce(apiKey, goal);
        Console.WriteLine();
        CommandPresentation.WriteQuestionResult(outcome, batch: false);
        return outcome;
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
            var outcome = await RunAskOnce(apiKey, goal);
            CommandPresentation.WriteQuestionResult(outcome, batch: true);
            return (int)outcome;
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
