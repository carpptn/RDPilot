using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;

namespace RDPilot.SelfTests.Regression;

internal static class ReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("failed mouse-up remains pending until cleanup succeeds", DragRelease),
        ("missing primary recovery store restores its backup", MissingPrimary),
        ("CLI validation rejects malformed values and missing arguments", InvalidCli),
        ("CLI profile syntaxes have identical override precedence", ProfileSyntax),
        ("CLI task values cannot be mistaken for profile flags", ProfileInTask),
        ("environment profile precedence warning is not a configuration error", EnvironmentProfilePrecedence),
        ("invalid boolean configuration is reported to the caller", InvalidBoolean),
        ("replay reports input validation failures", () => ReplayValidation().GetAwaiter().GetResult()),
        ("replay rejects incomplete HTTP success", () => ReplayHttp("incomplete", CommandOutcome.Failed).GetAwaiter().GetResult()),
        ("replay supports final SSE responses from streamed request logs", () => ReplayHttp("stream", CommandOutcome.Success).GetAwaiter().GetResult()),
        ("replay reports HTTP success", () => ReplayHttp("ok", CommandOutcome.Success).GetAwaiter().GetResult()),
        ("replay reports HTTP failure", () => ReplayHttp("error", CommandOutcome.Failed).GetAwaiter().GetResult()),
        ("replay reports cancellation and restores its run context", () => ReplayHttp("cancel", CommandOutcome.Cancelled).GetAwaiter().GetResult())
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    static string TempDirectory()
    {
        var root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }

    static void WriteText(string path, string text) =>
        File.WriteAllText(path, text.Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));

    sealed class DragInput(uint down, uint up) : Input.IInputBackend
    {
        public bool SetCursorPosition(int x, int y) => throw new InvalidOperationException("Unexpected cursor movement in test.");
        public void SetClipboardText(string text) => throw new InvalidOperationException("Unexpected clipboard write in test.");
        internal bool RejectRelease = true;
        internal bool Held;
        internal int UpAttempts;
        public int LastError => 5;
        public uint Send(Input.INPUT[] inputs)
        {
            var accepted = 0u;
            foreach (var input in inputs)
            {
                if (input.type == 0 && (input.U.mi.dwFlags & up) != 0)
                {
                    UpAttempts++;
                    if (RejectRelease) return accepted;
                    Held = false;
                }
                if (input.type == 0 && (input.U.mi.dwFlags & down) != 0) Held = true;
                accepted++;
            }
            return accepted;
        }
    }

    static void DragRelease()
    {
        using var scope = A.BeginRun();
        A.Options.SendInputMaxRetries = 0;
        var originalInput = Input.InputBackend;
        try
        {
            foreach (var (button, down, up) in new[] { ("left", 2u, 4u), ("right", 8u, 16u), ("middle", 32u, 64u) })
            {
                var input = new DragInput(down, up);
                Input.InputBackend = input;
                A.CurrentRun.CancelRequested = false;
                var failed = false;
                try { Input.MouseDragPath([new Point(10, 10), new Point(30, 30)], button, 100); }
                catch (InvalidOperationException) { failed = true; }
                Check(failed && input.Held && input.UpAttempts >= 2, "failed drag release was not tracked or retried");
                input.RejectRelease = false;
                A.CurrentRun.CancelRequested = true;
                var attempts = input.UpAttempts;
                Input.ReleaseAllHeldInputs();
                Check(!input.Held && input.UpAttempts == attempts + 1, "run cleanup lost the held mouse button");
                Input.ReleaseAllHeldInputs();
                Check(input.UpAttempts == attempts + 1, "cleanup replayed an already accepted release");
            }
        }
        finally
        {
            if (Input.InputBackend is DragInput input) input.RejectRelease = false;
            Input.ReleaseAllHeldInputs();
            Input.InputBackend = originalInput;
        }
    }

    static void MissingPrimary()
    {
        using var scope = A.BeginRun();
        var root = TempDirectory();
        var path = Path.Combine(root, "memory.json");
        A.Options.RecoveryMemoryEnabled = true;
        A.Options.RecoveryMemoryPath = path;
        A.Options.RecoveryMemoryArchivePath = Path.Combine(root, "archive.json");
        var json = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        var store = new RecoveryLessonStore { Lessons = [new() { Id = "valuable-lesson", UpdatedUtc = DateTime.UtcNow }] };
        try
        {
            WriteText(path + ".bak", JsonSerializer.Serialize(store, json));
            var lessons = RecoveryMemoryService.LoadRecoveryLessons();
            Check(lessons.Count == 1 && lessons[0].Id == "valuable-lesson", "backup was ignored when primary was missing");
            var save = typeof(RecoveryMemoryService).GetMethod("SaveRecoveryLessons", BindingFlags.NonPublic | BindingFlags.Static)!;
            Check((bool)save.Invoke(null, [lessons])!, "subsequent recovery save failed");
            foreach (var file in new[] { path, path + ".bak" })
                Check(JsonSerializer.Deserialize<RecoveryLessonStore>(File.ReadAllText(file), json)!.Lessons.Single().Id == "valuable-lesson", "valid recovery lesson was lost");
            Check(RecoveryMemoryService.LoadRecoveryLessons().Single().Id == "valuable-lesson", "restarting lost recovered data");
        }
        finally { Directory.Delete(root, true); }
    }

    static void InvalidCli()
    {
        using var scope = A.BeginRun();
        string[][] invalid =
        [
            ["--max-steps", "oops"], ["--max-steps", "-1"], ["--max-steps="],
            ["--max-steps", "--batch"], ["--max-steps"], ["--profile=unknown"],
            ["--profile", "unknown", "--quality"], ["--verify-low-confidence", "NaN"],
            ["--verify", "invalid"], ["--batch=false"], ["--task="],
            ["--task", "one", "--task=two"], ["--unrecognized-option"],
            ["--screenshot-format=gif"], ["--reasoning-effort=invalid"]
        ];
        foreach (var args in invalid)
        {
            var result = ConfigurationService.ParseCliArgs(args);
            Check(!result.Success && A.CurrentRun.CliArgumentError, "invalid CLI succeeded: " + string.Join(" ", args));
        }
        var valid = ConfigurationService.ParseCliArgs(["--batch", "--max-steps=7", "--task", "synthetic task"]);
        Check(valid.Success && !A.CurrentRun.CliArgumentError && A.Options.MaxSteps == 7 && valid.Task == "synthetic task", "validation leaked between parses");
    }

    static void ProfileSyntax()
    {
        using var scope = A.BeginRun();
        var spaced = ConfigurationService.ParseCliArgs(["--max-steps", "7", "--profile", "quality"]);
        var expected = A.Options.Snapshot();
        ConfigurationService.ApplyProfile("fast");
        var equals = ConfigurationService.ParseCliArgs(["--profile=quality", "--max-steps=7"]);
        Check(spaced.Success && equals.Success, "valid profile syntax was rejected");
        foreach (var field in typeof(AppOptions).GetFields(BindingFlags.Instance | BindingFlags.NonPublic))
            Check(Equals(field.GetValue(expected), field.GetValue(A.Options)), "profile syntax changed " + field.Name);
        Check(A.Options.MaxSteps == 7, "profile erased an explicit option override");
    }

    static void ProfileInTask()
    {
        using var scope = A.BeginRun();
        ConfigurationService.ApplyProfile("balanced");
        var expected = A.Options.Snapshot();
        var parsed = ConfigurationService.ParseCliArgs(["--task=--quality"]);
        Check(parsed.Success && parsed.Task == "--quality", "task text was reinterpreted as an option");
        Check(A.Options.MaxOutputTokens == expected.MaxOutputTokens && A.Options.ReasoningEffort == expected.ReasoningEffort, "task text selected a profile");
    }

    static void EnvironmentProfilePrecedence()
    {
        using var scope = A.BeginRun();
        var originalProfile = Environment.GetEnvironmentVariable("RDPILOT_PROFILE");
        var originalFast = Environment.GetEnvironmentVariable("FAST_MODE");
        try
        {
            Environment.SetEnvironmentVariable("RDPILOT_PROFILE", "quality");
            Environment.SetEnvironmentVariable("FAST_MODE", "1");
            ConfigurationService.ApplyEnvironmentProfileOverrides();
            Check(A.CurrentRun.ConfigurationErrors.Count == 0, "a valid profile override was classified as an error");
        }
        finally
        {
            Environment.SetEnvironmentVariable("RDPILOT_PROFILE", originalProfile);
            Environment.SetEnvironmentVariable("FAST_MODE", originalFast);
        }
    }

    static void InvalidBoolean()
    {
        using var scope = A.BeginRun();
        var assigned = false;
        ConfigurationService.ApplyBool("maybe", "test", _ => assigned = true);
        Check(!assigned && A.CurrentRun.ConfigurationErrors.Count == 1, "invalid boolean silently kept defaults");
    }

    static async Task ReplayValidation()
    {
        var root = TempDirectory();
        var path = Path.Combine(root, "request.json");
        try
        {
            Check(await ArtifactAnalysisService.ReplayRequestAsync(null, path, true) == CommandOutcome.Failed, "missing replay file succeeded");
            foreach (var content in new[]
            {
                "{broken", "[]", "{}", "{\"body\":{}}",
                "{\"model\":null,\"input\":[],\"text\":{}}",
                "{\"model\":\"synthetic\",\"input\":false,\"text\":{}}",
                "{\"model\":\"synthetic\",\"input\":[],\"text\":null}"
            })
            {
                WriteText(path, content);
                Check(await ArtifactAnalysisService.ReplayRequestAsync(null, path, true) == CommandOutcome.Failed, "invalid replay JSON succeeded");
            }
            WriteText(path, JsonSerializer.Serialize(new { model = "synthetic", input = new[] { new { image_url = "file://" + Path.Combine(root, "missing.png") } }, text = new { } }));
            Check(await ArtifactAnalysisService.ReplayRequestAsync(null, path, true) == CommandOutcome.Failed, "missing replay image succeeded");
            WriteText(path, "{\"model\":\"synthetic\",\"input\":[],\"text\":{}}");
            Check(await ArtifactAnalysisService.ReplayRequestAsync(null, path, true) == CommandOutcome.Success, "valid dry run failed");
            Check(await ArtifactAnalysisService.ReplayRequestAsync(null, path, false) == CommandOutcome.Failed, "live replay without a key succeeded");
        }
        finally { Directory.Delete(root, true); }
    }

    sealed class ReplayHandler(string scenario) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            if (scenario == "cancel") A.CurrentRun.CancelRequested = true;
            var response = JsonSerializer.Serialize(new { status = scenario == "incomplete" ? "incomplete" : "completed", output_text = "synthetic replay" });
            if (scenario == "stream")
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                { Content = new StringContent("data: {\"type\":\"response.completed\",\"response\":" + response + "}\n\n", Encoding.UTF8, "text/event-stream") });
            return Task.FromResult(new HttpResponseMessage(scenario is "ok" or "incomplete" ? HttpStatusCode.OK : HttpStatusCode.BadRequest)
            { Content = new StringContent(response, Encoding.UTF8, "application/json") });
        }
    }

    static async Task ReplayHttp(string scenario, CommandOutcome expected)
    {
        var root = TempDirectory();
        var path = Path.Combine(root, "request.json");
        var originalClient = OpenAiTransport.HttpClient;
        var originalContext = A.CurrentRun;
        using var client = new HttpClient(new ReplayHandler(scenario));
        try
        {
            WriteText(path, "{\"model\":\"synthetic\",\"input\":[],\"text\":{}}");
            OpenAiTransport.HttpClient = client;
            var result = await ArtifactAnalysisService.ReplayRequestAsync("synthetic-key", path, false);
            Check(result == expected, $"replay expected {expected}, got {result}");
            Check(ReferenceEquals(originalContext, A.CurrentRun), "replay leaked its run context");
        }
        finally
        {
            OpenAiTransport.HttpClient = originalClient;
            Directory.Delete(root, true);
        }
    }
}
