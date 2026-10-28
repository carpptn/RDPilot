using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;

namespace RDPilot.SelfTests.Regression;

internal static class CompletionRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("rejected completions reach the proposal cycle guard", RejectedCompletions),
        ("permanent verifier failure stops the goal", PermanentFailure),
        ("transient verifier failures have an independent bounded counter", TransientFailures),
        ("valid negative verdict resets verifier failures", ValidVerdictResetsFailures),
        ("invalid verifier payload fails instead of becoming a negative verdict", InvalidVerdict),
        ("cancelled completion cannot bypass verification", () => AlreadyCancelled(false)),
        ("completion honors cancellation token without the hotkey flag", () => AlreadyCancelled(true)),
        ("late verifier success cannot complete a cancelled goal", () => LateCancellation("yes")),
        ("late verifier rejection preserves cancellation", () => LateCancellation("no")),
        ("late verifier failure preserves cancellation", () => LateCancellation("http403")),
        ("verified completion restores the controller coordinate map", VerifiedCompletion),
        ("disabled verification needs no screenshot or API request", UnverifiedCompletion),
        ("loop replay returns successful regression counts", () => ReplayResult(false, true)),
        ("loop replay returns failures for independent cases", () => ReplayResult(true, true)),
        ("loop replay returns failures without independent labels", () => ReplayResult(true, false)),
        ("loop replay CLI returns zero for matching cases", () => ReplayCli(false)),
        ("loop replay CLI returns one for mismatched cases", () => ReplayCli(true))
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class Responses(params string[] replies) : HttpMessageHandler
    {
        internal int Calls;
        internal bool Cancel;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            var reply = replies[Math.Min(Calls++, replies.Length - 1)];
            if (Cancel) DesktopConsoleService.CancelRequested = true;
            var status = reply == "http403" ? HttpStatusCode.Forbidden :
                reply == "http503" ? HttpStatusCode.ServiceUnavailable : HttpStatusCode.OK;
            var payload = reply == "invalid" ? "{}" : JsonSerializer.Serialize(new { verdict = reply, reason = "test verdict" });
            return Task.FromResult(new HttpResponseMessage(status)
            {
                Content = new StringContent(JsonSerializer.Serialize(new { status = "completed", output_text = payload }), Encoding.UTF8, "application/json")
            });
        }
    }

    sealed class Scenario : IDisposable
    {
        internal readonly Fixture Fixture = new();
        internal readonly Responses Responses;
        readonly HttpClient previous = OpenAiTransport.HttpClient;
        readonly HttpClient client;
        internal Scenario(params string[] replies)
        {
            A.Options.VerifyMode = "always";
            A.Options.RefreshScreenshotBeforeVerify = false;
            A.Options.VerifyScreenshotMaxWidth = 160;
            A.Options.OpenAiMaxRetries = 0;
            A.Options.MaxModelFailuresBeforeAbort = 2;
            A.Options.MaxRejectedProposalRepeatsBeforeAbort = 2;
            Responses = new(replies);
            client = new(Responses);
            OpenAiTransport.HttpClient = client;
        }
        internal ControlStepState Turn(int step = 1)
        {
            var capture = ScreenshotService.ScreenshotToDataUrl(AppContext.BaseDirectory, "test", step, null);
            return new(step)
            {
                Decision = new ControlDecision(new() { Type = "done", Confidence = 1 }),
                promptContext = Fixture.Context,
                observationFrame = capture.ObservationFrame,
                dataUrl = capture.DataUrl,
                screenW = capture.ScreenWidth, screenH = capture.ScreenHeight,
                imageW = capture.ImageWidth, imageH = capture.ImageHeight
            };
        }
        internal StepFlow Execute(int step = 1) => Fixture.Execute(Turn(step));
        public void Dispose()
        {
            OpenAiTransport.HttpClient = previous;
            client.Dispose();
            Fixture.Dispose();
        }
    }

    static void RejectedCompletions()
    {
        using var s = new Scenario("no");
        for (var i = 1; i <= 3; i++) Check(s.Execute(i) == StepFlow.Next, "negative verdict aborted too early");
        var run = s.Fixture.Run;
        Check(run.rejectedProposalCycleCount == 2, "done refusals bypassed the proposal guard");
        Check(run.consecutiveActionFailures == 0 && run.consecutiveVerificationFailures == 0, "negative verdict counted as infrastructure/executor failure");
        Check(run.prevAction is null, "rejected completion became an executed action");
        var flow = ControlDecisionStage.RunAsync(run, new(4)).GetAwaiter().GetResult();
        Check(flow == StepFlow.Stop && run.runResult.Outcome == ControlRunOutcome.GuardStopped && s.Responses.Calls == 3,
            "guard did not stop repeated completions before another API call");
    }
    static void PermanentFailure()
    {
        using var s = new Scenario("http403");
        Check(s.Execute() == StepFlow.Stop && s.Fixture.Run.runResult.Outcome == ControlRunOutcome.Failed, "permanent verifier failure did not stop the goal");
        Check(s.Fixture.Run.runResult.Message.Contains("http_403") && s.Responses.Calls == 1, "verifier error lost or retried");
    }
    static void TransientFailures()
    {
        using var s = new Scenario("http503");
        Check(s.Execute() == StepFlow.Next, "transient error was not retried");
        // Every successful controller response clears this existing counter.
        s.Fixture.Run.consecutiveModelFailures = 0;
        Check(s.Execute(2) == StepFlow.Stop && s.Fixture.Run.runResult.Outcome == ControlRunOutcome.Failed, "verifier failures escaped their limit");
        Check(s.Fixture.Run.consecutiveVerificationFailures == 2 && s.Responses.Calls == 2, "wrong verifier attempt count");
    }
    static void ValidVerdictResetsFailures()
    {
        using var s = new Scenario("http503", "no", "http503", "yes");
        Check(s.Execute() == StepFlow.Next && s.Execute(2) == StepFlow.Next, "initial attempts failed");
        Check(s.Fixture.Run.consecutiveVerificationFailures == 0, "valid no did not reset verifier failures");
        Check(s.Execute(3) == StepFlow.Next && s.Execute(4) == StepFlow.Stop, "recovery after transient failure failed");
        Check(s.Fixture.Run.runResult.Completed && s.Fixture.Run.consecutiveVerificationFailures == 0, "valid yes did not complete/reset");
    }
    static void InvalidVerdict()
    {
        using var s = new Scenario("invalid");
        Check(s.Execute() == StepFlow.Stop && s.Fixture.Run.runResult.Outcome == ControlRunOutcome.Failed, "empty verdict did not fail");
        Check(s.Fixture.Run.rejectedProposalCycleCount == 0, "invalid payload was recorded as a negative verdict");
    }
    static void AlreadyCancelled(bool tokenOnly)
    {
        using var s = new Scenario("yes");
        A.Options.VerifyMode = "off";
        var turn = s.Turn();
        if (tokenOnly)
            ((CancellationTokenSource)typeof(CancellationMonitor).GetField("source", BindingFlags.Instance | BindingFlags.NonPublic)!
                .GetValue(s.Fixture.Run.cancelCts)!).Cancel();
        else DesktopConsoleService.CancelRequested = true;
        Check(s.Fixture.Execute(turn) == StepFlow.Stop && s.Fixture.Run.runResult.Outcome == ControlRunOutcome.Cancelled, "cancelled completion was accepted");
        Check(s.Responses.Calls == 0, "cancelled completion called the API");
    }
    static void LateCancellation(string reply)
    {
        using var s = new Scenario(reply);
        s.Responses.Cancel = true;
        var map = A.CurrentRun.CurrentScreenMap;
        Check(s.Execute() == StepFlow.Stop && s.Fixture.Run.runResult.Outcome == ControlRunOutcome.Cancelled, "late result overrode cancellation");
        Check(A.CurrentRun.CurrentScreenMap == map && s.Fixture.Run.consecutiveVerificationFailures == 0, "cancelled verification leaked map/failure state");
    }
    static void VerifiedCompletion()
    {
        using var s = new Scenario("yes");
        A.Options.RefreshScreenshotBeforeVerify = true;
        var map = A.CurrentRun.CurrentScreenMap;
        Check(s.Execute() == StepFlow.Stop && s.Fixture.Run.runResult.Completed, "verified goal did not complete");
        Check(s.Responses.Calls == 1 && A.CurrentRun.CurrentScreenMap == map, "helper map was not restored");
    }
    static void UnverifiedCompletion()
    {
        using var s = new Scenario("http403");
        A.Options.VerifyMode = "off";
        A.Options.RefreshScreenshotBeforeVerify = true;
        var turn = s.Turn();
        turn.dataUrl = "deliberately not an image";
        var captures = A.CurrentRun.RunScreenshotCount;
        Check(s.Fixture.Execute(turn) == StepFlow.Stop && s.Fixture.Run.runResult.Completed, "verification-off completion failed");
        Check(s.Responses.Calls == 0 && A.CurrentRun.RunScreenshotCount == captures, "disabled verifier performed work");
    }

    sealed class Corpus : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string FilePath => Path.Combine(Root, "corpus.json");
        internal Corpus(bool mismatch, bool independent = true)
        {
            Directory.CreateDirectory(Root);
            var json = JsonSerializer.Serialize(new { cases = new[] { new {
                name = "stable-screen", labelSource = independent ? "manual:reviewed" : "telemetry:regression",
                expectedLoop = mismatch, screenWidth = 320, screenHeight = 200,
                frames = new[] { new { screenFingerprintBase64 = "AAAA" }, new { screenFingerprintBase64 = "AAAA" } }
            } } });
            File.WriteAllText(FilePath, json, new UTF8Encoding(true));
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
    static void ReplayResult(bool mismatch, bool independent)
    {
        using var corpus = new Corpus(mismatch, independent);
        var result = LoopReplayService.ExecuteLoopReplay(corpus.FilePath);
        Check(result.CaseCount == 1 && result.FailureCount == (mismatch ? 1 : 0), "wrong replay counts");
        Check(result.TruePositive == 0 && result.FalsePositive == 0 && result.TrueNegative == (mismatch ? 0 : 1) &&
            result.FalseNegative == (mismatch ? 1 : 0), "wrong confusion matrix");
        Check(result.Outcome == (mismatch ? CommandOutcome.Failed : CommandOutcome.Success), "wrong replay outcome");
    }
    static void ReplayCli(bool mismatch)
    {
        using var corpus = new Corpus(mismatch);
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        {
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        foreach (var arg in new[] { "--command-main-probe", "--loop-replay", corpus.FilePath }) start.ArgumentList.Add(arg);
        using var child = Process.Start(start)!;
        try
        {
            var output = child.StandardOutput.ReadToEndAsync();
            var error = child.StandardError.ReadToEndAsync();
            Check(child.WaitForExit(10000), "loop replay CLI did not exit");
            Task.WhenAll(output, error).GetAwaiter().GetResult();
            Check(child.ExitCode == (mismatch ? 1 : 0), $"wrong replay exit code: {child.ExitCode}; {error.Result}");
            Check(output.Result.Contains(mismatch ? "[FAIL]" : "[PASS]"), "report disagreed with exit code");
        }
        finally { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } }
    }
}
