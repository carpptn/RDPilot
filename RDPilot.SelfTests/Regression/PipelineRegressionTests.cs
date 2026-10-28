using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class PipelineRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("invalid first batched action is recoverable within the failure budget", () => Run("invalid-first-batch", ControlRunOutcome.Completed, 2).GetAwaiter().GetResult()),
        ("invalid single action is recoverable within the failure budget", () => Run("invalid-single", ControlRunOutcome.Completed, 2).GetAwaiter().GetResult()),
        ("rejected decisions cannot leak queued text into the next step", () => Run("stale-queue", ControlRunOutcome.StepLimitReached, 3).GetAwaiter().GetResult()),
        ("streamed decisions retain safe follow-ups and their failure accounting", () => Run("streamed-batch-failure", ControlRunOutcome.GuardStopped).GetAwaiter().GetResult()),
        ("batch failures enforce the action failure limit", () => Run("batch-failure", ControlRunOutcome.GuardStopped).GetAwaiter().GetResult()),
        ("successful decisions reset preceding batch failures", () => Run("failure-reset", ControlRunOutcome.StepLimitReached, 3).GetAwaiter().GetResult()),
        ("batch cancellation produces a cancelled run", () => Run("batch-cancel", ControlRunOutcome.Cancelled).GetAwaiter().GetResult()),
        ("missing UIA index is recoverable within the failure budget", () => Run("uia-missing", ControlRunOutcome.Completed, 2).GetAwaiter().GetResult()),
        ("unknown UIA index obeys the failure budget", () => Run("uia-unknown", ControlRunOutcome.GuardStopped, 2).GetAwaiter().GetResult()),
        ("zero model failure limit permits transient retries", () => Run("retry-zero", ControlRunOutcome.Completed, 2).GetAwaiter().GetResult()),
        ("zero model failure limit still stops permanent errors", () => Run("permanent-zero", ControlRunOutcome.Failed).GetAwaiter().GetResult()),
        ("control context fallback restores an omitted screenshot", () => Run("context-image-fallback", ControlRunOutcome.Completed, 3).GetAwaiter().GetResult()),
        ("incomplete verifier cannot complete a control run", () => Run("incomplete-verifier", ControlRunOutcome.Failed, 2).GetAwaiter().GetResult()),
        ("control pipeline stops repeated rejected completions", () => Run("completion-reject", ControlRunOutcome.GuardStopped, 6).GetAwaiter().GetResult()),
        ("controller successes cannot hide repeated verifier failures", () => Run("completion-error", ControlRunOutcome.Failed, 4).GetAwaiter().GetResult()),
        ("control pipeline cancels even after a late verifier success", () => Run("completion-cancel", ControlRunOutcome.Cancelled, 2).GetAwaiter().GetResult()),
        ("control pipeline completes after observation and decision", () => Run("done", ControlRunOutcome.Completed).GetAwaiter().GetResult()),
        ("control pipeline honors the step limit after execution", () => Run("wait", ControlRunOutcome.StepLimitReached).GetAwaiter().GetResult()),
        ("control pipeline fails on permanent model errors", () => Run("bad-request", ControlRunOutcome.Failed).GetAwaiter().GetResult()),
        ("control pipeline retries transient model failures", () => Run("retry", ControlRunOutcome.Completed, 2).GetAwaiter().GetResult()),
        ("control pipeline propagates cancellation", () => Run("cancel", ControlRunOutcome.Cancelled).GetAwaiter().GetResult()),
        ("control pipeline cleans up after capture failure", () => Run("capture-failure", ControlRunOutcome.Failed, 0).GetAwaiter().GetResult())
    ];

    sealed class SyntheticCapture(bool fail) : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multipleMonitors) => new(0, 0, 320, 200);
        public void CopyPixels(Graphics destination, Rectangle source)
        {
            if (fail) throw new InvalidOperationException("synthetic capture failure");
            destination.Clear(Color.Navy);
            destination.FillRectangle(Brushes.White, 30, 30, 160, 100);
        }
    }

    sealed class Responses(string scenario) : HttpMessageHandler
    {
        internal int Calls;
        internal List<string> Bodies = [];
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Calls++;
            var payload = await request.Content!.ReadAsStringAsync(cancellationToken);
            Bodies.Add(payload);
            using var body = JsonDocument.Parse(payload);
            if (scenario.StartsWith("completion-", StringComparison.Ordinal) &&
                body.RootElement.GetProperty("text").GetProperty("format").GetProperty("name").GetString() == "VerifyGoal")
            {
                if (scenario == "completion-cancel") A.CurrentRun.CancelRequested = true;
                return new(scenario == "completion-error" ? HttpStatusCode.ServiceUnavailable : HttpStatusCode.OK)
                {
                    Content = new StringContent(JsonSerializer.Serialize(new {
                        status = "completed",
                        output_text = JsonSerializer.Serialize(new { verdict = scenario == "completion-reject" ? "no" : "yes", reason = "test verdict" })
                    }))
                };
            }
            if (scenario == "context-image-fallback" && Calls == 2)
                return new(HttpStatusCode.NotFound) { Content = new StringContent("{\"error\":{\"message\":\"previous_response_id not found\",\"param\":\"previous_response_id\"}}") };
            if (body.RootElement.TryGetProperty("temperature", out _))
                throw new InvalidOperationException("unsupported temperature in pipeline request");
            if (scenario == "incomplete-verifier" && Calls > 1)
                return new(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(new
                {
                    status = "incomplete", incomplete_details = new { reason = "max_output_tokens" },
                    output_text = "{\"verdict\":\"yes\",\"reason\":\"partial response\"}"
                })) };
            if (scenario == "cancel") A.CurrentRun.CancelRequested = true;
            if (scenario is "bad-request" or "cancel" or "permanent-zero" || scenario is "retry" or "retry-zero" && Calls == 1)
                return new HttpResponseMessage(scenario is "retry" or "retry-zero" ? HttpStatusCode.ServiceUnavailable : HttpStatusCode.BadRequest)
                { Content = new StringContent("{\"error\":{\"message\":\"synthetic failure\"}}") };
            var action = new ActionDto { Type = scenario == "wait" || scenario == "context-image-fallback" && Calls == 1 ? "wait" : "done", WaitSeconds = 0, Confidence = 0.99, Note = "synthetic test" };
            ActionDto[] actions = [action];
            ActionDto focus = new() { Type = "keys", Keys = ["ctrl+l"], Confidence = 0.99 };
            ActionDto text = new() { Type = "type_text", Text = "STALE", Confidence = 0.99 };
            ActionDto commit = new() { Type = "keys", Keys = ["enter"], Confidence = 0.99 };
            if (scenario is "batch-failure" or "streamed-batch-failure" or "batch-cancel" || scenario == "failure-reset" && Calls != 2)
                actions = [focus, text, commit];
            if (scenario == "stale-queue")
                actions = Calls == 1 ? [focus] : Calls == 2 ? [focus, text, commit] : [new() { Type = "wait", WaitSeconds = 0 }];
            if (scenario == "failure-reset" && Calls == 2)
                actions = [new() { Type = "wait", WaitSeconds = 0 }];
            if (scenario == "uia-missing" && Calls == 1 || scenario == "uia-unknown")
                actions = [new() { Type = "focus_uia", UiaIndex = scenario == "uia-unknown" ? 999 : null }];
            if (Calls == 1 && scenario is "invalid-first-batch" or "invalid-single")
                actions = scenario == "invalid-single" ? [new() { Type = "click" }] : [new() { Type = "click" }, new() { Type = "wait", WaitSeconds = 0 }];
            var response = JsonSerializer.Serialize(new
            {
                id = "response_test_" + Calls,
                status = "completed",
                output_text = JsonSerializer.Serialize(new { actions })
            });
            if (scenario == "streamed-batch-failure")
            {
                var streamEvent = JsonSerializer.Serialize(new { type = "response.output_text.done", text = JsonSerializer.Serialize(new { actions }) });
                return new HttpResponseMessage(HttpStatusCode.OK)
                { Content = new StringContent("data: " + streamEvent + "\n\n", Encoding.UTF8, "text/event-stream") };
            }
            return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(response, Encoding.UTF8, "application/json") };
        }
    }

    sealed class RecordedInput(string scenario) : DesktopInputService.IInputBackend
    {
        public bool SetCursorPosition(int x, int y) => throw new InvalidOperationException("Unexpected cursor movement in test.");
        public void SetClipboardText(string text) => throw new InvalidOperationException("Unexpected clipboard write in test.");
        internal int UnicodeAttempts;
        public int LastError => 5;
        public uint Send(DesktopInputService.INPUT[] inputs)
        {
            if (inputs.Any(input => input.type == 1 && (input.U.ki.dwFlags & 4) != 0))
            {
                UnicodeAttempts++;
                if (scenario == "batch-cancel")
                {
                    A.CurrentRun.CancelRequested = true;
                    throw new OperationCanceledException();
                }
                if (scenario is "batch-failure" or "streamed-batch-failure" or "failure-reset") return 0;
            }
            if (scenario is "stale-queue" or "batch-failure" or "streamed-batch-failure" or "failure-reset" or "batch-cancel")
                return (uint)inputs.Length;
            throw new InvalidOperationException("pipeline scenario must not send native input");
        }
    }

    static async Task Run(string scenario, ControlRunOutcome expected, int expectedCalls = 1)
    {
        var options = A.Options;
        var originalOptions = options.Snapshot();
        var context = A.CurrentRun;
        var originalClient = OpenAiTransport.HttpClient;
        var originalCapture = DesktopCapture.Backend;
        var originalContext = UiAutomationService.PromptContextProvider;
        var originalWindow = ScreenshotService.ActiveWindowRectangleProvider;
        var originalInput = DesktopInputService.InputBackend;
        var output = Console.Out;
        var error = Console.Error;
        var input = new RecordedInput(scenario);
        using var handler = new Responses(scenario);
        using var client = new HttpClient(handler);
        try
        {
            options.BatchMode = true;
            options.MaxSteps = expectedCalls == 0 ? 1 : expectedCalls;
            if (scenario is "batch-failure" or "streamed-batch-failure" or "batch-cancel") options.MaxSteps = 3;
            options.GoalMode = "finite";
            options.ObservationMode = "static";
            options.ObservationLogVerbose = scenario == "stale-queue";
            options.IncludeFocusUia = false;
            options.IncludeUiaTargets = false;
            options.LogScreens = false;
            options.LogRequests = false;
            options.MaxArtifactsPerDir = 0;
            options.RecoveryMemoryEnabled = false;
            options.LoopReplayAutoExportEnabled = false;
            options.AllowHighLevelActions = false;
            options.VerifyMode = scenario == "incomplete-verifier" || scenario.StartsWith("completion-", StringComparison.Ordinal) ? "always" : "off";
            if (scenario == "incomplete-verifier")
            {
                options.MaxSteps = 1;
                options.VerifyReasoningEffort = "low";
                options.VerifyReasoningEffortExplicit = true;
                options.IncompleteMaxOutputRetries = 0;
            }
            options.ScreenSanityChecks = false;
            options.ScreenPollingEnabled = false;
            options.WaitNoChangeExtraMs = 0;
            options.UiSettleDelayMs = 0;
            options.OpenAiMaxRetries = 0;
            options.UsePreviousResponseState = scenario == "context-image-fallback";
            if (scenario == "context-image-fallback")
            {
                options.MaxSteps = 2;
                options.OmitUnchangedScreenImageWithState = true;
            }
            options.ControlContextCompactionEnabled = false;
            options.MaxModelFailuresBeforeAbort = scenario is "retry-zero" or "permanent-zero" ? 0 : 2;
            options.MaxActionFailuresBeforeAbort = scenario is "batch-failure" or "streamed-batch-failure" ? 1 : 2;
            options.MaxStagnationStepsBeforeAbort = 0;
            options.MaxRepeatedActionBeforeAbort = 0;
            options.MaxRejectedProposalRepeatsBeforeAbort = scenario == "completion-reject" ? 2 : 0;
            options.ActionRepeatCooldownSteps = scenario == "stale-queue" ? 4 : 0;
            options.SendInputMaxRetries = 0;
            options.SendInputRetryDelayMs = 0;
            options.ExecuteMultiActionCandidates = true;
            options.MaxQueuedBatchActions = 4;
            OpenAiTransport.HttpClient = client;
            DesktopCapture.Backend = new SyntheticCapture(scenario == "capture-failure");
            UiAutomationService.PromptContextProvider = (_, _, _) => new("Synthetic", "synthetic", null, null, null, null);
            ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(0, 0, 320, 200);
            DesktopInputService.InputBackend = input;
            var result = await ControlLoopService.RunOnce("synthetic-key", "Inspect the synthetic test screen");
            if (result.Outcome != expected || handler.Calls != expectedCalls)
                throw new InvalidOperationException($"expected {expected}/{expectedCalls} requests, got {result.Outcome}/{handler.Calls}: {result}");
            if (!ReferenceEquals(context, A.CurrentRun) || !ReferenceEquals(output, Console.Out) || !ReferenceEquals(error, Console.Error))
                throw new InvalidOperationException("pipeline leaked context or console redirection");
            if (scenario == "stale-queue" && input.UnicodeAttempts != 0)
                throw new InvalidOperationException("stale text from a rejected decision was executed");
            if (scenario is "batch-failure" or "streamed-batch-failure" && input.UnicodeAttempts != 1)
                throw new InvalidOperationException("batch failure test did not reach exactly one failed text input");
            if (scenario == "failure-reset" && input.UnicodeAttempts != 2)
                throw new InvalidOperationException("successful decision did not separate the two failed batches");
            if (scenario == "context-image-fallback")
            {
                var chained = handler.Bodies[1];
                var fallback = handler.Bodies[2];
                if (!chained.Contains("previous_response_id") || !chained.Contains("SCREEN_IMAGE: omitted") ||
                    !fallback.Contains("input_image") || !fallback.Contains("SCREEN_IMAGE: attached") ||
                    fallback.Contains("SCREEN_IMAGE: omitted"))
                    throw new InvalidOperationException("control request did not rebuild visual evidence after context loss");
            }
            var expectedStep = scenario switch
            {
                "completion-reject" => 4,
                "context-image-fallback" or "completion-error" => 2,
                "incomplete-verifier" or "completion-cancel" => 1,
                _ => Math.Max(1, expectedCalls)
            };
            if (result.Step != expectedStep)
                throw new InvalidOperationException($"pipeline returned step {result.Step}, expected {expectedStep}");
        }
        finally
        {
            OpenAiTransport.HttpClient = originalClient;
            DesktopCapture.Backend = originalCapture;
            UiAutomationService.PromptContextProvider = originalContext;
            ScreenshotService.ActiveWindowRectangleProvider = originalWindow;
            DesktopInputService.ReleaseAllHeldInputs();
            DesktopInputService.InputBackend = originalInput;
            foreach (var field in typeof(AppOptions).GetFields(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic))
                if (!field.IsInitOnly) field.SetValue(options, field.GetValue(originalOptions));
        }
    }
}
