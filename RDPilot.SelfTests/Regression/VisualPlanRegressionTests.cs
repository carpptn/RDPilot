using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;

namespace RDPilot.SelfTests.Regression;

internal static class VisualPlanRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("rejected action cannot hide a changed screen from the model", () => RequestFreshness(true, false)),
        ("pre-action baseline cannot replace the image known to the model", () => RequestFreshness(true, true)),
        ("unchanged model evidence remains reusable after action baseline reset", () => RequestFreshness(false, false)),
        ("UIA reuse requires the image and context of the actual target list", TargetFreshness),
        ("screen bounds and semantic context invalidate image reuse", ImageContext),
        ("failed responses cannot advance model image evidence", FailedResponse),
        ("context restoration records the newly attached image", RestoredResponse),
        ("retained key batches bind against the remaining route", () => RemainingBatch(false)),
        ("retained click batches bind against the remaining route", () => RemainingBatch(true)),
        ("retained single click expands only the remaining route", () => RemainingBatch(true, true)),
        ("remaining route respects revision completion and state identity", RouteLifetime),
        ("WASD and arrows reject all immediate opposing directions", Reversals),
        ("WASD batching preserves physical key events", WasdExecution),
        ("click directions can advance a WASD route without rewriting its keys", MixedRoute),
        ("a physical WASD input cannot silently replace an arrow input", PhysicalIdentity)
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    sealed class Responses : HttpMessageHandler
    {
        internal readonly List<string> Bodies = [];
        internal ActionDto[] Actions = [new() { Type = "wait", WaitSeconds = 0 }];
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Bodies.Add(await request.Content!.ReadAsStringAsync(cancellationToken));
            return new(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(new
                {
                    id = "response_visual_" + Bodies.Count, status = "completed",
                    output_text = JsonSerializer.Serialize(new { actions = Actions })
                }))
            };
        }
    }

    static ControlStepState Observe(ControlRunState run, int step)
    {
        var turn = new ControlStepState(step);
        Check(ControlObservationStage.RunAsync(run, turn).GetAwaiter().GetResult() == StepFlow.Next, "observation stopped");
        return turn;
    }

    static void RequestFreshness(bool changed, bool hasActionBaseline)
    {
        using var f = new Fixture();
        A.Options.UsePreviousResponseState = true;
        A.Options.OmitUnchangedScreenImageWithState = true;
        A.Options.ControlContextCompactionEnabled = false;
        A.Options.ScreenSanityChecks = false;
        var run = new ControlRunState("synthetic", "Move the character", AppContext.BaseDirectory, AppContext.BaseDirectory) { cancelCts = f.Run.cancelCts };
        using var handler = new Responses();
        using var client = new HttpClient(handler);
        var oldClient = OpenAiTransport.HttpClient;
        try
        {
            OpenAiTransport.HttpClient = client;
            var first = Observe(run, 1);
            Check(ControlDecisionStage.RunAsync(run, first).GetAwaiter().GetResult() == StepFlow.Next, "first decision failed");
            if (changed) f.DelayedShift = 100;
            var baseline = ScreenshotService.CaptureObservationFrameProbe(true);
            run.lastObservation = run.observationSession.Assess(baseline, baseline, null, f.Context, f.Context, "finite");
            Check(run.lastObservation.IsNoChange, "fixture must start with stale stable evidence");
            run.prevObservationFrame = hasActionBaseline ? baseline : null;
            run.prevShotFingerprint = hasActionBaseline ? baseline.GlobalFingerprint : null;
            run.prevActiveWindowFingerprint = hasActionBaseline ? baseline.ActiveWindowFingerprint : null;
            run.prevObservationContext = hasActionBaseline ? f.Context : null;
            if (!hasActionBaseline)
            {
                var rejected = f.Turn(new() { Type = "click" });
                var flow = ControlExecutionStage.RunAsync(run, rejected).GetAwaiter().GetResult();
                if (flow == StepFlow.Next) ControlCompletionStage.RunAsync(run, rejected).GetAwaiter().GetResult();
                Check(rejected.actionExecutionFailed, "fixture did not reject the invalid click");
            }
            var next = Observe(run, 2);
            if (!hasActionBaseline)
                Check(run.lastObservation is null && double.IsNaN(run.lastDelta), "stale action outcome survived baseline reset");
            Check(ControlDecisionStage.RunAsync(run, next).GetAwaiter().GetResult() == StepFlow.Next, "second decision failed");
            Check(handler.Bodies[1].Contains("previous_response_id"), "test did not continue the model chain");
            Check(handler.Bodies[1].Contains("SCREEN_IMAGE: omitted") == !changed, "request used the wrong visual evidence");
            if (changed) Check(handler.Bodies[1].Contains("input_image"), "fresh screen was not sent");
        }
        finally { OpenAiTransport.HttpClient = oldClient; }
    }

    static ControlStepState Evidence(Fixture f, string image = "image-A")
    {
        var turn = f.Turn(new() { Type = "wait" });
        turn.dataUrl = image;
        turn.previousResponseIdForRequest = "response-A";
        return turn;
    }

    static void TargetFreshness()
    {
        using var f = new Fixture();
        var cache = f.Run.promptVisualState;
        var turn = Evidence(f);
        List<UiaTarget> targets = [new(1, "Continue", "Button", new(10, 10, 40, 20))];
        cache.RecordTargets(turn, targets);
        Check(cache.CanReuseTargets(turn, targets), "unchanged targets not reused");
        Check(!cache.CanReuseTargets(turn, targets.ToList()), "unscanned replacement list reused");
        turn.dataUrl = "image-B";
        Check(!cache.CanReuseTargets(turn, targets), "stale UIA targets reused on changed screen");
        turn.dataUrl = "image-A";
        turn.promptContext = f.Context with { ActiveWindowTitle = "Other window" };
        Check(!cache.CanReuseTargets(turn, targets), "stale UIA targets reused in another context");
    }

    static void ImageContext()
    {
        using var f = new Fixture();
        var turn = Evidence(f);
        var cache = f.Run.promptVisualState;
        cache.RecordResponse(turn, "response-A", true);
        Check(cache.CanOmitImage(turn), "identical evidence was not reusable");
        turn.promptContext = f.Context with { FocusedUiaSummary = "New focus" };
        Check(!cache.CanOmitImage(turn), "focus context was ignored");
        turn.promptContext = f.Context;
        turn.observationFrame = turn.observationFrame with { ScreenBounds = new(320, 0, 320, 200) };
        Check(!cache.CanOmitImage(turn), "desktop coordinate change was ignored");
        turn.observationFrame = turn.observationFrame with { ScreenBounds = new(0, 0, 320, 200) };
        turn.previousResponseIdForRequest = "unrelated-response";
        Check(!cache.CanOmitImage(turn), "image inherited from an unrelated chain");
    }

    static void FailedResponse()
    {
        using var f = new Fixture();
        var turn = Evidence(f);
        var cache = f.Run.promptVisualState;
        cache.RecordResponse(turn, "response-A", true);
        turn.dataUrl = "image-B";
        cache.RecordResponse(turn, null, true);
        Check(!cache.CanOmitImage(turn), "failed request advanced model image");
        turn.dataUrl = "image-A";
        Check(cache.CanOmitImage(turn), "failed request discarded valid preceding model evidence");
    }

    static void RestoredResponse()
    {
        using var f = new Fixture();
        var turn = Evidence(f);
        var cache = f.Run.promptVisualState;
        cache.RecordResponse(turn, "response-A", true);
        cache.RecordResponse(turn, "response-B", false);
        turn.previousResponseIdForRequest = "response-B";
        Check(cache.CanOmitImage(turn), "chained omission lost inherited image");
        turn.dataUrl = "restored-image";
        cache.RecordResponse(turn, "restored-response", true);
        turn.previousResponseIdForRequest = "restored-response";
        Check(cache.CanOmitImage(turn), "standalone restoration was not recorded");
    }

    static ActionDto Plan(string type = "keys") => new()
    {
        Type = type, Keys = type == "keys" ? ["right"] : null, XPx = 100, YPx = 100,
        PlanStatus = "active", ShortTermPlan = "Up, then right twice to the exit",
        PlannedInputs = ["Up", "Right", "Right"], PlanWaypoint = "exit", PlanConfidence = .99
    };

    static void RemainingBatch(bool clicks, bool expand = false)
    {
        using var f = new Fixture(turnBased: true) { Moving = true };
        A.Options.ActionRepeatCooldownSteps = 0;
        A.Options.ScreenSanityChecks = false;
        var turn = Observe(f.Run, 2);
        var region = new Rectangle(0, 0, 320, 200);
        f.Run.turnBasedInteractionRect = region;
        f.Run.turnBasedTransitions.ObserveState(turn.observationFrame, region);
        var state = f.Run.turnBasedTransitions.CurrentStateId!;
        var tracker = f.Run.shortTermPlan;
        var first = Plan(clicks ? "click" : "keys");
        tracker.Update(first, "prior-state");
        tracker.RecordDirectionalResult("up", "prior-state", state, true);
        var second = new ActionDto { Type = first.Type, Keys = first.Keys, XPx = 100, YPx = 100 };
        using var handler = new Responses { Actions = expand ? [first] : [first, second] };
        using var client = new HttpClient(handler);
        var oldClient = OpenAiTransport.HttpClient;
        try
        {
            OpenAiTransport.HttpClient = client;
            Check(ControlDecisionStage.RunAsync(f.Run, turn).GetAwaiter().GetResult() == StepFlow.Next, "route decision failed");
            var remaining = turn.plannedInputsForExecution!;
            Check(remaining.SequenceEqual(new[] { "ArrowRight", "ArrowRight" }), "route cursor reset");
            var decision = turn.decision;
            Check(decision.FollowUps.Count == (expand ? 0 : 1), "retained route has incorrect follow-ups");
            if (!expand)
            {
                Check(turn.action.ResolvedTurnInputLabel == "ArrowRight", "first continuation mislabeled as Up");
                Check(ControlBatchService.TryTakeQueuedObservedTurnActions(new(decision.FollowUps),
                    DesktopInputService.CaptureResolvedAction(turn.action, null), region, out var actions, remaining), "executor rebound the route from its origin");
                Check(actions.Count == 2 && actions.All(a => a.Action.ResolvedTurnInputLabel == "ArrowRight"), "executor corrupted direction labels");
            }
            f.Execute(turn);
            Check(f.Run.recentExecutedActions.Count == 2, "production execution lost the remaining route batch");
            Check(f.Run.recentExecutedActions.All(a => a.Action.ResolvedTurnInputLabel == "ArrowRight"), "execution mislabeled a remaining move");
            Check(BatchPlanner.SafeObservedTurnFollowUps([first, second], 4, []).Length == 0, "empty route fell back to stale model payload");
        }
        finally { OpenAiTransport.HttpClient = oldClient; }
    }

    static void RouteLifetime()
    {
        var tracker = new ShortTermPlanTracker();
        var action = Plan();
        tracker.Update(action, "S1");
        tracker.RecordDirectionalResult("up", "S1", "S2", true);
        Check(tracker.RemainingInputs("unrelated").Count == 0, "stale state retained an executable route");
        action.PlanStatus = "revised";
        action.PlannedInputs = ["left"];
        Check(tracker.PrepareInputsForProposal(action, "S2").SequenceEqual(new[] { "ArrowLeft" }), "revision retained old cursor");
        tracker.RecordDirectionalResult("left", "S2", "S3", true);
        Check(tracker.RemainingInputs("S3").Count == 0 && tracker.Status == "completed", "completed route remains executable");
        action.PlanStatus = "none";
        Check(tracker.PrepareInputsForProposal(action, "S3").Count == 0, "cleared plan revived raw inputs");
    }

    static void Reversals()
    {
        foreach (var pair in new[] { ("w", "s"), ("a", "d"), ("s", "w"), ("d", "a"), ("W", "ArrowDown"), ("left", "D") })
        {
            Check(BatchPlanner.HasImmediateDirectionalReversal([pair.Item1, pair.Item2]), "reversal accepted: " + pair);
            var tracker = new ShortTermPlanTracker();
            var action = Plan();
            action.PlannedInputs = [pair.Item1, pair.Item2];
            tracker.Update(action, "S1");
            Check(!tracker.HasActivePlan, "incoherent WASD plan accepted");
        }
        Check(!BatchPlanner.HasImmediateDirectionalReversal(["W", "d", "D"]), "valid turn was rejected");
    }

    static void WasdExecution()
    {
        using var f = new Fixture(turnBased: true) { Moving = true };
        A.Options.ActionRepeatCooldownSteps = 0;
        f.Run.turnBasedInteractionRect = new(0, 0, 320, 200);
        var first = new ActionDto { Type = "keys", Keys = ["w"], PlannedInputs = ["W", "W"], PlanConfidence = .99 };
        var turn = f.Turn(first, new() { Type = "keys", Keys = ["W"] }, 4);
        Check(turn.decision.FollowUps.Count == 1, "WASD follow-up was discarded");
        f.Execute(turn);
        Check(f.Input.Events.Count == 4 && f.Input.Events.All(e => e.U.ki.wVk == 'W'), "WASD was dropped or replaced by arrow events");
        Check(first.ResolvedTurnInputLabel == "W", "lowercase input lost physical identity");
    }

    static void MixedRoute()
    {
        var tracker = new ShortTermPlanTracker();
        var action = Plan();
        action.PlannedInputs = ["w", "d"];
        tracker.Update(action, "S1");
        tracker.RecordDirectionalResult("ArrowUp", "S1", "S2", true);
        Check(tracker.RemainingInputs("S2").SequenceEqual(new[] { "D" }), "click failed to advance WASD plan");
        Check(tracker.ProposedSequenceMatches(["d"], "S2", out _), "lowercase WASD identity differs");
    }

    static void PhysicalIdentity()
    {
        var action = new ActionDto { Type = "keys", Keys = ["w"] };
        Check(!BatchPlanner.TryBindObservedTurnActionToPlannedInput(action, "ArrowUp"), "physical key substitution accepted");
        Check(BatchPlanner.TryBindObservedTurnActionToPlannedInput(action, "W") && action.Keys![0] == "w", "canonicalization rewrote outgoing key");
        foreach (var pair in new[] { ("w", "ArrowUp"), ("a", "ArrowLeft"), ("s", "ArrowDown"), ("d", "ArrowRight") })
            Check(BatchPlanner.TryNormalizeDirectionalLabel(pair.Item1, out var label) && label == pair.Item2, "wrong logical direction");
    }
}
