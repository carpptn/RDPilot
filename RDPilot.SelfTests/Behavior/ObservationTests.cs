using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static void TestObservationActionGuard()
    {
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(1920, 1080, 640, 360);
        var guard = new RDPilot.Control.ObservationActionGuardState();
        ResolvedActionSnapshot Snapshot(ActionDto action) => RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(action, null);
        var boardCrop = Snapshot(new ActionDto { Type = "request_crop", Crop = new BBox { Left = 230, Top = 90, Right = 405, Bottom = 220 } });
        var controlsCrop = Snapshot(new ActionDto { Type = "request_crop", Crop = new BBox { Left = 230, Top = 215, Right = 410, Bottom = 310 } });
        var thirdCrop = Snapshot(new ActionDto { Type = "request_crop", Crop = new BBox { Left = 30, Top = 40, Right = 150, Bottom = 150 } });
        Assert(!guard.TryGetBlockReason(boardCrop, 2, out _), "first inspection was blocked");
        guard.RecordExecuted(boardCrop);
        Assert(guard.TryGetBlockReason(boardCrop, 2, out var duplicateReason) && duplicateReason.Contains("already inspected", StringComparison.Ordinal), "repeated inspection of the same region was not blocked");
        Assert(guard.TryGetBlockReason(controlsCrop, 2, singleInspectionBeforeInteraction: true, out var nestedTurnCropReason) && nestedTurnCropReason.Contains("already supplied", StringComparison.Ordinal), "a second nested crop was allowed in the same turn-based state");
        Assert(!guard.TryGetBlockReason(controlsCrop, 2, out _), "second distinct inspection was blocked");
        guard.RecordExecuted(controlsCrop);
        Assert(guard.TryGetBlockReason(thirdCrop, 2, out var budgetReason) && budgetReason.Contains("limit of 2", StringComparison.Ordinal), "inspection budget did not block an A-B-A style observation loop");
        var aim = Snapshot(new ActionDto { Type = "aim", BBox = new BBox { Left = 250, Top = 230, Right = 300, Bottom = 280 } });
        Assert(!guard.TryGetBlockReason(aim, 2, out _), "precision AIM was blocked after inspection");
        guard.SetAim(RDPilot.Desktop.DesktopInputService.ResolveAimRect(aim.Action)!.Value);
        Assert(guard.TryGetBlockReason(aim, 2, out var aimReason) && aimReason.Contains("already active", StringComparison.Ordinal), "repeated AIM was not blocked");
        guard.RecordExecuted(Snapshot(new ActionDto { Type = "wait", WaitSeconds = 1 }));
        Assert(guard.RequiresInteraction(2), "wait incorrectly reset the inspection budget");
        guard.RecordExecuted(Snapshot(new ActionDto { Type = "keys", Keys = ["left"] }));
        Assert(!guard.RequiresInteraction(2) && !guard.AimIssuedSinceInteraction, "state-changing input did not reset the inspection budget");
        Assert(!guard.TryGetBlockReason(thirdCrop, 2, out _), "new inspection remained blocked after interaction");
    }

    static void TestAdaptiveLocalObservation()
    {
        RDPilot.Configuration.ConfigurationService.ApplyObservationMode("auto", "self-test");
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
        var action = new ActionDto
        {
            Type = "drag_path",
            GestureKind = "draw",
            DurationMs = 500,
            Path = [new GesturePointDto
            {
                XPx = 10,
                YPx = 50
            }, new GesturePointDto
            {
                XPx = 90,
                YPx = 50
            }

            ]
        };
        var snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(action, null);
        var beforeDetail = Enumerable.Repeat((byte)255, 10_000).ToArray();
        var afterDetail = beforeDetail.ToArray();
        for (var x = 10; x <= 90; x++)
            afterDetail[50 * 100 + x] = 0;
        var stable = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
        var before = new ScreenObservationFrame(stable, stable, beforeDetail, 100, 100, new Rectangle(0, 0, 100, 100));
        var after = new ScreenObservationFrame(stable.ToArray(), stable.ToArray(), afterDetail, 100, 100, new Rectangle(0, 0, 100, 100));
        var session = new RDPilot.Observation.AdaptiveObservationSession();
        var assessment = session.Assess(before, after, snapshot, new UiPromptContext("Paint", "mspaint", "canvas", null, null, null), new UiPromptContext("Paint", "mspaint", "canvas", null, null, null), "finite");
        Assert(assessment.Profile == "local_editing", "draw path did not select local_editing");
        Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed, "thin local path change was not confirmed");
        Assert(assessment.GoalProgress == GoalProgressState.Progress, "confirmed path was not treated as progress");
        var click = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "click", XPx = 50, YPx = 50 }, null);
        var coarseOnly = new ScreenObservationFrame(stable, stable, [], 0, 0, new Rectangle(0, 0, 100, 100));
        var clickAssessment = session.Assess(coarseOnly, coarseOnly, click, new UiPromptContext("Paint", "mspaint", "canvas", null, null, null), new UiPromptContext("Paint", "mspaint", "canvas", null, null, null), "finite");
        Assert(clickAssessment.ActionPolicy == "static_ui", "a discrete click inherited the prior drawing policy");
    }

    static void TestAdaptiveColorObservation()
    {
        RDPilot.Configuration.ConfigurationService.ApplyObservationMode("auto", "self-test");
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
        var snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "drag_path", GestureKind = "draw", DurationMs = 500, Path = [new GesturePointDto { XPx = 10, YPx = 50 }, new GesturePointDto { XPx = 90, YPx = 50 }] }, null);
        var grayscale = Enumerable.Repeat((byte)128, 10_000).ToArray();
        var beforeColor = Enumerable.Repeat((byte)128, 30_000).ToArray();
        var afterColor = beforeColor.ToArray();
        for (var x = 10; x <= 90; x++)
        {
            var index = (50 * 100 + x) * 3;
            afterColor[index] = 255;
            afterColor[index + 1] = 32;
            afterColor[index + 2] = 128;
        }

        var global = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
        var before = new ScreenObservationFrame(global, global, grayscale, 100, 100, new Rectangle(0, 0, 100, 100))
        {
            DetailColorFingerprint = beforeColor
        };
        var after = new ScreenObservationFrame(global.ToArray(), global.ToArray(), grayscale.ToArray(), 100, 100, new Rectangle(0, 0, 100, 100))
        {
            DetailColorFingerprint = afterColor
        };
        var context = new UiPromptContext("Paint", "mspaint", "canvas", null, null, null);
        var assessment = new RDPilot.Observation.AdaptiveObservationSession().Assess(before, after, snapshot, context, context, "finite");
        Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed, "color-only path change was missed");
    }

    static void TestFocusedTextObservation()
    {
        RDPilot.Configuration.ConfigurationService.ApplyObservationMode("auto", "self-test");
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(100, 100, 100, 100);
        var snapshot = RDPilot.Observation.ControlObservationService.AttachFocusedTextObservationRegion(RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "type_text", Text = "1100" }, null), new Rectangle(35, 40, 30, 12));
        Assert(snapshot.ObservationRegion is not null, "focused text action did not receive a local observation region");
        Assert(RDPilot.Control.ControlActionPolicy.ShouldDeferFocusedTextStagnation(snapshot, true, 0), "first visually ambiguous focused text edit does not receive a grace attempt");
        Assert(!RDPilot.Control.ControlActionPolicy.ShouldDeferFocusedTextStagnation(snapshot, true, 1), "repeated ineffective text edits receive unlimited stagnation grace");
        var beforeDetail = Enumerable.Repeat((byte)230, 10_000).ToArray();
        var afterDetail = beforeDetail.ToArray();
        for (var y = 44; y < 49; y++)
            for (var x = 42; x < 58; x++)
                afterDetail[y * 100 + x] = 20;
        var stable = Enumerable.Repeat((byte)128, 96 * 54).ToArray();
        var before = new ScreenObservationFrame(stable, stable, beforeDetail, 100, 100, new Rectangle(0, 0, 100, 100));
        var after = new ScreenObservationFrame(stable.ToArray(), stable.ToArray(), afterDetail, 100, 100, new Rectangle(0, 0, 100, 100));
        var context = new UiPromptContext("Properties", "mspaint", "textbox edit", null, null, null);
        var assessment = new RDPilot.Observation.AdaptiveObservationSession().Assess(before, after, snapshot, context, context, "finite");
        Assert(assessment.ActionPolicy == "local_editing", "focused text edit did not select local observation");
        Assert(assessment.ActionOutcome == ActionOutcomeState.Confirmed && assessment.GoalProgress == GoalProgressState.Progress, "small focused text change was treated as stagnation");
    }

    static void TestRealtimeAmbientMotion()
    {
        RDPilot.Configuration.ConfigurationService.ApplyObservationMode("auto", "self-test");
        var zero = new byte[96 * 54];
        var four = Enumerable.Repeat((byte)4, 96 * 54).ToArray();
        var eight = Enumerable.Repeat((byte)8, 96 * 54).ToArray();
        var twelve = Enumerable.Repeat((byte)12, 96 * 54).ToArray();
        var detail = new byte[100];
        ScreenObservationFrame Frame(byte[] fingerprint) => new(fingerprint, fingerprint, detail, 10, 10, new Rectangle(0, 0, 100, 100));
        var action = new ResolvedActionSnapshot(new ActionDto { Type = "keys", Keys = ["space"] }, "keys [space]", "keys:space", null);
        var context = new UiPromptContext("Game", "game", "animated canvas", null, null, null);
        var transientSession = new RDPilot.Observation.AdaptiveObservationSession();
        transientSession.RecordAmbientMotion(Frame(zero), Frame(four));
        var transientAssessment = transientSession.Assess(Frame(four), Frame(eight), action, context, context, "finite");
        Assert(transientAssessment.Profile != "realtime_interaction", "a single transient animation immediately selected realtime observation");
        var session = new RDPilot.Observation.AdaptiveObservationSession();
        session.RecordAmbientMotion(Frame(zero), Frame(four));
        session.RecordAmbientMotion(Frame(four), Frame(eight));
        var assessment = session.Assess(Frame(eight), Frame(twelve), action, context, context, "finite");
        Assert(assessment.Profile == "realtime_interaction", "ambient motion did not select realtime observation");
        Assert(assessment.GoalProgress != GoalProgressState.Progress, "ambient animation was treated as goal progress");
        var shellSession = new RDPilot.Observation.AdaptiveObservationSession();
        shellSession.RecordAmbientMotion(Frame(zero), Frame(four));
        shellSession.RecordAmbientMotion(Frame(four), Frame(eight));
        var shellAction = new ResolvedActionSnapshot(new ActionDto { Type = "keys", Keys = ["enter"], Note = "open selected application" }, "keys [enter]", "keys:enter", null);
        var shellContext = new UiPromptContext("Search", "explorer", "search results", null, null, null);
        var shellAssessment = shellSession.Assess(Frame(eight), Frame(twelve), shellAction, shellContext, shellContext, "finite");
        Assert(shellAssessment.Profile != "realtime_interaction", "animated shell UI incorrectly selected realtime observation");
        Assert(shellAssessment.ActionPolicy == "static_ui", "discrete shell action lost settle observation while the UI was animated");
    }
}
