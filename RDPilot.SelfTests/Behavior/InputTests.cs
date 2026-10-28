using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static void TestInvalidBoundingBox()
    {
        var action = new ActionDto
        {
            Type = "click",
            BBox = new BBox
            {
                Left = 100,
                Top = 100,
                Right = 50,
                Bottom = 50
            }
        };
        var snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(action, null);
        Assert(!snapshot.IsValid, "inverted bbox must be invalid");
        var noOpDrag = new ActionDto
        {
            Type = "drag_drop",
            XPx = 100,
            YPx = 100,
            ToXPx = 101,
            ToYPx = 101
        };
        var dragSnapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(noOpDrag, null);
        Assert(!dragSnapshot.IsValid, "effectively zero-distance drag must be invalid");
    }

    static void TestCoordinateBounds()
    {
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(1920, 1080, 1280, 720);
        var pixel = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "click", XPx = 1280, YPx = 10 }, null);
        Assert(!pixel.IsValid, "pixel outside the screenshot was accepted");
        var normalized = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "move", X = double.NaN, Y = 0.5 }, null);
        Assert(!normalized.IsValid, "non-finite normalized coordinate was accepted");
    }

    static void TestCoordinateMapping()
    {
        var mapper = ScreenCoordinateMapper.Create(1920, 1080, 1280, 720);
        Assert(mapper.ImageToScreenPoint(0, 0) == (0, 0), "origin changed");
        Assert(mapper.ImageToScreenPoint(1279, 719) == (1919, 1079), "bottom-right is not inclusive");
        Assert(mapper.ScreenToImagePoint(1919, 1079) == (1279, 719), "round trip failed");
    }

    static void TestVirtualDesktopCoordinateMapping()
    {
        var mapper = ScreenCoordinateMapper.Create(-1920, -200, 4480, 1640, 2240, 820);
        Assert(mapper.ImageToScreenPoint(0, 0) == (-1920, -200), "virtual desktop origin was lost");
        Assert(mapper.ImageToScreenPoint(2239, 819) == (2559, 1439), "virtual desktop bottom-right was not mapped inclusively");
        Assert(mapper.ScreenToImagePoint(-1920, -200) == (0, 0), "negative screen origin did not round-trip");
    }

    static void TestGestureActionSchema()
    {
        var types = RDPilot.OpenAI.PromptAndRequestFactory.ControlActionTypes("finite");
        Assert(types.Contains("drag_path", StringComparer.OrdinalIgnoreCase), "drag_path is missing from the action schema");
        Assert(types.Contains("hold_keys", StringComparer.OrdinalIgnoreCase), "hold_keys is missing from the action schema");
        Assert(RDPilot.Control.ControlBatchService.IsSafeBatchedAction(new ActionDto { Type = "drag_path" }), "bounded draw batches are missing from the batch action gate");
        Assert(!RDPilot.Control.ControlBatchService.IsSafeBatchedAction(new ActionDto { Type = "hold_keys" }), "hold_keys must not be batchable");
    }

    static void TestDragPathMapping()
    {
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(1920, 1080, 960, 540);
        var action = new ActionDto
        {
            Type = "drag_path",
            GestureKind = "draw",
            DurationMs = 600,
            Path = [new GesturePointDto
            {
                XPx = 100,
                YPx = 100
            }, new GesturePointDto
            {
                XPx = 400,
                YPx = 200
            }, new GesturePointDto
            {
                XPx = 800,
                YPx = 400
            }

            ]
        };
        var snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(action, null);
        Assert(snapshot.IsValid, snapshot.ValidationError ?? "valid path was rejected");
        Assert(snapshot.ScreenPath.Count == 3, "resolved path lost points");
        Assert(snapshot.ScreenPath[0] == new Point(200, 200), "path start was not mapped to the desktop");
        Assert(snapshot.ObservationRegion is not null, "path did not capture an observation footprint");
        action.Path[1].XPx = 960;
        snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(action, null);
        Assert(!snapshot.IsValid, "out-of-range path point was accepted");
    }

    static void TestDragPathInputPlan()
    {
        Point[] path = [new Point(10, 20), new Point(40, 20), new Point(40, 50), new Point(10, 20)];
        var plan = RDPilot.Desktop.DesktopInputService.BuildMouseDragPlan(path, 600);
        Assert(plan.Moves.Length >= path.Length - 1, "drag plan lost movement events");
        Assert(plan.EffectiveDurationMs == 600, "drag plan changed the requested duration");
        Assert(plan.InitialHoldMs is >= 20 and <= 80, "drag plan produced an invalid initial hold");
        Assert(plan.Moves[^1] == path[^1], "drag plan did not finish at the final path point");
        foreach (var waypoint in path.Skip(1))
            Assert(plan.Moves.Contains(waypoint), $"drag plan skipped waypoint {waypoint}");
        const uint expectedFlags = 0x0001 | 0x2000 | 0x4000 | 0x8000;
        var topLeft = RDPilot.Desktop.DesktopInputService.BuildAbsoluteMouseMoveData(-1920, 0, -1920, 0, 3840, 1080);
        var bottomRight = RDPilot.Desktop.DesktopInputService.BuildAbsoluteMouseMoveData(1919, 1079, -1920, 0, 3840, 1080);
        Assert(topLeft.Dx == 0 && topLeft.Dy == 0, "virtual-desktop origin was normalized incorrectly");
        Assert(bottomRight.Dx == 65535 && bottomRight.Dy == 65535, "virtual-desktop edge was normalized incorrectly");
        Assert(topLeft.Flags == expectedFlags, "absolute drag movement is missing required SendInput flags");
        var firstDeadline = RDPilot.Desktop.DesktopInputService.GestureMoveTargetElapsedMs(plan, 0);
        var finalDeadline = RDPilot.Desktop.DesktopInputService.GestureMoveTargetElapsedMs(plan, plan.Moves.Length - 1);
        Assert(firstDeadline >= plan.InitialHoldMs, "first drag movement precedes the initial button hold");
        Assert(Math.Abs(finalDeadline - plan.EffectiveDurationMs) < 0.001, "drag schedule accumulates per-move sleep instead of ending at the absolute deadline");
    }

    static void TestHoldKeysValidation()
    {
        var valid = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "hold_keys", Keys = ["w", "d"], DurationMs = 500 }, null);
        Assert(valid.IsValid, valid.ValidationError ?? "valid key hold was rejected");
        var unbounded = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "hold_keys", Keys = ["w"], DurationMs = 60_000 }, null);
        Assert(!unbounded.IsValid, "unbounded key hold was accepted");
        var chord = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "hold_keys", Keys = ["shift+w"], DurationMs = 500 }, null);
        Assert(!chord.IsValid, "chord string was accepted as an individual held key");
        var systemKey = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "hold_keys", Keys = ["win"], DurationMs = 500 }, null);
        Assert(!systemKey.IsValid, "unsafe system key was accepted for holding");
        Assert(RDPilot.Desktop.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(["win+down"], ownConsoleForeground: true) && RDPilot.Desktop.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(["win", "ArrowDown"], ownConsoleForeground: true), "Win+Down aimed at RDPilot's own console was not routed through its window handle");
        Assert(!RDPilot.Desktop.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(["win+down"], ownConsoleForeground: false) && !RDPilot.Desktop.DesktopInputService.ShouldMinimizeOwnConsoleByHandle(["win+left"], ownConsoleForeground: true), "the own-console safeguard blocked a legitimate system window shortcut");
    }

    static void TestPathGestureLoopClassification()
    {
        RDPilot.Desktop.DesktopInputService.SetCurrentScreenMap(800, 600, 800, 600);
        var snapshot = RDPilot.Desktop.DesktopInputService.CaptureResolvedAction(new ActionDto { Type = "drag_path", GestureKind = "draw", DurationMs = 500, Note = "draw a cat outline on canvas", Path = [new GesturePointDto { XPx = 100, YPx = 100 }, new GesturePointDto { XPx = 200, YPx = 200 }] }, null);
        var kind = RDPilot.Persistence.RecoveryMemoryService.ClassifyLoopKind([snapshot], snapshot, new UiPromptContext("Paint", "mspaint", "canvas drawing surface", null, null, null));
        Assert(kind == RDPilot.Persistence.RecoveryMemoryService.RasterCanvasPointerLoop, $"draw path was classified as {kind}");
    }
}
