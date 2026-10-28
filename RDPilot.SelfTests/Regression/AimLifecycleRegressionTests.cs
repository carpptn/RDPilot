using System.Drawing;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class AimLifecycleRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("expired AIM permits re-aim and guarded click", ExpiredAim),
        ("stable AIM still blocks duplicate aiming and inspection", StableAim),
        ("window move invalidates AIM and permits re-aim", () => ChangedGeometry(false)),
        ("window resize invalidates AIM and permits re-aim", () => ChangedGeometry(true)),
        ("external epoch clears AIM and inspection evidence", ExternalEpoch),
        ("AIM expiration preserves the inspection budget", ExpirationPreservesBudget),
        ("successful interaction permits a new AIM while retaining click context", InteractionUnlocksAim),
        ("AIM invalidation preserves an unrelated requested crop", IndependentCrop)
    ];

    static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
    }

    static ActionDto Aim(int left = 10) => new()
    {
        Type = "aim", BBox = new() { Left = left, Top = 10, Right = left + 90, Bottom = 100 }
    };

    static ActionDto Crop() => new()
    {
        Type = "request_crop", Crop = new() { Left = 10, Top = 10, Right = 150, Bottom = 150 }
    };

    static void Configure()
    {
        A.Options.DirectClickWithoutAim = false;
        A.Options.AimExpireDelta = .08;
        A.Options.ScreenSanityChecks = false;
        A.Options.MaxArtifactsPerDir = 0;
        A.Options.RecoveryProgressVerificationEnabled = false;
        A.Options.MaxConsecutiveInspectionActions = 1;
    }

    static void Execute(SyntheticControlFixture fixture, ActionDto action)
    {
        var turn = fixture.Turn(action);
        var flow = fixture.Execute(turn);
        Check(flow == StepFlow.Next && turn.actionExecuted && !turn.actionExecutionFailed,
            $"{action.Type} failed: {fixture.Run.lastExecutorFailure}");
    }

    sealed class ChangedScreen : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multiple) => new(0, 0, 320, 200);
        public void CopyPixels(Graphics graphics, Rectangle region) => graphics.Clear(Color.White);
    }

    static void Expire(SyntheticControlFixture fixture)
    {
        DesktopCapture.Backend = new ChangedScreen();
        var turn = new ControlStepState(2);
        Check(ControlObservationStage.RunAsync(fixture.Run, turn).GetAwaiter().GetResult() == StepFlow.Next,
            "observation stopped unexpectedly");
        Check(fixture.Run.lastAimRect is null && !fixture.Run.observationActionGuard.AimIssuedSinceInteraction,
            "expired AIM retained its rectangle or pending interaction lock");
        Check(fixture.Run.nextFocusRect is null, "expired AIM left a queued focus crop");
    }

    static void ExpiredAim()
    {
        using var f = new SyntheticControlFixture();
        Configure();
        Execute(f, Aim());
        Expire(f);
        Execute(f, Aim(110));
        Execute(f, new() { Type = "click", XPx = 150, YPx = 50 });
        Check(f.Input.Sends > 0, "guarded click was never sent to the synthetic backend");
    }

    static void StableAim()
    {
        using var f = new SyntheticControlFixture();
        Configure();
        Execute(f, Aim());
        var original = f.Run.lastAimRect;
        var observation = new ControlStepState(2);
        Check(ControlObservationStage.RunAsync(f.Run, observation).GetAwaiter().GetResult() == StepFlow.Next,
            "stable observation stopped");
        foreach (var action in new[] { Aim(110), Crop() })
        {
            var turn = f.Turn(action);
            _ = f.Execute(turn);
            Check(!turn.actionExecuted && turn.actionExecutionFailed && f.Run.lastAimRect == original &&
                f.Run.observationActionGuard.AimIssuedSinceInteraction, "stable AIM was replaced or its guard bypassed");
        }
        Check(f.Input.Sends == 0, "rejected actions generated input");
    }

    static void ChangedGeometry(bool resize)
    {
        using var f = new SyntheticControlFixture(turnBased: true);
        Configure();
        var region = new Rectangle(10, 10, 250, 150);
        ScreenshotService.ActiveWindowRectangleProvider = () => region;
        f.Run.turnBasedInteractionRect = region;
        f.Run.turnBasedInteractionContext = TurnInteractionContext.Capture(f.Context);
        Execute(f, Aim());
        ScreenshotService.ActiveWindowRectangleProvider = () => resize
            ? new Rectangle(10, 10, 280, 180) : new Rectangle(20, 20, 250, 150);
        Check(ControlObservationStage.RunAsync(f.Run, new(2)).GetAwaiter().GetResult() == StepFlow.Next,
            "geometry observation stopped");
        Check(f.Run.stateEpoch == 1 && f.Run.lastAimRect is null &&
            !f.Run.observationActionGuard.AimIssuedSinceInteraction, "geometry reset retained AIM state");
        Execute(f, Aim(110));
    }

    static void ExternalEpoch()
    {
        using var f = new SyntheticControlFixture();
        Configure();
        Execute(f, Crop());
        Execute(f, Aim());
        Check(f.Run.observationActionGuard.RequiresInteraction(1), "fixture has no inspection budget to reset");
        f.Run.BeginExternalStateEpoch();
        Check(f.Run.lastAimRect is null && f.Run.nextFocusRect is null &&
            !f.Run.observationActionGuard.AimIssuedSinceInteraction &&
            f.Run.observationActionGuard.ConsecutiveInspectionActions == 0, "epoch retained visual inspection state");
        Execute(f, Crop()); // The same region belongs to a new external state.
        Execute(f, Aim());
    }

    static void ExpirationPreservesBudget()
    {
        using var f = new SyntheticControlFixture();
        Configure();
        Execute(f, Crop());
        Execute(f, Aim());
        Expire(f);
        Check(f.Run.observationActionGuard.RequiresInteraction(1), "expiration reset the inspection budget");
        var crop = f.Turn(Crop());
        _ = f.Execute(crop);
        Check(!crop.actionExecuted && crop.actionExecutionFailed, "expiration bypassed duplicate crop protection");
        Execute(f, Aim());
    }

    static void InteractionUnlocksAim()
    {
        using var f = new SyntheticControlFixture();
        Configure();
        Execute(f, Aim());
        var original = f.Run.lastAimRect;
        Execute(f, new() { Type = "click", XPx = 50, YPx = 50 });
        Check(f.Run.lastAimRect == original && !f.Run.observationActionGuard.AimIssuedSinceInteraction,
            "interaction lost geometric click context or retained the re-aim lock");
        Execute(f, Aim(110));
    }

    static void IndependentCrop()
    {
        using var f = new SyntheticControlFixture();
        Configure();
        Execute(f, Aim());
        var crop = new Rectangle(180, 110, 60, 60);
        f.Run.nextFocusRect = crop;
        f.Run.InvalidateAim();
        Check(f.Run.lastAimRect is null && !f.Run.observationActionGuard.AimIssuedSinceInteraction &&
            f.Run.nextFocusRect == crop, "AIM invalidation removed an unrelated requested crop");
    }
}
