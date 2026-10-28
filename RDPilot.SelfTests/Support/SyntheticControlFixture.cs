using System.Drawing;
using A = RDPilot.RDPilotApplication;
using Input = RDPilot.Desktop.DesktopInputService;

namespace RDPilot.SelfTests.Support;

internal sealed class SyntheticControlFixture : IDisposable
{
    internal sealed class Inputs : Input.IInputBackend
    {
        internal readonly List<Input.INPUT> Events = [];
        internal int Sends;
        public uint Send(Input.INPUT[] events) { Sends++; Events.AddRange(events); return (uint)events.Length; }
        public int LastError => 5;
        public bool SetCursorPosition(int x, int y) => true; // No native cursor movement.
        public void SetClipboardText(string text) => throw new InvalidOperationException("Native clipboard prohibited");
    }

    sealed class Capture(Func<int> position) : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multiple) => new(0, 0, 320, 200);
        public void CopyPixels(Graphics graphics, Rectangle region)
        {
            graphics.Clear(Color.Navy);
            graphics.FillRectangle(Brushes.White, position(), 20, 40, 40);
        }
    }

    readonly IDisposable scope = A.BeginRun();
    readonly Input.IInputBackend oldInput = DesktopInputService.InputBackend;
    readonly DesktopCapture.IBackend oldCapture = DesktopCapture.Backend;
    readonly Func<string?, int, int, UiPromptContext> oldContext = UiAutomationService.PromptContextProvider;
    readonly Func<Rectangle?> oldWindow = ScreenshotService.ActiveWindowRectangleProvider;
    internal readonly Inputs Input = new();
    internal readonly ControlRunState Run;
    internal bool Moving;
    internal int DelayedShift;
    internal readonly UiPromptContext Context = new("Game", "game", null, null, null, null);

    internal SyntheticControlFixture(bool turnBased = false)
    {
        A.Options.RecoveryMemoryEnabled = false;
        A.Options.LogScreens = A.Options.LogRequests = false;
        A.Options.IncludeUiaTargets = A.Options.IncludeFocusUia = false;
        A.Options.ScreenPollingEnabled = false;
        A.Options.ScreenPollTimeoutMs = A.Options.UiSettleDelayMs = A.Options.WaitNoChangeExtraMs = 0;
        A.Options.ClipboardPasteThreshold = 10000;
        A.Options.ObservationMode = turnBased ? "turn_based_interaction" : "static_ui";
        A.Options.ExecuteMultiActionCandidates = true;
        A.Options.MaxQueuedBatchActions = 4;
        A.Options.MaxActionFailuresBeforeAbort = 3;
        A.Options.ActionRepeatCooldownSteps = 2;
        A.CurrentRun.CurrentScreenMap = ScreenCoordinateMapper.Create(320, 200, 320, 200);
        DesktopInputService.InputBackend = Input;
        UiAutomationService.PromptContextProvider = (_, _, _) => Context;
        ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(0, 0, 320, 200);
        DesktopCapture.Backend = new Capture(() => 20 + (Moving && Input.Sends > 0 ? 10 : 0) + DelayedShift);
        Run = new("synthetic", "Move the character", AppContext.BaseDirectory, AppContext.BaseDirectory);
        Run.cancelCts = new CancellationMonitor(A.CurrentRun, () => false);
    }

    internal ControlStepState Turn(ActionDto first, ActionDto? second = null, int observedLimit = 0) => new(1)
    {
        Decision = BatchPlanner.CreateDecision(second is null ? [first] : [first, second], observedTurnBatchLimit: observedLimit),
        promptContext = Context,
        observationFrame = ScreenshotService.CaptureObservationFrameProbe(true)
    };

    internal StepFlow Execute(ControlStepState turn)
    {
        var flow = ControlExecutionStage.RunAsync(Run, turn).GetAwaiter().GetResult();
        return flow == StepFlow.Next ? ControlCompletionStage.RunAsync(Run, turn).GetAwaiter().GetResult() : flow;
    }

    public void Dispose()
    {
        Run.cancelCts!.DisposeAsync().AsTask().GetAwaiter().GetResult();
        DesktopInputService.InputBackend = oldInput;
        DesktopCapture.Backend = oldCapture;
        UiAutomationService.PromptContextProvider = oldContext;
        ScreenshotService.ActiveWindowRectangleProvider = oldWindow;
        scope.Dispose();
    }
}

