using System.Diagnostics;
using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;
using Fixture = RDPilot.SelfTests.Support.SyntheticControlFixture;

namespace RDPilot.SelfTests.Regression;

internal static class GeometryMaintenanceRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("pixel control templates survive image resizing", () => ControlScale("pixel")),
        ("normalized control templates survive image resizing", () => ControlScale("normalized")),
        ("bbox control templates survive image resizing", () => ControlScale("bbox")),
        ("historical click evidence retains its original physical coordinates", HistoricalControl),
        ("control recognition uses the current image scale", ControlRecognition),
        ("moved windows invalidate learned controls", ControlWindow),
        ("moving a window resets the observation region before taking a crop", () => WindowChange(false)),
        ("resizing a window resets the observation region before taking a crop", () => WindowChange(true)),
        ("image resizing alone preserves the interaction context", StableGeometry),
        ("window motion while planning blocks stale input", MovedDuringPlanning),
        ("QA normalized locations include the desktop origin", () => QaLocation("normalized")),
        ("QA pixel locations map from the reduced image to the desktop", () => QaLocation("pixel")),
        ("QA boxes label image and desktop coordinates separately", () => QaLocation("bbox")),
        ("memory prune CLI reports corrupt input as failure", () => MaintenanceCli("prune", "corrupt")),
        ("memory export CLI preserves its destination on corrupt input", () => MaintenanceCli("export", "corrupt")),
        ("memory export CLI preserves its destination on unsupported input", () => MaintenanceCli("export", "unsupported")),
        ("memory prune CLI reports successful persistence", () => MaintenanceCli("prune", "valid")),
        ("memory export CLI produces a readable current store", () => MaintenanceCli("export", "valid")),
        ("memory prune reports an actual persistence failure", () => FailedWrite("prune")),
        ("memory export preserves its destination on persistence failure", () => FailedWrite("export"))
    ];

    static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
    static ActionDto Click(int x, int y) => new() { Type = "click", XPx = x, YPx = y };
    static Point Position(ActionDto action) => DesktopInputService.CaptureResolvedAction(action, null).ScreenPoint!.Value;
    static void Near(Point actual, Point expected) => Check(Math.Abs(actual.X - expected.X) <= 1 && Math.Abs(actual.Y - expected.Y) <= 1,
        $"expected {expected}, got {actual}");

    static void ControlScale(string kind)
    {
        using var f = new Fixture(turnBased: true);
        var tracker = f.Run.turnBasedTransitions;
        var first = kind switch
        {
            "normalized" => new ActionDto { Type = "double_click", X = .3, Y = .2 },
            "bbox" => new ActionDto { Type = "click", BBox = new() { Left = 80, Top = 20, Right = 120, Bottom = 60 } },
            _ => Click(100, 40)
        };
        first.Button = "right";
        first.ResolvedTurnInputLabel = "ArrowUp";
        var original = Position(first);
        Check(tracker.TryBuildObservedDirectionalFollowUps(first, ["ArrowUp"], out _), "template was not learned");
        foreach (var width in new[] { 160, 320, 160 })
        {
            DesktopInputService.SetCurrentScreenMap(0, 0, 320, 200, width, width * 200 / 320);
            var right = Click(60, 30);
            right.ResolvedTurnInputLabel = "ArrowRight";
            Check(tracker.TryBuildObservedDirectionalFollowUps(right, ["ArrowUp"], out var followUps), "learned direction lost during resize");
            Near(Position(followUps[0]), original);
            Check(followUps[0].Button == "right" && followUps[0].Type == first.Type, "click semantics changed");
            followUps[0].XPx = 0; // Returned actions must not mutate retained evidence.
        }
    }

    static void HistoricalControl()
    {
        using var f = new Fixture();
        var snapshot = DesktopInputService.CaptureResolvedAction(Click(100, 40), null);
        DesktopInputService.SetCurrentScreenMap(0, 0, 320, 200, 160, 100);
        var memory = new DirectionalControlMemory();
        memory.Remember(snapshot, "up");
        Near(Position(memory.Recall("ArrowUp")!), new(100, 40));
    }

    static void ControlRecognition()
    {
        using var f = new Fixture();
        var memory = new DirectionalControlMemory();
        memory.Remember(Click(100, 40), "up");
        DesktopInputService.SetCurrentScreenMap(0, 0, 320, 200, 160, 100);
        Check(memory.TryResolve(DesktopInputService.CaptureResolvedAction(Click(50, 20), null), out var label) && label == "ArrowUp", "scaled target not recognized");
        Check(!memory.TryResolve(DesktopInputService.CaptureResolvedAction(Click(100, 40), null), out _), "old image coordinates still recognized");
    }

    static void ControlWindow()
    {
        using var f = new Fixture();
        var memory = new DirectionalControlMemory();
        memory.Remember(Click(100, 40), "up");
        ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(160, 0, 160, 200);
        Check(memory.Recall("ArrowUp") is null, "moved window retained physical controls");
    }

    static void InitializeRegion(Fixture f, Rectangle region)
    {
        f.Run.turnBasedInteractionRect = region;
        f.Run.turnBasedInteractionContext = TurnInteractionContext.Capture(f.Context);
        f.Run.nextFocusRect = region;
        f.Run.shortTermPlan.Update(new() { PlanStatus = "active", ShortTermPlan = "move twice", PlannedInputs = ["right", "right"], PlanConfidence = .99 });
    }

    static void WindowChange(bool resize)
    {
        using var f = new Fixture(turnBased: true);
        A.Options.ScreenSanityChecks = false;
        var old = new Rectangle(0, 0, 120, 100);
        ScreenshotService.ActiveWindowRectangleProvider = () => old;
        InitializeRegion(f, old);
        f.Run.SetAim(old);
        var moved = resize ? new Rectangle(0, 0, 220, 180) : new Rectangle(160, 0, 120, 100);
        ScreenshotService.ActiveWindowRectangleProvider = () => moved;
        var turn = new ControlStepState(2);
        Check(ControlObservationStage.RunAsync(f.Run, turn).GetAwaiter().GetResult() == StepFlow.Next, "observation stopped");
        Check(f.Run.turnBasedInteractionRect == moved, "old interaction rectangle survived");
        Check(turn.focusUrl is null && turn.appliedFocusRectForPrompt is null, "stale crop reached the request");
        Check(f.Run.nextFocusRect is null && f.Run.lastAimRect is null && !f.Run.shortTermPlan.HasActivePlan, "geometry-bound plan or aim survived");
        Check(f.Run.stateEpoch == 1, "context change did not start a fresh observation epoch");
    }

    static void StableGeometry()
    {
        using var f = new Fixture(turnBased: true);
        var region = new Rectangle(0, 0, 320, 200);
        InitializeRegion(f, region);
        var context = f.Run.turnBasedInteractionContext;
        DesktopInputService.SetCurrentScreenMap(0, 0, 320, 200, 160, 100);
        Check(context == TurnInteractionContext.Capture(f.Context) && context!.GeometryIsCurrent, "image scale changed physical context identity");
        var turn = new ControlStepState(2);
        ControlObservationStage.RunAsync(f.Run, turn).GetAwaiter().GetResult();
        Check(f.Run.stateEpoch == 0 && f.Run.shortTermPlan.HasActivePlan, "unchanged geometry unnecessarily reset plan");
    }

    static void MovedDuringPlanning()
    {
        using var f = new Fixture(turnBased: true);
        InitializeRegion(f, new(0, 0, 320, 200));
        var turn = f.Turn(new() { Type = "keys", Keys = ["right"] });
        ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(10, 0, 310, 200);
        f.Execute(turn);
        Check(turn.actionExecutionFailed && f.Input.Sends == 0, "stale plan sent native input after window movement");
    }

    sealed class Capture : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multiple) => new(-320, -100, 320, 200);
        public void CopyPixels(Graphics graphics, Rectangle region) => graphics.Clear(Color.Navy);
    }
    sealed class Response(QaLocateDto answer) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(JsonSerializer.Serialize(new
            { status = "completed", output_text = JsonSerializer.Serialize(answer) })) });
    }
    static void QaLocation(string kind)
    {
        var options = A.Options.Snapshot();
        var oldCapture = DesktopCapture.Backend;
        var oldContext = UiAutomationService.PromptContextProvider;
        var oldWindow = ScreenshotService.ActiveWindowRectangleProvider;
        var oldHttp = OpenAiTransport.HttpClient;
        var output = Console.Out;
        using var text = new StringWriter();
        var answer = kind switch
        {
            "normalized" => new QaLocateDto { X = .5, Y = .5 },
            "pixel" => new QaLocateDto { XPx = 80, YPx = 50 },
            _ => new QaLocateDto { BBox = new() { Left = 20, Top = 10, Right = 40, Bottom = 30 } }
        };
        using var client = new HttpClient(new Response(answer));
        try
        {
            A.Options.BatchMode = true;
            A.Options.LogScreens = A.Options.LogRequests = false;
            A.Options.IncludeFocusUia = A.Options.IncludeUiaTargets = false;
            A.Options.ScreenSanityChecks = false;
            A.Options.QaScreenshotMaxWidth = 160;
            A.Options.MultiMonitorEnabled = true;
            A.Options.MaxArtifactsPerDir = 0;
            DesktopCapture.Backend = new Capture();
            UiAutomationService.PromptContextProvider = (_, _, _) => new("Synthetic", "synthetic", null, null, null, null);
            ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(-320, -100, 320, 200);
            OpenAiTransport.HttpClient = client;
            Console.SetOut(text);
            Check(QuestionAnswerService.RunAskOnce("synthetic", "/ask Where?").GetAwaiter().GetResult() == CommandOutcome.Success, "QA failed");
            var expected = kind switch
            {
                "normalized" => "Location (desktop): -160:0",
                "pixel" => "Location (image): 80:50; desktop: -159:1",
                _ => "bbox (image)=(20,10)-(40,30); desktop=(-280,-80)-(-240,-40)"
            };
            Check(text.ToString().Contains(expected), "incorrect coordinate output: " + text);
        }
        finally
        {
            Console.SetOut(output);
            DesktopCapture.Backend = oldCapture;
            UiAutomationService.PromptContextProvider = oldContext;
            ScreenshotService.ActiveWindowRectangleProvider = oldWindow;
            OpenAiTransport.HttpClient = oldHttp;
            foreach (var field in typeof(AppOptions).GetFields(BindingFlags.Instance | BindingFlags.NonPublic))
                if (!field.IsInitOnly) field.SetValue(A.Options, field.GetValue(options));
        }
    }

    sealed class Files : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string Memory => Path.Combine(Root, "memory.json");
        internal string Export => Path.Combine(Root, "export.json");
        internal Files(string kind)
        {
            Directory.CreateDirectory(Root);
            Write(Memory, kind switch { "corrupt" => "{broken", "unsupported" => "{\"version\":999}", _ => "{\"version\":7,\"lessons\":[],\"calibration\":{}}" });
            Write(Export, "keep-existing-export");
        }
        internal static void Write(string path, string text) => File.WriteAllText(path, text + "\r\n", new UTF8Encoding(true));
        public void Dispose() => Directory.Delete(Root, true);
    }

    static void MaintenanceCli(string operation, string kind)
    {
        using var files = new Files(kind);
        var before = File.ReadAllBytes(files.Export);
        var source = File.ReadAllBytes(files.Memory);
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        foreach (var arg in new[] { "--command-main-probe", "--recovery-memory", "--recovery-memory-path", files.Memory,
            operation == "prune" ? "--memory-prune" : "--memory-export" }) start.ArgumentList.Add(arg);
        if (operation == "export") start.ArgumentList.Add(files.Export);
        using var child = Process.Start(start)!;
        try
        {
            var output = child.StandardOutput.ReadToEndAsync();
            var error = child.StandardError.ReadToEndAsync();
            Check(child.WaitForExit(10000), "maintenance CLI did not exit");
            Task.WhenAll(output, error).GetAwaiter().GetResult();
            Check(child.ExitCode == (kind == "valid" ? 0 : 1), $"unexpected exit {child.ExitCode}: {output.Result} {error.Result}");
            if (kind != "valid")
            {
                Check(File.ReadAllBytes(files.Export).SequenceEqual(before), "failed maintenance overwrote export");
                Check(File.ReadAllBytes(files.Memory).SequenceEqual(source), "failed maintenance changed unsupported or corrupt source");
            }
            else if (operation == "export")
                Check(RecoveryStoreReader.Read(files.Export).Status == RecoveryStoreReadStatus.Loaded, "successful export is unreadable");
        }
        finally { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } }
    }

    sealed class Unwritable(AtomicTextStore.IFileOperations inner) : AtomicTextStore.IFileOperations
    {
        public bool Exists(string path) => inner.Exists(path);
        public void Write(string path, string text) => throw new IOException("simulated full disk");
        public void Replace(string source, string destination, string? backup) => throw new IOException("simulated full disk");
        public void Move(string source, string destination) => throw new IOException("simulated full disk");
        public void Delete(string path) => inner.Delete(path);
    }
    static void FailedWrite(string operation)
    {
        using var scope = A.BeginRun();
        using var files = new Files("valid");
        A.Options.RecoveryMemoryEnabled = true;
        A.Options.RecoveryMemoryPath = files.Memory;
        var before = File.ReadAllBytes(files.Export);
        var oldFiles = AtomicTextStore.Files;
        try
        {
            AtomicTextStore.Files = new Unwritable(oldFiles);
            Check(RecoveryMemoryService.ExecuteRecoveryMemoryMaintenance(operation, files.Export) == CommandOutcome.Failed, "write failure reported success");
            Check(File.ReadAllBytes(files.Export).SequenceEqual(before), "failed export overwrote destination");
        }
        finally { AtomicTextStore.Files = oldFiles; }
    }
}
