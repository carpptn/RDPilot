using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class ReplayConfigurationRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
        (from kind in new[] { "version", "epoch", "screen", "active", "action", "history", "null-frame" }
         from index in new[] { 1, 11 }
         select ($"replay rejects {kind} corruption {(index == 1 ? "before" : "after")} loop detection",
             (Action)(() => InvalidFrame(kind, index))))
        .Concat(new (string, Action)[]
        {
            ("validated replay retains the first loop result", FirstDetection),
            ("invalid replay import preserves the entire existing corpus", ImportPreservesCorpus),
            ("valid independent replay import merges cases", ValidImport),
            ("replay CLI rejects a corrupt tail", () => InvalidCli(false)),
            ("import CLI rejects a corrupt tail without replacing its destination", () => InvalidCli(true))
        })
        .Concat(Enumerable.Range(0, 6).Select(stage =>
            ($"configuration precedence remains stable at stage {stage}", (Action)(() => ConfigurationPrecedence(stage)))))
        .ToArray();

    static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
    }

    static LoopReplayCase LoopCase()
    {
        var replay = new LoopReplayCase
        {
            Name = "reviewed-cycle", LabelSource = "manual:reviewed", ExpectedLoop = true,
            ScreenWidth = 320, ScreenHeight = 200
        };
        for (var index = 0; index < 12; index++)
        {
            var fingerprint = Convert.ToBase64String(Enumerable.Repeat((byte)(index % 2 == 0 ? 0 : 255), 96 * 54).ToArray());
            replay.Frames.Add(new()
            {
                ScreenFingerprintBase64 = fingerprint, ActiveProcess = "game", WindowTitle = "Game",
                GoalProgress = "Neutral", LastDelta = .5,
                PreviousAction = index % 4 == 0 ? null : new() { Type = "keys", Keys = [index % 2 == 0 ? "left" : "right"] }
            });
        }
        return replay;
    }

    static string Corrupt(LoopReplayCase replay, string kind, int index)
    {
        var frame = replay.Frames[index];
        switch (kind)
        {
            case "version": frame.ReplayStateVersion = 99; return "metadata";
            case "epoch": frame.StateEpoch = -1; return "metadata";
            case "screen": frame.ScreenFingerprintBase64 = "not base64!"; return "screen fingerprint";
            case "active": frame.ActiveWindowFingerprintBase64 = "not base64!"; return "active-window fingerprint";
            case "action": frame.PreviousAction = new() { Type = "click" }; return "source coordinates";
            case "history":
                frame.ReplayStateVersion = 2;
                frame.RecentActions = [];
                frame.PreviousActionIndex = 0;
                return "previous action index";
            default: replay.Frames[index] = null!; return "null frame";
        }
    }

    static void Reject(Action operation, string detail)
    {
        try { operation(); }
        catch (InvalidOperationException ex)
        {
            Check(ex.Message.Contains(detail, StringComparison.OrdinalIgnoreCase), "unexpected validation failure: " + ex.Message);
            return;
        }
        throw new InvalidOperationException("Invalid replay was accepted: " + detail);
    }

    static void InvalidFrame(string kind, int index)
    {
        using var fixture = new SyntheticControlFixture();
        var replay = LoopCase();
        var first = LoopReplayService.ReplayCase(replay);
        Check(first.Detected && first.Step < 12, "fixture does not detect a loop before its tail");
        var detail = Corrupt(replay, kind, index);
        Reject(() => LoopReplayService.ReplayCase(replay), detail);
    }

    static void FirstDetection()
    {
        using var fixture = new SyntheticControlFixture();
        var replay = LoopCase();
        var complete = LoopReplayService.ReplayCase(replay);
        Check(complete.Detected && complete.Step == 5, "first detection changed");
        replay.Frames = replay.Frames.Take(complete.Step).ToList();
        Check(LoopReplayService.ReplayCase(replay) == complete, "tail validation changed the first loop verdict");
    }

    sealed class Files : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        internal string Source => Path.Combine(Root, "source.json");
        internal string Destination => Path.Combine(Root, "corpus.json");
        internal Files() => Directory.CreateDirectory(Root);
        internal static void Write(string path, string text) => File.WriteAllText(path,
            text.Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n", new UTF8Encoding(true));
        internal static void Corpus(string path, params LoopReplayCase[] cases) =>
            Write(path, JsonSerializer.Serialize(new LoopReplayCorpus { Cases = cases.ToList() }, A.PrettyJson));
        public void Dispose() => Directory.Delete(Root, true);
    }

    static void ImportPreservesCorpus()
    {
        using var scope = A.BeginRun();
        using var files = new Files();
        A.Options.LoopReplayCorpusPath = files.Destination;
        Files.Corpus(files.Destination, LoopCase());
        var before = File.ReadAllBytes(files.Destination);
        var valid = LoopCase();
        valid.Name = "new-valid-case";
        var invalid = LoopCase();
        Corrupt(invalid, "screen", 11);
        Files.Corpus(files.Source, valid, invalid);
        Reject(() => LoopReplayService.ImportIndependentLoopReplayCorpus(files.Source), "screen fingerprint");
        Check(File.ReadAllBytes(files.Destination).SequenceEqual(before), "failed import partially changed the destination");
    }

    static void ValidImport()
    {
        using var scope = A.BeginRun();
        using var files = new Files();
        A.Options.LoopReplayCorpusPath = files.Destination;
        var existing = LoopCase();
        existing.Name = "existing";
        Files.Corpus(files.Destination, existing);
        Files.Corpus(files.Source, LoopCase());
        LoopReplayService.ImportIndependentLoopReplayCorpus(files.Source);
        var merged = JsonSerializer.Deserialize<LoopReplayCorpus>(File.ReadAllText(files.Destination), A.PrettyJson)!;
        Check(merged.Cases.Select(item => item.Name).SequenceEqual(new[] { "existing", "reviewed-cycle" }), "valid import lost a case");
    }

    static void InvalidCli(bool import)
    {
        using var files = new Files();
        var replay = LoopCase();
        Corrupt(replay, "version", 11);
        Files.Corpus(files.Source, replay);
        Files.Corpus(files.Destination, LoopCase());
        var before = File.ReadAllBytes(files.Destination);
        var start = new ProcessStartInfo(Environment.ProcessPath!)
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
            start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
        foreach (var arg in new[] { "--command-main-probe", import ? "--loop-replay-import" : "--loop-replay", files.Source,
            "--loop-replay-corpus", files.Destination }) start.ArgumentList.Add(arg);
        using var child = Process.Start(start)!;
        try
        {
            var output = child.StandardOutput.ReadToEndAsync();
            var error = child.StandardError.ReadToEndAsync();
            Check(child.WaitForExit(10000), "replay CLI did not exit");
            Task.WhenAll(output, error).GetAwaiter().GetResult();
            Check(child.ExitCode == 1 && error.Result.Contains("frame 12"), $"unexpected CLI result: {child.ExitCode}; {output.Result}; {error.Result}");
            Check(File.ReadAllBytes(files.Destination).SequenceEqual(before), "failed CLI operation changed the destination");
        }
        finally { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } }
    }

    static void ConfigurationPrecedence(int stage)
    {
        using var scope = A.BeginRun();
        using var files = new Files();
        var names = new[] { "RDPILOT_CONFIG", "RDPILOT_PROFILE", "FAST_MODE", "BALANCED_MODE", "QUALITY_MODE", "OPENAI_MODEL", "POST_ACTION_DELAY_MS" };
        var environment = names.ToDictionary(name => name, Environment.GetEnvironmentVariable);
        var oldDirectory = Environment.CurrentDirectory;
        try
        {
            foreach (var name in names) Environment.SetEnvironmentVariable(name, null);
            Environment.CurrentDirectory = files.Root;
            Files.Write(Path.Combine(files.Root, "appsettings.json"), "{\"profile\":\"custom\",\"model\":\"automatic-model\",\"postActionDelayMs\":101}");
            Files.Write(files.Source, "{\"profile\":\"custom\",\"model\":\"file-model\",\"postActionDelayMs\":202}");
            if (stage >= 1) Environment.SetEnvironmentVariable("RDPILOT_CONFIG", files.Source);
            if (stage >= 2) Environment.SetEnvironmentVariable("RDPILOT_PROFILE", "quality");
            if (stage >= 3)
            {
                Environment.SetEnvironmentVariable("OPENAI_MODEL", "environment-model");
                Environment.SetEnvironmentVariable("POST_ACTION_DELAY_MS", "303");
            }
            ConfigurationService.ApplyEnvironmentConfig();
            if (stage >= 4)
            {
                var arguments = stage == 4
                    ? new[] { "--delay", "404", "--model", "cli-model", "--profile", "fast" }
                    : new[] { "--profile", "fast", "--delay", "404", "--model", "cli-model" };
                Check(ConfigurationService.ParseCliArgs(arguments).Success, "CLI configuration was rejected");
            }
            ConfigurationService.NormalizeConfig();
            var expectedDelay = stage switch { 0 => 101, 1 => 202, 2 => 1000, 3 => 303, _ => 404 };
            var expectedModel = stage switch { 0 => "automatic-model", 1 or 2 => "file-model", 3 => "environment-model", _ => "cli-model" };
            var expectedProfile = stage switch { 0 or 1 => "custom", 2 or 3 => "quality", _ => "fast" };
            Check(A.Options.UiSettleDelayMs == expectedDelay && A.Options.Model == expectedModel && A.Options.RunProfile == expectedProfile,
                $"unexpected settings at stage {stage}: {A.Options.UiSettleDelayMs}, {A.Options.Model}, {A.Options.RunProfile}");
        }
        finally
        {
            Environment.CurrentDirectory = oldDirectory;
            foreach (var (name, value) in environment) Environment.SetEnvironmentVariable(name, value);
        }
    }
}
