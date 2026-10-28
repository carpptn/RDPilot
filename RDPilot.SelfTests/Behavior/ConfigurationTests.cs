using System.Drawing;
using System.Reflection;

namespace RDPilot.SelfTests.Behavior;
internal static partial class BehavioralTests
{
    static void TestMultiMonitorOptIn()
    {
        var root = typeof(RDPilotApplication);
        var field = TestState.Field(root, "MultiMonitorEnabled", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = field.GetValue(null);
        try
        {
            field.SetValue(null, false);
            _ = RDPilot.Configuration.ConfigurationService.ApplyCliArgs(["--multi-monitor"]);
            Assert((bool)field.GetValue(null)!, "--multi-monitor did not enable the virtual desktop");
            _ = RDPilot.Configuration.ConfigurationService.ApplyCliArgs(["--primary-monitor-only"]);
            Assert(!(bool)field.GetValue(null)!, "--primary-monitor-only did not restore the safe default");
        }
        finally
        {
            field.SetValue(null, original);
        }
    }

    static void TestBatchCliParsing()
    {
        var root = typeof(RDPilotApplication);
        var batchField = TestState.Field(root, "BatchMode", BindingFlags.Static | BindingFlags.NonPublic)!;
        var errorField = TestState.Field(root, "CliArgumentError", BindingFlags.Static | BindingFlags.NonPublic)!;
        var modelField = TestState.Field(root, "Model", BindingFlags.Static | BindingFlags.NonPublic)!;
        var webField = TestState.Field(root, "AllowWebSearch", BindingFlags.Static | BindingFlags.NonPublic)!;
        var memoryPathField = TestState.Field(root, "RecoveryMemoryPath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var archivePathField = TestState.Field(root, "RecoveryMemoryArchivePath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var corpusPathField = TestState.Field(root, "LoopReplayCorpusPath", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originals = new[]
        {
            batchField.GetValue(null),
            errorField.GetValue(null),
            modelField.GetValue(null),
            webField.GetValue(null),
            memoryPathField.GetValue(null),
            archivePathField.GetValue(null),
            corpusPathField.GetValue(null)
        };
        var tempRoot = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        var taskFile = Path.Combine(tempRoot, "task.txt");
        File.WriteAllText(taskFile, "first line\r\nsecond line");
        try
        {
            batchField.SetValue(null, false);
            errorField.SetValue(null, false);
            webField.SetValue(null, false);
            var task = RDPilot.Configuration.ConfigurationService.ApplyCliArgs(["--task", "open Paint", "--model", "gpt-test", "--allow-web-search"]);
            Assert(task == "open Paint" && (bool)batchField.GetValue(null)! && !(bool)errorField.GetValue(null)! && (string)modelField.GetValue(null)! == "gpt-test" && (bool)webField.GetValue(null)!, "--task did not enable one-shot batch mode alongside configuration flags");
            batchField.SetValue(null, false);
            errorField.SetValue(null, false);
            var fileTask = RDPilot.Configuration.ConfigurationService.ApplyCliArgs(["--task-file", taskFile, "--no-web-search"]);
            Assert(fileTask == "first line\r\nsecond line" && (bool)batchField.GetValue(null)! && !(bool)errorField.GetValue(null)! && !(bool)webField.GetValue(null)!, "--task-file did not preserve a multiline task or batch mode");
            Assert(RDPilotApplication.ControlRunExitCode(ControlRunOutcome.Completed) == 0 && RDPilotApplication.ControlRunExitCode(ControlRunOutcome.Failed) == 1 && RDPilotApplication.ControlRunExitCode(ControlRunOutcome.GuardStopped) == 3 && RDPilotApplication.ControlRunExitCode(ControlRunOutcome.StepLimitReached) == 4 && RDPilotApplication.ControlRunExitCode(ControlRunOutcome.Cancelled) == 130, "batch control outcomes do not map to stable process exit codes");
            Assert(Path.GetFullPath(RDPilot.Diagnostics.ArtifactStorageService.EnsureLogDir()).StartsWith(Path.GetFullPath(AppContext.BaseDirectory), StringComparison.OrdinalIgnoreCase) && Path.GetFullPath(RDPilot.Diagnostics.ArtifactStorageService.EnsureScreensDir()).StartsWith(Path.GetFullPath(AppContext.BaseDirectory), StringComparison.OrdinalIgnoreCase) && Path.GetFullPath(RDPilot.Diagnostics.ArtifactStorageService.EnsureRequestsDir()).StartsWith(Path.GetFullPath(AppContext.BaseDirectory), StringComparison.OrdinalIgnoreCase), "batch artifacts are not rooted beside the RDPilot executable");
            memoryPathField.SetValue(null, "memory\\custom-memory.json");
            archivePathField.SetValue(null, "memory\\custom-archive.json");
            corpusPathField.SetValue(null, "memory\\custom-corpus.json");
            Assert(RDPilot.Persistence.RecoveryMemoryService.EffectiveRecoveryMemoryPath() == Path.Combine(AppContext.BaseDirectory, "memory", "custom-memory.json") && RDPilot.Persistence.RecoveryMemoryService.EffectiveRecoveryMemoryArchivePath() == Path.Combine(AppContext.BaseDirectory, "memory", "custom-archive.json") && RDPilot.Persistence.RecoveryMemoryService.EffectiveLoopReplayCorpusPath() == Path.Combine(AppContext.BaseDirectory, "memory", "custom-corpus.json"), "relative durable-memory paths are not rooted beside the executable");
        }
        finally
        {
            batchField.SetValue(null, originals[0]);
            errorField.SetValue(null, originals[1]);
            modelField.SetValue(null, originals[2]);
            webField.SetValue(null, originals[3]);
            memoryPathField.SetValue(null, originals[4]);
            archivePathField.SetValue(null, originals[5]);
            corpusPathField.SetValue(null, originals[6]);
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    static void TestShippedAppSettingsDefaults()
    {
        const int expectedPropertyCount = 145;
        var configPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
        Assert(File.Exists(configPath), "appsettings.json was not copied beside the self-test executable");
        using var document = System.Text.Json.JsonDocument.Parse(File.ReadAllText(configPath));
        var properties = document.RootElement.EnumerateObject().ToArray();
        var propertyNames = properties.Select(property => property.Name).ToArray();
        Assert(properties.Length == expectedPropertyCount, $"appsettings.json exposes {properties.Length} properties instead of {expectedPropertyCount}");
        Assert(propertyNames.Distinct(StringComparer.Ordinal).Count() == expectedPropertyCount, "appsettings.json contains duplicate property names");
        var expectedNullProperties = new[]
        {
            "loopReplayCorpusPath",
            "qaModel",
            "qaReasoningEffort",
            "recoveryMemoryArchivePath",
            "recoveryMemoryPath",
            "verifyModel"
        };
        var actualNullProperties = properties.Where(property => property.Value.ValueKind == System.Text.Json.JsonValueKind.Null).Select(property => property.Name).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        Assert(actualNullProperties.SequenceEqual(expectedNullProperties, StringComparer.Ordinal), "appsettings.json null defaults do not match nullable code defaults");
        var root = typeof(RDPilotApplication);
        var fields = TestState.Fields(root, BindingFlags.Static | BindingFlags.NonPublic).Where(field => !field.IsLiteral && !field.IsInitOnly && IsComparableScalar(field.FieldType)).ToArray();
        var originalValues = fields.ToDictionary(field => field, field => field.GetValue(null));
        var recoveryMemoryField = TestState.Field(root, "RecoveryMemoryEnabled", BindingFlags.Static | BindingFlags.NonPublic)!;
        try
        {
            // The test harness disables recovery memory before running tests. Restore
            // its code default while validating the shipped default configuration.
            recoveryMemoryField.SetValue(null, true);
            var expectedValues = fields.ToDictionary(field => field, field => field.GetValue(null));
            RDPilot.Configuration.ConfigurationService.ApplyConfigFile(configPath);
            foreach (var field in fields)
            {
                // Loading a concrete control effort intentionally records that the
                // value came from configuration. The effective value remains equal.
                if (field.Name == "ReasoningEffortExplicit")
                    continue;
                var actual = field.GetValue(null);
                var expected = expectedValues[field];
                Assert(Equals(actual, expected), $"appsettings.json default changed {field.Name}: expected '{expected}', got '{actual}'");
            }
        }
        finally
        {
            foreach (var pair in originalValues)
                pair.Key.SetValue(null, pair.Value);
        }

        static bool IsComparableScalar(Type type) => type.IsPrimitive || type.IsEnum || type == typeof(string) || type == typeof(decimal);
    }

    static void TestAutomaticConfigLoading()
    {
        var modelField = TestState.Field(typeof(RDPilotApplication), "Model", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originalModel = modelField.GetValue(null);
        var tempRoot = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        var baseDirectory = Path.Combine(tempRoot, "base");
        var currentDirectory = Path.Combine(tempRoot, "cwd");
        Directory.CreateDirectory(baseDirectory);
        Directory.CreateDirectory(currentDirectory);
        var baseAppSettings = Path.Combine(baseDirectory, "appsettings.json");
        var cwdAppSettings = Path.Combine(currentDirectory, "appsettings.json");
        File.WriteAllText(baseAppSettings, """{"model":"base-appsettings"}""");
        File.WriteAllText(cwdAppSettings, """{"model":"cwd-appsettings"}""");
        try
        {
            var path = RDPilot.Configuration.ConfigurationService.ResolveAutomaticConfigPath(baseDirectory, currentDirectory);
            Assert(path == cwdAppSettings, "automatic config paths did not prefer the working-directory appsettings file");
            RDPilot.Configuration.ConfigurationService.ApplyAutomaticConfigFile(baseDirectory, currentDirectory);
            Assert((string? )modelField.GetValue(null) == "cwd-appsettings", "working-directory appsettings.json was not loaded");
            File.Delete(cwdAppSettings);
            path = RDPilot.Configuration.ConfigurationService.ResolveAutomaticConfigPath(baseDirectory, currentDirectory);
            Assert(path == baseAppSettings, "automatic config path did not fall back to the executable-directory file");
            RDPilot.Configuration.ConfigurationService.ApplyAutomaticConfigFile(baseDirectory, currentDirectory);
            Assert((string? )modelField.GetValue(null) == "base-appsettings", "executable-directory appsettings.json was not loaded");
        }
        finally
        {
            modelField.SetValue(null, originalModel);
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    static void TestAdvancedJsonTuning()
    {
        var root = typeof(RDPilotApplication);
        var expected = new Dictionary<string, object>
        {
            ["ApiUrl"] = "https://example.invalid/v1/responses",
            ["SendFocusCrop"] = false,
            ["FocusRingPadding"] = 9,
            ["FocusRingThickness"] = 4,
            ["FocusGlowThickness"] = 5,
            ["FocusCornerRadius"] = 12,
            ["LargeFocusOverlayAreaRatio"] = 0.31,
            ["ClickAimEdgeAdjustMaxAreaRatio"] = 0.33,
            ["ClickAimEdgeMarginRatio"] = 0.21,
            ["ClickAimEdgeMinMarginPx"] = 11,
            ["ClickAimEdgeMaxMarginPx"] = 45,
            ["GridLabelEveryPx"] = 75,
            ["GridMajorEveryPx"] = 150,
            ["AimExpireDelta"] = 0.12,
            ["NoChangeThreshold"] = 0.009,
            ["MaxGesturePathPoints"] = 256,
            ["MaxGestureDurationMs"] = 8000,
            ["MaxHeldKeys"] = 6,
            ["MaxKeyHoldDurationMs"] = 7000,
            ["IneffectiveMouseClusterPx"] = 120,
            ["MaxFocusUiaCropPixels"] = 250000,
            ["MaxArtifactsPerDir"] = 321,
            ["MaxBatchedGesturePoints"] = 333,
            ["MaxBatchedGestureDurationMs"] = 15000
        };
        var fields = expected.Keys.ToDictionary(name => name, name => TestState.Field(root, name, BindingFlags.Static | BindingFlags.NonPublic)!);
        var originals = fields.ToDictionary(pair => pair.Key, pair => pair.Value.GetValue(null));
        var tempRoot = Path.Combine(Path.GetTempPath(), "RDPilotSelfTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);
        var configPath = Path.Combine(tempRoot, "appsettings.json");
        File.WriteAllText(configPath, """
        {
          "apiUrl": "https://example.invalid/v1/responses",
          "sendFocusCrop": false,
          "focusRingPadding": 9,
          "focusRingThickness": 4,
          "focusGlowThickness": 5,
          "focusCornerRadius": 12,
          "largeFocusOverlayAreaRatio": 0.31,
          "clickAimEdgeAdjustMaxAreaRatio": 0.33,
          "clickAimEdgeMarginRatio": 0.21,
          "clickAimEdgeMinMarginPx": 11,
          "clickAimEdgeMaxMarginPx": 45,
          "gridLabelEveryPx": 75,
          "gridMajorEveryPx": 150,
          "aimExpireDelta": 0.12,
          "noChangeThreshold": 0.009,
          "maxGesturePathPoints": 256,
          "maxGestureDurationMs": 8000,
          "maxHeldKeys": 6,
          "maxKeyHoldDurationMs": 7000,
          "ineffectiveMouseClusterPx": 120,
          "maxFocusUiaCropPixels": 250000,
          "maxArtifactsPerDir": 321,
          "maxBatchedGesturePoints": 333,
          "maxBatchedGestureDurationMs": 15000
        }
        """);
        try
        {
            RDPilot.Configuration.ConfigurationService.ApplyConfigFile(configPath);
            RDPilot.Configuration.ConfigurationService.NormalizeConfig();
            foreach (var pair in expected)
            {
                var actual = fields[pair.Key].GetValue(null);
                if (pair.Value is double expectedDouble)
                {
                    Assert(actual is double actualDouble && Math.Abs(actualDouble - expectedDouble) < 0.000001, $"appsettings.json did not apply {pair.Key}");
                }
                else
                {
                    Assert(Equals(actual, pair.Value), $"appsettings.json did not apply {pair.Key}");
                }
            }
        }
        finally
        {
            foreach (var pair in originals)
                fields[pair.Key].SetValue(null, pair.Value);
            Directory.Delete(tempRoot, recursive: true);
        }
    }

    static void TestUnlimitedStepConfiguration()
    {
        var root = typeof(RDPilotApplication);
        var maxStepsField = TestState.Field(root, "MaxSteps", BindingFlags.Static | BindingFlags.NonPublic)!;
        var original = maxStepsField.GetValue(null);
        try
        {
            _ = RDPilot.Configuration.ConfigurationService.ApplyCliArgs(["--max-steps", "0"]);
            RDPilot.Configuration.ConfigurationService.NormalizeConfig();
            Assert((int)maxStepsField.GetValue(null)! == 0, "maxSteps=0 was normalized back to a finite limit");
        }
        finally
        {
            maxStepsField.SetValue(null, original);
        }
    }

    static void TestCustomProfileReset()
    {
        RDPilot.Configuration.ConfigurationService.ApplyProfile("fast");
        RDPilot.Configuration.ConfigurationService.ApplyProfile("custom");
        var root = typeof(RDPilotApplication);
        Assert(GetStatic<string>(root, "RunProfile") == "custom", "profile name did not reset");
        Assert(GetStatic<string?>(root, "ReasoningEffort") == "high", "reasoning did not reset to the high code default");
        Assert(GetStatic<string?>(root, "VerifyReasoningEffort") == "high", "verifier reasoning did not reset to the high code default");
        Assert(GetStatic<int>(root, "MaxStagnationStepsBeforeAbort") == 20, "stagnation default did not reset");
        Assert(GetStatic<int>(root, "RecoveryMemoryMaxLessons") == 500, "recovery-memory active default changed");
        Assert(GetStatic<int>(root, "RecoveryMemoryMaxQuarantinedLessons") == 500, "recovery-memory quarantine default changed");
    }

    static void TestMaxReasoningEffortDoesNotDowngrade()
    {
        var root = typeof(RDPilotApplication);
        var effortField = TestState.Field(root, "ReasoningEffort", BindingFlags.Static | BindingFlags.NonPublic)!;
        var adaptiveField = TestState.Field(root, "AdaptiveReasoningEffort", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originalEffort = effortField.GetValue(null);
        var originalAdaptive = adaptiveField.GetValue(null);
        try
        {
            effortField.SetValue(null, "max");
            adaptiveField.SetValue(null, true);
            var effective = (string? )InvokePrivate(typeof(RDPilot.Control.ControlGoalPolicy), "EffectiveReasoningEffort", 8, 4, 4);
            Assert(effective == "max", "adaptive effort downgraded max after stagnation");
        }
        finally
        {
            effortField.SetValue(null, originalEffort);
            adaptiveField.SetValue(null, originalAdaptive);
        }
    }

    static void TestProfilesPreserveStrongEffort()
    {
        var root = typeof(RDPilotApplication);
        var effortField = TestState.Field(root, "ReasoningEffort", BindingFlags.Static | BindingFlags.NonPublic)!;
        var explicitField = TestState.Field(root, "ReasoningEffortExplicit", BindingFlags.Static | BindingFlags.NonPublic)!;
        var originalEffort = effortField.GetValue(null);
        var originalExplicit = explicitField.GetValue(null);
        try
        {
            explicitField.SetValue(null, false);
            RDPilot.Configuration.ConfigurationService.ApplyProfile("custom");
            RDPilot.Configuration.ConfigurationService.ApplyProfile("fast");
            Assert(GetStatic<string?>(root, "ReasoningEffort") == "high", "fast profile lowered high reasoning effort");
            Assert(GetStatic<int>(root, "MaxOutputTokens") >= 10000, "fast profile lowered the max reasoning budget");
            Assert((string? )InvokePrivate(typeof(RDPilot.Configuration.ConfigurationService), "EffectiveQaReasoningEffort") == "high", "fast profile lowered QA reasoning effort");
            Assert(GetStatic<int>(root, "QaMaxOutputTokens") >= 4000, "fast profile lowered the QA reasoning budget");
            Assert((string? )InvokePrivate(typeof(RDPilot.Configuration.ConfigurationService), "EffectiveVerifyReasoningEffort") == "high", "fast profile lowered verifier reasoning effort");
            Assert(GetStatic<int>(root, "VerifyMaxOutputTokens") >= 6000, "fast profile lowered the verifier reasoning budget");
            Assert(GetStatic<int>(root, "TurnReanalysisMaxOutputTokens") >= 10000, "fast profile lowered the salient reanalysis budget");
            explicitField.SetValue(null, true);
            effortField.SetValue(null, "low");
            RDPilot.Configuration.ConfigurationService.ApplyProfile("quality");
            Assert(GetStatic<string?>(root, "ReasoningEffort") == "low", "explicit low effort was overwritten by a profile");
        }
        finally
        {
            effortField.SetValue(null, originalEffort);
            explicitField.SetValue(null, originalExplicit);
            RDPilot.Configuration.ConfigurationService.ApplyProfile("custom");
        }
    }
}
