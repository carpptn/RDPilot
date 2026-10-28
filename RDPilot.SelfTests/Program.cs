using System.Drawing;
using System.Reflection;

if (args.FirstOrDefault() == "--prompt-history-worker")
    return PromptHistoryRegressionTests.Worker(args);

if (args.FirstOrDefault() == "--storage-family-worker")
    return StorageFamilyRegressionTests.Worker(args);

if (args.FirstOrDefault() == "--recovery-writer-probe")
{
    Console.WriteLine(RecoveryWriterSession.CurrentId);
    Console.ReadLine();
    return 0;
}
if (args.FirstOrDefault() == "--command-main-probe")
    return await (Task<int>)typeof(RDPilotApplication).GetMethod("Main", BindingFlags.Static | BindingFlags.NonPublic)!
        .Invoke(null, new object[] { args.Skip(1).ToArray() })!;

if (args.Contains("--live-api-smoke", StringComparer.Ordinal))
    return await RDPilot.SelfTests.Integration.LiveApiSmoke.RunAsync();

TestState.Field(typeof(RDPilotApplication), "RecoveryMemoryEnabled", BindingFlags.Static | BindingFlags.NonPublic)!.SetValue(null, false);
var tests = BehavioralTests.Cases.Concat(AuditRegressionTests.Cases).Concat(RuntimeRegressionTests.Cases).Concat(PipelineRegressionTests.Cases).Concat(ReviewRegressionTests.Cases).Concat(ThirdReviewRegressionTests.Cases).Concat(FourthReviewRegressionTests.Cases).Concat(FifthReviewRegressionTests.Cases).Concat(SixthReviewRegressionTests.Cases).Concat(SeventhReviewRegressionTests.Cases).Concat(EighthReviewRegressionTests.Cases).Concat(NinthReviewRegressionTests.Cases).Concat(CompletionRegressionTests.Cases).Concat(VisualPlanRegressionTests.Cases).Concat(GeometryMaintenanceRegressionTests.Cases).Concat(ReplayConfigurationRegressionTests.Cases).Concat(AimLifecycleRegressionTests.Cases).Concat(RecoveryStoragePathRegressionTests.Cases).Concat(StorageFamilyRegressionTests.Cases).Concat(ArchiveRotationRegressionTests.Cases).Concat(PromptHistoryRegressionTests.Cases).ToArray();
var failures = new List<string>();
foreach (var test in tests)
{
    try
    {
        test.Run();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception ex)
    {
        failures.Add($"{test.Name}: {BehavioralTests.RootException(ex).Message}");
        Console.WriteLine($"FAIL {test.Name}: {BehavioralTests.RootException(ex).Message}");
    }
}

Console.WriteLine($"{tests.Length - failures.Count}/{tests.Length} self-tests passed.");
if (failures.Count > 0)
{
    foreach (var failure in failures)
        Console.Error.WriteLine(failure);
    return 1;
}

return 0;
