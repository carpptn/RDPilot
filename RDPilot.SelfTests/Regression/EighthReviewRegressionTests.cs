using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Regression;

internal static class EighthReviewRegressionTests
{
    internal static (string Name, Action Run)[] Cases =>
    [
        ("compacted calibration merges stale counters in either order", CalibrationOrder),
        ("repeated calibration merges do not resurrect compacted writers", CalibrationRepeated),
        ("calibration preserves the latest of two compaction cutoffs", CalibrationCutoffs),
        ("uncompacted calibration keeps per-writer maxima", CalibrationUncompacted),
        ("lesson counters share nullable cutoff merge semantics", LessonCounters),
        ("compaction cutoff handles empty and nullable snapshots", NullableCutoffs),
        ("Q&A rejects missing empty whitespace and note-only answers", EmptyAnswers),
        ("Q&A rejects incomplete and out-of-range locations", InvalidLocations),
        ("Q&A accepts valid boundary locations without text", ValidLocations),
        ("Q&A preserves text while discarding invalid optional locations", TextWithInvalidLocation),
        ("Q&A batch returns failure for a null-field payload", () => BatchQa(new(), false)),
        ("Q&A batch returns failure for whitespace and a note", () => BatchQa(new() { AnswerText = " \t", Note = "metadata" }, false)),
        ("Q&A batch returns failure for an invalid location", () => BatchQa(new() { XPx = 99999, YPx = 1 }, false)),
        ("Q&A batch returns success for a textual answer", () => BatchQa(new() { AnswerText = "A blue screen." }, true)),
        ("Q&A batch returns success for a valid location", () => BatchQa(new() { X = 1, Y = 0 }, true))
    ];

    static void Check(bool value, string message)
    {
        if (!value) throw new InvalidOperationException(message);
    }

    static readonly DateTime Cutoff = new(2020, 1, 2, 0, 0, 0, DateTimeKind.Utc);
    static string Writer(DateTime date) => $"{date.Ticks:x16}-test";
    static readonly string OldWriter = Writer(Cutoff);
    static readonly string NewWriter = Writer(Cutoff.AddDays(1));
    static Dictionary<string, LoopCalibrationBucket> Bucket(LoopCalibrationBucket bucket) => new() { ["loop"] = bucket };
    static Dictionary<string, LoopCalibrationBucket> Compacted() => Bucket(new()
    {
        CandidateCount = 10, ConfirmedCount = 5, RejectedCount = 3, InconclusiveCount = 2,
        CompactedCandidateCount = 10, CompactedConfirmedCount = 5, CompactedRejectedCount = 3, CompactedInconclusiveCount = 2,
        CountersCompactedBeforeUtc = Cutoff
    });
    static Dictionary<string, LoopCalibrationBucket> Stale() => Bucket(new()
    {
        CandidateByWriter = new() { [OldWriter] = 10, [NewWriter] = 4 },
        ConfirmedByWriter = new() { [OldWriter] = 5, [NewWriter] = 2 },
        RejectedByWriter = new() { [OldWriter] = 3, [NewWriter] = 1 },
        InconclusiveByWriter = new() { [OldWriter] = 2, [NewWriter] = 1 }
    });
    static object Invoke(string method, params object[] args) => typeof(RecoveryMemoryService)
        .GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, args)!;
    static Dictionary<string, LoopCalibrationBucket> Merge(Dictionary<string, LoopCalibrationBucket> left,
        Dictionary<string, LoopCalibrationBucket> right) =>
        (Dictionary<string, LoopCalibrationBucket>)Invoke("MergeCalibration", left, right);
    static void AssertTotals(LoopCalibrationBucket bucket)
    {
        Check(bucket.CandidateCount == 14 && bucket.ConfirmedCount == 7 && bucket.RejectedCount == 4 && bucket.InconclusiveCount == 3,
            "compacted writer was counted again or a newer writer was lost");
        Check(bucket.CountersCompactedBeforeUtc == Cutoff, "compaction cutoff was lost");
        Check(new[] { bucket.CandidateByWriter, bucket.ConfirmedByWriter, bucket.RejectedByWriter, bucket.InconclusiveByWriter }
            .All(map => map.Count == 1 && map.ContainsKey(NewWriter)), "stale component survived compaction");
    }
    static void CalibrationOrder()
    {
        AssertTotals(Merge(Compacted(), Stale())["loop"]);
        AssertTotals(Merge(Stale(), Compacted())["loop"]);
    }
    static void CalibrationRepeated()
    {
        var merged = Merge(Compacted(), Stale());
        for (var i = 0; i < 8; i++)
        {
            merged = Merge(Merge(Stale(), merged), Compacted());
            merged = Merge(merged, merged);
            AssertTotals(merged["loop"]);
        }
    }
    static void CalibrationCutoffs()
    {
        Dictionary<string, LoopCalibrationBucket> Later() => Bucket(new()
        {
            CandidateCount = 14, ConfirmedCount = 7, RejectedCount = 4, InconclusiveCount = 3,
            CompactedCandidateCount = 14, CompactedConfirmedCount = 7, CompactedRejectedCount = 4, CompactedInconclusiveCount = 3,
            CountersCompactedBeforeUtc = Cutoff.AddDays(1)
        });
        foreach (var merged in new[] { Merge(Merge(Compacted(), Stale()), Later()), Merge(Later(), Merge(Stale(), Compacted())) })
        {
            var value = merged["loop"];
            Check(value.CountersCompactedBeforeUtc == Cutoff.AddDays(1) && value.CandidateCount == 14 &&
                value.ConfirmedCount == 7 && value.RejectedCount == 4 && value.InconclusiveCount == 3 &&
                value.CandidateByWriter.Count == 0, "later compacted snapshot was not preserved");
        }
    }
    static void CalibrationUncompacted()
    {
        var result = Merge(Stale(), Bucket(new() { CandidateByWriter = new() { [NewWriter] = 6, ["third"] = 2 } }))["loop"];
        Check(result.CandidateCount == 18 && result.CountersCompactedBeforeUtc is null, "uncompacted maxima changed");
    }
    static void LessonCounters()
    {
        foreach (var reverse in new[] { false, true })
        {
            var compacted = new RecoveryLesson
            {
                Id = "same", SuccessCount = 10, FailureCount = 3, SelectionCount = 13, CumulativeReward = 5.5, RewardObservationCount = 10,
                CompactedSuccessCount = 10, CompactedFailureCount = 3, CompactedSelectionCount = 13,
                CompactedCumulativeReward = 5.5, CompactedRewardObservationCount = 10, CountersCompactedBeforeUtc = Cutoff
            };
            var stale = new RecoveryLesson
            {
                Id = "same", SuccessByWriter = new() { [OldWriter] = 10, [NewWriter] = 1 },
                FailureByWriter = new() { [OldWriter] = 3 }, SelectionByWriter = new() { [OldWriter] = 13, [NewWriter] = 1 },
                RewardByWriter = new() { [OldWriter] = 5.5, [NewWriter] = 0.5 }, RewardObservationByWriter = new() { [OldWriter] = 10, [NewWriter] = 1 }
            };
            var result = (List<RecoveryLesson>)Invoke("MergeRecoveryLessons",
                new List<RecoveryLesson> { reverse ? stale : compacted }, new List<RecoveryLesson> { reverse ? compacted : stale });
            var value = result.Single();
            Check(value.SuccessCount == 11 && value.FailureCount == 3 && value.SelectionCount == 14 &&
                value.CumulativeReward == 6 && value.RewardObservationCount == 11 && value.CountersCompactedBeforeUtc == Cutoff,
                "lesson merge diverged from calibration cutoff rules");
        }
    }
    static void NullableCutoffs()
    {
        Check(RecoveryWriterCounters.LatestCompactionCutoff([]) is null, "empty cutoff set changed");
        Check(RecoveryWriterCounters.LatestCompactionCutoff(null, null) is null, "null cutoffs changed");
        Check(RecoveryWriterCounters.LatestCompactionCutoff(Cutoff, null) == Cutoff &&
            RecoveryWriterCounters.LatestCompactionCutoff(null, Cutoff) == Cutoff &&
            RecoveryWriterCounters.LatestCompactionCutoff(new DateTime?[] { null, Cutoff, Cutoff.AddDays(1), null }) == Cutoff.AddDays(1),
            "nullable maximum is not order independent");
    }
    static QaLocateDto? Validate(QaLocateDto? answer) => QaAnswerValidator.Validate(answer, 320, 200);
    static void EmptyAnswers()
    {
        foreach (var value in new QaLocateDto?[] { null, new(), new() { AnswerText = " \r\n\t" }, new() { Note = "metadata only" } })
            Check(Validate(value) is null, "empty or note-only answer accepted");
    }
    static void InvalidLocations()
    {
        foreach (var value in new QaLocateDto[]
        {
            new() { X = 0.5 }, new() { X = double.NaN, Y = 0 }, new() { X = 0, Y = double.PositiveInfinity },
            new() { X = -0.1, Y = 0 }, new() { X = 1.1, Y = 0 }, new() { XPx = 1 },
            new() { XPx = 320, YPx = 1 }, new() { XPx = 0, YPx = 200 }, new() { XPx = -1, YPx = 0 },
            new() { BBox = new() }, new() { BBox = new() { Left = 2, Top = 2, Right = 1, Bottom = 3 } },
            new() { BBox = new() { Left = 0, Top = 0, Right = 321, Bottom = 200 } }
        }) Check(Validate(value) is null, "invalid geometry accepted as an answer");
    }
    static void ValidLocations()
    {
        foreach (var value in new QaLocateDto[]
        {
            new() { X = 0, Y = 1 }, new() { X = 1, Y = 0 }, new() { XPx = 0, YPx = 199 },
            new() { XPx = 319, YPx = 0 }, new() { BBox = new() { Left = 0, Top = 0, Right = 320, Bottom = 200 } }
        }) Check(Validate(value) is not null, "valid location-only answer rejected");
    }
    static void TextWithInvalidLocation()
    {
        var source = new QaLocateDto { AnswerText = "  answer  ", Note = "  note  ", X = double.NaN, Y = 0, XPx = 320, YPx = 1, BBox = new() };
        var value = Validate(source);
        Check(value is { AnswerText: "answer", Note: "note", X: null, Y: null, XPx: null, YPx: null, BBox: null }, "invalid optional location reached display");
        Check(source.XPx == 320 && source.BBox is not null, "validator mutated the original response");
        value = Validate(new() { X = 2, Y = 0, XPx = 2, YPx = 3 });
        Check(value is { X: null, Y: null, XPx: 2, YPx: 3 }, "valid alternate location was discarded");
    }

    sealed class Capture : DesktopCapture.IBackend
    {
        public Rectangle Bounds(bool multiple) => new(0, 0, 320, 200);
        public void CopyPixels(Graphics graphics, Rectangle region) => graphics.Clear(Color.Navy);
    }
    sealed class Response(QaLocateDto answer) : HttpMessageHandler
    {
        internal int Calls;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            Calls++;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(new
                {
                    id = "synthetic", status = "completed", output_text = JsonSerializer.Serialize(answer)
                }), Encoding.UTF8, "application/json")
            });
        }
    }
    static void BatchQa(QaLocateDto answer, bool expectedSuccess)
    {
        // RunBatchTaskAsync owns BeginRun, so restore configured options outside that scope.
        var options = typeof(AppOptions).GetFields(BindingFlags.Instance | BindingFlags.NonPublic)
            .Where(field => !field.IsLiteral && !field.IsInitOnly).Select(field => (Field: field, Value: field.GetValue(A.Options))).ToArray();
        var oldCapture = DesktopCapture.Backend;
        var oldWindow = ScreenshotService.ActiveWindowRectangleProvider;
        var oldContext = UiAutomationService.PromptContextProvider;
        var oldHttp = OpenAiTransport.HttpClient;
        using var handler = new Response(answer);
        using var client = new HttpClient(handler);
        try
        {
            A.Options.BatchMode = true;
            A.Options.LogScreens = A.Options.LogRequests = false;
            A.Options.IncludeFocusUia = A.Options.IncludeUiaTargets = false;
            A.Options.AutoHideConsoleDuringRun = false;
            A.Options.MaxArtifactsPerDir = 0;
            DesktopCapture.Backend = new Capture();
            ScreenshotService.ActiveWindowRectangleProvider = () => new Rectangle(0, 0, 320, 200);
            UiAutomationService.PromptContextProvider = (_, _, _) => new("Synthetic", "synthetic", null, null, null, null);
            OpenAiTransport.HttpClient = client;
            var exit = A.RunBatchTaskAsync("synthetic-key", "/ask What is displayed?").GetAwaiter().GetResult();
            Check(exit == (expectedSuccess ? 0 : 1), $"unexpected batch exit code {exit}");
            Check(handler.Calls == 1, "unexpected HTTP retries");
        }
        finally
        {
            OpenAiTransport.HttpClient = oldHttp;
            DesktopCapture.Backend = oldCapture;
            ScreenshotService.ActiveWindowRectangleProvider = oldWindow;
            UiAutomationService.PromptContextProvider = oldContext;
            foreach (var option in options) option.Field.SetValue(A.Options, option.Value);
        }
    }
}
