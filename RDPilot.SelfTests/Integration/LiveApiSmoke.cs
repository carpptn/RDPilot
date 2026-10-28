using System.Drawing;
using A = RDPilot.RDPilotApplication;

namespace RDPilot.SelfTests.Integration;

/// <summary>Opt-in, paid API check using only a synthetic image and the existing environment key.</summary>
internal static class LiveApiSmoke
{
    internal static async Task<int> RunAsync()
    {
        var key = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        if (string.IsNullOrWhiteSpace(key))
        {
            Console.Error.WriteLine("Live smoke test requires OPENAI_API_KEY.");
            return 1;
        }
        using var scope = A.BeginRun();
        A.Options.LogRequests = false;
        A.Options.LogScreens = false;
        A.Options.AllowWebSearch = false;
        A.Options.OpenAiMaxRetries = 0;
        A.Options.IncompleteMaxOutputRetries = 0;
        A.Options.OpenAiTimeoutSeconds = 60;
        A.Options.ReasoningEffort = "high";
        using var bitmap = new Bitmap(320, 200);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.FillRectangle(Brushes.Red, 80, 50, 160, 100);
        }
        var image = ScreenshotService.EncodeBitmapToDataUrl(bitmap, 0, "png", 80);
        var request = PromptAndRequestFactory.BuildQARequestBody(
            A.Options.Model, PromptAndRequestFactory.BuildQaSystemRules(),
            "What color is the rectangle? Answer in English using one word.",
            image, 320, 200, 0, 0, 0, 0, null, null,
            new("Synthetic test image", "RDPilot.SelfTests", null, null, null, null));
        var (answer, _) = await OpenAiResponsesService.CallOpenAIParsedAsync<QaLocateDto>(key, request);
        if (answer?.AnswerText?.Contains("red", StringComparison.OrdinalIgnoreCase) != true)
        {
            Console.Error.WriteLine("FAIL live API smoke: no correct structured answer.");
            return 1;
        }
        Console.WriteLine($"PASS live API smoke: {A.Options.Model}, reasoning=high, synthetic image recognized.");
        return 0;
    }
}
