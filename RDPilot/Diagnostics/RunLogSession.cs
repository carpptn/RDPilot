namespace RDPilot.Diagnostics;

/// <summary>Restores process console writers even when run cleanup throws.</summary>
internal sealed class RunLogSession : IDisposable
{
    readonly TextWriter previousOutput = Console.Out;
    readonly TextWriter previousError = Console.Error;
    readonly StreamWriter file;
    readonly TeeTextWriter output;
    readonly TeeTextWriter error;

    internal RunLogSession(string path)
    {
        file = new StreamWriter(path, append: false, new UTF8Encoding(true)) { AutoFlush = true };
        var synchronized = TextWriter.Synchronized(file);
        output = new TeeTextWriter(previousOutput, synchronized);
        error = new TeeTextWriter(previousError, synchronized);
        Console.SetOut(output);
        Console.SetError(error);
    }

    public void Dispose()
    {
        Console.SetOut(previousOutput);
        Console.SetError(previousError);
        output.Dispose();
        error.Dispose();
        file.Dispose();
    }
}
