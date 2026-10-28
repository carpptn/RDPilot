
namespace RDPilot.Persistence;

/// <summary>Writes a complete text file before atomically replacing its destination.</summary>
internal static class AtomicTextStore
{
    internal interface IFileOperations
    {
        bool Exists(string path);
        void Write(string path, string text);
        void Replace(string source, string destination, string? backup);
        void Move(string source, string destination);
        void Delete(string path);
    }

    sealed class SystemFiles : IFileOperations
    {
        public bool Exists(string path) => File.Exists(path);
        public void Write(string path, string text) => File.WriteAllText(path,
            text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\n", "\r\n", StringComparison.Ordinal),
            new UTF8Encoding(true));
        public void Replace(string source, string destination, string? backup) => File.Replace(source, destination, backup, true);
        public void Move(string source, string destination) => File.Move(source, destination);
        public void Delete(string path) => File.Delete(path);
    }

    internal static IFileOperations Files { get; set; } = new SystemFiles();

    internal static void Write(string path, string contents, string? backup = null)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        var temporary = $"{path}.{Guid.NewGuid():N}.tmp";
        var files = Files;
        try
        {
            files.Write(temporary, contents);
            if (files.Exists(path)) files.Replace(temporary, path, backup);
            else files.Move(temporary, path);
        }
        finally
        {
            // A failed replacement must never overwrite the existing file or its backup.
            if (files.Exists(temporary))
            {
                try { files.Delete(temporary); }
                catch { /* Preserve the original I/O failure. */ }
            }
        }
    }
}
