namespace RDPilot.Persistence;

/// <summary>Stages a replacement and reverses completed rotation moves on I/O failure. The caller owns the file-family lock.</summary>
internal static class RotatingFileStore
{
    internal interface IFileOperations
    {
        bool Exists(string path);
        bool IsDirectory(string path);
        void Write(string path, byte[] contents);
        void Move(string source, string destination);
        void Replace(string source, string destination);
        void Delete(string path);
    }

    sealed class SystemFiles : IFileOperations
    {
        public bool Exists(string path) => File.Exists(path);
        public bool IsDirectory(string path) => Directory.Exists(path);
        public void Write(string path, byte[] contents) => File.WriteAllBytes(path, contents);
        public void Move(string source, string destination) => File.Move(source, destination);
        public void Replace(string source, string destination) => File.Move(source, destination, overwrite: true);
        public void Delete(string path) => File.Delete(path);
    }

    internal static void Write(string path, byte[] contents, IReadOnlyList<string> generations,
        IFileOperations? operations = null)
    {
        var files = operations ?? new SystemFiles();
        var family = new[] { path }.Concat(generations).ToArray();
        foreach (var member in family)
            if (files.IsDirectory(member))
                throw new IOException($"Storage destination is a directory: {member}");

        var token = Guid.NewGuid().ToString("N");
        var temporary = $"{path}.{token}.tmp";
        var displaced = $"{path}.{token}.rollback";
        var moves = new List<(string Source, string Destination)>();
        void Move(string source, string destination)
        {
            files.Move(source, destination);
            moves.Add((source, destination));
        }

        try
        {
            // No retained data is touched until the complete replacement has been written.
            files.Write(temporary, contents);
            if (generations.Count == 0)
            {
                files.Replace(temporary, path);
                return;
            }

            try
            {
                if (files.Exists(family[^1])) Move(family[^1], displaced);
                for (var index = family.Length - 2; index >= 0; index--)
                    if (files.Exists(family[index])) Move(family[index], family[index + 1]);
                files.Move(temporary, path);
            }
            catch (Exception failure)
            {
                var errors = new List<Exception> { failure };
                foreach (var move in moves.AsEnumerable().Reverse())
                {
                    try { files.Move(move.Destination, move.Source); }
                    catch (Exception rollbackFailure) { errors.Add(rollbackFailure); }
                }
                if (errors.Count > 1)
                    throw new AggregateException(
                        $"Rotation failed and rollback was incomplete. Retained data was not deleted; inspect {string.Join(", ", family)} and {displaced} before retrying.",
                        errors);
                throw;
            }

            // Only a committed replacement permits dropping the oldest generation.
            Cleanup(displaced);
        }
        finally
        {
            Cleanup(temporary);
        }

        void Cleanup(string candidate)
        {
            try { if (files.Exists(candidate)) files.Delete(candidate); }
            catch (Exception ex)
            {
                // Cleanup must not turn an already committed write into a reported failure.
                Console.WriteLine($"[storage] could not remove temporary file {candidate}: {ex.Message}");
            }
        }
    }
}