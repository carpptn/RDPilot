internal static partial class RDPilotApplication
{
    /// <summary>
    /// Persists and prunes screenshots, request payloads, responses, and run logs.
    /// </summary>
    internal static class ArtifactStorageService
    {
            // subfolders
            internal static string EnsureScreensDir()
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "screens");
                Directory.CreateDirectory(dir);
                CleanupScreenArtifacts(dir);
                return dir;
            }
            internal static string EnsureRequestsDir()
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "requests");
                Directory.CreateDirectory(dir);
                CleanupArtifacts(dir, "*.json");
                return dir;
            }
            internal static string EnsureLogDir()
            {
                var dir = Path.Combine(AppContext.BaseDirectory, "logs");
                Directory.CreateDirectory(dir);
                CleanupArtifacts(dir, "*.log");
                return dir;
            }
        
            internal static void CleanupArtifacts(string dir, string pattern)
            {
                if (MaxArtifactsPerDir <= 0)
                    return;
        
                try
                {
                    var files = Directory.GetFiles(dir, pattern)
                        .Select(p => new FileInfo(p))
                        .OrderByDescending(f => f.LastWriteTimeUtc)
                        .GroupBy(ArtifactRunKey)
                        .OrderByDescending(g => g.Max(f => f.LastWriteTimeUtc))
                        .Skip(MaxArtifactsPerDir)
                        .SelectMany(g => g)
                        .ToArray();
        
                    foreach (var file in files)
                        file.Delete();
                }
                catch
                {
                    // Best effort cleanup only.
                }
            }
        
            internal static void CleanupScreenArtifacts(string dir)
            {
                if (MaxArtifactsPerDir <= 0)
                    return;
        
                try
                {
                    var files = Directory.GetFiles(dir, "*.*")
                        .Where(IsScreenImageArtifact)
                        .Select(p => new FileInfo(p))
                        .GroupBy(ArtifactRunKey)
                        .OrderByDescending(g => g.Max(f => f.LastWriteTimeUtc))
                        .Skip(MaxArtifactsPerDir)
                        .SelectMany(g => g)
                        .ToArray();
        
                    foreach (var file in files)
                        file.Delete();
                }
                catch
                {
                    // Best effort cleanup only.
                }
            }
        
            internal static bool IsScreenImageArtifact(string path)
            {
                var ext = Path.GetExtension(path);
                return ext.Equals(".png", StringComparison.OrdinalIgnoreCase) ||
                       ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                       ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase);
            }
        
            internal static string ArtifactRunKey(FileInfo file)
            {
                var name = Path.GetFileNameWithoutExtension(file.Name);
                var underscore = name.IndexOf('_');
                return underscore > 0 ? name[..underscore] : name;
            }
        
            internal static void SaveJson(string path, object o)
            {
                if (!LogRequests) return;
                var sw = Stopwatch.StartNew();
                try
                {
                    File.WriteAllBytes(path, JsonSerializer.SerializeToUtf8Bytes(o, PrettyRequestLogs ? PrettyJson : null));
                }
                finally
                {
                    sw.Stop();
                    RunArtifactLogWrites++;
                    RunArtifactLogElapsed += sw.Elapsed;
                }
            }
        
            internal static void SaveRaw(string path, string raw)
            {
                if (!LogRequests) return;
                var sw = Stopwatch.StartNew();
                try
                {
                    File.WriteAllText(path, raw, Encoding.UTF8);
                }
                finally
                {
                    sw.Stop();
                    RunArtifactLogWrites++;
                    RunArtifactLogElapsed += sw.Elapsed;
                }
            }
        
            internal static string LogImageRef(string? path) =>
                string.IsNullOrWhiteSpace(path) ? "(not saved; base64 omitted)" : $"file://{path}";
        
            internal static string ShotLabel(string? path, string commandId, int step) =>
                string.IsNullOrWhiteSpace(path) ? $"{commandId}_{step}.png (not saved)" : Path.GetFileName(path);
    }
}

