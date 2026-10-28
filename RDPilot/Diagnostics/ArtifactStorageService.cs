
namespace RDPilot.Diagnostics;

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
                if (Options.MaxArtifactsPerDir <= 0)
                    return;
        
                try
                {
                    var files = Directory.GetFiles(dir, pattern)
                        .Select(p => new FileInfo(p))
                        .OrderByDescending(f => f.LastWriteTimeUtc)
                        .GroupBy(ArtifactRunKey)
                        .OrderByDescending(g => g.Max(f => f.LastWriteTimeUtc))
                        .Skip(Options.MaxArtifactsPerDir)
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
                if (Options.MaxArtifactsPerDir <= 0)
                    return;
        
                try
                {
                    var files = Directory.GetFiles(dir, "*.*")
                        .Where(IsScreenImageArtifact)
                        .Select(p => new FileInfo(p))
                        .GroupBy(ArtifactRunKey)
                        .OrderByDescending(g => g.Max(f => f.LastWriteTimeUtc))
                        .Skip(Options.MaxArtifactsPerDir)
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
        
            internal static void SaveRequestLog(string path, object request)
            {
                if (!Options.LogRequests) return;
                var images = new Dictionary<string, string>(StringComparer.Ordinal);
                var imageIndex = 0;
                var redacted = CreateRequestLog(request, dataUrl =>
                {
                    if (!Options.LogScreens) return LogImageRef(null);
                    if (images.TryGetValue(dataUrl, out var existing)) return existing;
                    var separator = dataUrl.IndexOf(',');
                    if (separator < 0 || !dataUrl.StartsWith("data:image/", StringComparison.Ordinal))
                        return "(image omitted)";
                    var extension = dataUrl.StartsWith("data:image/jpeg;", StringComparison.Ordinal) ? ".jpg" : ".png";
                    var directory = Path.Combine(AppContext.BaseDirectory, "screens");
                    Directory.CreateDirectory(directory);
                    var imagePath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(path)}_input_{++imageIndex}{extension}");
                    File.WriteAllBytes(imagePath, Convert.FromBase64String(dataUrl[(separator + 1)..]));
                    var reference = LogImageRef(imagePath);
                    images[dataUrl] = reference;
                    return reference;
                });
                SaveJson(path, redacted);
            }

            internal static JsonNode CreateRequestLog(object request, Func<string, string> imageReference)
            {
                var copy = JsonSerializer.SerializeToNode(request)!;
                Redact(copy);
                return copy;

                void Redact(JsonNode? node)
                {
                    if (node is JsonObject obj)
                    {
                        foreach (var pair in obj.ToArray())
                        {
                            if (pair.Key == "image_url" && pair.Value is JsonValue value && value.TryGetValue<string>(out var url))
                                obj[pair.Key] = imageReference(url);
                            else Redact(pair.Value);
                        }
                    }
                    else if (node is JsonArray array)
                        foreach (var item in array) Redact(item);
                }
            }

            internal static void SaveJson(string path, object o)
            {
                if (!Options.LogRequests) return;
                var sw = RuntimeClock.StartTimer();
                try
                {
                    File.WriteAllBytes(path, JsonSerializer.SerializeToUtf8Bytes(o, Options.PrettyRequestLogs ? PrettyJson : null));
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunArtifactLogWrites++;
                    CurrentRun.RunArtifactLogElapsed += sw.Elapsed;
                }
            }
        
            internal static void SaveRaw(string path, string raw)
            {
                if (!Options.LogRequests) return;
                var sw = RuntimeClock.StartTimer();
                try
                {
                    File.WriteAllText(path, raw, Encoding.UTF8);
                }
                finally
                {
                    sw.Stop();
                    CurrentRun.RunArtifactLogWrites++;
                    CurrentRun.RunArtifactLogElapsed += sw.Elapsed;
                }
            }
        
            internal static string LogImageRef(string? path) =>
                string.IsNullOrWhiteSpace(path) ? "(not saved; base64 omitted)" : $"file://{path}";
        
            internal static string ShotLabel(string? path, string commandId, int step) =>
                string.IsNullOrWhiteSpace(path) ? $"{commandId}_{step}.png (not saved)" : Path.GetFileName(path);
    }
