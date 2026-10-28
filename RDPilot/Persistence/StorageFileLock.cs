namespace RDPilot.Persistence;

/// <summary>Synchronous, thread-affine lease keyed by the file actually being modified.</summary>
internal sealed class StorageFileLock : IDisposable
{
    readonly Mutex mutex;

    StorageFileLock(Mutex mutex) => this.mutex = mutex;

    internal static Mutex Create(string path)
    {
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
            RecoveryStoragePaths.ResolvePath(path).ToUpperInvariant())));
        return new Mutex(false, $@"Local\RDPilotRecoveryMemory_{hash[..24]}");
    }

    internal static bool TryEnter(Mutex mutex)
    {
        try { return mutex.WaitOne(TimeSpan.FromSeconds(5)); }
        catch (AbandonedMutexException) { return true; }
    }

    internal static StorageFileLock Acquire(string path)
    {
        var mutex = Create(path);
        try
        {
            if (!TryEnter(mutex))
                throw new IOException($"Could not acquire the storage lock for '{path}'.");
            return new(mutex);
        }
        catch { mutex.Dispose(); throw; }
    }

    public void Dispose()
    {
        try { mutex.ReleaseMutex(); }
        finally { mutex.Dispose(); }
    }
}
