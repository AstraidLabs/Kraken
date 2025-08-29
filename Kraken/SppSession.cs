using System;
using Serilog;

namespace Kraken;

/// <summary>
/// Represents a managed Software Protection Platform session.
/// </summary>
public sealed class SppSession : IDisposable
{
    private readonly SppApi.SppSafeHandle _handle;
    private bool _disposed;

    private SppSession(SppApi.SppSafeHandle handle)
    {
        _handle = handle;
    }

    /// <summary>
    /// Opens a new SPP session.
    /// </summary>
    /// <exception cref="SppException">Thrown when the session cannot be opened.</exception>
    public static SppSession Open()
    {
        if (!SppApi.TryOpenSession(out var handle))
        {
            Log.Error("Failed to open SPP session");
            throw new SppException("Cannot open SPP session.");
        }

        Log.Information("SPP session opened");
        return new SppSession(handle);
    }

    /// <summary>
    /// Gets the native SPP handle.
    /// </summary>
    public SppApi.SppSafeHandle Handle => _handle;

    /// <inheritdoc />
    public void Dispose()
    {
        if (!_disposed)
        {
            _handle.Dispose();
            _disposed = true;
            Log.Information("SPP session closed");
        }
    }
}

