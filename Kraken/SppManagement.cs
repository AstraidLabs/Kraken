using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Kraken;

/// <summary>
/// High‑level, disposable wrapper around <see cref="SppApi"/> that manages an SPP session
/// and exposes a convenient, type‑safe API for licence data.
/// </summary>
public sealed class SppManagement : IDisposable
{
    private readonly SppApi.SppSafeHandle _handle;
    private bool _disposed;

    /// <summary>
    /// Opens a new SPP session. Throws <see cref="InvalidOperationException"/> if the session cannot be opened.
    /// </summary>
    public SppManagement()
    {
        if (!SppApi.TryOpenSession(out var handle))
        {
            throw new InvalidOperationException("Cannot open SPP session.");
        }

        _handle = handle;
    }

    /// <summary>
    /// Indicates whether the underlying SPP session is still open.
    /// </summary>
    public bool IsOpen => !_disposed && !_handle.IsInvalid && !_handle.IsClosed;

    #region Core data retrieval

    /// <summary>
    /// Enumerates all SLID values (GUIDs) associated with the specified application.
    /// </summary>
    public IEnumerable<Guid> GetSLIDs(Guid appId)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetSLIDList(_handle, appId, out var ppGuids, out var count);

        if (hr != 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        // No SLIDs – return empty sequence
        if (count == 0 || ppGuids == IntPtr.Zero)
        {
            return Array.Empty<Guid>();
        }

        // Copy the native GUID array into managed memory
        var result = new Guid[count];
        int size = Marshal.SizeOf<Guid>();
        for (int i = 0; i < count; i++)
        {
            IntPtr ptr = IntPtr.Add(ppGuids, i * size);
            result[i] = Marshal.PtrToStructure<Guid>(ptr);
        }

        // Free the native buffer – SPP allocates it via CoTaskMem / GlobalAlloc
        Marshal.FreeHGlobal(ppGuids);
        return result;
    }

    /// <summary>
    /// Retrieves licensing status information for a specific application/SKU pair.
    /// </summary>
    public SppApi.SppLicenseStatus[] GetLicensingStatus(Guid appId, Guid skuId)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetLicensingStatusInformation(_handle, appId, skuId, out var ppStatus, out var count);

        if (hr != 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        if (count == 0 || ppStatus == IntPtr.Zero)
        {
            return Array.Empty<SppApi.SppLicenseStatus>();
        }

        var statuses = SppApi.ParseLicensingStatus(ppStatus, count);
        Marshal.FreeHGlobal(ppStatus);
        return statuses;
    }

    /// <summary>
    /// Retrieves a named value (string/uint/ulong) associated with a product key.
    /// </summary>
    public SppApi.SppValue GetPKeyInfo(Guid pkeyId, string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetPKeyInformation(_handle, pkeyId, name,
                                               out uint t, out uint c, out IntPtr b);

        if (hr != 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        var value = SppApi.InterpretValue(t, c, b);
        Marshal.FreeHGlobal(b);
        return value;
    }

    /// <summary>
    /// Retrieves a named value associated with a product SKU.
    /// </summary>
    public SppApi.SppValue GetProductSkuInfo(Guid skuId, string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetProductSkuInformation(_handle, skuId, name,
                                                     out uint t, out uint c, out IntPtr b);

        if (hr != 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        var value = SppApi.InterpretValue(t, c, b);
        Marshal.FreeHGlobal(b);
        return value;
    }

    /// <summary>
    /// Retrieves a named value from an application and returns its string representation.
    /// </summary>
    public string? GetApplicationInfo(Guid appId, string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetApplicationInformation(_handle, appId, name,
                                                      out uint t, out uint c, out IntPtr b);

        if (hr != 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        var value = SppApi.InterpretValue(t, c, b);
        Marshal.FreeHGlobal(b);
        return value.S;
    }

    /// <summary>
    /// Generates an offline installation identifier for the specified SKU.
    /// </summary>
    public string? GenerateOfflineInstallationId(Guid skuId) =>
        SppApi.GenerateOfflineInstallationId(_handle, skuId);

    #endregion

    #region Windows information helpers

    /// <summary>
    /// Retrieves a Windows string value from the SPP API.
    /// </summary>
    public string? GetWindowsString(string key) => SppApi.GetWindowsString(key);

    /// <summary>
    /// Retrieves a Windows DWORD value from the SPP API.
    /// </summary>
    public uint? GetWindowsDWord(string key) => SppApi.GetWindowsDWord(key);

    #endregion

    #region Value extraction helpers

    /// <summary>Gets the string component of an <see cref="SppApi.SppValue"/>.</summary>
    public static string? GetString(SppApi.SppValue v) => v.S;

    /// <summary>Gets the 32‑bit unsigned integer component of an <see cref="SppApi.SppValue"/>.</summary>
    public static uint? GetUInt32(SppApi.SppValue v) => v.U32;

    /// <summary>Gets the 64‑bit unsigned integer component of an <see cref="SppApi.SppValue"/>.</summary>
    public static ulong? GetUInt64(SppApi.SppValue v) => v.U64;

    #endregion

    #region IDisposable implementation

    /// <summary>
    /// Releases the underlying SPP session. After calling <c>Dispose</c>, the instance
    /// cannot be used again and <see cref="IsOpen"/> will return <c>false</c>.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _handle?.Dispose();   // triggers SPP's SLClose
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }

    #endregion

    #region Internals

    private void EnsureNotDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(SppManagement));
        }
    }

    #endregion
}
