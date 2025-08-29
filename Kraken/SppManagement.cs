using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Kraken;

/// <summary>
/// High-level wrapper around <see cref="SppApi"/> that manages an SPP session and
/// exposes type-safe helpers mirroring the reference PowerShell script.
/// </summary>
public sealed class SppManagement : IDisposable
{
    private readonly SppApi.SppSafeHandle _handle;
    private bool _disposed;
    private readonly Dictionary<string, object?> _collected = new();

    private SppManagement(SppApi.SppSafeHandle handle) => _handle = handle;

    /// <summary>
    /// Attempts to open a new SPP session.
    /// </summary>
    /// <param name="session">The opened session when successful.</param>
    /// <returns><c>true</c> if the session was opened; otherwise <c>false</c>.</returns>
    public static bool TryOpenSession(out SppManagement? session)
    {
        if (SppApi.TryOpenSession(out var handle))
        {
            session = new SppManagement(handle);
            return true;
        }
        session = null;
        return false;
    }

    /// <summary>
    /// Opens a new SPP session or throws when the platform is unavailable.
    /// </summary>
    public static SppManagement OpenSession()
    {
        if (!TryOpenSession(out var session))
            throw new InvalidOperationException("Cannot open SPP session.");
        return session!;
    }

    /// <summary>
    /// Indicates whether the session is open and usable.
    /// </summary>
    public bool IsOpen => !_disposed && !_handle.IsInvalid && !_handle.IsClosed;

    /// <summary>
    /// Enumerates all SLIDs for the specified application.
    /// </summary>
    public IEnumerable<Guid> GetSLIDs(Guid appId)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetSLIDList(_handle, appId, out var ppGuids, out var count);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            var result = new Guid[count];
            int size = Marshal.SizeOf<Guid>();
            for (int i = 0; i < count; i++)
            {
                IntPtr ptr = IntPtr.Add(ppGuids, i * size);
                result[i] = Marshal.PtrToStructure<Guid>(ptr);
            }
            _collected[$"SLIDs:{appId}"] = result;
            return result;
        }
        finally
        {
            if (ppGuids != IntPtr.Zero)
                Marshal.FreeHGlobal(ppGuids);
        }
    }

    /// <summary>
    /// Retrieves licensing status information for the specified application/SKU pair.
    /// </summary>
    public SppApi.SppLicenseStatus[] GetLicensingStatus(Guid appId, Guid skuId)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetLicensingStatusInformation(_handle, appId, skuId, out var pStatus, out var count);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            var statuses = SppApi.ParseLicensingStatus(pStatus, count);
            _collected[$"LicensingStatus:{appId}:{skuId}"] = statuses;
            return statuses;
        }
        finally
        {
            if (pStatus != IntPtr.Zero)
                Marshal.FreeHGlobal(pStatus);
        }
    }

    /// <summary>
    /// Sets a named licensing status value for an application/SKU pair.
    /// </summary>
    public void SetLicensingStatus(Guid appId, Guid skuId, string name, string value)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLSetLicensingStatusInformation(_handle, appId, skuId, name, value);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected[$"SetLicensingStatus:{appId}:{skuId}:{name}"] = value;
    }

    /// <summary>
    /// Retrieves a named application value as a string.
    /// </summary>
    public string? GetApplicationInfo(Guid appId, string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetApplicationInformation(_handle, appId, name, out uint t, out uint c, out IntPtr b);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            var value = SppApi.InterpretValue(t, c, b);
            _collected[$"ApplicationInfo:{appId}:{name}"] = value.S;
            return value.S;
        }
        finally
        {
            if (b != IntPtr.Zero)
                Marshal.FreeHGlobal(b);
        }
    }

    /// <summary>
    /// Sets a named application value.
    /// </summary>
    public void SetApplicationInfo(Guid appId, string name, string value)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLSetApplicationInformation(_handle, appId, name, value);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected[$"SetApplicationInfo:{appId}:{name}"] = value;
    }

    /// <summary>
    /// Retrieves a named service value as a string.
    /// </summary>
    public string? GetServiceInfo(string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetServiceInformation(_handle, name, out uint t, out uint c, out IntPtr b);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            var value = SppApi.InterpretValue(t, c, b);
            _collected[$"ServiceInfo:{name}"] = value.S;
            return value.S;
        }
        finally
        {
            if (b != IntPtr.Zero)
                Marshal.FreeHGlobal(b);
        }
    }

    /// <summary>
    /// Sets a named service value.
    /// </summary>
    public void SetServiceInfo(string name, string value)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLSetServiceInformation(_handle, name, value);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected[$"SetServiceInfo:{name}"] = value;
    }

    /// <summary>
    /// Retrieves a named product key value.
    /// </summary>
    public SppApi.SppValue GetPKeyInfo(Guid pkeyId, string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetPKeyInformation(_handle, pkeyId, name, out uint t, out uint c, out IntPtr b);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            var value = SppApi.InterpretValue(t, c, b);
            _collected[$"PKeyInfo:{pkeyId}:{name}"] = value;
            return value;
        }
        finally
        {
            if (b != IntPtr.Zero)
                Marshal.FreeHGlobal(b);
        }
    }

    /// <summary>
    /// Sets a named product key value.
    /// </summary>
    public void SetPKeyInfo(Guid pkeyId, string name, string value)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLSetPKeyInformation(_handle, pkeyId, name, value);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected[$"SetPKeyInfo:{pkeyId}:{name}"] = value;
    }

    /// <summary>
    /// Retrieves a named product SKU value.
    /// </summary>
    public SppApi.SppValue GetProductSkuInfo(Guid skuId, string name)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLGetProductSkuInformation(_handle, skuId, name, out uint t, out uint c, out IntPtr b);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            var value = SppApi.InterpretValue(t, c, b);
            _collected[$"ProductSkuInfo:{skuId}:{name}"] = value;
            return value;
        }
        finally
        {
            if (b != IntPtr.Zero)
                Marshal.FreeHGlobal(b);
        }
    }

    /// <summary>
    /// Sets a named product SKU value.
    /// </summary>
    public void SetProductSkuInfo(Guid skuId, string name, string value)
    {
        EnsureNotDisposed();
        int hr = SppApi.SLSetProductSkuInformation(_handle, skuId, name, value);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected[$"SetProductSkuInfo:{skuId}:{name}"] = value;
    }

    /// <summary>
    /// Retrieves a Windows information string.
    /// </summary>
    public string? GetWindowsString(string key)
    {
        int hr = SppApi.SLGetWindowsInformation(key, out uint t, out uint c, out IntPtr b);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        try
        {
            if (t != 1 || b == IntPtr.Zero) return null;
            string s = c > 0 ? Marshal.PtrToStringUni(b, (int)(c / 2))?.TrimEnd('\0') ?? string.Empty : Marshal.PtrToStringUni(b) ?? string.Empty;
            _collected[$"WindowsString:{key}"] = s;
            return s;
        }
        finally
        {
            if (b != IntPtr.Zero)
                Marshal.FreeHGlobal(b);
        }
    }

    /// <summary>
    /// Retrieves a Windows information DWORD value.
    /// </summary>
    public uint? GetWindowsDWord(string key)
    {
        int hr = SppApi.SLGetWindowsInformationDWORD(key, out uint dword);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected[$"WindowsDWord:{key}"] = dword;
        return dword;
    }

    /// <summary>
    /// Generates an offline installation identifier for the specified SKU.
    /// </summary>
    public string? GenerateOfflineInstallationId(Guid skuId)
    {
        EnsureNotDisposed();
        var id = SppApi.GenerateOfflineInstallationId(_handle, skuId);
        _collected[$"OfflineInstallationId:{skuId}"] = id;
        return id;
    }

    /// <summary>Gets the string component of a value.</summary>
    public static string? GetString(SppApi.SppValue v) => v.S;

    /// <summary>Gets the 32-bit unsigned integer component of a value.</summary>
    public static uint? GetUInt32(SppApi.SppValue v) => v.U32;

    /// <summary>Gets the 64-bit unsigned integer component of a value.</summary>
    public static ulong? GetUInt64(SppApi.SppValue v) => v.U64;

    /// <summary>
    /// Writes all collected information to a JSON file.
    /// </summary>
    public void ExportLicenseInfo(string path)
    {
        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(path, JsonSerializer.Serialize(_collected, options));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (!_disposed)
        {
            _handle.Dispose();
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }

    private void EnsureNotDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(SppManagement));
    }
}

