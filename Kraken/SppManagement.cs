using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Kraken
{
    /// <summary>
    /// Managed wrapper around <see cref="SppApi"/> that opens and maintains an SPP session.
    /// </summary>
    public sealed class SppManagement : IDisposable
    {
        private readonly SppApi.SppSafeHandle _handle;

        /// <summary>
        /// Initializes a new instance of the <see cref="SppManagement"/> class and opens an SPP session.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown when the session cannot be opened.</exception>
        public SppManagement()
        {
            if (!SppApi.TryOpenSession(out var handle))
            {
                throw new InvalidOperationException("Failed to open SPP session.");
            }

            _handle = handle;
        }

        /// <summary>
        /// Gets a value indicating whether the SPP session is currently open.
        /// </summary>
        public bool IsOpen => !_handle.IsInvalid && !_handle.IsClosed;

        /// <summary>
        /// Enumerates the SLID values for the specified application.
        /// </summary>
        /// <param name="appId">Application identifier.</param>
        /// <returns>Collection of SLID GUIDs.</returns>
        /// <exception cref="COMException">Thrown when the underlying SPP call fails.</exception>
        public IEnumerable<Guid> GetSLIDs(Guid appId)
        {
            int hr = SppApi.SLGetSLIDList(_handle, appId, out var pGuids, out var count);
            try
            {
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                if (count == 0 || pGuids == IntPtr.Zero)
                {
                    return Array.Empty<Guid>();
                }

                var result = new Guid[count];
                int size = Marshal.SizeOf<Guid>();
                for (int i = 0; i < count; i++)
                {
                    IntPtr ptr = IntPtr.Add(pGuids, i * size);
                    result[i] = Marshal.PtrToStructure<Guid>(ptr);
                }
                return result;
            }
            finally
            {
                if (pGuids != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(pGuids);
                }
            }
        }

        /// <summary>
        /// Retrieves licensing status information for the specified application and SKU.
        /// </summary>
        /// <param name="appId">Application identifier.</param>
        /// <param name="skuId">SKU identifier.</param>
        /// <returns>Array of licensing status entries.</returns>
        /// <exception cref="COMException">Thrown when the underlying SPP call fails.</exception>
        public SppApi.SppLicenseStatus[] GetLicensingStatus(Guid appId, Guid skuId)
        {
            int hr = SppApi.SLGetLicensingStatusInformation(_handle, appId, skuId, out var pStatus, out var count);
            try
            {
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                return SppApi.ParseLicensingStatus(pStatus, count);
            }
            finally
            {
                if (pStatus != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(pStatus);
                }
            }
        }

        /// <summary>
        /// Retrieves a named value associated with a product key.
        /// </summary>
        /// <param name="pkeyId">Product key identifier.</param>
        /// <param name="name">Name of the value to retrieve.</param>
        /// <returns>The interpreted value.</returns>
        /// <exception cref="COMException">Thrown when the underlying SPP call fails.</exception>
        public SppApi.SppValue GetPKeyInfo(Guid pkeyId, string name)
        {
            int hr = SppApi.SLGetPKeyInformation(_handle, pkeyId, name, out uint t, out uint c, out IntPtr p);
            try
            {
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                return SppApi.InterpretValue(t, c, p);
            }
            finally
            {
                if (p != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(p);
                }
            }
        }

        /// <summary>
        /// Retrieves a named value associated with a product SKU.
        /// </summary>
        /// <param name="skuId">SKU identifier.</param>
        /// <param name="name">Name of the value to retrieve.</param>
        /// <returns>The interpreted value.</returns>
        /// <exception cref="COMException">Thrown when the underlying SPP call fails.</exception>
        public SppApi.SppValue GetProductSkuInfo(Guid skuId, string name)
        {
            int hr = SppApi.SLGetProductSkuInformation(_handle, skuId, name, out uint t, out uint c, out IntPtr p);
            try
            {
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                return SppApi.InterpretValue(t, c, p);
            }
            finally
            {
                if (p != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(p);
                }
            }
        }

        /// <summary>
        /// Retrieves a named value associated with an application and returns its string representation.
        /// </summary>
        /// <param name="appId">Application identifier.</param>
        /// <param name="name">Name of the value to retrieve.</param>
        /// <returns>The string value if available; otherwise, <c>null</c>.</returns>
        /// <exception cref="COMException">Thrown when the underlying SPP call fails.</exception>
        public string? GetApplicationInfo(Guid appId, string name)
        {
            int hr = SppApi.SLGetApplicationInformation(_handle, appId, name, out uint t, out uint c, out IntPtr p);
            try
            {
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                var v = SppApi.InterpretValue(t, c, p);
                return v.S;
            }
            finally
            {
                if (p != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(p);
                }
            }
        }

        /// <summary>
        /// Generates an offline installation identifier for the specified SKU.
        /// </summary>
        /// <param name="skuId">SKU identifier.</param>
        /// <returns>The generated installation identifier or <c>null</c>.</returns>
        public string? GenerateOfflineInstallationId(Guid skuId) => SppApi.GenerateOfflineInstallationId(_handle, skuId);

        /// <summary>
        /// Retrieves a Windows information string.
        /// </summary>
        /// <param name="key">Registry key name.</param>
        /// <returns>The associated string or <c>null</c>.</returns>
        public string? GetWindowsString(string key) => SppApi.GetWindowsString(key);

        /// <summary>
        /// Retrieves a Windows information DWORD value.
        /// </summary>
        /// <param name="key">Registry key name.</param>
        /// <returns>The associated value or <c>null</c>.</returns>
        public uint? GetWindowsDWord(string key) => SppApi.GetWindowsDWord(key);

        /// <summary>
        /// Extracts the string component from an <see cref="SppApi.SppValue"/>.
        /// </summary>
        /// <param name="v">Value to inspect.</param>
        /// <returns>The string component if present; otherwise, <c>null</c>.</returns>
        public static string? GetString(SppApi.SppValue v) => v.S;

        /// <summary>
        /// Extracts the 32-bit unsigned integer component from an <see cref="SppApi.SppValue"/>.
        /// </summary>
        /// <param name="v">Value to inspect.</param>
        /// <returns>The 32-bit unsigned integer component if present; otherwise, <c>null</c>.</returns>
        public static uint? GetUInt32(SppApi.SppValue v) => v.U32;

        /// <summary>
        /// Extracts the 64-bit unsigned integer component from an <see cref="SppApi.SppValue"/>.
        /// </summary>
        /// <param name="v">Value to inspect.</param>
        /// <returns>The 64-bit unsigned integer component if present; otherwise, <c>null</c>.</returns>
        public static ulong? GetUInt64(SppApi.SppValue v) => v.U64;

        /// <summary>
        /// Releases the underlying SPP session.
        /// </summary>
        public void Dispose()
        {
            _handle.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}

