using System;
using System.Runtime.InteropServices;

namespace Kraken
{
    /// <summary>
    /// Provides a managed wrapper for Software Protection Platform (SPP) APIs.
    /// </summary>
    public static class SppApi
    {
        private const int ERROR_MOD_NOT_FOUND = unchecked((int)0x8007007E);
        private const int ERROR_PROC_NOT_FOUND = unchecked((int)0x8007007F);

        /// <summary>Represents licensing status information.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct SLSTATUS
        {
            public Guid SkuId;
            public uint eStatus;
            public ulong qwGraceTime;
            public uint hrReason;
        }

        /// <summary>Represents public key information.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct SLGETPKEYINFO
        {
            public Guid PKeyId;
            public Guid SkuId;
        }

        /// <summary>Represents SKU information.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct SLGETSKUID
        {
            public Guid SkuId;
        }

        /// <summary>Represents service information.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct SLGETSERVICEINFO
        {
            public uint cbSize;
            public IntPtr pwszValue;
        }

        /// <summary>Represents application information.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct SLGETAPPLICATIONINFO
        {
            public Guid AppId;
            public uint cbSize;
            public IntPtr pwszValue;
        }

        /// <summary>
        /// Represents a safe handle for an SPP session.
        /// </summary>
        public sealed class SppSafeHandle : SafeHandle
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="SppSafeHandle"/> class.
            /// </summary>
            public SppSafeHandle() : base(IntPtr.Zero, true)
            {
            }

            /// <inheritdoc />
            public override bool IsInvalid => handle == IntPtr.Zero;

            /// <summary>
            /// Initializes the handle with a native value.
            /// </summary>
            /// <param name="h">Native SPP handle.</param>
            internal void Initialize(IntPtr h)
            {
                SetHandle(h);
            }

            /// <inheritdoc />
            protected override bool ReleaseHandle()
            {
                int hr;
                try
                {
                    hr = NativeSppc.SLClose(handle);
                }
                catch (DllNotFoundException)
                {
                    try
                    {
                        hr = NativeSlc.SLClose(handle);
                    }
                    catch (DllNotFoundException)
                    {
                        return false;
                    }
                    catch (EntryPointNotFoundException)
                    {
                        return false;
                    }
                }
                catch (EntryPointNotFoundException)
                {
                    try
                    {
                        hr = NativeSlc.SLClose(handle);
                    }
                    catch (DllNotFoundException)
                    {
                        return false;
                    }
                    catch (EntryPointNotFoundException)
                    {
                        return false;
                    }
                }

                return hr == 0;
            }
        }

        /// <summary>
        /// Opens an SPP session.
        /// </summary>
        /// <param name="hSLC">Receives the session handle on success.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLOpen(out SppSafeHandle hSLC)
        {
            hSLC = new SppSafeHandle();
            IntPtr tmp = IntPtr.Zero;
            int hr;
            try
            {
                hr = NativeSppc.SLOpen(ref tmp);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLOpen(ref tmp);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLOpen(ref tmp);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }

            if (hr == 0 && tmp != IntPtr.Zero)
            {
                hSLC.Initialize(tmp);
            }
            return hr;
        }

        /// <summary>
        /// Attempts to open an SPP session and returns a boolean indicating success.
        /// </summary>
        public static bool TryOpenSession(out SppSafeHandle handle)
        {
            var hr = SLOpen(out handle);
            return hr == 0 && handle != null && !handle.IsInvalid;
        }

        /// <summary>
        /// Retrieves a list of SLIC identifiers.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="appId">Application identifier.</param>
        /// <param name="ppGuids">Receives the GUID array pointer.</param>
        /// <param name="count">Receives the number of GUIDs.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetSLIDList(SppSafeHandle h, Guid appId, out IntPtr ppGuids, out uint count)
        {
            ppGuids = IntPtr.Zero;
            count = 0;
            EnsureHandle(h);
            int hr;
            try
            {
                hr = NativeSppc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            return hr;
        }

        /// <summary>
        /// Retrieves licensing status information.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="appId">Application identifier.</param>
        /// <param name="skuId">SKU identifier.</param>
        /// <param name="ppStatus">Receives the status information pointer.</param>
        /// <param name="cStatus">Receives the number of status entries.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetLicensingStatusInformation(SppSafeHandle h, Guid appId, Guid skuId, out IntPtr ppStatus, out uint cStatus)
        {
            ppStatus = IntPtr.Zero;
            cStatus = 0;
            EnsureHandle(h);
            int hr;
            try
            {
                hr = NativeSppc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            return hr;
        }

        /// <summary>
        /// Gets licensing status structures for a given application and SKU.
        /// </summary>
        public static SLSTATUS[] GetLicensingStatus(SppSafeHandle h, Guid appId, Guid skuId)
        {
            var result = Array.Empty<SLSTATUS>();
            int hr = SLGetLicensingStatusInformation(h, appId, skuId, out var pStatus, out var count);
            if (hr != 0 || pStatus == IntPtr.Zero || count == 0)
            {
                return result;
            }
            try
            {
                int size = Marshal.SizeOf<SLSTATUS>();
                result = new SLSTATUS[count];
                for (int i = 0; i < count; i++)
                {
                    IntPtr item = IntPtr.Add(pStatus, i * size);
                    result[i] = Marshal.PtrToStructure<SLSTATUS>(item);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(pStatus);
            }
            return result;
        }

        /// <summary>
        /// Retrieves public key information.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="pkeyId">Public key identifier.</param>
        /// <param name="name">Value name.</param>
        /// <param name="tData">Receives the data type.</param>
        /// <param name="cData">Receives the data size.</param>
        /// <param name="bData">Receives the data pointer.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetPKeyInformation(SppSafeHandle h, Guid pkeyId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            tData = 0;
            cData = 0;
            bData = IntPtr.Zero;
            EnsureHandle(h);
            int hr;
            try
            {
                hr = NativeSppc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            return hr;
        }

        /// <summary>
        /// Retrieves a public key property as a string.
        /// </summary>
        public static string? GetPKeyInfo(SppSafeHandle h, Guid pkeyId, string name)
        {
            string? result = null;
            int hr = SLGetPKeyInformation(h, pkeyId, name, out uint t, out uint c, out IntPtr data);
            if (hr != 0 || data == IntPtr.Zero)
            {
                return null;
            }
            try
            {
                if (t == 1) // Unicode string
                {
                    result = Marshal.PtrToStringUni(data);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(data);
            }
            return result;
        }

        /// <summary>
        /// Retrieves product SKU information.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="skuId">SKU identifier.</param>
        /// <param name="name">Value name.</param>
        /// <param name="tData">Receives the data type.</param>
        /// <param name="cData">Receives the data size.</param>
        /// <param name="bData">Receives the data pointer.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetProductSkuInformation(SppSafeHandle h, Guid skuId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            tData = 0;
            cData = 0;
            bData = IntPtr.Zero;
            EnsureHandle(h);
            int hr;
            try
            {
                hr = NativeSppc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            return hr;
        }

        /// <summary>
        /// Retrieves a SKU property as a string.
        /// </summary>
        public static string? GetSkuInfo(SppSafeHandle h, Guid skuId, string name)
        {
            string? result = null;
            int hr = SLGetProductSkuInformation(h, skuId, name, out uint t, out uint c, out IntPtr data);
            if (hr != 0 || data == IntPtr.Zero)
            {
                return null;
            }
            try
            {
                if (t == 1)
                {
                    result = Marshal.PtrToStringUni(data);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(data);
            }
            return result;
        }

        /// <summary>
        /// Retrieves service information.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="name">Value name.</param>
        /// <param name="tData">Receives the data type.</param>
        /// <param name="cData">Receives the data size.</param>
        /// <param name="bData">Receives the data pointer.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetServiceInformation(SppSafeHandle h, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            tData = 0;
            cData = 0;
            bData = IntPtr.Zero;
            EnsureHandle(h);
            int hr;
            try
            {
                hr = NativeSppc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            return hr;
        }

        /// <summary>
        /// Retrieves service information as a string.
        /// </summary>
        public static string? GetServiceInfo(SppSafeHandle h, string name)
        {
            string? result = null;
            int hr = SLGetServiceInformation(h, name, out uint t, out uint c, out IntPtr data);
            if (hr != 0 || data == IntPtr.Zero)
            {
                return null;
            }
            try
            {
                if (t == 1)
                {
                    result = Marshal.PtrToStringUni(data);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(data);
            }
            return result;
        }

        /// <summary>
        /// Retrieves application information.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="appId">Application identifier.</param>
        /// <param name="name">Value name.</param>
        /// <param name="tData">Receives the data type.</param>
        /// <param name="cData">Receives the data size.</param>
        /// <param name="bData">Receives the data pointer.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetApplicationInformation(SppSafeHandle h, Guid appId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            tData = 0;
            cData = 0;
            bData = IntPtr.Zero;
            EnsureHandle(h);
            int hr;
            try
            {
                hr = NativeSppc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out tData, out cData, out bData);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out tData, out cData, out bData);
                }
                catch (DllNotFoundException)
                {
                    return ERROR_MOD_NOT_FOUND;
                }
                catch (EntryPointNotFoundException)
                {
                    return ERROR_PROC_NOT_FOUND;
                }
            }
            return hr;
        }

        /// <summary>
        /// Retrieves application information as a string.
        /// </summary>
        public static string? GetApplicationInfo(SppSafeHandle h, Guid appId, string name)
        {
            string? result = null;
            int hr = SLGetApplicationInformation(h, appId, name, out uint t, out uint c, out IntPtr data);
            if (hr != 0 || data == IntPtr.Zero)
            {
                return null;
            }
            try
            {
                if (t == 1)
                {
                    result = Marshal.PtrToStringUni(data);
                }
            }
            finally
            {
                Marshal.FreeHGlobal(data);
            }
            return result;
        }

        /// <summary>
        /// Retrieves Windows information.
        /// </summary>
        /// <param name="valueName">Name of the value.</param>
        /// <param name="tData">Receives the data type.</param>
        /// <param name="cData">Receives the data size.</param>
        /// <param name="bData">Receives the data pointer.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetWindowsInformation(string valueName, out uint tData, out uint cData, out IntPtr bData)
        {
            tData = 0;
            cData = 0;
            bData = IntPtr.Zero;
            int hr;
            try
            {
                hr = NativeSlc.SLGetWindowsInformation(valueName, out tData, out cData, out bData);
            }
            catch (DllNotFoundException)
            {
                return ERROR_MOD_NOT_FOUND;
            }
            catch (EntryPointNotFoundException)
            {
                return ERROR_PROC_NOT_FOUND;
            }
            return hr;
        }

        /// <summary>
        /// Retrieves a Windows DWORD value.
        /// </summary>
        /// <param name="valueName">Name of the value.</param>
        /// <param name="dword">Receives the DWORD.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLGetWindowsInformationDWORD(string valueName, out uint dword)
        {
            dword = 0;
            int hr;
            try
            {
                hr = NativeSlc.SLGetWindowsInformationDWORD(valueName, out dword);
            }
            catch (DllNotFoundException)
            {
                return ERROR_MOD_NOT_FOUND;
            }
            catch (EntryPointNotFoundException)
            {
                return ERROR_PROC_NOT_FOUND;
            }
            return hr;
        }

        /// <summary>
        /// Determines whether Windows is genuine.
        /// </summary>
        /// <param name="genuine">Receives the genuineness state.</param>
        /// <returns>HRESULT of the native call.</returns>
        public static int SLIsWindowsGenuineLocal(out uint genuine)
        {
            genuine = 0;
            int hr;
            try
            {
                hr = NativeSlc.SLIsWindowsGenuineLocal(out genuine);
            }
            catch (DllNotFoundException)
            {
                return ERROR_MOD_NOT_FOUND;
            }
            catch (EntryPointNotFoundException)
            {
                return ERROR_PROC_NOT_FOUND;
            }
            return hr;
        }

        /// <summary>
        /// Returns the offline installation ID string.
        /// </summary>
        /// <param name="h">Session handle.</param>
        /// <param name="skuId">SKU identifier.</param>
        /// <returns>Installation ID string or null on failure.</returns>
        public static string? GenerateOfflineInstallationId(SppSafeHandle h, Guid skuId)
        {
            EnsureHandle(h);
            IntPtr pStr = IntPtr.Zero;
            int hr;
            try
            {
                hr = NativeSppc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out pStr);
            }
            catch (DllNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out pStr);
                }
                catch (DllNotFoundException)
                {
                    return null;
                }
                catch (EntryPointNotFoundException)
                {
                    return null;
                }
            }
            catch (EntryPointNotFoundException)
            {
                try
                {
                    hr = NativeSlc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out pStr);
                }
                catch (DllNotFoundException)
                {
                    return null;
                }
                catch (EntryPointNotFoundException)
                {
                    return null;
                }
            }
            if (hr != 0 || pStr == IntPtr.Zero)
            {
                return null;
            }
            return Marshal.PtrToStringUni(pStr);
        }

        /// <summary>
        /// Gets a Windows string value.
        /// </summary>
        /// <param name="key">Value name.</param>
        /// <returns>String value or null on failure.</returns>
        public static string? GetWindowsString(string key)
        {
            uint tData;
            uint cData;
            IntPtr bData;
            int hr = SLGetWindowsInformation(key, out tData, out cData, out bData);
            if (hr != 0 || tData != 1 || bData == IntPtr.Zero)
            {
                return null;
            }
            return Marshal.PtrToStringUni(bData);
        }

        /// <summary>
        /// Gets a Windows DWORD value.
        /// </summary>
        /// <param name="key">Value name.</param>
        /// <returns>DWORD value or null on failure.</returns>
        public static uint? GetWindowsDWord(string key)
        {
            int hr = SLGetWindowsInformationDWORD(key, out uint dword);
            if (hr != 0)
            {
                return null;
            }
            return dword;
        }

        /// <summary>
        /// Attempts to obtain subscription status.
        /// </summary>
        /// <param name="status">Receives the subscription status.</param>
        /// <returns><c>true</c> on success; otherwise, <c>false</c>.</returns>
        public static bool TryGetSubscriptionStatus(out SubStatus status)
        {
            status = default;
            IntPtr buffer = Marshal.AllocHGlobal(Marshal.SizeOf<SubStatus>());
            try
            {
                IntPtr ptr = buffer;
                int hr = ClipGetSubscriptionStatus(ref ptr);
                if (hr != 0 || ptr == IntPtr.Zero)
                {
                    return false;
                }
                status = Marshal.PtrToStructure<SubStatus>(ptr);
                return true;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        /// <summary>
        /// Ensures the handle is valid.
        /// </summary>
        /// <param name="h">Handle to validate.</param>
        private static void EnsureHandle(SppSafeHandle h)
        {
            if (h == null || h.IsInvalid)
            {
                throw new ArgumentException("Invalid SPP handle.", nameof(h));
            }
        }

        /// <summary>
        /// Determines whether a product is genuine (Windows 7 or earlier).
        /// </summary>
        /// <param name="appId">Application identifier.</param>
        /// <param name="pdwGenuine">Receives the genuineness state.</param>
        /// <param name="pvReserved">Reserved parameter.</param>
        /// <returns>HRESULT of the native call.</returns>
        [DllImport("slwga.dll", CharSet = CharSet.Unicode)]
        public static extern int SLIsGenuineLocal(ref Guid appId, out uint pdwGenuine, IntPtr pvReserved);

        [DllImport("clipc.dll", CharSet = CharSet.Unicode)]
        private static extern int ClipGetSubscriptionStatus(ref IntPtr pStatus);

        /// <summary>
        /// Subscription status information.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct SubStatus
        {
            /// <summary>Indicates whether subscription is enabled.</summary>
            public uint dwEnabled;
            /// <summary>Subscription SKU.</summary>
            public uint dwSku;
            /// <summary>Subscription state.</summary>
            public uint dwState;
        }

        // Native import classes
        internal static class NativeSppc
        {
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLOpen(ref IntPtr hSLC);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLClose(IntPtr hSLC);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid skuId, out IntPtr pwszOfflineId);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetSLIDList(IntPtr hSLC, uint dwFlags, ref Guid appId, uint dwCount, out uint pcSLIDs, out IntPtr ppSLIDs);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pwszProductVersion, out uint pcStatus, out IntPtr ppStatus);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetServiceInformation(IntPtr hSLC, string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("sppc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        }

        internal static class NativeSlc
        {
            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLOpen(ref IntPtr hSLC);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLClose(IntPtr hSLC);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid skuId, out IntPtr pwszOfflineId);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetSLIDList(IntPtr hSLC, uint dwFlags, ref Guid appId, uint dwCount, out uint pcSLIDs, out IntPtr ppSLIDs);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pwszProductVersion, out uint pcStatus, out IntPtr ppStatus);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetServiceInformation(IntPtr hSLC, string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("slc.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Winapi)]
            internal static extern int SLIsWindowsGenuineLocal(out uint pdwGenuine);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetWindowsInformation(string valueName, out uint tData, out uint cData, out IntPtr bData);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetWindowsInformationDWORD(string valueName, out uint dword);

            [DllImport("slc.dll", CharSet = CharSet.Unicode)]
            internal static extern int SLGetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        }
    }
}
