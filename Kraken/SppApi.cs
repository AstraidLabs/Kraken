using System;
using System.Runtime.InteropServices;

namespace Kraken
{
    /// <summary>
    /// Robust P/Invoke wrapper for Software Protection Platform (SPP).
    /// Prefer sppc.dll; fall back to slc.dll. Returned SPP buffers are NOT freed here
    /// (lifetime is tied to the SPP session and reclaimed on SLClose).
    /// </summary>
    public static class SppApi
    {
        private const int E_MOD_NOT_FOUND = unchecked((int)0x8007007E);
        private const int E_PROC_NOT_FOUND = unchecked((int)0x8007007F);

        // ---------------- Safe handle for SPP session ----------------
        public sealed class SppSafeHandle : SafeHandle
        {
            public SppSafeHandle() : base(IntPtr.Zero, ownsHandle: true) { }
            public override bool IsInvalid => handle == IntPtr.Zero;

            internal void Initialize(IntPtr h) => SetHandle(h);

            protected override bool ReleaseHandle()
            {
                try
                {
                    try { NativeSppc.SLClose(handle); }
                    catch (DllNotFoundException) { NativeSlc.SLClose(handle); }
                    catch (EntryPointNotFoundException) { NativeSlc.SLClose(handle); }
                }
                catch { /* ignore */ }
                return true;
            }
        }

        /// <summary>Open SPP session (sppc.dll then slc.dll).</summary>
        public static int SLOpen(out SppSafeHandle hSLC)
        {
            hSLC = new SppSafeHandle();
            IntPtr tmp = IntPtr.Zero;
            int hr;
            try { hr = NativeSppc.SLOpen(ref tmp); }
            catch (DllNotFoundException) { hr = TryOpenSlc(ref tmp); }
            catch (EntryPointNotFoundException) { hr = TryOpenSlc(ref tmp); }

            if (hr == 0 && tmp != IntPtr.Zero) hSLC.Initialize(tmp);
            return hr;

            static int TryOpenSlc(ref IntPtr h)
            {
                try { return NativeSlc.SLOpen(ref h); }
                catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
                catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
            }
        }

        public static bool TryOpenSession(out SppSafeHandle handle)
        {
            var hr = SLOpen(out handle);
            return hr == 0 && handle is not null && !handle.IsInvalid;
        }

        // ---------------- Core list/status APIs ----------------

        /// <summary>Enumerate SLIDs for an application.</summary>
        public static int SLGetSLIDList(SppSafeHandle h, Guid appId, out IntPtr ppGuids, out uint count)
        {
            Ensure(h);
            ppGuids = IntPtr.Zero; count = 0;
            try { return NativeSppc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids); }
            catch (DllNotFoundException) { return CallSlc(() => NativeSlc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids)); }
            catch (EntryPointNotFoundException) { return CallSlc(() => NativeSlc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids)); }
        }

        /// <summary>Get licensing status info array pointer; parse with <see cref="ParseLicensingStatus"/>.</summary>
        public static int SLGetLicensingStatusInformation(SppSafeHandle h, Guid appId, Guid skuId, out IntPtr ppStatus, out uint cStatus)
        {
            Ensure(h);
            ppStatus = IntPtr.Zero; cStatus = 0;
            try { return NativeSppc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus); }
            catch (DllNotFoundException) { return CallSlc(() => NativeSlc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus)); }
            catch (EntryPointNotFoundException) { return CallSlc(() => NativeSlc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus)); }
        }

        // The native status layout is not officially documented; the reference PS script treats each entry as 40 bytes:
        // offset +16: DWORD dwStatus
        // offset +20: DWORD dwGrace (minutes)
        // offset +28: DWORD hrReason
        // offset +32: QWORD qwValidity (FILETIME)
        public readonly record struct SppLicenseStatus(uint Status, uint GraceMinutes, uint ReasonHResult, ulong ValidityFileTimeUtc);

        /// <summary>Parse status entries from native buffer (no frees here).</summary>
        public static SppLicenseStatus[] ParseLicensingStatus(IntPtr pStatus, uint count)
        {
            if (pStatus == IntPtr.Zero || count == 0) return Array.Empty<SppLicenseStatus>();
            var arr = new SppLicenseStatus[count];
            const int stride = 40;

            for (uint i = 0; i < count; i++)
            {
                IntPtr entry = IntPtr.Add(pStatus, (int)(i * stride));
                uint status = (uint)Marshal.ReadInt32(entry, 16);
                uint grace = (uint)Marshal.ReadInt32(entry, 20);
                uint reason = (uint)Marshal.ReadInt32(entry, 28);
                ulong valid = (ulong)(long)Marshal.ReadInt64(entry, 32);

                // Mirror PS fixups:
                if (status == 3) status = 5; // Notification
                if (status == 2)
                {
                    if (reason == 0x4004F00D) status = 3;       // Additional grace (KMS expired) 
                    else if (reason == 0x4004F065) status = 4;  // Non-genuine grace
                    else if (reason == 0x4004FC06) status = 6;  // Extended grace
                }

                arr[i] = new SppLicenseStatus(status, grace, reason, valid);
            }
            return arr;
        }

        // ---------------- Value getters (tData/cData/bData) ----------------

        public enum SppValueKind { Unknown = 0, String = 1, UInt32 = 4, UInt64 = 8 }

        public readonly record struct SppValue(SppValueKind Kind, string? S, uint? U32, ulong? U64)
        {
            public static SppValue FromString(string s) => new(SppValueKind.String, s, null, null);
            public static SppValue FromUInt32(uint v) => new(SppValueKind.UInt32, null, v, null);
            public static SppValue FromUInt64(ulong v) => new(SppValueKind.UInt64, null, null, v);
            public static readonly SppValue Empty = new(SppValueKind.Unknown, null, null, null);
        }

        public static int SLGetPKeyInformation(SppSafeHandle h, Guid pkeyId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            Ensure(h);
            tData = 0; cData = 0; bData = IntPtr.Zero;
            try { return NativeSppc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData); }
            catch (DllNotFoundException) { return CallSlc(() => NativeSlc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData)); }
            catch (EntryPointNotFoundException) { return CallSlc(() => NativeSlc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData)); }
        }

        public static int SLGetProductSkuInformation(SppSafeHandle h, Guid skuId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            Ensure(h);
            tData = 0; cData = 0; bData = IntPtr.Zero;
            try { return NativeSppc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData); }
            catch (DllNotFoundException) { return CallSlc(() => NativeSlc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData)); }
            catch (EntryPointNotFoundException) { return CallSlc(() => NativeSlc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData)); }
        }

        public static int SLGetServiceInformation(SppSafeHandle h, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            Ensure(h);
            tData = 0; cData = 0; bData = IntPtr.Zero;
            try { return NativeSppc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData); }
            catch (DllNotFoundException) { return CallSlc(() => NativeSlc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData)); }
            catch (EntryPointNotFoundException) { return CallSlc(() => NativeSlc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData)); }
        }

        public static int SLGetApplicationInformation(SppSafeHandle h, Guid appId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            Ensure(h);
            tData = 0; cData = 0; bData = IntPtr.Zero;

            try { return NativeSppc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out tData, out cData, out bData); }
            catch (DllNotFoundException) { /* try slc */ }
            catch (EntryPointNotFoundException) { /* try slc */ }

            try { return NativeSlc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out tData, out cData, out bData); }
            catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
        }

        /// <summary>
        /// Interpret (tData,cData,bData) into a managed value. No freeing is performed here.
        /// </summary>
        public static SppValue InterpretValue(uint tData, uint cData, IntPtr bData)
        {
            if (bData == IntPtr.Zero || cData == 0) return SppValue.Empty;

            if (tData == 1)
            {
                // Prefer explicit length if provided; cData is in bytes.
                int chars = (int)(cData / 2);
                string s = chars > 0 ? Marshal.PtrToStringUni(bData, chars).TrimEnd('\0') : Marshal.PtrToStringUni(bData);
                return SppValue.FromString(s ?? string.Empty);
            }
            if (tData == 4)
            {
                return SppValue.FromUInt32((uint)Marshal.ReadInt32(bData));
            }
            if (tData == 3 && cData == 8)
            {
                return SppValue.FromUInt64((ulong)(long)Marshal.ReadInt64(bData));
            }
            return SppValue.Empty;
        }

        // ---------------- Windows info helpers (slc.dll only) ----------------

        public static int SLGetWindowsInformation(string valueName, out uint tData, out uint cData, out IntPtr bData)
        {
            tData = 0; cData = 0; bData = IntPtr.Zero;
            try { return NativeSlc.SLGetWindowsInformation(valueName, out tData, out cData, out bData); }
            catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
        }

        public static int SLGetWindowsInformationDWORD(string valueName, out uint dword)
        {
            dword = 0;
            try { return NativeSlc.SLGetWindowsInformationDWORD(valueName, out dword); }
            catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
        }

        public static int SLIsWindowsGenuineLocal(out uint genuine)
        {
            genuine = 0;
            try { return NativeSlc.SLIsWindowsGenuineLocal(out genuine); }
            catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
        }

        /// <summary>
        /// Generate offline Installation ID; pointer managed by SPP, do not free.
        /// </summary>
        public static string? GenerateOfflineInstallationId(SppSafeHandle h, Guid skuId)
        {
            Ensure(h);
            IntPtr pStr = IntPtr.Zero;
            int hr;
            try { hr = NativeSppc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out pStr); }
            catch (DllNotFoundException) { hr = TrySlcGenerate(h, ref skuId, out pStr); }
            catch (EntryPointNotFoundException) { hr = TrySlcGenerate(h, ref skuId, out pStr); }

            if (hr != 0 || pStr == IntPtr.Zero) return null;
            return Marshal.PtrToStringUni(pStr);

            static int TrySlcGenerate(SppSafeHandle h, ref Guid sku, out IntPtr p)
            {
                p = IntPtr.Zero;
                try { return NativeSlc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref sku, out p); }
                catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
                catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
            }
        }

        public static string? GetWindowsString(string key)
        {
            int hr = SLGetWindowsInformation(key, out uint t, out uint c, out IntPtr p);
            if (hr != 0 || p == IntPtr.Zero || t != 1) return null;

            // c is bytes; if 0, trust null-termination.
            return c > 0 ? Marshal.PtrToStringUni(p, (int)(c / 2))?.TrimEnd('\0') : Marshal.PtrToStringUni(p);
        }

        public static uint? GetWindowsDWord(string key)
        {
            int hr = SLGetWindowsInformationDWORD(key, out uint d);
            return hr == 0 ? d : (uint?)null;
        }

        // ---------------- Subscription (clipc.dll) ----------------

        [StructLayout(LayoutKind.Sequential)]
        public struct SubStatus
        {
            public uint dwEnabled;
            public uint dwSku;
            public uint dwState;
            // Intentionally only the fields present in the PS script.
        }

        [DllImport("clipc.dll", CharSet = CharSet.Unicode)]
        private static extern int ClipGetSubscriptionStatus(ref IntPtr pStatus);

        public static bool TryGetSubscriptionStatus(out SubStatus status)
        {
            status = default;
            IntPtr buf = IntPtr.Zero;
            try
            {
                buf = Marshal.AllocHGlobal(Marshal.SizeOf<SubStatus>());
                IntPtr p = buf;
                int hr = ClipGetSubscriptionStatus(ref p);
                if (hr != 0 || p == IntPtr.Zero) return false;
                status = Marshal.PtrToStructure<SubStatus>(p);
                return true;
            }
            catch { return false; }
            finally
            {
                if (buf != IntPtr.Zero) Marshal.FreeHGlobal(buf);
            }
        }

        // ---------------- Windows 7 genuine (slwga.dll) ----------------
        [DllImport("slwga.dll", CharSet = CharSet.Unicode)]
        public static extern int SLIsGenuineLocal(ref Guid appId, out uint pdwGenuine, IntPtr pvReserved);

        // ---------------- Internals ----------------

        private static void Ensure(SppSafeHandle h)
        {
            if (h is null || h.IsInvalid)
                throw new ArgumentException("Invalid SPP handle.", nameof(h));
        }

        private static int CallSlc(Func<int> f)
        {
            try { return f(); }
            catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
        }

        // ---------------- Native imports ----------------

        internal static class NativeSppc
        {
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLOpen(ref IntPtr hSLC);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLClose(IntPtr hSLC);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid skuId, out IntPtr pwszOfflineId);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetSLIDList(IntPtr hSLC, uint dwFlags, ref Guid appId, uint dwCount, out uint pcSLIDs, out IntPtr ppSLIDs);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pwszProductVersion, out uint pcStatus, out IntPtr ppStatus);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetServiceInformation(IntPtr hSLC, string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        }

        internal static class NativeSlc
        {
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLOpen(ref IntPtr hSLC);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLClose(IntPtr hSLC);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid skuId, out IntPtr pwszOfflineId);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetSLIDList(IntPtr hSLC, uint dwFlags, ref Guid appId, uint dwCount, out uint pcSLIDs, out IntPtr ppSLIDs);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pwszProductVersion, out uint pcStatus, out IntPtr ppStatus);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetServiceInformation(IntPtr hSLC, string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("slc.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.Winapi)] internal static extern int SLIsWindowsGenuineLocal(out uint pdwGenuine);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetWindowsInformation(string valueName, out uint tData, out uint cData, out IntPtr bData);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetWindowsInformationDWORD(string valueName, out uint dword);
            [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        }
    }
}
