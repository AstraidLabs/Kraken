using System.Runtime.InteropServices;

namespace Kraken.SppSdk;

/// <summary>
/// P/Invoke wrapper for the Software Protection Platform (SPP).
/// Functions are attempted first on sppc.dll and fall back to slc.dll when
/// sppc.dll is not available.
/// </summary>
public static class SppApi
{
    private const int E_MOD_NOT_FOUND = unchecked((int)0x8007007E);
    private const int E_PROC_NOT_FOUND = unchecked((int)0x8007007F);

    // ---------------- Safe handle ----------------
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

    // ---------------- Session ----------------
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

    // ---------------- Core list/status APIs ----------------
    public static int SLGetSLIDList(SppSafeHandle h, Guid appId, out IntPtr ppGuids, out uint count)
    {
        Ensure(h);
        ppGuids = IntPtr.Zero; count = 0;
        try { return NativeSppc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out count, out ppGuids); }
        catch (DllNotFoundException) { return TrySlc(h, ref appId, out ppGuids, out count); }
        catch (EntryPointNotFoundException) { return TrySlc(h, ref appId, out ppGuids, out count); }

        static int TrySlc(SppSafeHandle h, ref Guid appId, out IntPtr ppGuids, out uint count)
        {
            IntPtr p = IntPtr.Zero; uint c = 0;
            try
            {
                int hr = NativeSlc.SLGetSLIDList(h.DangerousGetHandle(), 0, ref appId, 1, out c, out p);
                ppGuids = p; count = c; return hr;
            }
            catch (DllNotFoundException) { ppGuids = IntPtr.Zero; count = 0; return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { ppGuids = IntPtr.Zero; count = 0; return E_PROC_NOT_FOUND; }
        }
    }

    public static int SLGetLicensingStatusInformation(SppSafeHandle h, Guid appId, Guid skuId, out IntPtr ppStatus, out uint cStatus)
    {
        Ensure(h);
        ppStatus = IntPtr.Zero; cStatus = 0;
        try { return NativeSppc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out cStatus, out ppStatus); }
        catch (DllNotFoundException) { return TrySlc(h, ref appId, ref skuId, out ppStatus, out cStatus); }
        catch (EntryPointNotFoundException) { return TrySlc(h, ref appId, ref skuId, out ppStatus, out cStatus); }

        static int TrySlc(SppSafeHandle h, ref Guid appId, ref Guid skuId, out IntPtr ppStatus, out uint cStatus)
        {
            IntPtr p = IntPtr.Zero; uint c = 0;
            try
            {
                int hr = NativeSlc.SLGetLicensingStatusInformation(h.DangerousGetHandle(), ref appId, ref skuId, IntPtr.Zero, out c, out p);
                ppStatus = p; cStatus = c; return hr;
            }
            catch (DllNotFoundException) { ppStatus = IntPtr.Zero; cStatus = 0; return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { ppStatus = IntPtr.Zero; cStatus = 0; return E_PROC_NOT_FOUND; }
        }
    }

    public static int SLGetPKeyInformation(SppSafeHandle h, Guid pkeyId, string name, out uint tData, out uint cData, out IntPtr bData)
    {
        Ensure(h);
        tData = 0; cData = 0; bData = IntPtr.Zero;
        try { return NativeSppc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out tData, out cData, out bData); }
        catch (DllNotFoundException) { return TrySlc(h, ref pkeyId, name, out tData, out cData, out bData); }
        catch (EntryPointNotFoundException) { return TrySlc(h, ref pkeyId, name, out tData, out cData, out bData); }

        static int TrySlc(SppSafeHandle h, ref Guid pkeyId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            uint t = 0; uint c = 0; IntPtr b = IntPtr.Zero;
            try
            {
                int hr = NativeSlc.SLGetPKeyInformation(h.DangerousGetHandle(), ref pkeyId, name, out t, out c, out b);
                tData = t; cData = c; bData = b; return hr;
            }
            catch (DllNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_PROC_NOT_FOUND; }
        }
    }

    public static int SLGetProductSkuInformation(SppSafeHandle h, Guid skuId, string name, out uint tData, out uint cData, out IntPtr bData)
    {
        Ensure(h);
        tData = 0; cData = 0; bData = IntPtr.Zero;
        try { return NativeSppc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out tData, out cData, out bData); }
        catch (DllNotFoundException) { return TrySlc(h, ref skuId, name, out tData, out cData, out bData); }
        catch (EntryPointNotFoundException) { return TrySlc(h, ref skuId, name, out tData, out cData, out bData); }

        static int TrySlc(SppSafeHandle h, ref Guid skuId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            uint t = 0; uint c = 0; IntPtr b = IntPtr.Zero;
            try
            {
                int hr = NativeSlc.SLGetProductSkuInformation(h.DangerousGetHandle(), ref skuId, name, out t, out c, out b);
                tData = t; cData = c; bData = b; return hr;
            }
            catch (DllNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_PROC_NOT_FOUND; }
        }
    }

    public static int SLGetServiceInformation(SppSafeHandle h, string name, out uint tData, out uint cData, out IntPtr bData)
    {
        Ensure(h);
        tData = 0; cData = 0; bData = IntPtr.Zero;
        try { return NativeSppc.SLGetServiceInformation(h.DangerousGetHandle(), name, out tData, out cData, out bData); }
        catch (DllNotFoundException) { return TrySlc(h, name, out tData, out cData, out bData); }
        catch (EntryPointNotFoundException) { return TrySlc(h, name, out tData, out cData, out bData); }

        static int TrySlc(SppSafeHandle h, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            uint t = 0; uint c = 0; IntPtr b = IntPtr.Zero;
            try
            {
                int hr = NativeSlc.SLGetServiceInformation(h.DangerousGetHandle(), name, out t, out c, out b);
                tData = t; cData = c; bData = b; return hr;
            }
            catch (DllNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_PROC_NOT_FOUND; }
        }
    }

    public static int SLGetApplicationInformation(SppSafeHandle h, Guid appId, string name, out uint tData, out uint cData, out IntPtr bData)
    {
        Ensure(h);
        tData = 0; cData = 0; bData = IntPtr.Zero;
        try { return NativeSppc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out tData, out cData, out bData); }
        catch (DllNotFoundException) { return TrySlc(h, ref appId, name, out tData, out cData, out bData); }
        catch (EntryPointNotFoundException) { return TrySlc(h, ref appId, name, out tData, out cData, out bData); }

        static int TrySlc(SppSafeHandle h, ref Guid appId, string name, out uint tData, out uint cData, out IntPtr bData)
        {
            uint t = 0; uint c = 0; IntPtr b = IntPtr.Zero;
            try
            {
                int hr = NativeSlc.SLGetApplicationInformation(h.DangerousGetHandle(), ref appId, name, out t, out c, out b);
                tData = t; cData = c; bData = b; return hr;
            }
            catch (DllNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_MOD_NOT_FOUND; }
            catch (EntryPointNotFoundException) { tData = 0; cData = 0; bData = IntPtr.Zero; return E_PROC_NOT_FOUND; }
        }
    }

    public static int SLGetWindowsInformation(string name, out uint tData, out uint cData, out IntPtr bData)
    {
        tData = 0; cData = 0; bData = IntPtr.Zero;
        try { return NativeSlc.SLGetWindowsInformation(name, out tData, out cData, out bData); }
        catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
        catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
    }

    public static int SLGetWindowsInformationDWORD(string name, out uint dword)
    {
        dword = 0;
        try { return NativeSlc.SLGetWindowsInformationDWORD(name, out dword); }
        catch (DllNotFoundException) { return E_MOD_NOT_FOUND; }
        catch (EntryPointNotFoundException) { return E_PROC_NOT_FOUND; }
    }

    public static int GenerateOfflineInstallationId(SppSafeHandle h, Guid skuId, out string? offlineId)
    {
        Ensure(h);
        offlineId = null;
        IntPtr p = IntPtr.Zero;
        try
        {
            int hr;
            try { hr = NativeSppc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out p); }
            catch (DllNotFoundException) { hr = NativeSlc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out p); }
            catch (EntryPointNotFoundException) { hr = NativeSlc.SLGenerateOfflineInstallationId(h.DangerousGetHandle(), ref skuId, out p); }
            if (hr != 0) return hr;
            offlineId = Marshal.PtrToStringUni(p);
            return hr;
        }
        finally
        {
            if (p != IntPtr.Zero)
                Marshal.FreeHGlobal(p);
        }
    }

    // Subscription status (clipc.dll)
    [StructLayout(LayoutKind.Sequential)]
    public struct RawSubStatus
    {
        public uint dwEnabled;
        public uint dwSku;
        public uint dwState;
        public uint dwLicenseExpiration;
        public uint dwSubscriptionType;
    }

    [DllImport("clipc.dll", CharSet = CharSet.Unicode)]
    private static extern int ClipGetSubscriptionStatusNative(ref IntPtr pStatus);

    public static int ClipGetSubscriptionStatus(out SubStatus status)
    {
        status = default;
        IntPtr buf = IntPtr.Zero;
        try
        {
            buf = Marshal.AllocHGlobal(Marshal.SizeOf<SubStatus>());
            IntPtr p = buf;
            int hr = ClipGetSubscriptionStatusNative(ref p);
            if (hr != 0 || p == IntPtr.Zero) return hr != 0 ? hr : -1;
            status = Marshal.PtrToStructure<SubStatus>(p);
            return 0;
        }
        finally
        {
            if (buf != IntPtr.Zero) Marshal.FreeHGlobal(buf);
        }
    }

    // ---------------- Helpers ----------------
    public enum SppValueKind { Unknown = 0, String = 1, UInt32 = 4, UInt64 = 8 }
    public readonly record struct SppValue(SppValueKind Kind, string? S, uint? U32, ulong? U64)
    {
        public static SppValue FromString(string s) => new(SppValueKind.String, s, null, null);
        public static SppValue FromUInt32(uint v) => new(SppValueKind.UInt32, null, v, null);
        public static SppValue FromUInt64(ulong v) => new(SppValueKind.UInt64, null, null, v);
        public static readonly SppValue Empty = new(SppValueKind.Unknown, null, null, null);
    }

    public static SppValue InterpretValue(uint tData, uint cData, IntPtr bData)
    {
        if (tData == 1 && bData != IntPtr.Zero)
        {
            string s = cData > 0 ? Marshal.PtrToStringUni(bData, (int)(cData / 2))?.TrimEnd('\0') ?? string.Empty : Marshal.PtrToStringUni(bData) ?? string.Empty;
            return SppValue.FromString(s);
        }
        if (tData == 4) return SppValue.FromUInt32((uint)Marshal.ReadInt32(bData));
        if (tData == 8) return SppValue.FromUInt64((ulong)Marshal.ReadInt64(bData));
        return SppValue.Empty;
    }

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
            if (status == 3) status = 5;
            if (status == 2)
            {
                if (reason == 0x4004F00D) status = 3;
                else if (reason == 0x4004F065) status = 4;
                else if (reason == 0x4004FC06) status = 6;
            }
            arr[i] = new SppLicenseStatus(status, grace, reason, valid);
        }
        return arr;
    }

    private static void Ensure(SppSafeHandle h)
    {
        if (h is null || h.IsInvalid)
            throw new ArgumentException("Invalid SPP handle.", nameof(h));
    }

    // ---------------- Native imports ----------------
    internal static class NativeSppc
    {
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLOpen(ref IntPtr hSLC);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLClose(IntPtr hSLC);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetSLIDList(IntPtr hSLC, uint dwFlags, ref Guid appId, uint dwCount, out uint pcSLIDs, out IntPtr ppSLIDs);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pwszProductVersion, out uint pcStatus, out IntPtr ppStatus);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetServiceInformation(IntPtr hSLC, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid skuId, out IntPtr pwszOfflineId);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetServiceInformation(IntPtr hSLC, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("sppc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, uint tData, uint cData, IntPtr bData);
    }

    internal static class NativeSlc
    {
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLOpen(ref IntPtr hSLC);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLClose(IntPtr hSLC);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetSLIDList(IntPtr hSLC, uint dwFlags, ref Guid appId, uint dwCount, out uint pcSLIDs, out IntPtr ppSLIDs);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pwszProductVersion, out uint pcStatus, out IntPtr ppStatus);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetServiceInformation(IntPtr hSLC, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGenerateOfflineInstallationId(IntPtr hSLC, ref Guid skuId, out IntPtr pwszOfflineId);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetWindowsInformation(string valueName, out uint tData, out uint cData, out IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLGetWindowsInformationDWORD(string valueName, out uint dword);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetLicensingStatusInformation(IntPtr hSLC, ref Guid appId, ref Guid skuId, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetApplicationInformation(IntPtr hSLC, ref Guid appId, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetServiceInformation(IntPtr hSLC, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetPKeyInformation(IntPtr hSLC, ref Guid pkeyId, string valueName, uint tData, uint cData, IntPtr bData);
        [DllImport("slc.dll", CharSet = CharSet.Unicode)] internal static extern int SLSetProductSkuInformation(IntPtr hSLC, ref Guid skuId, string valueName, uint tData, uint cData, IntPtr bData);
    }
}
