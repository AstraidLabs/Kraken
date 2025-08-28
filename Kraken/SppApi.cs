using System;
using System.Runtime.InteropServices;

namespace Kraken
{
    /// <summary>
    /// P/Invoke wrapper around slc.dll / sppc.dll (Software Protection Platform APIs).
    /// All signatures are taken from the PowerShell script’s DefinePInvokeMethod calls.
    /// </summary>
    public static class SppApi
    {
        /* ------------------------------------------------------------------ */
        /* 1) Basic SL (Software Licensing) handles                           */
        /* ------------------------------------------------------------------ */

        // SLOpen / SLClose – open/close the licensing context.
        // The script uses "sppc.dll" (or "slc.dll") for these.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLOpen(ref IntPtr hSLC);

        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLClose(IntPtr hSLC);

        /* ------------------------------------------------------------------ */
        /* 2) License / SKU related functions                                 */
        /* ------------------------------------------------------------------ */

        // SLGenerateOfflineInstallationId – generates an offline ID string.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGenerateOfflineInstallationId(
            IntPtr hSLC,
            ref Guid skuId,
            out IntPtr pBstrOfflineId);   // returned string pointer

        // SLGetSLIDList – list all SLIDs (subscription IDs) for an app.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetSLIDList(
            IntPtr hSLC,
            uint dwFlags,
            ref Guid appId,
            uint dwCount,                // usually 1
            out uint pcSLIDs,
            out IntPtr ppSLIDs);         // array of 16‑byte GUIDs

        // SLGetLicensingStatusInformation – get licensing state of an SKU.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetLicensingStatusInformation(
            IntPtr hSLC,
            ref Guid appId,
            ref Guid skuId,
            IntPtr pwszProductVersion,   // null in the script
            out uint pcStatus,
            out IntPtr ppStatus);        // array of status structs

        // SLGetPKeyInformation – query a public key attribute.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetPKeyInformation(
            IntPtr hSLC,
            ref Guid skuId,
            string pszKey,
            out uint ptData,
            out uint pcData,
            out IntPtr ppData);          // returned data pointer

        // SLGetProductSkuInformation – query a SKU attribute.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetProductSkuInformation(
            IntPtr hSLC,
            ref Guid skuId,
            string pszKey,
            out uint ptData,
            out uint pcData,
            out IntPtr ppData);          // returned data pointer

        // SLGetServiceInformation – query a service attribute.
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetServiceInformation(
            IntPtr hSLC,
            string pszService,
            out uint ptData,
            out uint pcData,
            out IntPtr ppData);          // returned data pointer

        // SLGetApplicationInformation – query an app attribute (Office only).
        [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetApplicationInformation(
            IntPtr hSLC,
            ref Guid appId,
            string pszApp,
            out uint ptData,
            out uint pcData,
            out IntPtr ppData);          // returned data pointer

        /* ------------------------------------------------------------------ */
        /* 3) Windows‑specific helpers (slc.dll)                              */
        /* ------------------------------------------------------------------ */

        // SLGetWindowsInformation – generic key/value query (string).
        [DllImport("slc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetWindowsInformation(
            string pszKey,
            ref uint pcchVal,            // input: buffer size, output: size needed
            out uint pdwVal,             // used for DWORD‑style values
            out IntPtr ppwszVal);        // returned Unicode string pointer

        // SLGetWindowsInformationDWORD – shortcut for DWORD values.
        [DllImport("slc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLGetWindowsInformationDWORD(
            string pszKey,
            out uint pdwVal);

        // SLIsGenuineLocal – query whether a product is genuine (Win‑7 or earlier).
        [DllImport("slwga.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int SLIsGenuineLocal(
            ref Guid appId,
            out uint pdwGenuine,
            IntPtr pwszVal);

        // SLIsWindowsGenuineLocal – query Windows genuineness (Win‑8+).
        [DllImport("slc.dll", CharSet = CharSet.Unicode, SetLastError = true,
                   CallingConvention = CallingConvention.Winapi)]
        public static extern int SLIsWindowsGenuineLocal(
            out uint pdwGenuine);

        /* ------------------------------------------------------------------ */
        /* 4) Subscription helpers (Clipc.dll)                               */
        /* ------------------------------------------------------------------ */

        // ClipGetSubscriptionStatus – returns a pointer to SubStatus struct.
        [DllImport("Clipc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int ClipGetSubscriptionStatus(
            ref IntPtr ppSubStatus);

        /* ------------------------------------------------------------------ */
        /* 5) Helper struct for ClipGetSubscriptionStatus                   */
        /* ------------------------------------------------------------------ */

        [StructLayout(LayoutKind.Sequential)]
        public struct SubStatus
        {
            public uint dwEnabled;   // 0 = disabled, 1 = enabled
            public uint dwSku;       // SKU ID
            public uint dwState;     // subscription state
        }

        /* ------------------------------------------------------------------ */
        /* 6) Optional – P/Invoke for SLGenerateOfflineInstallationId        */
        /* ------------------------------------------------------------------ */

        // The function is already defined above; we expose a small wrapper.
        public static string GenerateOfflineInstallationId(IntPtr hSLC, Guid skuId)
        {
            if (SLGenerateOfflineInstallationId(hSLC, ref skuId, out IntPtr ptr) != 0 || ptr == IntPtr.Zero)
                return null;

            string result = Marshal.PtrToStringUni(ptr);
            Marshal.FreeHGlobal(ptr);
            return result;
        }

        /* ------------------------------------------------------------------ */
        /* 7) Optional – helper to read a UTF‑16 string from SLGetWindowsInformation */
        /* ------------------------------------------------------------------ */

        public static string GetWindowsString(string key)
        {
            uint size = 0;
            uint dummy;
            if (SLGetWindowsInformation(key, ref size, out dummy, out IntPtr ptr) != 0 || ptr == IntPtr.Zero)
                return null;

            string result = Marshal.PtrToStringUni(ptr);
            Marshal.FreeHGlobal(ptr);
            return result;
        }

        /* ------------------------------------------------------------------ */
        /* 8) Optional – helper to read a DWORD from SLGetWindowsInformationDWORD */
        /* ------------------------------------------------------------------ */

        public static uint? GetWindowsDWord(string key)
        {
            if (SLGetWindowsInformationDWORD(key, out uint val) != 0)
                return null;
            return val;
        }
    }
}
