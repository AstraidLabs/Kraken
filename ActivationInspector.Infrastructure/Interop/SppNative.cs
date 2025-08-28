using System;
using System.Runtime.InteropServices;

namespace ActivationInspector.Infrastructure.Interop;

internal static class SppNative
{
    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLOpen(out IntPtr hSLC);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLClose(IntPtr hSLC);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGetSLIDList(
        IntPtr hSLC, uint eQuery, ref Guid appId, uint flags,
        ref uint pcReturnIds, ref IntPtr pReturnIds);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGetLicensingStatusInformation(
        IntPtr hSLC, ref Guid appId, ref Guid skuId, IntPtr pInfoQuery,
        ref uint pcStatus, ref IntPtr ppStatus);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGetPKeyInformation(
        IntPtr hSLC, ref Guid pkeyId, string value,
        ref uint pType, ref uint pcbValue, ref IntPtr ppbValue);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGetProductSkuInformation(
        IntPtr hSLC, ref Guid skuId, string value,
        ref uint pType, ref uint pcbValue, ref IntPtr ppbValue);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGetServiceInformation(
        IntPtr hSLC, string value, ref uint pType,
        ref uint pcbValue, ref IntPtr ppbValue);

    [DllImport("slc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGetWindowsInformation(
        string value, ref uint pType, ref uint pcbValue, ref IntPtr ppbValue);

    [DllImport("slc.dll", SetLastError = true)]
    internal static extern int SLGetWindowsInformationDWORD(string value, ref uint valueOut);

    [DllImport("slc.dll", SetLastError = true)]
    internal static extern int SLIsWindowsGenuineLocal(ref uint dwGenuine);

    [DllImport("sppc.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int SLGenerateOfflineInstallationId(
        IntPtr hSLC, ref Guid skuId, ref IntPtr pwszInstallationId);
}
