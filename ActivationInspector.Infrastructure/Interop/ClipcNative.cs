using System;
using System.Runtime.InteropServices;

namespace ActivationInspector.Infrastructure.Interop;

internal static class ClipcNative
{
    [DllImport("Clipc.dll", SetLastError = true)]
    internal static extern int ClipGetSubscriptionStatus(out IntPtr pStatus);

    [StructLayout(LayoutKind.Sequential)]
    internal struct SubStatus
    {
        public uint dwEnabled;
        public uint dwSku;
        public uint dwState;
    }
}
