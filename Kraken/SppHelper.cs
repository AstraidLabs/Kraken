using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text.Json;
using Microsoft.Win32;

namespace Kraken;

/// <summary>Activation details for a specific SKU.</summary>
public record SppActivationInfo(
    string? CustomerPID,
    string? KeyManagementServiceName,
    string? VLActivationType,
    string? VLActivationInterval);

/// <summary>Represents a vNext/Click-To-Run license file.</summary>
public record VNextLicense(
    string FileName,
    string ProductReleaseId,
    string Status,
    DateTime? Expiry);

/// <summary>Subscription status information returned by clipc.dll.</summary>
[StructLayout(LayoutKind.Sequential)]
public readonly record struct SubStatus(
    int LicenseStatus,
    int LicenseState,
    int GenuineStatus,
    int GenuineState);

/// <summary>
/// High level helpers for querying Software Protection Platform data.
/// </summary>
public static class SppHelper
{
    private static readonly Dictionary<string, object?> _collected = new();

    private static readonly string[] _licensingStates =
    {
        "Unlicensed",
        "Licensed",
        "Out-of-Box Grace",
        "Out-of-Tolerance Grace",
        "Non-Genuine Grace",
        "Notification",
        "Extended Grace"
    };

    private static readonly string[] _genuineStates =
    {
        "Unknown",
        "Genuine",
        "Non-Genuine",
        "Unknown"
    };

    [DllImport("clipc.dll", CharSet = CharSet.Unicode)]
    private static extern int ClipGetSubscriptionStatus(ref SubStatus status);

    [ComImport]
    [Guid("00000000-0000-0000-0000-000000000000")] // Replace with actual GUID
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IEditionUpgradeManager
    {
        [PreserveSig]
        int AcquireModernLicenseForWindows();
    }

    /// <summary>
    /// Retrieves activation information for the specified SKU.
    /// </summary>
    /// <param name="skuId">SKU identifier.</param>
    /// <returns>Activation information objects.</returns>
    /// <exception cref="InvalidOperationException">SPP session could not be opened.</exception>
    public static IEnumerable<SppActivationInfo> GetSkuActivationInfo(Guid skuId)
    {
        if (!SppApi.TryOpenSession(out var handle))
            throw new InvalidOperationException("Cannot open SPP session.");
        using (handle)
        {
            var info = new SppActivationInfo(
                GetSkuString(handle, skuId, "CustomerPID"),
                GetSkuString(handle, skuId, "KeyManagementServiceName"),
                GetSkuString(handle, skuId, "VLActivationType"),
                GetSkuString(handle, skuId, "VLActivationInterval"));
            _collected[$"SkuActivationInfo:{skuId}"] = info;
            return new[] { info };
        }

        static string? GetSkuString(SppApi.SppSafeHandle h, Guid sku, string name)
        {
            int hr = SppApi.SLGetProductSkuInformation(h, sku, name, out uint t, out uint c, out IntPtr b);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            try
            {
                return SppApi.InterpretValue(t, c, b).S;
            }
            finally
            {
                if (b != IntPtr.Zero)
                    Marshal.FreeHGlobal(b);
            }
        }
    }

    /// <summary>
    /// Retrieves vNext licenses by reading registry entries and decoding license files.
    /// </summary>
    public static IEnumerable<VNextLicense> GetVNextLicenses()
    {
        var list = new List<VNextLicense>();
        using var office = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Office");
        if (office == null) return list;

        string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "Office", "Licenses");

        foreach (var ver in office.GetSubKeyNames())
        {
            using var key = office.OpenSubKey($"{ver}\\Common\\Licensing\\LicensingNext");
            if (key == null) continue;
            foreach (var sub in key.GetSubKeyNames())
            {
                string file = Path.Combine(baseDir, sub + ".json");
                if (!File.Exists(file)) continue;
                string encoded = File.ReadAllText(file);
                var data = Convert.FromBase64String(encoded);
                using var doc = JsonDocument.Parse(data);
                string release = doc.RootElement.GetProperty("ProductReleaseId").GetString() ?? string.Empty;
                string status = doc.RootElement.GetProperty("Status").GetString() ?? string.Empty;
                DateTime? expiry = null;
                if (doc.RootElement.TryGetProperty("Expiry", out var ex))
                {
                    if (ex.ValueKind == JsonValueKind.String && DateTime.TryParse(ex.GetString(), out var dt))
                        expiry = dt;
                    else if (ex.ValueKind == JsonValueKind.Number)
                        expiry = DateTimeOffset.FromUnixTimeSeconds(ex.GetInt64()).DateTime;
                }
                list.Add(new VNextLicense(Path.GetFileName(file), release, status, expiry));
            }
        }
        _collected["VNextLicenses"] = list;
        return list;
    }

    /// <summary>
    /// Retrieves subscription status using clipc.dll.
    /// </summary>
    /// <returns>The subscription status.</returns>
    public static SubStatus GetSubscriptionStatus()
    {
        var status = new SubStatus();
        int hr = ClipGetSubscriptionStatus(ref status);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);
        _collected["SubscriptionStatus"] = status;
        return status;
    }

    /// <summary>
    /// Calls the EditionUpgradeManager COM interface to acquire a modern license.
    /// </summary>
    /// <returns><c>true</c> when the call succeeds.</returns>
    public static bool AcquireModernLicense()
    {
        Guid clsid = new("00000000-0000-0000-0000-000000000000"); // Replace with actual CLSID
        Type t = Type.GetTypeFromCLSID(clsid, true);
        object? obj = null;
        try
        {
            obj = Activator.CreateInstance(t);
            var mgr = (IEditionUpgradeManager)obj!;
            int hr = mgr.AcquireModernLicenseForWindows();
            Marshal.ThrowExceptionForHR(hr);
            bool success = hr == 0;
            _collected["AcquireModernLicense"] = success;
            return success;
        }
        finally
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.ReleaseComObject(obj);
        }
    }

    /// <summary>Looks up a licensing state description.</summary>
    public static string? LookupLicensingState(int code) =>
        code >= 0 && code < _licensingStates.Length ? _licensingStates[code] : null;

    /// <summary>Looks up a genuine state description.</summary>
    public static string? LookupGenuineState(int code) =>
        code >= 0 && code < _genuineStates.Length ? _genuineStates[code] : null;

    /// <summary>
    /// Ensures that the Software Protection Platform service is running.
    /// </summary>
    public static void EnsureSppServiceRunning()
    {
        using var sc = new ServiceController("sppsvc");
        if (sc.Status != ServiceControllerStatus.Running)
        {
            sc.Start();
            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
        }
    }

    /// <summary>
    /// Prints the collected information in a simple table.
    /// </summary>
    public static void PrintCollectedInfo()
    {
        int width = 0;
        foreach (var key in _collected.Keys)
            width = Math.Max(width, key.Length);
        foreach (var kv in _collected)
            Console.WriteLine($"{kv.Key.PadRight(width)} : {kv.Value}");
    }
}

