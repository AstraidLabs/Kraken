// ==========================================================================
// LicenceService.cs
// ==========================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;

namespace Kraken;

/// <summary>
/// Holds a *complete* snapshot of all licence‑related information that
/// can be retrieved from the machine (Windows & Office + SPP + WMI + registry).
/// </summary>
public sealed class LicenseSummary
{
    public WindowsLicenseInfo? WindowsLicense { get; set; }
    public List<OfficeLicenseInfo> OfficeLicenses { get; set; } = new();
    public KmsServerInfo? KmsServer { get; set; }
    public KmsClientInfo? KmsClient { get; set; }
    public SubscriptionInfo? Subscription { get; set; }
    public VNextInfo? VNext { get; set; }
    public AdActivationInfo? AdActivation { get; set; }
    public VmActivationInfo? VmActivation { get; set; }
    public bool IsDigitalLicense { get; set; } = false;
    public bool SharedComputerLicensingEnabled { get; set; } = false;
}

/// <summary>
/// Windows licence – gathered from WMI + SPP.
/// </summary>
public sealed class WindowsLicenseInfo
{
    public string ProductName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string PartialProductKey { get; set; } = string.Empty;
    public LicenseStatus Status { get; set; }
    public int GraceMinutes { get; set; }
    public DateTime? Expiration { get; set; }
    public string Channel { get; set; } = string.Empty;
    public DateTime? EvaluationEndDate { get; set; }

    // SPP‑based data
    public DateTime? LastActivationTime { get; set; }
    public int? LastActivationHResult { get; set; }
    public DateTime? KernelTimeBomb { get; set; }
    public DateTime? SystemTimeBomb { get; set; }
    public DateTime? TrustedTime { get; set; }
    public int? RearmCount { get; set; }
    public bool? IsWindowsGenuineLocal { get; set; }

    // License key files
    public List<string> LicenseKeyFiles { get; } = new();
}

/// <summary>
/// Office licence – gathered from WMI + SPP.
/// </summary>
public sealed class OfficeLicenseInfo
{
    public string SLID { get; set; } = string.Empty;
    public string ProductName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string PartialProductKey { get; set; } = string.Empty;
    public LicenseStatus Status { get; set; }
    public int GraceMinutes { get; set; }
    public DateTime? Expiration { get; set; }
    public string Channel { get; set; } = string.Empty;
    public string SkuId { get; set; } = string.Empty;

    // SPP‑based data
    public string OfflineInstallationId { get; set; } = string.Empty;
    public int? RearmCount { get; set; }
    public DateTime? TrustedTime { get; set; }

    // License key files
    public List<string> LicenseKeyFiles { get; } = new();
}

/// <summary>
/// Information about a local KMS server (if the machine is a server).  
/// Data is pulled from the `SoftwareLicensingService` WMI class.
/// </summary>
public sealed class KmsServerInfo
{
    public int CurrentClients { get; set; }
    public int TotalRequests { get; set; }
    public int FailedRequests { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Port { get; set; }
    public string IPAddress { get; set; } = string.Empty;
    public string ServiceState { get; set; } = string.Empty;
    public string ServiceStatus { get; set; } = string.Empty;
}

/// <summary>
/// Information about a local KMS client (if the machine is a client).  
/// Data is pulled from the `SoftwareLicensingService` WMI class.
/// </summary>
public sealed class KmsClientInfo
{
    public string ClientPid { get; set; } = string.Empty;
    public string ServerName { get; set; } = string.Empty;
    public int ServerPort { get; set; }
    public string DiscoveredName { get; set; } = string.Empty;
    public int DiscoveredPort { get; set; }
    public string DiscoveredIP { get; set; } = string.Empty;
    public int RenewalInterval { get; set; }
    public string ActivationId { get; set; } = string.Empty;
    public DateTime? LastActivationTime { get; set; }
    public int RearmCount { get; set; }
}

/// <summary>
/// Active Directory‑based activation data (if present).  
/// </summary>
public sealed class AdActivationInfo
{
    public string ObjectName { get; set; } = string.Empty;
    public string ObjectDN { get; set; } = string.Empty;
    public string CsvlkPid { get; set; } = string.Empty;
    public string CsvlkSkuId { get; set; } = string.Empty;
    public string ActivationId { get; set; } = string.Empty;
}

/// <summary>
/// Virtual‑machine activation data (if present).  
/// </summary>
public sealed class VmActivationInfo
{
    public string InheritedActivationId { get; set; } = string.Empty;
    public string HostMachineName { get; set; } = string.Empty;
    public string HostDigitalPid2 { get; set; } = string.Empty;
    public DateTime? ActivationTime { get; set; }
    public string HostMachineID { get; set; } = string.Empty;
    public string HostMachineVersion { get; set; } = string.Empty;
}

/// <summary>
/// vNext/Click‑To‑Run diagnostic information.
/// </summary>
public sealed class VNextInfo
{
    public bool SharedComputerLicensing { get; set; }
    public Dictionary<string, string> ModePerProductReleaseId { get; } = new();
    public List<string> Licenses { get; } = new();
    public List<string> LicenseKeyFiles { get; } = new();
}

/// <summary>
/// Subscription status returned by clipc.dll.
/// </summary>
public sealed class SubscriptionInfo
{
    public bool Enabled { get; set; }
    public uint Sku { get; set; }
    public uint State { get; set; }
    public DateTime? LicenseExpiration { get; set; }
    public uint SubscriptionType { get; set; }
}

/// <summary>
/// Licence status for WMI licences.
/// </summary>
public enum LicenseStatus
{
    Unknown = -1,
    Unlicensed,
    Licensed,
    Grace,
    Notification,
    Expired,
    ExtendedGrace
}

/// <summary>
/// Helper class that aggregates all licence information in one place.
/// </summary>
public static class LicenseService
{
    /// <summary>
    /// Retrieves a summary of all licence information available on the machine.
    /// </summary>
    public static LicenseSummary GetLicenseSummary()
    {
        var summary = new LicenseSummary();

        // 1. Windows + Office licences
        if (SppManagement.TryOpenSession(out var session))
        {
            using (session)
            {
                summary.WindowsLicense = GetWindowsLicense(session);
                summary.OfficeLicenses = GetOfficeLicenses(session);
            }
        }
        else
        {
            summary.WindowsLicense = GetWindowsLicense(null);
            summary.OfficeLicenses = GetOfficeLicenses(null);
        }

        // 2. KMS server/client (if present)
        summary.KmsServer = GetKmsServerInfo();
        summary.KmsClient = GetKmsClientInfo();

        // 3. Subscription (Office 365, etc.)
        summary.Subscription = GetSubscriptionInfo();

        // 4. vNext / Click‑To‑Run diagnostics
        summary.VNext = GetVNextInfo();

        // 5. AD / VM activations
        summary.AdActivation = GetAdActivationInfo();
        summary.VmActivation = GetVmActivationInfo();

        // 6. Digital licence & shared‑computer licensing
        summary.IsDigitalLicense = IsDigitalLicense();
        summary.SharedComputerLicensingEnabled = IsSharedComputerLicensingEnabled();

        return summary;
    }

    /// <summary>
    /// Collects basic Windows licence information using WMI.
    /// If an SPP session is provided, additional SPP data is retrieved.
    /// </summary>
    public static WindowsLicenseInfo? GetWindowsLicense(SppManagement? s)
    {
        try
        {
            const string wql = @"
                SELECT Name, Description, ID, PartialProductKey, LicenseStatus,
                       GracePeriodRemaining, EvaluationEndDate, ProductKeyChannel
                FROM SoftwareLicensingProduct
                WHERE PartialProductKey IS NOT NULL";

            using var searcher = new ManagementObjectSearcher(wql);
            foreach (ManagementObject obj in searcher.Get())
            {
                var info = new WindowsLicenseInfo
                {
                    ProductName = obj["Name"]?.ToString() ?? string.Empty,
                    Description = obj["Description"]?.ToString() ?? string.Empty,
                    PartialProductKey = obj["PartialProductKey"]?.ToString() ?? string.Empty,
                    Status = ParseLicenseStatus(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null
                                      ? Convert.ToInt32(obj["GracePeriodRemaining"])
                                      : 0,
                    Expiration = ParseDate(obj["EvaluationEndDate"]),
                    Channel = obj["ProductKeyChannel"]?.ToString() ?? string.Empty,
                    EvaluationEndDate = ParseDate(obj["EvaluationEndDate"])
                };

                // ----- SPP‑based data -----------------------------------------
                if (s != null)
                {
                    info.LastActivationTime = ParseDate(s.GetWindowsString("LastActivationTime"));
                    var hres = s.GetWindowsDWord("LastActivationHR");
                    if (hres.HasValue) info.LastActivationHResult = (int)hres.Value;

                    info.KernelTimeBomb = ParseDate(s.GetWindowsString("KernelDebuggerTimeBomb"));
                    info.SystemTimeBomb = ParseDate(s.GetWindowsString("TimeBomb"));
                    info.TrustedTime = ParseDate(s.GetWindowsString("TrustedTime"));

                    var rearm = s.GetWindowsDWord("RemainingWindowsRearmCount");
                    if (rearm.HasValue) info.RearmCount = (int)rearm.Value;

                    if (SppApi.SLIsWindowsGenuineLocal(out uint genuine) == 0)
                        info.IsWindowsGenuineLocal = genuine != 0;
                }

                // ----- License key files --------------------------------------
                var licDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Office", "Licenses");

                if (Directory.Exists(licDir))
                {
                    foreach (var file in Directory.GetFiles(licDir, "*.json",
                                 SearchOption.AllDirectories))
                    {
                        info.LicenseKeyFiles.Add(Path.GetFileName(file));
                        try
                        {
                            using var doc = JsonDocument.Parse(File.ReadAllText(file));
                            if (doc.RootElement.TryGetProperty("PartialProductKey", out var pk))
                            {
                                var val = pk.GetString();
                                if (!string.IsNullOrEmpty(val))
                                    info.PartialProductKey = val;
                            }
                        }
                        catch
                        {
                            // ignore malformed JSON
                        }
                    }
                }

                return info;
            }
        }
        catch
        {
            // swallow everything – method is tolerant
        }
        return null;
    }

    /// <summary>
    /// Retrieves Office licence information using WMI.
    /// If an SPP session is provided, installation IDs are generated for each product.
    /// </summary>
    public static List<OfficeLicenseInfo> GetOfficeLicenses(SppManagement? s)
    {
        var list = new List<OfficeLicenseInfo>();
        try
        {
            const string query = @"
                SELECT Name, Description, ID, PartialProductKey, LicenseStatus,
                       GracePeriodRemaining, EvaluationEndDate, ProductKeyChannel
                FROM OfficeSoftwareProtectionProduct
                WHERE PartialProductKey IS NOT NULL";

            using var searcher = new ManagementObjectSearcher(query);
            foreach (ManagementObject obj in searcher.Get())
            {
                var slid = obj["ID"]?.ToString() ?? string.Empty;
                var lic = new OfficeLicenseInfo
                {
                    SLID = slid,
                    ProductName = obj["Name"]?.ToString() ?? string.Empty,
                    Description = obj["Description"]?.ToString() ?? string.Empty,
                    PartialProductKey = obj["PartialProductKey"]?.ToString() ?? string.Empty,
                    Status = ParseLicenseStatus(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null
                                      ? Convert.ToInt32(obj["GracePeriodRemaining"])
                                      : 0,
                    Expiration = ParseDate(obj["EvaluationEndDate"]),
                    Channel = obj["ProductKeyChannel"]?.ToString() ?? string.Empty,
                    SkuId = obj["ID"]?.ToString() ?? string.Empty
                };

                if (s != null && Guid.TryParse(slid, out var guid))
                {
                    lic.OfflineInstallationId = s.GenerateOfflineInstallationId(guid) ?? string.Empty;

                    var rearm = s.GetWindowsDWord("RemainingAppRearmCount");
                    if (rearm.HasValue) lic.RearmCount = (int)rearm.Value;

                    lic.TrustedTime = ParseDate(s.GetWindowsString("TrustedTime"));
                }

                // License key files
                string licDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Office", "Licenses", slid);

                if (Directory.Exists(licDir))
                {
                    foreach (var file in Directory.GetFiles(licDir, "*.json",
                                 SearchOption.AllDirectories))
                    {
                        lic.LicenseKeyFiles.Add(Path.GetFileName(file));
                    }
                }

                list.Add(lic);
            }
        }
        catch
        {
            // tolerant – return what we have
        }
        return list;
    }

    /// <summary>
    /// Retrieves information about the local KMS server (if the machine is a server).
    /// </summary>
    public static KmsServerInfo? GetKmsServerInfo()
    {
        try
        {
            const string wql = @"
                SELECT KeyManagementServiceCurrentCount, KeyManagementServiceTotalRequests,
                       KeyManagementServiceFailedRequests, KeyManagementServiceMachineName,
                       KeyManagementServicePort, KeyManagementServiceMachineIpAddress,
                       ServiceState, ServiceStatus
                FROM SoftwareLicensingService";

            using var searcher = new ManagementObjectSearcher(wql);
            foreach (ManagementObject obj in searcher.Get())
            {
                return new KmsServerInfo
                {
                    CurrentClients = Convert.ToInt32(obj["KeyManagementServiceCurrentCount"] ?? 0),
                    TotalRequests = Convert.ToInt32(obj["KeyManagementServiceTotalRequests"] ?? 0),
                    FailedRequests = Convert.ToInt32(obj["KeyManagementServiceFailedRequests"] ?? 0),
                    Name = obj["KeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    Port = obj["KeyManagementServicePort"] != null
                                      ? Convert.ToInt32(obj["KeyManagementServicePort"])
                                      : 0,
                    IPAddress = obj["KeyManagementServiceMachineIpAddress"]?.ToString() ?? string.Empty,
                    ServiceState = obj["ServiceState"]?.ToString() ?? string.Empty,
                    ServiceStatus = obj["ServiceStatus"]?.ToString() ?? string.Empty
                };
            }
        }
        catch
        {
            // ignore
        }
        return null;
    }

    /// <summary>
    /// Retrieves information about the KMS client (if the machine is a client).
    /// </summary>
    public static KmsClientInfo? GetKmsClientInfo()
    {
        try
        {
            const string wql = @"
                SELECT ClientMachineID, KeyManagementServiceMachineName, KeyManagementServicePort,
                       DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort,
                       DiscoveredKeyManagementServiceMachineIpAddress, VLActivationInterval,
                       KeyManagementServiceActivationID, ProductActivationLastAttemptTime,
                       RemainingWindowsRearmCount
                FROM SoftwareLicensingService";

            using var searcher = new ManagementObjectSearcher(wql);
            foreach (ManagementObject obj in searcher.Get())
            {
                return new KmsClientInfo
                {
                    ClientPid = obj["ClientMachineID"]?.ToString() ?? string.Empty,
                    ServerName = obj["KeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    ServerPort = obj["KeyManagementServicePort"] != null
                                          ? Convert.ToInt32(obj["KeyManagementServicePort"])
                                          : 0,
                    DiscoveredName = obj["DiscoveredKeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    DiscoveredPort = obj["DiscoveredKeyManagementServiceMachinePort"] != null
                                          ? Convert.ToInt32(obj["DiscoveredKeyManagementServiceMachinePort"])
                                          : 0,
                    DiscoveredIP = obj["DiscoveredKeyManagementServiceMachineIpAddress"]?.ToString() ?? string.Empty,
                    RenewalInterval = obj["VLActivationInterval"] != null
                                          ? Convert.ToInt32(obj["VLActivationInterval"])
                                          : 0,
                    ActivationId = obj["KeyManagementServiceActivationID"]?.ToString() ?? string.Empty,
                    LastActivationTime = ParseDate(obj["ProductActivationLastAttemptTime"]),
                    RearmCount = obj["RemainingWindowsRearmCount"] != null
                                          ? Convert.ToInt32(obj["RemainingWindowsRearmCount"])
                                          : 0
                };
            }
        }
        catch
        {
            // ignore
        }
        return null;
    }

    /// <summary>
    /// Checks whether the system has a digital licence (EditionUpgradeManager).
    /// </summary>
    public static bool IsDigitalLicense()
    {
        try
        {
            var type = Type.GetTypeFromProgID("EditionUpgradeManagerObj.EditionUpgradeManager");
            if (type == null) return false;

            dynamic obj = Activator.CreateInstance(type);
            obj?.AcquireModernLicenseForWindows();      // method exists, returns HRESULT
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Gets subscription information, if available (Office 365, etc.).
    /// </summary>
    public static SubscriptionInfo? GetSubscriptionInfo()
    {
        try
        {
            if (SppApi.TryGetSubscriptionStatus(out var status))
            {
                return new SubscriptionInfo
                {
                    Enabled = status.dwEnabled != 0,
                    Sku = status.dwSku,
                    State = status.dwState,
                    LicenseExpiration = status.dwLicenseExpiration != 0
                                         ? DateTimeOffset.FromUnixTimeSeconds(status.dwLicenseExpiration).DateTime
                                         : null,
                    SubscriptionType = status.dwSubscriptionType
                };
            }
        }
        catch
        {
            // ignore
        }
        return null;
    }

    /// <summary>
    /// Retrieves vNext diagnostic information from the registry.
    /// </summary>
    public static VNextInfo? GetVNextInfo()
    {
        try
        {
            var info = new VNextInfo();

            using var cfgKey = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (cfgKey != null)
            {
                info.SharedComputerLicensing = Convert.ToInt32(cfgKey.GetValue("SharedComputerLicensing", 0)) != 0;
            }

            using var modeKey = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Office\ClickToRun\Scenario");
            if (modeKey != null)
            {
                foreach (var name in modeKey.GetValueNames())
                {
                    info.ModePerProductReleaseId[name] = modeKey.GetValue(name)?.ToString() ?? string.Empty;
                }
            }

            using var licKey = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Office\Licenses");
            if (licKey != null)
            {
                foreach (var sub in licKey.GetSubKeyNames())
                {
                    info.Licenses.Add(sub);
                }
            }

            string licDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft", "Office", "Licenses");
            if (Directory.Exists(licDir))
            {
                foreach (var file in Directory.GetFiles(licDir, "*.json",
                                 SearchOption.AllDirectories))
                {
                    info.LicenseKeyFiles.Add(Path.GetFileName(file));
                }
            }

            return info;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Retrieves Active Directory activation information.
    /// </summary>
    public static AdActivationInfo? GetAdActivationInfo()
    {
        try
        {
            const string wql = @"
                SELECT ADActivationObjectName, ADActivationObjectDN,
                       KeyManagementServiceProductKeyID, KeyManagementServiceSkuID,
                       KeyManagementServiceActivationID
                FROM SoftwareLicensingService";

            using var searcher = new ManagementObjectSearcher(wql);
            foreach (ManagementObject obj in searcher.Get())
            {
                return new AdActivationInfo
                {
                    ObjectName = obj["ADActivationObjectName"]?.ToString() ?? string.Empty,
                    ObjectDN = obj["ADActivationObjectDN"]?.ToString() ?? string.Empty,
                    CsvlkPid = obj["KeyManagementServiceProductKeyID"]?.ToString() ?? string.Empty,
                    CsvlkSkuId = obj["KeyManagementServiceSkuID"]?.ToString() ?? string.Empty,
                    ActivationId = obj["KeyManagementServiceActivationID"]?.ToString() ?? string.Empty
                };
            }
        }
        catch
        {
            // ignore
        }
        return null;
    }

    /// <summary>
    /// Retrieves VM activation information.
    /// </summary>
    public static VmActivationInfo? GetVmActivationInfo()
    {
        try
        {
            const string wql = @"
                SELECT InheritedActivationID, VirtualizationHostMachineName,
                       VirtualizationHostDigitalPid2, VirtualizationActivationTime,
                       VirtualizationHostMachineID, VirtualizationHostMachineVersion
                FROM SoftwareLicensingService";

            using var searcher = new ManagementObjectSearcher(wql);
            foreach (ManagementObject obj in searcher.Get())
            {
                return new VmActivationInfo
                {
                    InheritedActivationId = obj["InheritedActivationID"]?.ToString() ?? string.Empty,
                    HostMachineName = obj["VirtualizationHostMachineName"]?.ToString() ?? string.Empty,
                    HostDigitalPid2 = obj["VirtualizationHostDigitalPid2"]?.ToString() ?? string.Empty,
                    ActivationTime = ParseDate(obj["VirtualizationActivationTime"]),
                    HostMachineID = obj["VirtualizationHostMachineID"]?.ToString() ?? string.Empty,
                    HostMachineVersion = obj["VirtualizationHostMachineVersion"]?.ToString() ?? string.Empty
                };
            }
        }
        catch
        {
            // ignore
        }
        return null;
    }

    /// <summary>
    /// Determines whether shared computer licensing is enabled for Office.
    /// </summary>
    public static bool IsSharedComputerLicensingEnabled()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (key != null)
            {
                return Convert.ToInt32(key.GetValue("SharedComputerLicensing", 0)) != 0;
            }
        }
        catch
        {
            // ignore
        }
        return false;
    }

    // --------------------------------------------------------------------
    // Helper methods (parsing)
    // --------------------------------------------------------------------
    private static LicenseStatus ParseLicenseStatus(object? statusObj)
    {
        if (statusObj == null) return LicenseStatus.Unknown;
        try
        {
            var status = Convert.ToInt32(statusObj);
            return status switch
            {
                0 => LicenseStatus.Unlicensed,
                1 => LicenseStatus.Licensed,
                2 => LicenseStatus.Grace,
                3 => LicenseStatus.Notification,
                4 => LicenseStatus.Expired,
                5 => LicenseStatus.ExtendedGrace,
                _ => LicenseStatus.Unknown
            };
        }
        catch
        {
            return LicenseStatus.Unknown;
        }
    }

    private static DateTime? ParseDate(object? obj)
    {
        if (obj == null) return null;
        try
        {
            return ManagementDateTimeConverter.ToDateTime(obj.ToString());
        }
        catch
        {
            return null;
        }
    }
}
