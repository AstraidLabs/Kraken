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
                var sku = obj["ID"]?.ToString() ?? string.Empty;
                var info = new WindowsLicenseInfo
                {
                    ProductName = obj["Name"]?.ToString() ?? string.Empty,
                    Description = obj["Description"]?.ToString() ?? string.Empty,
                    SkuId = sku,
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

                    if (Guid.TryParse(sku, out var g))
                        info.OfflineInstallationId = s.GenerateOfflineInstallationId(g) ?? string.Empty;

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
