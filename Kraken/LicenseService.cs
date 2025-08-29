using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Text.Json;
using Microsoft.Win32;

namespace Kraken;

/// <summary>
/// High level helpers that use SppManagement, SppApi and WMI/registry to build license information objects.
/// </summary>
public static class LicenseService
{
    /// <summary>
    /// Retrieves a summary of all licence information available on the machine.
    /// </summary>
    public static LicenseSummary GetLicenseSummary()
    {
        var summary = new LicenseSummary();
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
        summary.KmsServer = GetKmsServerInfo();
        summary.KmsClient = GetKmsClientInfo();
        summary.Subscription = GetSubscriptionInfo();
        summary.VNext = GetVNextInfo();
        summary.AdActivation = GetAdActivationInfo();
        summary.VmActivation = GetVmActivationInfo();
        if (summary.WindowsLicense != null)
        {
            summary.WindowsLicense.IsDigitalLicense = IsDigitalLicense();
        }
        return summary;
    }

    /// <summary>
    /// Collects basic Windows licence information using WMI. If an SPP session is provided,
    /// additional SPP data is retrieved.
    /// </summary>
    public static WindowsLicenseInfo? GetWindowsLicense(SppManagement? s)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Name, Description, ID, PartialProductKey, LicenseStatus, GracePeriodRemaining, EvaluationEndDate, ProductKeyChannel FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL");
            foreach (var obj in searcher.Get())
            {
                var info = new WindowsLicenseInfo
                {
                    ProductName = obj["Name"]?.ToString() ?? string.Empty,
                    Description = obj["Description"]?.ToString() ?? string.Empty,
                    PartialProductKey = obj["PartialProductKey"]?.ToString() ?? string.Empty,
                    Status = ParseLicenseStatus(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null ? Convert.ToInt32(obj["GracePeriodRemaining"]) : 0,
                    Expiration = ParseDate(obj["EvaluationEndDate"]),
                    Channel = obj["ProductKeyChannel"]?.ToString() ?? string.Empty,
                    EvaluationEndDate = ParseDate(obj["EvaluationEndDate"])
                };

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

                var licDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Office", "Licenses");
                if (Directory.Exists(licDir))
                {
                    foreach (var file in Directory.GetFiles(licDir, "*.json", SearchOption.AllDirectories))
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
                        }
                    }
                }

                return info;
            }
        }
        catch
        {
            // ignore
        }
        return null;
    }

    /// <summary>
    /// Retrieves Office licence information using WMI. If an SPP session is provided,
    /// installation IDs are generated for each product.
    /// </summary>
    public static List<OfficeLicenseInfo> GetOfficeLicenses(SppManagement? s)
    {
        var list = new List<OfficeLicenseInfo>();
        try
        {
            const string query = "SELECT Name, Description, ID, PartialProductKey, LicenseStatus, GracePeriodRemaining, EvaluationEndDate, ProductKeyChannel FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey IS NOT NULL";
            using var searcher = new ManagementObjectSearcher(query);
            foreach (var obj in searcher.Get())
            {
                var slid = obj["ID"]?.ToString() ?? string.Empty;
                var lic = new OfficeLicenseInfo
                {
                    SLID = slid,
                    ProductName = obj["Name"]?.ToString() ?? string.Empty,
                    Description = obj["Description"]?.ToString() ?? string.Empty,
                    PartialProductKey = obj["PartialProductKey"]?.ToString() ?? string.Empty,
                    Status = ParseLicenseStatus(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null ? Convert.ToInt32(obj["GracePeriodRemaining"]) : 0,
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

                string licDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Office", "Licenses", slid);
                if (Directory.Exists(licDir))
                {
                    foreach (var file in Directory.GetFiles(licDir, "*.json", SearchOption.AllDirectories))
                    {
                        lic.LicenseKeyFiles.Add(Path.GetFileName(file));
                    }
                }

                list.Add(lic);
            }
        }
        catch
        {
            // ignore
        }
        return list;
    }

    /// <summary>
    /// Retrieves information about the local KMS server.
    /// </summary>
    public static KmsServerInfo? GetKmsServerInfo()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT KeyManagementServiceCurrentCount, KeyManagementServiceTotalRequests, KeyManagementServiceFailedRequests, KeyManagementServiceMachineName, KeyManagementServicePort, KeyManagementServiceMachineIpAddress, ServiceState, ServiceStatus FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
            {
                return new KmsServerInfo
                {
                    CurrentClients = Convert.ToInt32(obj["KeyManagementServiceCurrentCount"] ?? 0),
                    TotalRequests = Convert.ToInt32(obj["KeyManagementServiceTotalRequests"] ?? 0),
                    FailedRequests = Convert.ToInt32(obj["KeyManagementServiceFailedRequests"] ?? 0),
                    Name = obj["KeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    Port = obj["KeyManagementServicePort"] != null ? Convert.ToInt32(obj["KeyManagementServicePort"]) : 0,
                    IPAddress = obj["KeyManagementServiceMachineIpAddress"]?.ToString() ?? string.Empty,
                    ServiceState = obj["ServiceState"]?.ToString() ?? string.Empty,
                    ServiceStatus = obj["ServiceStatus"]?.ToString() ?? string.Empty
                };
            }
        }
        catch
        {
        }
        return null;
    }

    /// <summary>
    /// Retrieves information about the KMS client.
    /// </summary>
    public static KmsClientInfo? GetKmsClientInfo()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT ClientMachineID, KeyManagementServiceMachineName, KeyManagementServicePort, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, DiscoveredKeyManagementServiceMachineIpAddress, VLActivationInterval, KeyManagementServiceActivationID, ProductActivationLastAttemptTime, RemainingWindowsRearmCount FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
            {
                return new KmsClientInfo
                {
                    ClientPid = obj["ClientMachineID"]?.ToString() ?? string.Empty,
                    ServerName = obj["KeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    ServerPort = obj["KeyManagementServicePort"] != null ? Convert.ToInt32(obj["KeyManagementServicePort"]) : 0,
                    DiscoveredName = obj["DiscoveredKeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    DiscoveredPort = obj["DiscoveredKeyManagementServiceMachinePort"] != null ? Convert.ToInt32(obj["DiscoveredKeyManagementServiceMachinePort"]) : 0,
                    DiscoveredIP = obj["DiscoveredKeyManagementServiceMachineIpAddress"]?.ToString() ?? string.Empty,
                    RenewalInterval = obj["VLActivationInterval"] != null ? Convert.ToInt32(obj["VLActivationInterval"]) : 0,
                    ActivationId = obj["KeyManagementServiceActivationID"]?.ToString() ?? string.Empty,
                    LastActivationTime = ParseDate(obj["ProductActivationLastAttemptTime"]),
                    RearmCount = obj["RemainingWindowsRearmCount"] != null ? Convert.ToInt32(obj["RemainingWindowsRearmCount"]) : 0
                };
            }
        }
        catch
        {
        }
        return null;
    }

    /// <summary>
    /// Checks whether the system has a digital licence.
    /// </summary>
    public static bool IsDigitalLicense()
    {
        try
        {
            var type = Type.GetTypeFromProgID("EditionUpgradeManagerObj.EditionUpgradeManager");
            if (type == null)
                return false;
            dynamic obj = Activator.CreateInstance(type);
            obj?.AcquireModernLicenseForWindows();
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Gets subscription information, if available.
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
                    LicenseExpiration = status.dwLicenseExpiration != 0 ? DateTimeOffset.FromUnixTimeSeconds(status.dwLicenseExpiration).DateTime : null,
                    SubscriptionType = status.dwSubscriptionType
                };
            }
        }
        catch
        {
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
            using var cfgKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (cfgKey != null)
            {
                info.SharedComputerLicensing = Convert.ToInt32(cfgKey.GetValue("SharedComputerLicensing", 0)) != 0;
            }

            using var modeKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Scenario");
            if (modeKey != null)
            {
                foreach (var name in modeKey.GetValueNames())
                {
                    info.ModePerProductReleaseId[name] = modeKey.GetValue(name)?.ToString() ?? string.Empty;
                }
            }

            using var licKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\Licenses");
            if (licKey != null)
            {
                foreach (var sub in licKey.GetSubKeyNames())
                {
                    info.Licenses.Add(sub);
                }
            }

            string licDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft", "Office", "Licenses");
            if (Directory.Exists(licDir))
            {
                foreach (var file in Directory.GetFiles(licDir, "*.json", SearchOption.AllDirectories))
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
            using var searcher = new ManagementObjectSearcher(
                "SELECT ADActivationObjectName, ADActivationObjectDN, KeyManagementServiceProductKeyID, KeyManagementServiceSkuID, KeyManagementServiceActivationID FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
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
            using var searcher = new ManagementObjectSearcher(
                "SELECT InheritedActivationID, VirtualizationHostMachineName, VirtualizationHostDigitalPid2, VirtualizationActivationTime, VirtualizationHostMachineID, VirtualizationHostMachineVersion FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
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
        }
        return null;
    }

    /// <summary>
    /// Determines if shared computer licensing is enabled for Office.
    /// </summary>
    public static bool IsSharedComputerLicensingEnabled()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (key != null)
            {
                return Convert.ToInt32(key.GetValue("SharedComputerLicensing", 0)) != 0;
            }
        }
        catch
        {
        }
        return false;
    }

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
