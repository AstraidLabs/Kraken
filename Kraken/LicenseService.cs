using System;
using System.Collections.Generic;
using System.Management;
using Microsoft.Win32;

namespace Kraken;

/// <summary>
/// High level helpers that use SppApi and WMI/registry to build license information objects.
/// </summary>
public static class LicenseService
{
    /// <summary>
    /// Retrieves a summary of all licence information available on the machine.
    /// </summary>
    public static LicenseSummary GetLicenseSummary()
    {
        var summary = new LicenseSummary();
        if (SppApi.TryOpenSession(out var handle))
        {
            using (handle)
            {
                summary.WindowsLicense = GetWindowsLicense(handle);
                summary.OfficeLicenses = GetOfficeLicenses(handle);
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
    /// an installation ID is generated.
    /// </summary>
    public static WindowsLicenseInfo? GetWindowsLicense(SppApi.SppSafeHandle? h)
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
                    Status = LicenseStatusToString(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null ? Convert.ToInt32(obj["GracePeriodRemaining"]) : 0,
                    Expiration = ParseDate(obj["EvaluationEndDate"]),
                    Channel = obj["ProductKeyChannel"]?.ToString() ?? string.Empty,
                    EvaluationEndDate = ParseDate(obj["EvaluationEndDate"])
                };
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
    /// Retrieves Office licence information using WMI. If SPP handle provided, installation IDs
    /// are generated for each product.
    /// </summary>
    public static List<OfficeLicenseInfo> GetOfficeLicenses(SppApi.SppSafeHandle? h)
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
                    Status = LicenseStatusToString(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null ? Convert.ToInt32(obj["GracePeriodRemaining"]) : 0,
                    Expiration = ParseDate(obj["EvaluationEndDate"]),
                    Channel = obj["ProductKeyChannel"]?.ToString() ?? string.Empty
                };
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
                "SELECT KeyManagementServiceCurrentCount, KeyManagementServiceTotalRequests, KeyManagementServiceFailedRequests, KeyManagementServiceMachineName, KeyManagementServicePort, KeyManagementServiceMachineIpAddress FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
            {
                return new KmsServerInfo
                {
                    CurrentClients = Convert.ToInt32(obj["KeyManagementServiceCurrentCount"] ?? 0),
                    TotalRequests = Convert.ToInt32(obj["KeyManagementServiceTotalRequests"] ?? 0),
                    FailedRequests = Convert.ToInt32(obj["KeyManagementServiceFailedRequests"] ?? 0),
                    Name = obj["KeyManagementServiceMachineName"]?.ToString() ?? string.Empty,
                    Port = obj["KeyManagementServicePort"] != null ? Convert.ToInt32(obj["KeyManagementServicePort"]) : 0,
                    IPAddress = obj["KeyManagementServiceMachineIpAddress"]?.ToString() ?? string.Empty
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
                "SELECT ClientMachineID, KeyManagementServiceMachineName, KeyManagementServicePort, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, DiscoveredKeyManagementServiceMachineIpAddress, VLActivationInterval FROM SoftwareLicensingService");
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
                    RenewalInterval = obj["VLActivationInterval"] != null ? Convert.ToInt32(obj["VLActivationInterval"]) : 0
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
                    State = status.dwState
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
                "SELECT ADActivationObjectName, ADActivationObjectDN, KeyManagementServiceProductKeyID, KeyManagementServiceSkuID FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
            {
                return new AdActivationInfo
                {
                    ObjectName = obj["ADActivationObjectName"]?.ToString() ?? string.Empty,
                    ObjectDN = obj["ADActivationObjectDN"]?.ToString() ?? string.Empty,
                    CsvlkPid = obj["KeyManagementServiceProductKeyID"]?.ToString() ?? string.Empty,
                    CsvlkSkuId = obj["KeyManagementServiceSkuID"]?.ToString() ?? string.Empty
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
                "SELECT InheritedActivationID, VirtualizationHostMachineName, VirtualizationHostDigitalPid2, VirtualizationActivationTime FROM SoftwareLicensingService");
            foreach (var obj in searcher.Get())
            {
                return new VmActivationInfo
                {
                    InheritedActivationId = obj["InheritedActivationID"]?.ToString() ?? string.Empty,
                    HostMachineName = obj["VirtualizationHostMachineName"]?.ToString() ?? string.Empty,
                    HostDigitalPid2 = obj["VirtualizationHostDigitalPid2"]?.ToString() ?? string.Empty,
                    ActivationTime = ParseDate(obj["VirtualizationActivationTime"])
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

    private static string LicenseStatusToString(object? statusObj)
    {
        if (statusObj == null) return string.Empty;
        int status = Convert.ToInt32(statusObj);
        return status switch
        {
            0 => "Unlicensed",
            1 => "Licensed",
            2 => "Grace",
            3 => "Notification",
            4 => "Expired",
            5 => "Extended Grace",
            _ => "Unknown"
        };
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
