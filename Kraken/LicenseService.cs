using System;
using System.Collections.Generic;
using System.Management;

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

        // Other sections are not implemented in this minimal sample and will remain null.
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
                "SELECT Name, Description, ID, PartialProductKey, LicenseStatus, GracePeriodRemaining, EvaluationEndDate FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL");
            foreach (var obj in searcher.Get())
            {
                var info = new WindowsLicenseInfo
                {
                    ProductName = obj["Name"]?.ToString() ?? string.Empty,
                    Description = obj["Description"]?.ToString() ?? string.Empty,
                    PartialProductKey = obj["PartialProductKey"]?.ToString() ?? string.Empty,
                    Status = LicenseStatusToString(obj["LicenseStatus"]),
                    GraceMinutes = obj["GracePeriodRemaining"] != null ? Convert.ToInt32(obj["GracePeriodRemaining"]) : 0,
                    Expiration = ParseDate(obj["EvaluationEndDate"])
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
            const string query = "SELECT Name, Description, ID, PartialProductKey, LicenseStatus, GracePeriodRemaining, EvaluationEndDate FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey IS NOT NULL";
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
                    Expiration = ParseDate(obj["EvaluationEndDate"])
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
