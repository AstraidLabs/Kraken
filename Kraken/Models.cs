using System;
using System.Collections.Generic;

namespace Kraken;

/// <summary>
/// POCO models representing various SPP license information blocks.
/// </summary>
public class WindowsLicenseInfo
{
    public string ProductName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public int GraceMinutes { get; set; }
    public DateTime? Expiration { get; set; }
    public string PartialProductKey { get; set; } = string.Empty;
    public string Channel { get; set; } = string.Empty;
    public DateTime? EvaluationEndDate { get; set; }
    public DateTime? LastActivationTime { get; set; }
    public int LastActivationHResult { get; set; }
    public DateTime? KernelTimeBomb { get; set; }
    public DateTime? SystemTimeBomb { get; set; }
    public DateTime? TrustedTime { get; set; }
    public int RearmCount { get; set; }
    public bool IsDigitalLicense { get; set; }
    public bool IsWindowsGenuineLocal { get; set; }
    public List<string> LicenseKeyFiles { get; set; } = new();
}

public class OfficeLicenseInfo
{
    public string SLID { get; set; } = string.Empty;
    public string SkuId { get; set; } = string.Empty;
    public string ProductName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public int GraceMinutes { get; set; }
    public DateTime? Expiration { get; set; }
    public string PartialProductKey { get; set; } = string.Empty;
    public string Channel { get; set; } = string.Empty;
    public string OfflineInstallationId { get; set; } = string.Empty;
    public DateTime? TrustedTime { get; set; }
    public int RearmCount { get; set; }
    public List<string> LicenseKeyFiles { get; set; } = new();
}

public class KmsServerInfo
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

public class KmsClientInfo
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

public class SubscriptionInfo
{
    public bool Enabled { get; set; }
    public uint Sku { get; set; }
    public uint State { get; set; }
    public DateTime? LicenseExpiration { get; set; }
    public uint SubscriptionType { get; set; }
}

public class VNextInfo
{
    public Dictionary<string, string> ModePerProductReleaseId { get; set; } = new();
    public bool SharedComputerLicensing { get; set; }
    public List<string> Licenses { get; set; } = new();
    public List<string> LicenseKeyFiles { get; set; } = new();
}

public class AdActivationInfo
{
    public string ObjectName { get; set; } = string.Empty;
    public string ObjectDN { get; set; } = string.Empty;
    public string CsvlkPid { get; set; } = string.Empty;
    public string CsvlkSkuId { get; set; } = string.Empty;
    public string ActivationId { get; set; } = string.Empty;
}

public class VmActivationInfo
{
    public string InheritedActivationId { get; set; } = string.Empty;
    public string HostMachineName { get; set; } = string.Empty;
    public string HostDigitalPid2 { get; set; } = string.Empty;
    public DateTime? ActivationTime { get; set; }
    public string HostMachineID { get; set; } = string.Empty;
    public string HostMachineVersion { get; set; } = string.Empty;
}

public class LicenseSummary
{
    public WindowsLicenseInfo? WindowsLicense { get; set; }
    public List<OfficeLicenseInfo> OfficeLicenses { get; set; } = new();
    public KmsServerInfo? KmsServer { get; set; }
    public KmsClientInfo? KmsClient { get; set; }
    public SubscriptionInfo? Subscription { get; set; }
    public VNextInfo? VNext { get; set; }
    public AdActivationInfo? AdActivation { get; set; }
    public VmActivationInfo? VmActivation { get; set; }
}
