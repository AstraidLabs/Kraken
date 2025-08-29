namespace Kraken.SppSdk;

public interface ISppService
{
    Task<ISppSession> OpenSessionAsync(CancellationToken ct = default);
    Task<WindowsLicenseInfo> GetWindowsLicenseAsync(CancellationToken ct = default);
    Task<IReadOnlyCollection<OfficeLicenseInfo>> GetOfficeLicensesAsync(CancellationToken ct = default);
    Task<SppLicenseStatus[]> GetLicensingStatusAsync(Guid appId, Guid skuId, CancellationToken ct = default);
    Task<string?> GetApplicationInfoAsync(Guid appId, string name, CancellationToken ct = default);
    Task<IReadOnlyCollection<VNextLicense>> GetVNextLicensesAsync(CancellationToken ct = default);
    Task<SubStatus> GetSubscriptionStatusAsync(CancellationToken ct = default);
    Task EnsureSppServiceRunningAsync(CancellationToken ct = default);
}

public interface ISppSession : IAsyncDisposable
{
    Task<WindowsLicenseInfo> GetWindowsLicenseAsync(CancellationToken ct = default);
    Task<IReadOnlyCollection<OfficeLicenseInfo>> GetOfficeLicensesAsync(CancellationToken ct = default);
    Task<SppLicenseStatus[]> GetLicensingStatusAsync(Guid appId, Guid skuId, CancellationToken ct = default);
    Task<string?> GetApplicationInfoAsync(Guid appId, string name, CancellationToken ct = default);
    Task<IReadOnlyCollection<VNextLicense>> GetVNextLicensesAsync(CancellationToken ct = default);
    Task<SubStatus> GetSubscriptionStatusAsync(CancellationToken ct = default);
    Task EnsureSppServiceRunningAsync(CancellationToken ct = default);
}
