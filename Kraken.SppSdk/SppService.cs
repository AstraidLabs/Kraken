using Serilog;

namespace Kraken.SppSdk;

public sealed class SppService : ISppService
{
    private readonly ILogger _logger;

    public SppService(ILogger logger) => _logger = logger;

    public async Task<ISppSession> OpenSessionAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(OpenSessionAsync));
        var handle = await Task.Run(() =>
        {
            int hr = SppApi.SLOpen(out var h);
            if (hr != 0) throw new SppException(hr, nameof(SppApi.SLOpen));
            return h;
        }, ct);
        _logger.Debug("Exiting {Method}", nameof(OpenSessionAsync));
        return new SppSession(handle, _logger);
    }

    public async Task<WindowsLicenseInfo> GetWindowsLicenseAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetWindowsLicenseAsync));
        await using var session = await OpenSessionAsync(ct);
        var result = await session.GetWindowsLicenseAsync(ct);
        _logger.Debug("Exiting {Method}", nameof(GetWindowsLicenseAsync));
        return result;
    }

    public async Task<IReadOnlyCollection<OfficeLicenseInfo>> GetOfficeLicensesAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetOfficeLicensesAsync));
        await using var session = await OpenSessionAsync(ct);
        var result = await session.GetOfficeLicensesAsync(ct);
        _logger.Debug("Exiting {Method}", nameof(GetOfficeLicensesAsync));
        return result;
    }

    public async Task<SppLicenseStatus[]> GetLicensingStatusAsync(Guid appId, Guid skuId, CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetLicensingStatusAsync));
        await using var session = await OpenSessionAsync(ct);
        var result = await session.GetLicensingStatusAsync(appId, skuId, ct);
        _logger.Debug("Exiting {Method}", nameof(GetLicensingStatusAsync));
        return result;
    }

    public async Task<string?> GetApplicationInfoAsync(Guid appId, string name, CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetApplicationInfoAsync));
        await using var session = await OpenSessionAsync(ct);
        var result = await session.GetApplicationInfoAsync(appId, name, ct);
        _logger.Debug("Exiting {Method}", nameof(GetApplicationInfoAsync));
        return result;
    }

    public async Task<IReadOnlyCollection<VNextLicense>> GetVNextLicensesAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetVNextLicensesAsync));
        await using var session = await OpenSessionAsync(ct);
        var result = await session.GetVNextLicensesAsync(ct);
        _logger.Debug("Exiting {Method}", nameof(GetVNextLicensesAsync));
        return result;
    }

    public async Task<SubStatus> GetSubscriptionStatusAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetSubscriptionStatusAsync));
        await using var session = await OpenSessionAsync(ct);
        var result = await session.GetSubscriptionStatusAsync(ct);
        _logger.Debug("Exiting {Method}", nameof(GetSubscriptionStatusAsync));
        return result;
    }

    public async Task EnsureSppServiceRunningAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(EnsureSppServiceRunningAsync));
        await using var session = await OpenSessionAsync(ct);
        await session.EnsureSppServiceRunningAsync(ct);
        _logger.Debug("Exiting {Method}", nameof(EnsureSppServiceRunningAsync));
    }
}
