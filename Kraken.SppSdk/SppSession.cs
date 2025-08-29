using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text.Json;
using Microsoft.Win32;
using Serilog;

namespace Kraken.SppSdk;

internal sealed class SppSession : ISppSession
{
    private readonly SppApi.SppSafeHandle _handle;
    private readonly ILogger _logger;

    internal SppSession(SppApi.SppSafeHandle handle, ILogger logger)
    {
        _handle = handle;
        _logger = logger;
    }

    public ValueTask DisposeAsync()
    {
        _handle.Dispose();
        return ValueTask.CompletedTask;
    }

    private static void ThrowIfFailed(int hr, string function) => throw new SppException(hr, function);

    private IReadOnlyCollection<Guid> GetSlids(Guid appId)
    {
        int hr = SppApi.SLGetSLIDList(_handle, appId, out var pGuids, out var count);
        if (hr != 0) ThrowIfFailed(hr, nameof(SppApi.SLGetSLIDList));
        try
        {
            var guids = new Guid[count];
            int size = Marshal.SizeOf<Guid>();
            for (int i = 0; i < count; i++)
            {
                IntPtr ptr = IntPtr.Add(pGuids, i * size);
                guids[i] = Marshal.PtrToStructure<Guid>(ptr);
            }
            return guids;
        }
        finally
        {
            if (pGuids != IntPtr.Zero) Marshal.FreeHGlobal(pGuids);
        }
    }

    private string GetProductKey(Guid slid)
    {
        int hr = SppApi.SLGetPKeyInformation(_handle, slid, "ProductKey", out uint t, out uint c, out IntPtr b);
        if (hr != 0) ThrowIfFailed(hr, nameof(SppApi.SLGetPKeyInformation));
        try { return SppApi.InterpretValue(t, c, b).S ?? string.Empty; }
        finally { if (b != IntPtr.Zero) Marshal.FreeHGlobal(b); }
    }

    private string GetProductSkuString(Guid skuId, string name)
    {
        int hr = SppApi.SLGetProductSkuInformation(_handle, skuId, name, out uint t, out uint c, out IntPtr b);
        if (hr != 0) ThrowIfFailed(hr, nameof(SppApi.SLGetProductSkuInformation));
        try { return SppApi.InterpretValue(t, c, b).S ?? string.Empty; }
        finally { if (b != IntPtr.Zero) Marshal.FreeHGlobal(b); }
    }

    private SppLicenseStatus[] GetLicensingStatus(Guid appId, Guid skuId)
    {
        int hr = SppApi.SLGetLicensingStatusInformation(_handle, appId, skuId, out IntPtr p, out uint c);
        if (hr != 0) ThrowIfFailed(hr, nameof(SppApi.SLGetLicensingStatusInformation));
        try { return SppApi.ParseLicensingStatus(p, c); }
        finally { if (p != IntPtr.Zero) Marshal.FreeHGlobal(p); }
    }

    public Task<WindowsLicenseInfo> GetWindowsLicenseAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetWindowsLicenseAsync));
        return Task.Run(() =>
        {
            var appId = new Guid("55C92734-D682-4D71-983E-D6EC3F16059F");
            foreach (var slid in GetSlids(appId))
            {
                var key = GetProductKey(slid);
                var status = GetLicensingStatus(appId, slid);
                DateTime? expiry = status.Length > 0 && status[0].ValidityFileTimeUtc != 0 ? DateTime.FromFileTimeUtc((long)status[0].ValidityFileTimeUtc) : null;
                var state = status.Length > 0 ? (LicenseState)status[0].Status : LicenseState.Unlicensed;
                var info = new WindowsLicenseInfo(slid, key, expiry, state);
                _logger.Debug("Exiting {Method}", nameof(GetWindowsLicenseAsync));
                return info;
            }
            _logger.Debug("Exiting {Method}", nameof(GetWindowsLicenseAsync));
            return new WindowsLicenseInfo(Guid.Empty, string.Empty, null, LicenseState.Unlicensed);
        }, ct);
    }

    public Task<IReadOnlyCollection<OfficeLicenseInfo>> GetOfficeLicensesAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetOfficeLicensesAsync));
        return Task.Run(() =>
        {
            var appId = new Guid("0FF1CE15-A989-479D-AF46-F275C6370663");
            var list = new List<OfficeLicenseInfo>();
            foreach (var slid in GetSlids(appId))
            {
                var key = GetProductKey(slid);
                var status = GetLicensingStatus(appId, slid);
                DateTime? expiry = status.Length > 0 && status[0].ValidityFileTimeUtc != 0 ? DateTime.FromFileTimeUtc((long)status[0].ValidityFileTimeUtc) : null;
                var state = status.Length > 0 ? (LicenseState)status[0].Status : LicenseState.Unlicensed;
                string edition = string.Empty;
                try { edition = GetProductSkuString(slid, "Name"); } catch { }
                list.Add(new OfficeLicenseInfo(slid, key, expiry, state, edition));
            }
            _logger.Debug("Exiting {Method}", nameof(GetOfficeLicensesAsync));
            return (IReadOnlyCollection<OfficeLicenseInfo>)list;
        }, ct);
    }

    public Task<SppLicenseStatus[]> GetLicensingStatusAsync(Guid appId, Guid skuId, CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetLicensingStatusAsync));
        return Task.Run(() =>
        {
            var status = GetLicensingStatus(appId, skuId);
            _logger.Debug("Exiting {Method}", nameof(GetLicensingStatusAsync));
            return status;
        }, ct);
    }

    public Task<string?> GetApplicationInfoAsync(Guid appId, string name, CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetApplicationInfoAsync));
        return Task.Run(() =>
        {
            int hr = SppApi.SLGetApplicationInformation(_handle, appId, name, out uint t, out uint c, out IntPtr b);
            if (hr != 0) ThrowIfFailed(hr, nameof(SppApi.SLGetApplicationInformation));
            try { return SppApi.InterpretValue(t, c, b).S; }
            finally { if (b != IntPtr.Zero) Marshal.FreeHGlobal(b); _logger.Debug("Exiting {Method}", nameof(GetApplicationInfoAsync)); }
        }, ct);
    }

    public Task<IReadOnlyCollection<VNextLicense>> GetVNextLicensesAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetVNextLicensesAsync));
        return Task.Run(() =>
        {
            var list = new List<VNextLicense>();
            using var office = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Office");
            if (office != null)
            {
                string baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Office", "Licenses");
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
                            if (ex.ValueKind == JsonValueKind.String && DateTime.TryParse(ex.GetString(), out var dt)) expiry = dt;
                            else if (ex.ValueKind == JsonValueKind.Number) expiry = DateTimeOffset.FromUnixTimeSeconds(ex.GetInt64()).DateTime;
                        }
                        list.Add(new VNextLicense(Path.GetFileName(file), release, status, expiry));
                    }
                }
            }
            _logger.Debug("Exiting {Method}", nameof(GetVNextLicensesAsync));
            return (IReadOnlyCollection<VNextLicense>)list;
        }, ct);
    }

    public Task<SubStatus> GetSubscriptionStatusAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(GetSubscriptionStatusAsync));
        return Task.Run(() =>
        {
            int hr = SppApi.ClipGetSubscriptionStatus(out var status);
            if (hr != 0) ThrowIfFailed(hr, nameof(SppApi.ClipGetSubscriptionStatus));
            _logger.Debug("Exiting {Method}", nameof(GetSubscriptionStatusAsync));
            return status;
        }, ct);
    }

    public Task EnsureSppServiceRunningAsync(CancellationToken ct = default)
    {
        _logger.Debug("Entering {Method}", nameof(EnsureSppServiceRunningAsync));
        return Task.Run(() =>
        {
            using var sc = new ServiceController("sppsvc");
            if (sc.Status != ServiceControllerStatus.Running)
            {
                sc.Start();
                sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
            }
            _logger.Debug("Exiting {Method}", nameof(EnsureSppServiceRunningAsync));
        }, ct);
    }
}
