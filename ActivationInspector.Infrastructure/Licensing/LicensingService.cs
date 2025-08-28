using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ActivationInspector.Application.Interfaces;
using ActivationInspector.Domain;

namespace ActivationInspector.Infrastructure.Licensing;

/// <summary>
/// Facade that exposes simplified access points for retrieving licensing
/// information. Additional providers (Office, diagnostics, etc.) can be added
/// following the same pattern.
/// </summary>
public class LicensingService : ILicensingService
{
    private readonly WindowsLicensingProvider _windowsProvider = new();

    public Task<IReadOnlyList<WindowsLicense>> GetWindowsLicensesAsync(CancellationToken token = default)
        => _windowsProvider.GetLicensesAsync(token);
}
