using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ActivationInspector.Core.Models;

namespace ActivationInspector.Core.Licensing;

/// <summary>
/// Facade that exposes simplified access points for retrieving licensing
/// information. Additional providers (Office, diagnostics, etc.) can be added
/// following the same pattern.
/// </summary>
public class LicensingService
{
    private readonly WindowsLicensingProvider _windowsProvider = new();

    public Task<IReadOnlyList<WindowsLicenseDto>> GetWindowsLicensesAsync(CancellationToken token = default)
        => _windowsProvider.GetLicensesAsync(token);
}
