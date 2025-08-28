using ActivationInspector.Domain;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ActivationInspector.Application.Interfaces;

public interface ILicensingService
{
    Task<IReadOnlyList<WindowsLicense>> GetWindowsLicensesAsync(CancellationToken token = default);
}
