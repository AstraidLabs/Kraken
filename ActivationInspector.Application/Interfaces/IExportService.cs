using ActivationInspector.Domain;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ActivationInspector.Application.Interfaces;

public interface IExportService
{
    Task<string> ExportJsonAsync(IEnumerable<WindowsLicense> windowsLicenses);
}
