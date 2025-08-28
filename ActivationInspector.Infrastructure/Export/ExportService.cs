using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using ActivationInspector.Application.Interfaces;
using ActivationInspector.Domain;

namespace ActivationInspector.Infrastructure.Export;

/// <summary>
/// Handles simple data export scenarios (JSON/TXT/CSV). Only JSON is implemented
/// for brevity, the other formats can be added later following the same idea.
/// </summary>
public class ExportService : IExportService
{
    public Task<string> ExportJsonAsync(IEnumerable<WindowsLicense> windowsLicenses)
    {
        return Task.Run(() =>
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            return JsonSerializer.Serialize(windowsLicenses, options);
        });
    }
}
