using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;
using ActivationInspector.Core.Models;

namespace ActivationInspector.Core.Export;

/// <summary>
/// Handles simple data export scenarios (JSON/TXT/CSV). Only JSON is implemented
/// for brevity, the other formats can be added later following the same idea.
/// </summary>
public static class ExportService
{
    public static Task<string> ExportJsonAsync(IEnumerable<WindowsLicenseDto> windowsLicenses)
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
