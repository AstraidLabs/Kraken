using ActivationInspector.Core.Export;
using ActivationInspector.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Text.Json;
using System.Threading.Tasks;

namespace ActivationInspector.Tests;

[TestClass]
public class ExportServiceTests
{
    [TestMethod]
    public async Task ExportJsonProducesExpectedStructure()
    {
        var licenses = new List<WindowsLicenseDto>
        {
            new() { Name = "Windows", LicenseStatus = "Licensed" }
        };

        string json = await ExportService.ExportJsonAsync(licenses);
        var doc = JsonDocument.Parse(json);
        Assert.AreEqual("Windows", doc.RootElement[0].GetProperty("Name").GetString());
        Assert.AreEqual("Licensed", doc.RootElement[0].GetProperty("LicenseStatus").GetString());
    }
}
