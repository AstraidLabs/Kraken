using ActivationInspector.Domain;
using ActivationInspector.Infrastructure.Export;
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
        var licenses = new List<WindowsLicense>
        {
            new() { Name = "Windows", LicenseStatus = "Licensed" }
        };

        var service = new ExportService();
        string json = await service.ExportJsonAsync(licenses);
        var doc = JsonDocument.Parse(json);
        Assert.AreEqual("Windows", doc.RootElement[0].GetProperty("Name").GetString());
        Assert.AreEqual("Licensed", doc.RootElement[0].GetProperty("LicenseStatus").GetString());
    }
}
