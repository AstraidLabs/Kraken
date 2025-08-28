using System.Windows;
using ActivationInspector.Application.Interfaces;
using ActivationInspector.Infrastructure.Export;
using ActivationInspector.Infrastructure.Licensing;
using ActivationInspector.UI.ViewModels;

namespace ActivationInspector.UI;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        ILicensingService licensingService = new LicensingService();
        IExportService exportService = new ExportService();
        var vm = new MainViewModel(licensingService, exportService);
        var window = new MainWindow { DataContext = vm };
        window.Show();
    }
}
