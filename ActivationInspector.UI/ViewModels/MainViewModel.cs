using System.Collections.ObjectModel;
using System.Threading.Tasks;
using ActivationInspector.Application.Interfaces;
using ActivationInspector.Domain;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ActivationInspector.UI.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ILicensingService _licensingService;
    private readonly IExportService _exportService;

    public ObservableCollection<WindowsLicense> WindowsLicenses { get; } = new();

    [ObservableProperty]
    private bool allFlag;

    [ObservableProperty]
    private bool dlvFlag;

    [ObservableProperty]
    private bool iidFlag;

    [ObservableProperty]
    private bool noClearFlag;

    [ObservableProperty]
    private string logText = string.Empty;

    public IAsyncRelayCommand RefreshCommand { get; }
    public IAsyncRelayCommand ExportJsonCommand { get; }

    public MainViewModel(ILicensingService licensingService, IExportService exportService)
    {
        _licensingService = licensingService;
        _exportService = exportService;
        RefreshCommand = new AsyncRelayCommand(RefreshAsync);
        ExportJsonCommand = new AsyncRelayCommand(ExportJsonAsync);
    }

    private async Task RefreshAsync()
    {
        var licenses = await _licensingService.GetWindowsLicensesAsync();
        WindowsLicenses.Clear();
        foreach (var l in licenses)
        {
            WindowsLicenses.Add(l);
        }
        if (!NoClearFlag)
            LogText = string.Empty;
        LogText += $"Scanned {licenses.Count} license(s)\n";
    }

    private async Task ExportJsonAsync()
    {
        var json = await _exportService.ExportJsonAsync(WindowsLicenses);
        LogText += $"Exported JSON length: {json.Length}\n";
    }
}
