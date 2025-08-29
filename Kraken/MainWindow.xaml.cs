using System;
using System.Windows;
using System.Management;
using System.Runtime.InteropServices;

namespace Kraken;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private LicenseSummary? _summary;

    public MainWindow()
    {
        InitializeComponent();
        DisplaySystemInfo();
        RefreshData();
    }

    private void RefreshData()
    {
        _summary = LicenseService.GetLicenseSummary();
        if (_summary != null)
        {
            LicensesGrid.ItemsSource = _summary.OfficeLicenses;
        }
    }

    private void DisplaySystemInfo()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem");
            foreach (var obj in searcher.Get())
            {
                var caption = obj["Caption"]?.ToString()?.Trim() ?? "Windows";
                var version = obj["Version"]?.ToString() ?? string.Empty;
                var build = obj["BuildNumber"]?.ToString() ?? string.Empty;
                var arch = obj["OSArchitecture"]?.ToString() ?? RuntimeInformation.OSArchitecture.ToString();
                SystemInfoText.Text = $"{caption} Version {version} (Build {build}, {arch})";
                break;
            }
        }
        catch (ManagementException)
        {
            SystemInfoText.Text = $"{RuntimeInformation.OSDescription} ({RuntimeInformation.OSArchitecture})";
        }
    }
}
