using System.Collections.ObjectModel;
using System.Management;
using System.Windows;

namespace Kraken;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private ObservableCollection<LicenseInfo> Licenses { get; } = new();

    public MainWindow()
    {
        InitializeComponent();
        LicensesGrid.ItemsSource = Licenses;
        LoadLicenses();
    }

    private void LoadLicenses()
    {
        using var searcher = new ManagementObjectSearcher("SELECT Name, LicenseStatus FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL");
        foreach (var obj in searcher.Get())
        {
            var name = obj["Name"]?.ToString() ?? string.Empty;
            var statusCode = obj["LicenseStatus"] != null ? Convert.ToInt32(obj["LicenseStatus"]) : 0;
            var status = statusCode switch
            {
                0 => "Unlicensed",
                1 => "Licensed",
                2 => "Grace",
                3 => "Notification",
                4 => "Expired",
                5 => "Extended Grace",
                _ => "Unknown"
            };
            Licenses.Add(new LicenseInfo { Name = name, Status = status });
        }
    }
}
