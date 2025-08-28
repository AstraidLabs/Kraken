using System;
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
        LoadLicensesFromClass("SoftwareLicensingProduct", "Windows");
        LoadLicensesFromClass("OfficeSoftwareProtectionProduct", "Office");
    }

    private void LoadLicensesFromClass(string wmiClass, string application)
    {
        var query =
            $"SELECT Name, Description, ID, PartialProductKey, LicenseStatus, GracePeriodRemaining, EvaluationEndDate FROM {wmiClass} WHERE PartialProductKey IS NOT NULL";

        try
        {
            using var searcher = new ManagementObjectSearcher(query);
            foreach (var obj in searcher.Get())
            {
                var name = obj["Name"]?.ToString() ?? string.Empty;
                var description = obj["Description"]?.ToString() ?? string.Empty;
                var activationId = obj["ID"]?.ToString() ?? string.Empty;
                var partialKey = obj["PartialProductKey"]?.ToString() ?? string.Empty;
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
                var grace = obj["GracePeriodRemaining"] != null ? Convert.ToInt32(obj["GracePeriodRemaining"]) : 0;
                DateTime? evalEnd = null;
                if (obj["EvaluationEndDate"] != null)
                {
                    try
                    {
                        evalEnd = ManagementDateTimeConverter.ToDateTime(obj["EvaluationEndDate"].ToString());
                    }
                    catch
                    {
                        evalEnd = null;
                    }
                }
                Licenses.Add(new LicenseInfo
                {
                    Application = application,
                    Name = name,
                    Description = description,
                    ActivationId = activationId,
                    PartialProductKey = partialKey,
                    Status = status,
                    GraceMinutes = grace,
                    EvaluationEndDate = evalEnd
                });
            }
        }
        catch (ManagementException)
        {
            // The WMI class may not exist on this system; ignore and continue.
        }
    }
}
