using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Input;
using Microsoft.Win32;

namespace Kraken;

/// <summary>
/// View model for the main window.
/// </summary>
public class MainWindowViewModel : INotifyPropertyChanged
{
    private LicenseSummary _summary = new();

    /// <summary>Gets or sets the licence summary.</summary>
    public LicenseSummary Summary
    {
        get => _summary;
        set
        {
            _summary = value;
            OnPropertyChanged(nameof(Summary));
            OnPropertyChanged(nameof(WindowsLicenseList));
            OnPropertyChanged(nameof(KmsServerList));
            OnPropertyChanged(nameof(KmsClientList));
            OnPropertyChanged(nameof(OfficeLicenses));
        }
    }

    /// <summary>Information about the system.</summary>
    public string SystemInfo { get; }

    /// <summary>Collection for DataGrid binding.</summary>
    public IEnumerable<WindowsLicenseInfo> WindowsLicenseList =>
        Summary.WindowsLicense != null ? new[] { Summary.WindowsLicense } : Array.Empty<WindowsLicenseInfo>();

    /// <summary>Collection for DataGrid binding.</summary>
    public IEnumerable<KmsServerInfo> KmsServerList =>
        Summary.KmsServer != null ? new[] { Summary.KmsServer } : Array.Empty<KmsServerInfo>();

    /// <summary>Collection for DataGrid binding.</summary>
    public IEnumerable<KmsClientInfo> KmsClientList =>
        Summary.KmsClient != null ? new[] { Summary.KmsClient } : Array.Empty<KmsClientInfo>();

    /// <summary>Collection for Office licences.</summary>
    public IEnumerable<OfficeLicenseInfo> OfficeLicenses => Summary.OfficeLicenses;

    public ICommand RefreshCommand { get; }

    public ICommand SaveJsonCommand { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="MainWindowViewModel"/> class.
    /// </summary>
    public MainWindowViewModel()
    {
        SystemInfo = BuildSystemInfo();
        RefreshCommand = new RelayCommand(_ => Refresh());
        SaveJsonCommand = new RelayCommand(_ => SaveJson());
        Refresh();
    }

    private void Refresh()
    {
        try
        {
            Summary = LicenseService.GetLicenseSummary();
        }
        catch
        {
            Summary = new LicenseSummary();
        }
    }

    private void SaveJson()
    {
        try
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var path = Path.Combine(desktop, "LicenseSummary.json");
            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(path, JsonSerializer.Serialize(Summary, options));
        }
        catch
        {
        }
    }


    private string BuildSystemInfo()
    {
        string os = RuntimeInformation.OSDescription + " (" + RuntimeInformation.OSArchitecture + ")";
        string ps = GetPowerShellVersion();
        return os + " - PowerShell " + ps;
    }

    private string GetPowerShellVersion()
    {
        try
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine", "PowerShellVersion", "Unknown")?.ToString() ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    /// <inheritdoc />
    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
