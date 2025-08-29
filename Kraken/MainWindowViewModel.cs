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
        }
    }

    /// <summary>Command to refresh data.</summary>
    public ICommand RefreshCommand { get; }

    /// <summary>Command to save JSON report.</summary>
    public ICommand SaveJsonCommand { get; }

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

    /// <summary>
    /// Initializes a new instance of the <see cref="MainWindowViewModel"/> class.
    /// </summary>
    public MainWindowViewModel()
    {
        RefreshCommand = new RelayCommand(_ => Refresh());
        SaveJsonCommand = new RelayCommand(_ => SaveJson(), _ => Summary != null);
        SystemInfo = BuildSystemInfo();
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
            var dlg = new SaveFileDialog { Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*" };
            if (dlg.ShowDialog() == true)
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(Summary, options);
                File.WriteAllText(dlg.FileName, json);
            }
        }
        catch
        {
            // ignore
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
