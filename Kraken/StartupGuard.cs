using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Windows;
using Microsoft.Win32;

namespace Kraken;

public static class StartupGuard
{
    private const string AppTitle = "ActivationInspector";

    public static void CheckAndExitIfBlocked()
    {
        if (IsSetupInProgress())
            Block("Unavailable while Windows is completing setup. Please try again later.", 10);

        if (IsSafeMode())
            Block("Unavailable in Safe Mode. Please restart Windows normally and try again.", 11);

        if (IsWindowsUpdateInProgress())
            Block("Unavailable while Windows is installing updates. Please try again later.", 12);

        if (IsMsiInProgress())
            Block("Unavailable while an installer is in progress. Please try again later.", 13);

        if (IsOfficeClickToRunInProgress())
            Block("Unavailable while Office is updating. Please try again later.", 14);
    }

    private static void Block(string message, int exitCode)
    {
        MessageBox.Show(message, AppTitle, MessageBoxButton.OK, MessageBoxImage.Information);
        Environment.Exit(exitCode);
    }

    private static bool IsSetupInProgress()
    {
        try
        {
            using var setup = Registry.LocalMachine.OpenSubKey(@"SYSTEM\\Setup");
            if (setup != null)
            {
                if (Convert.ToInt32(setup.GetValue("SystemSetupInProgress", 0)) != 0)
                    return true;
                if (Convert.ToInt32(setup.GetValue("OOBEInProgress", 0)) != 0)
                    return true;
                if (Convert.ToInt32(setup.GetValue("SetupPhase", 0)) > 0)
                    return true;
            }

            using var miniNt = Registry.LocalMachine.OpenSubKey(@"SYSTEM\\CurrentControlSet\\Control\\MiniNT");
            if (miniNt != null)
                return true;
        }
        catch
        {
        }

        return false;
    }

    private static bool IsSafeMode()
    {
        try
        {
            var boot = GetSystemMetrics(SM_CLEANBOOT);
            if (boot == 1 || boot == 2)
                return true;
        }
        catch
        {
        }

        if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("SAFEBOOT_OPTION")))
            return true;

        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\\CurrentControlSet\\Control\\SafeBoot\\Option");
            if (key != null && Convert.ToInt32(key.GetValue("OptionValue", 0)) > 0)
                return true;
        }
        catch
        {
        }

        return false;
    }

    private static bool IsWindowsUpdateInProgress()
    {
        if (IsServiceActive("TrustedInstaller") ||
            IsServiceActive("UsoSvc") ||
            IsServiceActive("wuauserv"))
            return true;

        if (Process.GetProcessesByName("TiWorker").Length > 0 ||
            Process.GetProcessesByName("TrustedInstaller").Length > 0 ||
            Process.GetProcessesByName("MoUsoCoreWorker").Length > 0)
            return true;

        try
        {
            var t = Type.GetTypeFromProgID("Microsoft.Update.Session", throwOnError: false);
            if (t != null)
            {
                dynamic session = Activator.CreateInstance(t);
                dynamic installer = session.CreateUpdateInstaller();
                if (installer.IsBusy == true)
                    return true;
            }
        }
        catch
        {
        }

        return false;
    }

    private static bool IsMsiInProgress()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Installer\\InProgress");
            if (key != null)
                return true;
        }
        catch
        {
        }

        var serviceActive = IsServiceActive("msiserver");
        if (serviceActive && Process.GetProcessesByName("msiexec").Length > 0)
            return true;

        return false;
    }

    private static bool IsOfficeClickToRunInProgress()
    {
        var serviceActive = IsServiceActive("ClickToRunSvc");
        if (!serviceActive)
            return false;

        if (Process.GetProcessesByName("OfficeClickToRun").Length > 0 ||
            Process.GetProcessesByName("OfficeC2RClient").Length > 0)
            return true;

        return false;
    }

    private static bool IsServiceActive(string name)
    {
        try
        {
            using var svc = new ServiceController(name);
            return svc.Status == ServiceControllerStatus.Running ||
                   svc.Status == ServiceControllerStatus.StartPending;
        }
        catch
        {
            return false;
        }
    }

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);
    private const int SM_CLEANBOOT = 0x43;
}
