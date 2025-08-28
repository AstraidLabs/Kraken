// StartupGuard.cs
using System;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using Microsoft.Win32;

namespace Kraken
{
    /// <summary>
    /// Blokuje spuštění aplikace v nevhodných stavech systému:
    /// - Během dokončování instalace Windows / OOBE / WinPE
    /// - V nouzovém režimu (Safe Mode)
    /// - Během instalace/servisování Windows Update (CBS/TiWorker)
    /// - Během probíhající MSI instalace/odinstalace
    /// - Během instalace/aktualizace Office (Click-to-Run)
    /// 
    /// Vše je "read-only" (bez zásahů do služeb/registrů).
    /// </summary>
    public static class StartupGuard
    {
        private const string AppTitle = "Kraken";
        private const int SM_CLEANBOOT = 0x43;

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
            try
            {
                MessageBox.Show(message, AppTitle, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch
            {
                // V některých režimech (např. WinPE/Setup) nemusí být GUI dostupné – ignoruj.
            }
            Environment.Exit(exitCode);
        }

        // --------------------------
        // A) Windows Setup / OOBE / WinPE
        // --------------------------
        private static bool IsSetupInProgress()
        {
            try
            {
                using var setup = OpenHKLM(@"SYSTEM\Setup");
                if (setup != null)
                {
                    if (ToInt(setup.GetValue("SystemSetupInProgress", 0)) != 0) return true;
                    if (ToInt(setup.GetValue("OOBEInProgress", 0)) != 0) return true;
                    if (ToInt(setup.GetValue("SetupPhase", 0)) > 0) return true;
                }

                // WinPE/Setup indikátor
                using var miniNt = OpenHKLM(@"SYSTEM\CurrentControlSet\Control\MiniNT");
                if (miniNt != null) return true;
            }
            catch { }
            return false;
        }

        // --------------------------
        // B) Safe Mode
        // --------------------------
        private static bool IsSafeMode()
        {
            try
            {
                int boot = GetSystemMetrics(SM_CLEANBOOT);
                if (boot == 1 || boot == 2) return true; // Minimal/Network
            }
            catch { }

            // Další signály Safe Mode
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("SAFEBOOT_OPTION")))
                return true;

            try
            {
                using var key = OpenHKLM(@"SYSTEM\CurrentControlSet\Control\SafeBoot\Option");
                if (key != null && ToInt(key.GetValue("OptionValue", 0)) > 0) return true;
            }
            catch { }

            return false;
        }

        // --------------------------
        // C) Windows Update / CBS (robustní, s předběžnou kontrolou a krátkým CPU samplingem)
        // --------------------------
        private static bool IsWindowsUpdateInProgress()
        {
            const double cpuThresholdPercent = 2.5;                  // agregát přes všechna jádra
            TimeSpan sampleWindow = TimeSpan.FromMilliseconds(800);  // rychlé, ale dostačující okno

            bool tiServiceRunning = IsServiceRunning("TrustedInstaller"); // silný signál
            bool usoSvcActive = IsServiceActive("UsoSvc");            // Running | StartPending
            bool wuauservActive = IsServiceActive("wuauserv");

            string[] wuProcNames = { "TiWorker", "TrustedInstaller", "MoUsoCoreWorker" };
            bool anyWuProcPresent = wuProcNames.Any(n => Process.GetProcessesByName(n).Length > 0);

            // Předběžný rychlý návrat: pokud neběží TI, nejsou přítomny WU procesy
            // a není aktivní orchestrátor s MoUsoCoreWorker, nedělej drahý sampling.
            bool orchestratorActive = (usoSvcActive || wuauservActive) && Process.GetProcessesByName("MoUsoCoreWorker").Length > 0;
            if (!tiServiceRunning && !anyWuProcPresent && !orchestratorActive)
                return false;

            // CPU aktivita – měříme až teď
            double wuCpuPercent = GetAggregateCpuPercent(wuProcNames, sampleWindow);
            bool wuCpuActive = wuCpuPercent >= cpuThresholdPercent;

            // COM IsBusy ber v potaz jen pokud je aspoň nějaká systémová aktivita
            bool wuServiceSignal = tiServiceRunning || usoSvcActive || wuauservActive || anyWuProcPresent;
            bool wuComBusy = wuServiceSignal && TryIsWuInstallerBusy();

            int score = 0;
            if (tiServiceRunning) score++;   // silný hlas
            if (anyWuProcPresent) score++;
            if (wuCpuActive) score++;
            if (wuComBusy) score++;

            // Potřebujeme alespoň 2 nezávislé hlasy a zároveň "silný" (TI) NEBO CPU aktivitu
            return (score >= 2) && (tiServiceRunning || wuCpuActive);
        }

        // --------------------------
        // D) MSI instalace/odinstalace v běhu
        // --------------------------
        private static bool IsMsiInProgress()
        {
            try
            {
                using var key = OpenHKLM(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\InProgress");
                if (key != null) return true;
            }
            catch { }

            bool msiserverActive = IsServiceActive("msiserver");
            return msiserverActive && Process.GetProcessesByName("msiexec").Length > 0;
        }

        // --------------------------
        // E) Office Click-to-Run aktivní
        // --------------------------
        private static bool IsOfficeClickToRunInProgress()
        {
            bool c2rActive = IsServiceActive("ClickToRunSvc");
            if (!c2rActive) return false;

            if (Process.GetProcessesByName("OfficeClickToRun").Length > 0 ||
                Process.GetProcessesByName("OfficeC2RClient").Length > 0)
                return true;

            // Volitelný doplňkový signál (heuristika):
            // UpdateState != "Idle" = pravděpodobně aktivní update/stream
            try
            {
                string state = ReadStringHKLM(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "UpdateState");
                if (!string.IsNullOrWhiteSpace(state) &&
                    !state.Equals("Idle", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }

            return false;
        }

        // --------------------------
        // Helpers: služby, CPU sampling, WU COM IsBusy, registry
        // --------------------------
        private static bool IsServiceRunning(string name)
        {
            try
            {
                using var sc = new ServiceController(name);
                return sc.Status == ServiceControllerStatus.Running;
            }
            catch { return false; }
        }

        private static bool IsServiceActive(string name)
        {
            try
            {
                using var sc = new ServiceController(name);
                return sc.Status == ServiceControllerStatus.Running ||
                       sc.Status == ServiceControllerStatus.StartPending;
            }
            catch { return false; }
        }

        private static double GetAggregateCpuPercent(IEnumerable<string> procNames, TimeSpan window)
        {
            var set = new HashSet<string>(procNames, StringComparer.OrdinalIgnoreCase);

            TimeSpan before = SumCpu(set);
            var sw = Stopwatch.StartNew();
            Thread.Sleep(window);
            TimeSpan after = SumCpu(set);
            sw.Stop();

            double elapsedMs = sw.Elapsed.TotalMilliseconds;
            if (elapsedMs <= 0) return 0;

            double deltaMs = (after - before).TotalMilliseconds;
            double cpuPct = (deltaMs / (elapsedMs * Environment.ProcessorCount)) * 100.0;
            return cpuPct < 0 ? 0 : cpuPct;

            static TimeSpan SumCpu(HashSet<string> names)
            {
                TimeSpan sum = TimeSpan.Zero;
                foreach (var p in Process.GetProcesses())
                {
                    try
                    {
                        if (names.Contains(p.ProcessName))
                            sum += p.TotalProcessorTime;
                    }
                    catch
                    {
                        // Access denied / rychle končící procesy – ignoruj
                    }
                    finally
                    {
                        try { p.Dispose(); } catch { }
                    }
                }
                return sum;
            }
        }

        private static bool TryIsWuInstallerBusy()
        {
            // Reflexe COMu "Microsoft.Update.Session" → CreateUpdateInstaller().IsBusy
            try
            {
                var t = Type.GetTypeFromProgID("Microsoft.Update.Session", throwOnError: false);
                if (t == null) return false;

                object? session = Activator.CreateInstance(t);
                if (session == null) return false;

                object? installer = session.GetType().GetMethod("CreateUpdateInstaller")?.Invoke(session, null);
                if (installer == null)
                {
                    ReleaseCom(session);
                    return false;
                }

                bool busy = false;
                var prop = installer.GetType().GetProperty("IsBusy");
                if (prop != null)
                    busy = prop.GetValue(installer) is bool b && b;

                ReleaseCom(installer);
                ReleaseCom(session);
                return busy;
            }
            catch
            {
                return false;
            }

            static void ReleaseCom(object o)
            {
                try { Marshal.FinalReleaseComObject(o); } catch { }
            }
        }

        private static RegistryKey? OpenHKLM(string subKey)
        {
            try
            {
                var k64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                                     .OpenSubKey(subKey, writable: false);
                if (k64 != null) return k64;
            }
            catch { }

            try
            {
                return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                                  .OpenSubKey(subKey, writable: false);
            }
            catch
            {
                return null;
            }
        }

        private static int ToInt(object? value)
        {
            try
            {
                return value switch
                {
                    int i => i,
                    string s when int.TryParse(s, out var j) => j,
                    _ => 0
                };
            }
            catch { return 0; }
        }

        private static string ReadStringHKLM(string subKey, string valueName)
        {
            try
            {
                using var k = OpenHKLM(subKey);
                return k?.GetValue(valueName)?.ToString() ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);
    }
}
