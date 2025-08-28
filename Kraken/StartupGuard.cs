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
    /// - Instalace Windows / OOBE / WinPE
    /// - Safe Mode
    /// - Servicing Windows Update (CBS/TiWorker)
    /// - Probíhající MSI instalace
    /// - Instalace/aktualizace Office (C2R/MSI)
    /// Vše read-only (bez zásahů do služeb/registrů).
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

            if (IsOfficeInstallOrUpdateInProgress())
                Block("Unavailable while Office is updating. Please try again later.", 14);
        }

        private static void Block(string message, int exitCode)
        {
            try { MessageBox.Show(message, AppTitle, MessageBoxButton.OK, MessageBoxImage.Information); }
            catch { /* UI nemusí být k dispozici (WinPE/Setup) */ }
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
        // C) Windows Update / CBS (předběžná kontrola + krátký CPU sampling)
        // --------------------------
        private static bool IsWindowsUpdateInProgress()
        {
            const double cpuThresholdPercent = 2.5;
            TimeSpan sampleWindow = TimeSpan.FromMilliseconds(800);

            bool tiServiceRunning = IsServiceRunning("TrustedInstaller"); // silný signál
            bool usoSvcActive = IsServiceActive("UsoSvc");            // Running | StartPending
            bool wuauservActive = IsServiceActive("wuauserv");

            string[] wuProcNames = { "TiWorker", "TrustedInstaller", "MoUsoCoreWorker" };
            bool anyWuProcPresent = wuProcNames.Any(n => Process.GetProcessesByName(n).Length > 0);

            // Pokud neběží TI, nejsou WU procesy a orchestrátor není aktivní, nečekáme.
            bool orchestratorActive = (usoSvcActive || wuauservActive) &&
                                      Process.GetProcessesByName("MoUsoCoreWorker").Length > 0;
            if (!tiServiceRunning && !anyWuProcPresent && !orchestratorActive)
                return false;

            // CPU aktivita – až nyní
            double wuCpuPercent = GetAggregateCpuPercent(wuProcNames, sampleWindow);
            bool wuCpuActive = wuCpuPercent >= cpuThresholdPercent;

            // COM IsBusy bereme v potaz jen pokud jsou služby/procesy naznačeny
            bool wuServiceSignal = tiServiceRunning || usoSvcActive || wuauservActive || anyWuProcPresent;
            bool wuComBusy = wuServiceSignal && TryIsWuInstallerBusy();

            int score = 0;
            if (tiServiceRunning) score++;   // silný hlas
            if (anyWuProcPresent) score++;
            if (wuCpuActive) score++;
            if (wuComBusy) score++;

            // Alespoň 2 hlasy a zároveň TI nebo měřitelná CPU aktivita
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
        // E) Office instalace/aktualizace (C2R + MSI) – robustní
        // --------------------------
        private static bool IsOfficeInstallOrUpdateInProgress()
        {
            // Primární signály
            bool c2rSvcActive = IsServiceActive("ClickToRunSvc"); // Running | StartPending
            bool procPresent = AnyProcessRunning("OfficeClickToRun", "OfficeC2RClient");

            // Nic nenasvědčuje aktivitě? rychle skonči
            if (!c2rSvcActive && !procPresent && !IsAnyOfficeUpdateTaskRunning())
                return false;

            // Lehký CPU sample (jen když je důvod)
            double cpuPct = GetAggregateCpuPercent(new[] { "OfficeClickToRun", "OfficeC2RClient" }, TimeSpan.FromMilliseconds(700));
            bool cpuActive = cpuPct >= 2.0;

            // Registry heuristiky (C2R)
            string updateState = ReadStringHKLM(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "UpdateState");
            int streamingFinished = ReadDwordHKLM(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "StreamingFinished");
            bool regBusy = !string.IsNullOrWhiteSpace(updateState) &&
                           !updateState.Equals("Idle", StringComparison.OrdinalIgnoreCase);

            // Streaming (instalace/aktualizace probíhá)
            bool streaming = procPresent && streamingFinished == 0;

            // Plánovač úloh – Office Automatic Updates
            bool taskRunning = IsAnyOfficeUpdateTaskRunning();

            // MSI fallback (Office MSI)
            bool msiBusy = IsServiceActive("msiserver") && AnyProcessRunning("msiexec");

            // Scoring
            int score = 0;
            if (c2rSvcActive) score++;
            if (procPresent) score++;
            if (cpuActive) score++;
            if (regBusy) score++;
            if (streaming) score++;
            if (taskRunning) score++;
            if (msiBusy) score++;

            bool doingWork = cpuActive || regBusy || streaming || msiBusy;
            return (score >= 2) && doingWork;
        }

        // --------------------------
        // Helpers: služby, procesy, CPU sampling, COM, registry
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

        private static bool AnyProcessRunning(params string[] names)
        {
            var set = new HashSet<string>(names.Select(n => n.ToLowerInvariant()));
            try
            {
                foreach (var p in Process.GetProcesses())
                {
                    try { if (set.Contains(p.ProcessName.ToLowerInvariant())) return true; }
                    catch { /* přístup odepřen / proces skončil */ }
                    finally { try { p.Dispose(); } catch { } }
                }
            }
            catch { }
            return false;
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
                    try { if (names.Contains(p.ProcessName)) sum += p.TotalProcessorTime; }
                    catch { }
                    finally { try { p.Dispose(); } catch { } }
                }
                return sum;
            }
        }

        private static bool TryIsWuInstallerBusy()
        {
            // COM: Microsoft.Update.Session → CreateUpdateInstaller().IsBusy
            try
            {
                var t = Type.GetTypeFromProgID("Microsoft.Update.Session", throwOnError: false);
                if (t == null) return false;

                object? session = Activator.CreateInstance(t);
                if (session == null) return false;

                object? installer = session.GetType().GetMethod("CreateUpdateInstaller")?.Invoke(session, null);
                if (installer == null) { ReleaseCom(session); return false; }

                bool busy = installer.GetType().GetProperty("IsBusy")?.GetValue(installer) is bool b && b;

                ReleaseCom(installer);
                ReleaseCom(session);
                return busy;
            }
            catch { return false; }
        }

        // Task Scheduler COM – detekce běžících Office update úloh
        private static bool IsAnyOfficeUpdateTaskRunning()
        {
            try
            {
                var t = Type.GetTypeFromProgID("Schedule.Service", throwOnError: false);
                if (t == null) return false;

                dynamic svc = Activator.CreateInstance(t)!;
                svc.Connect();

                bool running = false;
                // Standardní složka plánovače
                dynamic folder = svc.GetFolder(@"\Microsoft\Office");
                dynamic tasks = folder.GetTasks(0);

                string[] nameHints = {
                    "Office Automatic Updates", // i "Office Automatic Updates 2.0"
                    "ClickToRun"
                };

                for (int i = 1; i <= (int)tasks.Count; i++)
                {
                    dynamic task = tasks[i];
                    string name = task.Name as string ?? string.Empty;
                    int state = (int)task.State; // TASK_STATE_RUNNING == 4
                    if (state == 4 && nameHints.Any(h => name.IndexOf(h, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        running = true;
                    }
                    ReleaseCom(task);
                    if (running) break;
                }

                ReleaseCom(tasks);
                ReleaseCom(folder);
                ReleaseCom(svc);
                return running;
            }
            catch
            {
                // Task Scheduler nemusí být dostupný; heuristiky níže stačí
                return false;
            }
        }

        private static void ReleaseCom(object o)
        {
            try { Marshal.FinalReleaseComObject(o); } catch { }
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
            catch { return null; }
        }

        private static string ReadStringHKLM(string subKey, string valueName)
        {
            try { using var k = OpenHKLM(subKey); return k?.GetValue(valueName)?.ToString() ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static int ReadDwordHKLM(string subKey, string valueName)
        {
            try
            {
                using var k = OpenHKLM(subKey);
                var v = k?.GetValue(valueName);
                return v is int i ? i : (v is string s && int.TryParse(s, out var j) ? j : 0);
            }
            catch { return 0; }
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

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);
    }
}
