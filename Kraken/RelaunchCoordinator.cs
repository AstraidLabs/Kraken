using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Kraken;

/// <summary>
/// Coordinates architecture-aware relaunch logic so the application can
/// access the appropriate system binaries and registry view.
/// </summary>
public static class RelaunchCoordinator
{
    /// <summary>
    /// Determines whether the current process should relaunch through a
    /// different cmd.exe in order to access architecture-specific resources.
    /// </summary>
    /// <param name="args">Command line arguments excluding the executable.</param>
    /// <returns>True if a relaunch was started and the caller should exit.</returns>
    public static bool HandleRelaunch(string[] args)
    {
        try
        {
            bool hasRe1 = args.Any(a => string.Equals(a, "re1", StringComparison.OrdinalIgnoreCase));
            bool hasRe2 = args.Any(a => string.Equals(a, "re2", StringComparison.OrdinalIgnoreCase));

            string systemRoot = Environment.GetEnvironmentVariable("SystemRoot") ?? "C\\Windows";
            string sysnativeCmd = Path.Combine(systemRoot, "Sysnative", "cmd.exe");
            string sysArm32Cmd = Path.Combine(systemRoot, "SysArm32", "cmd.exe");

            // WOW64: 32-bit process on 64-bit OS. Relaunch via Sysnative\cmd.exe.
            if (Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess &&
                File.Exists(sysnativeCmd) && !hasRe1)
            {
                Relaunch(sysnativeCmd, args.Append("re1"));
                return true;
            }

            // ARM64 OS running an AMD64 build. Relaunch via SysArm32\cmd.exe.
            if (RuntimeInformation.OSArchitecture == Architecture.Arm64 &&
                RuntimeInformation.ProcessArchitecture == Architecture.X64 &&
                File.Exists(sysArm32Cmd) && !hasRe2)
            {
                Relaunch(sysArm32Cmd, args.Append("re2"));
                return true;
            }
        }
        catch
        {
            // Suppress errors and continue normally.
        }

        return false;
    }

    private static void Relaunch(string cmdPath, IEnumerable<string> args)
    {
        try
        {
            string? exe = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exe) || !File.Exists(exe))
                return;

            var psi = new ProcessStartInfo
            {
                FileName = cmdPath,
                UseShellExecute = false,
                RedirectStandardInput = false,
                RedirectStandardOutput = false,
                RedirectStandardError = false
            };

            psi.ArgumentList.Add("/c");

            // Quote executable and arguments to survive spaces.
            var quotedArgs = args.Select(Quote);
            var command = Quote(exe);
            if (quotedArgs.Any())
            {
                command += " " + string.Join(" ", quotedArgs);
            }

            psi.ArgumentList.Add(command);

            Process.Start(psi);
        }
        catch
        {
            // Best-effort; ignore failures.
        }
    }

    private static string Quote(string value) =>
        value.Contains(' ') ? $"\"{value}\"" : value;
}

