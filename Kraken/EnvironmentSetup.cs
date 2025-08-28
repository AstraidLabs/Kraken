using System;
using System.Collections.Generic;
using System.IO;

namespace Kraken;

/// <summary>
/// Normalizes critical environment variables so the application can reliably
/// reach system utilities regardless of process bitness.
/// </summary>
public static class EnvironmentSetup
{
    /// <summary>
    /// Applies best-effort environment variable settings scoped to the
    /// current process. Failures are swallowed so startup can continue.
    /// </summary>
    public static void Normalize()
    {
        try
        {
            var systemRoot = Environment.GetEnvironmentVariable("SystemRoot");
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                systemRoot = Path.GetDirectoryName(Environment.SystemDirectory) ?? "C\\Windows";
            }

            var system32 = Path.Combine(systemRoot, "System32");
            var sysnative = Path.Combine(systemRoot, "Sysnative");

            // Prefer Sysnative when available so a WOW64 32-bit process can
            // access native 64-bit binaries like reg.exe.
            var preferredSystemDir = system32;
            if (Environment.Is64BitOperatingSystem && Directory.Exists(sysnative) &&
                File.Exists(Path.Combine(sysnative, "reg.exe")))
            {
                preferredSystemDir = sysnative;
            }

            // Construct a PATH starting with the preferred system directories.
            var pathEntries = new List<string>
            {
                preferredSystemDir,
                Path.Combine(preferredSystemDir, "Wbem"),
                Path.Combine(preferredSystemDir, "WindowsPowerShell", "v1.0")
            };

            var existingPath = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrWhiteSpace(existingPath))
            {
                pathEntries.Add(existingPath);
            }

            Environment.SetEnvironmentVariable("PATH", string.Join(Path.PathSeparator, pathEntries));

            // Ensure PATHEXT has standard Windows executable extensions.
            var pathext = Environment.GetEnvironmentVariable("PATHEXT");
            if (string.IsNullOrWhiteSpace(pathext))
            {
                Environment.SetEnvironmentVariable(
                    "PATHEXT",
                    ".COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC");
            }

            // Point ComSpec at the cmd.exe located in the preferred system directory.
            Environment.SetEnvironmentVariable("ComSpec", Path.Combine(preferredSystemDir, "cmd.exe"));

            // Prepend the standard PowerShell module directory.
            var psModules = Path.Combine(preferredSystemDir, "WindowsPowerShell", "v1.0", "Modules");
            var psModulePath = Environment.GetEnvironmentVariable("PSModulePath");
            if (string.IsNullOrWhiteSpace(psModulePath))
            {
                Environment.SetEnvironmentVariable("PSModulePath", psModules);
            }
            else
            {
                Environment.SetEnvironmentVariable("PSModulePath", psModules + Path.PathSeparator + psModulePath);
            }
        }
        catch
        {
            // Suppress all exceptions – environment normalization is best-effort.
        }
    }
}

