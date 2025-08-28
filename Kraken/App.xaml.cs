using System.Windows;

namespace Kraken;

/// <summary>
/// Application entry point.
/// Normalizes environment variables and performs architecture-aware
/// relaunching before allowing the main window to display.
/// </summary>
public partial class App : Application
{
    /// <inheritdoc />
    protected override void OnStartup(StartupEventArgs e)
    {
        // Immediately abort startup if Windows is in a state where running
        // the application could interfere with system maintenance.
        StartupGuard.CheckAndExitIfBlocked();

        // Normalize key environment variables so the process can locate
        // expected Windows system binaries regardless of architecture.
        EnvironmentSetup.Normalize();

        // If a relaunch was initiated (to switch architectures), shut down
        // this instance immediately. The new process will continue normally.
        if (RelaunchCoordinator.HandleRelaunch(e.Args))
        {
            Shutdown();
            return;
        }

        base.OnStartup(e);
    }
}

