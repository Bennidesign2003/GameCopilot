using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace GameCopilot.Services;

/// <summary>
/// Migrated from WPF GamesPage.xaml.cs.
/// Contains all game launch logic: SteamVR, Pimax, OpenXR, MSFS start, process monitoring.
/// Cross-platform safe: Windows functions only run on Windows.
/// </summary>
public class GameLaunchService
{
    private readonly AppConfigService _config;

    public event Action<string>? StatusChanged;
    public event Action<string>? ErrorOccurred;

    public GameLaunchService(AppConfigService config)
    {
        _config = config;
    }

    // ======================================================
    // VR START - Migrated from WPF Button_VRStart
    // ======================================================
    public async Task<bool> StartVRAsync(
        Action<string> setButtonText,
        Action<bool> setButtonEnabled,
        Action resetButton)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            ErrorOccurred?.Invoke("VR-Start ist nur unter Windows verfuegbar.");
            return false;
        }

        setButtonEnabled(false);
        setButtonText("VR wird nun gestartet ...");

        // 1. Start Pimax Play (from WPF StartPimaxPlay)
        StartPimaxPlay();

        // 2. Set OpenXR to SteamVR/Pimax (from WPF SetPimaxOpenXR)
        SetPimaxOpenXR();

        try
        {
            // 3. Start SteamVR (from WPF StartSteamVR logic inside Button_VRStart)
            if (File.Exists(_config.SteamVrPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _config.SteamVrPath,
                    UseShellExecute = true
                });
                StatusChanged?.Invoke("SteamVR wird gestartet...");
            }
            else
            {
                ErrorOccurred?.Invoke("SteamVR nicht gefunden! Bitte pruefen Sie den Pfad.");
                resetButton();
                return false;
            }

            await Task.Delay(5000);

            // 4. Set OpenXR Runtime to Pimax (from WPF Registry write)
            try
            {
                SetOpenXRActiveRuntime(_config.PimaxOpenXrJson);
            }
            catch (UnauthorizedAccessException)
            {
                ErrorOccurred?.Invoke("Admin-Rechte noetig, um OpenXR Runtime zu aendern!");
                resetButton();
                return false;
            }

            await Task.Delay(3000);

            // 5. Start MSFS 2024 via Steam (from WPF cmd /c start steam://rungameid/)
            LaunchMsfsSteam();

            setButtonText("VR laeuft gerade ...");
            StatusChanged?.Invoke("MSFS 2024 VR gestartet");

            // 6. Monitor MSFS process (from WPF MonitorMSFSProcess)
            _ = Task.Run(async () =>
            {
                await MonitorMSFSProcess();
                resetButton();
            });

            return true;
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke("Fehler beim Starten von VR: " + ex.Message);
            resetButton();
            return false;
        }
    }

    // ======================================================
    // DESKTOP START - Migrated from WPF DesktopStartButton_Click
    // ======================================================
    public async Task<bool> StartDesktopAsync(
        Action<string> setButtonText,
        Action<bool> setButtonEnabled,
        Action resetButton)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            ErrorOccurred?.Invoke("MSFS-Start ist nur unter Windows verfuegbar.");
            return false;
        }

        setButtonEnabled(false);
        setButtonText("MSFS wird gestartet ...");

        try
        {
            // 1. Check if Steam is running (from WPF)
            bool steamRunning = Process.GetProcessesByName("Steam").Any();

            if (!steamRunning)
            {
                if (File.Exists(_config.SteamExePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = _config.SteamExePath,
                        UseShellExecute = true
                    });
                    setButtonText("Steam wird gestartet ...");
                    StatusChanged?.Invoke("Steam wird gestartet...");
                    await Task.Delay(5000);
                }
                else
                {
                    ErrorOccurred?.Invoke("Steam konnte nicht gefunden werden! Bitte Pfad pruefen.");
                    resetButton();
                    return false;
                }
            }

            await Task.Delay(2000);

            // 2. Start MSFS 2024 Desktop (from WPF)
            LaunchMsfsSteam();

            setButtonText("MSFS laeuft gerade ...");
            StatusChanged?.Invoke("MSFS 2024 Desktop gestartet");

            // 3. Monitor MSFS process (from WPF MonitorMSFSProcessDesktop)
            _ = Task.Run(async () =>
            {
                await MonitorMSFSProcess();
                resetButton();
            });

            return true;
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke("Fehler beim Starten von MSFS 2024: " + ex.Message);
            resetButton();
            return false;
        }
    }

    // ======================================================
    // HELPER METHODS - All migrated from WPF GamesPage.xaml.cs
    // ======================================================

    /// <summary>
    /// Migrated from WPF StartPimaxPlay()
    /// </summary>
    private void StartPimaxPlay()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

        try
        {
            if (!File.Exists(_config.PimaxClientPath))
            {
                StatusChanged?.Invoke("Pimax Play nicht gefunden - wird uebersprungen");
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = _config.PimaxClientPath,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Minimized
            });

            StatusChanged?.Invoke("Pimax Play wird gestartet...");
        }
        catch (Exception ex)
        {
            StatusChanged?.Invoke($"Pimax Play Fehler: {ex.Message}");
        }
    }

    /// <summary>
    /// Migrated from WPF SetPimaxOpenXR()
    /// Sets the OpenXR ActiveRuntime registry key.
    /// </summary>
    private void SetPimaxOpenXR()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

        try
        {
            SetOpenXRActiveRuntime(_config.PimaxOpenXrJson);
            StatusChanged?.Invoke("OpenXR Runtime gesetzt");
        }
        catch (UnauthorizedAccessException)
        {
            ErrorOccurred?.Invoke("Administratorrechte noetig fuer OpenXR Runtime!");
        }
        catch (Exception ex)
        {
            ErrorOccurred?.Invoke("Fehler beim Setzen der OpenXR Runtime: " + ex.Message);
        }
    }

    /// <summary>
    /// Writes the OpenXR ActiveRuntime registry value.
    /// Migrated from WPF Registry writes in Button_VRStart and SetPimaxOpenXR.
    /// </summary>
    private static void SetOpenXRActiveRuntime(string runtimePath)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

        // Uses Microsoft.Win32.Registry via reflection to stay cross-platform compilable
        var registryType = Type.GetType("Microsoft.Win32.Registry, Microsoft.Win32.Registry");
        if (registryType == null) return;

        var localMachine = registryType.GetField("LocalMachine")?.GetValue(null);
        if (localMachine == null) return;

        var createSubKeyMethod = localMachine.GetType().GetMethod("CreateSubKey", new[] { typeof(string) });
        var key = createSubKeyMethod?.Invoke(localMachine, new object[] { @"SOFTWARE\Khronos\OpenXR\1\InstalledRuntimes" });
        if (key == null) return;

        var setValueMethod = key.GetType().GetMethod("SetValue", new[] { typeof(string), typeof(object), Type.GetType("Microsoft.Win32.RegistryValueKind, Microsoft.Win32.Registry")! });
        // RegistryValueKind.String = 1
        setValueMethod?.Invoke(key, new object[] { "ActiveRuntime", runtimePath, 1 });

        var closeMethod = key.GetType().GetMethod("Close");
        closeMethod?.Invoke(key, null);
    }

    /// <summary>
    /// Launches MSFS 2024 via Steam protocol.
    /// Migrated from WPF: cmd /c start steam://rungameid/{appId}
    /// </summary>
    private void LaunchMsfsSteam()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

        var startInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c start steam://rungameid/{_config.MsfsAppId}",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        Process.Start(startInfo);
    }

    /// <summary>
    /// Monitors the MSFS process and calls the callback when it exits.
    /// Migrated from WPF MonitorMSFSProcess() and MonitorMSFSProcessDesktop().
    /// Both were identical, so merged into one method.
    /// </summary>
    private async Task MonitorMSFSProcess()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

        try
        {
            // Wait for MSFS 2024 to start (max 60 seconds) - from WPF
            Process? msfsProcess = null;
            for (int i = 0; i < 60; i++)
            {
                msfsProcess = Process.GetProcessesByName("FlightSimulator").FirstOrDefault()
                    ?? Process.GetProcessesByName("FlightSimulator2024").FirstOrDefault();

                if (msfsProcess != null)
                    break;

                await Task.Delay(1000);
            }

            if (msfsProcess == null)
            {
                // MSFS not found after 60s, wait another 60s then give up - from WPF
                await Task.Delay(60000);
                return;
            }

            // Wait for MSFS to exit - from WPF
            await msfsProcess.WaitForExitAsync();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Ueberwachen des MSFS-Prozesses: {ex.Message}");
        }
    }

    /// <summary>
    /// Checks if SteamVR is installed.
    /// Migrated from WPF SimulationEnvironmentService.CheckSteamVR()
    /// </summary>
    public bool IsSteamVRInstalled()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return false;
        return File.Exists(_config.SteamVrPath);
    }

    /// <summary>
    /// Checks if MSFS 2024 is installed.
    /// Migrated from WPF SplashScreenPage.CheckMSFSInstallation()
    /// </summary>
    public bool IsMsfsInstalled()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return false;

        string steamPath = @"C:\Program Files (x86)\Steam\steamapps\common\MicrosoftFlightSimulator24";
        string msStorePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "WindowsApps", "Microsoft.FlightSimulator");

        return Directory.Exists(steamPath) || Directory.Exists(msStorePath);
    }
}
