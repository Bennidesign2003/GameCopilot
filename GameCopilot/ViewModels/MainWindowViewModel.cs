using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GameCopilot.Models;
using GameCopilot.Services;

namespace GameCopilot.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly AppConfigService _configService;
    private readonly GameLaunchService _gameLaunchService;
    private readonly ModService _modService;
    private readonly PresetService _presetService;
    private readonly ReShadeService _reshadeService;
    private readonly OllamaService _ollamaService;
    private readonly CodexService _codexService;
    private readonly McpClientService _mcpService;
    private readonly UpdateService _updateService;
    private readonly ChatHistoryService _historyService = new();

    private ReShadePreset? _originalPresetState;

    // ── Generation cancellation ──────────────────────────────────────────────
    private CancellationTokenSource? _generationCts;
    [ObservableProperty] private bool _isGenerating;

    // ── Debounced history save ────────────────────────────────────────────────
    private Timer? _historySaveTimer;

    // ── System prompt cache ───────────────────────────────────────────────────
    private string? _cachedSystemPrompt;
    private bool    _cachedMcpState;
    private string  _cachedGpuName = "";

    // ── Ollama live connection status ─────────────────────────────────────────
    /// <summary>"connected" | "disconnected" | "checking"</summary>
    [ObservableProperty] private string _ollamaConnStatus = "checking";
    [ObservableProperty] private string _ollamaConnModelName = "";
    private Timer? _ollamaStatusTimer;

    /// <summary>Dot fill colour bound directly in AXAML — no converter needed.</summary>
    public string OllamaConnDotColor => OllamaConnStatus switch
    {
        "connected"    => "#4ade80",
        "disconnected" => "#ff4444",
        _              => "#fbbf24"   // checking = yellow
    };

    /// <summary>Display text in the status bar, e.g. "● qwen3:8b" or "Offline".</summary>
    public string OllamaConnDisplayText => OllamaConnStatus switch
    {
        "connected"    => string.IsNullOrEmpty(OllamaConnModelName)
                              ? "Ollama verbunden"
                              : OllamaConnModelName,
        "disconnected" => "Ollama offline",
        _              => "Überprüfe..."
    };

    // ── Codex API key show/hide & save confirmation ──────────────────────────
    [ObservableProperty] private bool _showApiKeyText;
    [ObservableProperty] private bool _codexKeySaved;

    // ── Settings page extras ─────────────────────────────────────────────────
    [ObservableProperty] private string _settingsOllamaUrl = "http://localhost:11434";
    [ObservableProperty] private bool   _settingsWarmOnStartup = true;
    [ObservableProperty] private bool   _settingsLoadChatHistory = true;
    [ObservableProperty] private string _mcpVersionDisplay = "";
    [ObservableProperty] private string _mcpServerPath = "";

    // Detected GPU name (read once at startup via WMI, used in system prompt)
    private string _detectedGpuName = "";

    // -- Navigation --
    [ObservableProperty] private string _currentPage = "Splash";
    [ObservableProperty] private bool _sidebarVisible;
    [ObservableProperty] private string _activeNav = "Launch";
    [ObservableProperty] private bool _isLaunchActive = true;
    [ObservableProperty] private bool _isModsActive;
    [ObservableProperty] private bool _isReshadeActive;
    [ObservableProperty] private bool _isSettingsActive;
    [ObservableProperty] private bool _isUpdatesActive;
    [ObservableProperty] private bool _isAiChatActive;

    private void UpdateNavStates(string active)
    {
        IsLaunchActive = active == "Launch";
        IsModsActive = active == "Mods";
        IsReshadeActive = active == "ReShade";
        IsSettingsActive = active == "Settings";
        IsUpdatesActive = active == "Updates";
        IsAiChatActive = active == "AiChat";
    }

    // -- Splash Screen --
    [ObservableProperty] private string _splashStatus = "Initialisiere System...";
    [ObservableProperty] private double _splashProgress;
    [ObservableProperty] private bool _splashDone;
    [ObservableProperty] private bool _msfsInstalled;
    [ObservableProperty] private bool _steamVrInstalled;

    // -- Games Page --
    [ObservableProperty] private string _vrButtonText = "In VR starten";
    [ObservableProperty] private bool _vrButtonEnabled = true;
    [ObservableProperty] private string _vrStatusText = "VR-Headset Status wird geprueft...";
    [ObservableProperty] private string _vrStatusColor = "#B0B0B0";
    [ObservableProperty] private string _desktopButtonText = "Normal starten";
    [ObservableProperty] private bool _desktopButtonEnabled = true;
    [ObservableProperty] private string _vrErrorText = "";

    // -- Mods Page --
    public ObservableCollection<ModItem> Mods { get; } = new();
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private string _modCategoryFilter = "All";
    public ObservableCollection<ModItem> FilteredMods { get; } = new();
    public ObservableCollection<object> ModGridItems { get; } = new();
    [ObservableProperty] private int _modCount;
    [ObservableProperty] private string _totalModSize = "0 B";
    [ObservableProperty] private bool _deleteModalVisible;
    [ObservableProperty] private string _deleteModName = "";
    private ModItem? _pendingDeleteMod;

    partial void OnSearchTextChanged(string value) => ApplyModFilter();
    partial void OnModCategoryFilterChanged(string value) => ApplyModFilter();

    [RelayCommand]
    private void SetModCategory(string category)
    {
        ModCategoryFilter = category;
    }

    // -- ReShade Presets --
    public ObservableCollection<ReShadePreset> Presets { get; } = new();
    [ObservableProperty] private ReShadePreset? _selectedPreset;

    partial void OnSelectedPresetChanged(ReShadePreset? value)
    {
        if (value != null)
        {
            _originalPresetState = value.Clone();
            SharpenEnabled = value.SharpenEnabled;
            BloomEnabled = value.BloomEnabled;
            VibranceEnabled = value.VibranceEnabled;
            TonemapEnabled = value.TonemapEnabled;
            SharpenStrength = value.SharpenStrength;
            BloomStrength = value.BloomStrength;
            VibranceStrength = value.VibranceStrength;
            Contrast = value.Contrast;
            Brightness = value.Brightness;
        }
        OnPropertyChanged(nameof(HasSelection));
    }

    public bool HasSelection => SelectedPreset != null;

    // -- Shader toggles --
    [ObservableProperty] private bool _sharpenEnabled;
    [ObservableProperty] private bool _bloomEnabled;
    [ObservableProperty] private bool _vibranceEnabled;
    [ObservableProperty] private bool _tonemapEnabled;

    // -- Shader sliders --
    [ObservableProperty] private double _sharpenStrength = 0.5;
    [ObservableProperty] private double _bloomStrength = 0.3;
    [ObservableProperty] private double _vibranceStrength = 0.5;
    [ObservableProperty] private double _contrast = 0.5;
    [ObservableProperty] private double _brightness = 0.5;

    // -- Detail panel visibility --
    [ObservableProperty] private bool _showShaderDetail;

    // -- Status bar --
    [ObservableProperty] private string _msfsPath = "Nicht erkannt";
    [ObservableProperty] private string _reshadePath = "Nicht erkannt";
    [ObservableProperty] private bool _isReshadeFound;
    [ObservableProperty] private string _activePresetName = "Kein Preset aktiv";
    [ObservableProperty] private string _ollamaStatus = "Nicht verbunden";
    [ObservableProperty] private string _statusMessage = "Bereit";

    // -- Settings Page --
    [ObservableProperty] private string _settingsCommunityPath = "";
    [ObservableProperty] private string _settingsMsfsGamePath = "";
    [ObservableProperty] private string _settingsSteamVrPath = "";
    [ObservableProperty] private string _settingsSteamExePath = "";
    [ObservableProperty] private string _settingsPimaxClientPath = "";
    [ObservableProperty] private string _settingsMsfsAppId = "";

    // -- Updates Page --
    [ObservableProperty] private string _currentVersionDisplay = "v1.0.0";
    [ObservableProperty] private string _updateStatusText = "CHECKING...";
    [ObservableProperty] private bool _isCheckingUpdates;
    [ObservableProperty] private string _lastUpdateCheck = "—";
    [ObservableProperty] private bool _isUpdateAvailable;
    [ObservableProperty] private string _latestVersionDisplay = "";
    [ObservableProperty] private double _downloadProgress;
    [ObservableProperty] private bool _isDownloading;
    [ObservableProperty] private string _downloadStatusText = "";
    public ObservableCollection<ReleaseEntry> Releases { get; } = new();
    /// <summary>True when no releases have been loaded yet — used for the empty-state panel.</summary>
    public bool NoReleasesYet => Releases.Count == 0;

    // -- Toast --
    [ObservableProperty] private bool _toastVisible;
    [ObservableProperty] private string _toastTitle = "";
    [ObservableProperty] private string _toastMessage = "";

    // -- Window title (shows version) --
    [ObservableProperty] private string _windowTitle = "Game Copilot";

    // -- MCP disconnect banner --
    [ObservableProperty] private bool _mcpDisconnected;

    // -- Chat empty state --
    public bool NoMessages => ChatMessages.Count == 0;

    public MainWindowViewModel()
    {
        _configService = new AppConfigService();
        _configService.Load();

        _gameLaunchService = new GameLaunchService(_configService);
        _modService = new ModService(_configService);
        _presetService = new PresetService();
        _reshadeService = new ReShadeService(_configService);
        _ollamaService = new OllamaService();
        _ollamaService.BaseUrl = _configService.OllamaUrl;
        _codexService = new CodexService();
        _mcpService = new McpClientService();
        _updateService = new UpdateService();

        // Wire up GameLaunchService events
        _gameLaunchService.StatusChanged += msg =>
            Dispatcher.UIThread.Post(() => StatusMessage = msg);
        _gameLaunchService.ErrorOccurred += msg =>
            Dispatcher.UIThread.Post(() => VrErrorText = msg);

        // Load settings into UI
        SettingsCommunityPath = _configService.CommunityPath;
        SettingsMsfsGamePath = _configService.MsfsGamePath;
        SettingsSteamVrPath = _configService.SteamVrPath;
        SettingsSteamExePath = _configService.SteamExePath;
        SettingsPimaxClientPath = _configService.PimaxClientPath;
        SettingsMsfsAppId = _configService.MsfsAppId;

        CurrentVersionDisplay = $"v{_configService.CurrentVersion}";
        WindowTitle = $"Game Copilot v{_configService.CurrentVersion}";

        // Keep NoReleasesYet in sync so the compiled binding refreshes
        Releases.CollectionChanged += (_, _) => OnPropertyChanged(nameof(NoReleasesYet));

        // Keep NoMessages in sync for empty state placeholder
        ChatMessages.CollectionChanged += (_, _) => OnPropertyChanged(nameof(NoMessages));

        // ── Load Codex / provider settings from config ────────────────────────
        _codexApiKey = _configService.CodexApiKey;
        _selectedCodexModel = _configService.CodexModel;
        _chatProvider = _configService.ChatProvider;
        _isCodexSelected = _chatProvider == "codex";

        // ── Load new settings fields ──────────────────────────────────────────
        _settingsOllamaUrl       = _configService.OllamaUrl;
        _settingsWarmOnStartup   = _configService.WarmOnStartup;
        _settingsLoadChatHistory = _configService.LoadChatHistory;

        // ── MCP server path display ───────────────────────────────────────────
        _mcpServerPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GameCopilot", "mcp-server.py");

        // ── Start Ollama connection-status polling (every 30 s) ───────────────
        _ollamaStatusTimer = new Timer(
            _ => _ = PollOllamaStatusAsync(),
            null,
            TimeSpan.FromSeconds(5),      // first check after 5 s
            TimeSpan.FromSeconds(30));    // then every 30 s

        // Populate Codex model list (cloud models are always "installed")
        foreach (var (id, label, size, desc) in CodexService.SupportedModels)
        {
            CodexModelList.Add(new ChatModelItem
            {
                Id = id,
                Label = label,
                Size = size,
                Description = desc,
                IsInstalled = true,
            });
        }
        _selectedCodexModelItem = CodexModelList.FirstOrDefault(m => m.Id == _selectedCodexModel)
                                  ?? CodexModelList.FirstOrDefault();
    }

    // ======================================================
    // OLLAMA LIVE STATUS POLLING
    // ======================================================

    private async Task PollOllamaStatusAsync()
    {
        try
        {
            Dispatcher.UIThread.Post(() =>
            {
                OllamaConnStatus = "checking";
                OnPropertyChanged(nameof(OllamaConnDotColor));
                OnPropertyChanged(nameof(OllamaConnDisplayText));
            });
            var connected = await _ollamaService.CheckConnectionAsync().ConfigureAwait(false);
            var modelName = connected && SelectedChatModel != null ? SelectedChatModel.Id : "";
            Dispatcher.UIThread.Post(() =>
            {
                OllamaConnStatus    = connected ? "connected" : "disconnected";
                OllamaConnModelName = modelName;
                OnPropertyChanged(nameof(OllamaConnDotColor));
                OnPropertyChanged(nameof(OllamaConnDisplayText));
            });
        }
        catch
        {
            Dispatcher.UIThread.Post(() =>
            {
                OllamaConnStatus    = "disconnected";
                OllamaConnModelName = "";
                OnPropertyChanged(nameof(OllamaConnDotColor));
                OnPropertyChanged(nameof(OllamaConnDisplayText));
            });
        }
    }

    // ======================================================
    // GENERATION CANCELLATION
    // ======================================================

    [RelayCommand]
    private void CancelGeneration()
    {
        _generationCts?.Cancel();
    }

    // ======================================================
    // CHAT HISTORY HELPERS
    // ======================================================

    private void ScheduleHistorySave()
    {
        // Debounce: reset the timer each time a message is added (max once per 2 s).
        // The timer callback runs on a thread-pool thread, so we must snapshot
        // ChatMessages on the UI thread before handing off to Task.Run.
        _historySaveTimer?.Dispose();
        _historySaveTimer = new Timer(
            _ => Dispatcher.UIThread.Post(() =>
            {
                var snapshot = ChatMessages.ToList();
                Task.Run(() => _historyService.Save(snapshot));
            }),
            null,
            TimeSpan.FromSeconds(2),
            Timeout.InfiniteTimeSpan);
    }

    // ======================================================
    // CODEX API KEY UX
    // ======================================================

    [RelayCommand]
    private void ToggleApiKeyVisibility() => ShowApiKeyText = !ShowApiKeyText;

    [RelayCommand]
    private async Task SaveCodexKeyWithConfirmation()
    {
        // The key is already auto-saved via OnCodexApiKeyChanged.
        // This command just shows the inline "✓ Gespeichert" confirmation.
        _configService.CodexApiKey = CodexApiKey;
        _configService.Save();
        CodexKeySaved = true;
        await Task.Delay(2500);
        CodexKeySaved = false;
    }

    // ======================================================
    // RETRY LAST MESSAGE
    // ======================================================

    [RelayCommand]
    private void RetryLastMessage()
    {
        // Cancel any in-flight generation first
        _generationCts?.Cancel();

        // Find the last user message
        var lastUser = ChatMessages.LastOrDefault(m => m.IsUser);
        if (lastUser == null) return;

        var text = lastUser.Content ?? "";

        // Remove everything from that message onward
        var idx = ChatMessages.IndexOf(lastUser);
        while (ChatMessages.Count > idx)
            ChatMessages.RemoveAt(idx);

        ChatInput = text;
        SendChatMessageCommand.Execute(null);
    }

    // ======================================================
    // SETTINGS PAGE COMMANDS
    // ======================================================

    [RelayCommand]
    private void SaveAppSettings()
    {
        _configService.CommunityPath    = SettingsCommunityPath;
        _configService.MsfsGamePath     = SettingsMsfsGamePath;
        _configService.SteamVrPath      = SettingsSteamVrPath;
        _configService.SteamExePath     = SettingsSteamExePath;
        _configService.PimaxClientPath  = SettingsPimaxClientPath;
        _configService.MsfsAppId        = SettingsMsfsAppId;
        _configService.OllamaUrl        = SettingsOllamaUrl;
        _ollamaService.BaseUrl          = SettingsOllamaUrl;
        _configService.WarmOnStartup    = SettingsWarmOnStartup;
        _configService.LoadChatHistory  = SettingsLoadChatHistory;
        _configService.Save();
        ShowToast("Einstellungen gespeichert", "Alle Einstellungen wurden gespeichert.");

        // Reload mods / paths
        LoadAllMods();
        _reshadeService.DetectPaths();
        DetectEnvironment();
        LoadPresets();
    }

    [RelayCommand]
    private async Task TestOllamaConnection()
    {
        OllamaConnStatus = "checking";
        await PollOllamaStatusAsync();
        ShowToast(
            OllamaConnStatus == "connected" ? "Ollama verbunden" : "Ollama nicht erreichbar",
            OllamaConnStatus == "connected"
                ? $"Verbunden mit {SettingsOllamaUrl}"
                : $"Kein Ollama unter {SettingsOllamaUrl} gefunden.");
    }

    // ======================================================
    // SPLASH SEQUENCE - Migrated from WPF SplashScreenPage.cs
    // PerformRealLoading() with actual system checks
    // ======================================================
    public async Task RunSplashSequenceAsync()
    {
        // Step 1: Init system (from WPF InitializeSystem)
        SplashStatus = "Initialisiere System...";
        SplashProgress = 1.0 / 7;
        await Task.Delay(500);

        // Step 2: Load config (from WPF LoadConfiguration)
        SplashStatus = "Lade Konfiguration...";
        SplashProgress = 2.0 / 7;
        _configService.Load();
        await Task.Delay(500);

        // Step 3: Check SteamVR (from WPF SimulationEnvironmentService.CheckSteamVR)
        SplashStatus = "Pruefe SteamVR Installation...";
        SplashProgress = 3.0 / 7;
        SteamVrInstalled = _gameLaunchService.IsSteamVRInstalled();
        await Task.Delay(500);

        // Step 3b: Detect GPU name via Registry for system prompt context
        try
        {
            // GPU name stored in Display adapter registry class key
            await Task.Run(() =>
            {
                try
                {
                    const string gpuClassKey =
                        @"SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}";
                    using var classKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(gpuClassKey);
                    if (classKey != null)
                    {
                        foreach (var subName in classKey.GetSubKeyNames())
                        {
                            if (subName == "Properties") continue;
                            using var sub = classKey.OpenSubKey(subName);
                            var desc = sub?.GetValue("DriverDesc")?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(desc) &&
                                (desc.Contains("NVIDIA", StringComparison.OrdinalIgnoreCase) ||
                                 desc.Contains("Radeon", StringComparison.OrdinalIgnoreCase) ||
                                 desc.Contains("Intel", StringComparison.OrdinalIgnoreCase)))
                            {
                                _detectedGpuName = desc;
                                break;
                            }
                        }
                    }
                }
                catch { /* Registry read is best-effort */ }
            });
        }
        catch { /* GPU detection is best-effort */ }

        // Step 4: Check VR Headset (from WPF SimulationEnvironmentService.CheckVRHeadset)
        SplashStatus = "Pruefe VR-Headset...";
        SplashProgress = 4.0 / 7;
        await Task.Delay(500);

        // Update VR status based on checks
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            VrStatusText = "VR nur unter Windows verfuegbar";
            VrStatusColor = "#FF9500";
        }
        else if (SteamVrInstalled)
        {
            VrStatusText = "SteamVR gefunden";
            VrStatusColor = "#00FF88";
        }
        else
        {
            VrStatusText = "SteamVR nicht gefunden";
            VrStatusColor = "#FF4444";
        }

        // Step 5: Check MSFS (from WPF CheckMSFSInstallation)
        SplashStatus = "Pruefe MSFS Installation...";
        SplashProgress = 5.0 / 7;
        MsfsInstalled = _gameLaunchService.IsMsfsInstalled();
        await Task.Delay(500);

        // Step 6: Load mods (from WPF LoadMods)
        SplashStatus = "Lade Mods...";
        SplashProgress = 6.0 / 7;
        LoadAllMods();
        await Task.Delay(400);

        // Step 7: Prepare UI
        SplashStatus = "Bereite Benutzeroberflaeche vor...";
        SplashProgress = 7.0 / 8;
        LoadPresets();
        DetectEnvironment();
        await Task.Delay(300);

        // Step 8: Pre-warm Nvidia MCP Server so it's ready before first chat message
        SplashStatus = "Nvidia MCP Server wird gestartet...";
        SplashProgress = 1.0;
        _ = Task.Run(async () =>
        {
            try
            {
                await _mcpService.StartAsync(progress =>
                    Dispatcher.UIThread.Post(() => McpStatus = progress));
            }
            catch
            {
                // Non-fatal: MCP will lazy-start on first chat message if pre-warm fails
            }
        });
        await Task.Delay(400);

        SplashStatus = "Bereit!";
        await Task.Delay(400);

        SplashDone = true;
        SidebarVisible = true;
        CurrentPage = "Launch";
        ActiveNav = "Launch";
        UpdateNavStates("Launch");

        // ── Restore previous chat session ─────────────────────────────────────
        if (_configService.LoadChatHistory)
        {
            var history = _historyService.Load();
            if (history.Count > 0)
            {
                ChatMessages.Add(new ChatMessage
                {
                    Role      = "assistant",
                    Content   = "── Letzte Sitzung ──",
                    Timestamp = ""
                });
                foreach (var m in history)
                    ChatMessages.Add(m);
            }
        }

        // ── Read MCP version for Settings display ─────────────────────────────
        try
        {
            var mcpSrc = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "GameCopilot", "mcp-server.py");
            if (!System.IO.File.Exists(mcpSrc))
                mcpSrc = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "mcp-server.py");
            if (System.IO.File.Exists(mcpSrc))
            {
                var firstLine = System.IO.File.ReadLines(mcpSrc).FirstOrDefault() ?? "";
                var vMatch = System.Text.RegularExpressions.Regex.Match(firstLine, @"""([^""]+)""");
                McpVersionDisplay = vMatch.Success ? $"v{vMatch.Groups[1].Value}" : "—";
            }
        }
        catch { McpVersionDisplay = "—"; }

        // Fire-and-forget: check for MCP server update and announce in AI chat if found
        _ = CheckAndAnnounceMcpUpdateAsync();

        // Fire-and-forget: check for updates and announce in AI chat if found
        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(3000); // Wait for UI to settle
                var hasUpdate = await _updateService.CheckForUpdatesAsync(_configService.CurrentVersion);
                if (!hasUpdate) return;

                var releases = await _updateService.FetchReleasesAsync();
                var latest = releases.FirstOrDefault();
                var changelog = latest?.Changes?.Take(5).ToList() ?? new List<string>();
                var changelogText = changelog.Count > 0
                    ? string.Join("\n", changelog.Select(c => $"• {c}"))
                    : "Verbesserungen und Bugfixes.";

                var announcement =
                    $"🔔 **Update verfügbar: v{_updateService.LatestVersion}**\n\n" +
                    $"Aktuelle Version: v{_configService.CurrentVersion}\n\n" +
                    $"**Was ist neu:**\n{changelogText}\n\n" +
                    $"Soll ich das Update jetzt herunterladen und installieren?";

                Dispatcher.UIThread.Post(() =>
                {
                    ChatMessages.Add(new ChatMessage
                    {
                        Role = "assistant",
                        Content = announcement,
                        Timestamp = DateTime.Now.ToString("HH:mm")
                    });
                    IsUpdateAvailable = true;
                    LatestVersionDisplay = $"v{_updateService.LatestVersion}";
                });
            }
            catch
            {
                // Non-fatal: update check failure should not crash or spam the user
            }
        });
    }

    // ======================================================
    // MCP UPDATE CHECK
    // ======================================================

    /// <summary>
    /// Checks if an MCP server update is available and posts a special chat message with Yes/No buttons.
    /// Works with both Codex and local Ollama providers.
    /// </summary>
    private async Task CheckAndAnnounceMcpUpdateAsync()
    {
        try
        {
            // Wait for the MCP server to be ready and for the UI to settle after splash
            await Task.Delay(4000);

            if (!_mcpService.IsRunning) return;

            var result = await _mcpService.CallToolAsync("check_for_mcp_server_update", "{}");

            // Guard against error responses
            if (string.IsNullOrWhiteSpace(result) || result.Contains("\"error\"")) return;

            // Parse the result JSON
            bool updateAvailable = false;
            string githubVersion = "";

            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(result);
                var root = doc.RootElement;

                // The tool may return "update_verfügbar" (German) or "update_available" (English)
                if (root.TryGetProperty("update_verfügbar", out var uvProp))
                    updateAvailable = uvProp.ValueKind == System.Text.Json.JsonValueKind.True;
                else if (root.TryGetProperty("update_available", out var uaProp))
                    updateAvailable = uaProp.ValueKind == System.Text.Json.JsonValueKind.True;

                if (root.TryGetProperty("github_version", out var gvProp))
                    githubVersion = gvProp.GetString() ?? "";
            }
            catch
            {
                return; // Malformed JSON — skip silently
            }

            if (!updateAvailable || string.IsNullOrEmpty(githubVersion)) return;

            var msg = new ChatMessage
            {
                Role = "assistant",
                IsUpdatePrompt = true,
                UpdateVersion = githubVersion,
                Content = $"👋 Hey Benjamin! MCP Server Update **{githubVersion}** ist verfügbar.\nMöchtest du jetzt aktualisieren?",
                Timestamp = DateTime.Now.ToString("HH:mm")
            };

            Dispatcher.UIThread.Post(() => ChatMessages.Add(msg));
        }
        catch
        {
            // Non-fatal: MCP update check failure should not crash or spam the user
        }
    }

    [RelayCommand]
    private async Task ConfirmMcpUpdate(ChatMessage msg)
    {
        msg.UpdateDismissed = true;

        // Show a "working..." message while the update runs
        var typingMsg = new ChatMessage
        {
            Role = "assistant",
            Content = "⏳ MCP Server wird aktualisiert...",
            Timestamp = DateTime.Now.ToString("HH:mm"),
            IsStreaming = true
        };
        Dispatcher.UIThread.Post(() => ChatMessages.Add(typingMsg));

        try
        {
            var result = await _mcpService.CallToolAsync("update_mcp_server", "{\"confirmed\": true}");

            typingMsg.IsStreaming = false;

            // Check if a restart is required
            bool restartRequired = false;
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(result);
                if (doc.RootElement.TryGetProperty("neustart_erforderlich", out var nr))
                    restartRequired = nr.ValueKind == System.Text.Json.JsonValueKind.True;
            }
            catch { /* ignore parse failures */ }

            if (restartRequired)
            {
                typingMsg.Content = "✅ MCP Server wurde aktualisiert!\n\n🔄 Bitte sage \"MCP Server neustarten\" damit die neue Version aktiv wird.";
            }
            else
            {
                typingMsg.Content = "✅ MCP Server wurde erfolgreich aktualisiert! Die neue Version ist jetzt aktiv.";
            }
        }
        catch (Exception ex)
        {
            typingMsg.IsStreaming = false;
            typingMsg.Content = $"❌ Update fehlgeschlagen: {ex.Message}";
        }
    }

    [RelayCommand]
    private void DismissMcpUpdate(ChatMessage msg)
    {
        msg.UpdateDismissed = true;
        Dispatcher.UIThread.Post(() =>
        {
            ChatMessages.Add(new ChatMessage
            {
                Role = "assistant",
                Content = "Ok, kein Update. Du kannst das Update jederzeit über den AI-Chat starten.",
                Timestamp = DateTime.Now.ToString("HH:mm")
            });
        });
    }

    // ======================================================
    // NAVIGATION
    // ======================================================
    [RelayCommand]
    private void NavigateLaunch()
    {
        CurrentPage = "Launch"; ActiveNav = "Launch"; UpdateNavStates("Launch");
    }

    [RelayCommand]
    private void NavigateMods()
    {
        CurrentPage = "Mods"; ActiveNav = "Mods"; UpdateNavStates("Mods");
    }

    [RelayCommand]
    private void NavigateReshade()
    {
        CurrentPage = "ReShade"; ActiveNav = "ReShade"; UpdateNavStates("ReShade");
    }

    [RelayCommand]
    private void NavigateSettings()
    {
        CurrentPage = "Settings"; ActiveNav = "Settings"; UpdateNavStates("Settings");
    }

    [RelayCommand]
    private void NavigateUpdates()
    {
        CurrentPage = "Updates"; ActiveNav = "Updates"; UpdateNavStates("Updates");
    }

    [RelayCommand]
    private async Task NavigateAiChat()
    {
        CurrentPage = "AiChat"; ActiveNav = "AiChat"; UpdateNavStates("AiChat");
        // Skip Ollama init when using cloud provider — only init when local is selected
        if (ChatProvider != "codex" && (ChatModelList.Count == 0 || !_ollamaService.IsAvailable))
            await InitializeOllamaChat();
    }

    // ======================================================
    // AI CHAT PAGE - Ollama Local AI Integration
    // ======================================================

    public ObservableCollection<ChatMessage> ChatMessages { get; } = new();
    [ObservableProperty] private string _chatInput = "";
    [ObservableProperty] private bool _isChatLoading;
    [ObservableProperty] private string _chatStreamingText = "";
    [ObservableProperty] private string _thinkingStatusText = "Analysiere...";
    [ObservableProperty] private int _thinkingDotCount;
    [ObservableProperty] private bool _isToolRunning;

    // Model selection & install state
    public ObservableCollection<ChatModelItem> ChatModelList { get; } = new();
    [ObservableProperty] private ChatModelItem? _selectedChatModel;
    [ObservableProperty] private string _chatModelStatus = "";
    [ObservableProperty] private bool _isModelReady;
    [ObservableProperty] private bool _isPullingModel;
    [ObservableProperty] private string _pullProgressText = "";
    [ObservableProperty] private bool _isInstallingOllama;
    [ObservableProperty] private string _mcpStatus = "";
    [ObservableProperty] private bool _showModelPicker = true;
    [ObservableProperty] private bool _isOllamaOffline;

    // ── Codex / provider switching ────────────────────────────────────────────
    /// <summary>"local" = Ollama; "codex" = OpenAI cloud.</summary>
    [ObservableProperty] private string _chatProvider = "local";
    [ObservableProperty] private bool _isCodexSelected;
    [ObservableProperty] private string _codexApiKey = "";
    [ObservableProperty] private string _selectedCodexModel = "codex-mini-latest";
    [ObservableProperty] private bool _codexApiKeyMissing;
    [ObservableProperty] private ChatModelItem? _selectedCodexModelItem;

    public ObservableCollection<ChatModelItem> CodexModelList { get; } = new();

    partial void OnChatProviderChanged(string value)
    {
        IsCodexSelected = value == "codex";
        CodexApiKeyMissing = false; // reset warning on switch
        _configService.ChatProvider = value;
        _configService.Save();
    }

    partial void OnCodexApiKeyChanged(string value)
    {
        if (!string.IsNullOrEmpty(value))
            CodexApiKeyMissing = false;
        _configService.CodexApiKey = value;
        _configService.Save();
    }

    partial void OnSelectedCodexModelChanged(string value)
    {
        _configService.CodexModel = value;
        _configService.Save();
        // Keep SelectedCodexModelItem in sync
        var item = CodexModelList.FirstOrDefault(m => m.Id == value);
        if (item != null && SelectedCodexModelItem != item)
            SelectedCodexModelItem = item;
    }

    partial void OnSelectedCodexModelItemChanged(ChatModelItem? value)
    {
        if (value != null && SelectedCodexModel != value.Id)
            SelectedCodexModel = value.Id;
    }

    [RelayCommand]
    private void SetProviderLocal()
    {
        ChatProvider = "local";
    }

    [RelayCommand]
    private void SetProviderCodex()
    {
        ChatProvider = "codex";
    }

    [RelayCommand]
    private async Task RetryOllamaConnection()
    {
        IsOllamaOffline = false;
        await InitializeOllamaChat();
    }

    /// <summary>
    /// Background retry when Ollama was slow to start.
    /// Checks every 5s for up to 60s, then refreshes model list + starts MCP.
    /// </summary>
    private async Task RetryConnectionInBackgroundAsync()
    {
        for (int i = 0; i < 12; i++)
        {
            await Task.Delay(5000);
            // Stop retrying if user manually reconnected or we're already online
            if (_ollamaService.IsAvailable) return;

            if (await _ollamaService.CheckConnectionAsync())
            {
                Dispatcher.UIThread.Post(() =>
                {
                    IsOllamaOffline = false;
                    ChatModelStatus = "Ollama verbunden";
                    foreach (var m in ChatModelList)
                        m.IsInstalled = _ollamaService.HasModel(m.Id);
                    SelectedChatModel ??= ChatModelList.FirstOrDefault(m => m.IsInstalled);
                    ShowModelPicker = SelectedChatModel == null;
                    UpdateModelStatus();
                });

                // Start MCP now that Ollama is up
                try
                {
                    await _mcpService.StartAsync(progress =>
                        Dispatcher.UIThread.Post(() => McpStatus = progress));
                }
                catch { /* non-fatal — MCP will retry on next chat message */ }
                Dispatcher.UIThread.Post(() =>
                    McpStatus = _mcpService.IsRunning
                        ? $"MCP: {_mcpService.Tools.Count} Tools bereit"
                        : _mcpService.Status);
                return;
            }
        }
    }

    /// <summary>
    /// Initialize Ollama when navigating to AI Chat: install if needed, start server, check models.
    /// </summary>
    [RelayCommand]
    private async Task InitializeOllamaChat()
    {
        ChatModelStatus = "Ollama wird geprueft...";
        IsModelReady = false;
        PullProgressText = "";
        IsOllamaOffline = false;

        // ALWAYS populate model list first so user sees something immediately
        if (ChatModelList.Count == 0)
        {
            foreach (var (id, label, size, desc) in OllamaService.SupportedModels)
            {
                ChatModelList.Add(new ChatModelItem
                {
                    Id = id,
                    Label = label,
                    Size = size,
                    Description = desc,
                    IsInstalled = false
                });
            }
        }

        try
        {
            // Quick check — if Ollama is already running, skip startup entirely
            if (await _ollamaService.CheckConnectionAsync())
            {
                IsOllamaOffline = false;
            }
            else
            {
                // Show progress banner for ANY Ollama startup (install or just starting)
                IsInstallingOllama = true;
                ChatModelStatus = "Ollama wird gestartet...";

                try
                {
                    await _ollamaService.EnsureRunningAsync(
                        progress => Dispatcher.UIThread.Post(() =>
                        {
                            ChatModelStatus = progress;
                            PullProgressText = progress;
                        }));
                }
                finally
                {
                    IsInstallingOllama = false;
                }

                if (!_ollamaService.IsAvailable)
                {
                    ChatModelStatus = _ollamaService.Status;
                    IsOllamaOffline = true;
                    // Ollama might just be slow — keep retrying in background
                    _ = RetryConnectionInBackgroundAsync();
                    return;
                }
                IsOllamaOffline = false;
            }

            // Update installed status now that we know what's downloaded
            foreach (var m in ChatModelList)
                m.IsInstalled = _ollamaService.HasModel(m.Id);

            // Detect available VRAM (NVIDIA via nvidia-smi on Win/Linux,
            // unified memory via sysctl on Apple Silicon) and prefer the
            // recommended model for this hardware if it's installed. Falls
            // back to the first installed model if nothing matches.
            var vramGb = await Services.OllamaService.DetectVramGbAsync();
            var recommendedId = vramGb > 0
                ? Services.OllamaService.RecommendedModelForVram(vramGb)
                : null;
            var recommendedInstalled = recommendedId != null
                ? ChatModelList.FirstOrDefault(m => m.IsInstalled && m.Id == recommendedId)
                : null;
            SelectedChatModel = recommendedInstalled
                ?? ChatModelList.FirstOrDefault(m => m.IsInstalled);
            ShowModelPicker = SelectedChatModel == null;
            UpdateModelStatus();
            // Surface the detection AFTER UpdateModelStatus so it doesn't get
            // overwritten by the default "<id> bereit" message.
            if (vramGb > 0)
            {
                if (SelectedChatModel != null)
                    ChatModelStatus = $"{SelectedChatModel.Id} bereit · {vramGb} GB erkannt";
                else if (recommendedId != null)
                    ChatModelStatus = $"{vramGb} GB erkannt → empfohlen: {recommendedId} (installieren)";
            }

            // Start MCP server in background, then arm the disconnect monitor
            McpStatus = "MCP wird gestartet...";
            _ = Task.Run(async () =>
            {
                await _mcpService.StartAsync(progress =>
                    Dispatcher.UIThread.Post(() => McpStatus = progress));
                Dispatcher.UIThread.Post(() =>
                    McpStatus = _mcpService.IsRunning
                        ? $"MCP: {_mcpService.Tools.Count} Tools bereit"
                        : _mcpService.Status);

                // Arm process health monitor after MCP is up
                if (_mcpService.IsRunning)
                    _ = MonitorMcpProcessAsync();
            });
        }
        catch (Exception ex)
        {
            ChatModelStatus = $"Fehler: {ex.Message}";
            IsOllamaOffline = true;
            IsInstallingOllama = false;
            ShowModelPicker = true;
        }
    }

    partial void OnSelectedChatModelChanged(ChatModelItem? value)
    {
        UpdateModelStatus();
    }

    private void UpdateModelStatus()
    {
        if (SelectedChatModel == null)
        {
            ChatModelStatus = "Waehle ein AI-Modell aus";
            IsModelReady = false;
            return;
        }

        if (SelectedChatModel.IsInstalled)
        {
            ChatModelStatus = $"{SelectedChatModel.Id} bereit";
            IsModelReady = true;
            ShowModelPicker = false;
        }
        else
        {
            ChatModelStatus = $"{SelectedChatModel.Id} muss installiert werden";
            IsModelReady = false;
        }
    }

    private string GetSelectedModelId()
    {
        return ChatProvider == "codex" ? SelectedCodexModel : (SelectedChatModel?.Id ?? "");
    }

    [RelayCommand]
    private async Task InstallModel(ChatModelItem? model)
    {
        if (model == null || model.IsInstalled) return;

        model.IsDownloading = true;
        model.DownloadProgress = "Ollama wird geprueft...";
        model.DownloadPercent = 0;
        model.DownloadSizeText = "";
        IsPullingModel = true;

        try
        {
            // Always verify Ollama is actually reachable, start if needed
            if (!await _ollamaService.CheckConnectionAsync())
            {
                model.DownloadProgress = "Ollama wird gestartet...";
                model.DownloadSizeText = "Bitte warten...";
                await _ollamaService.EnsureRunningAsync(
                    progress => Dispatcher.UIThread.Post(() =>
                    {
                        model.DownloadProgress = progress;
                        ChatModelStatus = progress;
                    }));
            }

            if (!_ollamaService.IsAvailable)
            {
                model.DownloadProgress = "Ollama nicht erreichbar - bitte manuell starten";
                model.DownloadSizeText = "";
                return;
            }

            model.DownloadProgress = "Ollama verbunden, Pull startet...";
            model.DownloadSizeText = "";

            await _ollamaService.PullModelAsync(model.Id,
                onDetailedProgress: detailed =>
                {
                    Dispatcher.UIThread.Post(() =>
                    {
                        if (detailed.TotalBytes > 0)
                        {
                            var dlMb = detailed.CompletedBytes / (1024.0 * 1024.0);
                            var totalMb = detailed.TotalBytes / (1024.0 * 1024.0);
                            var pct = (int)(detailed.Percent * 100);
                            model.DownloadProgress = $"{detailed.Status} — {pct}%";
                            model.DownloadPercent = detailed.Percent;
                            model.DownloadSizeText = $"{dlMb:F0} / {totalMb:F0} MB";
                            PullProgressText = $"{dlMb:F0}/{totalMb:F0} MB ({pct}%)";
                        }
                        else
                        {
                            model.DownloadProgress = detailed.Status;
                            model.DownloadSizeText = "Bitte warten...";
                            PullProgressText = detailed.Status;
                        }
                    });
                });

            await _ollamaService.CheckConnectionAsync();
            model.IsInstalled = true;
            model.DownloadProgress = "";
            model.DownloadPercent = 0;
            model.DownloadSizeText = "";
            SelectedChatModel = model;
            ShowModelPicker = false;
            UpdateModelStatus();
        }
        catch (Exception ex)
        {
            model.DownloadProgress = $"Fehler: {ex.Message}";
            model.DownloadSizeText = "";
        }
        finally
        {
            model.IsDownloading = false;
            IsPullingModel = false;
        }
    }

    [RelayCommand]
    private void ShowModelSelector()
    {
        ShowModelPicker = true;
    }

    [RelayCommand]
    private void SelectInstalledModel(ChatModelItem? model)
    {
        if (model == null || !model.IsInstalled) return;
        SelectedChatModel = model;
        ShowModelPicker = false;
        UpdateModelStatus();
    }

    // Map MCP tool names to user-friendly German status messages
    private static string ToolStatusText(string toolName) => toolName switch
    {
        // Phase 1: System Scan
        "get_gpu_status" => "NVIDIA GPU wird analysiert...",
        "get_system_info" => "System wird gescannt (CPU, RAM, OS)...",
        "diagnose_msfs_config" => "MSFS 2024 Installation wird gesucht...",
        "analyze_msfs_graphics" => "MSFS Grafikeinstellungen werden ausgewertet...",
        "diagnose_pimax" => "Pimax VR-Headset wird erkannt...",
        "analyze_pimax_settings" => "Pimax-Konfiguration wird gelesen...",
        "get_openxr_runtime" => "OpenXR-Runtime wird identifiziert...",
        "analyze_openxr" => "OpenXR Foveated Rendering Status wird gelesen...",
        "analyze_reshade" => "ReShade-Effekte werden analysiert...",
        // Phase 2: GPU
        "check_and_install_driver" => "NVIDIA Treiber wird geprueft...",
        // Phase 3: MSFS
        "optimize_msfs_graphics" => "MSFS Grafik wird auf PS5-Niveau optimiert...",
        "set_msfs_setting" => "MSFS-Einstellung wird gesetzt...",
        "backup_msfs_graphics" => "MSFS Sicherheitskopie wird erstellt...",
        "restore_msfs_graphics" => "MSFS Grafik wird wiederhergestellt...",
        // Phase 4: Pimax
        "optimize_pimax_settings" => "Pimax wird fuer MSFS VR optimiert...",
        "set_pimax_setting" => "Pimax-Einstellung wird konfiguriert...",
        // Phase 5+6: OpenXR
        "set_openxr_setting" => "OpenXR Foveated Rendering wird konfiguriert...",
        "set_openxr_runtime" => "OpenXR Runtime wird umgestellt...",
        // Phase 7: ReShade
        "set_reshade_effect" => "ReShade VR-Schaerfe wird optimiert...",
        "list_reshade_presets" => "ReShade-Presets werden geladen...",
        // Other
        "launch_msfs_vr" => "MSFS VR wird gestartet...",
        "fix_msfs" => "MSFS wird repariert...",
        _ => $"{toolName} wird ausgefuehrt..."
    };

    private string GetSystemPrompt(bool hasMcpTools, string gpuName = "")
    {
        // Rebuild only when the MCP-availability or GPU name changes.
        // Date/time is injected fresh to avoid stale context.
        if (_cachedSystemPrompt != null
            && _cachedMcpState == hasMcpTools
            && _cachedGpuName == gpuName)
        {
            // Patch timestamp in cached prompt (cheap string replace)
            var now = DateTime.Now;
            var timestamped = Regex.Replace(
                _cachedSystemPrompt,
                @"Aktuelles Datum/Uhrzeit: .+",
                $"Aktuelles Datum/Uhrzeit: {now:dddd, dd.MM.yyyy HH:mm}");
            return timestamped;
        }
        var toolCount = _mcpService.Tools?.Count > 0 ? _mcpService.Tools.Count : 142;
        _cachedSystemPrompt = BuildSystemPrompt(hasMcpTools, gpuName, toolCount);
        _cachedMcpState     = hasMcpTools;
        _cachedGpuName      = gpuName;
        return _cachedSystemPrompt;
    }

    private static string BuildSystemPrompt(bool hasMcpTools, string gpuName = "", int toolCount = 142)
    {
        // Professional system prompt: clear identity, explicit
        // capabilities + limits, output format, and honesty rules.
        var sb = new System.Text.StringBuilder();
        var now = DateTime.Now;

        sb.AppendLine("# IDENTITAET");
        sb.AppendLine("Du bist PILOT SUPPORT AI - ein hochprofessioneller Flight-Sim- und VR-System-Engineer.");
        sb.AppendLine("Spezialgebiete: Microsoft Flight Simulator 2024, Pimax VR Headsets, OpenXR, ReShade, GPU-Tuning.");
        sb.AppendLine("Du laeufst lokal auf dem PC des Users (keine Cloud, keine Telemetrie).");
        sb.AppendLine();

        // Inject current date/time so AI can timestamp logs and know context
        sb.AppendLine($"# KONTEXT");
        sb.AppendLine($"Aktuelles Datum/Uhrzeit: {now:dddd, dd.MM.yyyy HH:mm}");
        if (!string.IsNullOrEmpty(gpuName))
            sb.AppendLine($"GPU des Users: {gpuName}");
        sb.AppendLine("OS: Windows 10/11 (x64)");
        sb.AppendLine();

        sb.AppendLine("# STIL");
        sb.AppendLine("- Antworte IMMER auf Deutsch.");
        sb.AppendLine("- Sei praezise und direkt. Keine Floskeln, keine Wiederholungen der Frage.");
        sb.AppendLine("- Strukturiere mit Markdown: Ueberschriften (##), Listen, Code-Bloecke, Tabellen.");
        sb.AppendLine("- Bei technischen Werten: nenne Einheiten (MB, FPS, ms, Hz, °C) und gib den Bereich an, der normal vs. problematisch ist.");
        sb.AppendLine("- Wenn der User nach Optimierung fragt: erst Diagnose, dann konkrete Empfehlung mit Begruendung.");
        sb.AppendLine("- Halte dich kurz wenn moeglich, ausfuehrlich wenn noetig.");
        sb.AppendLine();

        sb.AppendLine("# EHRLICHKEIT");
        sb.AppendLine("- Erfinde NIEMALS Hardware-Daten, FPS-Zahlen, Einstellungen oder Pfade.");
        sb.AppendLine("- Wenn du etwas nicht weisst, sage es. Wenn ein Tool fehlschlaegt, melde den Fehler ehrlich statt zu raten.");
        sb.AppendLine("- Markiere Annahmen explizit als 'vermutlich' oder 'typischerweise' wenn keine echten Daten vorliegen.");
        sb.AppendLine();

        if (hasMcpTools)
        {
            sb.AppendLine($"# TOOLS — VOLLSTAENDIGE UEBERSICHT ({toolCount} Tools verfuegbar)");
            sb.AppendLine("Du bist ein AI-Agent mit System-Tools. Du KANNST und SOLLST Tools aufrufen, um echte Werte vom System zu lesen statt zu raten.");
            sb.AppendLine();

            sb.AppendLine("## MSFS Konfiguration & Diagnose");
            sb.AppendLine("- find_msfs_config          → Findet UserCfg.opt (alle MSFS-Versionen)");
            sb.AppendLine("- diagnose_msfs_config       → Prüft Config auf Probleme");
            sb.AppendLine("- analyze_msfs_graphics      → Analysiert aktuelle Grafik-Einstellungen");
            sb.AppendLine("- optimize_msfs_graphics     → Optimiert Einstellungen für Hardware-Tier");
            sb.AppendLine("- set_msfs_setting           → Setzt einzelne MSFS-Einstellung");
            sb.AppendLine("- set_msfs_multiple_settings → Setzt mehrere Einstellungen auf einmal");
            sb.AppendLine("- get_msfs_current_settings  → Liest ALLE aktuellen MSFS-Einstellungen (Before/After-Vergleich)");
            sb.AppendLine("- restore_msfs_graphics      → Stellt Backup wieder her");
            sb.AppendLine("- backup_msfs_graphics       → Erstellt manuelles Backup");
            sb.AppendLine("- revert_last_change         → Macht letzte Aenderung rueckgaengig");
            sb.AppendLine("- fix_msfs                   → Cache/Shader-Reparatur, Rolling Cache leeren");
            sb.AppendLine("- manage_msfs_rolling_cache  → Rolling Cache verwalten");
            sb.AppendLine("- set_msfs_multiplayer_settings → Live Traffic, Live Weather, Multiplayer ein/aus");
            sb.AppendLine("- get_msfs_weather_settings  → Liest Wetter-Einstellungen (Live Weather, Wolken, Sicht)");
            sb.AppendLine("- set_msfs_weather_settings  → Live Weather ein/aus, Wolken-Qualität setzen");
            sb.AppendLine("- get_msfs_traffic_settings  → KI-Traffic-Dichte, Airline/GA, Online-Spieler");
            sb.AppendLine("- get_vr_headset_connected   → Prüft ob VR-Headset erkannt (Pimax/SteamVR/WMR/Meta)");
            sb.AppendLine("- restart_steamvr            → Beendet und startet SteamVR neu");
            sb.AppendLine("- get_gpu_overclock_status   → GPU Übertaktung via MSI Afterburner Registry + NVML");
            sb.AppendLine("- get_vr_comfort_settings    → VR-Komfort: Motion Blur, Vignette, DLSS-Status");
            sb.AppendLine();

            sb.AppendLine("## MSFS Laufzeit & Prozess-Monitoring");
            sb.AppendLine("- is_msfs_running            → Prüft ob MSFS gerade läuft (Boolean)");
            sb.AppendLine("- get_msfs_process_info      → PID, CPU%, RAM, Fenstertitel von MSFS");
            sb.AppendLine("- get_msfs_fps_estimate      → GPU/CPU-Last als FPS-Proxy (echte FPS: Strg+Z in MSFS)");
            sb.AppendLine("- launch_msfs_vr             → Startet MSFS direkt in VR-Modus");
            sb.AppendLine();

            sb.AppendLine("## MSFS Community & Add-ons");
            sb.AppendLine("- get_msfs_community_folder  → Findet den Community-Ordner-Pfad");
            sb.AppendLine("- list_community_addons      → Listet alle installierten Add-ons (Name, Typ, Groesse)");
            sb.AppendLine("- get_msfs_official_folder   → Findet den Official-Ordner");
            sb.AppendLine("- install_mod                → Installiert Mod aus URL");
            sb.AppendLine("- find_and_install_mod       → Sucht und installiert Mod von flightsim.to");
            sb.AppendLine("- download_mod_via_browser   → Lädt Mod via Browser herunter");
            sb.AppendLine();

            sb.AppendLine("## VR Gesamt-Diagnose & Optimierung");
            sb.AppendLine("- diagnose_vr_complete       → VOLLSTAENDIGE VR-Diagnose (Hardware+MSFS+OpenXR+SteamVR+Pimax+Temps)");
            sb.AppendLine("- optimize_all_for_vr        → Optimiert ALLES fuer VR (MSFS+SteamVR+Pimax+Windows)");
            sb.AppendLine("- status_check               → Schneller Ueberblick: Hardware, MSFS, Pimax");
            sb.AppendLine("- get_vr_render_resolution   → Berechnet effektive Renderaufloesung (Pimax × SteamVR SS)");
            sb.AppendLine();

            sb.AppendLine("## Pimax VR Headset");
            sb.AppendLine("- get_pimax_headset_info     → Headset-Modell, Display-Specs, Firmware");
            sb.AppendLine("- get_pimax_settings         → Alle Pimax-Einstellungen");
            sb.AppendLine("- set_pimax_setting          → Einzelne Pimax-Einstellung setzen");
            sb.AppendLine("- set_pimax_settings         → Mehrere Pimax-Einstellungen");
            sb.AppendLine("- analyze_pimax_settings     → Analysiert Pimax auf Probleme");
            sb.AppendLine("- diagnose_pimax             → Tiefe Pimax-Diagnose");
            sb.AppendLine("- optimize_pimax_settings    → Optimiert Pimax allgemein");
            sb.AppendLine("- optimize_pimax_for_msfs    → Optimiert Pimax speziell fuer MSFS");
            sb.AppendLine("- get_pimax_recommended_settings → Empfohlene Einstellungen fuer Hardware-Tier");
            sb.AppendLine("- get_pimax_openxr_settings  → Pimax OpenXR-Einstellungen");
            sb.AppendLine("- set_pimax_openxr_settings  → Pimax OpenXR-Einstellungen setzen");
            sb.AppendLine("- adjust_pimax_brightness    → Helligkeit/Kontrast anpassen");
            sb.AppendLine("- improve_image_clarity      → Schaerfe und Klarheit verbessern");
            sb.AppendLine("- improve_performance        → Performance-Profil anwenden");
            sb.AppendLine("- restart_pimax              → Pimax-Dienste neu starten");
            sb.AppendLine("- restore_pimax_settings     → Pimax-Backup wiederherstellen");
            sb.AppendLine();

            sb.AppendLine("## OpenXR");
            sb.AppendLine("- get_openxr_runtime         → Aktive OpenXR-Runtime (SteamVR/Pimax/WMR)");
            sb.AppendLine("- set_openxr_runtime         → Wechselt OpenXR-Runtime");
            sb.AppendLine("- analyze_openxr             → Analysiert OpenXR-Konfiguration");
            sb.AppendLine("- set_openxr_setting         → OpenXR-Einstellung setzen");
            sb.AppendLine("- apply_openxr_preset        → OpenXR-Preset anwenden");
            sb.AppendLine("- get_openxr_layers          → Listet aktive OpenXR API-Layers");
            sb.AppendLine();

            sb.AppendLine("## SteamVR");
            sb.AppendLine("- get_steamvr_settings       → Alle SteamVR-Einstellungen");
            sb.AppendLine("- set_steamvr_setting        → Einzelne SteamVR-Einstellung");
            sb.AppendLine("- optimize_steamvr_for_hardware → Optimiert SteamVR fuer Hardware-Tier");
            sb.AppendLine();

            sb.AppendLine("## Hardware & System");
            sb.AppendLine("- get_gpu_status             → GPU-Auslastung, Temp, VRAM, Treiber (NVML)");
            sb.AppendLine("- get_nvidia_driver_info     → Nvidia Treiber-Version und Details");
            sb.AppendLine("- check_and_install_driver   → Prueft und installiert neuesten Nvidia-Treiber");
            sb.AppendLine("- get_detailed_hardware_profile → Vollstaendiges Hardware-Profil mit VR-Tier");
            sb.AppendLine("- get_cpu_info               → CPU: Kerne, Takt, Auslastung, Temperatur");
            sb.AppendLine("- get_ram_info               → RAM: Gesamt, verfügbar, Takt, Dual-Channel");
            sb.AppendLine("- get_disk_info              → Laufwerke: SSD/HDD, freier Speicher, MSFS-Laufwerk");
            sb.AppendLine("- get_system_temps           → CPU- und GPU-Temperaturen");
            sb.AppendLine("- get_system_info            → OS, CPU, RAM, Laufwerke, Netzwerk-Adapter");
            sb.AppendLine("- get_display_info           → Monitore: Auflösung, Refresh, HDR");
            sb.AppendLine();

            sb.AppendLine("## Audio");
            sb.AppendLine("- get_audio_devices          → Alle Audio-Geräte (Wiedergabe + Aufnahme)");
            sb.AppendLine("- set_default_audio_device   → Setzt VR-Headset als Standard-Audio");
            sb.AppendLine("- get_spatial_audio_status   → Windows Sonic / Dolby Atmos Status");
            sb.AppendLine();

            sb.AppendLine("## Windows Optimierung");
            sb.AppendLine("- optimize_windows_for_vr    → Windows-Tweaks fuer VR (HAGS, GameMode, etc.)");
            sb.AppendLine("- set_power_plan             → Power-Plan setzen (Ultimate/High/Balanced)");
            sb.AppendLine("- set_hardware_accelerated_gpu_scheduling → HAGS ein/aus");
            sb.AppendLine("- set_virtual_memory         → Auslagerungsdatei anpassen");
            sb.AppendLine("- set_nvidia_low_latency_mode → Nvidia Low Latency Modus");
            sb.AppendLine("- set_windows_display_scaling → DPI-Skalierung setzen");
            sb.AppendLine("- kill_background_processes  → Beendet VR-feindliche Hintergrundprozesse");
            sb.AppendLine("- manage_processes           → Prozesse auflisten, suchen, beenden");
            sb.AppendLine("- manage_services            → Windows-Dienste verwalten");
            sb.AppendLine("- manage_startup_programs    → Autostart-Programme verwalten");
            sb.AppendLine();

            sb.AppendLine("## Netzwerk");
            sb.AppendLine("- get_network_info           → Ping zu MSFS-Servern, LAN/WLAN-Check");
            sb.AppendLine("- network_diagnostics        → Ping, Traceroute, DNS, Verbindungen");
            sb.AppendLine();

            sb.AppendLine("## Logging & Diagnose");
            sb.AppendLine("- get_msfs_log               → Letzte N Zeilen des MSFS-Logs");
            sb.AppendLine("- get_event_log_errors       → Windows Event Log: GPU-TDRs, MSFS-Fehler");
            sb.AppendLine("- export_vr_diagnostic_report → Vollständiger HTML-Bericht auf Desktop");
            sb.AppendLine();

            sb.AppendLine("## ReShade");
            sb.AppendLine("- analyze_reshade                 → ReShade-Konfiguration analysieren");
            sb.AppendLine("- set_reshade_effect              → Einzelnen Effekt aktivieren/deaktivieren");
            sb.AppendLine("- apply_reshade_preset            → VR-optimiertes Preset anwenden");
            sb.AppendLine("- list_reshade_presets            → Alle verfügbaren Presets");
            sb.AppendLine("- uninstall_reshade               → ReShade entfernen");
            sb.AppendLine("- get_reshade_status              → ReShade-Installationsstatus + DLL-Version");
            sb.AppendLine("- get_reshade_settings            → Alle ReShade.ini-Einstellungen lesen");
            sb.AppendLine("- optimize_reshade_for_vr         → Alle VR-Einstellungen automatisch optimieren");
            sb.AppendLine("- get_reshade_effects             → Alle Shader-Effekte mit GPU-Kosten");
            sb.AppendLine("- enable_reshade_technique        → Einzelne Technik ein-/ausschalten");
            sb.AppendLine("- set_reshade_performance_mode    → Performance-Modus ein/aus");
            sb.AppendLine("- set_reshade_preset              → Auf anderes Preset umschalten");
            sb.AppendLine("- get_reshade_presets             → Alle Presets mit Effektliste");
            sb.AppendLine("- diagnose_reshade_vr             → Vollständige VR-Diagnose (DLL, Tiefenpuffer, Performance)");
            sb.AppendLine("- create_reshade_vr_preset        → Standard VR-Preset erstellen");
            sb.AppendLine("- create_reshade_vr_visual_preset → GPU-Tier-basiertes Qualitäts-Preset (HDR-aware)");
            sb.AppendLine("- tune_reshade_sharpness          → CAS-Schärfe anpassen (low/medium/high/ultra)");
            sb.AppendLine("- tune_reshade_colors             → Farben anpassen (Sättigung/Wärme/Kontrast)");
            sb.AppendLine("- apply_reshade_style             → Visuellen Stil anwenden (cinematic/night/day/sunset/dawn/dusk/vivid/natural/sharp)");
            sb.AppendLine("- get_reshade_visual_analysis     → Bildqualitäts-Score + Verbesserungsvorschläge");
            sb.AppendLine("- auto_apply_reshade_for_time     → ReShade automatisch an Tageszeit anpassen");
            sb.AppendLine("- get_reshade_addons              → Installierte .addon-Dateien auflisten (ReShade 5+)");
            sb.AppendLine("- check_reshade_vr_compatibility  → DLL-Version prüfen, ReShade 5+ für VR/OpenXR");
            sb.AppendLine("- backup_reshade_preset           → Aktives Preset mit Datum sichern");
            sb.AppendLine("- list_reshade_preset_backups     → Alle Preset-Backups anzeigen");
            sb.AppendLine("- restore_reshade_preset          → Preset aus Backup wiederherstellen");
            sb.AppendLine("- save_reshade_profile_as         → Aktuelles Preset unter eigenem Namen speichern");
            sb.AppendLine("- list_reshade_custom_profiles    → Alle eigenen Presets auflisten");
            sb.AppendLine();

            sb.AppendLine("## Updates");
            sb.AppendLine("- check_for_gameCopilot_update  → Prüft ob neue Game Copilot Version verfügbar");
            sb.AppendLine("- install_gameCopilot_update    → Installiert Update (confirmed=true nach Bestätigung)");
            sb.AppendLine("- check_for_mcp_server_update   → Prüft ob neue MCP-Server-Version verfügbar");
            sb.AppendLine("- update_mcp_server             → Lädt neue MCP-Server-Version (confirmed=true)");
            sb.AppendLine();

            sb.AppendLine("## Neue Tools (v3.6.2)");
            sb.AppendLine("- set_windows_theme             → Windows Dark/Light Mode umschalten");
            sb.AppendLine("- get_msfs_installed_version    → Installierte MSFS 2024 Version lesen");
            sb.AppendLine("- set_cpu_priority_msfs         → MSFS CPU-Priorität auf High/Realtime setzen");
            sb.AppendLine("- get_msfs_addon_count          → Schnelle Add-on-Anzahl im Community-Ordner");
            sb.AppendLine("- set_steamvr_supersampling     → SteamVR Supersampling direkt setzen");
            sb.AppendLine("- get_pimax_play_version        → Pimax Play / PiTool Version lesen");
            sb.AppendLine("- take_vr_screenshot            → MSFS Screenshot per Tastendruck auslösen");
            sb.AppendLine("- open_msfs_devmode             → MSFS Developer Mode ein/ausschalten");
            sb.AppendLine("- get_frame_generation_status   → DLSS/FSR Frame Generation Status lesen");
            sb.AppendLine("- set_frame_generation          → Frame Generation ein/ausschalten");
            sb.AppendLine();

            sb.AppendLine("## Browser & Misc");
            sb.AppendLine("- ensure_chrome_debug / browser_navigate / browser_click / browser_type / browser_read_page");
            sb.AppendLine("- manage_firewall / disk_analysis / windows_update_status / run_shell_command");
            sb.AppendLine();

            sb.AppendLine("## ENTSCHEIDUNGSHILFE — Wann welches Tool:");
            sb.AppendLine("- 'VR ruckelt / schlechte FPS'      → diagnose_vr_complete");
            sb.AppendLine("- 'Alles optimieren'                 → optimize_all_for_vr");
            sb.AppendLine("- 'Pimax einstellen'                 → get_pimax_headset_info → optimize_pimax_for_msfs");
            sb.AppendLine("- 'MSFS-Einstellungen anpassen'      → get_msfs_current_settings → set_msfs_setting");
            sb.AppendLine("- 'Temperaturproblem'                → get_system_temps");
            sb.AppendLine("- 'MSFS stürzt ab'                   → get_event_log_errors + get_msfs_log + fix_msfs");
            sb.AppendLine("- 'RAM/CPU/Disk-Frage'               → get_ram_info / get_cpu_info / get_disk_info");
            sb.AppendLine("- 'Add-ons prüfen'                   → list_community_addons");
            sb.AppendLine("- 'Diagnose-Bericht erstellen'        → export_vr_diagnostic_report");
            sb.AppendLine("- 'Updates prüfen'                   → check_for_gameCopilot_update");
            sb.AppendLine();

            sb.AppendLine("## TOOL-REIHENFOLGE (PFLICHT)");
            sb.AppendLine("1. IMMER erst LESEN (get_*, analyze_*, diagnose_*) bevor du schreibst.");
            sb.AppendLine("2. SCHREIBEN (set_*, optimize_*, fix_*) NUR mit expliziter User-Bestätigung.");
            sb.AppendLine("3. Bei Fehlern: melde den Fehler klar — erfinde KEINE Werte.");
            sb.AppendLine();
            sb.AppendLine("WICHTIG: Wenn der Nutzer fragt ob es Updates gibt, rufe IMMER zuerst check_for_gameCopilot_update auf.");
            sb.AppendLine("Wenn install_gameCopilot_update 'neustart: true' zurückgibt → User informieren dass App neu startet.");
            sb.AppendLine("Wenn update_mcp_server neustart_erforderlich=true → User soll 'MCP Server neustarten' sagen.");
            sb.AppendLine();

            sb.AppendLine("# AUSGABE-FORMAT FUER TOOL-ERGEBNISSE");
            sb.AppendLine("Tool-Output IMMER als Markdown-Tabelle: | Einstellung | Aktuell | Empfehlung |");
            sb.AppendLine("Hinter der Tabelle: 1-3 Saetze mit Erklaerung WARUM die Empfehlung sinnvoll ist.");
            sb.AppendLine("Bei langen Listen: gruppiere thematisch (Grafik / Audio / Steuerung etc.).");
            sb.AppendLine();
            sb.AppendLine("Kontext: Der User hat ein Pimax VR Headset.");
            // Qwen3 /no_think: blocks hidden chain-of-thought that makes tool calls 10-60s slower
            sb.AppendLine();
            sb.AppendLine("/no_think");
        }
        else
        {
            sb.AppendLine("# KEINE TOOLS VERFUEGBAR");
            sb.AppendLine("Der MCP-Server ist gerade nicht verbunden. Du hast KEINEN Zugriff auf System-Daten.");
            sb.AppendLine("Sage dem User klar: 'MCP-Server ist nicht verbunden - bitte App neu starten.' wenn er nach echten Werten fragt.");
            sb.AppendLine("Allgemeine Fragen zu MSFS, VR oder Hardware kannst du trotzdem beantworten - markiere alle Werte als allgemeine Richtwerte, nicht als seine echten Daten.");
        }

        return sb.ToString().TrimEnd();
    }

    [RelayCommand]
    private async Task SendChatMessage()
    {
        var prompt = ChatInput?.Trim();
        if (string.IsNullOrEmpty(prompt)) return;
        if (IsGenerating) return;  // guard against double-send

        var isCodex = ChatProvider == "codex";
        var modelId = GetSelectedModelId();

        if (isCodex)
        {
            // ── Cloud / Codex path ──────────────────────────────────────────
            if (string.IsNullOrEmpty(CodexApiKey))
            {
                CodexApiKeyMissing = true;
                ChatMessages.Add(new ChatMessage
                {
                    Role = "assistant",
                    Content = "Bitte trage deinen OpenAI API-Key ein, um Codex zu nutzen.",
                    Timestamp = DateTime.Now.ToString("HH:mm")
                });
                return;
            }
            if (string.IsNullOrEmpty(modelId))
                modelId = "codex-mini-latest";
        }
        else
        {
            // ── Local / Ollama path ─────────────────────────────────────────
            if (!_ollamaService.IsAvailable)
            {
                await InitializeOllamaChat();
                if (!_ollamaService.IsAvailable) return;
            }

            if (string.IsNullOrEmpty(modelId) || !IsModelReady)
            {
                ShowModelPicker = true;
                ChatModelStatus = "Bitte installiere zuerst ein AI-Modell";
                return;
            }
        }

        ChatMessages.Add(new ChatMessage
        {
            Role = "user",
            Content = prompt,
            Timestamp = DateTime.Now.ToString("HH:mm")
        });
        ChatInput = "";
        IsChatLoading = true;
        IsGenerating = true;
        ThinkingStatusText = "Analysiere...";
        IsToolRunning = false;

        // Create a fresh cancellation source for this generation
        _generationCts?.Dispose();
        _generationCts = new CancellationTokenSource();
        var ct = _generationCts.Token;

        var streamMsg = new ChatMessage
        {
            Role = "assistant",
            Content = "",
            Timestamp = DateTime.Now.ToString("HH:mm")
        };
        bool messageAdded = false;

        try
        {
            // Auto-start MCP if not running — always use tools
            if (!_mcpService.IsRunning)
            {
                Dispatcher.UIThread.Post(() => ThinkingStatusText = "MCP Server wird gestartet...");
                await Task.Run(async () => await _mcpService.StartAsync().ConfigureAwait(false));
            }

            var hasMcp = _mcpService.IsRunning;
            var tools = hasMcp ? _mcpService.GetOllamaToolDefinitions() : null;
            var systemPrompt = GetSystemPrompt(hasMcp, _detectedGpuName);

            // Build messages list for Ollama (last 20 messages for context)
            var ollamaMessages = new System.Collections.Generic.List<object>();
            ollamaMessages.Add(new { role = "system", content = systemPrompt });

            var recent = ChatMessages.Skip(Math.Max(0, ChatMessages.Count - 21)).Take(20);
            foreach (var msg in recent)
            {
                if (msg == ChatMessages.Last() && msg.IsUser) continue;
                ollamaMessages.Add(new { role = msg.IsUser ? "user" : "assistant", content = msg.Content ?? "" });
            }
            ollamaMessages.Add(new { role = "user", content = prompt });

            // Agent-style tool-calling loop: each tool is a visible step
            if (hasMcp && tools != null && tools.Count > 0)
            {
                // Show assistant message immediately so user sees agent activity
                Dispatcher.UIThread.Post(() =>
                {
                    streamMsg.IsAgentWorking = true;
                    streamMsg.HasAgentSteps = true;
                    streamMsg.Content = "";
                    ChatMessages.Add(streamMsg);
                    IsChatLoading = false;
                    messageAdded = true;
                });

                // Run the entire tool-calling loop on a background thread
                // to keep the UI responsive during long Ollama/MCP calls
                await Task.Run(async () =>
                {
                    const int maxToolRounds = 15;
                    int totalToolCalls = 0;
                    bool reachedFinalAnswer = false;
                    for (int round = 0; round < maxToolRounds; round++)
                    {
                        // Show "thinking" step while waiting for model
                        var thinkingStep = new AgentStep
                        {
                            Icon = "🧠",
                            StatusText = totalToolCalls == 0
                                ? "Anfrage wird analysiert..."
                                : "Naechste Schritte werden geplant...",
                            IconColor = "#a78bfa"
                        };
                        Dispatcher.UIThread.Post(() =>
                        {
                            ThinkingStatusText = totalToolCalls == 0
                                ? "Anfrage wird analysiert..."
                                : "KI verarbeitet Ergebnisse...";
                            streamMsg.AgentSteps.Add(thinkingStep);
                        });

                        // Stream tokens directly into the chat bubble. If the
                        // model decides to call tools instead of answering,
                        // onToken is simply not invoked and ToolCalls fills.
                        var streamedSoFar = new StringBuilder();
                        var firstTokenSeen = false;
                        Action<string> onToken = token =>
                        {
                            streamedSoFar.Append(token);
                            if (!firstTokenSeen)
                            {
                                firstTokenSeen = true;
                                var snapshot1 = streamedSoFar.ToString();
                                Dispatcher.UIThread.Post(() =>
                                {
                                    // Hide the "thinking" placeholder once the
                                    // model actually starts producing answer text.
                                    streamMsg.AgentSteps.Remove(thinkingStep);
                                    IsToolRunning = false;
                                    streamMsg.IsAgentWorking = false;
                                    streamMsg.IsStreaming = true;
                                    streamMsg.Content = CleanMarkdown(snapshot1);
                                });
                            }
                            else
                            {
                                var snapshot2 = streamedSoFar.ToString();
                                Dispatcher.UIThread.Post(() =>
                                {
                                    streamMsg.Content = CleanMarkdown(snapshot2);
                                });
                            }
                        };

                        OllamaStreamWithToolsResult response;
                        if (isCodex)
                        {
                            response = await _codexService.StreamChatWithToolsAsync(
                                modelId, ollamaMessages, tools, onToken, CodexApiKey, ct).ConfigureAwait(false);
                        }
                        else
                        {
                            response = await _ollamaService.StreamChatWithToolsAsync(
                                modelId, ollamaMessages, tools, onToken, ct).ConfigureAwait(false);
                        }

                        // If we never saw a content token (model jumped straight
                        // to a tool call), make sure the thinking step is removed.
                        if (!firstTokenSeen)
                        {
                            Dispatcher.UIThread.Post(() =>
                            {
                                streamMsg.AgentSteps.Remove(thinkingStep);
                            });
                        }

                        if (!response.HasToolCalls)
                        {
                            // Final answer already streamed in — just make sure
                            // the bubble shows the cleaned full text.
                            var finalText = CleanMarkdown(response.Content);
                            var tokenCount = response.TotalTokens;
                            reachedFinalAnswer = true;
                            Dispatcher.UIThread.Post(() =>
                            {
                                IsToolRunning = false;
                                streamMsg.IsAgentWorking = false;
                                streamMsg.Content = finalText;
                                if (tokenCount > 0)
                                    streamMsg.TotalTokens = tokenCount;
                            });
                            break;
                        }

                        // Add assistant message with tool calls to conversation
                        ollamaMessages.Add(System.Text.Json.JsonSerializer.Deserialize<object>(response.RawMessageJson)!);

                        // Execute each tool call as a visible agent step
                        foreach (var tc in response.ToolCalls)
                        {
                            totalToolCalls++;
                            var statusText = ToolStatusText(tc.Name);
                            var step = new AgentStep
                            {
                                Icon = "⏳",
                                StatusText = statusText,
                                IconColor = "#69daff"
                            };

                            Dispatcher.UIThread.Post(() =>
                            {
                                ThinkingStatusText = statusText;
                                IsToolRunning = true;
                                streamMsg.AgentSteps.Add(step);
                            });

                            var toolResult = await _mcpService.CallToolAsync(tc.Name, tc.ArgumentsJson, ct).ConfigureAwait(false);

                            // Parse a short summary from the tool result
                            var summary = ExtractToolSummary(tc.Name, toolResult);

                            var isError = toolResult.Contains("\"error\"") || toolResult.Contains("NOT_INSTALLED") || toolResult.Contains("NOT_FOUND");
                            Dispatcher.UIThread.Post(() =>
                            {
                                step.Icon = isError ? "⚠" : "✓";
                                step.IsComplete = true;
                                step.IconColor = isError ? "#fbbf24" : "#4ade80";
                                step.ResultSummary = summary;
                            });

                            // Truncate large tool results to prevent context overflow
                            var truncated = toolResult.Length > 4000
                                ? toolResult[..4000] + "\n... (gekuerzt)"
                                : toolResult;

                            // If tool returned an error, prefix with explicit instruction
                            // to prevent the AI from inventing/hallucinating data
                            if (truncated.Contains("\"error\"") || truncated.Contains("NOT_INSTALLED") || truncated.Contains("NOT_FOUND"))
                            {
                                truncated = "[TOOL FEHLER - Zeige diese Fehlermeldung dem User. ERFINDE KEINE Daten!]\n" + truncated;
                            }

                            // OpenAI requires tool_call_id; Ollama ignores it.
                            if (isCodex)
                                ollamaMessages.Add(new { role = "tool", tool_call_id = tc.Id, content = truncated });
                            else
                                ollamaMessages.Add(new { role = "tool", content = truncated });
                        }
                    } // end for-loop over tool rounds
                    // If the model used all 15 rounds without a final answer,
                    // surface a clear message so the user isn't left with a blank bubble.
                    if (!reachedFinalAnswer)
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            IsToolRunning = false;
                            streamMsg.IsAgentWorking = false;
                            streamMsg.Content = string.IsNullOrWhiteSpace(streamMsg.Content)
                                ? "Das Modell hat die maximale Anzahl an Tool-Runden erreicht (15). Bitte versuche es mit einer praeziiseren Anfrage."
                                : streamMsg.Content;
                        });
                    }
                });
            }
            else
            {
                // No MCP available — simple streaming chat on background thread
                var fullText = "";
                await Task.Run(async () =>
                {
                    Action<string> tokenHandler = token =>
                    {
                        fullText += token;
                        var cleaned = CleanMarkdown(fullText);
                        Dispatcher.UIThread.Post(() =>
                        {
                            if (!messageAdded)
                            {
                                ChatMessages.Add(streamMsg);
                                IsChatLoading = false;
                                streamMsg.IsStreaming = true;
                                messageAdded = true;
                            }
                            streamMsg.Content = cleaned;
                        });
                    };

                    if (isCodex)
                    {
                        // Codex uses the already-built ollamaMessages (List<object>)
                        await _codexService.StreamChatAsync(
                            modelId, ollamaMessages, tokenHandler, CodexApiKey).ConfigureAwait(false);
                    }
                    else
                    {
                        var chatHistory = new System.Collections.Generic.List<(string role, string content)>(
                            ChatMessages.Skip(Math.Max(0, ChatMessages.Count - 21)).Take(20)
                                .Where(m => !(m == ChatMessages.Last() && m.IsUser))
                                .Select(m => (m.IsUser ? "user" : "assistant", m.Content ?? ""))
                                .Append(("user", prompt)));

                        await _ollamaService.StreamChatAsync(
                            modelId, systemPrompt, chatHistory, tokenHandler).ConfigureAwait(false);
                    }
                });

                if (fullText.Length > 0)
                    Dispatcher.UIThread.Post(() => streamMsg.Content = CleanMarkdown(fullText));
            }

            if (!messageAdded)
            {
                streamMsg.Content = "Keine Antwort erhalten.";
                ChatMessages.Add(streamMsg);
                messageAdded = true;
            }

            streamMsg.Timestamp = DateTime.Now.ToString("HH:mm");
            Dispatcher.UIThread.Post(() => ParseRecommendations(streamMsg));
        }
        catch (Exception ex)
        {
            var errText = $"Fehler: {ex.Message}";
            Dispatcher.UIThread.Post(() => streamMsg.Content = errText);
            if (!messageAdded)
            {
                messageAdded = true;
                Dispatcher.UIThread.Post(() =>
                {
                    ChatMessages.Add(streamMsg);
                    IsChatLoading = false;
                });
            }
        }
        finally
        {
            Dispatcher.UIThread.Post(() =>
            {
                IsChatLoading  = false;
                IsGenerating   = false;
                IsToolRunning  = false;
                streamMsg.IsStreaming    = false;
                streamMsg.IsAgentWorking = false;   // safety net on exception/timeout
                ThinkingStatusText = "Analysiere...";
            });
            // Persist conversation history (debounced)
            ScheduleHistorySave();
        }
    }

    private void ParseRecommendations(ChatMessage msg)
    {
        var text = msg.Content ?? "";

        // Parse [SETTINGS_CARDS]...[/SETTINGS_CARDS]
        var cardsMatch = Regex.Match(text,
            @"\[SETTINGS_CARDS\](.*?)\[/SETTINGS_CARDS\]",
            RegexOptions.Singleline);

        if (cardsMatch.Success)
        {
            var sections = cardsMatch.Groups[1].Value.Split("---");
            foreach (var section in sections)
            {
                var lines = section.Trim().Split('\n', StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) continue;

                var card = new RecommendationCard();
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase))
                        card.Title = trimmed.Substring(6).Trim();
                    else if (trimmed.Contains('='))
                    {
                        var parts = trimmed.Split('=', 2);
                        card.Items.Add(new RecommendationItem
                        {
                            Label = parts[0].Trim(),
                            Value = parts[1].Trim()
                        });
                    }
                }
                if (!string.IsNullOrEmpty(card.Title))
                    msg.Cards.Add(card);
            }
        }

        // Parse [ACTION_BUTTONS]...[/ACTION_BUTTONS]
        var actionsMatch = Regex.Match(text,
            @"\[ACTION_BUTTONS\](.*?)\[/ACTION_BUTTONS\]",
            RegexOptions.Singleline);

        if (actionsMatch.Success)
        {
            var lines = actionsMatch.Groups[1].Value.Trim()
                .Split('\n', StringSplitOptions.RemoveEmptyEntries);
            bool first = true;
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                {
                    msg.ActionButtons.Add(new ActionButton
                    {
                        Text = trimmed,
                        IsPrimary = first
                    });
                    first = false;
                }
            }
        }

        // Remove parsed blocks from displayed text
        if (cardsMatch.Success || actionsMatch.Success)
        {
            var cleaned = text;
            if (cardsMatch.Success)
                cleaned = cleaned.Replace(cardsMatch.Value, "");
            if (actionsMatch.Success)
                cleaned = cleaned.Replace(actionsMatch.Value, "");
            msg.Content = cleaned.Trim();
        }

        msg.HasCards = msg.Cards.Count > 0;
        msg.HasActions = msg.ActionButtons.Count > 0;
    }

    private static string CleanMarkdown(string text)
    {
        // Remove Qwen3 <think>...</think> blocks (reasoning tokens not meant for display)
        text = Regex.Replace(text, @"<think>[\s\S]*?</think>", "", RegexOptions.IgnoreCase);
        // Also remove unclosed <think> (still streaming or incomplete)
        text = Regex.Replace(text, @"<think>[\s\S]*$", "", RegexOptions.IgnoreCase);
        // Remove ### ## # headers but keep the text
        text = Regex.Replace(text, @"^#{1,6}\s*", "", RegexOptions.Multiline);
        // Remove bold **text** → text
        text = Regex.Replace(text, @"\*\*(.+?)\*\*", "$1");
        // Remove italic *text* → text
        text = Regex.Replace(text, @"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", "$1");
        // Remove --- horizontal rules
        text = Regex.Replace(text, @"^-{3,}\s*$", "", RegexOptions.Multiline);
        // Convert - bullets to •
        text = Regex.Replace(text, @"^-\s+", "• ", RegexOptions.Multiline);
        // Remove inline code backticks
        text = Regex.Replace(text, @"`([^`]+)`", "$1");
        // Remove code block markers
        text = Regex.Replace(text, @"^```\w*\s*$", "", RegexOptions.Multiline);
        // Clean up triple+ blank lines
        text = Regex.Replace(text, @"\n{3,}", "\n\n");
        return text.Trim();
    }

    /// <summary>
    /// Extract a short human-readable summary from a tool's JSON result.
    /// Shows the user what was found/changed without raw JSON.
    /// </summary>
    private static string ExtractToolSummary(string toolName, string resultJson)
    {
        try
        {
            var doc = System.Text.Json.JsonDocument.Parse(resultJson);
            var root = doc.RootElement;

            // Error case
            if (root.TryGetProperty("error", out var err))
                return $"Fehler: {Truncate(err.GetString() ?? "", 80)}";

            return toolName switch
            {
                "get_gpu_status" => TryGet(root, "gpu_name", "gpu_temp_c", "vram_used_mb", "vram_total_mb")
                    is (string name, string temp, string used, string total)
                    ? $"{name}, {temp}°C, VRAM {used}/{total} MB"
                    : TryGetAny(root, "gpu_name", "status"),

                "get_system_info" => TryGet(root, "cpu", "ram_total_gb")
                    is (string cpu, string ram, _, _)
                    ? $"{Truncate(cpu, 40)}, {ram} GB RAM"
                    : TryGetAny(root, "os", "status"),

                "analyze_msfs_graphics" or "diagnose_msfs_config" =>
                    TryGetAny(root, "dlss_mode", "gpu_tier", "status", "file"),

                "optimize_msfs_graphics" => TryGet(root, "gpu", "gpu_tier", "mode")
                    is (string gpu, string tier, string mode, _)
                    ? $"{gpu} ({tier}) — {mode}-Preset angewendet"
                    : TryGetAny(root, "verifiziert", "status"),

                "analyze_pimax_settings" or "diagnose_pimax" =>
                    TryGetAny(root, "renderResolution", "refreshRate", "fov", "status"),

                "optimize_pimax_settings" => TryGet(root, "preset")
                    is (string preset, _, _, _)
                    ? $"Preset \"{preset}\" angewendet"
                    : TryGetAny(root, "status"),

                "set_openxr_setting" => TryGet(root, "einstellung", "neuer_wert_display")
                    is (string setting, string val, _, _)
                    ? $"{setting} → {val}"
                    : TryGetAny(root, "einstellung", "status"),

                "set_pimax_setting" =>
                    TryGetAny(root, "einstellung", "neuer_wert", "status"),

                "analyze_openxr" or "get_openxr_runtime" =>
                    TryGetAny(root, "runtime", "vrs", "status"),

                "analyze_reshade" =>
                    TryGetAny(root, "active_preset", "effects_count", "status"),

                "set_reshade_effect" =>
                    TryGetAny(root, "effekt", "status"),

                "set_msfs_setting" => TryGet(root, "einstellung", "neuer_wert")
                    is (string setting, string val, _, _)
                    ? $"{setting} → {val}"
                    : TryGetAny(root, "einstellung", "status"),

                "fix_msfs" => TryGet(root, "aktion", "ergebnis")
                    is (string action, string result, _, _)
                    ? $"{Truncate(action, 30)}: {Truncate(result, 40)}"
                    : TryGetAny(root, "aktion", "status", "ergebnis"),

                "check_and_install_driver" => TryGet(root, "installiert", "version")
                    is (string installed, string version, _, _)
                    ? (installed == "True" ? $"Treiber {version} installiert" : $"Aktuell: {version}")
                    : TryGetAny(root, "version", "status"),

                _ => TryGetAny(root, "status", "result", "message") ?? "Fertig"
            };
        }
        catch
        {
            return Truncate(resultJson, 60);
        }
    }

    private static (string, string, string, string)? TryGet(
        System.Text.Json.JsonElement root, string k1, string k2 = "", string k3 = "", string k4 = "")
    {
        string? v1 = null, v2 = null, v3 = null, v4 = null;
        if (root.TryGetProperty(k1, out var e1)) v1 = e1.ToString();
        if (!string.IsNullOrEmpty(k2) && root.TryGetProperty(k2, out var e2)) v2 = e2.ToString();
        if (!string.IsNullOrEmpty(k3) && root.TryGetProperty(k3, out var e3)) v3 = e3.ToString();
        if (!string.IsNullOrEmpty(k4) && root.TryGetProperty(k4, out var e4)) v4 = e4.ToString();
        return v1 != null ? (v1, v2 ?? "", v3 ?? "", v4 ?? "") : null;
    }

    private static string TryGetAny(System.Text.Json.JsonElement root, params string[] keys)
    {
        var parts = new System.Collections.Generic.List<string>();
        foreach (var k in keys)
        {
            if (root.TryGetProperty(k, out var val))
            {
                var s = val.ToString();
                if (!string.IsNullOrEmpty(s))
                    parts.Add($"{k}: {Truncate(s, 40)}");
            }
            if (parts.Count >= 2) break;
        }
        return parts.Count > 0 ? string.Join(" | ", parts) : "OK";
    }

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "...";

    /// <summary>Clear all chat messages and erase history from disk.</summary>
    [RelayCommand]
    private void ClearChat()
    {
        ChatMessages.Clear();
        _historyService.Clear();
        ShowToast("Chat geleert", "Alle Nachrichten wurden entfernt.");
    }

    /// <summary>
    /// Start a background task that checks if the MCP process is still alive every 15s.
    /// Shows the yellow disconnect banner if the process has exited unexpectedly.
    /// </summary>
    private async Task MonitorMcpProcessAsync()
    {
        // Wait for MCP to be running before we start watching
        for (int i = 0; i < 30 && !_mcpService.IsRunning; i++)
            await Task.Delay(2000);

        while (_mcpService.IsRunning)
        {
            await Task.Delay(15_000);

            if (_mcpService.HasProcessExited && _mcpService.IsRunning)
            {
                _mcpService.IsRunning = false;
                Dispatcher.UIThread.Post(() =>
                {
                    McpDisconnected = true;
                    McpStatus = "MCP getrennt";
                });
                break;
            }
        }
    }

    /// <summary>Reconnect a dropped MCP server.</summary>
    [RelayCommand]
    private async Task ReconnectMcp()
    {
        McpDisconnected = false;
        McpStatus = "MCP wird neugestartet...";
        try
        {
            await Task.Run(async () =>
            {
                await _mcpService.RestartAsync(progress =>
                    Dispatcher.UIThread.Post(() => McpStatus = progress));
            });
            Dispatcher.UIThread.Post(() =>
                McpStatus = _mcpService.IsRunning
                    ? $"MCP: {_mcpService.Tools.Count} Tools bereit"
                    : _mcpService.Status);

            // Re-arm the monitor
            _ = Task.Run(MonitorMcpProcessAsync);
        }
        catch (Exception ex)
        {
            McpStatus = $"MCP Fehler: {ex.Message}";
        }
    }

    [RelayCommand]
    private void SendSuggestion(string suggestion)
    {
        ChatInput = suggestion;
        SendChatMessageCommand.Execute(null);
    }

    [RelayCommand]
    private void ChatAction(string action)
    {
        ChatInput = action;
        SendChatMessageCommand.Execute(null);
    }

    // ======================================================
    // GAMES PAGE - Migrated from WPF GamesPage.xaml.cs
    // Button_VRStart + DesktopStartButton_Click
    // ======================================================

    [RelayCommand]
    private async Task StartVR()
    {
        VrErrorText = "";
        await _gameLaunchService.StartVRAsync(
            text => Dispatcher.UIThread.Post(() => VrButtonText = text),
            enabled => Dispatcher.UIThread.Post(() => VrButtonEnabled = enabled),
            () => Dispatcher.UIThread.Post(() =>
            {
                VrButtonEnabled = true;
                VrButtonText = "In VR starten";
            })
        );
    }

    [RelayCommand]
    private async Task StartDesktop()
    {
        await _gameLaunchService.StartDesktopAsync(
            text => Dispatcher.UIThread.Post(() => DesktopButtonText = text),
            enabled => Dispatcher.UIThread.Post(() => DesktopButtonEnabled = enabled),
            () => Dispatcher.UIThread.Post(() =>
            {
                DesktopButtonEnabled = true;
                DesktopButtonText = "Normal starten";
            })
        );
    }

    // ======================================================
    // MODS PAGE - Migrated from WPF ModsPage.xaml.cs
    // LoadMods, DeleteMod, RefreshMods, filter
    // ======================================================

    private void LoadAllMods()
    {
        Mods.Clear();
        FilteredMods.Clear();

        // Only load real mods from filesystem - exactly like WPF ModsPage.LoadMods()
        var mods = _modService.LoadMods();
        foreach (var mod in mods)
        {
            Mods.Add(mod);
            FilteredMods.Add(mod);
        }
        ModCount = Mods.Count;
        TotalModSize = CalculateTotalModSize();
        RebuildModGrid();

        if (!System.IO.Directory.Exists(_configService.CommunityPath))
            StatusMessage = "Community-Ordner nicht gefunden - bitte in Einstellungen konfigurieren";
    }

    /// <summary>Migrated from WPF ModsPage.FilterMods()</summary>
    private void ApplyModFilter()
    {
        FilteredMods.Clear();
        foreach (var mod in Mods)
        {
            if (!string.IsNullOrWhiteSpace(SearchText)
                && !mod.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase)
                && !mod.FullPath.Contains(SearchText, StringComparison.OrdinalIgnoreCase))
                continue;

            if (ModCategoryFilter != "All" && mod.Category != ModCategoryFilter)
                continue;

            FilteredMods.Add(mod);
        }
        RebuildModGrid();
    }

    private void RebuildModGrid()
    {
        ModGridItems.Clear();
        foreach (var mod in FilteredMods)
            ModGridItems.Add(mod);
        ModGridItems.Add(new AddModPlaceholder());
    }

    private string CalculateTotalModSize()
    {
        long total = 0;
        foreach (var mod in Mods)
        {
            try
            {
                if (System.IO.Directory.Exists(mod.FullPath))
                    total += new System.IO.DirectoryInfo(mod.FullPath)
                        .EnumerateFiles("*", System.IO.SearchOption.AllDirectories)
                        .Sum(f => f.Length);
                else if (System.IO.File.Exists(mod.FullPath))
                    total += new System.IO.FileInfo(mod.FullPath).Length;
            }
            catch { }
        }
        return total switch
        {
            >= 1_073_741_824 => $"{total / 1_073_741_824.0:F1} GB",
            >= 1_048_576 => $"{total / 1_048_576.0:F0} MB",
            >= 1_024 => $"{total / 1_024.0:F0} KB",
            _ => $"{total} B"
        };
    }

    /// <summary>Migrated from WPF ModsPage.RefreshMods_Click()</summary>
    [RelayCommand]
    private void RefreshMods()
    {
        LoadAllMods();
        SearchText = "";
        StatusMessage = $"Mods aktualisiert - {ModCount} Mods gefunden";
    }

    /// <summary>Shows confirmation modal before deleting.</summary>
    [RelayCommand]
    private void DeleteMod(ModItem? mod)
    {
        if (mod == null) return;
        _pendingDeleteMod = mod;
        DeleteModName = mod.Name;
        DeleteModalVisible = true;
    }

    [RelayCommand]
    private void ConfirmDeleteMod()
    {
        if (_pendingDeleteMod == null) return;

        _modService.DeleteMod(_pendingDeleteMod);
        Mods.Remove(_pendingDeleteMod);
        FilteredMods.Remove(_pendingDeleteMod);
        ModCount = Mods.Count;
        TotalModSize = CalculateTotalModSize();
        RebuildModGrid();
        StatusMessage = $"Mod '{_pendingDeleteMod.Name}' geloescht";
        _pendingDeleteMod = null;
        DeleteModalVisible = false;
    }

    [RelayCommand]
    private void CancelDeleteMod()
    {
        _pendingDeleteMod = null;
        DeleteModalVisible = false;
    }

    /// <summary>Called from code-behind after file picker. Migrated from WPF AddMod_Click()</summary>
    public void AddModFromFile(string filePath)
    {
        try
        {
            var mod = _modService.AddMod(filePath);
            if (mod != null)
            {
                Mods.Add(mod);
                FilteredMods.Add(mod);
                ModCount = Mods.Count;
                TotalModSize = CalculateTotalModSize();
                RebuildModGrid();
                StatusMessage = $"Mod '{mod.Name}' hinzugefuegt";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Fehler: {ex.Message}";
        }
    }

    /// <summary>Called from code-behind after input dialog. Migrated from WPF RenameMod_Click()</summary>
    public void RenameMod(ModItem mod, string newName)
    {
        var (success, newPath) = _modService.RenameMod(mod, newName);
        if (success)
        {
            mod.Name = newName;
            mod.FullPath = newPath;
            ApplyModFilter();
            StatusMessage = $"Mod umbenannt zu '{newName}'";
        }
        else
        {
            StatusMessage = "Umbenennen fehlgeschlagen";
        }
    }

    // ======================================================
    // RESHADE PRESETS
    // ======================================================

    private void LoadPresets()
    {
        Presets.Clear();

        // Primary: load real ReShade .ini presets from game directory
        var reshadePresets = _reshadeService.LoadAllPresets();
        if (reshadePresets.Count > 0)
        {
            foreach (var preset in reshadePresets)
                Presets.Add(preset);

            // Select the currently active preset (from ReShade.ini PresetPath)
            var activePath = _reshadeService.ActivePresetPath;
            var active = activePath != null
                ? Presets.FirstOrDefault(p =>
                    p.FilePath.Equals(activePath, StringComparison.OrdinalIgnoreCase))
                : null;

            SelectedPreset = active ?? Presets[0];
            if (active != null)
                ActivePresetName = active.Name;
        }
        else
        {
            // Fallback: load from JSON presets (no ReShade installation found)
            foreach (var preset in _presetService.LoadAllPresets())
                Presets.Add(preset);
            if (Presets.Count > 0)
                SelectedPreset = Presets[0];
        }
    }

    private void DetectEnvironment()
    {
        MsfsPath = _reshadeService.MsfsGamePath ?? _reshadeService.MsfsConfigPath ?? "Nicht erkannt";
        ReshadePath = _reshadeService.ReshadePath ?? "Nicht erkannt";
        IsReshadeFound = _reshadeService.IsReshadeFound;
        // Fire-and-forget Ollama status check — run on thread pool, marshal result back to UI thread.
        _ = Task.Run(async () =>
        {
            try
            {
                await _ollamaService.CheckConnectionAsync().ConfigureAwait(false);
            }
            catch { /* best-effort — failure means Ollama is offline */ }
            Dispatcher.UIThread.Post(() => OllamaStatus = _ollamaService.Status);
        });
    }

    private void SyncToSelectedPreset()
    {
        if (SelectedPreset == null) return;
        SelectedPreset.SharpenEnabled = SharpenEnabled;
        SelectedPreset.BloomEnabled = BloomEnabled;
        SelectedPreset.VibranceEnabled = VibranceEnabled;
        SelectedPreset.TonemapEnabled = TonemapEnabled;
        SelectedPreset.SharpenStrength = SharpenStrength;
        SelectedPreset.BloomStrength = BloomStrength;
        SelectedPreset.VibranceStrength = VibranceStrength;
        SelectedPreset.Contrast = Contrast;
        SelectedPreset.Brightness = Brightness;
    }

    [RelayCommand]
    private void SelectPresetCard(ReShadePreset? preset)
    {
        if (preset == null) return;
        SelectedPreset = preset;
        ShowShaderDetail = true;
        StatusMessage = $"Preset '{preset.Name}' geladen";
    }

    [RelayCommand] private void LoadPresetCmd()
    {
        if (SelectedPreset == null) return;
        OnSelectedPresetChanged(SelectedPreset);
        StatusMessage = $"Preset '{SelectedPreset.Name}' geladen";
    }

    [RelayCommand] private void SavePreset()
    {
        if (SelectedPreset == null) return;
        SyncToSelectedPreset();

        // Save to .ini if it's a ReShade preset, otherwise JSON
        if (!string.IsNullOrEmpty(SelectedPreset.FilePath))
            _reshadeService.SavePresetFile(SelectedPreset);
        else
            _presetService.SavePreset(SelectedPreset);

        _originalPresetState = SelectedPreset.Clone();
        StatusMessage = $"Preset '{SelectedPreset.Name}' gespeichert";
    }

    [RelayCommand] private void NewPreset()
    {
        var name = $"Neues Preset {Presets.Count + 1}";

        // Try to create as .ini in ReShade presets directory
        var basedOn = SelectedPreset;
        if (basedOn != null) SyncToSelectedPreset();

        var preset = _reshadeService.CreateNewPreset(name, basedOn);
        if (preset == null)
        {
            // Fallback to JSON preset if no ReShade directory available
            preset = new ReShadePreset
            {
                Name = name, Description = "Benutzerdefiniertes Preset",
                SharpenEnabled = SharpenEnabled, BloomEnabled = BloomEnabled,
                VibranceEnabled = VibranceEnabled, TonemapEnabled = TonemapEnabled,
                SharpenStrength = SharpenStrength, BloomStrength = BloomStrength,
                VibranceStrength = VibranceStrength, Contrast = Contrast, Brightness = Brightness,
            };
            _presetService.SavePreset(preset);
        }

        Presets.Add(preset);
        SelectedPreset = preset;
        StatusMessage = $"Neues Preset '{name}' erstellt";
    }

    [RelayCommand] private void DeletePreset()
    {
        if (SelectedPreset == null) return;
        DeletePresetItem(SelectedPreset);
    }

    [RelayCommand]
    private void DeletePresetItem(ReShadePreset? preset)
    {
        if (preset == null) return;
        var name = preset.Name;

        if (!string.IsNullOrEmpty(preset.FilePath))
            _reshadeService.DeletePresetFile(preset);
        else
            _presetService.DeletePreset(preset);

        Presets.Remove(preset);
        if (SelectedPreset == preset)
            SelectedPreset = Presets.FirstOrDefault();
        StatusMessage = $"Preset '{name}' geloescht";
    }

    [RelayCommand] private void ApplyPreset()
    {
        if (SelectedPreset == null) return;
        SyncToSelectedPreset();
        _reshadeService.ApplyPreset(SelectedPreset);
        StatusMessage = $"Preset '{SelectedPreset.Name}' angewendet";
        ActivePresetName = SelectedPreset.Name;
    }

    [RelayCommand] private void ResetPreset()
    {
        if (_originalPresetState == null) return;
        SharpenEnabled = _originalPresetState.SharpenEnabled;
        BloomEnabled = _originalPresetState.BloomEnabled;
        VibranceEnabled = _originalPresetState.VibranceEnabled;
        TonemapEnabled = _originalPresetState.TonemapEnabled;
        SharpenStrength = _originalPresetState.SharpenStrength;
        BloomStrength = _originalPresetState.BloomStrength;
        VibranceStrength = _originalPresetState.VibranceStrength;
        Contrast = _originalPresetState.Contrast;
        Brightness = _originalPresetState.Brightness;
        StatusMessage = "Aenderungen zurueckgesetzt";
    }

    public void ExportToFile(string path)
    {
        if (SelectedPreset == null) return;
        SyncToSelectedPreset();
        _presetService.ExportPreset(SelectedPreset, path);
        StatusMessage = "Preset exportiert";
    }

    public void ImportFromFile(string path)
    {
        var preset = _presetService.ImportPreset(path);
        if (preset != null)
        {
            Presets.Add(preset);
            SelectedPreset = preset;
            StatusMessage = $"Preset '{preset.Name}' importiert";
        }
    }

    // ======================================================
    // SETTINGS PAGE
    // ======================================================

    // ======================================================
    // UPDATES PAGE
    // ======================================================

    [RelayCommand]
    private async Task CheckForUpdates()
    {
        if (IsCheckingUpdates) return;
        IsCheckingUpdates = true;
        UpdateStatusText = "CHECKING...";
        StatusMessage = "Suche nach Updates...";

        try
        {
            var hasUpdate = await _updateService.CheckForUpdatesAsync(_configService.CurrentVersion);

            // Fetch releases for chronology
            var releases = await _updateService.FetchReleasesAsync();
            Releases.Clear();
            foreach (var r in releases)
                Releases.Add(r);

            CurrentVersionDisplay = $"v{_configService.CurrentVersion}";
            LastUpdateCheck = DateTime.Now.ToString("dd MMM yyyy - HH:mm");

            if (hasUpdate)
            {
                LatestVersionDisplay = $"v{_updateService.LatestVersion}";
                UpdateStatusText = $"UPDATE AVAILABLE: {LatestVersionDisplay}";
                IsUpdateAvailable = true;
                StatusMessage = $"Update verfuegbar: {LatestVersionDisplay}";
                ShowToast("UPDATE AVAILABLE", $"Version {_updateService.LatestVersion} is ready to download.");
            }
            else
            {
                UpdateStatusText = "SYSTEMS UP TO DATE";
                IsUpdateAvailable = false;
                StatusMessage = "System ist aktuell";
                ShowToast("SYNC COMPLETE", "All remote repositories are up to date.");
            }
        }
        catch (Exception ex)
        {
            UpdateStatusText = "CHECK FAILED";
            StatusMessage = $"Update-Check fehlgeschlagen: {ex.Message}";
            ShowToast("ERROR", "Could not reach update server.");
        }
        finally
        {
            IsCheckingUpdates = false;
        }
    }

    [RelayCommand]
    private async Task DownloadUpdate()
    {
        if (IsDownloading || !IsUpdateAvailable) return;
        IsDownloading = true;
        DownloadProgress = 0;
        DownloadStatusText = "Downloading...";
        StatusMessage = "Update wird heruntergeladen...";

        try
        {
            var progress = new Progress<double>(p =>
                Dispatcher.UIThread.Post(() =>
                {
                    DownloadProgress = p;
                    DownloadStatusText = $"Downloading... {(int)(p * 100)}%";
                }));

            var zipPath = await _updateService.DownloadUpdateAsync(progress);

            DownloadStatusText = "Extracting...";

            var extractDir = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "GameCopilotUpdate", "extracted");
            if (System.IO.Directory.Exists(extractDir))
                System.IO.Directory.Delete(extractDir, true);
            UpdateService.ExtractUpdate(zipPath, extractDir);

            DownloadProgress = 1.0;
            DownloadStatusText = "Installing...";

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                // Find the extracted exe (inside MSFS24.Game.Manager.Update/ folder)
                var appDir = AppDomain.CurrentDomain.BaseDirectory;
                var currentExe = Environment.ProcessPath
                    ?? System.IO.Path.Combine(appDir, "MSFS Mod Manager.exe");
                var exeName = System.IO.Path.GetFileName(currentExe);

                // Look for new exe in extracted folder (may be in subfolder)
                string? newExe = null;
                foreach (var f in System.IO.Directory.GetFiles(extractDir, "*.exe",
                    System.IO.SearchOption.AllDirectories))
                {
                    newExe = f;
                    break;
                }

                if (newExe == null)
                {
                    DownloadStatusText = "Error: No exe in update";
                    return;
                }

                // Create batch script: wait for app to close, copy new exe, restart
                var batchPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), "GameCopilotUpdate", "update.bat");
                var batchContent = $"""
                    @echo off
                    echo Warte auf Beendigung...
                    timeout /t 2 /nobreak >nul
                    echo Installiere Update...
                    copy /Y "{newExe}" "{currentExe}"
                    echo Update abgeschlossen. Starte neu...
                    start "" "{currentExe}"
                    del "%~f0"
                    """;
                await System.IO.File.WriteAllTextAsync(batchPath, batchContent);

                // Launch updater batch and exit
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"{batchPath}\"",
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };
                System.Diagnostics.Process.Start(psi);

                // Close the app so the exe can be replaced
                if (Avalonia.Application.Current?.ApplicationLifetime
                    is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
                {
                    desktop.Shutdown();
                }
            }
            else
            {
                DownloadStatusText = $"Update {LatestVersionDisplay} downloaded";
                StatusMessage = $"Update entpackt nach: {extractDir}";
                ShowToast("DOWNLOAD COMPLETE", "Update extracted. Please replace manually.");
            }

            IsUpdateAvailable = false;
            UpdateStatusText = "UPDATE INSTALLED";
        }
        catch (Exception ex)
        {
            DownloadStatusText = "Download failed";
            StatusMessage = $"Download fehlgeschlagen: {ex.Message}";
            ShowToast("DOWNLOAD FAILED", ex.Message);
        }
        finally
        {
            IsDownloading = false;
        }
    }

    private async void ShowToast(string title, string message)
    {
        // async void: must catch all exceptions or they become unhandled and crash the app.
        try
        {
            ToastTitle = title;
            ToastMessage = message;
            ToastVisible = true;
            await Task.Delay(4000);
            ToastVisible = false;
        }
        catch (OperationCanceledException) { /* app shutting down — no-op */ }
        catch { /* swallow any other unexpected error in a fire-and-forget UI method */ }
    }

    // ======================================================
    // SETTINGS PAGE
    // ======================================================

    [RelayCommand]
    private async Task RestartMcp()
    {
        McpStatus = "MCP Server wird neugestartet...";
        await _mcpService.RestartAsync(progress =>
            Dispatcher.UIThread.Post(() => McpStatus = progress));
        McpStatus = _mcpService.Status;
    }

    [RelayCommand]
    private void SaveSettings()
    {
        _configService.CommunityPath = SettingsCommunityPath;
        _configService.MsfsGamePath = SettingsMsfsGamePath;
        _configService.SteamVrPath = SettingsSteamVrPath;
        _configService.SteamExePath = SettingsSteamExePath;
        _configService.PimaxClientPath = SettingsPimaxClientPath;
        _configService.MsfsAppId = SettingsMsfsAppId;
        _configService.Save();
        StatusMessage = "Einstellungen gespeichert";

        // Reload mods with new path
        LoadAllMods();

        // Re-detect ReShade with new game path and reload presets
        _reshadeService.DetectPaths();
        DetectEnvironment();
        LoadPresets();
    }
}
