using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace GameCopilot.Services;

/// <summary>
/// Migrated from WPF AppConfig.cs + settings.config pattern.
/// Loads/saves app settings as JSON. Replaces the key=value config from WPF.
/// </summary>
public class AppConfigService
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "GameCopilot"
    );

    private static readonly string ConfigPath = Path.Combine(ConfigDir, "appconfig.json");

    public string CurrentVersion =>
        System.Reflection.Assembly.GetExecutingAssembly()
            .GetName().Version?.ToString(3) ?? "3.5.3";
    public string CommunityPath { get; set; } = "";
    public string MsfsGamePath { get; set; } = "";
    public string SteamVrPath { get; set; } = @"C:\Program Files (x86)\Steam\steamapps\common\SteamVR\bin\win64\vrstartup.exe";
    public string SteamExePath { get; set; } = @"C:\Program Files (x86)\Steam\Steam.exe";
    public string PimaxClientPath { get; set; } = @"C:\Program Files\Pimax\PimaxClient\pimaxui\PimaxClient.exe";
    public string PimaxOpenXrJson { get; set; } = @"C:\Program Files (x86)\Steam\steamapps\common\SteamVR\steamxr_win64.json";
    public string MsfsAppId { get; set; } = "2537590";

    // ── AI Provider settings ──────────────────────────────────────────────────
    /// <summary>OpenAI API key for the Codex/GPT cloud provider.</summary>
    public string CodexApiKey { get; set; } = "";
    /// <summary>Last-used Codex model id.</summary>
    public string CodexModel { get; set; } = "codex-mini-latest";
    /// <summary>"local" = Ollama; "codex" = OpenAI Codex/GPT cloud.</summary>
    public string ChatProvider { get; set; } = "local";
    /// <summary>Ollama API base URL.</summary>
    public string OllamaUrl { get; set; } = "http://localhost:11434";

    // ── Appearance / behaviour ────────────────────────────────────────────────
    /// <summary>Pre-warm the Ollama / MCP server at launch (faster first chat).</summary>
    public bool WarmOnStartup { get; set; } = true;
    /// <summary>Load the previous session's chat messages on startup.</summary>
    public bool LoadChatHistory { get; set; } = true;

    public AppConfigService()
    {
        if (!Directory.Exists(ConfigDir))
            Directory.CreateDirectory(ConfigDir);

        // Default community path (same as WPF ModsPage default)
        if (string.IsNullOrEmpty(CommunityPath))
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            CommunityPath = Path.Combine(appData, "Microsoft Flight Simulator 2024", "Packages", "Community");
        }
    }

    public void Load()
    {
        if (!File.Exists(ConfigPath)) return;

        try
        {
            var json = File.ReadAllText(ConfigPath);
            var data = JsonSerializer.Deserialize<ConfigData>(json);
            if (data == null) return;

            CommunityPath = data.CommunityPath ?? CommunityPath;
            MsfsGamePath = data.MsfsGamePath ?? MsfsGamePath;
            SteamVrPath = data.SteamVrPath ?? SteamVrPath;
            SteamExePath = data.SteamExePath ?? SteamExePath;
            PimaxClientPath = data.PimaxClientPath ?? PimaxClientPath;
            PimaxOpenXrJson = data.PimaxOpenXrJson ?? PimaxOpenXrJson;
            MsfsAppId = data.MsfsAppId ?? MsfsAppId;
            CodexApiKey     = data.CodexApiKey     ?? CodexApiKey;
            CodexModel      = data.CodexModel      ?? CodexModel;
            ChatProvider    = data.ChatProvider    ?? ChatProvider;
            OllamaUrl       = data.OllamaUrl       ?? OllamaUrl;
            WarmOnStartup   = data.WarmOnStartup   ?? WarmOnStartup;
            LoadChatHistory = data.LoadChatHistory ?? LoadChatHistory;
        }
        catch
        {
            // Corrupt config, use defaults
        }
    }

    public void Save()
    {
        try
        {
            var data = new ConfigData
            {
                CommunityPath = CommunityPath,
                MsfsGamePath = MsfsGamePath,
                SteamVrPath = SteamVrPath,
                SteamExePath = SteamExePath,
                PimaxClientPath = PimaxClientPath,
                PimaxOpenXrJson = PimaxOpenXrJson,
                MsfsAppId = MsfsAppId,
                CodexApiKey     = CodexApiKey,
                CodexModel      = CodexModel,
                ChatProvider    = ChatProvider,
                OllamaUrl       = OllamaUrl,
                WarmOnStartup   = WarmOnStartup,
                LoadChatHistory = LoadChatHistory,
            };

            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            // Write atomically: write to temp file then rename so a crash during write
            // never leaves a truncated/empty config.
            var tmp = ConfigPath + ".tmp";
            File.WriteAllText(tmp, json);
            File.Move(tmp, ConfigPath, overwrite: true);
        }
        catch (Exception ex)
        {
            // Non-fatal: log to Debug output but don't crash the app.
            System.Diagnostics.Debug.WriteLine($"[AppConfigService] Save failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Migrated from WPF ModsPage.GetSetting() key=value pattern.
    /// Now uses typed properties instead of string lookup.
    /// </summary>
    private class ConfigData
    {
        public string? CommunityPath { get; set; }
        public string? MsfsGamePath { get; set; }
        public string? SteamVrPath { get; set; }
        public string? SteamExePath { get; set; }
        public string? PimaxClientPath { get; set; }
        public string? PimaxOpenXrJson { get; set; }
        public string? MsfsAppId { get; set; }
        public string? CodexApiKey     { get; set; }
        public string? CodexModel      { get; set; }
        public string? ChatProvider    { get; set; }
        public string? OllamaUrl       { get; set; }
        public bool?   WarmOnStartup   { get; set; }
        public bool?   LoadChatHistory { get; set; }
    }
}
