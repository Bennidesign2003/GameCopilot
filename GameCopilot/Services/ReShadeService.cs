using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using GameCopilot.Models;

namespace GameCopilot.Services;

/// <summary>
/// Handles reading/writing ReShade .ini preset files and the main ReShade.ini config.
/// Detects ReShade installation, parses presets, and writes changes back losslessly.
/// </summary>
public class ReShadeService
{
    private readonly AppConfigService _config;

    public string? MsfsConfigPath { get; private set; }
    public string? MsfsGamePath { get; private set; }
    public string? ReshadePath { get; private set; }
    public bool IsReshadeFound { get; private set; }
    public string? ActivePresetPath { get; private set; }
    public string? PresetsDirectory { get; private set; }

    // Known shader file patterns for technique detection
    private static readonly string[] SharpenFiles = { "CAS.fx", "LumaSharpen.fx", "AdaptiveSharpen.fx", "FilmicSharpen.fx" };
    private static readonly string[] BloomFiles = { "MagicBloom.fx", "Bloom.fx", "HDRBloom.fx", "qUINT_bloom.fx" };
    private static readonly string[] VibranceFiles = { "Vibrance.fx", "ColorVibrancy.fx" };
    private static readonly string[] TonemapFiles = { "Tonemap.fx", "FakeHDR.fx", "LiftGammaGain.fx", "DPX.fx" };

    // Known parameter mappings: (sectionContains, paramKey) → model property name
    private static readonly (string filePattern, string paramKey, string modelProp)[] ParamMappings =
    {
        ("CAS", "Sharpness", "SharpenStrength"),
        ("LumaSharpen", "sharp_strength", "SharpenStrength"),
        ("AdaptiveSharpen", "Sharpening", "SharpenStrength"),
        ("FilmicSharpen", "Strength", "SharpenStrength"),
        ("MagicBloom", "fBloom_Intensity", "BloomStrength"),
        ("Bloom", "BloomIntensity", "BloomStrength"),
        ("Bloom", "fBloom_Intensity", "BloomStrength"),
        ("Vibrance", "Vibrance", "VibranceStrength"),
        ("Tonemap", "Gamma", "Brightness"),
        ("Tonemap", "Exposure", "Contrast"),
        ("FakeHDR", "Power", "Contrast"),
        ("FakeHDR", "HDRPower", "Contrast"),
        ("LiftGammaGain", "RGB_Gamma", "Brightness"),
        ("LiftGammaGain", "RGB_Lift", "Contrast"),
    };

    public ReShadeService(AppConfigService config)
    {
        _config = config;
        DetectPaths();
    }

    public void DetectPaths()
    {
        MsfsConfigPath = null;
        MsfsGamePath = null;
        ReshadePath = null;
        IsReshadeFound = false;
        ActivePresetPath = null;
        PresetsDirectory = null;

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            MsfsConfigPath = "(nur unter Windows verfuegbar)";
            return;
        }

        // 1. MSFS config path (%APPDATA%)
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var msfsConfig = Path.Combine(appData, "Microsoft Flight Simulator 2024");
        if (Directory.Exists(msfsConfig))
            MsfsConfigPath = msfsConfig;

        // 2. MSFS game path (where ReShade lives)
        // Check user-configured path first
        if (!string.IsNullOrEmpty(_config.MsfsGamePath) && Directory.Exists(_config.MsfsGamePath))
        {
            MsfsGamePath = _config.MsfsGamePath;
        }
        else
        {
            // Auto-detect Steam installation
            string[] possiblePaths =
            {
                @"C:\Program Files (x86)\Steam\steamapps\common\MicrosoftFlightSimulator2024",
                @"D:\Steam\steamapps\common\MicrosoftFlightSimulator2024",
                @"E:\Steam\steamapps\common\MicrosoftFlightSimulator2024",
                @"D:\SteamLibrary\steamapps\common\MicrosoftFlightSimulator2024",
                @"E:\SteamLibrary\steamapps\common\MicrosoftFlightSimulator2024",
            };

            foreach (var path in possiblePaths)
            {
                if (Directory.Exists(path))
                {
                    MsfsGamePath = path;
                    break;
                }
            }
        }

        // 3. Detect ReShade in game directory
        if (MsfsGamePath != null)
        {
            var reshadeIni = Path.Combine(MsfsGamePath, "ReShade.ini");
            var reshadeDll = Path.Combine(MsfsGamePath, "dxgi.dll");

            if (File.Exists(reshadeIni) || File.Exists(reshadeDll))
            {
                ReshadePath = MsfsGamePath;
                IsReshadeFound = true;

                // Parse ReShade.ini for active preset and preset directory
                if (File.Exists(reshadeIni))
                    ParseReshadeIni(reshadeIni);
            }
        }
    }

    /// <summary>
    /// Parses the main ReShade.ini to find the active preset path and preset directories.
    /// </summary>
    private void ParseReshadeIni(string reshadeIniPath)
    {
        try
        {
            var lines = File.ReadAllLines(reshadeIniPath);
            var gameDir = Path.GetDirectoryName(reshadeIniPath)!;

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("PresetPath=", StringComparison.OrdinalIgnoreCase))
                {
                    var presetPath = trimmed.Substring("PresetPath=".Length).Trim();
                    if (!Path.IsPathRooted(presetPath))
                        presetPath = Path.GetFullPath(Path.Combine(gameDir, presetPath));
                    ActivePresetPath = presetPath;

                    // Use the directory of the active preset as the presets directory
                    var dir = Path.GetDirectoryName(presetPath);
                    if (dir != null && Directory.Exists(dir))
                        PresetsDirectory = dir;
                }
            }

            // Also check common preset subdirectories
            if (PresetsDirectory == null)
            {
                string[] presetDirs =
                {
                    Path.Combine(gameDir, "reshade-presets"),
                    Path.Combine(gameDir, "ReShadePresets"),
                    Path.Combine(gameDir, "Presets"),
                };
                foreach (var dir in presetDirs)
                {
                    if (Directory.Exists(dir))
                    {
                        PresetsDirectory = dir;
                        break;
                    }
                }
            }

            // Fallback: game directory itself
            PresetsDirectory ??= gameDir;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Lesen von ReShade.ini: {ex.Message}");
        }
    }

    /// <summary>
    /// Finds all .ini preset files in the presets directory.
    /// </summary>
    public List<ReShadePreset> LoadAllPresets()
    {
        var presets = new List<ReShadePreset>();

        if (PresetsDirectory == null || !Directory.Exists(PresetsDirectory))
            return presets;

        foreach (var file in Directory.GetFiles(PresetsDirectory, "*.ini"))
        {
            // Skip ReShade.ini itself
            if (Path.GetFileName(file).Equals("ReShade.ini", StringComparison.OrdinalIgnoreCase))
                continue;
            // Skip ReShadePreset.ini (internal ReShade file)
            if (Path.GetFileName(file).Equals("ReShadePreset.ini", StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                var preset = ParsePresetFile(file);
                if (preset != null)
                    presets.Add(preset);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Fehler beim Lesen von Preset '{file}': {ex.Message}");
            }
        }

        return presets;
    }

    /// <summary>
    /// Parses a ReShade preset .ini file into a ReShadePreset model.
    /// </summary>
    public ReShadePreset? ParsePresetFile(string filePath)
    {
        if (!File.Exists(filePath)) return null;

        var lines = File.ReadAllLines(filePath);
        var preset = new ReShadePreset
        {
            Name = Path.GetFileNameWithoutExtension(filePath),
            FilePath = filePath,
            Description = GenerateDescription(filePath),
        };

        string currentSection = "";

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith(';'))
                continue;

            // Section header
            if (trimmed.StartsWith('[') && trimmed.EndsWith(']'))
            {
                currentSection = trimmed.Substring(1, trimmed.Length - 2);
                if (!preset.Sections.ContainsKey(currentSection))
                    preset.Sections[currentSection] = new Dictionary<string, string>();
                continue;
            }

            var eqIdx = trimmed.IndexOf('=');
            if (eqIdx < 0) continue;

            var key = trimmed.Substring(0, eqIdx).Trim();
            var value = trimmed.Substring(eqIdx + 1).Trim();

            if (string.IsNullOrEmpty(currentSection))
            {
                // Global entries
                if (key.Equals("Techniques", StringComparison.OrdinalIgnoreCase))
                {
                    preset.EnabledTechniques = value
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .ToList();
                }
                else if (key.Equals("TechniqueSorting", StringComparison.OrdinalIgnoreCase))
                {
                    preset.AllTechniques = value
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .ToList();
                }
                else if (key.Equals("PreprocessorDefinitions", StringComparison.OrdinalIgnoreCase))
                {
                    preset.PreprocessorDefinitions = value;
                }
            }
            else
            {
                preset.Sections[currentSection][key] = value;
            }
        }

        // Map techniques to toggle properties
        MapTechniquesToToggles(preset);
        // Map parameters to slider properties
        MapParametersToSliders(preset);

        return preset;
    }

    /// <summary>
    /// Maps enabled techniques to our 4 toggle properties.
    /// </summary>
    private static void MapTechniquesToToggles(ReShadePreset preset)
    {
        foreach (var tech in preset.EnabledTechniques)
        {
            // Technique format: "TechniqueName@File.fx"
            var atIdx = tech.IndexOf('@');
            var file = atIdx >= 0 ? tech.Substring(atIdx + 1) : tech;

            if (preset.SharpenFile == null && MatchesAny(file, SharpenFiles))
            {
                preset.SharpenEnabled = true;
                preset.SharpenFile = file;
            }
            else if (preset.BloomFile == null && MatchesAny(file, BloomFiles))
            {
                preset.BloomEnabled = true;
                preset.BloomFile = file;
            }
            else if (preset.VibranceFile == null && MatchesAny(file, VibranceFiles))
            {
                preset.VibranceEnabled = true;
                preset.VibranceFile = file;
            }
            else if (preset.TonemapFile == null && MatchesAny(file, TonemapFiles))
            {
                preset.TonemapEnabled = true;
                preset.TonemapFile = file;
            }
        }

        // Also check AllTechniques for files (they exist but might not be enabled)
        foreach (var tech in preset.AllTechniques)
        {
            var atIdx = tech.IndexOf('@');
            var file = atIdx >= 0 ? tech.Substring(atIdx + 1) : tech;

            if (preset.SharpenFile == null && MatchesAny(file, SharpenFiles))
                preset.SharpenFile = file;
            else if (preset.BloomFile == null && MatchesAny(file, BloomFiles))
                preset.BloomFile = file;
            else if (preset.VibranceFile == null && MatchesAny(file, VibranceFiles))
                preset.VibranceFile = file;
            else if (preset.TonemapFile == null && MatchesAny(file, TonemapFiles))
                preset.TonemapFile = file;
        }
    }

    /// <summary>
    /// Maps .ini section parameters to our 5 slider properties.
    /// </summary>
    private static void MapParametersToSliders(ReShadePreset preset)
    {
        foreach (var (section, entries) in preset.Sections)
        {
            foreach (var (key, value) in entries)
            {
                foreach (var (filePattern, paramKey, modelProp) in ParamMappings)
                {
                    if (!section.Contains(filePattern, StringComparison.OrdinalIgnoreCase) ||
                        !key.Equals(paramKey, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (!TryParseDouble(value, out var numValue))
                        continue;

                    switch (modelProp)
                    {
                        case "SharpenStrength":
                            if (preset.SharpenParamKey != null) break;
                            preset.SharpenStrength = ClampTo01(numValue);
                            preset.SharpenParamKey = $"{section}.{key}";
                            break;
                        case "BloomStrength":
                            if (preset.BloomParamKey != null) break;
                            preset.BloomStrength = ClampTo01(numValue);
                            preset.BloomParamKey = $"{section}.{key}";
                            break;
                        case "VibranceStrength":
                            if (preset.VibranceParamKey != null) break;
                            // Vibrance range is often -1 to 1, normalize to 0-1
                            preset.VibranceStrength = ClampTo01((numValue + 1.0) / 2.0);
                            preset.VibranceParamKey = $"{section}.{key}";
                            break;
                        case "Contrast":
                            if (preset.ContrastParamKey != null) break;
                            preset.Contrast = ClampTo01(numValue);
                            preset.ContrastParamKey = $"{section}.{key}";
                            break;
                        case "Brightness":
                            if (preset.BrightnessParamKey != null) break;
                            preset.Brightness = ClampTo01(numValue);
                            preset.BrightnessParamKey = $"{section}.{key}";
                            break;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Writes a preset back to its .ini file, applying UI changes to the known parameters.
    /// Preserves all unknown sections and parameters for lossless round-tripping.
    /// </summary>
    public bool SavePresetFile(ReShadePreset preset)
    {
        if (string.IsNullOrEmpty(preset.FilePath))
            return false;

        try
        {
            // Apply UI toggle changes to technique lists
            ApplyTogglesToTechniques(preset);

            // Apply UI slider changes to section parameters
            ApplySliderToSection(preset, preset.SharpenParamKey, preset.SharpenStrength);
            ApplySliderToSection(preset, preset.BloomParamKey, preset.BloomStrength);
            // Vibrance: denormalize from 0-1 back to -1..1
            if (preset.VibranceParamKey != null)
                ApplySliderToSection(preset, preset.VibranceParamKey, preset.VibranceStrength * 2.0 - 1.0);
            ApplySliderToSection(preset, preset.ContrastParamKey, preset.Contrast);
            ApplySliderToSection(preset, preset.BrightnessParamKey, preset.Brightness);

            // Reconstruct the .ini file
            var lines = new List<string>();

            // Global entries
            lines.Add($"PreprocessorDefinitions={preset.PreprocessorDefinitions}");
            lines.Add($"Techniques={string.Join(",", preset.EnabledTechniques)}");
            lines.Add($"TechniqueSorting={string.Join(",", preset.AllTechniques)}");
            lines.Add("");

            // Sections
            foreach (var (section, entries) in preset.Sections)
            {
                lines.Add($"[{section}]");
                foreach (var (key, value) in entries)
                    lines.Add($"{key}={value}");
                lines.Add("");
            }

            File.WriteAllLines(preset.FilePath, lines);
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Speichern: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Applies a preset by saving it and setting it as active in ReShade.ini.
    /// </summary>
    public bool ApplyPreset(ReShadePreset preset)
    {
        if (!SavePresetFile(preset))
            return false;

        // Set as active preset in ReShade.ini
        SetActivePreset(preset.FilePath);
        return true;
    }

    /// <summary>
    /// Sets the active preset path in ReShade.ini.
    /// </summary>
    public void SetActivePreset(string presetPath)
    {
        if (ReshadePath == null) return;

        var reshadeIni = Path.Combine(ReshadePath, "ReShade.ini");
        if (!File.Exists(reshadeIni)) return;

        try
        {
            var lines = File.ReadAllLines(reshadeIni).ToList();
            bool found = false;

            // Make path relative to game directory if possible
            var relativePath = presetPath;
            if (presetPath.StartsWith(ReshadePath, StringComparison.OrdinalIgnoreCase))
                relativePath = ".\\" + presetPath.Substring(ReshadePath.Length).TrimStart('\\', '/');

            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].TrimStart().StartsWith("PresetPath=", StringComparison.OrdinalIgnoreCase))
                {
                    lines[i] = $"PresetPath={relativePath}";
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                // Add under [GENERAL] section
                for (int i = 0; i < lines.Count; i++)
                {
                    if (lines[i].Trim().Equals("[GENERAL]", StringComparison.OrdinalIgnoreCase))
                    {
                        lines.Insert(i + 1, $"PresetPath={relativePath}");
                        found = true;
                        break;
                    }
                }
            }

            if (found)
            {
                File.WriteAllLines(reshadeIni, lines);
                ActivePresetPath = presetPath;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Setzen des aktiven Presets: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new preset .ini file with current slider/toggle values.
    /// </summary>
    public ReShadePreset? CreateNewPreset(string name, ReShadePreset? basedOn = null)
    {
        if (PresetsDirectory == null) return null;

        var filePath = Path.Combine(PresetsDirectory, SanitizeFileName(name) + ".ini");

        var preset = basedOn?.Clone() ?? new ReShadePreset();
        preset.Name = name;
        preset.FilePath = filePath;
        preset.Description = "Benutzerdefiniertes Preset";

        SavePresetFile(preset);
        return preset;
    }

    /// <summary>
    /// Deletes a preset .ini file from disk.
    /// </summary>
    public bool DeletePresetFile(ReShadePreset preset)
    {
        if (string.IsNullOrEmpty(preset.FilePath) || !File.Exists(preset.FilePath))
            return false;

        try
        {
            File.Delete(preset.FilePath);
            return true;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Loeschen: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Reads the currently active preset (the one referenced in ReShade.ini).
    /// </summary>
    public ReShadePreset? ReadCurrentConfig()
    {
        if (ActivePresetPath == null || !File.Exists(ActivePresetPath))
            return null;
        return ParsePresetFile(ActivePresetPath);
    }

    // ===== Helper methods =====

    private static void ApplyTogglesToTechniques(ReShadePreset preset)
    {
        UpdateTechniqueEnabled(preset, preset.SharpenFile, preset.SharpenEnabled);
        UpdateTechniqueEnabled(preset, preset.BloomFile, preset.BloomEnabled);
        UpdateTechniqueEnabled(preset, preset.VibranceFile, preset.VibranceEnabled);
        UpdateTechniqueEnabled(preset, preset.TonemapFile, preset.TonemapEnabled);
    }

    private static void UpdateTechniqueEnabled(ReShadePreset preset, string? shaderFile, bool enabled)
    {
        if (shaderFile == null) return;

        // Find the technique entry that references this shader file
        var technique = preset.AllTechniques.FirstOrDefault(t =>
            t.EndsWith("@" + shaderFile, StringComparison.OrdinalIgnoreCase) ||
            t.Equals(shaderFile, StringComparison.OrdinalIgnoreCase));

        if (technique == null) return;

        bool isCurrentlyEnabled = preset.EnabledTechniques.Any(t =>
            t.Equals(technique, StringComparison.OrdinalIgnoreCase));

        if (enabled && !isCurrentlyEnabled)
            preset.EnabledTechniques.Add(technique);
        else if (!enabled && isCurrentlyEnabled)
            preset.EnabledTechniques.RemoveAll(t =>
                t.Equals(technique, StringComparison.OrdinalIgnoreCase));
    }

    private static void ApplySliderToSection(ReShadePreset preset, string? paramKey, double value)
    {
        if (paramKey == null) return;

        var dotIdx = paramKey.IndexOf('.');
        if (dotIdx < 0) return;

        var section = paramKey.Substring(0, dotIdx);
        var key = paramKey.Substring(dotIdx + 1);

        if (preset.Sections.TryGetValue(section, out var entries))
            entries[key] = value.ToString("F6", CultureInfo.InvariantCulture);
    }

    private static bool MatchesAny(string file, string[] patterns)
    {
        return patterns.Any(p => file.Equals(p, StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryParseDouble(string value, out double result)
    {
        // ReShade uses invariant culture (e.g., "0.700000")
        // Handle comma-separated vectors by taking first value
        var firstValue = value.Split(',')[0].Trim();
        return double.TryParse(firstValue, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
    }

    private static double ClampTo01(double value)
    {
        return Math.Max(0.0, Math.Min(1.0, value));
    }

    private static string GenerateDescription(string filePath)
    {
        var name = Path.GetFileNameWithoutExtension(filePath);
        var dir = Path.GetDirectoryName(filePath);
        var dirName = dir != null ? Path.GetFileName(dir) : "";
        return $"ReShade Preset aus {dirName}";
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
