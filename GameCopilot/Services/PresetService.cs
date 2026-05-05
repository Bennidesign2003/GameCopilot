using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using GameCopilot.Models;

namespace GameCopilot.Services;

/// <summary>
/// Handles loading and saving ReShade presets as JSON files.
/// Migrated from WPF AppConfig pattern, adapted for preset management.
/// </summary>
public class PresetService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _presetsDirectory;

    public PresetService()
    {
        _presetsDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GameCopilot",
            "Presets"
        );

        if (!Directory.Exists(_presetsDirectory))
            Directory.CreateDirectory(_presetsDirectory);
    }

    public List<ReShadePreset> LoadAllPresets()
    {
        var presets = new List<ReShadePreset>();

        if (!Directory.Exists(_presetsDirectory))
            return GetDefaultPresets();

        foreach (var file in Directory.GetFiles(_presetsDirectory, "*.json"))
        {
            try
            {
                var json = File.ReadAllText(file);
                var preset = JsonSerializer.Deserialize<ReShadePreset>(json, JsonOptions);
                if (preset != null)
                    presets.Add(preset);
            }
            catch
            {
                // Skip corrupt files
            }
        }

        if (presets.Count == 0)
        {
            presets = GetDefaultPresets();
            foreach (var p in presets)
                SavePreset(p);
        }

        return presets;
    }

    public void SavePreset(ReShadePreset preset)
    {
        var fileName = SanitizeFileName(preset.Name) + ".json";
        var filePath = Path.Combine(_presetsDirectory, fileName);
        var json = JsonSerializer.Serialize(preset, JsonOptions);
        File.WriteAllText(filePath, json);
    }

    public void DeletePreset(ReShadePreset preset)
    {
        var fileName = SanitizeFileName(preset.Name) + ".json";
        var filePath = Path.Combine(_presetsDirectory, fileName);
        if (File.Exists(filePath))
            File.Delete(filePath);
    }

    public string ExportPreset(ReShadePreset preset, string targetPath)
    {
        var json = JsonSerializer.Serialize(preset, JsonOptions);
        File.WriteAllText(targetPath, json);
        return targetPath;
    }

    public ReShadePreset? ImportPreset(string filePath)
    {
        if (!File.Exists(filePath)) return null;

        var json = File.ReadAllText(filePath);
        var preset = JsonSerializer.Deserialize<ReShadePreset>(json, JsonOptions);
        if (preset != null)
            SavePreset(preset);

        return preset;
    }

    public static List<ReShadePreset> GetDefaultPresets()
    {
        return new List<ReShadePreset>
        {
            new()
            {
                Name = "VR Smooth",
                Description = "Optimiert fuer VR mit weichen Uebergaengen",
                SharpenEnabled = true,
                BloomEnabled = false,
                VibranceEnabled = true,
                TonemapEnabled = true,
                SharpenStrength = 0.3,
                BloomStrength = 0.0,
                VibranceStrength = 0.4,
                Contrast = 0.45,
                Brightness = 0.55,
            },
            new()
            {
                Name = "VR Balanced",
                Description = "Ausgewogenes VR-Preset fuer Schaerfe und Performance",
                SharpenEnabled = true,
                BloomEnabled = true,
                VibranceEnabled = true,
                TonemapEnabled = true,
                SharpenStrength = 0.5,
                BloomStrength = 0.2,
                VibranceStrength = 0.5,
                Contrast = 0.5,
                Brightness = 0.5,
            },
            new()
            {
                Name = "VR Quality",
                Description = "Maximale Bildqualitaet fuer VR",
                SharpenEnabled = true,
                BloomEnabled = true,
                VibranceEnabled = true,
                TonemapEnabled = true,
                SharpenStrength = 0.7,
                BloomStrength = 0.4,
                VibranceStrength = 0.6,
                Contrast = 0.55,
                Brightness = 0.5,
            },
            new()
            {
                Name = "Night Flight",
                Description = "Optimiert fuer Nachtfluege mit reduzierter Helligkeit",
                SharpenEnabled = true,
                BloomEnabled = true,
                VibranceEnabled = false,
                TonemapEnabled = true,
                SharpenStrength = 0.4,
                BloomStrength = 0.6,
                VibranceStrength = 0.3,
                Contrast = 0.6,
                Brightness = 0.35,
            },
            new()
            {
                Name = "Cinematic",
                Description = "Filmischer Look mit starkem Bloom und Kontrast",
                SharpenEnabled = false,
                BloomEnabled = true,
                VibranceEnabled = true,
                TonemapEnabled = true,
                SharpenStrength = 0.2,
                BloomStrength = 0.7,
                VibranceStrength = 0.7,
                Contrast = 0.65,
                Brightness = 0.45,
            },
        };
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
