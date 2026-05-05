using System.Collections.Generic;

namespace GameCopilot.Models;

/// <summary>
/// Represents a ReShade preset parsed from a .ini file.
/// Stores both mapped UI properties and raw .ini data for lossless round-tripping.
/// </summary>
public class ReShadePreset
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;

    /// <summary>Path to the .ini preset file on disk (empty for JSON-only presets).</summary>
    public string FilePath { get; set; } = string.Empty;

    // Shader toggles (mapped from Techniques= line in .ini)
    public bool SharpenEnabled { get; set; }
    public bool BloomEnabled { get; set; }
    public bool VibranceEnabled { get; set; }
    public bool TonemapEnabled { get; set; }

    // Shader parameters (mapped from known .ini sections, 0.0 - 1.0 UI range)
    public double SharpenStrength { get; set; } = 0.5;
    public double BloomStrength { get; set; } = 0.3;
    public double VibranceStrength { get; set; } = 0.5;
    public double Contrast { get; set; } = 0.5;
    public double Brightness { get; set; } = 0.5;

    // Raw .ini data for lossless round-tripping
    /// <summary>All enabled techniques from the Techniques= line.</summary>
    public List<string> EnabledTechniques { get; set; } = new();
    /// <summary>All techniques (sort order) from TechniqueSorting= line.</summary>
    public List<string> AllTechniques { get; set; } = new();
    /// <summary>PreprocessorDefinitions= line value.</summary>
    public string PreprocessorDefinitions { get; set; } = string.Empty;
    /// <summary>Parsed sections: section name → (key → value).</summary>
    public Dictionary<string, Dictionary<string, string>> Sections { get; set; } = new();

    // Tracking which shader files mapped to our UI properties
    public string? SharpenFile { get; set; }
    public string? BloomFile { get; set; }
    public string? VibranceFile { get; set; }
    public string? TonemapFile { get; set; }
    public string? SharpenParamKey { get; set; }
    public string? BloomParamKey { get; set; }
    public string? VibranceParamKey { get; set; }
    public string? ContrastParamKey { get; set; }
    public string? BrightnessParamKey { get; set; }

    public ReShadePreset Clone()
    {
        return new ReShadePreset
        {
            Name = Name,
            Description = Description,
            FilePath = FilePath,
            SharpenEnabled = SharpenEnabled,
            BloomEnabled = BloomEnabled,
            VibranceEnabled = VibranceEnabled,
            TonemapEnabled = TonemapEnabled,
            SharpenStrength = SharpenStrength,
            BloomStrength = BloomStrength,
            VibranceStrength = VibranceStrength,
            Contrast = Contrast,
            Brightness = Brightness,
            EnabledTechniques = new List<string>(EnabledTechniques),
            AllTechniques = new List<string>(AllTechniques),
            PreprocessorDefinitions = PreprocessorDefinitions,
            Sections = CloneSections(Sections),
            SharpenFile = SharpenFile,
            BloomFile = BloomFile,
            VibranceFile = VibranceFile,
            TonemapFile = TonemapFile,
            SharpenParamKey = SharpenParamKey,
            BloomParamKey = BloomParamKey,
            VibranceParamKey = VibranceParamKey,
            ContrastParamKey = ContrastParamKey,
            BrightnessParamKey = BrightnessParamKey,
        };
    }

    private static Dictionary<string, Dictionary<string, string>> CloneSections(
        Dictionary<string, Dictionary<string, string>> source)
    {
        var clone = new Dictionary<string, Dictionary<string, string>>();
        foreach (var (section, entries) in source)
            clone[section] = new Dictionary<string, string>(entries);
        return clone;
    }
}
