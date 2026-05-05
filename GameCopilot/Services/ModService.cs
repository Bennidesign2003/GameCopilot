using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using GameCopilot.Models;

namespace GameCopilot.Services;

/// <summary>
/// Migrated from WPF ModsPage.xaml.cs.
/// Handles all mod filesystem operations: scan, add, delete, rename.
/// </summary>
public class ModService
{
    private readonly AppConfigService _config;

    public ModService(AppConfigService config)
    {
        _config = config;
    }

    /// <summary>
    /// Loads all mods from the Community folder.
    /// Migrated from WPF ModsPage.LoadMods()
    /// </summary>
    public List<ModItem> LoadMods()
    {
        var mods = new List<ModItem>();
        var modsFolder = _config.CommunityPath;

        if (!Directory.Exists(modsFolder))
            return mods;

        foreach (var dir in Directory.GetDirectories(modsFolder))
        {
            mods.Add(new ModItem
            {
                Name = Path.GetFileName(dir),
                FullPath = dir
            });
        }

        return mods;
    }

    /// <summary>
    /// Deletes a mod folder.
    /// Migrated from WPF ModsPage.DeleteMod_Click()
    /// </summary>
    public bool DeleteMod(ModItem mod)
    {
        try
        {
            if (Directory.Exists(mod.FullPath))
            {
                Directory.Delete(mod.FullPath, true);
                return true;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Loeschen: {ex.Message}");
        }
        return false;
    }

    /// <summary>
    /// Renames a mod folder.
    /// Migrated from WPF ModsPage.RenameMod_Click()
    /// </summary>
    public (bool success, string newPath) RenameMod(ModItem mod, string newName)
    {
        if (string.IsNullOrWhiteSpace(newName) || newName == mod.Name)
            return (false, mod.FullPath);

        var parentFolder = Path.GetDirectoryName(mod.FullPath);
        if (parentFolder == null) return (false, mod.FullPath);

        var newPath = Path.Combine(parentFolder, newName);

        if (Directory.Exists(newPath))
            return (false, mod.FullPath);

        try
        {
            Directory.Move(mod.FullPath, newPath);
            return (true, newPath);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Umbenennen: {ex.Message}");
            return (false, mod.FullPath);
        }
    }

    /// <summary>
    /// Adds a mod from a file (ZIP or RAR archive, or plain file).
    /// Migrated from WPF ModsPage.AddMod_Click()
    /// </summary>
    public ModItem? AddMod(string sourceFile)
    {
        var modsFolder = _config.CommunityPath;

        if (!Directory.Exists(modsFolder))
            Directory.CreateDirectory(modsFolder);

        var extension = Path.GetExtension(sourceFile).ToLower();

        try
        {
            if (extension == ".zip")
            {
                var tempFolder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                Directory.CreateDirectory(tempFolder);
                ZipFile.ExtractToDirectory(sourceFile, tempFolder);
                return ProcessExtractedFolder(tempFolder, modsFolder, sourceFile);
            }
            else if (extension == ".rar")
            {
                var tempFolder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                Directory.CreateDirectory(tempFolder);
                ExtractRarArchive(sourceFile, tempFolder);
                return ProcessExtractedFolder(tempFolder, modsFolder, sourceFile);
            }
            else
            {
                var destFile = Path.Combine(modsFolder, Path.GetFileName(sourceFile));
                File.Copy(sourceFile, destFile, overwrite: true);
                return new ModItem
                {
                    Name = Path.GetFileNameWithoutExtension(destFile),
                    FullPath = destFile
                };
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Fehler beim Hinzufuegen: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Processes extracted archive folders.
    /// Migrated from WPF ModsPage.ProcessExtractedFolder()
    /// </summary>
    private static ModItem ProcessExtractedFolder(string tempFolder, string modsFolder, string sourceFile)
    {
        var directories = Directory.GetDirectories(tempFolder);
        string modFolder;

        if (directories.Length == 1)
        {
            // Single folder → move directly (from WPF)
            modFolder = Path.Combine(modsFolder, Path.GetFileName(directories[0]));
            if (Directory.Exists(modFolder))
                Directory.Delete(modFolder, true);
            Directory.Move(directories[0], modFolder);

            try { Directory.Delete(tempFolder, true); } catch { }
        }
        else
        {
            // Multiple files → move all into new folder (from WPF)
            modFolder = Path.Combine(modsFolder, Path.GetFileNameWithoutExtension(sourceFile));
            if (Directory.Exists(modFolder))
                Directory.Delete(modFolder, true);
            Directory.Move(tempFolder, modFolder);
        }

        return new ModItem
        {
            Name = Path.GetFileName(modFolder),
            FullPath = modFolder
        };
    }

    /// <summary>
    /// Extracts RAR archives using WinRAR.
    /// Migrated from WPF ModsPage.ExtractRarArchive()
    /// </summary>
    private static void ExtractRarArchive(string rarFile, string destinationFolder)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            throw new PlatformNotSupportedException("RAR-Entpacken nur unter Windows mit WinRAR moeglich.");

        // Possible WinRAR paths (from WPF)
        string[] possiblePaths =
        {
            @"C:\Program Files\WinRAR\WinRAR.exe",
            @"C:\Program Files (x86)\WinRAR\WinRAR.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"WinRAR\WinRAR.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"WinRAR\WinRAR.exe")
        };

        string? winrarPath = null;
        foreach (var path in possiblePaths)
        {
            if (File.Exists(path))
            {
                winrarPath = path;
                break;
            }
        }

        if (string.IsNullOrEmpty(winrarPath))
        {
            throw new FileNotFoundException(
                "WinRAR wurde nicht gefunden!\n\n" +
                "Bitte installiere WinRAR von: https://www.win-rar.com\n\n" +
                "Oder verwende ZIP-Dateien stattdessen."
            );
        }

        var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = winrarPath,
                Arguments = $"x -ibck -y \"{rarFile}\" \"{destinationFolder}\\\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };

        process.Start();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            var error = process.StandardError.ReadToEnd();
            throw new Exception($"RAR-Entpacken fehlgeschlagen.\n\nExitCode: {process.ExitCode}\n{error}");
        }
    }
}
