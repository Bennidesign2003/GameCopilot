using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace GameCopilot.Models;

/// <summary>
/// Represents an installed mod. Migrated from WPF ModItem class in ModsPage.xaml.cs.
/// </summary>
public class ModItem : INotifyPropertyChanged
{
    public string Name { get; set; } = string.Empty;
    public string FullPath { get; set; } = string.Empty;

    private string _category = "All";
    public string Category
    {
        get => _category;
        set { _category = value; OnPropertyChanged(); }
    }

    public static string[] Categories => new[] { "All", "Aircraft", "Helicopters", "Scenery" };

    public string SizeFormatted
    {
        get
        {
            try
            {
                if (Directory.Exists(FullPath))
                {
                    var bytes = new DirectoryInfo(FullPath)
                        .EnumerateFiles("*", SearchOption.AllDirectories)
                        .Sum(f => f.Length);
                    return FormatBytes(bytes);
                }
                if (File.Exists(FullPath))
                    return FormatBytes(new FileInfo(FullPath).Length);
            }
            catch { }
            return "—";
        }
    }

    private static string FormatBytes(long bytes) => bytes switch
    {
        >= 1_073_741_824 => $"{bytes / 1_073_741_824.0:F1} GB",
        >= 1_048_576 => $"{bytes / 1_048_576.0:F0} MB",
        >= 1_024 => $"{bytes / 1_024.0:F0} KB",
        _ => $"{bytes} B"
    };

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
