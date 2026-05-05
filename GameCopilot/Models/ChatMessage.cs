using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace GameCopilot.Models;

public class ChatMessage : INotifyPropertyChanged
{
    public string Role { get; set; } = "user";

    private string _content = "";
    public string Content
    {
        get => _content;
        set { _content = value; OnPropertyChanged(); }
    }

    private string _timestamp = "";
    public string Timestamp
    {
        get => _timestamp;
        set { _timestamp = value; OnPropertyChanged(); }
    }

    public bool IsUser => Role == "user";
    public bool IsAssistant => Role == "assistant";
    public bool IsSystem => Role == "system";

    // True while the model is streaming tokens into this bubble
    private bool _isStreaming;
    public bool IsStreaming
    {
        get => _isStreaming;
        set { _isStreaming = value; OnPropertyChanged(); }
    }

    // Agent activity steps (tool calls shown live)
    public ObservableCollection<AgentStep> AgentSteps { get; } = new();

    private bool _hasAgentSteps;
    public bool HasAgentSteps
    {
        get => _hasAgentSteps;
        set { _hasAgentSteps = value; OnPropertyChanged(); }
    }

    private bool _isAgentWorking;
    public bool IsAgentWorking
    {
        get => _isAgentWorking;
        set { _isAgentWorking = value; OnPropertyChanged(); }
    }

    // Recommendation cards & action buttons
    public ObservableCollection<RecommendationCard> Cards { get; } = new();
    public ObservableCollection<ActionButton> ActionButtons { get; } = new();

    private bool _hasCards;
    public bool HasCards
    {
        get => _hasCards;
        set { _hasCards = value; OnPropertyChanged(); }
    }

    private bool _hasActions;
    public bool HasActions
    {
        get => _hasActions;
        set { _hasActions = value; OnPropertyChanged(); }
    }

    /// <summary>Token usage from Codex API (0 = not tracked / Ollama).</summary>
    private int _totalTokens;
    public int TotalTokens
    {
        get => _totalTokens;
        set { _totalTokens = value; OnPropertyChanged(); OnPropertyChanged(nameof(TokenLabel)); }
    }

    /// <summary>Human-readable token count, e.g. "~1.2k tokens". Empty when not available.</summary>
    public string TokenLabel => TotalTokens > 0
        ? (TotalTokens >= 1000 ? $"~{TotalTokens / 1000.0:F1}k tokens" : $"{TotalTokens} tokens")
        : "";

    // MCP Update prompt support
    private bool _isUpdatePrompt;
    public bool IsUpdatePrompt
    {
        get => _isUpdatePrompt;
        set { _isUpdatePrompt = value; OnPropertyChanged(); }
    }

    public string? UpdateVersion { get; set; }

    private bool _updateDismissed;
    public bool UpdateDismissed
    {
        get => _updateDismissed;
        set { _updateDismissed = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class AgentStep : INotifyPropertyChanged
{
    private string _icon = "⏳";
    public string Icon
    {
        get => _icon;
        set { _icon = value; OnPropertyChanged(); }
    }

    private string _statusText = "";
    public string StatusText
    {
        get => _statusText;
        set { _statusText = value; OnPropertyChanged(); }
    }

    private string _resultSummary = "";
    public string ResultSummary
    {
        get => _resultSummary;
        set { _resultSummary = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasResult)); }
    }

    public bool HasResult => !string.IsNullOrEmpty(_resultSummary);

    private bool _isComplete;
    public bool IsComplete
    {
        get => _isComplete;
        set { _isComplete = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsWorking)); }
    }

    public bool IsWorking => !_isComplete;

    private string _iconColor = "#69daff";
    public string IconColor
    {
        get => _iconColor;
        set { _iconColor = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class RecommendationCard
{
    public string Title { get; set; } = "";
    public ObservableCollection<RecommendationItem> Items { get; } = new();
}

public class RecommendationItem
{
    public string Label { get; set; } = "";
    public string Value { get; set; } = "";
}

public class ActionButton
{
    public string Text { get; set; } = "";
    public bool IsPrimary { get; set; }
}

public class ChatModelItem : INotifyPropertyChanged
{
    public string Id { get; set; } = "";
    public string Label { get; set; } = "";
    public string Size { get; set; } = "";
    public string Description { get; set; } = "";

    private bool _isInstalled;
    public bool IsInstalled
    {
        get => _isInstalled;
        set { _isInstalled = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNotInstalled)); }
    }
    public bool IsNotInstalled => !_isInstalled;

    private bool _isDownloading;
    public bool IsDownloading
    {
        get => _isDownloading;
        set { _isDownloading = value; OnPropertyChanged(); }
    }

    private string _downloadProgress = "";
    public string DownloadProgress
    {
        get => _downloadProgress;
        set { _downloadProgress = value; OnPropertyChanged(); }
    }

    private double _downloadPercent;
    public double DownloadPercent
    {
        get => _downloadPercent;
        set { _downloadPercent = value; OnPropertyChanged(); }
    }

    private string _downloadSizeText = "";
    public string DownloadSizeText
    {
        get => _downloadSizeText;
        set { _downloadSizeText = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
