using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using GameCopilot.Models;
using GameCopilot.ViewModels;

namespace GameCopilot.Views;

public partial class MainWindow : Window
{
    private DispatcherTimer? _thinkingTimer;
    private int _dotPhase;
    private bool _shimmerForward = true;

    public MainWindow()
    {
        InitializeComponent();

        // Start splash sequence (mirrors WPF SplashScreenPage_Loaded)
        Loaded += async (_, _) =>
        {
            if (DataContext is MainWindowViewModel vm)
                await vm.RunSplashSequenceAsync();
        };

        // Window drag (mirrors WPF Window_MouseLeftButtonDown + DragMove)
        PointerPressed += (_, e) =>
        {
            if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
                BeginMoveDrag(e);
        };

        // Apple-style thinking animation timer
        _thinkingTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
        _thinkingTimer.Tick += OnThinkingTick;
        _thinkingTimer.Start();

        // Auto-scroll chat to bottom when new messages arrive
        Loaded += (_, _) =>
        {
            if (DataContext is MainWindowViewModel vm2)
            {
                vm2.ChatMessages.CollectionChanged += (_, _) =>
                {
                    DispatcherTimer.RunOnce(() =>
                    {
                        var sv = this.FindControl<ScrollViewer>("ChatScrollViewer");
                        sv?.ScrollToEnd();
                    }, TimeSpan.FromMilliseconds(50));
                };
            }
        };

        // ── Global keyboard shortcuts ──────────────────────────────────────────
        // Register on the Window so they fire regardless of which control has focus.
        KeyDown += OnWindowKeyDown;
    }

    // ── Global keyboard shortcut handler ──────────────────────────────────────
    private void OnWindowKeyDown(object? sender, KeyEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var ctrl  = e.KeyModifiers.HasFlag(KeyModifiers.Control);
        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

        // Ctrl+K — clear chat
        if (ctrl && e.Key == Key.K)
        {
            vm.ClearChatCommand.Execute(null);
            e.Handled = true;
            return;
        }

        // Escape — cancel ongoing generation
        if (e.Key == Key.Escape && vm.IsGenerating)
        {
            vm.CancelGenerationCommand.Execute(null);
            e.Handled = true;
            return;
        }

        // Ctrl+/ — focus the message input box
        if (ctrl && e.Key == Key.OemQuestion)
        {
            var box = this.FindControl<TextBox>("ChatInputBox");
            box?.Focus();
            e.Handled = true;
            return;
        }
    }

    // ── Chat TextBox: Enter = send, Shift+Enter = new line ───────────────────
    private void OnChatInputKeyDown(object? sender, KeyEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

        if (e.Key == Key.Enter && !shift)
        {
            // Plain Enter → send
            if (!string.IsNullOrWhiteSpace(vm.ChatInput))
                vm.SendChatMessageCommand.Execute(null);
            e.Handled = true;
        }
        // Shift+Enter: Avalonia TextBox with AcceptsReturn="True" already inserts
        // a newline on Enter, but we let Shift+Enter through naturally while blocking
        // plain Enter above — so this branch requires no extra code.
    }

    private void OnThinkingTick(object? sender, EventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        // Auto-scroll chat while model is streaming (text grows, view follows)
        var lastMsg = vm.ChatMessages.Count > 0 ? vm.ChatMessages[^1] : null;
        if (vm.IsChatLoading || lastMsg?.IsStreaming == true)
        {
            var sv = this.FindControl<ScrollViewer>("ChatScrollViewer");
            sv?.ScrollToEnd();
        }

        if (!vm.IsChatLoading) return;

        // Pulsing dots animation (Apple-style sequential fade)
        var dot1 = this.FindControl<Border>("Dot1");
        var dot2 = this.FindControl<Border>("Dot2");
        var dot3 = this.FindControl<Border>("Dot3");
        var glow = this.FindControl<Border>("ThinkingGlow");
        var shimmer = this.FindControl<Border>("ShimmerBar");

        if (dot1 != null && dot2 != null && dot3 != null)
        {
            _dotPhase = (_dotPhase + 1) % 3;
            dot1.Opacity = _dotPhase == 0 ? 1.0 : 0.25;
            dot2.Opacity = _dotPhase == 1 ? 1.0 : 0.25;
            dot3.Opacity = _dotPhase == 2 ? 1.0 : 0.25;
        }

        // Glow pulse
        if (glow != null)
            glow.Opacity = glow.Opacity > 0.5 ? 0.2 : 0.8;

        // Shimmer bar animation
        if (shimmer != null)
        {
            if (_shimmerForward)
                shimmer.Margin = new Thickness(30, 0, 0, 0);
            else
                shimmer.Margin = new Thickness(0, 0, 0, 0);
            _shimmerForward = !_shimmerForward;
        }
    }

    // Copy assistant message content to clipboard
    private void OnCopyMessageClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { DataContext: GameCopilot.Models.ChatMessage msg })
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard != null && !string.IsNullOrEmpty(msg.Content))
                _ = clipboard.SetTextAsync(msg.Content);
        }
    }

    // Mirrors WPF CloseButton_Click
    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();

    // Mirrors WPF MinimizeButton_Click
    private void OnMinimizeClick(object? sender, RoutedEventArgs e) =>
        WindowState = WindowState.Minimized;

    // Mirrors WPF ModsPage.AddMod_Click - file picker then add
    private async void OnAddModClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Waehle eine Mod-Datei",
            AllowMultiple = false,
            FileTypeFilter = new List<FilePickerFileType>
            {
                new("ZIP & RAR Archive") { Patterns = new[] { "*.zip", "*.rar" } },
                new("Alle Dateien") { Patterns = new[] { "*" } }
            }
        });

        if (files.Count > 0)
            vm.AddModFromFile(files[0].Path.LocalPath);
    }

    // Export preset
    private async void OnExportClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Preset exportieren",
            DefaultExtension = "json",
            FileTypeChoices = new List<FilePickerFileType>
            {
                new("JSON Preset") { Patterns = new[] { "*.json" } }
            },
            SuggestedFileName = vm.SelectedPreset?.Name ?? "preset"
        });

        if (file != null)
            vm.ExportToFile(file.Path.LocalPath);
    }

    // Import preset
    private async void OnImportClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Preset importieren",
            AllowMultiple = false,
            FileTypeFilter = new List<FilePickerFileType>
            {
                new("JSON Preset") { Patterns = new[] { "*.json" } }
            }
        });

        if (files.Count > 0)
            vm.ImportFromFile(files[0].Path.LocalPath);
    }
}
