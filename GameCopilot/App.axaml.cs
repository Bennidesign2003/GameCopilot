using System;
using System.IO;
using System.Threading;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Core;
using Avalonia.Data.Core.Plugins;
using System.Linq;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using GameCopilot.ViewModels;
using GameCopilot.Views;

namespace GameCopilot;

public partial class App : Application
{
    private static readonly string CrashLogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "GameCopilot", "crash.log");

    // Named mutex prevents multiple instances on the same user session.
    private static Mutex? _instanceMutex;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);

        // ── Global unhandled exception handlers ──────────────────────────────
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException      += OnUnobservedTaskException;
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // ── Single-instance guard ─────────────────────────────────────────
            const string mutexName = "Global\\GameCopilot_SingleInstance";
            _instanceMutex = new Mutex(
                initiallyOwned: true,
                name: mutexName,
                out bool createdNew);

            if (!createdNew)
            {
                // Another instance is already running — exit silently.
                _instanceMutex.Dispose();
                desktop.Shutdown();
                return;
            }

            // ── --reset-config argument ───────────────────────────────────────
            var args = desktop.Args ?? Array.Empty<string>();
            if (args.Contains("--reset-config", StringComparer.OrdinalIgnoreCase))
            {
                var configDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "GameCopilot");
                try
                {
                    var cfg = Path.Combine(configDir, "appconfig.json");
                    if (File.Exists(cfg)) File.Delete(cfg);

                    var history = Path.Combine(configDir, "chat_history.json");
                    if (File.Exists(history)) File.Delete(history);

                    WriteCrashLog("--reset-config: cleared appconfig.json + chat_history.json", null);
                }
                catch (Exception ex)
                {
                    WriteCrashLog("--reset-config: failed to clear config", ex);
                }
            }

            desktop.MainWindow = new MainWindow
            {
                DataContext = new MainWindowViewModel(),
            };

            // Release mutex when the app exits
            desktop.Exit += (_, _) =>
            {
                try { _instanceMutex?.ReleaseMutex(); } catch { /* already released */ }
                _instanceMutex?.Dispose();
            };
        }

        base.OnFrameworkInitializationCompleted();
    }

    // ── Crash handling helpers ────────────────────────────────────────────────

    private static void WriteCrashLog(string context, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(CrashLogPath)!);
            var entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}\n{ex}\n\n";
            File.AppendAllText(CrashLogPath, entry);
        }
        catch { /* swallow — crash log write must never throw */ }
    }

    private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        var ex = e.ExceptionObject as Exception;
        WriteCrashLog("AppDomain.UnhandledException", ex);

        // Show a friendly dialog on the UI thread if possible
        try
        {
            Dispatcher.UIThread.Post(async () =>
            {
                var win = new Window
                {
                    Title = "Game Copilot – Unerwarteter Fehler",
                    Width = 500, Height = 260,
                    Background = Avalonia.Media.Brushes.Black,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    CanResize = false,
                };
                var tb = new TextBlock
                {
                    Text = $"Game Copilot ist abgestürzt.\n\nFehler: {ex?.Message ?? "Unbekannt"}\n\n" +
                           $"Ein Crash-Log wurde gespeichert unter:\n{CrashLogPath}",
                    Foreground = Avalonia.Media.Brushes.White,
                    Margin = new Avalonia.Thickness(24),
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    FontSize = 13,
                };
                win.Content = tb;
                win.Show();
            });
        }
        catch { /* ignore — best effort */ }
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        e.SetObserved(); // prevent process crash for fire-and-forget tasks
        WriteCrashLog("UnobservedTaskException", e.Exception?.InnerException ?? e.Exception);
    }
}
