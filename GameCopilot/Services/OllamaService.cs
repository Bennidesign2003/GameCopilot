using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace GameCopilot.Services;

public class OllamaService
{
    // Short timeout for health checks — prevents UI freeze if Ollama hangs
    private static readonly HttpClient PingHttp = new() { Timeout = TimeSpan.FromSeconds(5) };
    // Chat timeout — 2 minutes for tool calls, enough for most responses
    private static readonly HttpClient ChatHttp = new() { Timeout = TimeSpan.FromMinutes(2) };
    // Long timeout for model downloads (can take hours for large models)
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(30) };
    private const string BaseUrl = "http://localhost:11434";
    private const string OllamaInstallScript = "https://ollama.com/install.ps1";

    public bool IsAvailable { get; private set; }
    public bool IsInstalled { get; private set; }
    public string Status { get; private set; } = "Nicht verbunden";
    public string[] AvailableModels { get; private set; } = Array.Empty<string>();

    // Preferred models in display order — sorted by recommended-for-most-users.
    // Newer / stronger families appear first so the picker leads with the best.
    public static readonly (string Id, string Label, string Size, string Desc)[] SupportedModels = new[]
    {
        // ── Qwen 3 family (Alibaba, latest as of 2025) ──
        ("qwen3:4b",          "Qwen 3",         "4B   ~3 GB",  "Sehr schnell, klein, gut fuer einfache Fragen"),
        ("qwen3:8b",          "Qwen 3",         "8B   ~5 GB",  "Schnell, gut fuer alltaegliche Fragen"),
        ("qwen3:14b",         "Qwen 3",         "14B  ~9 GB",  "Ausgewogen, gute Qualitaet"),
        ("qwen3:30b-a3b",     "Qwen 3 MoE",     "30B  ~18 GB", "Mixture-of-Experts: nur 3B aktive Params, sehr schnell trotz 30B Wissen"),
        ("qwen3:32b",         "Qwen 3",         "32B  ~20 GB", "Beste Qualitaet, braucht viel VRAM"),

        // ── Code-specialized (besser als Standard-Modelle fuer technische Anfragen) ──
        ("qwen2.5-coder:7b",  "Qwen 2.5 Coder", "7B   ~4 GB",  "Code-Spezialist, schnell"),
        ("qwen2.5-coder:14b", "Qwen 2.5 Coder", "14B  ~9 GB",  "Code-Spezialist, sehr gut bei Tools/MCP"),
        ("qwen2.5-coder:32b", "Qwen 2.5 Coder", "32B  ~20 GB", "Top Code-Qualitaet, braucht viel VRAM"),

        // ── OpenAI Open-Weights (Aug 2025 release) ──
        ("gpt-oss:20b",       "GPT-OSS",        "20B  ~13 GB", "OpenAI Open-Weights, sehr stark bei Reasoning"),
        ("gpt-oss:120b",      "GPT-OSS",        "120B ~65 GB", "Flagship Open-Weights, braucht 80GB+ VRAM"),

        // ── Reasoning-spezialisiert ──
        ("deepseek-r1:8b",    "DeepSeek R1",    "8B   ~5 GB",  "Reasoning-Modell, denkt Schritt fuer Schritt"),
        ("deepseek-r1:14b",   "DeepSeek R1",    "14B  ~9 GB",  "Starkes Reasoning, gute Balance"),
        ("deepseek-r1:32b",   "DeepSeek R1",    "32B  ~20 GB", "Top Reasoning, braucht viel VRAM"),

        // ── Meta Llama (Dec 2024 release, sehr stark) ──
        ("llama3.3:70b",      "Llama 3.3",      "70B  ~40 GB", "Meta Flagship, GPT-4-Klasse, braucht viel VRAM"),

        // ── Google Gemma 3 (multimodal, lange Kontexte) ──
        ("gemma3:12b",        "Gemma 3",        "12B  ~8 GB",  "128k Kontext, gut bei langen Texten"),
        ("gemma3:27b",        "Gemma 3",        "27B  ~17 GB", "Top Gemma-Variante, 128k Kontext"),
    };

    /// <summary>
    /// Check if Ollama is running and list available models.
    /// </summary>
    public async Task<bool> CheckConnectionAsync()
    {
        try
        {
            var resp = await PingHttp.GetAsync($"{BaseUrl}/api/tags");
            if (!resp.IsSuccessStatusCode)
            {
                IsAvailable = false;
                Status = "Ollama nicht erreichbar";
                return false;
            }

            var json = await resp.Content.ReadAsStringAsync();
            var doc = JsonDocument.Parse(json);
            var models = new List<string>();
            if (doc.RootElement.TryGetProperty("models", out var arr))
            {
                foreach (var m in arr.EnumerateArray())
                {
                    if (m.TryGetProperty("name", out var name))
                        models.Add(name.GetString() ?? "");
                }
            }

            AvailableModels = models.ToArray();
            IsAvailable = true;
            Status = $"Verbunden ({models.Count} Modelle)";
            return true;
        }
        catch
        {
            IsAvailable = false;
            Status = "Ollama nicht erreichbar";
            return false;
        }
    }

    /// <summary>
    /// Check if the ollama binary exists on this machine.
    /// </summary>
    public bool CheckInstalled()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            // Check common install locations + PowerShell install path + PATH
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var candidates = new[]
            {
                Path.Combine(localAppData, "Programs", "Ollama", "ollama.exe"),
                Path.Combine(programFiles, "Ollama", "ollama.exe"),
                Path.Combine(localAppData, "Ollama", "ollama.exe"),
                Path.Combine(userProfile, ".ollama", "ollama.exe"),
                Path.Combine(userProfile, "AppData", "Local", "Programs", "Ollama", "ollama.exe"),
            };

            if (candidates.Any(File.Exists))
            {
                IsInstalled = true;
                return true;
            }

            // Check PATH (short timeout to avoid UI freeze)
            try
            {
                var psi = new ProcessStartInfo("where", "ollama")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                var p = Process.Start(psi);
                p?.WaitForExit(2000);
                IsInstalled = p?.ExitCode == 0;
                return IsInstalled;
            }
            catch { }
        }
        else
        {
            try
            {
                var psi = new ProcessStartInfo("which", "ollama")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                var p = Process.Start(psi);
                p?.WaitForExit(2000);
                IsInstalled = p?.ExitCode == 0;
                return IsInstalled;
            }
            catch { }
        }

        IsInstalled = false;
        return false;
    }

    /// <summary>
    /// Download and install Ollama silently (Windows only).
    /// Reports progress via callback.
    /// </summary>
    public async Task InstallAsync(Action<string>? onProgress = null, CancellationToken ct = default)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Status = "Automatische Installation nur auf Windows. Bitte manuell installieren: ollama.com";
            return;
        }

        try
        {
            onProgress?.Invoke("Ollama wird installiert (PowerShell)...");
            Status = "Ollama wird installiert...";

            // Install via official PowerShell script: irm https://ollama.com/install.ps1 | iex
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"irm {OllamaInstallScript} | iex\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            var proc = Process.Start(psi);
            if (proc != null)
            {
                // Read output for progress feedback
                _ = Task.Run(async () =>
                {
                    try
                    {
                        while (!proc.StandardOutput.EndOfStream)
                        {
                            var line = await proc.StandardOutput.ReadLineAsync(ct);
                            if (!string.IsNullOrWhiteSpace(line))
                                onProgress?.Invoke($"Ollama: {line}");
                        }
                    }
                    catch { }
                }, ct);

                await proc.WaitForExitAsync(ct);
                await Task.Delay(2000, ct);
            }

            // Verify installation
            if (CheckInstalled())
            {
                onProgress?.Invoke("Ollama erfolgreich installiert!");
                Status = "Ollama installiert";
            }
            else
            {
                onProgress?.Invoke("Installation abgeschlossen - ollama evtl. noch nicht im PATH");
                Status = "Installiert (evtl. Neustart noetig)";
                IsInstalled = true;
            }
        }
        catch (OperationCanceledException)
        {
            onProgress?.Invoke("Installation abgebrochen");
            Status = "Installation abgebrochen";
        }
        catch (Exception ex)
        {
            onProgress?.Invoke($"Installationsfehler: {ex.Message}");
            Status = $"Installationsfehler: {ex.Message}";
        }
    }

    /// <summary>
    /// Full startup: check installed -> install if needed -> start serve -> verify connection.
    /// </summary>
    public async Task EnsureRunningAsync(Action<string>? onProgress = null, CancellationToken ct = default)
    {
        // Already running?
        if (await CheckConnectionAsync()) return;

        // Check if installed
        if (!CheckInstalled())
        {
            onProgress?.Invoke("Ollama nicht gefunden - wird automatisch installiert...");
            Status = "Ollama wird installiert...";
            await InstallAsync(onProgress, ct);

            if (!IsInstalled)
            {
                Status = "Ollama konnte nicht installiert werden";
                return;
            }

            // After fresh install via PowerShell, ollama might already be in PATH.
            // Give it a moment, then try to start serve.
            onProgress?.Invoke("Warte auf Ollama Start nach Installation...");
            for (int i = 0; i < 20; i++)
            {
                await Task.Delay(1000, ct);
                onProgress?.Invoke($"Ollama wird gestartet... ({i + 1}s)");
                if (await CheckConnectionAsync())
                {
                    onProgress?.Invoke("Ollama laeuft!");
                    return;
                }
            }
        }

        // Check again in case it started in the meantime
        if (await CheckConnectionAsync())
        {
            onProgress?.Invoke("Ollama laeuft!");
            return;
        }

        // Check if an ollama process is already running but slow to respond
        try
        {
            var procs = Process.GetProcessesByName("ollama");
            if (procs.Length > 0)
            {
                onProgress?.Invoke("Ollama Prozess gefunden, warte auf Server...");
                if (await WaitForConnectionAsync(8, "Ollama Prozess laeuft", onProgress, ct))
                    return;
            }
        }
        catch { }

        // Try to start ollama serve — first with known path, then via system PATH
        var ollamaExe = FindOllamaExe();
        string? serveExe = (ollamaExe != "ollama.exe" && ollamaExe != "ollama") ? ollamaExe : null;

        // If not found directly, try resolving via cmd.exe "where" (fresh system PATH)
        if (serveExe == null && RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            serveExe = await ResolveOllamaViaShellAsync(ct);

        if (!string.IsNullOrEmpty(serveExe))
        {
            try
            {
                onProgress?.Invoke("Ollama serve wird gestartet...");
                Status = "Ollama wird gestartet...";

                Process.Start(new ProcessStartInfo
                {
                    FileName = serveExe,
                    Arguments = "serve",
                    UseShellExecute = false,
                    CreateNoWindow = true
                });

                if (await WaitForConnectionAsync(15, "Warte auf Ollama", onProgress, ct))
                    return;
            }
            catch { }
        }

        Status = "Ollama Server antwortet nicht";
        onProgress?.Invoke("Ollama antwortet nicht - bitte Ollama manuell starten oder PC neustarten");
    }

    /// <summary>
    /// Use cmd.exe to resolve "ollama" from the current system PATH.
    /// This gets the fresh PATH (not the one inherited at process start).
    /// Returns full path to ollama.exe or null if not found.
    /// </summary>
    private async Task<string?> ResolveOllamaViaShellAsync(CancellationToken ct)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c where ollama",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            var p = Process.Start(psi);
            if (p != null)
            {
                var output = await p.StandardOutput.ReadToEndAsync(ct);
                await p.WaitForExitAsync(ct);
                if (p.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                {
                    var path = output.Split('\n')[0].Trim();
                    if (File.Exists(path)) return path;
                }
            }
        }
        catch { }
        return null;
    }

    private async Task<bool> WaitForConnectionAsync(int seconds, string label,
        Action<string>? onProgress, CancellationToken ct)
    {
        for (int i = 0; i < seconds; i++)
        {
            await Task.Delay(1000, ct);
            onProgress?.Invoke($"{label}... ({i + 1}s)");
            if (await CheckConnectionAsync())
            {
                onProgress?.Invoke("Ollama laeuft!");
                return true;
            }
        }
        return false;
    }

    private string FindOllamaExe()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var candidates = new[]
            {
                Path.Combine(localAppData, "Programs", "Ollama", "ollama.exe"),
                Path.Combine(programFiles, "Ollama", "ollama.exe"),
                Path.Combine(localAppData, "Ollama", "ollama.exe"),
                Path.Combine(userProfile, ".ollama", "ollama.exe"),
                Path.Combine(userProfile, "AppData", "Local", "Programs", "Ollama", "ollama.exe"),
            };

            var found = candidates.FirstOrDefault(File.Exists);
            // Don't do blocking "where" here — ResolveOllamaViaShellAsync does it async
            return found ?? "ollama.exe";
        }
        return "ollama";
    }

    /// <summary>
    /// Pull a model if not already downloaded. Reports progress via callback.
    /// </summary>
    /// <summary>
    /// Progress info for model download — includes bytes for progress bar.
    /// </summary>
    public class PullProgress
    {
        public string Status { get; set; } = "";
        public long CompletedBytes { get; set; }
        public long TotalBytes { get; set; }
        public double Percent => TotalBytes > 0 ? (double)CompletedBytes / TotalBytes : 0;
    }

    public async Task PullModelAsync(string model, Action<string>? onProgress = null,
        Action<PullProgress>? onDetailedProgress = null, CancellationToken ct = default)
    {
        // Verify Ollama is reachable before starting pull
        if (!await CheckConnectionAsync())
            throw new Exception("Ollama nicht erreichbar - bitte Ollama starten");

        onProgress?.Invoke("Verbinde mit Ollama...");
        onDetailedProgress?.Invoke(new PullProgress { Status = "Verbinde..." });

        var body = JsonSerializer.Serialize(new { name = model, stream = true });
        var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/api/pull")
        {
            Content = new StringContent(body, Encoding.UTF8, "application/json")
        };

        using var resp = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
        if (!resp.IsSuccessStatusCode)
        {
            var errBody = await resp.Content.ReadAsStringAsync(ct);
            throw new Exception($"Ollama Fehler ({resp.StatusCode}): {errBody}");
        }

        using var stream = await resp.Content.ReadAsStreamAsync(ct);
        using var reader = new StreamReader(stream);

        onProgress?.Invoke("Pulling manifest...");
        onDetailedProgress?.Invoke(new PullProgress { Status = "Pulling manifest..." });

        while (!reader.EndOfStream)
        {
            ct.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync();
            if (string.IsNullOrEmpty(line)) continue;

            try
            {
                var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;

                // Check for error response from Ollama
                if (root.TryGetProperty("error", out var error))
                {
                    var errMsg = error.GetString() ?? "Unbekannter Fehler";
                    throw new Exception($"Ollama: {errMsg}");
                }

                if (root.TryGetProperty("status", out var status))
                {
                    var statusText = status.GetString() ?? "";
                    if (root.TryGetProperty("completed", out var completed) &&
                        root.TryGetProperty("total", out var total))
                    {
                        var completedBytes = completed.GetInt64();
                        var totalBytes = total.GetInt64();
                        var pct = totalBytes > 0 ? (int)(completedBytes * 100 / totalBytes) : 0;
                        var completedMb = completedBytes / (1024.0 * 1024.0);
                        var totalMb = totalBytes / (1024.0 * 1024.0);

                        onProgress?.Invoke($"{statusText} — {completedMb:F0}/{totalMb:F0} MB ({pct}%)");
                        onDetailedProgress?.Invoke(new PullProgress
                        {
                            Status = statusText,
                            CompletedBytes = completedBytes,
                            TotalBytes = totalBytes
                        });
                    }
                    else
                    {
                        onProgress?.Invoke(statusText);
                        onDetailedProgress?.Invoke(new PullProgress { Status = statusText });
                    }
                }
            }
            catch (JsonException) { /* skip non-JSON lines */ }
        }

        // Refresh model list
        await CheckConnectionAsync();
    }

    /// <summary>
    /// Check if a specific model is already downloaded.
    /// </summary>
    public bool HasModel(string model)
    {
        // Exact match: "qwen3:8b" against downloaded "qwen3:8b"
        // Also handle Ollama's ":latest" suffix: "qwen3:8b" matches "qwen3:8b-*" variants
        var modelBase = model.Split(':')[0];
        var modelTag = model.Contains(':') ? model.Split(':')[1] : "";
        return AvailableModels.Any(m =>
        {
            if (m.Equals(model, StringComparison.OrdinalIgnoreCase)) return true;
            // Match "qwen3:8b" against "qwen3:8b-instruct-q4_0" but NOT "qwen3:14b"
            if (!string.IsNullOrEmpty(modelTag))
            {
                var mTag = m.Contains(':') ? m.Split(':')[1] : "";
                if (mTag.StartsWith(modelTag, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        });
    }

    /// <summary>
    /// Stream a chat completion. Calls onToken for each text chunk as it arrives.
    /// </summary>
    public async Task StreamChatAsync(
        string model,
        string systemPrompt,
        List<(string role, string content)> messages,
        Action<string> onToken,
        CancellationToken ct = default)
    {
        var msgArray = new List<object>();

        // System message
        msgArray.Add(new { role = "system", content = systemPrompt });

        // Conversation history
        foreach (var (role, content) in messages)
            msgArray.Add(new { role, content });

        var payload = JsonSerializer.Serialize(new
        {
            model,
            messages = msgArray,
            stream = true,
            // Keep the model loaded in VRAM for 30 minutes between requests so
            // the user only pays the warm-up cost once per session. Default
            // Ollama lifetime is 5 min — far too short for an interactive chat.
            keep_alive = "30m",
            options = new
            {
                temperature = 0.7,
                num_predict = 8192,
                num_ctx = 16384
            }
        });

        var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/api/chat")
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };

        using var resp = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
        resp.EnsureSuccessStatusCode();

        using var stream = await resp.Content.ReadAsStreamAsync(ct);
        using var reader = new StreamReader(stream);

        while (!reader.EndOfStream)
        {
            ct.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync();
            if (string.IsNullOrEmpty(line)) continue;

            try
            {
                var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                if (root.TryGetProperty("message", out var msg) &&
                    msg.TryGetProperty("content", out var content))
                {
                    var token = content.GetString();
                    if (!string.IsNullOrEmpty(token))
                        onToken(token);
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// Non-streaming chat call that supports tool calling.
    /// Returns the full response including any tool_calls.
    /// </summary>
    public async Task<OllamaChatResponse> ChatWithToolsAsync(
        string model,
        List<object> messages,
        List<object>? tools = null,
        CancellationToken ct = default)
    {
        var hasTools = tools != null && tools.Count > 0;
        var payloadObj = new Dictionary<string, object>
        {
            ["model"] = model,
            ["messages"] = messages,
            ["stream"] = false,
            // Same keep_alive as the streaming path — VRAM-resident model = no
            // 5–10s reload delay between turns.
            ["keep_alive"] = "30m",
            // Tool-calling: moderate output for tool selection + tables in final answer
            ["options"] = new { temperature = 0.7, num_predict = 4096, num_ctx = hasTools ? 8192 : 16384 }
        };
        if (hasTools)
            payloadObj["tools"] = tools!;

        var payload = JsonSerializer.Serialize(payloadObj);

        var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/api/chat")
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };

        using var resp = await ChatHttp.SendAsync(request, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();
        var json = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var result = new OllamaChatResponse();

        if (root.TryGetProperty("message", out var msg))
        {
            if (msg.TryGetProperty("content", out var content))
                result.Content = content.GetString() ?? "";

            if (msg.TryGetProperty("tool_calls", out var toolCalls))
            {
                foreach (var tc in toolCalls.EnumerateArray())
                {
                    if (tc.TryGetProperty("function", out var fn))
                    {
                        var call = new OllamaToolCall
                        {
                            Name = fn.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "",
                            ArgumentsJson = fn.TryGetProperty("arguments", out var a) ? a.GetRawText() : "{}"
                        };
                        result.ToolCalls.Add(call);
                    }
                }
            }

            // Store the raw message element for conversation threading
            result.RawMessageJson = msg.GetRawText();
        }

        return result;
    }
}

public class OllamaChatResponse
{
    public string Content { get; set; } = "";
    public List<OllamaToolCall> ToolCalls { get; } = new();
    public bool HasToolCalls => ToolCalls.Count > 0;
    public string RawMessageJson { get; set; } = "";
}

public class OllamaToolCall
{
    public string Name { get; set; } = "";
    public string ArgumentsJson { get; set; } = "{}";
}
