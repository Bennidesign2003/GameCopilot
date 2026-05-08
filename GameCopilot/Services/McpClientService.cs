using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace GameCopilot.Services;

/// <summary>
/// MCP (Model Context Protocol) client that manages the nvidia-gpu MCP server
/// as a subprocess and forwards tool calls from Ollama.
/// </summary>
public class McpClientService : IDisposable
{
    private Process? _process;
    private StreamWriter? _writer;
    private StreamReader? _reader;
    private int _nextId;
    private readonly SemaphoreSlim _rpcLock = new(1, 1);

    public bool IsRunning { get; set; }
    public bool HasProcessExited => _process?.HasExited ?? true;
    public string Status { get; set; } = "MCP nicht gestartet";
    public List<McpToolDef> Tools { get; } = new();

    /// <summary>
    /// Extract the bundled MCP server.py from embedded resources to AppData.
    /// Version-aware: keeps the AppData copy if it has a newer __mcp_version__ than the embedded one.
    /// </summary>
    private static string ExtractBundledServer()
    {
        var mcpDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "GameCopilot", "mcp-server");
        Directory.CreateDirectory(mcpDir);

        var serverPath = Path.Combine(mcpDir, "server.py");

        // Read embedded version
        using var stream = typeof(McpClientService).Assembly
            .GetManifestResourceStream("mcp-server.py");
        if (stream == null) return serverPath;

        // Read first 200 bytes to extract __mcp_version__ from embedded resource
        var headerBuf = new byte[200];
        var headerLen = stream.Read(headerBuf, 0, headerBuf.Length);
        var embeddedHeader = System.Text.Encoding.UTF8.GetString(headerBuf, 0, headerLen);
        var embeddedVersion = ExtractMcpVersion(embeddedHeader);

        // Check if AppData has a newer user-updated version
        if (File.Exists(serverPath))
        {
            var existingHeader = File.ReadLines(serverPath).Take(3)
                .Aggregate("", (a, b) => a + b + "\n");
            var existingVersion = ExtractMcpVersion(existingHeader);

            if (!string.IsNullOrEmpty(existingVersion) &&
                !string.IsNullOrEmpty(embeddedVersion) &&
                Version.TryParse(existingVersion, out var ev) &&
                Version.TryParse(embeddedVersion, out var bv) &&
                ev > bv)
            {
                // AppData has a newer version (user-updated) — keep it
                return serverPath;
            }
        }

        // Write embedded version (reset stream first)
        stream.Seek(0, SeekOrigin.Begin);
        using var fs = new FileStream(serverPath, FileMode.Create, FileAccess.Write);
        stream.CopyTo(fs);

        return serverPath;
    }

    private static string ExtractMcpVersion(string text)
    {
        // Looks for: # __mcp_version__ = "3.5.3"
        var m = System.Text.RegularExpressions.Regex.Match(
            text, @"__mcp_version__\s*=\s*""([^""]+)""");
        return m.Success ? m.Groups[1].Value : "";
    }

    /// <summary>
    /// Start the MCP server and discover available tools.
    /// </summary>
    public async Task StartAsync(Action<string>? onProgress = null)
    {
        if (IsRunning) return;

        // Extract bundled server.py to AppData (offline-safe baseline)
        onProgress?.Invoke("MCP Server wird extrahiert...");
        var serverPath = ExtractBundledServer();

        // Then check the nvidia-mcp GitHub release channel for a newer server.py.
        // Failures here are non-fatal — the embedded copy keeps us functional offline.
        onProgress?.Invoke("MCP Server: prüfe Online-Updates...");
        var updated = await UpdateService.TryUpdateMcpServerAsync(serverPath).ConfigureAwait(false);
        if (updated)
            onProgress?.Invoke("MCP Server wurde auf neueste Version aktualisiert");

        // Check if an MCP server update was applied — clear the restart flag
        var mcpDir = Path.GetDirectoryName(serverPath)!;
        var restartFlag = Path.Combine(mcpDir, "__mcp_restart_requested__");
        if (File.Exists(restartFlag))
        {
            File.Delete(restartFlag);
            onProgress?.Invoke("Neuer MCP Server wird gestartet...");
        }

        if (!File.Exists(serverPath))
        {
            Status = "MCP server.py konnte nicht extrahiert werden";
            onProgress?.Invoke(Status);
            return;
        }

        onProgress?.Invoke("MCP Server wird gestartet...");
        Status = "MCP Server wird gestartet...";

        // Ensure uv is installed
        var uvPath = await EnsureUvInstalledAsync(onProgress);
        if (string.IsNullOrEmpty(uvPath))
        {
            Status = "uv konnte nicht installiert werden";
            onProgress?.Invoke(Status);
            return;
        }

        // Build uv command to run the MCP server in stdio mode
        // Run uv directly (not via cmd.exe) to avoid Windows quoting issues with spaces in paths
        var uvArgs = $"run --with \"mcp[cli]\" --with nvidia-ml-py --with httpx --with py7zr --with rarfile --with websocket-client mcp run \"{serverPath}\"";

        var psi = new ProcessStartInfo
        {
            FileName = uvPath,
            Arguments = uvArgs,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        try
        {
            _process = Process.Start(psi);
            if (_process == null)
            {
                Status = "MCP Server konnte nicht gestartet werden";
                onProgress?.Invoke(Status);
                return;
            }

            _writer = _process.StandardInput;
            _writer.AutoFlush = true;
            _reader = _process.StandardOutput;

            // Capture stderr in background for error reporting
            var stderrBuilder = new StringBuilder();
            _ = Task.Run(async () =>
            {
                try
                {
                    while (!_process.StandardError.EndOfStream)
                    {
                        var line = await _process.StandardError.ReadLineAsync();
                        if (line != null) stderrBuilder.AppendLine(line);
                    }
                }
                catch { }
            });

            // Give server a moment to initialize
            await Task.Delay(3000);

            // Check if process crashed during startup
            if (_process.HasExited)
            {
                var stderr = stderrBuilder.ToString().Trim();
                var lastLines = stderr.Length > 300 ? stderr[^300..] : stderr;
                Status = $"MCP Server abgestuerzt: {(string.IsNullOrEmpty(lastLines) ? $"Exit-Code {_process.ExitCode}" : lastLines)}";
                onProgress?.Invoke(Status);
                Stop();
                return;
            }

            // MCP Initialize handshake
            onProgress?.Invoke("MCP Handshake...");
            var initResult = await SendRpcAsync("initialize", new
            {
                protocolVersion = "2024-11-05",
                capabilities = new { },
                clientInfo = new { name = "GameCopilot", version = "3.5.1" }
            });

            // Send initialized notification (no response expected)
            await SendNotificationAsync("notifications/initialized");

            // Discover tools
            onProgress?.Invoke("MCP Tools werden geladen...");
            var toolsResult = await SendRpcAsync("tools/list", new { });

            Tools.Clear();
            if (toolsResult.TryGetProperty("tools", out var toolsArr))
            {
                foreach (var t in toolsArr.EnumerateArray())
                {
                    var tool = new McpToolDef
                    {
                        Name = t.GetProperty("name").GetString() ?? "",
                        Description = t.TryGetProperty("description", out var desc)
                            ? desc.GetString() ?? "" : "",
                    };
                    if (t.TryGetProperty("inputSchema", out var schema))
                        tool.InputSchemaJson = schema.GetRawText();

                    Tools.Add(tool);
                }
            }

            IsRunning = true;
            Status = $"MCP bereit ({Tools.Count} Tools)";
            onProgress?.Invoke(Status);
        }
        catch (Exception ex)
        {
            Status = $"MCP Fehler: {ex.Message}";
            onProgress?.Invoke(Status);
            Stop();
        }
    }

    /// <summary>
    /// Stop and restart the MCP server process (e.g. after a server.py update).
    /// </summary>
    public async Task RestartAsync(Action<string>? onProgress = null)
    {
        onProgress?.Invoke("MCP Server wird neugestartet...");
        // Kill current process
        try
        {
            _process?.Kill(entireProcessTree: true);
            _process?.Dispose();
            _process = null;
        }
        catch { }

        IsRunning = false;
        Status = "Neustart...";

        await Task.Delay(1500);
        await StartAsync(onProgress);
    }

    /// <summary>
    /// Find uv on this machine, or install it automatically.
    /// Returns the full path to the uv binary, or empty string on failure.
    /// </summary>
    private async Task<string> EnsureUvInstalledAsync(Action<string>? onProgress = null)
    {
        var isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

        // 1) Check common install locations
        if (isWindows)
        {
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var candidates = new[]
            {
                Path.Combine(userProfile, ".local", "bin", "uv.exe"),
                Path.Combine(userProfile, ".cargo", "bin", "uv.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "uv", "uv.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "uv", "uv.exe"),
            };
            foreach (var c in candidates)
            {
                if (File.Exists(c)) return c;
            }
        }

        // 2) Check PATH
        try
        {
            var checkCmd = isWindows ? "where" : "which";
            var psi = new ProcessStartInfo(checkCmd, "uv")
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            var p = Process.Start(psi);
            if (p != null)
            {
                var output = (await p.StandardOutput.ReadToEndAsync()).Trim();
                p.WaitForExit(5000);
                if (p.ExitCode == 0 && !string.IsNullOrEmpty(output))
                    return output.Split('\n')[0].Trim();
            }
        }
        catch { }

        // 3) Not found - install automatically
        onProgress?.Invoke("uv wird installiert...");

        try
        {
            ProcessStartInfo installPsi;
            if (isWindows)
            {
                // Official uv installer for Windows via PowerShell
                installPsi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"irm https://astral.sh/uv/install.ps1 | iex\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
            }
            else
            {
                installPsi = new ProcessStartInfo
                {
                    FileName = "/bin/sh",
                    Arguments = "-c \"curl -LsSf https://astral.sh/uv/install.sh | sh\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
            }

            var proc = Process.Start(installPsi);
            if (proc != null)
            {
                await proc.WaitForExitAsync();
                await Task.Delay(1000);
            }

            // Check again after install
            if (isWindows)
            {
                var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var postInstallPaths = new[]
                {
                    Path.Combine(userProfile, ".local", "bin", "uv.exe"),
                    Path.Combine(userProfile, ".cargo", "bin", "uv.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "uv", "uv.exe"),
                };
                foreach (var c in postInstallPaths)
                {
                    if (File.Exists(c))
                    {
                        onProgress?.Invoke("uv erfolgreich installiert!");
                        return c;
                    }
                }
            }

            // Try PATH again after install
            try
            {
                var checkCmd = isWindows ? "where" : "which";
                var psi2 = new ProcessStartInfo(checkCmd, "uv")
                {
                    UseShellExecute = false, CreateNoWindow = true,
                    RedirectStandardOutput = true, RedirectStandardError = true
                };
                var p2 = Process.Start(psi2);
                if (p2 != null)
                {
                    var output = (await p2.StandardOutput.ReadToEndAsync()).Trim();
                    p2.WaitForExit(5000);
                    if (p2.ExitCode == 0 && !string.IsNullOrEmpty(output))
                    {
                        onProgress?.Invoke("uv erfolgreich installiert!");
                        return output.Split('\n')[0].Trim();
                    }
                }
            }
            catch { }

            onProgress?.Invoke("uv Installation fehlgeschlagen");
            return "";
        }
        catch (Exception ex)
        {
            onProgress?.Invoke($"uv Install-Fehler: {ex.Message}");
            return "";
        }
    }

    /// <summary>
    /// Call an MCP tool by name with given arguments.
    /// </summary>
    // Tools that can take a long time (MSFS restart, driver install, model pull)
    private static readonly HashSet<string> SlowTools = new(StringComparer.OrdinalIgnoreCase)
    {
        "optimize_msfs_graphics", "restore_msfs_graphics",
        "check_and_install_driver", "launch_msfs_vr", "fix_msfs",
        "set_msfs_setting",        // writes + optionally restarts MSFS
    };

    public async Task<string> CallToolAsync(string name, string argumentsJson, CancellationToken ct = default)
    {
        if (!IsRunning) return "{\"error\": \"MCP server not running\"}";

        try
        {
            var args = string.IsNullOrWhiteSpace(argumentsJson) || argumentsJson == "null"
                ? new { }
                : (object)JsonDocument.Parse(argumentsJson).RootElement;

            // Slow tools get 120s timeout, fast tools get 30s
            var timeout = SlowTools.Contains(name) ? 120_000 : 30_000;
            var result = await SendRpcAsync("tools/call", new { name, arguments = args }, timeout, ct);

            // MCP tool results have { content: [ { type: "text", text: "..." } ] }
            if (result.TryGetProperty("content", out var contentArr))
            {
                var sb = new StringBuilder();
                foreach (var c in contentArr.EnumerateArray())
                {
                    if (c.TryGetProperty("text", out var textEl))
                        sb.AppendLine(textEl.GetString());
                }
                return sb.ToString().TrimEnd();
            }

            return result.GetRawText();
        }
        catch (Exception ex)
        {
            return $"{{\"error\": \"{ex.Message.Replace("\"", "\\\"").Replace("\n", " ")}\"}}";
        }
    }

    /// <summary>
    /// Convert all MCP tools to Ollama tool-calling format.
    /// </summary>
    public List<object> GetOllamaToolDefinitions()
    {
        var ollamaTools = new List<object>();

        foreach (var tool in Tools)
        {
            object parameters;
            if (!string.IsNullOrEmpty(tool.InputSchemaJson))
            {
                try { parameters = JsonSerializer.Deserialize<object>(tool.InputSchemaJson)!; }
                catch { parameters = new { type = "object", properties = new { } }; }
            }
            else
            {
                parameters = new { type = "object", properties = new { } };
            }

            // Keep descriptions short to reduce context size
            var desc = tool.Description;
            var firstLine = desc.Contains('\n') ? desc[..desc.IndexOf('\n')] : desc;
            if (firstLine.Length > 200) firstLine = firstLine[..200];

            ollamaTools.Add(new
            {
                type = "function",
                function = new
                {
                    name = tool.Name,
                    description = firstLine,
                    parameters
                }
            });
        }

        return ollamaTools;
    }

    private async Task<JsonElement> SendRpcAsync(string method, object @params, int timeoutMs = 30000, CancellationToken ct = default)
    {
        await _rpcLock.WaitAsync(ct).ConfigureAwait(false);
        try
        {
            var id = Interlocked.Increment(ref _nextId);
            var request = JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id,
                method,
                @params
            });

            await _writer!.WriteLineAsync(request).ConfigureAwait(false);

            // Read lines until we find the response with matching id
            using var cts = new CancellationTokenSource(timeoutMs);
            while (!cts.Token.IsCancellationRequested)
            {
                var line = await _reader!.ReadLineAsync(cts.Token).ConfigureAwait(false);
                if (string.IsNullOrEmpty(line)) continue;

                try
                {
                    var doc = JsonDocument.Parse(line);
                    if (doc.RootElement.TryGetProperty("id", out var respId) &&
                        respId.GetInt32() == id)
                    {
                        if (doc.RootElement.TryGetProperty("result", out var result))
                            return result;
                        if (doc.RootElement.TryGetProperty("error", out var error))
                            throw new Exception($"MCP error: {error.GetRawText()}");
                        return doc.RootElement;
                    }
                    // Not our response (notification or other), keep reading
                }
                catch (JsonException) { /* skip non-JSON lines (stderr leaking?) */ }
            }

            throw new TimeoutException($"MCP request '{method}' timed out after {timeoutMs}ms");
        }
        finally
        {
            _rpcLock.Release();
        }
    }

    private async Task SendNotificationAsync(string method, object? @params = null)
    {
        await _rpcLock.WaitAsync().ConfigureAwait(false);
        try
        {
            // Guard: _writer may be null if Stop() was called concurrently.
            if (_writer == null) return;

            var notification = @params != null
                ? JsonSerializer.Serialize(new { jsonrpc = "2.0", method, @params })
                : JsonSerializer.Serialize(new { jsonrpc = "2.0", method });
            await _writer.WriteLineAsync(notification).ConfigureAwait(false);
        }
        finally
        {
            _rpcLock.Release();
        }
    }

    public void Stop()
    {
        IsRunning = false;
        try
        {
            _writer?.Close();
            _reader?.Close();   // was previously leaked (set to null without disposing)
            if (_process is { HasExited: false })
            {
                _process.Kill();
                _process.WaitForExit(3000);
            }
            _process?.Dispose();
        }
        catch { }
        _process = null;
        _writer = null;
        _reader = null;
    }

    public void Dispose()
    {
        Stop();
        _rpcLock.Dispose();
    }
}

public class McpToolDef
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string InputSchemaJson { get; set; } = "";
}
