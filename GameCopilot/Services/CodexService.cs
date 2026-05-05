using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace GameCopilot.Services;

/// <summary>
/// OpenAI Chat Completions backend (the API powering OpenAI Codex).
/// Supports SSE streaming and tool calling — mirrors OllamaService's
/// StreamChatWithToolsAsync signature so MainWindowViewModel can swap
/// providers without duplicating the agent loop.
/// </summary>
public class CodexService
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };
    private const string BaseUrl = "https://api.openai.com/v1/chat/completions";

    /// <summary>
    /// Cloud models exposed in the provider picker (analogous to OllamaService.SupportedModels).
    /// </summary>
    public static readonly (string Id, string Label, string Size, string Desc)[] SupportedModels =
    {
        ("codex-mini-latest", "Codex Mini",   "Cloud", "OpenAI Codex Mini — schnell, Code-fokussiert"),
        ("gpt-4.1",           "GPT-4.1",      "Cloud", "OpenAI GPT-4.1 — stärkstes Modell"),
        ("gpt-4.1-mini",      "GPT-4.1 Mini", "Cloud", "OpenAI GPT-4.1 Mini — ausgewogen"),
        ("o4-mini",           "o4-mini",      "Cloud", "OpenAI o4-mini — Reasoning-Spezialist"),
    };

    /// <summary>
    /// Streaming chat call with tool support (OpenAI SSE format).
    /// Tokens are delivered live via <paramref name="onToken"/>;
    /// tool calls are collected and returned in the result.
    /// The returned <see cref="OllamaStreamWithToolsResult.RawMessageJson"/>
    /// is formatted for OpenAI's conversation threading (includes tool_call ids).
    /// </summary>
    public async Task<OllamaStreamWithToolsResult> StreamChatWithToolsAsync(
        string model,
        List<object> messages,
        List<object>? tools,
        Action<string> onToken,
        string apiKey,
        CancellationToken ct = default)
    {
        var hasTools = tools != null && tools.Count > 0;

        var payloadObj = new Dictionary<string, object>
        {
            ["model"] = model,
            ["messages"] = messages,
            ["stream"] = true,
            ["max_tokens"] = 4096,
            ["temperature"] = hasTools ? 0.1 : 0.65,
        };
        if (hasTools)
            payloadObj["tools"] = tools!;

        var payload = JsonSerializer.Serialize(payloadObj);

        var request = new HttpRequestMessage(HttpMethod.Post, BaseUrl)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        using var resp = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);

        // Handle rate limit (429) gracefully
        if ((int)resp.StatusCode == 429)
        {
            var retryAfter = resp.Headers.RetryAfter?.Delta?.Seconds ?? 60;
            var limitMsg = $"\n\n⏱ Rate limit erreicht. Bitte warte {retryAfter}s und versuche es erneut.";
            onToken(limitMsg);
            return new OllamaStreamWithToolsResult { Content = limitMsg };
        }

        resp.EnsureSuccessStatusCode();

        using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var reader = new StreamReader(stream);

        var result = new OllamaStreamWithToolsResult();
        var contentBuffer = new StringBuilder();
        int totalTokens = 0;

        // Accumulate tool call fragments by index.
        // OpenAI streams partial arguments across multiple chunks.
        var toolCallMap = new Dictionary<int, ToolCallAccumulator>();

        while (!reader.EndOfStream)
        {
            ct.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync().ConfigureAwait(false);
            if (string.IsNullOrEmpty(line)) continue;
            if (!line.StartsWith("data: ", StringComparison.Ordinal)) continue;

            var data = line[6..]; // strip "data: "
            if (data == "[DONE]") break;

            try
            {
                using var doc = JsonDocument.Parse(data);
                var root = doc.RootElement;

                // ── Token usage (appears in final chunk from OpenAI) ──────────
                if (root.TryGetProperty("usage", out var usage) &&
                    usage.TryGetProperty("total_tokens", out var totalEl))
                    totalTokens = totalEl.GetInt32();

                if (!root.TryGetProperty("choices", out var choices)) continue;
                if (choices.GetArrayLength() == 0) continue;
                var choice = choices[0];
                if (!choice.TryGetProperty("delta", out var delta)) continue;

                // ── Content tokens ────────────────────────────────────────────
                if (delta.TryGetProperty("content", out var contentEl) &&
                    contentEl.ValueKind == JsonValueKind.String)
                {
                    var token = contentEl.GetString();
                    if (!string.IsNullOrEmpty(token))
                    {
                        contentBuffer.Append(token);
                        onToken(token);
                    }
                }

                // ── Tool call fragments ───────────────────────────────────────
                if (delta.TryGetProperty("tool_calls", out var tcArr) &&
                    tcArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var tc in tcArr.EnumerateArray())
                    {
                        var idx = tc.TryGetProperty("index", out var idxEl) ? idxEl.GetInt32() : 0;
                        if (!toolCallMap.ContainsKey(idx))
                            toolCallMap[idx] = new ToolCallAccumulator();

                        var acc = toolCallMap[idx];

                        if (tc.TryGetProperty("id", out var idEl))
                            acc.Id = idEl.GetString() ?? acc.Id;

                        if (tc.TryGetProperty("function", out var fn))
                        {
                            if (fn.TryGetProperty("name", out var nameEl))
                                acc.Name = nameEl.GetString() ?? acc.Name;
                            if (fn.TryGetProperty("arguments", out var argsEl))
                                acc.Args.Append(argsEl.GetString() ?? "");
                        }
                    }
                }
            }
            catch { /* skip malformed chunk */ }
        }

        result.Content = contentBuffer.ToString();
        if (totalTokens > 0)
            result.TotalTokens = totalTokens;

        // ── Build result tool calls + OpenAI-format echo message ─────────────
        var openAiToolCallsForEcho = new List<object>();

        foreach (var (_, acc) in toolCallMap)
        {
            var argsStr = acc.Args.Length > 0 ? acc.Args.ToString() : "{}";

            result.ToolCalls.Add(new OllamaToolCall
            {
                Id = acc.Id,
                Name = acc.Name,
                ArgumentsJson = argsStr,
            });

            openAiToolCallsForEcho.Add(new
            {
                id = acc.Id,
                type = "function",
                function = new { name = acc.Name, arguments = argsStr }
            });
        }

        // The assistant message we echo back on the next conversation turn.
        // OpenAI format: content is null when there are tool calls.
        var echoObj = new Dictionary<string, object?>
        {
            ["role"] = "assistant",
            ["content"] = result.Content.Length > 0 ? result.Content : null,
        };
        if (openAiToolCallsForEcho.Count > 0)
            echoObj["tool_calls"] = openAiToolCallsForEcho;

        result.RawMessageJson = JsonSerializer.Serialize(echoObj);
        return result;
    }

    /// <summary>
    /// Simple streaming chat without tool support (fallback / no-MCP path).
    /// </summary>
    public async Task StreamChatAsync(
        string model,
        List<object> messages,
        Action<string> onToken,
        string apiKey,
        CancellationToken ct = default)
    {
        var payloadObj = new Dictionary<string, object>
        {
            ["model"] = model,
            ["messages"] = messages,
            ["stream"] = true,
            ["max_tokens"] = 8192,
            ["temperature"] = 0.65,
        };

        var payload = JsonSerializer.Serialize(payloadObj);

        var request = new HttpRequestMessage(HttpMethod.Post, BaseUrl)
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        using var resp = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var reader = new StreamReader(stream);

        while (!reader.EndOfStream)
        {
            ct.ThrowIfCancellationRequested();
            var line = await reader.ReadLineAsync().ConfigureAwait(false);
            if (string.IsNullOrEmpty(line)) continue;
            if (!line.StartsWith("data: ", StringComparison.Ordinal)) continue;

            var data = line[6..];
            if (data == "[DONE]") break;

            try
            {
                using var doc = JsonDocument.Parse(data);
                var root = doc.RootElement;
                if (!root.TryGetProperty("choices", out var choices)) continue;
                if (choices.GetArrayLength() == 0) continue;
                var choice = choices[0];
                if (!choice.TryGetProperty("delta", out var delta)) continue;

                if (delta.TryGetProperty("content", out var contentEl) &&
                    contentEl.ValueKind == JsonValueKind.String)
                {
                    var token = contentEl.GetString();
                    if (!string.IsNullOrEmpty(token))
                        onToken(token);
                }
            }
            catch { /* skip malformed chunk */ }
        }
    }

    // ── API key validation ────────────────────────────────────────────────────

    /// <summary>
    /// Sends a minimal test request to OpenAI to verify the API key is valid.
    /// Returns true and clears the error on success; returns false with a reason string on failure.
    /// </summary>
    public async Task<(bool IsValid, string Reason)> ValidateApiKeyAsync(
        string apiKey, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
            return (false, "API-Key ist leer");

        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(TimeSpan.FromSeconds(8));

            // Use the models list endpoint — it's lightweight and requires a valid key
            var req = new HttpRequestMessage(HttpMethod.Get, "https://api.openai.com/v1/models");
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

            using var resp = await Http.SendAsync(req, cts.Token).ConfigureAwait(false);
            if (resp.IsSuccessStatusCode) return (true, "");

            if ((int)resp.StatusCode == 401)
                return (false, "Ungültiger API-Key (401)");
            if ((int)resp.StatusCode == 429)
                return (false, "Rate limit (429) — Key ist gültig");

            return (false, $"HTTP {(int)resp.StatusCode}");
        }
        catch (OperationCanceledException)
        {
            return (false, "Zeitüberschreitung (8s)");
        }
        catch (Exception ex)
        {
            return (false, $"Netzwerkfehler: {ex.Message}");
        }
    }

    // ── Internal helpers ──────────────────────────────────────────────────────

    private sealed class ToolCallAccumulator
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public StringBuilder Args { get; } = new();
    }
}
