using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using GameCopilot.Models;

namespace GameCopilot.Services;

/// <summary>
/// Persists chat messages to %AppData%\GameCopilot\chat_history.json
/// and restores them on the next app launch.
/// </summary>
public class ChatHistoryService
{
    private static readonly string HistoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "GameCopilot", "chat_history.json");

    /// <summary>
    /// Serialize the last 20 non-ephemeral messages to disk.
    /// IsUpdatePrompt and still-streaming messages are excluded.
    /// </summary>
    public void Save(IEnumerable<ChatMessage> messages)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(HistoryPath)!);

            var toSave = messages
                .Where(m => !m.IsUpdatePrompt && !m.IsStreaming && !m.IsAgentWorking
                         && !(m.Content ?? "").Contains("Letzte Sitzung"))
                .TakeLast(20)
                .Select(m => new HistoryEntry
                {
                    Role      = m.Role      ?? "assistant",
                    Content   = m.Content   ?? "",
                    Timestamp = m.Timestamp ?? ""
                })
                .ToList();

            var json = JsonSerializer.Serialize(toSave,
                new JsonSerializerOptions { WriteIndented = false });

            // Atomic write: temp-file + rename
            var tmp = HistoryPath + ".tmp";
            File.WriteAllText(tmp, json);
            File.Move(tmp, HistoryPath, overwrite: true);
        }
        catch
        {
            // Non-fatal — history persistence must never crash the app
        }
    }

    /// <summary>
    /// Load up to 20 messages from the history file.
    /// Returns an empty list if the file is missing or corrupt.
    /// </summary>
    public List<ChatMessage> Load()
    {
        try
        {
            if (!File.Exists(HistoryPath)) return new();

            var json    = File.ReadAllText(HistoryPath);
            var entries = JsonSerializer.Deserialize<List<HistoryEntry>>(json);
            if (entries is not { Count: > 0 }) return new();

            return entries
                .TakeLast(20)
                .Select(e => new ChatMessage
                {
                    Role      = e.Role,
                    Content   = e.Content,
                    Timestamp = e.Timestamp
                })
                .ToList();
        }
        catch
        {
            return new();
        }
    }

    /// <summary>Delete the history file on explicit clear.</summary>
    public void Clear()
    {
        try
        {
            if (File.Exists(HistoryPath))
                File.Delete(HistoryPath);
        }
        catch { }
    }

    // ── Serialisation DTO ────────────────────────────────────────────────────
    private class HistoryEntry
    {
        public string Role      { get; set; } = "";
        public string Content   { get; set; } = "";
        public string Timestamp { get; set; } = "";
    }
}
