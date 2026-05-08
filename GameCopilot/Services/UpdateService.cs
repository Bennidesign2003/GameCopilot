using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using GameCopilot.Models;

namespace GameCopilot.Services;

public class UpdateService
{
    private const string UpdateJsonUrl =
        "https://github.com/Bennidesign2003/GodotRenderingAI/releases/download/MSFS24/update.json";

    private const string ReleasesApiUrl =
        "https://api.github.com/repos/Bennidesign2003/GodotRenderingAI/releases";

    // nvidia-mcp release channel — kept in its own GitHub repo, fetched at app start
    // so the MCP server upgrades independently of GameCopilot's own release cadence.
    private const string McpUpdateJsonUrl =
        "https://github.com/Bennidesign2003/nvidia-mcp/releases/latest/download/update.json";

    private static readonly HttpClient Http = new()
    {
        DefaultRequestHeaders =
        {
            { "User-Agent", "GameCopilot" }
        }
    };

    public string? LatestVersion { get; private set; }
    public string? DownloadUrl { get; private set; }
    public List<string> Changelog { get; private set; } = new();
    public bool IsUpdateAvailable { get; private set; }

    /// <summary>
    /// Checks update.json from GitHub release for new version.
    /// Throws <see cref="HttpRequestException"/> or <see cref="TaskCanceledException"/>
    /// on network failure — callers are expected to wrap in try/catch.
    /// </summary>
    public async Task<bool> CheckForUpdatesAsync(string currentVersion)
    {
        using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(30));
        var json = await Http.GetStringAsync(UpdateJsonUrl, cts.Token).ConfigureAwait(false);
        var data = JsonSerializer.Deserialize<UpdateJson>(json);
        if (data == null) return false;

        LatestVersion = data.LatestVersion;
        DownloadUrl = data.DownloadUrl;
        Changelog = data.Changelog ?? new List<string>();

        if (Version.TryParse(currentVersion, out var cur) &&
            Version.TryParse(LatestVersion, out var latest))
        {
            IsUpdateAvailable = latest > cur;
        }

        return IsUpdateAvailable;
    }

    /// <summary>
    /// Fetches all releases from GitHub API for the chronology view.
    /// </summary>
    public async Task<List<ReleaseEntry>> FetchReleasesAsync()
    {
        var entries = new List<ReleaseEntry>();

        var response = await Http.GetAsync(ReleasesApiUrl);
        if (!response.IsSuccessStatusCode)
            return entries;

        var releases = await response.Content.ReadFromJsonAsync<List<GitHubRelease>>();
        if (releases == null) return entries;

        for (var i = 0; i < releases.Count; i++)
        {
            var r = releases[i];
            var changes = new List<string>();

            // Parse body lines as changelog items
            if (!string.IsNullOrWhiteSpace(r.Body))
            {
                foreach (var line in r.Body.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var trimmed = line.TrimStart('-', '*', ' ', '\t', '\r');
                    if (!string.IsNullOrWhiteSpace(trimmed))
                        changes.Add(trimmed);
                }
            }

            // Also include changelog from update.json for latest
            if (i == 0 && Changelog.Count > 0 && changes.Count == 0)
                changes.AddRange(Changelog);

            var dateStr = "";
            if (DateTime.TryParse(r.PublishedAt, out var dt))
                dateStr = dt.ToString("dd MMM yyyy");

            entries.Add(new ReleaseEntry
            {
                TagName = r.TagName ?? "",
                Title = r.Name ?? r.TagName ?? "",
                Label = i == 0 ? "CURRENT BUILD" : "RELEASE",
                Date = dateStr,
                Body = r.Body ?? "",
                Changes = changes,
                IsCurrent = i == 0,
            });
        }

        return entries;
    }

    /// <summary>
    /// Downloads the update zip with progress reporting.
    /// Returns path to downloaded zip file.
    /// </summary>
    public async Task<string> DownloadUpdateAsync(IProgress<double>? progress = null)
    {
        if (string.IsNullOrEmpty(DownloadUrl))
            throw new InvalidOperationException("No update URL available");

        var tempDir = Path.Combine(Path.GetTempPath(), "GameCopilotUpdate");
        if (Directory.Exists(tempDir))
            Directory.Delete(tempDir, true);
        Directory.CreateDirectory(tempDir);

        var zipPath = Path.Combine(tempDir, "update.zip");

        using var response = await Http.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength ?? -1;

        await using var stream = await response.Content.ReadAsStreamAsync();
        await using var fs = File.Create(zipPath);

        var buffer = new byte[81920];
        long downloaded = 0;
        int read;
        while ((read = await stream.ReadAsync(buffer)) > 0)
        {
            await fs.WriteAsync(buffer.AsMemory(0, read));
            downloaded += read;
            if (totalBytes > 0)
                progress?.Report((double)downloaded / totalBytes);
        }

        return zipPath;
    }

    /// <summary>
    /// Extracts a downloaded update zip to target directory.
    /// </summary>
    public static void ExtractUpdate(string zipPath, string targetDir)
    {
        System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, targetDir, overwriteFiles: true);
    }

    /// <summary>
    /// Fetches nvidia-mcp's release manifest, compares against the local
    /// <c># __mcp_version__</c> marker, and (if newer + SHA256 matches) writes
    /// the downloaded server.py to <paramref name="serverPath"/>. Network and
    /// disk errors are swallowed — a failed online update is non-fatal because
    /// the embedded fallback already provides a working server.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the file on disk was replaced, <c>false</c> if it was
    /// already current or any step failed.
    /// </returns>
    public static async Task<bool> TryUpdateMcpServerAsync(string serverPath)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            var manifestJson = await Http.GetStringAsync(McpUpdateJsonUrl, cts.Token).ConfigureAwait(false);
            var manifest = JsonSerializer.Deserialize<McpUpdateJson>(manifestJson);
            if (manifest == null || string.IsNullOrWhiteSpace(manifest.Version)
                || string.IsNullOrWhiteSpace(manifest.DownloadUrl)
                || string.IsNullOrWhiteSpace(manifest.Sha256))
                return false;

            var localVersion = ReadMcpVersionFromFile(serverPath);
            if (Version.TryParse(localVersion, out var lv) && Version.TryParse(manifest.Version, out var rv))
            {
                if (rv <= lv) return false; // already up-to-date
            }

            var bytes = await Http.GetByteArrayAsync(manifest.DownloadUrl!, cts.Token).ConfigureAwait(false);

            using var sha = SHA256.Create();
            var hash = Convert.ToHexString(sha.ComputeHash(bytes)).ToLowerInvariant();
            if (!string.Equals(hash, manifest.Sha256!.Trim().ToLowerInvariant(), StringComparison.Ordinal))
                return false; // checksum mismatch — refuse to write a tampered file

            var dir = Path.GetDirectoryName(serverPath)!;
            var backupPath = Path.Combine(dir, "server.py.bak");
            if (File.Exists(serverPath))
                File.Copy(serverPath, backupPath, overwrite: true);

            var tmpPath = serverPath + ".new";
            await File.WriteAllBytesAsync(tmpPath, bytes, cts.Token).ConfigureAwait(false);
            File.Move(tmpPath, serverPath, overwrite: true);

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string ReadMcpVersionFromFile(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var buf = new byte[200];
            var n = fs.Read(buf, 0, buf.Length);
            var head = Encoding.UTF8.GetString(buf, 0, n);
            var m = Regex.Match(head, "__mcp_version__\\s*=\\s*\"([^\"]+)\"");
            return m.Success ? m.Groups[1].Value : "";
        }
        catch
        {
            return "";
        }
    }

    private class UpdateJson
    {
        public string? LatestVersion { get; set; }
        public string? DownloadUrl { get; set; }
        public List<string>? Changelog { get; set; }
    }

    private class McpUpdateJson
    {
        [JsonPropertyName("version")]
        public string? Version { get; set; }

        [JsonPropertyName("download_url")]
        public string? DownloadUrl { get; set; }

        [JsonPropertyName("sha256")]
        public string? Sha256 { get; set; }
    }

    private class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string? TagName { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("body")]
        public string? Body { get; set; }

        [JsonPropertyName("published_at")]
        public string? PublishedAt { get; set; }

        [JsonPropertyName("assets")]
        public List<GitHubAsset>? Assets { get; set; }
    }

    private class GitHubAsset
    {
        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("size")]
        public long Size { get; set; }

        [JsonPropertyName("browser_download_url")]
        public string? BrowserDownloadUrl { get; set; }
    }
}
