using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using GameCopilot.Models;

namespace GameCopilot.Services;

public class UpdateService
{
    private const string UpdateJsonUrl =
        "https://github.com/Bennidesign2003/GodotRenderingAI/releases/download/MSFS24/update.json";

    private const string ReleasesApiUrl =
        "https://api.github.com/repos/Bennidesign2003/GodotRenderingAI/releases";

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

    private class UpdateJson
    {
        public string? LatestVersion { get; set; }
        public string? DownloadUrl { get; set; }
        public List<string>? Changelog { get; set; }
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
