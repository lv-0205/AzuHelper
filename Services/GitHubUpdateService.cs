using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using AzuHelper_v2.Models;

namespace AzuHelper_v2.Services;

public sealed class GitHubUpdateService
{
    private readonly HttpClient _httpClient;

    public GitHubUpdateService(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
        if (!_httpClient.DefaultRequestHeaders.UserAgent.Any())
        {
            _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("AzuHelper_v2");
        }
    }

    public async Task<UpdateResult> CheckAndPrepareUpdateAsync(AppConfig config, IProgress<string>? progress, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(config.GitHubOwner) || string.IsNullOrWhiteSpace(config.GitHubRepo))
        {
            return UpdateResult.NotConfigured();
        }

        progress?.Report("Suche nach Updates...");
        var release = await GetLatestReleaseAsync(config, cancellationToken).ConfigureAwait(false);
        if (release is null)
        {
            return UpdateResult.Failed("Release konnte nicht geladen werden.");
        }

        var latestVersion = ParseVersion(release.TagName);
        if (latestVersion is null)
        {
            return UpdateResult.Failed("Release-Version ist ungültig.");
        }

        var currentVersion = GetCurrentVersion();
        if (currentVersion is null)
        {
            return UpdateResult.Failed("Aktuelle Version konnte nicht ermittelt werden.");
        }

        if (latestVersion <= currentVersion)
        {
            return UpdateResult.UpToDate();
        }

        var asset = SelectAsset(release, config.GitHubAssetName);
        if (asset is null)
        {
            return UpdateResult.Failed("Kein passendes Update-Asset gefunden.");
        }

        progress?.Report($"Lade Version {latestVersion}...");
        var updateRoot = Path.Combine(Path.GetTempPath(), "AzuHelperUpdate");
        Directory.CreateDirectory(updateRoot);

        var downloadPath = Path.Combine(updateRoot, asset.Name);
        await DownloadAssetAsync(asset.DownloadUrl, downloadPath, cancellationToken).ConfigureAwait(false);

        if (!downloadPath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return UpdateResult.Failed("Update-Asset muss eine ZIP-Datei sein.");
        }

        var extractPath = Path.Combine(updateRoot, "extracted");
        if (Directory.Exists(extractPath))
        {
            Directory.Delete(extractPath, true);
        }
        ZipFile.ExtractToDirectory(downloadPath, extractPath, true);

        var sourcePath = ResolveExtractRoot(extractPath);
        var exePath = Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(exePath))
        {
            return UpdateResult.Failed("Ausführbare Datei konnte nicht ermittelt werden.");
        }

        var targetDir = Path.GetDirectoryName(exePath);
        if (string.IsNullOrWhiteSpace(targetDir))
        {
            return UpdateResult.Failed("Zielverzeichnis konnte nicht ermittelt werden.");
        }

        if (!File.Exists(Path.Combine(sourcePath, Path.GetFileName(exePath))))
        {
            return UpdateResult.Failed("Update-Paket enthält keine ausführbare Datei.");
        }

        var scriptPath = Path.Combine(updateRoot, $"update-{Process.GetCurrentProcess().Id}.cmd");
        WriteUpdateScript(scriptPath, sourcePath, targetDir, exePath, Process.GetCurrentProcess().Id);

        return UpdateResult.ReadyToApply(scriptPath, latestVersion.ToString());
    }

    public void LaunchUpdateScript(string scriptPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c \"{scriptPath}\"",
            CreateNoWindow = true,
            UseShellExecute = false
        };

        Process.Start(startInfo);
    }

    private async Task<GitHubRelease?> GetLatestReleaseAsync(AppConfig config, CancellationToken cancellationToken)
    {
        var url = $"https://api.github.com/repos/{config.GitHubOwner}/{config.GitHubRepo}/releases/latest";
        using var response = await _httpClient.GetAsync(url, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        return await JsonSerializer.DeserializeAsync<GitHubRelease>(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    private static Version? ParseVersion(string? tagName)
    {
        if (string.IsNullOrWhiteSpace(tagName))
        {
            return null;
        }

        var trimmed = tagName.Trim();
        if (trimmed.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed[1..];
        }

        return Version.TryParse(trimmed, out var version) ? version : null;
    }

    private static Version? GetCurrentVersion()
    {
        var version = typeof(GitHubUpdateService).Assembly.GetName().Version;
        return version is null ? null : new Version(version.Major, version.Minor, version.Build, version.Revision);
    }

    private static GitHubAsset? SelectAsset(GitHubRelease release, string? preferredName)
    {
        if (!string.IsNullOrWhiteSpace(preferredName))
        {
            return release.Assets.FirstOrDefault(asset => string.Equals(asset.Name, preferredName, StringComparison.OrdinalIgnoreCase));
        }

        return release.Assets.FirstOrDefault(asset => asset.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase));
    }

    private static async Task DownloadAssetAsync(string url, string destination, CancellationToken cancellationToken)
    {
        using var response = await new HttpClient().GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        await using var fileStream = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None);
        await stream.CopyToAsync(fileStream, cancellationToken).ConfigureAwait(false);
    }

    private static string ResolveExtractRoot(string extractPath)
    {
        var directories = Directory.GetDirectories(extractPath);
        var files = Directory.GetFiles(extractPath);
        if (files.Length == 0 && directories.Length == 1)
        {
            return directories[0];
        }

        return extractPath;
    }

    private static void WriteUpdateScript(string scriptPath, string sourceDir, string targetDir, string exePath, int processId)
    {
        var script = $"@echo off{Environment.NewLine}" +
                     "setlocal" + Environment.NewLine +
                     $"set PID={processId}{Environment.NewLine}" +
                     ":wait" + Environment.NewLine +
                     $"tasklist /FI \"PID eq %PID%\" | find \"%PID%\" >nul{Environment.NewLine}" +
                     "if not errorlevel 1 (" + Environment.NewLine +
                     "  timeout /t 1 >nul" + Environment.NewLine +
                     "  goto wait" + Environment.NewLine +
                     ")" + Environment.NewLine +
                     $"xcopy /E /Y /I \"{sourceDir}\" \"{targetDir}\"{Environment.NewLine}" +
                     $"start \"\" \"{exePath}\"{Environment.NewLine}" +
                     "exit /b 0" + Environment.NewLine;

        File.WriteAllText(scriptPath, script);
    }

    private sealed record GitHubRelease(
        [property: JsonPropertyName("tag_name")] string TagName,
        [property: JsonPropertyName("assets")] List<GitHubAsset> Assets);

    private sealed record GitHubAsset(
        [property: JsonPropertyName("name")] string Name,
        [property: JsonPropertyName("browser_download_url")] string DownloadUrl);
}

public sealed record UpdateResult(UpdateResultKind Kind, string Message, string? ScriptPath = null, string? LatestVersion = null)
{
    public static UpdateResult NotConfigured() => new(UpdateResultKind.NotConfigured, "GitHub-Update ist nicht konfiguriert.");

    public static UpdateResult UpToDate() => new(UpdateResultKind.UpToDate, "App ist aktuell.");

    public static UpdateResult Failed(string message) => new(UpdateResultKind.Failed, message);

    public static UpdateResult ReadyToApply(string scriptPath, string latestVersion) =>
        new(UpdateResultKind.ReadyToApply, $"Update {latestVersion} wird installiert...", scriptPath, latestVersion);
}

public enum UpdateResultKind
{
    NotConfigured,
    UpToDate,
    ReadyToApply,
    Failed
}
