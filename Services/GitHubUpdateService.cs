using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using AzuHelper_v2.Models;

namespace AzuHelper_v2.Services;

public sealed class GitHubUpdateService
{
    private readonly HttpClient _httpClient;

    public GitHubUpdateService(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
        _httpClient.Timeout = TimeSpan.FromMinutes(5);
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

        var latestVersion = ParseVersion(release.TagName, release.Name);
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

        var asset = SelectAsset(release);
        if (asset is null)
        {
            return UpdateResult.Failed("Kein passendes Update-Asset gefunden.");
        }

        progress?.Report($"Lade Version {latestVersion}...");
        var updateRoot = Path.Combine(Path.GetTempPath(), "AzuHelperUpdate");
        Directory.CreateDirectory(updateRoot);

        var downloadPath = Path.Combine(updateRoot, asset.Name);
        await DownloadAssetAsync(asset.DownloadUrl, downloadPath, progress, cancellationToken).ConfigureAwait(false);

        if (!downloadPath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return UpdateResult.Failed("Update-Asset muss eine ZIP-Datei sein.");
        }
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

        var extractPath = Path.Combine(updateRoot, "extracted");
        progress?.Report("Bereite Installation vor...");
        var scriptPath = Path.Combine(updateRoot, $"update-{Process.GetCurrentProcess().Id}.cmd");
        var ps1Path = Path.ChangeExtension(scriptPath, ".ps1");
        WriteUpdateScript(scriptPath, ps1Path, downloadPath, extractPath, targetDir, exePath, Process.GetCurrentProcess().Id);

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

    private static Version? ParseVersion(string? tagName, string? releaseName)
    {
        var source = string.IsNullOrWhiteSpace(tagName) ? releaseName : tagName;
        if (string.IsNullOrWhiteSpace(source))
        {
            return null;
        }

        var match = Regex.Match(source, @"\d+(\.\d+){0,3}");
        return match.Success && Version.TryParse(match.Value, out var version) ? version : null;
    }

    private static Version? GetCurrentVersion()
    {
        var version = typeof(GitHubUpdateService).Assembly.GetName().Version;
        return version is null ? null : new Version(version.Major, version.Minor, version.Build, version.Revision);
    }

    private static GitHubAsset? SelectAsset(GitHubRelease release)
        => release.Assets.FirstOrDefault(asset => asset.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase));

    private async Task DownloadAssetAsync(string url, string destination, IProgress<string>? progress, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        var contentLength = response.Content.Headers.ContentLength;
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        await using var fileStream = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None);

        var buffer = new byte[81920];
        long totalRead = 0;
        int read;
        while (true)
        {
            var readTask = stream.ReadAsync(buffer, cancellationToken).AsTask();
            var completed = await Task.WhenAny(readTask, Task.Delay(TimeSpan.FromSeconds(10), cancellationToken)).ConfigureAwait(false);
            if (completed != readTask)
            {
                break;
            }

            read = await readTask.ConfigureAwait(false);
            if (read <= 0)
            {
                break;
            }

            await fileStream.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            if (contentLength.HasValue)
            {
                totalRead += read;
                var percent = Math.Clamp((int)(totalRead * 100 / contentLength.Value), 0, 99);
                progress?.Report($"Lade Version {percent}%...");
                if (totalRead >= contentLength.Value)
                {
                    break;
                }
            }
        }

        if (contentLength.HasValue)
        {
            progress?.Report("Lade Version 99%...");
        }

        progress?.Report("Download abgeschlossen...");
    }

    private static void WriteUpdateScript(string scriptPath, string ps1Path, string zipPath, string extractPath, string targetDir, string exePath, int processId)
    {
        var escapedZip = EscapeForPowerShell(zipPath);
        var escapedExtract = EscapeForPowerShell(extractPath);
        var escapedTarget = EscapeForPowerShell(targetDir);
        var ps1 = $"$zip = '{escapedZip}'{Environment.NewLine}" +
                  $"$extract = '{escapedExtract}'{Environment.NewLine}" +
                  "if (Test-Path $extract) { Remove-Item $extract -Recurse -Force }" + Environment.NewLine +
                  "Expand-Archive -Path $zip -DestinationPath $extract -Force" + Environment.NewLine +
                  "$dirs = Get-ChildItem -Path $extract -Directory" + Environment.NewLine +
                  "$files = Get-ChildItem -Path $extract -File" + Environment.NewLine +
                  "$src = if ($files.Count -eq 0 -and $dirs.Count -eq 1) { $dirs[0].FullName } else { $extract }" + Environment.NewLine +
                  $"robocopy $src '{escapedTarget}' /E /R:1 /W:1 /NP /NFL /NDL /NJH /NJS | Out-Null" + Environment.NewLine;

        File.WriteAllText(ps1Path, ps1);

        var script = $"@echo off{Environment.NewLine}" +
                     "setlocal" + Environment.NewLine +
                     $"set PID={processId}{Environment.NewLine}" +
                     ":wait" + Environment.NewLine +
                     $"tasklist /FI \"PID eq %PID%\" | find \"%PID%\" >nul{Environment.NewLine}" +
                     "if not errorlevel 1 (" + Environment.NewLine +
                     "  timeout /t 1 >nul" + Environment.NewLine +
                     "  goto wait" + Environment.NewLine +
                     ")" + Environment.NewLine +
                     $"powershell -NoProfile -ExecutionPolicy Bypass -File \"{ps1Path}\"{Environment.NewLine}" +
                     $"start \"\" \"{exePath}\"{Environment.NewLine}" +
                     "exit /b 0" + Environment.NewLine;

        File.WriteAllText(scriptPath, script);
    }

    private static string EscapeForPowerShell(string value) => value.Replace("'", "''");

    private sealed record GitHubRelease(
        [property: JsonPropertyName("tag_name")] string TagName,
        [property: JsonPropertyName("name")] string? Name,
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
