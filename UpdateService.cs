using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;

namespace BsePuller;

internal sealed record UpdateCheckResult(
    bool CheckedSuccessfully,
    bool IsUpdateAvailable,
    string CurrentTag,
    string? LatestTag,
    string? DownloadUrl,
    string? ErrorMessage);

internal sealed record InstallerDownloadResult(
    bool Success,
    string? FilePath,
    string? ErrorMessage);

internal static class UpdateService
{
    private const string LatestReleaseEndpoint = "https://api.github.com/repos/mcjeston/bse-puller/releases/latest";
    private const string InstallerAssetName = "BsePullerSetup.exe";
    private static readonly HttpClient HttpClient = CreateHttpClient();

    public static string GetCurrentReleaseTag()
    {
        var informationalVersion = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;

        var cleanVersion = string.IsNullOrWhiteSpace(informationalVersion)
            ? Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0.0"
            : informationalVersion.Split('+', StringSplitOptions.TrimEntries)[0];

        return cleanVersion.StartsWith("v", StringComparison.OrdinalIgnoreCase)
            ? cleanVersion
            : $"v{cleanVersion}";
    }

    public static async Task<UpdateCheckResult> CheckForUpdatesAsync(CancellationToken cancellationToken)
    {
        var currentTag = GetCurrentReleaseTag();

        try
        {
            using var response = await HttpClient.GetAsync(LatestReleaseEndpoint, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                return new UpdateCheckResult(
                    CheckedSuccessfully: false,
                    IsUpdateAvailable: false,
                    CurrentTag: currentTag,
                    LatestTag: null,
                    DownloadUrl: null,
                    ErrorMessage: $"GitHub returned {(int)response.StatusCode} {response.ReasonPhrase}.");
            }

            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            if (!root.TryGetProperty("tag_name", out var tagProperty))
            {
                return new UpdateCheckResult(
                    CheckedSuccessfully: false,
                    IsUpdateAvailable: false,
                    CurrentTag: currentTag,
                    LatestTag: null,
                    DownloadUrl: null,
                    ErrorMessage: "GitHub response did not include tag_name.");
            }

            var latestTag = tagProperty.GetString();
            if (string.IsNullOrWhiteSpace(latestTag))
            {
                return new UpdateCheckResult(
                    CheckedSuccessfully: false,
                    IsUpdateAvailable: false,
                    CurrentTag: currentTag,
                    LatestTag: null,
                    DownloadUrl: null,
                    ErrorMessage: "GitHub response contained an empty release tag.");
            }

            var downloadUrl = FindInstallerDownloadUrl(root);
            var isUpdateAvailable = IsNewerRelease(latestTag, currentTag);

            return new UpdateCheckResult(
                CheckedSuccessfully: true,
                IsUpdateAvailable: isUpdateAvailable,
                CurrentTag: currentTag,
                LatestTag: latestTag,
                DownloadUrl: downloadUrl,
                ErrorMessage: null);
        }
        catch (Exception ex)
        {
            return new UpdateCheckResult(
                CheckedSuccessfully: false,
                IsUpdateAvailable: false,
                CurrentTag: currentTag,
                LatestTag: null,
                DownloadUrl: null,
                ErrorMessage: ex.Message);
        }
    }

    public static async Task<InstallerDownloadResult> DownloadInstallerAsync(string downloadUrl, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(downloadUrl))
        {
            return new InstallerDownloadResult(false, null, "The release did not contain a valid installer download URL.");
        }

        try
        {
            var targetPath = Path.Combine(
                Path.GetTempPath(),
                $"BsePullerSetup-{DateTime.Now:yyyyMMdd-HHmmss}.exe");

            using var response = await HttpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return new InstallerDownloadResult(
                    false,
                    null,
                    $"Installer download failed with {(int)response.StatusCode} {response.ReasonPhrase}.");
            }

            await using var responseStream = await response.Content.ReadAsStreamAsync(cancellationToken);
            await using var targetStream = File.Create(targetPath);
            await responseStream.CopyToAsync(targetStream, cancellationToken);

            return new InstallerDownloadResult(true, targetPath, null);
        }
        catch (Exception ex)
        {
            return new InstallerDownloadResult(false, null, ex.Message);
        }
    }

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(45)
        };

        client.DefaultRequestHeaders.UserAgent.Clear();
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("BsePuller", "1.0"));
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));

        return client;
    }

    private static string? FindInstallerDownloadUrl(JsonElement releaseRoot)
    {
        if (!releaseRoot.TryGetProperty("assets", out var assetsElement) || assetsElement.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var asset in assetsElement.EnumerateArray())
        {
            if (asset.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var name = asset.TryGetProperty("name", out var nameProperty) ? nameProperty.GetString() : null;
            if (!string.Equals(name, InstallerAssetName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (asset.TryGetProperty("browser_download_url", out var urlProperty))
            {
                return urlProperty.GetString();
            }
        }

        return null;
    }

    private static bool IsNewerRelease(string latestTag, string currentTag)
    {
        if (TryParseTagVersion(latestTag, out var latestParts) && TryParseTagVersion(currentTag, out var currentParts))
        {
            return CompareVersionParts(latestParts, currentParts) > 0;
        }

        return !string.Equals(latestTag, currentTag, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseTagVersion(string tag, out IReadOnlyList<int> parts)
    {
        var cleanTag = tag.Trim();
        if (cleanTag.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            cleanTag = cleanTag[1..];
        }

        var tokens = cleanTag.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (tokens.Length == 0)
        {
            parts = Array.Empty<int>();
            return false;
        }

        var values = new List<int>(tokens.Length);
        foreach (var token in tokens)
        {
            if (!int.TryParse(token, out var numericPart))
            {
                parts = Array.Empty<int>();
                return false;
            }

            values.Add(numericPart);
        }

        parts = values;
        return true;
    }

    private static int CompareVersionParts(IReadOnlyList<int> left, IReadOnlyList<int> right)
    {
        var maxLength = Math.Max(left.Count, right.Count);
        for (var index = 0; index < maxLength; index++)
        {
            var leftPart = index < left.Count ? left[index] : 0;
            var rightPart = index < right.Count ? right[index] : 0;

            var comparison = leftPart.CompareTo(rightPart);
            if (comparison != 0)
            {
                return comparison;
            }
        }

        return 0;
    }
}
