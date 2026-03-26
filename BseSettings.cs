using System.Text.Json;

namespace BsePuller;

internal static class BseSettings
{
    private const string TokenPlaceholder = "PUT_YOUR_BILL_API_TOKEN_HERE";
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static AppUserSettings? _cachedSettings;

    public static readonly Uri BaseUri = new("https://gateway.prod.bill.com/connect/");

    public const string TransactionsPath = "v3/spend/transactions";
    public const string TransactionsApiFilter = "type:ne:DECLINE,syncStatus:eq:NOT_SYNCED,complete:eq:true";
    public const int PageSize = 50;

    public static string ApiToken => Load().ApiToken;

    public static bool IsConfigured =>
        !string.IsNullOrWhiteSpace(ApiToken) &&
        !ApiToken.Equals(TokenPlaceholder, StringComparison.OrdinalIgnoreCase);

    public static DateTimeOffset? GetLastUpdateCheckUtc()
    {
        return Load().LastUpdateCheckUtc;
    }

    public static void SaveLastUpdateCheckUtc(DateTimeOffset checkedAtUtc)
    {
        var settings = Load();
        settings.LastUpdateCheckUtc = checkedAtUtc.ToUniversalTime();
        Save(settings);
    }

    public static string GetExportsFolder()
    {
        return Path.Combine(AppContext.BaseDirectory, "CSV exports");
    }

    public static string GetInstalledAppFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs",
            "BsePuller");
    }

    public static string GetStartMenuShortcutFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Programs),
            "BSE Puller");
    }

    public static string GetUserSettingsFolder()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BsePuller");
    }

    public static string GetUserSettingsPath()
    {
        return Path.Combine(GetUserSettingsFolder(), "settings.json");
    }

    public static void SaveApiToken(string apiToken)
    {
        var cleanToken = apiToken?.Trim() ?? string.Empty;
        var settings = Load();
        settings.ApiToken = cleanToken;
        Save(settings);
    }

    public static void ClearApiToken()
    {
        var settings = Load();
        settings.ApiToken = string.Empty;

        if (settings.LastUpdateCheckUtc is not null)
        {
            Save(settings);
            return;
        }

        DeleteSettingsFile();
        _cachedSettings = new AppUserSettings();
    }

    public static void Reload()
    {
        _cachedSettings = null;
    }

    public static bool IsRunningInstalledCopy()
    {
        var currentDirectory = NormalizePath(AppContext.BaseDirectory);
        var installedDirectory = NormalizePath(GetInstalledAppFolder());
        return string.Equals(currentDirectory, installedDirectory, StringComparison.OrdinalIgnoreCase);
    }

    private static AppUserSettings Load()
    {
        if (_cachedSettings is not null)
        {
            return _cachedSettings;
        }

        var path = GetUserSettingsPath();
        if (!File.Exists(path))
        {
            _cachedSettings = new AppUserSettings();
            return _cachedSettings;
        }

        try
        {
            var json = File.ReadAllText(path);
            var settings = JsonSerializer.Deserialize<AppUserSettings>(json, JsonOptions) ?? new AppUserSettings();
            settings.ApiToken = NormalizeToken(settings.ApiToken);
            _cachedSettings = settings;
            return settings;
        }
        catch
        {
            _cachedSettings = new AppUserSettings();
            return _cachedSettings;
        }
    }

    private static void Save(AppUserSettings settings)
    {
        Directory.CreateDirectory(GetUserSettingsFolder());
        File.WriteAllText(GetUserSettingsPath(), JsonSerializer.Serialize(settings, JsonOptions));
        _cachedSettings = settings;
    }

    private static void DeleteSettingsFile()
    {
        var settingsPath = GetUserSettingsPath();
        if (File.Exists(settingsPath))
        {
            File.Delete(settingsPath);
        }

        var settingsFolder = GetUserSettingsFolder();
        if (Directory.Exists(settingsFolder) &&
            !Directory.EnumerateFileSystemEntries(settingsFolder).Any())
        {
            Directory.Delete(settingsFolder);
        }
    }

    private static string NormalizeToken(string? token)
    {
        var cleanToken = token?.Trim() ?? string.Empty;
        return string.Equals(cleanToken, TokenPlaceholder, StringComparison.OrdinalIgnoreCase)
            ? string.Empty
            : cleanToken;
    }

    private static string NormalizePath(string path)
    {
        return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private sealed class AppUserSettings
    {
        public string ApiToken { get; set; } = string.Empty;
        public DateTimeOffset? LastUpdateCheckUtc { get; set; }
    }
}
