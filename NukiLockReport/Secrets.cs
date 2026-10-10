using System.Text.Json;

namespace NukiLockReport;

/// <summary>
/// Reads the API token and settings without them ever living in the project folder.
/// Each value comes from its environment variable first, then from .NET user-secrets.
/// </summary>
public static class Secrets
{
    // Must match <UserSecretsId> in NukiLockReport.csproj (shared with NukiApiProbe)
    private const string UserSecretsId = "37283e67-de66-45d7-8338-c1dda63a28ed";

    public const string TokenKey = "Nuki:ApiToken";
    public const string TokenEnv = "NUKI_API_TOKEN";
    public const string TimeZoneKey = "Nuki:TimeZone";
    public const string TimeZoneEnv = "NUKI_TIME_ZONE";

    /// <summary>The user-secrets file (outside the project folder, under %APPDATA%).</summary>
    public static string FilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Microsoft", "UserSecrets", UserSecretsId, "secrets.json");

    public static string? GetToken() => Get(TokenKey, TokenEnv).Value;

    /// <summary>Returns the value and where it came from, or (null, null) when it isn't set.</summary>
    public static (string? Value, string? Source) Get(string key, string envVar)
    {
        string? fromEnv = Environment.GetEnvironmentVariable(envVar);
        if (!string.IsNullOrWhiteSpace(fromEnv))
            return (fromEnv.Trim(), $"{envVar} environment variable");

        if (!File.Exists(FilePath))
            return (null, null);

        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(FilePath));
            string? value = doc.RootElement.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.String
                ? v.GetString()?.Trim()
                : null;
            return string.IsNullOrEmpty(value) ? (null, null) : (value, $"\"{key}\" in user-secrets");
        }
        catch (Exception ex) when (ex is JsonException or IOException)
        {
            return (null, null);
        }
    }
}
