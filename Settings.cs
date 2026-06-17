using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PdfMergerGui;

/// <summary>
/// Strongly-typed settings loaded from / saved to settings.json (next to the EXE).
/// </summary>
internal sealed class AppSettings
{
    // ── Properties (map 1:1 to JSON keys) ─────────────────────────────────────

    [JsonPropertyName("TempFolder")]
    public string TempFolder { get; set; } = @"%TEMP%\PdfMergerGui";

    [JsonPropertyName("DefaultOutputFolder")]
    public string DefaultOutputFolder { get; set; } = @"%USERPROFILE%\Documents";

    [JsonPropertyName("EnableLogging")]
    public bool EnableLogging { get; set; } = true;

    [JsonPropertyName("EnableFastMode")]
    public bool EnableFastMode { get; set; } = true;

    [JsonPropertyName("RetryWordWithFreshInstance")]
    public bool RetryWordWithFreshInstance { get; set; } = true;

    [JsonPropertyName("ForceLocalPrinterForWord")]
    public bool ForceLocalPrinterForWord { get; set; } = true;

    [JsonPropertyName("PreferredWordPrinter")]
    public string PreferredWordPrinter { get; set; } = "Microsoft Print to PDF";

    // ── Static API ─────────────────────────────────────────────────────────────

    /// <summary>The loaded (or default) settings instance.</summary>
    public static AppSettings Current { get; private set; } = new();

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        AllowTrailingCommas = true,
        ReadCommentHandling = JsonCommentHandling.Skip
    };

    private static string SettingsFilePath =>
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");

    /// <summary>
    /// Loads settings.json from the application directory.
    /// Creates the file with defaults if it does not exist.
    /// Falls back to defaults silently on any parse error.
    /// </summary>
    public static void Load()
    {
        try
        {
            if (File.Exists(SettingsFilePath))
            {
                string json = File.ReadAllText(SettingsFilePath);
                Current = JsonSerializer.Deserialize<AppSettings>(json, SerializerOptions)
                          ?? new AppSettings();
            }
            else
            {
                Current = new AppSettings();
                Save(); // write defaults so the user can see and edit them
            }
        }
        catch
        {
            Current = new AppSettings();
        }
    }

    /// <summary>Persists the current settings to settings.json.</summary>
    public static void Save()
    {
        try
        {
            string json = JsonSerializer.Serialize(Current, SerializerOptions);
            File.WriteAllText(SettingsFilePath, json);
        }
        catch { /* best-effort */ }
    }

    // ── Helpers ────────────────────────────────────────────────────────────────

    /// <summary>Returns the temp folder path with environment variables expanded.</summary>
    public static string GetTempFolder() =>
        Environment.ExpandEnvironmentVariables(Current.TempFolder);

    /// <summary>Returns the default output folder with environment variables expanded.</summary>
    public static string GetDefaultOutputFolder() =>
        Environment.ExpandEnvironmentVariables(Current.DefaultOutputFolder);
}
