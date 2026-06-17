using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PdfMergerGui;

/// <summary>
/// Extracts bundled native tools (qpdf.exe) from the assembly's embedded resources
/// to a well-known temp location, re-using existing extractions on subsequent runs.
/// </summary>
internal static class EmbeddedTools
{
    /// <summary>
    /// Extracts qpdf.exe to %TEMP%\PdfMergerGui\qpdf.exe (or the configured temp folder)
    /// and returns the full path to the executable.
    ///
    /// The file is extracted only once per machine/version; subsequent calls
    /// return the cached path immediately.
    /// </summary>
    public static string ExtractQpdf()
    {
        string targetDir = AppSettings.GetTempFolder();
        string targetPath = Path.Combine(targetDir, "qpdf.exe");

        if (File.Exists(targetPath))
        {
            AppLogger.Log($"qpdf.exe already present at: {targetPath}");
            return targetPath;
        }

        Directory.CreateDirectory(targetDir);

        // Locate the embedded resource regardless of how MSBuild named it.
        // The default name is "<RootNamespace>.qpdf.exe", e.g. "PdfMergerGui.qpdf.exe",
        // but we search by suffix so the code works even if the project is renamed.
        Assembly assembly = Assembly.GetExecutingAssembly();
        string? resourceName = assembly
            .GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith("qpdf.exe", StringComparison.OrdinalIgnoreCase));

        if (resourceName is null)
        {
            string available = string.Join(", ", assembly.GetManifestResourceNames());
            throw new InvalidOperationException(
                "qpdf.exe was not found as an embedded resource. " +
                "Please add qpdf.exe to the project root and set its Build Action to " +
                "'Embedded Resource' in Visual Studio (or add it to the .csproj as " +
                "<EmbeddedResource Include=\"qpdf.exe\" />).\n\n" +
                $"Available embedded resources: [{available}]");
        }

        AppLogger.Log($"Extracting embedded resource '{resourceName}' to: {targetPath}");

        using Stream resourceStream = assembly.GetManifestResourceStream(resourceName)!;
        using FileStream fileStream = new(targetPath, FileMode.Create, FileAccess.Write, FileShare.None);
        resourceStream.CopyTo(fileStream);

        AppLogger.Log("qpdf.exe extracted successfully.");
        return targetPath;
    }
}
