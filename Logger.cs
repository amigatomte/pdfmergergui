using System;
using System.IO;

namespace PdfMergerGui;

/// <summary>
/// Simple thread-safe file logger that appends to logs/app.log next to the EXE.
/// All methods are no-throw – logging failures are silently ignored so they
/// never affect the user experience.
/// </summary>
internal static class AppLogger
{
    private static string _logPath = string.Empty;
    private static readonly object _lock = new();

    /// <summary>
    /// Must be called once at startup (after AppSettings.Load).
    /// Creates the logs directory and sets the log file path.
    /// </summary>
    public static void Initialize()
    {
        try
        {
            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            _logPath = Path.Combine(logDir, "app.log");
        }
        catch
        {
            // If we can't create the log directory, logging is silently disabled.
        }
    }

    /// <summary>Appends a timestamped message to the log file.</summary>
    public static void Log(string message)
    {
        if (!AppSettings.Current.EnableLogging) return;
        if (string.IsNullOrEmpty(_logPath)) return;

        try
        {
            lock (_lock)
            {
                File.AppendAllText(
                    _logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
            }
        }
        catch { /* best-effort */ }
    }

    /// <summary>Logs an exception with its full stack trace.</summary>
    public static void LogException(string context, Exception ex)
    {
        Log($"ERROR in {context}: {ex.GetType().Name}: {ex.Message}");
        Log($"  StackTrace: {ex.StackTrace}");

        if (ex.InnerException != null)
            Log($"  Inner: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
    }
}
