using System.Windows.Forms;

namespace PdfMergerGui;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        // Load settings before anything else so all subsystems see them.
        AppSettings.Load();

        // Initialise the logger (creates the logs/ directory if needed).
        AppLogger.Initialize();
        AppLogger.Log("=== PdfMergerGui started ===");

        Application.Run(new MainForm());

        AppLogger.Log("=== PdfMergerGui exited ===");
    }
}
