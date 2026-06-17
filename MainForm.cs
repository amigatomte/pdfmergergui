using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Word       = Microsoft.Office.Interop.Word;
using Excel      = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PdfMergerGui;

internal sealed class MainForm : Form
{
    // ── Supported file extensions ──────────────────────────────────────────────

    private static readonly HashSet<string> SupportedExtensions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ".pdf",
            ".doc", ".docx",
            ".xls", ".xlsx",
            ".ppt", ".pptx",
            ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"
        };

    private const string OpenFileFilter =
        "Supported Files|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;" +
                          "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|" +
        "PDF Files|*.pdf|" +
        "Word Files|*.doc;*.docx|" +
        "Excel Files|*.xls;*.xlsx|" +
        "PowerPoint Files|*.ppt;*.pptx|" +
        "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|" +
        "All Files|*.*";

    // ── UI controls ───────────────────────────────────────────────────────────

    private ListBox     _lstFiles    = null!;
    private Button      _btnAdd      = null!;
    private Button      _btnRemove   = null!;
    private Button      _btnMoveUp   = null!;
    private Button      _btnMoveDown = null!;
    private Button      _btnAbout    = null!;
    private Button      _btnMerge    = null!;
    private ProgressBar _progressBar = null!;
    private Label       _lblStatus   = null!;

    // ── Constructor ───────────────────────────────────────────────────────────

    public MainForm()
    {
        InitializeComponent();
        AppLogger.Log("MainForm initialised.");
    }

    // ── UI setup ──────────────────────────────────────────────────────────────

    private void InitializeComponent()
    {
        SuspendLayout();

        Text            = "PDF Merger";
        Size            = new Size(760, 520);
        MinimumSize     = new Size(560, 380);
        StartPosition   = FormStartPosition.CenterScreen;
        Font            = new Font("Segoe UI", 9f);

        // ── Bottom panel: status label + progress bar ─────────────────────────
        var pnlBottom = new Panel
        {
            Dock    = DockStyle.Bottom,
            Height  = 52,
            Padding = new Padding(6, 4, 6, 4)
        };

        _progressBar = new ProgressBar
        {
            Dock   = DockStyle.Bottom,
            Height = 20,
            Minimum = 0,
            Maximum = 100,
            Value   = 0,
            Style   = ProgressBarStyle.Continuous
        };

        _lblStatus = new Label
        {
            Dock      = DockStyle.Fill,
            Text      = "Ready. Add files and click Merge.",
            TextAlign = ContentAlignment.MiddleLeft
        };

        pnlBottom.Controls.Add(_lblStatus);
        pnlBottom.Controls.Add(_progressBar);

        // ── Right panel: action buttons ───────────────────────────────────────
        var pnlButtons = new Panel
        {
            Dock  = DockStyle.Right,
            Width = 170
        };

        _btnAdd      = MakeButton("Add...",    0);
        _btnRemove   = MakeButton("Remove",    1);
        _btnMoveUp   = MakeButton("Move Up",   2);
        _btnMoveDown = MakeButton("Move Down", 3);
        _btnAbout    = MakeButton("About",     4);
        _btnMerge    = MakeButton("Merge...",  6);  // visual gap before merge

        _btnMerge.Font      = new Font(_btnMerge.Font, FontStyle.Bold);
        _btnMerge.BackColor = Color.FromArgb(0, 120, 212);
        _btnMerge.ForeColor = Color.White;
        _btnMerge.FlatStyle = FlatStyle.Flat;
        _btnMerge.FlatAppearance.BorderColor = Color.FromArgb(0, 100, 180);
        _btnMerge.UseCompatibleTextRendering = true;
        _btnMerge.Padding = new Padding(0, 0, 0, 2);

        LayoutButtons(pnlButtons, _btnAdd, _btnRemove, _btnMoveUp, _btnMoveDown, _btnAbout, _btnMerge);

        _btnAdd.Click      += BtnAdd_Click;
        _btnRemove.Click   += BtnRemove_Click;
        _btnMoveUp.Click   += BtnMoveUp_Click;
        _btnMoveDown.Click += BtnMoveDown_Click;
        _btnAbout.Click    += BtnAbout_Click;
        _btnMerge.Click    += BtnMerge_Click;

        pnlButtons.Controls.AddRange(
            new Control[] { _btnAdd, _btnRemove, _btnMoveUp, _btnMoveDown, _btnAbout, _btnMerge });

        // ── File list box ─────────────────────────────────────────────────────
        _lstFiles = new ListBox
        {
            Dock              = DockStyle.Fill,
            SelectionMode     = SelectionMode.MultiExtended,
            Font              = new Font("Consolas", 8.5f),
            AllowDrop         = true,
            IntegralHeight    = false,
            HorizontalScrollbar = true
        };

        _lstFiles.SelectedIndexChanged += (_, _) => UpdateButtonStates();
        _lstFiles.DragEnter            += LstFiles_DragEnter;
        _lstFiles.DragDrop             += LstFiles_DragDrop;

        // ── Add controls (order determines dock priority) ─────────────────────
        // Bottom and Right are added before Fill so Fill takes the remaining space.
        Controls.Add(pnlBottom);   // DockStyle.Bottom
        Controls.Add(pnlButtons);  // DockStyle.Right
        Controls.Add(_lstFiles);   // DockStyle.Fill

        ResumeLayout(performLayout: true);
        UpdateButtonStates();
    }

    /// <summary>Creates a button inside the right buttons panel. Final size/position is applied by LayoutButtons.</summary>
    private static Button MakeButton(string text, int slotIndex)
    {
        return new Button
        {
            Text     = text,
            Location = new Point(10, 10 + slotIndex * 40),
            Size     = new Size(140, 36),
            Anchor   = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            UseCompatibleTextRendering = true
        };
    }

    private static void LayoutButtons(Panel panel, params Button[] buttons)
    {
        const int sideMargin = 10;
        const int topMargin = 10;
        const int rowGap = 8;
        const int specialMergeGap = 12;

        // Measure required content size for each button and use the largest width/height.
        int requiredTextWidth = 0;
        int requiredButtonHeight = 36;

        foreach (var b in buttons)
        {
            Size text = TextRenderer.MeasureText(
                b.Text + "  ",
                b.Font,
                new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);

            requiredTextWidth = Math.Max(requiredTextWidth, text.Width);

            // Extra vertical padding prevents clipping for bold labels like Merge.
            int h = Math.Max(38, text.Height + 18);
            requiredButtonHeight = Math.Max(requiredButtonHeight, h);
        }

        int buttonWidth = requiredTextWidth + 24;
        int panelWidth = buttonWidth + sideMargin * 2;
        panel.Width = Math.Max(panel.Width, panelWidth);

        int y = topMargin;
        for (int i = 0; i < buttons.Length; i++)
        {
            var b = buttons[i];
            b.Location = new Point(sideMargin, y);
            b.Size = new Size(buttonWidth, requiredButtonHeight);

            y += requiredButtonHeight + rowGap;

            // Keep a visual gap before the Merge button (last one).
            if (i == buttons.Length - 2)
                y += specialMergeGap;
        }
    }

    // ── Button-state management ───────────────────────────────────────────────

    private void UpdateButtonStates()
    {
        bool hasItems    = _lstFiles.Items.Count > 0;
        bool hasSelected = _lstFiles.SelectedIndices.Count > 0;
        int  firstSel    = hasSelected ? _lstFiles.SelectedIndices[0] : -1;
        int  lastSel     = hasSelected ? _lstFiles.SelectedIndices[^1] : -1;

        _btnRemove.Enabled   = hasSelected;
        _btnMoveUp.Enabled   = hasSelected && firstSel > 0;
        _btnMoveDown.Enabled = hasSelected && lastSel < _lstFiles.Items.Count - 1;
        _btnMerge.Enabled    = hasItems;
    }

    // ── Event handlers: Add / Remove / Move ──────────────────────────────────

    private void BtnAdd_Click(object? sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog
        {
            Title       = "Add Files",
            Filter      = OpenFileFilter,
            Multiselect = true
        };
        if (ofd.ShowDialog(this) == DialogResult.OK)
            AddFilesToList(ofd.FileNames);
    }

    private void BtnRemove_Click(object? sender, EventArgs e)
    {
        // Remove in reverse order to keep indices stable.
        var indices = _lstFiles.SelectedIndices
            .Cast<int>()
            .OrderByDescending(i => i)
            .ToList();

        _lstFiles.BeginUpdate();
        foreach (int idx in indices)
            _lstFiles.Items.RemoveAt(idx);
        _lstFiles.EndUpdate();

        UpdateButtonStates();
    }

    private void BtnMoveUp_Click(object? sender, EventArgs e)
    {
        MoveSelected(-1);
    }

    private void BtnMoveDown_Click(object? sender, EventArgs e)
    {
        MoveSelected(+1);
    }

    private void BtnAbout_Click(object? sender, EventArgs e)
    {
        string assemblyVersion =
            Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "Unknown";
        string productVersion = Application.ProductVersion;
        string today = DateTime.Now.ToString("yyyy-MM-dd");

        MessageBox.Show(
            "PdfMergerGui\n\n" +
            "Author: Christian Roth\n" +
            $"Date: {today}\n" +
            $"Version: {assemblyVersion}\n" +
            $"Product Version: {productVersion}",
            "About PdfMergerGui",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void MoveSelected(int direction)
    {
        var indices = _lstFiles.SelectedIndices.Cast<int>().OrderBy(i => i).ToList();
        if (direction > 0) indices.Reverse();

        _lstFiles.BeginUpdate();
        foreach (int idx in indices)
        {
            int newIdx = idx + direction;
            if (newIdx < 0 || newIdx >= _lstFiles.Items.Count) break;

            object item = _lstFiles.Items[idx];
            _lstFiles.Items.RemoveAt(idx);
            _lstFiles.Items.Insert(newIdx, item);
        }

        // Restore selection at new positions.
        _lstFiles.SelectedIndices.Clear();
        foreach (int idx in indices)
        {
            int newIdx = idx + direction;
            if (newIdx >= 0 && newIdx < _lstFiles.Items.Count)
                _lstFiles.SetSelected(newIdx, true);
        }
        _lstFiles.EndUpdate();

        UpdateButtonStates();
    }

    // ── Drag-and-drop ─────────────────────────────────────────────────────────

    private void LstFiles_DragEnter(object? sender, DragEventArgs e)
    {
        e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    private void LstFiles_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] droppedPaths)
            AddFilesToList(droppedPaths);
    }

    private void AddFilesToList(IEnumerable<string> paths)
    {
        _lstFiles.BeginUpdate();
        foreach (string path in paths)
        {
            string ext = Path.GetExtension(path);
            if (SupportedExtensions.Contains(ext))
            {
                _lstFiles.Items.Add(path);
                AppLogger.Log($"Added to list: {path}");
            }
            else
            {
                MessageBox.Show(
                    $"Unsupported file type: {ext}\n\nFile skipped:\n{path}",
                    "Unsupported File",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        _lstFiles.EndUpdate();
        UpdateButtonStates();
    }

    // ── Merge button ──────────────────────────────────────────────────────────

    private async void BtnMerge_Click(object? sender, EventArgs e)
    {
        if (_lstFiles.Items.Count == 0)
        {
            MessageBox.Show(
                "Please add at least one file before merging.",
                "No Files",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        // Prompt for output path.
        string defaultDir = AppSettings.GetDefaultOutputFolder();
        if (!Directory.Exists(defaultDir))
            defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        using var sfd = new SaveFileDialog
        {
            Title          = "Save Merged PDF As",
            Filter         = "PDF Files (*.pdf)|*.pdf",
            DefaultExt     = "pdf",
            FileName       = "merged.pdf",
            InitialDirectory = defaultDir
        };

        if (sfd.ShowDialog(this) != DialogResult.OK) return;

        string outputPath = sfd.FileName;
        var files = _lstFiles.Items.Cast<string>().ToList();

        SetUIEnabled(false);
        _progressBar.Value = 0;

        // Progress is reported back on the UI thread via Progress<T>.
        var progress = new Progress<(int Value, string Message)>(report =>
        {
            _progressBar.Value = Math.Clamp(report.Value, 0, _progressBar.Maximum);
            _lblStatus.Text    = report.Message;
        });

        try
        {
            await MergeAsync(files, outputPath, progress);
            MessageBox.Show(
                $"Merged successfully!\n\nOutput file:\n{outputPath}",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (FileLoadException fex)
        {
            AppLogger.LogException(nameof(BtnMerge_Click), fex);
            string detail = fex.Message.Contains("office", StringComparison.OrdinalIgnoreCase)
                ? "The Office assembly 'office' could not be loaded.\n\n" +
                  "This typically means Microsoft Office (Word, Excel, PowerPoint) is not properly installed.\n\n" +
                  "Please ensure Office is installed on this machine, then try again.\n\n" +
                  "Check logs/app.log for details."
                : $"Failed to load a required assembly:\n\n{fex.Message}\n\nCheck logs/app.log for details.";

            MessageBox.Show(detail, "Assembly Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _lblStatus.Text    = "Error — assembly not found. See logs/app.log.";
            _progressBar.Value = 0;
        }
        catch (Exception ex)
        {
            AppLogger.LogException(nameof(BtnMerge_Click), ex);
            MessageBox.Show(
                $"Merge failed:\n\n{ex.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            _lblStatus.Text    = "Error — see logs/app.log for details.";
            _progressBar.Value = 0;
        }
        finally
        {
            SetUIEnabled(true);
        }
    }

    private void SetUIEnabled(bool enabled)
    {
        _btnAdd.Enabled  = enabled;
        _lstFiles.Enabled = enabled;
        if (enabled) UpdateButtonStates();
        else
        {
            _btnRemove.Enabled   = false;
            _btnMoveUp.Enabled   = false;
            _btnMoveDown.Enabled = false;
            _btnMerge.Enabled    = false;
        }
    }

    // ── Core merge pipeline ───────────────────────────────────────────────────

    private static async Task MergeAsync(
        IReadOnlyList<string> files,
        string outputPath,
        IProgress<(int Value, string Message)> progress)
    {
        // Create a per-session temp directory to hold converted PDFs.
        string sessionTempDir = Path.Combine(
            AppSettings.GetTempFolder(), $"session_{Guid.NewGuid():N}");
        Directory.CreateDirectory(sessionTempDir);

        var tempPdfsToDelete = new List<string>();   // converted files to clean up
        var allPdfPaths      = new List<string>();   // ordered list passed to qpdf

        try
        {
            // ── Step 1: Convert every file to PDF (~80 % of progress) ──────────
            // Run a single STA conversion pass and reuse Office app instances.
            await RunOnStaAsync(() => ConvertAllFilesToPdfOnSta(
                files,
                sessionTempDir,
                allPdfPaths,
                tempPdfsToDelete,
                progress));

            // ── Step 2: Merge with qpdf (~20 % of progress, 80–100) ───────────
            progress.Report((82, "Merging PDFs..."));
            await MergeWithQpdfAsync(allPdfPaths, outputPath);

            progress.Report((100, "Done!"));
            AppLogger.Log($"Merge complete: {outputPath}");
        }
        finally
        {
            // Clean up temp PDFs regardless of success or failure.
            foreach (string f in tempPdfsToDelete)
                try { File.Delete(f); } catch { }

            try { Directory.Delete(sessionTempDir, recursive: false); } catch { }
        }
    }

    private static void ConvertAllFilesToPdfOnSta(
        IReadOnlyList<string> files,
        string sessionTempDir,
        List<string> allPdfPaths,
        List<string> tempPdfsToDelete,
        IProgress<(int Value, string Message)> progress)
    {
        object? wordApp = null;
        object? excelApp = null;
        object? pptApp = null;

        try
        {
            for (int i = 0; i < files.Count; i++)
            {
                string file = files[i];
                string fileName = Path.GetFileName(file);
                int pct = (int)((double)i / files.Count * 80.0);

                progress.Report((pct, $"Converting {fileName}..."));
                AppLogger.Log($"Processing [{i + 1}/{files.Count}]: {file}");

                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".pdf")
                {
                    allPdfPaths.Add(file);
                    continue;
                }

                string tempPdf = Path.Combine(sessionTempDir, $"{i:D4}_{Guid.NewGuid():N}.pdf");

                switch (ext)
                {
                    case ".doc":
                    case ".docx":
                        wordApp ??= CreateWordApp();
                        try
                        {
                            ConvertWordToPdfWithApp(wordApp, file, tempPdf);
                        }
                        catch (Exception firstEx) when (AppSettings.Current.RetryWordWithFreshInstance)
                        {
                            AppLogger.Log($"Word conversion failed on first attempt; recycling Word instance. File: {file}");
                            AppLogger.LogException("ConvertWordToPdfWithApp(first-attempt)", firstEx);
                            SafeQuitComApp(ref wordApp);
                            wordApp = CreateWordApp();
                            ConvertWordToPdfWithApp(wordApp, file, tempPdf);
                        }
                        break;
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                    case ".bmp":
                    case ".tif":
                    case ".tiff":
                        ConvertImageToPdfDirect(file, tempPdf);
                        break;
                    case ".xls":
                    case ".xlsx":
                        excelApp ??= CreateExcelApp();
                        ConvertExcelToPdfWithApp(excelApp, file, tempPdf);
                        break;
                    case ".ppt":
                    case ".pptx":
                        pptApp ??= CreatePowerPointApp();
                        ConvertPowerPointToPdfWithApp(pptApp, file, tempPdf);
                        break;
                    default:
                        throw new NotSupportedException($"Unsupported file type: {ext}");
                }

                allPdfPaths.Add(tempPdf);
                tempPdfsToDelete.Add(tempPdf);
                AppLogger.Log($"  → Temp PDF: {tempPdf}");
            }
        }
        finally
        {
            SafeQuitComApp(ref wordApp);
            SafeQuitComApp(ref excelApp);
            SafeQuitComApp(ref pptApp);
        }
    }

    private static object CreateWordApp()
    {
        var wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new Exception("Word.Application not registered");
        object app = Activator.CreateInstance(wordType)
            ?? throw new Exception("Failed to create Word.Application instance");
        dynamic appDyn = app;
        appDyn.Visible = false;
        appDyn.DisplayAlerts = 0;
        bool fastMode = AppSettings.Current.EnableFastMode;
        try { appDyn.ScreenUpdating = !fastMode ? true : false; } catch { }
        try { appDyn.DisplayStatusBar = !fastMode ? true : false; } catch { }
        try { appDyn.Options.PrintBackground = false; } catch { }
        try { appDyn.Options.CheckGrammarAsYouType = false; } catch { }
        try { appDyn.Options.CheckSpellingAsYouType = false; } catch { }
        TryForceWordActivePrinter(appDyn);
        return app;
    }

    private static void TryForceWordActivePrinter(dynamic wordApp)
    {
        if (!AppSettings.Current.ForceLocalPrinterForWord)
            return;

        try
        {
            string preferred = AppSettings.Current.PreferredWordPrinter;
            if (!string.IsNullOrWhiteSpace(preferred) &&
                IsInstalledPrinter(preferred))
            {
                wordApp.ActivePrinter = preferred;
                AppLogger.Log($"Word ActivePrinter forced to preferred printer: {preferred}");
                return;
            }

            string[] fallbackCandidates =
            {
                "Microsoft Print to PDF",
                "Microsoft XPS Document Writer",
                "Fax"
            };

            foreach (string candidate in fallbackCandidates)
            {
                if (!IsInstalledPrinter(candidate)) continue;
                wordApp.ActivePrinter = candidate;
                AppLogger.Log($"Word ActivePrinter forced to fallback printer: {candidate}");
                return;
            }

            AppLogger.Log("No local fallback printer found for Word ActivePrinter override.");
        }
        catch (Exception ex)
        {
            AppLogger.LogException("TryForceWordActivePrinter", ex);
        }
    }

    private static bool IsInstalledPrinter(string printerName)
    {
        foreach (string p in PrinterSettings.InstalledPrinters)
        {
            if (string.Equals(p, printerName, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static object CreateExcelApp()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new Exception("Excel.Application not registered");
        object app = Activator.CreateInstance(excelType)
            ?? throw new Exception("Failed to create Excel.Application instance");
        dynamic appDyn = app;
        appDyn.Visible = false;
        appDyn.DisplayAlerts = false;
        appDyn.AskToUpdateLinks = false;
        return app;
    }

    private static object CreatePowerPointApp()
    {
        var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
            ?? throw new Exception("PowerPoint.Application not registered");
        return Activator.CreateInstance(pptType)
            ?? throw new Exception("Failed to create PowerPoint.Application instance");
    }

    private static void SafeQuitComApp(ref object? app)
    {
        if (app == null) return;
        try { ((dynamic)app).Quit(); } catch { }
        try { Marshal.ReleaseComObject(app); } catch { }
        app = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>Check if a file has an Office extension.</summary>
    private static bool IsOfficeFile(string ext) =>
        new[] { ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx" }
            .Contains(ext, StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Pre-flight check: try to create Office applications via reflection to ensure they're available.
    /// If this fails, throw a descriptive exception that tells the user Office is not installed.
    /// </summary>
    private static void CheckOfficeAvailable()
    {
        AppLogger.Log("Pre-flight: checking Office COM availability via reflection...");

        string? failedApp = null;
        try
        {
            failedApp = "Word";
            var wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null) throw new Exception("Word.Application ProgID not registered");
            var word = Activator.CreateInstance(wordType);
            try { ((dynamic?)word)?.Quit(); } catch { }
            if (word != null) Marshal.ReleaseComObject(word);

            failedApp = "Excel";
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) throw new Exception("Excel.Application ProgID not registered");
            var excel = Activator.CreateInstance(excelType);
            try { ((dynamic?)excel)?.Quit(); } catch { }
            if (excel != null) Marshal.ReleaseComObject(excel);

            failedApp = "PowerPoint";
            var pptType = Type.GetTypeFromProgID("PowerPoint.Application");
            if (pptType == null) throw new Exception("PowerPoint.Application ProgID not registered");
            var ppt = Activator.CreateInstance(pptType);
            try { ((dynamic?)ppt)?.Quit(); } catch { }
            if (ppt != null) Marshal.ReleaseComObject(ppt);

            AppLogger.Log("Pre-flight check: all Office apps available via reflection.");
        }
        catch (Exception ex)
        {
            AppLogger.Log($"Pre-flight check FAILED at {failedApp}: {ex.Message}");
            throw new Exception(
                $"Microsoft Office {failedApp} is not available or not properly installed.\n\n" +
                $"Error: {ex.Message}\n\n" +
                "Please ensure Microsoft Office (Word, Excel, PowerPoint) is installed on this machine.",
                ex);
        }
    }

    // ── Word → PDF (via reflection-based COM) ──────────────────────────────────

    private static void ConvertWordToPdfWithApp(object app, string inputPath, string outputPdfPath)
    {
        AppLogger.Log($"Word → PDF: {inputPath}");
        object? doc = null;

        try
        {
            dynamic appDyn = app;

            dynamic docsDyn = appDyn.Documents;
            doc = docsDyn.Open(
                FileName: inputPath,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                OpenAndRepair: false,
                NoEncodingDialog: true);
            dynamic docDyn = doc!;

            // ExportAsFixedFormat(OutputFileName, ExportFormat)
            // wdExportFormatPDF = 17
            docDyn.ExportAsFixedFormat(outputPdfPath, 17);
        }
        finally
        {
            try { if (doc != null) { ((dynamic)doc).Close(SaveChanges: false); Marshal.ReleaseComObject(doc); } } catch { }
        }
    }

    // ── Excel → PDF (via reflection-based COM) ────────────────────────────────

    private static void ConvertExcelToPdfWithApp(object app, string inputPath, string outputPdfPath)
    {
        AppLogger.Log($"Excel → PDF: {inputPath}");
        object? workbook = null;

        try
        {
            dynamic appDyn = app;

            dynamic wbsDyn = appDyn.Workbooks;
            workbook = wbsDyn.Open(inputPath, ReadOnly: true, UpdateLinks: false);
            dynamic wbDyn = workbook!;

            // ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, OpenAfterPublish)
            // xlTypePDF = 0
            // xlQualityStandard = 0
            wbDyn.ExportAsFixedFormat(
                Type: 0, Filename: outputPdfPath, Quality: 0,
                IncludeDocProperties: true, IgnorePrintAreas: false, OpenAfterPublish: false);
        }
        finally
        {
            try { if (workbook != null) { ((dynamic)workbook).Close(SaveChanges: false); Marshal.ReleaseComObject(workbook); } } catch { }
        }
    }

    // ── PowerPoint → PDF (via reflection-based COM) ────────────────────────────

    private static void ConvertPowerPointToPdfWithApp(object app, string inputPath, string outputPdfPath)
    {
        AppLogger.Log($"PowerPoint → PDF: {inputPath}");
        object? presObj = null;

        try
        {
            dynamic appDyn = app;

            dynamic presentationsDyn = appDyn.Presentations;
            // Open(Filename, ReadOnly, Untitled, WithWindow)
            // ReadOnly: -1 (true), Untitled: 0 (false), WithWindow: 0 (false)
            presObj = presentationsDyn.Open(inputPath, -1, 0, 0);
            dynamic presDyn = presObj!;

            // SaveAs(FileName, FileFormat)
            // ppSaveAsPDF = 32
            presDyn.SaveAs(FileName: outputPdfPath, FileFormat: 32);
        }
        finally
        {
            try { if (presObj != null) { ((dynamic)presObj).Close(); Marshal.ReleaseComObject(presObj); } } catch { }
        }
    }

    // ── Image → PDF (direct, no Word COM) ────────────────────────────────────

    private static void ConvertImageToPdfDirect(string inputPath, string outputPdfPath)
    {
        AppLogger.Log($"Image → PDF (direct): {inputPath}");

        using var document = new PdfDocument();
        using var image = XImage.FromFile(inputPath);

        var page = document.AddPage();

        // Keep a predictable page size (A4 portrait) and fit image while preserving aspect ratio.
        page.Size = PdfSharp.PageSize.A4;

        using var gfx = XGraphics.FromPdfPage(page);

        double pageWidth = page.Width.Point;
        double pageHeight = page.Height.Point;
        double margin = 20.0;
        double maxWidth = pageWidth - margin * 2;
        double maxHeight = pageHeight - margin * 2;

        double imgWidth = image.PointWidth;
        double imgHeight = image.PointHeight;
        double scale = Math.Min(maxWidth / imgWidth, maxHeight / imgHeight);

        double drawWidth = imgWidth * scale;
        double drawHeight = imgHeight * scale;
        double x = (pageWidth - drawWidth) / 2.0;
        double y = (pageHeight - drawHeight) / 2.0;

        gfx.DrawImage(image, x, y, drawWidth, drawHeight);
        document.Save(outputPdfPath);
    }

    // ── qpdf merge ────────────────────────────────────────────────────────────

    private static async Task MergeWithQpdfAsync(
        IReadOnlyList<string> pdfFiles,
        string outputPath)
    {
        string qpdfExe = EmbeddedTools.ExtractQpdf();

        // Build: qpdf --empty --pages "f1.pdf" "f2.pdf" ... -- "out.pdf"
        var sb = new StringBuilder("--empty --pages ");
        foreach (string pdf in pdfFiles)
            sb.Append('"').Append(pdf).Append("\" ");
        sb.Append("-- \"").Append(outputPath).Append('"');

        string args = sb.ToString();
        AppLogger.Log($"qpdf command: \"{qpdfExe}\" {args}");

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName               = qpdfExe,
                Arguments              = args,
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true
            }
        };

        process.Start();

        // Read stdout and stderr concurrently to avoid deadlocks.
        Task<string> stdoutTask = process.StandardOutput.ReadToEndAsync();
        Task<string> stderrTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();

        string stdout = await stdoutTask;
        string stderr = await stderrTask;

        if (!string.IsNullOrWhiteSpace(stdout))
            AppLogger.Log($"qpdf stdout: {stdout.Trim()}");
        if (!string.IsNullOrWhiteSpace(stderr))
            AppLogger.Log($"qpdf stderr: {stderr.Trim()}");

        if (process.ExitCode != 0)
        {
            string detail = string.IsNullOrWhiteSpace(stderr) ? "(no stderr)" : stderr.Trim();
            throw new Exception(
                $"qpdf exited with code {process.ExitCode}.\n\n{detail}");
        }
    }

    // ── STA thread helper for COM calls ───────────────────────────────────────

    /// <summary>
    /// Runs <paramref name="action"/> on a newly created STA thread and
    /// returns a Task that completes (or faults) when it finishes.
    /// Required because Office COM objects must be created on STA threads,
    /// while Task.Run uses MTA thread-pool threads.
    /// </summary>
    private static Task RunOnStaAsync(Action action)
    {
        var tcs    = new TaskCompletionSource<bool>(
                         TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                action();
                tcs.SetResult(true);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        return tcs.Task;
    }
}
