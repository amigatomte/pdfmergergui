# PdfMergerGui

A WinForms application that merges PDF, Word, Excel, PowerPoint, and image files
into a single PDF.  Office documents and images are first converted to PDF via
Microsoft Office COM automation; the PDFs are then merged with the embedded
`qpdf.exe` tool.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| .NET 8 SDK | <https://dotnet.microsoft.com/download> |
| Microsoft Office (Word / Excel / PowerPoint) | Must be installed on the build **and** runtime machine; used for COM-based conversion |
| `qpdf.exe` | Must be added to the project root manually (see below) |

---

## Adding qpdf.exe

1. Download the latest **Windows static binary** from the [qpdf releases page](https://github.com/qpdf/qpdf/releases).
   Look for `qpdf-<version>-bin-msvc64.zip`.
2. Extract `bin\qpdf.exe` from the archive.
3. Copy `qpdf.exe` into the project root (next to `PdfMergerGui.csproj`).
4. The `.csproj` already declares it as an `EmbeddedResource`; no further action needed.

> **Note:** Some qpdf releases ship with companion DLLs.  If `qpdf.exe` requires
> DLLs, add them alongside `qpdf.exe` in the temp extraction directory
> (`%TEMP%\PdfMergerGui\`) or bundle them as additional embedded resources and
> extract them in `EmbeddedTools.cs`.

---

## Building

```bash
# Restore NuGet packages
dotnet restore

# Debug build
dotnet build

# Release build
dotnet build -c Release
```

---

## Running in development

```bash
dotnet run
```

---

## Publishing as a single-file self-contained EXE

### Option A — using the included publish profile

```bash
dotnet publish -p:PublishProfile=FolderProfile
```

Output: `bin\Release\net8.0-windows\publish\win-x64\PdfMergerGui.exe`

`settings.json` is placed next to the EXE automatically.

### Option B — command line (equivalent)

```bash
dotnet publish -c Release -r win-x64 ^
  --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:EnableCompressionInSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true
```

### What is bundled into the EXE

| Item | How bundled |
|---|---|
| .NET 8 runtime | Embedded (self-contained) |
| Application assemblies | Embedded (single-file) |
| `qpdf.exe` | Embedded resource — extracted to `%TEMP%\PdfMergerGui\` at first run |
| `settings.json` | **Not** embedded — copied next to the EXE so the user can edit it |
| `logs\app.log` | Created at runtime next to the EXE |

---

## Configuration — `settings.json`

```json
{
  "TempFolder": "%TEMP%\\PdfMergerGui",
  "DefaultOutputFolder": "%USERPROFILE%\\Documents",
  "EnableLogging": true
}
```

| Key | Default | Description |
|---|---|---|
| `TempFolder` | `%TEMP%\PdfMergerGui` | Where converted PDFs and qpdf.exe are stored temporarily |
| `DefaultOutputFolder` | `%USERPROFILE%\Documents` | Initial directory shown in the "Save As" dialog |
| `EnableLogging` | `true` | Write to `logs\app.log` next to the EXE |

Environment variables (e.g. `%TEMP%`) are expanded at runtime.

---

## Project structure

```
PdfMergerGui/
├── PdfMergerGui.csproj        Project file
├── Program.cs                 Entry point
├── MainForm.cs                UI + all conversion / merge logic
├── EmbeddedTools.cs           Extracts qpdf.exe from the embedded resource
├── Settings.cs                Loads / saves settings.json
├── Logger.cs                  Thread-safe file logger
├── settings.json              Default settings (copied to output dir)
├── qpdf.exe                   ← YOU MUST ADD THIS (see above)
└── Properties/
    └── PublishProfiles/
        └── FolderProfile.pubxml  Single-file publish profile
```

---

## Supported input formats

| Category | Extensions |
|---|---|
| PDF | `.pdf` |
| Word | `.doc`, `.docx` |
| Excel | `.xls`, `.xlsx` |
| PowerPoint | `.ppt`, `.pptx` |
| Images | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tif`, `.tiff` |

---

## Architecture notes

* **Office COM interop** requires Microsoft Office to be installed.  The conversion
  helpers (`ConvertWordToPdf`, `ConvertExcelToPdf`, etc.) always create fresh
  application instances, export to a temp PDF, and immediately call `Quit()` plus
  `Marshal.ReleaseComObject` followed by `GC.Collect` / `GC.WaitForPendingFinalizers`
  to ensure the COM objects are released.
* **STA requirement**: Office COM objects must be created on STA threads.  The
  `RunOnStaAsync` helper in `MainForm.cs` spawns a dedicated STA thread for every
  conversion task; this keeps the UI responsive while being compatible with the
  Office threading model.
* **qpdf** is invoked as a subprocess with `Process.WaitForExitAsync` and parallel
  stdout/stderr reads to prevent deadlocks.
