# PdfMergerGui

PdfMergerGui is a WinForms application that merges mixed input files into one PDF.

- Office files are converted through Microsoft Office COM automation.
- Images are converted directly to PDF (no Word COM for images).
- PDFs are merged using embedded `qpdf.exe`.

## Prerequisites

| Requirement | Notes |
|---|---|
| .NET 8 SDK | Needed for local build/publish: <https://dotnet.microsoft.com/download> |
| Microsoft Office (Word / Excel / PowerPoint) | Required at runtime for Office document conversion |
| qpdf binaries | `qpdf.exe` and its companion DLLs must be present in the repo (see below) |

## qpdf files (exe + DLLs)

1. Download the latest Windows qpdf package from <https://github.com/qpdf/qpdf/releases>.
2. Copy `qpdf.exe` to the project root.
3. Copy qpdf companion DLLs to `qpdf\` (or root).

The project embeds `qpdf.exe` and matching DLLs as resources, then extracts them to:

`%TEMP%\PdfMergerGui`

at runtime via `EmbeddedTools.cs`.

## Build

```bash
dotnet restore
dotnet build
dotnet build -c Release
```

If build fails with file-lock errors (`MSB3021`/`MSB3027`), close/stop `PdfMergerGui.exe` and build again.

## Run in development

```bash
dotnet run
```

## Publish local single-file EXE

### Publish profile

```bash
dotnet publish -p:PublishProfile=FolderProfile
```

### Command line

```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:EnableCompressionInSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true
```

Output directory:

`bin\Release\net8.0-windows\win-x64\publish`

## GitHub binary releases (tag-based)

This repo includes `.github/workflows/release.yml`.

Pushing a version tag (`v*`) triggers remote build + GitHub Release asset creation (`PdfMergerGui-win-x64.zip`).

```bash
git tag v1.0.0
git push origin v1.0.0
```

Release page:

`https://github.com/amigatomte/pdfmergergui/releases`

## Configuration (`settings.json`)

```json
{
  "TempFolder": "%TEMP%\\PdfMergerGui",
  "DefaultOutputFolder": "%USERPROFILE%\\Documents",
  "EnableLogging": true,
  "EnableFastMode": true,
  "RetryWordWithFreshInstance": true,
  "ForceLocalPrinterForWord": true,
  "PreferredWordPrinter": "Microsoft Print to PDF"
}
```

| Key | Description |
|---|---|
| `TempFolder` | Temp directory for extracted qpdf and intermediate PDFs |
| `DefaultOutputFolder` | Initial Save dialog folder |
| `EnableLogging` | Enables `logs\app.log` |
| `EnableFastMode` | Applies conversion speed optimizations |
| `RetryWordWithFreshInstance` | Retries failed Word conversion with a fresh Word instance |
| `ForceLocalPrinterForWord` | Forces Word to use a local printer to avoid printer connection delays |
| `PreferredWordPrinter` | Preferred printer name for Word (fallbacks are applied if missing) |

## Supported input formats

| Category | Extensions |
|---|---|
| PDF | `.pdf` |
| Word | `.doc`, `.docx` |
| Excel | `.xls`, `.xlsx` |
| PowerPoint | `.ppt`, `.pptx` |
| Images | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tif`, `.tiff` |

## Architecture notes

- Office COM conversion runs on STA and reuses app instances in batch for performance.
- Image conversion is direct (`PdfSharp`) and does not invoke Word.
- Word conversion includes printer forcing + retry/recycle behavior to avoid stalls.
- qpdf merge is executed via subprocess with async stdout/stderr reads.
