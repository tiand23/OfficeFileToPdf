# OfficeFileToPdf

[中文](README.md) | [日本語](README.ja.md) | [English](README.en.md)

## Overview

This is a Windows-only batch conversion tool.

It recursively scans a target folder and exports supported Office documents and PDFs into a `res/` folder while preserving the original folder structure as much as possible.

All generated output files are PDFs.

The tool also creates `source_to_pdf_map.json` in the root of `res/` to record the mapping between source files and generated PDFs.

## Main Features

- Recursive folder scanning
- Output to `res/` while keeping the original directory structure
- Convert Excel, Word, and PowerPoint files to PDF
- Re-process existing PDF files and write them into `res/`
- Automatically trim PDF white margins
- Generate a JSON mapping file for source-to-output relationships
- Supports double-click execution through `ToPdf.bat`
- Supports drag-and-drop folders onto `ToPdf.bat`

## Requirements

- The operating system must be Windows
- Microsoft Office must be installed
- Office document conversion uses Microsoft Office COM only
- LibreOffice is not used
- On first run, `ToPdf.bat` will try to install Python if Python is missing
- On first run, the script creates `.venv` and installs Python dependencies

## Supported Input Types

Files that can be converted to PDF:

- Excel: `.xls` `.xlsx` `.xlsm` `.xlsb` `.csv` `.ods`
- Word: `.doc` `.docx` `.docm` `.rtf` `.odt`
- PowerPoint: `.ppt` `.pptx` `.pptm` `.odp`
- PDF: `.pdf`

Files that are skipped:

- Image files: `.png` `.jpg` `.jpeg` `.bmp` `.gif` `.svg` `.tif` `.tiff` `.webp` `.heic` `.ico`
- Text-like files: `.txt` `.md` `.json` `.yaml` `.yml` `.xml` `.toml` `.ini` `.cfg` `.log` `.rst`
- Scripts or executables: `.py` `.bat` `.cmd` `.ps1` `.exe` `.dll` `.sh`
- Files without an extension
- Other file types that are not currently supported

## Output Rules

- The default output folder is `res/` under the target folder
- The directory structure is preserved as much as possible
- All output files use the `.pdf` extension
- If a target filename collision happens, the tool will rename the output automatically to avoid overwriting

Example:

Source folder:

```text
input/
  report.xlsx
  deck.pptx
  docs/
    plan.docx
    old.pdf
    note.txt
```

Output folder:

```text
input/
  res/
    report.pdf
    deck.pdf
    docs/
      plan.pdf
      old.pdf
    source_to_pdf_map.json
```

Notes:

- `note.txt` is skipped
- `old.pdf` is processed again and written to `res/docs/old.pdf`
- `source_to_pdf_map.json` stores the source-to-output relationships

## How to Run

### Option 1: Double-click

Double-click `ToPdf.bat`.

By default, it processes the folder where `ToPdf.bat` is located and creates `res/` next to it.

### Option 2: Drag and drop a folder

Drag a target folder onto `ToPdf.bat`.

That folder becomes the processing target.

### Option 3: Run from command line

In Windows `cmd`:

```bat
cd /d D:\OfficeFileToPdf
ToPdf.bat "D:\YourInputFolder"
```

To process the script folder itself:

```bat
cd /d D:\OfficeFileToPdf
ToPdf.bat
```

## Mapping JSON

The tool writes `res/source_to_pdf_map.json`.

Example structure:

```json
{
  "source_root": "D:/input",
  "output_root": "D:/input/res",
  "mappings": [
    {
      "source_path": "D:/input/docs/plan.docx",
      "source_relative_path": "docs/plan.docx",
      "output_pdf_path": "D:/input/res/docs/plan.pdf",
      "output_relative_path": "docs/plan.pdf",
      "kind": "word"
    }
  ]
}
```

Field descriptions:

- `source_root`: source root folder
- `output_root`: output folder
- `mappings`: list of successfully generated mappings
- `source_path`: absolute path of the source file
- `source_relative_path`: relative path of the source file
- `output_pdf_path`: absolute path of the generated PDF
- `output_relative_path`: path relative to `res/`
- `kind`: one of `excel`, `word`, `powerpoint`, or `pdf`

## Log Files

Two log files are generated during execution:

- `ToPdf_run.log`
- `ToPdf_python.log`

Purpose:

- `ToPdf_run.log`: logs the batch launcher flow
- `ToPdf_python.log`: logs Python output and errors

If the process fails, check these two files first.

## Notes

- This tool is Windows-only
- Office document conversion depends on whether Microsoft Office can open the file correctly
- Even if a file extension is listed as supported, conversion may still fail if Microsoft Office cannot open that file
- If PDF trimming fails, the tool falls back to copying the PDF directly so the whole batch job does not stop
- Skipped files are not included in `mappings`
- Failed files are not currently written into `mappings`

## FAQ

### The window closes immediately after double-clicking

Use the latest `ToPdf.bat`.

The latest version pauses before exiting and prints the log file locations.

### It cannot start `Excel.Application`, `Word.Application`, or `PowerPoint.Application`

Common reasons:

- Microsoft Office is not installed
- Office COM registration is broken
- Office is not available in the current Windows environment

### Some files are missing in the output

Please check:

- Whether the file belongs to a skipped category
- Whether `ToPdf_python.log` contains a conversion failure
- Whether `source_to_pdf_map.json` contains a mapping record for that file

## Files

- `ToPdf.py`: main program
- `ToPdf.bat`: Windows launcher
- `requirements.txt`: Python dependencies
- `pdfToPng.py`: legacy PDF trimming script

## License

This project is released under the MIT License.

See [`LICENSE`](LICENSE) for details.
