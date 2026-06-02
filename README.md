# Professional Document to Corpus TXT Converter

A production-style Python project that converts document files into raw corpus `.txt` files for downstream tokenization and language model workflows.

## Features

- Professional PyQt6 desktop UI
- App icon and menu bar actions (including Open Output Folder)
- Tesseract OCR executable path selection (`tesseract.exe`)
- Automatic OCR language discovery from installed Tesseract language packs
- Multi-language OCR selection (e.g., `eng`, `deu`, etc.)
- Configurable Office COM ProgIDs for `.doc` extraction (Word/WPS/custom)
- Single-file or full-folder conversion
- Extension filtering in UI
- OCR fallback for low-text PDF pages
- Structured corpus output with paragraph-level `[SEP]` separators
- Persistent settings (last paths, selected languages, options)
- Built-in **System Check** button for deployment diagnostics

## Supported Input Extensions

- `.pdf`
- `.docx`
- `.doc` (requires Microsoft Word + `pywin32` OR LibreOffice)
- `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif`, `.tiff`

## Project Structure

```text
doc2pdf/
  app/
    assets/
      app_icon.svg
    config.py
    services/
      converter.py
      extractors.py
      text_processing.py
    ui/
      main_window.py
      worker.py
  main.py
  requirements.txt
  scripts/
    build_windows.ps1
```

## Setup

1. Install Python 3.10+.
2. Install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) for Windows.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run Desktop UI (Recommended)

```bash
python main.py
```

Workflow:

1. Select `tesseract.exe`
2. Click **Load & Detect Languages**
3. Choose OCR languages
4. (Optional) Set Office COM ProgIDs for `.doc` compatibility
5. Choose single file or input folder
6. Choose output folder
7. Configure options and click **Start Conversion**

## Build Windows EXE (PyInstaller)

Specify the virtual environment used for the build (recommended):

```powershell
.\scripts\build_windows.ps1 -VenvPath "D:\code\language_model\tokenization\llmenv"
```

You can pass either the venv folder or `python.exe` directly:

```powershell
.\scripts\build_windows.ps1 -VenvPath "D:\code\language_model\tokenization\llmenv\Scripts\python.exe"
```

If `-VenvPath` is omitted, the script tries in order:

1. `$env:VIRTUAL_ENV` (if you already activated a venv)
2. `.venv` in the project folder
3. `..\llmenv` next to the project folder
4. `python` on PATH (with a warning)

Clean build (remove previous `build/` and `dist/` first):

```powershell
.\scripts\build_windows.ps1 -VenvPath "D:\path\to\your\venv" -Clean
```

Debug build with visible console output (recommended when EXE behaves differently):

```powershell
.\scripts\build_windows.ps1 -VenvPath "D:\path\to\your\venv" -Clean -Console
```

Generated executable:

- `dist/CorpusConverter/CorpusConverter.exe`
- Runtime log file: `%USERPROFILE%\CorpusConverter\logs\app.log`

## Notes

- `.doc` extraction uses Word COM first, then LibreOffice fallback.
- If OCR languages are missing, install additional Tesseract language data (`.traineddata` files).
- Output files are generated in UTF-8 encoding.

