param(
    [Parameter(HelpMessage = "Path to virtual environment folder (e.g. D:\envs\llmenv) or to python.exe")]
    [string]$VenvPath = "",

    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $ProjectRoot

function Resolve-PythonExecutable {
    param([string]$PathInput)

    if ([string]::IsNullOrWhiteSpace($PathInput)) {
        if ($env:VIRTUAL_ENV) {
            $PathInput = $env:VIRTUAL_ENV
        }
        elseif (Test-Path (Join-Path $ProjectRoot ".venv")) {
            $PathInput = Join-Path $ProjectRoot ".venv"
        }
        elseif (Test-Path (Join-Path $ProjectRoot "..\llmenv")) {
            $PathInput = (Resolve-Path (Join-Path $ProjectRoot "..\llmenv")).Path
        }
    }

    if ([string]::IsNullOrWhiteSpace($PathInput)) {
        $fallback = Get-Command python -ErrorAction SilentlyContinue
        if ($fallback) {
            Write-Warning "No -VenvPath given; using python on PATH: $($fallback.Source)"
            return $fallback.Source
        }
        throw @"
No Python environment found.

Pass a virtual environment path, for example:
  .\scripts\build_windows.ps1 -VenvPath 'D:\code\language_model\tokenization\llmenv'

Or activate a venv first, then run without -VenvPath (uses `$env:VIRTUAL_ENV`).
"@
    }

    $PathInput = $PathInput.Trim().Trim('"')

    if ($PathInput -match '\\python\.exe$|/python\.exe$') {
        $candidate = $PathInput
    }
    else {
        $candidate = Join-Path $PathInput "Scripts\python.exe"
    }

    if (-not (Test-Path $candidate)) {
        throw "Python not found at: $candidate`nCheck -VenvPath (venv folder or full path to python.exe)."
    }

    return (Resolve-Path $candidate).Path
}

$PythonExe = Resolve-PythonExecutable -PathInput $VenvPath
Write-Host "Using Python: $PythonExe"
Write-Host "Project root: $ProjectRoot"
Write-Host ""

if ($Clean) {
    if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
    if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }
    Write-Host "Cleaned build/ and dist/"
}

& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install -r requirements.txt
& $PythonExe -m pip install pyinstaller

& $PythonExe -m PyInstaller `
  --noconfirm `
  --windowed `
  --name "CorpusConverter" `
  --add-data "app/assets;app/assets" `
  main.py

Write-Host ""
Write-Host "Build complete."
Write-Host "Executable: $ProjectRoot\dist\CorpusConverter\CorpusConverter.exe"
