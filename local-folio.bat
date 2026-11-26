@echo off
setlocal EnableDelayedExpansion
title Local-Folio - Local PDF, RapidOCR ^& MarkItDown Studio
cd /d "%~dp0"

set "PY_CMD="

:: 1. Prefer existing local portable Python if already bootstrapped
if exist "%~dp0python-portable\python.exe" (
    set "PY_CMD="%~dp0python-portable\python.exe""
    goto :check_deps
)

:: 2. Check system Python in PATH (ignores Windows Store 0-byte stub)
python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :check_deps
)

:: 3. Check Windows Python Launcher (py.exe)
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py -3"
    goto :check_deps
)

:: 4. No Python found -> Auto-download official Portable Python 3.12 (~11 MB)
echo ============================================================================
echo  Python 3.10+ was not found on this computer.
echo  Downloading official Portable Python 3.12.8 into .\python-portable\ ...
echo ============================================================================
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference = 'Stop'; " ^
  "$dir = Join-Path (Get-Location) 'python-portable'; " ^
  "New-Item -ItemType Directory -Force -Path $dir | Out-Null; " ^
  "$zip = Join-Path $dir 'python-embed.zip'; " ^
  "Write-Host '  [1/3] Downloading python-3.12.8-embed-amd64.zip from python.org...'; " ^
  "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip' -OutFile $zip; " ^
  "Write-Host '  [2/3] Extracting portable Python runtime...'; " ^
  "Expand-Archive -Path $zip -DestinationPath $dir -Force; " ^
  "Remove-Item $zip -Force; " ^
  "$pth = Join-Path $dir 'python312._pth'; " ^
  "@('python312.zip', '.', '..', 'Lib\site-packages', 'import site') | Set-Content -Path $pth -Encoding ASCII; " ^
  "Write-Host '  [3/3] Bootstrapping pip package manager...'; " ^
  "$getpip = Join-Path $dir 'get-pip.py'; " ^
  "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getpip; " ^
  "& (Join-Path $dir 'python.exe') $getpip --no-warn-script-location; " ^
  "Remove-Item $getpip -Force"

if not exist "%~dp0python-portable\python.exe" (
    echo.
    echo Failed to download Portable Python. Please check your internet connection.
    pause
    exit /b 1
)
set "PY_CMD="%~dp0python-portable\python.exe""

:check_deps
%PY_CMD% -c "import importlib.util as u, sys; sys.exit(0 if all(u.find_spec(m) for m in ('fastapi','uvicorn','pymupdf','pdf2docx')) and (u.find_spec('rapidocr_onnxruntime') or u.find_spec('rapidocr')) else 1)" >nul 2>&1
if errorlevel 1 (
    echo First-time setup: installing Local-Folio dependencies...
    %PY_CMD% -m pip install -r requirements.txt --no-warn-script-location
    if errorlevel 1 (
        echo.
        echo Failed to install dependencies. Please check your internet connection and try again.
        pause
        exit /b 1
    )
)

echo Starting Local-Folio at http://127.0.0.1:8765 ...
start "" http://127.0.0.1:8765
%PY_CMD% local_folio_server.py
pause
