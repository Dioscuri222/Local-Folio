@echo off
title Local-Folio - Local PDF, RapidOCR ^& MarkItDown Studio
cd /d "%~dp0"

python -c "import importlib.util as u, sys; sys.exit(0 if all(u.find_spec(m) for m in ('fastapi','uvicorn','pymupdf','pdf2docx')) and (u.find_spec('rapidocr_onnxruntime') or u.find_spec('rapidocr')) else 1)" >nul 2>&1
if errorlevel 1 (
    echo First-time setup: installing Local-Folio dependencies...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Failed to install dependencies. Make sure Python 3.10+ and pip are installed.
        pause
        exit /b 1
    )
)

echo Starting Local-Folio at http://127.0.0.1:8765 ...
start "" http://127.0.0.1:8765
python local_folio_server.py
pause
