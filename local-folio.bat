@echo off
title Local-Folio - Local PDF, RapidOCR ^& MarkItDown Studio
cd /d "%~dp0"

python -c "import fastapi, uvicorn, rapidocr_onnxruntime, pymupdf" >nul 2>&1
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
