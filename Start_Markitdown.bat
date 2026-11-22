@echo off
title Local-Folio - Local PDF, RapidOCR & MarkItDown Studio
cd /d "%~dp0"
echo Starting Local-Folio at http://127.0.0.1:8765 ...
start "" http://127.0.0.1:8765
python markitdown_server.py
pause
