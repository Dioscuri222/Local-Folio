# Local-Folio

**Local-Folio** is a 100% local, privacy-first web suite for **document-to-Markdown conversion**, **offline batch OCR (including rotated photos & scanned PDFs)**, and a **29-tool iLovePDF-style PDF Studio** — with zero file-count limits and zero cloud uploads.

## Key Capabilities

Local-Folio runs completely on your own machine (`http://127.0.0.1:8765`) with a built-in **Global Feature Search (`Ctrl+K` or magnifying glass icon)** and **Live Per-Page Progress Monitor with Cancel**:

1. **MarkItDown Converter** — Unlimited batch conversion for PDF, DOCX, PPTX, XLSX/CSV, HTML, EPUB, ZIP, Audio, and Web/YouTube URLs, with automatic local RapidOCR fallback for scanned PDFs.
2. **Local Batch Image & PDF OCR** — Offline batch text extraction from `.jpg`, `.jpeg`, `.png`, `.webp`, `.bmp`, and scanned `.pdf` using RapidOCR (ONNX) with smart auto-orientation (`0° / 90° / 180° / 270°`), `Selected` / `Combined All` view, and `.ZIP` export.
3. **Local PDF Toolkit (29 Tools)**:
   - **Organize PDF**: Merge PDF, Split PDF, Remove Pages, Organize PDF (visual thumbnails + reorder), Rotate PDF, Crop PDF.
   - **Optimize PDF**: Compress PDF, Repair PDF, OCR PDF (Smart Hybrid + Force RapidOCR), PDF to PDF/A.
   - **Convert PDF**: PDF to Word (`.docx`), PDF to PowerPoint (`.pptx`), PDF to Excel (`.xlsx`), PDF to JPG/PNG, PDF to Markdown, Word to PDF, PowerPoint to PDF, Excel to PDF, JPG/PNG to PDF, HTML/URL to PDF.
   - **Edit & Review**: Watermark PDF, Page Numbers, Sign PDF (interactive signature pad), Compare PDF (side-by-side HTML diff).
   - **PDF Security**: Protect PDF, Unlock PDF, Redact PDF (permanent text stream removal).
   - **AI & Workflows**: AI Summarizer (uses your Settings LLM: Ollama, LM Studio, or OpenAI), Custom Workflow Builder (chain Merge, Rotate, Watermark, Page Numbers, Compress, Protect in one click).

## Installation

Make sure you have **Python 3.10+** installed, then clone this repository and install dependencies from PyPI:

```bash
git clone https://github.com/Dioscuri222/Local-Folio.git
cd Local-Folio

# (Optional) Create and activate a virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
# source .venv/bin/activate

pip install -r requirements.txt
```

> **Note for Office-to-PDF & HTML-to-PDF:** Word/Excel/PowerPoint-to-PDF uses native Microsoft Office COM on Windows (or LibreOffice `soffice` on macOS/Linux). HTML-to-PDF uses local headless Microsoft Edge or Google Chrome.

## Running the App

### Windows (One-Click)
Double-click **`local-folio.bat`** — this starts the local FastAPI server and opens `http://127.0.0.1:8765` in your browser.

### Command Line (Windows / macOS / Linux)
```bash
python local_folio_server.py
```
Then open **[http://127.0.0.1:8765](http://127.0.0.1:8765)** in your browser.

## Project Structure

- `local_folio_server.py` — Core FastAPI backend (MarkItDown, RapidOCR, core PDF tools, live job progress & cancel APIs).
- `pdf_extended_routes.py` — Extended PDF & Office conversion router (Watermark, Page Numbers, Crop, Redact, Repair, Organize, Sign, Compare, PDF ↔ Word/Excel/PowerPoint, HTML → PDF, PDF/A, AI Summarizer, Custom Workflow).
- `local-folio.html` — Single-file frontend UI with Expanding Label tab navigation, Global Feature Search (`Ctrl+K`), full-tab drag-and-drop, and live progress monitor.
- `local-folio.bat` — Windows launcher script.
- `requirements.txt` — Python dependencies installed from PyPI.

## Credits & Open-Source Attribution

This standalone project is built on top of open-source libraries installed via PyPI:

- **[Microsoft MarkItDown](https://github.com/microsoft/markitdown)** (MIT License) — Document-to-Markdown conversion engine (`markitdown[all]`).
- **[RapidOCR](https://github.com/RapidAI/RapidOCR)** (Apache-2.0 License) — Fast CPU ONNX runtime port of **Baidu PaddleOCR**.
- **[PyMuPDF](https://pymupdf.readthedocs.io/)**, **[pikepdf](https://github.com/pikepdf/pikepdf)**, **[pypdf](https://github.com/py-pdf/pypdf)**, **[pypdfium2](https://github.com/pypdfium2-team/pypdfium2)**, **[pdf2docx](https://github.com/ArtifexSoftware/pdf2docx)**, **[pdfplumber](https://github.com/jsvine/pdfplumber)**, **[openpyxl](https://openpyxl.readthedocs.io/)**, and **[python-pptx](https://python-pptx.readthedocs.io/)**.
- **[FastAPI](https://fastapi.tiangolo.com/)**, **[Uvicorn](https://www.uvicorn.org/)**, and **[Pillow](https://python-pillow.org/)**.
- Client-side UI utilities: **PDF.js**, **Tesseract.js** (browser fallback), **Marked**, **JSZip**, and **Lucide Icons**.
