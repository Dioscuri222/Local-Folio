# Folio-Local

**Folio-Local** is a 100% local, privacy-first web suite for **document-to-Markdown conversion**, **offline OCR (including rotated photos & scanned PDFs)**, and **iLovePDF-style PDF tools** — with zero file-count limits and zero cloud uploads.

## Problem & Solution

- **Problem:** Most online PDF and OCR tools cap batch uploads (often 2–3 files), require uploading sensitive documents to third-party servers, and fail on scanned PDFs or rotated phone photos of screens and documents.
- **Solution:** Folio-Local runs completely on your own machine (`http://127.0.0.1:8765`) by combining three local workflows in one interface:
  1. **MarkItDown Converter** — Unlimited batch conversion for PDF, DOCX, PPTX, XLSX/CSV, HTML, EPUB, ZIP, Audio, and Web/YouTube URLs, with automatic local OCR fallback for scanned PDFs.
  2. **Local Image & PDF OCR** — Offline text extraction from `.jpg`, `.jpeg`, `.png`, `.webp`, `.bmp`, and scanned `.pdf` using RapidOCR (ONNX) with smart auto-orientation (`0° / 90° / 180° / 270°`).
  3. **Local PDF Toolkit** — Merge PDF, Split / Extract Pages, Rotate PDF, Compress PDF, PDF to JPG/PNG (`.zip`), and JPG/PNG to PDF.

## Installation

Make sure you have **Python 3.10+** installed, then clone this repository and install dependencies from PyPI:

```bash
git clone https://github.com/Dioscuri222/Folio-Local.git
cd Folio-Local

# (Optional) Create and activate a virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
# source .venv/bin/activate

pip install -r requirements.txt
```

## Running the App

### Windows (One-Click)
Double-click **`Start_Markitdown.bat`** — this starts the local FastAPI server and opens `http://127.0.0.1:8765` in your browser.

### Command Line (Windows / macOS / Linux)
```bash
python markitdown_server.py
```
Then open **[http://127.0.0.1:8765](http://127.0.0.1:8765)** in your browser.

## Project Structure

- `markitdown_server.py` — Local FastAPI backend (Markdown conversion, RapidOCR, and PDF toolkit APIs).
- `Markitdown_Lokal.html` — Single-file frontend UI with Expanding Label tab navigation.
- `Start_Markitdown.bat` — Windows launcher script.
- `requirements.txt` — Python dependencies installed from PyPI.

## Credits & Open-Source Attribution

This standalone project is built on top of open-source libraries installed via PyPI:

- **[Microsoft MarkItDown](https://github.com/microsoft/markitdown)** (MIT License) — Document-to-Markdown conversion engine (`markitdown[all]`).
- **[RapidOCR](https://github.com/RapidAI/RapidOCR)** (Apache-2.0 License) — Fast CPU ONNX runtime port of **Baidu PaddleOCR**.
- **[pypdf](https://github.com/py-pdf/pypdf)** (BSD-3-Clause) & **[pypdfium2](https://github.com/pypdfium2-team/pypdfium2)** (Apache-2.0 / BSD-3-Clause) — Local PDF manipulation and page rendering.
- **[FastAPI](https://fastapi.tiangolo.com/)**, **[Uvicorn](https://www.uvicorn.org/)**, and **[Pillow](https://python-pillow.org/)**.
- Client-side UI utilities: **PDF.js**, **Tesseract.js** (browser fallback), **Marked**, **JSZip**, and **Lucide Icons**.
