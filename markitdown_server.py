import io
import os
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional, Tuple

import numpy as np
import pypdfium2 as pdfium
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from markitdown import MarkItDown
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter
from rapidocr_onnxruntime import RapidOCR

APP_DIR = Path(__file__).resolve().parent
HTML_PATH = APP_DIR / "Markitdown_Lokal.html"

app = FastAPI(title="Folio-Local — Local PDF, OCR & Markdown Suite")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Original-Size", "X-Compressed-Size", "X-Page-Count", "Content-Disposition"],
)

# Lazy singleton for RapidOCR (Baidu PaddleOCR ONNX runtime)
_ocr_engine: Optional[RapidOCR] = None


def get_ocr_engine() -> RapidOCR:
    global _ocr_engine
    if _ocr_engine is None:
        _ocr_engine = RapidOCR()
    return _ocr_engine


def build_frontmatter(filename: str, size: int, content_type: str, method: str) -> str:
    now = datetime.now(timezone.utc).isoformat()
    return (
        f'---\n'
        f'title: "{filename}"\n'
        f'converted_at: "{now}"\n'
        f'original_size: {size}\n'
        f'original_type: "{content_type or "unknown"}"\n'
        f'engine: "{method}"\n'
        f'---\n\n'
    )


def create_markitdown_instance(
    llm_api_key: Optional[str] = None,
    llm_base_url: Optional[str] = None,
    llm_model: Optional[str] = None,
) -> MarkItDown:
    if llm_model and (llm_api_key or llm_base_url):
        from openai import OpenAI

        client = OpenAI(
            api_key=llm_api_key or "ollama",
            base_url=llm_base_url or None,
        )
        return MarkItDown(llm_client=client, llm_model=llm_model)
    return MarkItDown()


def _score_ocr_result(result) -> Tuple[float, List[str], float]:
    """Return (weighted_score, lines, avg_confidence) favoring readable horizontal lines."""
    if not result:
        return 0.0, [], 0.0

    lines: List[str] = []
    confs: List[float] = []
    horiz_weight = 0.0

    for box, text, score in result:
        clean = (text or "").strip()
        if not clean:
            continue
        lines.append(clean)
        conf = float(score or 0.0)
        confs.append(conf)
        # Measure box width vs height to reward upright horizontal text
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        w = max(xs) - min(xs)
        h = max(ys) - min(ys)
        orientation_bonus = 1.25 if w >= h else 0.7
        horiz_weight += len(clean) * conf * orientation_bonus

    avg_conf = (sum(confs) / len(confs)) if confs else 0.0
    return horiz_weight, lines, avg_conf


def _downscale_for_ocr(img: Image.Image, max_side: int = 2200) -> Image.Image:
    w, h = img.size
    if max(w, h) <= max_side:
        return img
    scale = max_side / float(max(w, h))
    return img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)


def run_rapidocr_on_image(
    img: Image.Image,
    rotate_mode: str = "auto",
) -> Tuple[str, int, float, int]:
    """Run RapidOCR on a PIL Image with smart auto-orientation for rotated phone photos."""
    engine = get_ocr_engine()
    base_img = ImageOps.exif_transpose(img).convert("RGB")
    base_img = _downscale_for_ocr(base_img)

    if rotate_mode in {"0", "90", "180", "270"}:
        deg = int(rotate_mode)
        rotated = base_img.rotate(-deg, expand=True) if deg else base_img
        res, _ = engine(np.array(rotated))
        _, lines, avg_conf = _score_ocr_result(res)
        return "\n".join(lines), deg, avg_conf, len(lines)

    # Auto mode: test 0 deg first; if result looks rotated or weak, also test 90 & 270 deg
    candidates = [0]
    res0, _ = engine(np.array(base_img))
    best_score, best_lines, best_conf = _score_ocr_result(res0)
    best_deg = 0

    # Check if many boxes in 0-deg result are taller than wide (vertical text -> rotated photo)
    vertical_boxes = 0
    total_boxes = len(res0) if res0 else 0
    if res0:
        for box, _, _ in res0:
            xs = [pt[0] for pt in box]
            ys = [pt[1] for pt in box]
            if (max(ys) - min(ys)) > (max(xs) - min(xs)) * 1.15:
                vertical_boxes += 1

    needs_rotation_check = (
        total_boxes == 0
        or best_conf < 0.82
        or (total_boxes > 0 and vertical_boxes / total_boxes >= 0.35)
    )

    if needs_rotation_check:
        for deg in (90, 270, 180):
            rotated = base_img.rotate(-deg, expand=True)
            res, _ = engine(np.array(rotated))
            score, lines, conf = _score_ocr_result(res)
            if score > best_score * 1.08:
                best_score, best_lines, best_conf, best_deg = score, lines, conf, deg

    return "\n".join(best_lines), best_deg, best_conf, len(best_lines)


def render_pdf_pages_pil(pdf_bytes: bytes, scale: float = 2.0) -> List[Image.Image]:
    doc = pdfium.PdfDocument(pdf_bytes)
    images: List[Image.Image] = []
    try:
        for idx in range(len(doc)):
            page = doc[idx]
            bitmap = page.render(scale=scale)
            images.append(bitmap.to_pil())
            page.close()
    finally:
        doc.close()
    return images


def parse_page_ranges(range_str: str, total_pages: int) -> List[int]:
    """Parse 1-based page range string like '1-3, 5, 7-9' into sorted 0-based indices."""
    indices: List[int] = []
    seen = set()
    for part in range_str.split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            start_s, end_s = token.split("-", 1)
            start = max(1, int(start_s.strip()))
            end = min(total_pages, int(end_s.strip()))
            for p in range(start, end + 1):
                if p - 1 not in seen:
                    seen.add(p - 1)
                    indices.append(p - 1)
        else:
            p = int(token)
            if 1 <= p <= total_pages and (p - 1) not in seen:
                seen.add(p - 1)
                indices.append(p - 1)
    return indices


# ─────────────────────────────────────────────────────────────────────────────
# Core Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/")
async def serve_frontend():
    if not HTML_PATH.exists():
        raise HTTPException(status_code=404, detail="Markitdown_Lokal.html not found")
    return FileResponse(HTML_PATH)


@app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "engine": "folio-local",
        "features": ["microsoft-markitdown", "rapidocr-onnx", "pdf-tools"],
    }


# ─────────────────────────────────────────────────────────────────────────────
# Tab 1: Microsoft MarkItDown + RapidOCR Fallback
# ─────────────────────────────────────────────────────────────────────────────

@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    enable_frontmatter: bool = Form(True),
    enable_ocr: bool = Form(True),
    llm_api_key: Optional[str] = Form(None),
    llm_base_url: Optional[str] = Form(None),
    llm_model: Optional[str] = Form(None),
):
    filename = file.filename or "untitled"
    ext = Path(filename).suffix.lower()
    raw_bytes = await file.read()
    size = len(raw_bytes)
    method = (
        "Microsoft MarkItDown + Vision"
        if (llm_model and (llm_api_key or llm_base_url))
        else "Microsoft MarkItDown"
    )

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir) / f"input{ext}"
        tmp_path.write_bytes(raw_bytes)

        try:
            md_engine = create_markitdown_instance(
                llm_api_key=llm_api_key.strip() if llm_api_key else None,
                llm_base_url=llm_base_url.strip() if llm_base_url else None,
                llm_model=llm_model.strip() if llm_model else None,
            )
            result = md_engine.convert(str(tmp_path))
            text = (result.text_content or "").strip()
        except Exception as exc:
            return JSONResponse(status_code=500, content={"error": str(exc)})

    is_empty_pdf = ext == ".pdf" and len(text) == 0

    # Automatic server-side RapidOCR fallback for scanned/image-only PDFs
    if is_empty_pdf and enable_ocr:
        try:
            pages = render_pdf_pages_pil(raw_bytes, scale=2.0)
            ocr_sections: List[str] = []
            for idx, page_img in enumerate(pages, start=1):
                page_txt, _, _, _ = run_rapidocr_on_image(page_img, rotate_mode="auto")
                if len(pages) > 1:
                    ocr_sections.append(f"## Page {idx}\n\n{page_txt or '> [No text detected]'}")
                else:
                    ocr_sections.append(page_txt or "> [No text detected]")
            text = f"# {Path(filename).stem}\n\n" + "\n\n".join(ocr_sections)
            method = "RapidOCR (PDF Scan)"
            is_empty_pdf = False
        except Exception as ocr_exc:
            text = f"> [RapidOCR fallback failed: {ocr_exc}]"

    markdown = (
        build_frontmatter(filename, size, file.content_type or "", method) if enable_frontmatter else ""
    ) + text

    return {
        "filename": filename,
        "markdown": markdown,
        "raw_text_length": len(text),
        "is_empty_pdf": is_empty_pdf,
        "method": method,
    }


@app.post("/api/convert-url")
async def convert_url(
    url: str = Form(...),
    enable_frontmatter: bool = Form(True),
    llm_api_key: Optional[str] = Form(None),
    llm_base_url: Optional[str] = Form(None),
    llm_model: Optional[str] = Form(None),
):
    try:
        md_engine = create_markitdown_instance(
            llm_api_key=llm_api_key.strip() if llm_api_key else None,
            llm_base_url=llm_base_url.strip() if llm_base_url else None,
            llm_model=llm_model.strip() if llm_model else None,
        )
        result = md_engine.convert_url(url.strip())
        text = (result.text_content or "").strip()
    except Exception as exc:
        return JSONResponse(status_code=500, content={"error": str(exc)})

    method = "Microsoft MarkItDown (URL)"
    markdown = (
        build_frontmatter(url.strip(), len(text.encode("utf-8")), "text/uri-list", method)
        if enable_frontmatter
        else ""
    ) + text

    return {
        "filename": url.strip(),
        "markdown": markdown,
        "raw_text_length": len(text),
        "method": method,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Tab 2: Local RapidOCR (Image & Scanned PDF -> Clean Text Only)
# ─────────────────────────────────────────────────────────────────────────────

@app.post("/api/ocr")
async def ocr_extract(
    file: UploadFile = File(...),
    rotate: str = Form("auto"),
):
    filename = file.filename or "image.jpg"
    ext = Path(filename).suffix.lower()
    raw_bytes = await file.read()

    try:
        if ext == ".pdf":
            pages = render_pdf_pages_pil(raw_bytes, scale=2.0)
            parts: List[str] = []
            total_lines = 0
            confs: List[float] = []
            applied_deg = 0
            for idx, page_img in enumerate(pages, start=1):
                txt, deg, conf, count = run_rapidocr_on_image(page_img, rotate_mode=rotate)
                applied_deg = deg
                total_lines += count
                if conf > 0:
                    confs.append(conf)
                parts.append(f"--- Page {idx} ---\n{txt}" if len(pages) > 1 else txt)
            full_text = "\n\n".join(parts).strip()
            avg_conf = round((sum(confs) / len(confs)) * 100, 1) if confs else 0.0
            return {
                "filename": filename,
                "text": full_text,
                "rotation_applied": applied_deg,
                "line_count": total_lines,
                "page_count": len(pages),
                "avg_confidence": avg_conf,
                "engine": "RapidOCR (PaddleOCR ONNX)",
            }

        img = Image.open(io.BytesIO(raw_bytes))
        text, deg, conf, count = run_rapidocr_on_image(img, rotate_mode=rotate)
        return {
            "filename": filename,
            "text": text,
            "rotation_applied": deg,
            "line_count": count,
            "page_count": 1,
            "avg_confidence": round(conf * 100, 1),
            "engine": "RapidOCR (PaddleOCR ONNX)",
        }
    except Exception as exc:
        return JSONResponse(status_code=500, content={"error": str(exc)})


# ─────────────────────────────────────────────────────────────────────────────
# Tab 3: Local PDF Toolkit (iLovePDF Lokal)
# ─────────────────────────────────────────────────────────────────────────────

@app.post("/api/pdf/merge")
async def pdf_merge(files: List[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Upload at least 2 PDF files to merge.")
    writer = PdfWriter()
    try:
        for f in files:
            data = await f.read()
            reader = PdfReader(io.BytesIO(data))
            for page in reader.pages:
                writer.add_page(page)
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": 'attachment; filename="folio_merged.pdf"',
                "X-Page-Count": str(len(writer.pages)),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/split")
async def pdf_split(
    file: UploadFile = File(...),
    pages: str = Form(""),
    mode: str = Form("extract"),  # "extract" (single PDF of chosen pages) | "burst" (ZIP of 1-page PDFs)
):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        reader = PdfReader(io.BytesIO(raw))
        total = len(reader.pages)
        indices = parse_page_ranges(pages, total) if pages.strip() else list(range(total))
        if not indices:
            raise HTTPException(status_code=400, detail=f"No valid pages in range (PDF has {total} pages).")

        if mode == "burst":
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx in indices:
                    w = PdfWriter()
                    w.add_page(reader.pages[idx])
                    p_buf = io.BytesIO()
                    w.write(p_buf)
                    zf.writestr(f"{stem}_page_{idx + 1}.pdf", p_buf.getvalue())
            zip_buf.seek(0)
            return StreamingResponse(
                zip_buf,
                media_type="application/zip",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_split.zip"',
                    "X-Page-Count": str(len(indices)),
                },
            )

        writer = PdfWriter()
        for idx in indices:
            writer.add_page(reader.pages[idx])
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_extracted.pdf"',
                "X-Page-Count": str(len(indices)),
            },
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/rotate")
async def pdf_rotate(
    file: UploadFile = File(...),
    angle: int = Form(90),
    pages: str = Form(""),
):
    if angle not in (90, 180, 270):
        raise HTTPException(status_code=400, detail="Angle must be 90, 180, or 270.")
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        reader = PdfReader(io.BytesIO(raw))
        total = len(reader.pages)
        target_set = set(parse_page_ranges(pages, total)) if pages.strip() else set(range(total))
        writer = PdfWriter()
        for idx, page in enumerate(reader.pages):
            if idx in target_set:
                page.rotate(angle)
            writer.add_page(page)
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_rotated_{angle}.pdf"',
                "X-Page-Count": str(total),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/compress")
async def pdf_compress(file: UploadFile = File(...)):
    raw = await file.read()
    orig_size = len(raw)
    stem = Path(file.filename or "document").stem
    try:
        reader = PdfReader(io.BytesIO(raw))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        for page in writer.pages:
            try:
                page.compress_content_streams()
            except Exception:
                pass
            for img_file in getattr(page, "images", []):
                try:
                    pil_img = img_file.image
                    if pil_img is not None:
                        img_file.replace(pil_img.convert("RGB"), quality=72)
                except Exception:
                    pass
        try:
            writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
        except Exception:
            pass
        out = io.BytesIO()
        writer.write(out)
        compressed_bytes = out.getvalue()
        return StreamingResponse(
            io.BytesIO(compressed_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_compressed.pdf"',
                "X-Original-Size": str(orig_size),
                "X-Compressed-Size": str(len(compressed_bytes)),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/to-images")
async def pdf_to_images(
    file: UploadFile = File(...),
    img_format: str = Form("jpeg"),  # "jpeg" | "png"
    dpi: int = Form(150),
):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    fmt = "PNG" if img_format.lower() == "png" else "JPEG"
    ext = "png" if fmt == "PNG" else "jpg"
    scale = max(1.0, min(4.0, dpi / 72.0))

    try:
        pages = render_pdf_pages_pil(raw, scale=scale)
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, img in enumerate(pages, start=1):
                buf = io.BytesIO()
                if fmt == "JPEG":
                    img.convert("RGB").save(buf, format="JPEG", quality=90)
                else:
                    img.save(buf, format="PNG")
                zf.writestr(f"{stem}_page_{idx}.{ext}", buf.getvalue())
        zip_buf.seek(0)
        return StreamingResponse(
            zip_buf,
            media_type="application/zip",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_images.zip"',
                "X-Page-Count": str(len(pages)),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/from-images")
async def pdf_from_images(files: List[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="Upload at least 1 image file.")
    try:
        pil_images: List[Image.Image] = []
        for f in files:
            data = await f.read()
            im = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
            pil_images.append(im)
        out = io.BytesIO()
        pil_images[0].save(out, format="PDF", save_all=True, append_images=pil_images[1:])
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": 'attachment; filename="folio_images.pdf"',
                "X-Page-Count": str(len(pil_images)),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/remove-pages")
async def pdf_remove_pages(
    file: UploadFile = File(...),
    pages: str = Form(...),
):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        reader = PdfReader(io.BytesIO(raw))
        total = len(reader.pages)
        remove_set = set(parse_page_ranges(pages, total))
        if not remove_set:
            raise HTTPException(status_code=400, detail=f"Specify pages to remove (1-{total}).")
        kept = [reader.pages[i] for i in range(total) if i not in remove_set]
        if not kept:
            raise HTTPException(status_code=400, detail="Cannot remove all pages from a PDF.")
        writer = PdfWriter()
        for page in kept:
            writer.add_page(page)
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_trimmed.pdf"',
                "X-Page-Count": str(len(kept)),
            },
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/protect")
async def pdf_protect(
    file: UploadFile = File(...),
    password: str = Form(...),
):
    if not password.strip():
        raise HTTPException(status_code=400, detail="Password cannot be empty.")
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        reader = PdfReader(io.BytesIO(raw))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_protected.pdf"',
                "X-Page-Count": str(len(writer.pages)),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/unlock")
async def pdf_unlock(
    file: UploadFile = File(...),
    password: str = Form(""),
):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        reader = PdfReader(io.BytesIO(raw))
        if reader.is_encrypted:
            res = reader.decrypt(password)
            if res == 0:
                raise HTTPException(status_code=400, detail="Incorrect password for this encrypted PDF.")
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_unlocked.pdf"',
                "X-Page-Count": str(len(writer.pages)),
            },
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/to-markdown")
async def pdf_to_markdown(file: UploadFile = File(...)):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir) / "input.pdf"
        tmp_path.write_bytes(raw)
        try:
            md_engine = MarkItDown()
            result = md_engine.convert(str(tmp_path))
            text = (result.text_content or "").strip()
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc))

    if not text:
        pages = render_pdf_pages_pil(raw, scale=2.0)
        sections = []
        for idx, page_img in enumerate(pages, start=1):
            p_txt, _, _, _ = run_rapidocr_on_image(page_img, rotate_mode="auto")
            sections.append(f"## Page {idx}\n\n{p_txt}" if len(pages) > 1 else p_txt)
        text = f"# {stem}\n\n" + "\n\n".join(sections)

    return StreamingResponse(
        io.BytesIO(text.encode("utf-8")),
        media_type="text/markdown; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{stem}.md"'},
    )


@app.post("/api/pdf/ocr-pdf")
async def pdf_ocr_download(file: UploadFile = File(...)):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        pages = render_pdf_pages_pil(raw, scale=2.0)
        parts: List[str] = []
        for idx, page_img in enumerate(pages, start=1):
            txt, _, _, _ = run_rapidocr_on_image(page_img, rotate_mode="auto")
            parts.append(f"--- Page {idx} ---\n{txt}" if len(pages) > 1 else txt)
        out_text = "\n\n".join(parts).strip()
        return StreamingResponse(
            io.BytesIO(out_text.encode("utf-8")),
            media_type="text/plain; charset=utf-8",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_ocr.txt"',
                "X-Page-Count": str(len(pages)),
            },
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8765"))
    print(f"Folio-Local running at http://127.0.0.1:{port}")
    uvicorn.run(app, host="127.0.0.1", port=port)
