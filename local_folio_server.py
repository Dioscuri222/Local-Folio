import asyncio
import io
import os
import tempfile
import threading
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pypdfium2 as pdfium
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse
from markitdown import MarkItDown
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter
from rapidocr_onnxruntime import RapidOCR

APP_DIR = Path(__file__).resolve().parent
HTML_PATH = APP_DIR / "local-folio.html"

app = FastAPI(title="Local-Folio — Local PDF, OCR & Markdown Suite")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=[
        "X-Original-Size",
        "X-Compressed-Size",
        "X-Page-Count",
        "X-OCR-Pages",
        "X-Digital-Pages",
        "X-Redaction-Count",
        "X-Diff-Added",
        "X-Diff-Removed",
        "Content-Disposition",
    ],
)

# Thread-safe live job progress & cancellation registry
_jobs_lock = threading.Lock()
_jobs: Dict[str, dict] = {}


class JobCancelledError(Exception):
    """Raised when a user cancels an active processing job."""


def _job_init(job_id: Optional[str], total: int = 0, detail: str = "Starting...") -> None:
    if not job_id:
        return
    now = time.time()
    with _jobs_lock:
        if _jobs.get(job_id, {}).get("cancel_requested"):
            raise JobCancelledError("Cancelled by user.")
        # Prune stale finished jobs if registry grows large
        if len(_jobs) > 100:
            stale = [k for k, v in _jobs.items() if now - v.get("updated_at", 0) > 1800]
            for k in stale:
                _jobs.pop(k, None)
        _jobs[job_id] = {
            "status": "running",
            "current": 0,
            "total": total,
            "detail": detail,
            "cancel_requested": False,
            "updated_at": now,
        }


def _job_update(job_id: Optional[str], current: int, total: int, detail: str) -> None:
    if not job_id:
        return
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return
        if job.get("cancel_requested"):
            raise JobCancelledError("Cancelled by user.")
        job["current"] = current
        job["total"] = total
        job["detail"] = detail
        job["updated_at"] = time.time()


def _job_check_cancel(job_id: Optional[str]) -> None:
    if not job_id:
        return
    with _jobs_lock:
        job = _jobs.get(job_id)
        if job and job.get("cancel_requested"):
            raise JobCancelledError("Cancelled by user.")


def _job_finish(job_id: Optional[str], status: str, detail: str) -> None:
    if not job_id:
        return
    with _jobs_lock:
        job = _jobs.get(job_id)
        if job:
            job["status"] = status
            job["detail"] = detail
            if status == "done" and job["total"] > 0:
                job["current"] = job["total"]
            job["updated_at"] = time.time()

# Lazy thread-safe singleton for RapidOCR (auto-accelerates on CUDA / DirectML GPU when available)
_ocr_engine: Optional[RapidOCR] = None
_ocr_lock = threading.Lock()
_ocr_infer_lock = threading.Lock()


def get_ocr_engine() -> RapidOCR:
    global _ocr_engine
    if _ocr_engine is not None:
        return _ocr_engine

    with _ocr_lock:
        if _ocr_engine is not None:
            return _ocr_engine

        providers: List[str] = []
        try:
            import onnxruntime as ort

            providers = list(ort.get_available_providers())
        except Exception:
            providers = []

        # 1. Try NVIDIA CUDA GPU if onnxruntime-gpu + CUDA/cuDNN are available
        if "CUDAExecutionProvider" in providers:
            try:
                engine = RapidOCR(det_use_cuda=True, cls_use_cuda=True, rec_use_cuda=True)
                det_eps = engine.text_det.infer.session.get_providers()
                if "CUDAExecutionProvider" in det_eps:
                    _ocr_engine = engine
                    print("[Local-Folio] RapidOCR backend: CUDAExecutionProvider (GPU)")
                    return _ocr_engine
            except Exception as exc:
                print(f"[Local-Folio] CUDA GPU init skipped ({exc}); trying fallback...")

        # 2. Try Windows DirectML GPU (works on NVIDIA, AMD, and Intel GPUs via DirectX 12)
        if "DmlExecutionProvider" in providers:
            try:
                engine = RapidOCR(det_use_dml=True, cls_use_dml=True, rec_use_dml=True)
                det_eps = engine.text_det.infer.session.get_providers()
                if "DmlExecutionProvider" in det_eps:
                    _ocr_engine = engine
                    print("[Local-Folio] RapidOCR backend: DmlExecutionProvider (GPU)")
                    return _ocr_engine
            except Exception as exc:
                print(f"[Local-Folio] DirectML GPU init skipped ({exc}); falling back to CPU...")

        # 3. Reliable CPU fallback
        _ocr_engine = RapidOCR()
        print("[Local-Folio] RapidOCR backend: CPUExecutionProvider")
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


def _downscale_for_ocr(img: Image.Image, max_side: int = 2000) -> Image.Image:
    w, h = img.size
    if max(w, h) <= max_side:
        return img
    scale = max_side / float(max(w, h))
    return img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)


def _is_nearly_blank(img: Image.Image) -> bool:
    """Quickly detect blank/solid pages so RapidOCR does not waste extra passes."""
    small = img.convert("L").resize((64, 64))
    arr = np.array(small, dtype=np.float32)
    return float(arr.std()) < 4.0


def run_rapidocr_on_image(
    img: Image.Image,
    rotate_mode: str = "auto",
) -> Tuple[str, int, float, int]:
    """Run RapidOCR on a PIL Image with smart auto-orientation for rotated phone photos."""
    engine = get_ocr_engine()
    base_img = ImageOps.exif_transpose(img).convert("RGB")
    base_img = _downscale_for_ocr(base_img)

    if _is_nearly_blank(base_img):
        return "", 0, 0.0, 0

    if rotate_mode in {"0", "90", "180", "270"}:
        deg = int(rotate_mode)
        rotated = base_img.rotate(-deg, expand=True) if deg else base_img
        with _ocr_infer_lock:
            res, _ = engine(np.array(rotated))
        _, lines, avg_conf = _score_ocr_result(res)
        return "\n".join(lines), deg, avg_conf, len(lines)

    # Auto mode: test 0 deg first; only retry rotated orientations if boxes are vertical or very weak
    with _ocr_infer_lock:
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
        or (total_boxes > 0 and vertical_boxes / total_boxes >= 0.35)
        or (best_conf < 0.72 and total_boxes < 15)
    )

    if needs_rotation_check:
        for deg in (90, 270, 180):
            rotated = base_img.rotate(-deg, expand=True)
            with _ocr_infer_lock:
                res, _ = engine(np.array(rotated))
            score, lines, conf = _score_ocr_result(res)
            if score > best_score * 1.08:
                best_score, best_lines, best_conf, best_deg = score, lines, conf, deg

    return "\n".join(best_lines), best_deg, best_conf, len(best_lines)


def render_pdf_pages_pil(pdf_bytes: bytes, scale: float = 1.5) -> List[Image.Image]:
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


def iter_pdf_pages_pil(pdf_bytes: bytes, scale: float = 1.5):
    """Yield (1-based page_num, total_pages, pil_image) one page at a time to save memory."""
    doc = pdfium.PdfDocument(pdf_bytes)
    try:
        total = len(doc)
        for idx in range(total):
            page = doc[idx]
            try:
                bitmap = page.render(scale=scale)
                img = bitmap.to_pil()
            finally:
                page.close()
            yield idx + 1, total, img
    finally:
        doc.close()


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

FAVICON_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">'
    '<rect x="2" y="2" width="60" height="60" rx="16" fill="#efe6da" stroke="#e2d5c3" stroke-width="3"/>'
    '<g transform="translate(14 14) scale(1.5)" fill="none" stroke="#8f4f24" '
    'stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round">'
    '<path d="M10 2v2"/><path d="M14 2v2"/>'
    '<path d="M16 8a1 1 0 0 1 1 1v8a4 4 0 0 1-4 4H7a4 4 0 0 1-4-4V9a1 1 0 0 1 1-1h14a4 4 0 1 1 0 8h-1"/>'
    '<path d="M6 2v2"/></g></svg>'
)


@app.get("/")
async def serve_frontend():
    if not HTML_PATH.exists():
        raise HTTPException(status_code=404, detail="local-folio.html not found")
    return FileResponse(HTML_PATH)


@app.get("/favicon.ico")
async def serve_favicon():
    return Response(content=FAVICON_SVG, media_type="image/svg+xml")


@app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "engine": "local-folio",
        "features": ["microsoft-markitdown", "rapidocr-onnx", "pdf-tools", "live-progress"],
    }


@app.get("/api/jobs/{job_id}")
async def get_job_progress(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            return {"status": "unknown", "current": 0, "total": 0, "pct": 0, "detail": "Waiting..."}
        total = job["total"]
        current = job["current"]
        pct = int(round((current / total) * 100)) if total > 0 else 0
        return {
            "status": job["status"],
            "current": current,
            "total": total,
            "pct": min(100, max(0, pct)),
            "detail": job["detail"],
        }


@app.post("/api/jobs/{job_id}/cancel")
async def cancel_job(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
        if not job:
            _jobs[job_id] = {
                "status": "cancelled",
                "current": 0,
                "total": 0,
                "detail": "Cancelled by user.",
                "cancel_requested": True,
                "updated_at": time.time(),
            }
        else:
            job["cancel_requested"] = True
            job["status"] = "cancelled"
            job["detail"] = "Cancelling..."
            job["updated_at"] = time.time()
    return {"status": "cancelled", "job_id": job_id}


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
    job_id: Optional[str] = Form(None),
):
    filename = file.filename or "untitled"
    ext = Path(filename).suffix.lower()
    raw_bytes = await file.read()
    size = len(raw_bytes)
    content_type = file.content_type or ""

    def _run():
        method = (
            "Microsoft MarkItDown + Vision"
            if (llm_model and (llm_api_key or llm_base_url))
            else "Microsoft MarkItDown"
        )
        _job_init(job_id, total=0, detail=f"Converting {filename}...")

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir) / f"input{ext}"
            tmp_path.write_bytes(raw_bytes)

            md_engine = create_markitdown_instance(
                llm_api_key=llm_api_key.strip() if llm_api_key else None,
                llm_base_url=llm_base_url.strip() if llm_base_url else None,
                llm_model=llm_model.strip() if llm_model else None,
            )
            result = md_engine.convert(str(tmp_path))
            text = (result.text_content or "").strip()

        is_empty_pdf = ext == ".pdf" and len(text) == 0

        # Automatic server-side RapidOCR fallback for scanned/image-only PDFs
        if is_empty_pdf and enable_ocr:
            try:
                ocr_sections: List[str] = []
                total_pages = 0
                for page_num, total, page_img in iter_pdf_pages_pil(raw_bytes, scale=1.5):
                    total_pages = total
                    _job_update(job_id, page_num - 1, total, f"OCR page {page_num} of {total} ({filename})")
                    page_txt, _, _, _ = run_rapidocr_on_image(page_img, rotate_mode="0")
                    if total > 1:
                        ocr_sections.append(f"## Page {page_num}\n\n{page_txt or '> [No text detected]'}")
                    else:
                        ocr_sections.append(page_txt or "> [No text detected]")
                    _job_update(job_id, page_num, total, f"OCR page {page_num} of {total} ({filename})")
                text = f"# {Path(filename).stem}\n\n" + "\n\n".join(ocr_sections)
                method = "RapidOCR (PDF Scan)"
                is_empty_pdf = False
            except JobCancelledError:
                raise
            except Exception as ocr_exc:
                text = f"> [RapidOCR fallback failed: {ocr_exc}]"

        markdown = (
            build_frontmatter(filename, size, content_type, method) if enable_frontmatter else ""
        ) + text
        _job_finish(job_id, "done", f"Converted {filename}")
        return {
            "filename": filename,
            "markdown": markdown,
            "raw_text_length": len(text),
            "is_empty_pdf": is_empty_pdf,
            "method": method,
        }

    try:
        return await asyncio.to_thread(_run)
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        return JSONResponse(status_code=499, content={"error": str(exc)})
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
        return JSONResponse(status_code=500, content={"error": str(exc)})


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
    job_id: Optional[str] = Form(None),
):
    filename = file.filename or "image.jpg"
    ext = Path(filename).suffix.lower()
    raw_bytes = await file.read()

    def _run():
        if ext == ".pdf":
            pdf_rotate = "0" if rotate == "auto" else rotate
            parts: List[str] = []
            total_lines = 0
            confs: List[float] = []
            applied_deg = 0
            total_pages = 0
            _job_init(job_id, total=0, detail=f"Opening {filename}...")
            for page_num, total, page_img in iter_pdf_pages_pil(raw_bytes, scale=1.5):
                total_pages = total
                _job_update(job_id, page_num - 1, total, f"OCR page {page_num} of {total} ({filename})")
                txt, deg, conf, count = run_rapidocr_on_image(page_img, rotate_mode=pdf_rotate)
                applied_deg = deg
                total_lines += count
                if conf > 0:
                    confs.append(conf)
                parts.append(f"--- Page {page_num} ---\n{txt}" if total > 1 else txt)
                _job_update(job_id, page_num, total, f"OCR page {page_num} of {total} ({filename})")
            full_text = "\n\n".join(parts).strip()
            avg_conf = round((sum(confs) / len(confs)) * 100, 1) if confs else 0.0
            _job_finish(job_id, "done", f"Completed {filename}")
            return {
                "filename": filename,
                "text": full_text,
                "rotation_applied": applied_deg,
                "line_count": total_lines,
                "page_count": total_pages,
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

    try:
        return await asyncio.to_thread(_run)
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        return JSONResponse(status_code=499, content={"error": str(exc)})
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
        return JSONResponse(status_code=500, content={"error": str(exc)})


# ─────────────────────────────────────────────────────────────────────────────
# Tab 3: Local PDF Toolkit (Local iLovePDF)
# ─────────────────────────────────────────────────────────────────────────────

@app.post("/api/pdf/merge")
async def pdf_merge(
    files: List[UploadFile] = File(...),
    add_bookmarks: bool = Form(True),
    page_size: str = Form("keep"),  # "keep" | "a4" | "letter"
    job_id: Optional[str] = Form(None),
):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Upload at least 2 PDF files to merge.")
    raw_items = [(f.filename or f"file_{i + 1}.pdf", await f.read()) for i, f in enumerate(files)]

    def _run():
        import pymupdf

        _job_init(job_id, total=len(raw_items), detail=f"Merging {len(raw_items)} PDF files...")
        out_doc = pymupdf.open()
        toc = []
        size_map = {"a4": (595.28, 841.89), "letter": (612.0, 792.0)}
        target_dims = size_map.get(page_size.lower())

        try:
            for idx, (fname, data) in enumerate(raw_items, start=1):
                _job_update(job_id, idx - 1, len(raw_items), f"Adding {fname} ({idx}/{len(raw_items)})...")
                src = pymupdf.open(stream=data, filetype="pdf")
                try:
                    start_page = len(out_doc) + 1
                    if target_dims:
                        tw, th = target_dims
                        for p_i in range(len(src)):
                            sr = src[p_i].rect
                            # Keep each page's orientation when normalizing page size
                            pw, ph = (th, tw) if sr.width > sr.height else (tw, th)
                            new_p = out_doc.new_page(width=pw, height=ph)
                            new_p.show_pdf_page(new_p.rect, src, p_i, keep_proportion=True)
                    else:
                        out_doc.insert_pdf(src)
                    if add_bookmarks and len(src) > 0:
                        toc.append([1, Path(fname).stem, start_page])
                finally:
                    src.close()
                _job_update(job_id, idx, len(raw_items), f"Added {fname} ({idx}/{len(raw_items)})")

            if add_bookmarks and toc:
                out_doc.set_toc(toc)
            total_pages = len(out_doc)
            out_bytes = out_doc.tobytes(deflate=True)
        finally:
            out_doc.close()

        _job_finish(job_id, "done", f"Merged {len(raw_items)} files ({total_pages} pages)")
        return out_bytes, total_pages

    try:
        out_bytes, total_pages = await asyncio.to_thread(_run)
        return StreamingResponse(
            io.BytesIO(out_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": 'attachment; filename="local-folio_merged.pdf"',
                "X-Page-Count": str(total_pages),
            },
        )
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        raise HTTPException(status_code=499, detail=str(exc))
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
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
async def pdf_compress(
    file: UploadFile = File(...),
    level: str = Form("recommended"),  # "low" | "recommended" | "extreme"
    job_id: Optional[str] = Form(None),
):
    raw = await file.read()
    orig_size = len(raw)
    stem = Path(file.filename or "document").stem
    lvl = level.lower()
    quality = 88 if lvl == "low" else 45 if lvl == "extreme" else 72
    max_dim = 2400 if lvl == "low" else 1100 if lvl == "extreme" else 1600

    def _run():
        reader = PdfReader(io.BytesIO(raw))
        total = len(reader.pages)
        _job_init(job_id, total=total, detail=f"Preparing {total} pages ({lvl} compression)...")
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        for idx, page in enumerate(writer.pages, start=1):
            _job_update(job_id, idx - 1, total, f"Compressing page {idx} of {total} ({lvl})...")
            try:
                page.compress_content_streams()
            except Exception:
                pass
            for img_file in getattr(page, "images", []):
                _job_check_cancel(job_id)
                try:
                    pil_img = img_file.image
                    if pil_img is not None:
                        rgb = pil_img.convert("RGB")
                        w, h = rgb.size
                        if max(w, h) > max_dim:
                            scale = max_dim / float(max(w, h))
                            rgb = rgb.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.Resampling.LANCZOS)
                        img_file.replace(rgb, quality=quality)
                except Exception:
                    pass
            _job_update(job_id, idx, total, f"Compressed page {idx} of {total}")
        _job_check_cancel(job_id)
        try:
            writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
        except Exception:
            pass
        out = io.BytesIO()
        writer.write(out)
        compressed_bytes = out.getvalue()
        _job_finish(job_id, "done", f"Compressed {total} pages ({lvl})")
        return compressed_bytes

    try:
        compressed_bytes = await asyncio.to_thread(_run)
        return StreamingResponse(
            io.BytesIO(compressed_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_compressed.pdf"',
                "X-Original-Size": str(orig_size),
                "X-Compressed-Size": str(len(compressed_bytes)),
            },
        )
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        raise HTTPException(status_code=499, detail=str(exc))
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/to-images")
async def pdf_to_images(
    file: UploadFile = File(...),
    img_format: str = Form("jpeg"),  # "jpeg" | "png"
    dpi: int = Form(150),
    pages: str = Form(""),
    job_id: Optional[str] = Form(None),
):
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    fmt = "PNG" if img_format.lower() == "png" else "JPEG"
    ext = "png" if fmt == "PNG" else "jpg"
    scale = max(1.0, min(4.0, dpi / 72.0))

    def _run():
        doc = pdfium.PdfDocument(raw)
        try:
            total = len(doc)
            indices = parse_page_ranges(pages, total) if pages.strip() else list(range(total))
            if not indices:
                raise ValueError(f"No valid pages in range (PDF has {total} pages).")
            _job_init(job_id, total=len(indices), detail=f"Rendering {len(indices)} page(s)...")
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for step_i, idx in enumerate(indices, start=1):
                    _job_update(job_id, step_i - 1, len(indices), f"Rendering page {idx + 1} ({step_i}/{len(indices)}, {dpi} DPI)...")
                    page = doc[idx]
                    try:
                        img = page.render(scale=scale).to_pil()
                    finally:
                        page.close()
                    buf = io.BytesIO()
                    if fmt == "JPEG":
                        img.convert("RGB").save(buf, format="JPEG", quality=90)
                    else:
                        img.save(buf, format="PNG")
                    zf.writestr(f"{stem}_page_{idx + 1}.{ext}", buf.getvalue())
                    _job_update(job_id, step_i, len(indices), f"Rendered page {idx + 1} ({step_i}/{len(indices)})")
        finally:
            doc.close()
        zip_buf.seek(0)
        _job_finish(job_id, "done", f"Rendered {len(indices)} pages")
        return zip_buf, len(indices)

    try:
        zip_buf, total_pages = await asyncio.to_thread(_run)
        return StreamingResponse(
            zip_buf,
            media_type="application/zip",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_images.zip"',
                "X-Page-Count": str(total_pages),
            },
        )
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        raise HTTPException(status_code=499, detail=str(exc))
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
        raise HTTPException(status_code=400 if isinstance(exc, ValueError) else 500, detail=str(exc))


@app.post("/api/pdf/from-images")
async def pdf_from_images(
    files: List[UploadFile] = File(...),
    orientation: str = Form("portrait"),  # "portrait" | "landscape"
    page_size: str = Form("a4"),          # "fit" | "a4" | "letter"
    margin: str = Form("none"),           # "none" | "small" | "big"
    merge_all: bool = Form(True),
    job_id: Optional[str] = Form(None),
):
    if not files:
        raise HTTPException(status_code=400, detail="Upload at least 1 image file.")
    raw_items = [(f.filename or f"image_{i + 1}.jpg", await f.read()) for i, f in enumerate(files)]

    def _add_image_page(doc, img_bytes: bytes):
        import pymupdf

        im = ImageOps.exif_transpose(Image.open(io.BytesIO(img_bytes))).convert("RGB")
        iw, ih = im.size
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=92)
        clean_bytes = buf.getvalue()

        # Margin in PDF points (1 mm = 72 / 25.4 pt)
        margin_pt = 0.0 if margin == "none" else (28.35 if margin == "small" else 70.87)  # 0mm / 10mm / 25mm

        ps = page_size.lower()
        if ps == "fit":
            # Match image aspect ratio at 72 pt/in (or 96 dpi -> 0.75 scale) + optional margin
            pw = max(72.0, iw * 0.75) + margin_pt * 2
            ph = max(72.0, ih * 0.75) + margin_pt * 2
            if orientation == "landscape" and ph > pw:
                pw, ph = ph, pw
            elif orientation == "portrait" and pw > ph:
                pw, ph = ph, pw
        else:
            base_w, base_h = (612.0, 792.0) if ps == "letter" else (595.28, 841.89)  # US Letter or A4
            pw, ph = (base_h, base_w) if orientation == "landscape" else (base_w, base_h)

        page = doc.new_page(width=pw, height=ph)
        inner = pymupdf.Rect(margin_pt, margin_pt, max(margin_pt + 10, pw - margin_pt), max(margin_pt + 10, ph - margin_pt))
        page.insert_image(inner, stream=clean_bytes, keep_proportion=True)

    def _run():
        import pymupdf

        total = len(raw_items)
        _job_init(job_id, total=total, detail=f"Converting {total} image(s) to PDF...")

        if merge_all or total == 1:
            doc = pymupdf.open()
            try:
                for idx, (fname, data) in enumerate(raw_items, start=1):
                    _job_update(job_id, idx - 1, total, f"Placing {fname} ({idx}/{total})...")
                    _add_image_page(doc, data)
                    _job_update(job_id, idx, total, f"Added {fname} ({idx}/{total})")
                out_bytes = doc.tobytes(deflate=True)
            finally:
                doc.close()
            _job_finish(job_id, "done", f"Created PDF from {total} image(s)")
            return out_bytes, "application/pdf", "local-folio_images.pdf", total

        # Separate 1-page PDF per image inside a .ZIP archive
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, (fname, data) in enumerate(raw_items, start=1):
                _job_update(job_id, idx - 1, total, f"Creating PDF for {fname} ({idx}/{total})...")
                single = pymupdf.open()
                try:
                    _add_image_page(single, data)
                    zf.writestr(f"{Path(fname).stem}_{idx}.pdf", single.tobytes(deflate=True))
                finally:
                    single.close()
                _job_update(job_id, idx, total, f"Created PDF {idx}/{total}")
        zip_buf.seek(0)
        _job_finish(job_id, "done", f"Created {total} separate PDFs (.ZIP)")
        return zip_buf.getvalue(), "application/zip", "local-folio_images_pdfs.zip", total

    try:
        out_bytes, media_type, out_name, count = await asyncio.to_thread(_run)
        return StreamingResponse(
            io.BytesIO(out_bytes),
            media_type=media_type,
            headers={
                "Content-Disposition": f'attachment; filename="{out_name}"',
                "X-Page-Count": str(count),
            },
        )
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        raise HTTPException(status_code=499, detail=str(exc))
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
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
    allow_print: bool = Form(True),
    allow_copy: bool = Form(True),
):
    if not password.strip():
        raise HTTPException(status_code=400, detail="Password cannot be empty.")
    raw = await file.read()
    stem = Path(file.filename or "document").stem
    try:
        import pymupdf

        doc = pymupdf.open(stream=raw, filetype="pdf")
        try:
            perms = pymupdf.PDF_PERM_ACCESSIBILITY
            if allow_print:
                perms |= pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ
            if allow_copy:
                perms |= pymupdf.PDF_PERM_COPY
            out_bytes = doc.tobytes(
                encryption=pymupdf.PDF_ENCRYPT_AES_256,
                user_pw=password,
                owner_pw=password,
                permissions=perms,
                deflate=True,
            )
            page_count = len(doc)
        finally:
            doc.close()
        return StreamingResponse(
            io.BytesIO(out_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_protected.pdf"',
                "X-Page-Count": str(page_count),
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
async def pdf_to_markdown(
    file: UploadFile = File(...),
    enable_frontmatter: bool = Form(False),
    enable_ocr: bool = Form(True),
    job_id: Optional[str] = Form(None),
):
    raw = await file.read()
    filename = file.filename or "document.pdf"
    stem = Path(filename).stem

    def _run():
        _job_init(job_id, total=0, detail=f"Converting {stem}.pdf to Markdown...")
        method = "Microsoft MarkItDown"
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir) / "input.pdf"
            tmp_path.write_bytes(raw)
            md_engine = MarkItDown()
            result = md_engine.convert(str(tmp_path))
            text = (result.text_content or "").strip()

        if not text and enable_ocr:
            method = "RapidOCR (PDF Scan)"
            sections = []
            for page_num, total, page_img in iter_pdf_pages_pil(raw, scale=1.5):
                _job_update(job_id, page_num - 1, total, f"Scanned PDF — OCR page {page_num} of {total}...")
                p_txt, _, _, _ = run_rapidocr_on_image(page_img, rotate_mode="0")
                sections.append(f"## Page {page_num}\n\n{p_txt}" if total > 1 else p_txt)
                _job_update(job_id, page_num, total, f"Scanned PDF — OCR page {page_num} of {total}")
            text = f"# {stem}\n\n" + "\n\n".join(sections)

        if enable_frontmatter:
            text = build_frontmatter(filename, len(raw), "application/pdf", method) + text

        _job_finish(job_id, "done", f"Converted {stem}.pdf to Markdown")
        return text

    try:
        text = await asyncio.to_thread(_run)
        return StreamingResponse(
            io.BytesIO(text.encode("utf-8")),
            media_type="text/markdown; charset=utf-8",
            headers={"Content-Disposition": f'attachment; filename="{stem}.md"'},
        )
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        raise HTTPException(status_code=499, detail=str(exc))
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/pdf/ocr-pdf")
async def pdf_ocr_download(
    file: UploadFile = File(...),
    ocr_mode: str = Form("smart"),  # "smart" (digital text + OCR scanned pages) | "force" (OCR every page)
    job_id: Optional[str] = Form(None),
):
    raw = await file.read()
    stem = Path(file.filename or "document").stem

    def _run():
        # Read embedded text layer first so Smart Hybrid mode skips OCR on digital text pages
        embedded_texts: List[str] = []
        if ocr_mode != "force":
            try:
                reader = PdfReader(io.BytesIO(raw))
                for p in reader.pages:
                    embedded_texts.append((p.extract_text() or "").strip())
            except Exception:
                embedded_texts = []

        doc = pdfium.PdfDocument(raw)
        parts: List[str] = []
        ocr_pages = 0
        digital_pages = 0
        try:
            total = len(doc)
            _job_init(job_id, total=total, detail=f"Preparing {total} pages...")
            for idx in range(total):
                page_num = idx + 1
                _job_check_cancel(job_id)

                digital_txt = embedded_texts[idx] if idx < len(embedded_texts) else ""
                if ocr_mode != "force" and len(digital_txt) >= 25:
                    digital_pages += 1
                    _job_update(job_id, page_num, total, f"Page {page_num} of {total} (digital text)")
                    parts.append(f"--- Page {page_num} ---\n{digital_txt}" if total > 1 else digital_txt)
                    continue

                _job_update(job_id, page_num - 1, total, f"Page {page_num} of {total} — Running RapidOCR...")
                page = doc[idx]
                try:
                    bitmap = page.render(scale=1.5)
                    page_img = bitmap.to_pil()
                finally:
                    page.close()

                txt, _, _, _ = run_rapidocr_on_image(page_img, rotate_mode="0")
                ocr_pages += 1
                parts.append(f"--- Page {page_num} ---\n{txt}" if total > 1 else txt)
                _job_update(job_id, page_num, total, f"Page {page_num} of {total} — OCR done")
        finally:
            doc.close()

        out_text = "\n\n".join(parts).strip()
        summary = f"Completed {total} pages ({digital_pages} digital, {ocr_pages} OCR)"
        _job_finish(job_id, "done", summary)
        return out_text, total, ocr_pages, digital_pages

    try:
        out_text, total, ocr_pages, digital_pages = await asyncio.to_thread(_run)
        return StreamingResponse(
            io.BytesIO(out_text.encode("utf-8")),
            media_type="text/plain; charset=utf-8",
            headers={
                "Content-Disposition": f'attachment; filename="{stem}_ocr.txt"',
                "X-Page-Count": str(total),
                "X-OCR-Pages": str(ocr_pages),
                "X-Digital-Pages": str(digital_pages),
            },
        )
    except JobCancelledError as exc:
        _job_finish(job_id, "cancelled", str(exc))
        raise HTTPException(status_code=499, detail=str(exc))
    except Exception as exc:
        _job_finish(job_id, "error", str(exc))
        raise HTTPException(status_code=500, detail=str(exc))


# Register all extended PDF & Office tools (Watermark, Page Numbers, Crop, Redact, Repair,
# Organize, Sign, Compare, PDF<->Word/PPTX/Excel, HTML->PDF, PDF/A, AI Summarizer, Workflow)
from pdf_extended_routes import build_extended_pdf_router

app.include_router(
    build_extended_pdf_router(
        job_init=_job_init,
        job_update=_job_update,
        job_check_cancel=_job_check_cancel,
        job_finish=_job_finish,
        job_cancelled_exc=JobCancelledError,
        parse_page_ranges=parse_page_ranges,
        iter_pdf_pages_pil=iter_pdf_pages_pil,
        run_rapidocr_on_image=run_rapidocr_on_image,
    )
)


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8765"))
    print(f"Local-Folio running at http://127.0.0.1:{port}")
    uvicorn.run(app, host="127.0.0.1", port=port)
