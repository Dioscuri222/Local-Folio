import asyncio
import base64
import difflib
import hashlib
import io
import json
import os
import shutil
import subprocess
import tempfile
import threading
from collections import OrderedDict
from pathlib import Path
from typing import Callable, List, Optional

import pdfplumber
import pikepdf
import pymupdf
from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse
from openpyxl import Workbook
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from pypdf import PdfReader, PdfWriter


def _find_browser_binary() -> Optional[str]:
    candidates = [
        shutil.which("msedge"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        shutil.which("chrome"),
        shutil.which("google-chrome"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        shutil.which("chromium"),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


def _find_soffice_binary() -> Optional[str]:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


def _find_gs_binary() -> Optional[str]:
    for name in ("gswin64c", "gswin32c", "gs"):
        found = shutil.which(name)
        if found:
            return found
    return None


def _convert_office_to_pdf_sync(
    input_path: Path,
    output_path: Path,
    app_type: str,
    quality: str = "standard",
    orientation: str = "auto",
    fit_columns: bool = False,
) -> None:
    """Convert Word/Excel/PowerPoint to PDF via Windows MS Office COM or LibreOffice headless."""
    # 1. Try native Microsoft Office COM on Windows (highest fidelity)
    if os.name == "nt":
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            try:
                in_str = str(input_path.resolve())
                out_str = str(output_path.resolve())
                if app_type == "word":
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                    try:
                        doc = word.Documents.Open(in_str, ReadOnly=True)
                        try:
                            opt_for = 1 if quality == "minimal" else 0  # 0=Print, 1=OnScreen
                            doc.ExportAsFixedFormat(out_str, 17, False, opt_for)
                        finally:
                            doc.Close(False)
                    finally:
                        word.Quit()
                    if output_path.exists():
                        return
                elif app_type == "excel":
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    try:
                        wb = excel.Workbooks.Open(in_str, ReadOnly=True)
                        try:
                            if orientation in ("portrait", "landscape") or fit_columns:
                                for ws in wb.Worksheets:
                                    if orientation == "portrait":
                                        ws.PageSetup.Orientation = 1  # xlPortrait
                                    elif orientation == "landscape":
                                        ws.PageSetup.Orientation = 2  # xlLandscape
                                    if fit_columns:
                                        ws.PageSetup.Zoom = False
                                        ws.PageSetup.FitToPagesWide = 1
                                        ws.PageSetup.FitToPagesTall = False
                            xl_quality = 1 if quality == "minimal" else 0
                            wb.ExportAsFixedFormat(0, out_str, xl_quality)
                        finally:
                            wb.Close(False)
                    finally:
                        excel.Quit()
                    if output_path.exists():
                        return
                elif app_type == "powerpoint":
                    ppt = win32com.client.DispatchEx("PowerPoint.Application")
                    try:
                        pres = ppt.Presentations.Open(in_str, WithWindow=False)
                        try:
                            pres.SaveAs(out_str, 32)  # 32 = ppSaveAsPDF
                        finally:
                            pres.Close()
                    finally:
                        ppt.Quit()
                    if output_path.exists():
                        return
            finally:
                pythoncom.CoUninitialize()
        except Exception as com_exc:
            soffice = _find_soffice_binary()
            if not soffice:
                raise RuntimeError(
                    f"Microsoft {app_type.title()} COM conversion failed ({com_exc}) and LibreOffice was not found."
                )

    # 2. Fallback to LibreOffice headless
    soffice = _find_soffice_binary()
    if not soffice:
        raise RuntimeError(
            "Neither Microsoft Office (Windows COM) nor LibreOffice (soffice) was found on this machine."
        )
    res = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_path.parent), str(input_path)],
        capture_output=True,
        text=True,
        timeout=180,
    )
    generated = output_path.parent / f"{input_path.stem}.pdf"
    if res.returncode != 0 or not generated.exists():
        raise RuntimeError(res.stderr.strip() or res.stdout.strip() or "LibreOffice conversion failed.")
    if generated != output_path:
        generated.replace(output_path)


def _apply_watermark_doc(
    doc: pymupdf.Document,
    text: str,
    opacity: float = 0.25,
    angle: int = 45,
    font_size: int = 42,
    position: str = "center",
    job_update: Optional[Callable] = None,
) -> None:
    total = len(doc)
    clean_text = text.strip() or "CONFIDENTIAL"
    alpha = max(0.05, min(1.0, opacity))
    # PyMuPDF shape fill_opacity gives clean transparency across every PDF viewer
    for idx, page in enumerate(doc):
        if job_update:
            job_update(idx, total, f"Watermarking page {idx + 1} of {total}...")
        rect = page.rect
        cx, cy = rect.width / 2.0, rect.height / 2.0
        if position == "top":
            cy = rect.height * 0.14
        elif position == "bottom":
            cy = rect.height * 0.88

        # Build centered textbox/point with rotation via Morph
        tw = pymupdf.get_text_length(clean_text, fontname="helv", fontsize=font_size)
        pt = pymupdf.Point(cx - tw / 2.0, cy + font_size * 0.35)
        pivot = pymupdf.Point(cx, cy)
        mat = pymupdf.Matrix(-angle)
        page.insert_text(
            pt,
            clean_text,
            fontname="helv",
            fontsize=font_size,
            color=(0.35, 0.24, 0.16),
            fill_opacity=alpha,
            morph=(pivot, mat),
            overlay=True,
        )
        if job_update:
            job_update(idx + 1, total, f"Watermarked page {idx + 1} of {total}")


def _apply_page_numbers_doc(
    doc: pymupdf.Document,
    position: str = "bottom-center",
    fmt: str = "page_of",
    start_num: int = 1,
    font_size: int = 11,
    job_update: Optional[Callable] = None,
) -> None:
    total = len(doc)
    for idx, page in enumerate(doc):
        if job_update:
            job_update(idx, total, f"Numbering page {idx + 1} of {total}...")
        num = start_num + idx
        total_display = start_num + total - 1
        if fmt == "simple":
            label = str(num)
        elif fmt == "slash":
            label = f"{num} / {total_display}"
        else:
            label = f"Page {num} of {total_display}"

        rect = page.rect
        margin_x = 36
        margin_y = 26
        tw = pymupdf.get_text_length(label, fontname="helv", fontsize=font_size)

        if "left" in position:
            x = margin_x
        elif "right" in position:
            x = max(margin_x, rect.width - margin_x - tw)
        else:
            x = (rect.width - tw) / 2.0

        y = margin_y + font_size if position.startswith("top") else rect.height - margin_y

        page.insert_text(
            pymupdf.Point(x, y),
            label,
            fontname="helv",
            fontsize=font_size,
            color=(0.25, 0.20, 0.16),
            overlay=True,
        )
        if job_update:
            job_update(idx + 1, total, f"Numbered page {idx + 1} of {total}")


def build_extended_pdf_router(
    *,
    job_init: Callable,
    job_update: Callable,
    job_check_cancel: Callable,
    job_finish: Callable,
    job_cancelled_exc: type,
    parse_page_ranges: Callable,
    iter_pdf_pages_pil: Callable,
    run_rapidocr_on_image: Callable,
) -> APIRouter:
    router = APIRouter()

    # ── 1. Watermark PDF ──────────────────────────────────────────────────────
    @router.post("/api/pdf/watermark")
    async def pdf_watermark(
        file: UploadFile = File(...),
        text: str = Form("CONFIDENTIAL"),
        opacity: float = Form(0.25),
        angle: int = Form(45),
        font_size: int = Form(44),
        position: str = Form("center"),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            doc = pymupdf.open(stream=raw, filetype="pdf")
            try:
                job_init(job_id, total=len(doc), detail="Applying watermark...")
                _apply_watermark_doc(
                    doc,
                    text=text,
                    opacity=opacity,
                    angle=angle,
                    font_size=font_size,
                    position=position,
                    job_update=lambda c, t, d: job_update(job_id, c, t, d),
                )
                out = doc.tobytes(deflate=True)
                total = len(doc)
            finally:
                doc.close()
            job_finish(job_id, "done", f"Watermarked {total} pages")
            return out, total

        try:
            out, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_watermarked.pdf"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 2. Page Numbers ───────────────────────────────────────────────────────
    @router.post("/api/pdf/page-numbers")
    async def pdf_page_numbers(
        file: UploadFile = File(...),
        position: str = Form("bottom-center"),
        fmt: str = Form("page_of"),
        start_num: int = Form(1),
        font_size: int = Form(11),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            doc = pymupdf.open(stream=raw, filetype="pdf")
            try:
                job_init(job_id, total=len(doc), detail="Adding page numbers...")
                _apply_page_numbers_doc(
                    doc,
                    position=position,
                    fmt=fmt,
                    start_num=start_num,
                    font_size=font_size,
                    job_update=lambda c, t, d: job_update(job_id, c, t, d),
                )
                out = doc.tobytes(deflate=True)
                total = len(doc)
            finally:
                doc.close()
            job_finish(job_id, "done", f"Numbered {total} pages")
            return out, total

        try:
            out, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_numbered.pdf"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 3. Crop PDF ───────────────────────────────────────────────────────────
    @router.post("/api/pdf/crop")
    async def pdf_crop(
        file: UploadFile = File(...),
        top: float = Form(20.0),
        right: float = Form(20.0),
        bottom: float = Form(20.0),
        left: float = Form(20.0),
        unit: str = Form("mm"),  # "mm" | "pt" | "pct"
        pages: str = Form(""),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            doc = pymupdf.open(stream=raw, filetype="pdf")
            try:
                total = len(doc)
                job_init(job_id, total=total, detail="Cropping PDF margins...")
                target_set = set(parse_page_ranges(pages, total)) if pages.strip() else set(range(total))
                for idx, page in enumerate(doc):
                    job_update(job_id, idx, total, f"Cropping page {idx + 1} of {total}...")
                    if idx not in target_set:
                        continue
                    r = page.rect
                    if unit == "pct":
                        dt = r.height * (top / 100.0)
                        db = r.height * (bottom / 100.0)
                        dl = r.width * (left / 100.0)
                        dr = r.width * (right / 100.0)
                    elif unit == "mm":
                        factor = 72.0 / 25.4
                        dt, dr, db, dl = top * factor, right * factor, bottom * factor, left * factor
                    else:
                        dt, dr, db, dl = top, right, bottom, left

                    new_rect = pymupdf.Rect(
                        r.x0 + dl,
                        r.y0 + dt,
                        r.x1 - dr,
                        r.y1 - db,
                    )
                    if new_rect.width < 20 or new_rect.height < 20:
                        raise ValueError("Crop margins are too large for this page size.")
                    page.set_cropbox(new_rect)
                    job_update(job_id, idx + 1, total, f"Cropped page {idx + 1} of {total}")
                out = doc.tobytes(deflate=True)
            finally:
                doc.close()
            job_finish(job_id, "done", f"Cropped {total} pages")
            return out, total

        try:
            out, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_cropped.pdf"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=400 if isinstance(exc, ValueError) else 500, detail=str(exc))

    # ── 4. Redact PDF (Permanent Text Removal) ────────────────────────────────
    @router.post("/api/pdf/redact")
    async def pdf_redact(
        file: UploadFile = File(...),
        terms: str = Form(...),
        fill_color: str = Form("black"),  # "black" | "brown" | "white"
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem
        term_list = [t.strip() for part in terms.splitlines() for t in part.split(",") if t.strip()]
        if not term_list:
            raise HTTPException(status_code=400, detail="Enter at least one word or phrase to redact.")
        color_map = {"black": (0, 0, 0), "brown": (0.28, 0.17, 0.11), "white": (1, 1, 1)}
        rgb = color_map.get(fill_color.lower(), (0, 0, 0))

        def _run():
            doc = pymupdf.open(stream=raw, filetype="pdf")
            redaction_count = 0
            try:
                total = len(doc)
                job_init(job_id, total=total, detail="Scanning & redacting sensitive text...")
                for idx, page in enumerate(doc):
                    job_update(job_id, idx, total, f"Redacting page {idx + 1} of {total}...")
                    page_hits = 0
                    for term in term_list:
                        rects = page.search_for(term)
                        for rect in rects:
                            page.add_redact_annot(rect, fill=rgb)
                            page_hits += 1
                    if page_hits > 0:
                        page.apply_redactions()
                        redaction_count += page_hits
                    job_update(job_id, idx + 1, total, f"Redacted page {idx + 1} of {total} ({redaction_count} matches)")
                out = doc.tobytes(garbage=4, deflate=True)
            finally:
                doc.close()
            job_finish(job_id, "done", f"Redacted {redaction_count} matches across {total} pages")
            return out, total, redaction_count

        try:
            out, total, redaction_count = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_redacted.pdf"',
                    "X-Page-Count": str(total),
                    "X-Redaction-Count": str(redaction_count),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 5. Repair PDF ─────────────────────────────────────────────────────────
    @router.post("/api/pdf/repair")
    async def pdf_repair(
        file: UploadFile = File(...),
        mode: str = Form("standard"),  # "standard" | "deep"
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            job_init(job_id, total=2, detail="Rebuilding PDF xref & object streams...")
            if mode == "deep":
                doc = pymupdf.open(stream=raw, filetype="pdf")
                try:
                    total = len(doc)
                    job_update(job_id, 1, 2, f"Deep-cleaning {total} pages with PyMuPDF...")
                    out = doc.tobytes(garbage=4, deflate=True, clean=True)
                finally:
                    doc.close()
            else:
                try:
                    with pikepdf.open(io.BytesIO(raw)) as pdf:
                        total = len(pdf.pages)
                        job_update(job_id, 1, 2, f"Validating {total} pages with pikepdf...")
                        buf = io.BytesIO()
                        pdf.save(buf, linearize=False, fix_metadata_version=True)
                        out = buf.getvalue()
                except Exception:
                    doc = pymupdf.open(stream=raw, filetype="pdf")
                    try:
                        total = len(doc)
                        job_update(job_id, 1, 2, f"Reconstructing {total} pages with PyMuPDF...")
                        out = doc.tobytes(garbage=4, deflate=True, clean=True)
                    finally:
                        doc.close()
            job_finish(job_id, "done", f"Repaired {total} pages cleanly")
            return out, total

        try:
            out, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_repaired.pdf"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=f"Unable to repair PDF: {exc}")

    # Bounded in-memory LRU cache for rendered WebP page thumbnails (local equivalent of Redis cache)
    _thumb_cache_lock = threading.Lock()
    _thumb_cache: OrderedDict = OrderedDict()
    _THUMB_CACHE_MAX = 16

    # ── 6. Organize PDF (+ Paginated WebP Thumbnails Preview Helper) ──────────
    @router.post("/api/pdf/thumbnails")
    async def pdf_thumbnails(
        file: UploadFile = File(...),
        offset: int = Form(0),
        limit: int = Form(24),
    ):
        raw = await file.read()
        safe_offset = max(0, offset)
        safe_limit = max(1, min(60, limit))
        cache_key = (hashlib.sha256(raw).hexdigest()[:16], safe_offset, safe_limit)

        with _thumb_cache_lock:
            cached = _thumb_cache.get(cache_key)
            if cached is not None:
                _thumb_cache.move_to_end(cache_key)
                return cached

        def _run():
            doc = pymupdf.open(stream=raw, filetype="pdf")
            thumbs = []
            try:
                total = len(doc)
                end = min(total, safe_offset + safe_limit)
                for idx in range(safe_offset, end):
                    pix = doc[idx].get_pixmap(matrix=pymupdf.Matrix(0.32, 0.32), alpha=False)
                    im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    wbuf = io.BytesIO()
                    im.save(wbuf, format="WEBP", quality=78, method=4)
                    b64 = base64.b64encode(wbuf.getvalue()).decode("ascii")
                    thumbs.append({"page": idx + 1, "src": f"data:image/webp;base64,{b64}"})
            finally:
                doc.close()
            return {
                "total_pages": total,
                "offset": safe_offset,
                "limit": safe_limit,
                "has_more": end < total,
                "thumbnails": thumbs,
            }

        try:
            payload = await asyncio.to_thread(_run)
            with _thumb_cache_lock:
                _thumb_cache[cache_key] = payload
                _thumb_cache.move_to_end(cache_key)
                while len(_thumb_cache) > _THUMB_CACHE_MAX:
                    _thumb_cache.popitem(last=False)
            return payload
        except Exception as exc:
            raise HTTPException(status_code=500, detail=str(exc))

    @router.post("/api/pdf/organize")
    async def pdf_organize(
        file: UploadFile = File(...),
        order: str = Form(...),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            reader = PdfReader(io.BytesIO(raw))
            total = len(reader.pages)
            order_clean = order.strip().lower()
            if order_clean == "reverse":
                indices = list(range(total - 1, -1, -1))
            else:
                indices = []
                for part in order.split(","):
                    tok = part.strip()
                    if not tok:
                        continue
                    if "-" in tok:
                        a, b = tok.split("-", 1)
                        ia, ib = int(a.strip()), int(b.strip())
                        step = 1 if ia <= ib else -1
                        for p in range(ia, ib + step, step):
                            if 1 <= p <= total:
                                indices.append(p - 1)
                    else:
                        p = int(tok)
                        if 1 <= p <= total:
                            indices.append(p - 1)
            if not indices:
                raise ValueError(f"Specify a valid page order (1-{total}, e.g. '3, 1, 2' or 'reverse').")

            job_init(job_id, total=len(indices), detail="Reordering pages...")
            writer = PdfWriter()
            for step_i, p_idx in enumerate(indices, start=1):
                job_check_cancel(job_id)
                writer.add_page(reader.pages[p_idx])
                job_update(job_id, step_i, len(indices), f"Added page {p_idx + 1} ({step_i}/{len(indices)})")
            buf = io.BytesIO()
            writer.write(buf)
            job_finish(job_id, "done", f"Organized {len(indices)} pages")
            return buf.getvalue(), len(indices)

        try:
            out, count = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_organized.pdf"',
                    "X-Page-Count": str(count),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=400 if isinstance(exc, ValueError) else 500, detail=str(exc))

    # ── 7. Sign PDF ───────────────────────────────────────────────────────────
    @router.post("/api/pdf/sign")
    async def pdf_sign(
        file: UploadFile = File(...),
        signature_data: str = Form(""),
        signature_file: Optional[UploadFile] = File(None),
        page_num: str = Form("last"),
        position: str = Form("bottom-right"),
        width_pct: float = Form(24.0),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem
        sig_bytes: Optional[bytes] = None
        if signature_file is not None:
            sig_bytes = await signature_file.read()
        elif signature_data.strip():
            payload = signature_data.split(",", 1)[-1]
            sig_bytes = base64.b64decode(payload)

        if not sig_bytes:
            raise HTTPException(status_code=400, detail="Draw a signature or upload a signature PNG image first.")

        def _run():
            img = Image.open(io.BytesIO(sig_bytes)).convert("RGBA")
            bbox = img.getbbox()
            if bbox:
                img = img.crop(bbox)
            img_buf = io.BytesIO()
            img.save(img_buf, format="PNG")
            clean_sig = img_buf.getvalue()
            aspect = img.height / max(1, img.width)

            doc = pymupdf.open(stream=raw, filetype="pdf")
            try:
                total = len(doc)
                job_init(job_id, total=total, detail="Stamping signature...")
                p_clean = page_num.strip().lower()
                if p_clean == "all":
                    targets = list(range(total))
                elif p_clean == "last":
                    targets = [total - 1]
                else:
                    targets = parse_page_ranges(p_clean, total) or [total - 1]

                w_ratio = max(0.08, min(0.65, width_pct / 100.0))
                for idx in targets:
                    job_check_cancel(job_id)
                    page = doc[idx]
                    r = page.rect
                    w = r.width * w_ratio
                    h = w * aspect
                    margin = 36
                    if position == "bottom-left":
                        x0, y0 = margin, r.height - margin - h
                    elif position == "bottom-center":
                        x0, y0 = (r.width - w) / 2.0, r.height - margin - h
                    elif position == "center":
                        x0, y0 = (r.width - w) / 2.0, (r.height - h) / 2.0
                    elif position == "top-right":
                        x0, y0 = r.width - margin - w, margin
                    else:  # bottom-right
                        x0, y0 = r.width - margin - w, r.height - margin - h

                    page.insert_image(pymupdf.Rect(x0, y0, x0 + w, y0 + h), stream=clean_sig, overlay=True)
                out = doc.tobytes(deflate=True)
            finally:
                doc.close()
            job_finish(job_id, "done", f"Signed {len(targets)} page(s)")
            return out, total

        try:
            out, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}_signed.pdf"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 8. Compare PDF (Side-by-Side Diff Report) ─────────────────────────────
    @router.post("/api/pdf/compare")
    async def pdf_compare(
        files: List[UploadFile] = File(...),
        ignore_whitespace: bool = Form(True),
        ignore_case: bool = Form(False),
        job_id: Optional[str] = Form(None),
    ):
        if len(files) < 2:
            raise HTTPException(status_code=400, detail="Select 2 PDF files to compare.")
        raw_a = await files[0].read()
        raw_b = await files[1].read()
        name_a = files[0].filename or "original.pdf"
        name_b = files[1].filename or "modified.pdf"

        def _norm(lines: List[str]) -> List[str]:
            out = []
            for ln in lines:
                s = " ".join(ln.split()) if ignore_whitespace else ln
                if ignore_whitespace and not s:
                    continue
                if ignore_case:
                    s = s.lower()
                out.append(s)
            return out

        def _run():
            doc_a = pymupdf.open(stream=raw_a, filetype="pdf")
            doc_b = pymupdf.open(stream=raw_b, filetype="pdf")
            try:
                max_pages = max(len(doc_a), len(doc_b))
                job_init(job_id, total=max_pages, detail="Comparing PDF documents...")
                lines_a: List[str] = []
                lines_b: List[str] = []
                for i in range(max_pages):
                    job_update(job_id, i, max_pages, f"Reading page {i + 1} of {max_pages}...")
                    ta = _norm(doc_a[i].get_text("text").splitlines()) if i < len(doc_a) else []
                    tb = _norm(doc_b[i].get_text("text").splitlines()) if i < len(doc_b) else []
                    lines_a.append(f"=== Page {i + 1} ===")
                    lines_a.extend(ta)
                    lines_b.append(f"=== Page {i + 1} ===")
                    lines_b.extend(tb)
                    job_update(job_id, i + 1, max_pages, f"Compared page {i + 1} of {max_pages}")
            finally:
                doc_a.close()
                doc_b.close()

            added = 0
            removed = 0
            for d in difflib.unified_diff(lines_a, lines_b, lineterm=""):
                if d.startswith("+") and not d.startswith("+++"):
                    added += 1
                elif d.startswith("-") and not d.startswith("---"):
                    removed += 1

            html_diff = difflib.HtmlDiff(wrapcolumn=80).make_file(
                lines_a, lines_b, fromdesc=name_a, todesc=name_b, context=True, numlines=3
            )
            job_finish(job_id, "done", f"Diff complete (+{added} / -{removed} lines)")
            return html_diff, max_pages, added, removed

        try:
            html_diff, max_pages, added, removed = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(html_diff.encode("utf-8")),
                media_type="text/html; charset=utf-8",
                headers={
                    "Content-Disposition": 'attachment; filename="local-folio_compare_report.html"',
                    "X-Page-Count": str(max_pages),
                    "X-Diff-Added": str(added),
                    "X-Diff-Removed": str(removed),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 9. PDF to Word (.docx) ────────────────────────────────────────────────
    @router.post("/api/pdf/to-word")
    async def pdf_to_word(
        file: UploadFile = File(...),
        pages: str = Form(""),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            from pdf2docx import Converter

            with tempfile.TemporaryDirectory() as tmp_dir:
                in_pdf = Path(tmp_dir) / "input.pdf"
                out_docx = Path(tmp_dir) / f"{stem}.docx"
                in_pdf.write_bytes(raw)
                doc_check = pymupdf.open(str(in_pdf))
                total = len(doc_check)
                doc_check.close()
                selected = parse_page_ranges(pages, total) if pages.strip() else None
                count = len(selected) if selected else total
                job_init(job_id, total=count, detail=f"Converting {count} page(s) to Word (.docx)...")
                cv = Converter(str(in_pdf))
                try:
                    if selected:
                        cv.convert(str(out_docx), pages=selected)
                    else:
                        cv.convert(str(out_docx), start=0, end=None)
                finally:
                    cv.close()
                job_finish(job_id, "done", f"Converted {count} pages to Word")
                return out_docx.read_bytes(), count

        try:
            data, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(data),
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}.docx"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 10. PDF to PowerPoint (.pptx) ─────────────────────────────────────────
    @router.post("/api/pdf/to-pptx")
    async def pdf_to_pptx(
        file: UploadFile = File(...),
        dpi: int = Form(150),
        slide_size: str = Form("match"),  # "match" | "16:9" | "4:3"
        pages: str = Form(""),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem
        scale = max(1.0, min(3.0, dpi / 72.0))

        def _run():
            import pypdfium2 as pdfium

            prs = Presentation()
            blank_layout = prs.slide_layouts[6]
            if slide_size == "16:9":
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
            elif slide_size == "4:3":
                prs.slide_width = Inches(10.0)
                prs.slide_height = Inches(7.5)

            doc = pdfium.PdfDocument(raw)
            try:
                total = len(doc)
                indices = parse_page_ranges(pages, total) if pages.strip() else list(range(total))
                if not indices:
                    raise ValueError(f"No valid pages in range (PDF has {total} pages).")
                job_init(job_id, total=len(indices), detail=f"Creating {len(indices)} PowerPoint slide(s)...")
                for step_i, idx in enumerate(indices, start=1):
                    job_update(job_id, step_i - 1, len(indices), f"Creating slide {step_i} of {len(indices)}...")
                    page = doc[idx]
                    try:
                        img = page.render(scale=scale).to_pil()
                    finally:
                        page.close()
                    if step_i == 1 and slide_size == "match":
                        prs.slide_width = Inches(10.0)
                        prs.slide_height = Inches(10.0 * (img.height / max(1, img.width)))

                    slide = prs.slides.add_slide(blank_layout)
                    img_buf = io.BytesIO()
                    img.convert("RGB").save(img_buf, format="JPEG", quality=90)
                    img_buf.seek(0)

                    # Fit proportionally centered inside slide
                    sw, sh = prs.slide_width, prs.slide_height
                    img_ratio = img.width / max(1, img.height)
                    slide_ratio = sw / max(1, sh)
                    if img_ratio > slide_ratio:
                        pw = sw
                        ph = int(sw / img_ratio)
                        left_pos, top_pos = 0, int((sh - ph) / 2)
                    else:
                        ph = sh
                        pw = int(sh * img_ratio)
                        left_pos, top_pos = int((sw - pw) / 2), 0
                    slide.shapes.add_picture(img_buf, left_pos, top_pos, width=pw, height=ph)
                    job_update(job_id, step_i, len(indices), f"Created slide {step_i} of {len(indices)}")
            finally:
                doc.close()

            out_buf = io.BytesIO()
            prs.save(out_buf)
            job_finish(job_id, "done", f"Created {len(indices)} slides")
            return out_buf.getvalue(), len(indices)

        try:
            data, total_pages = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(data),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}.pptx"',
                    "X-Page-Count": str(total_pages),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=400 if isinstance(exc, ValueError) else 500, detail=str(exc))

    # ── 11. PDF to Excel (.xlsx) ──────────────────────────────────────────────
    @router.post("/api/pdf/to-excel")
    async def pdf_to_excel(
        file: UploadFile = File(...),
        sheet_mode: str = Form("per_page"),  # "per_page" | "single"
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem

        def _run():
            wb = Workbook()
            ws_default = wb.active
            used_default = False
            if sheet_mode == "single":
                ws_default.title = "Extracted_Tables"
                used_default = True
            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                total = len(pdf.pages)
                job_init(job_id, total=total, detail=f"Extracting tables from {total} pages...")
                for idx, page in enumerate(pdf.pages, start=1):
                    job_update(job_id, idx - 1, total, f"Scanning page {idx} of {total} for tables...")
                    tables = page.extract_tables() or []
                    if sheet_mode == "single":
                        ws = ws_default
                        if idx > 1:
                            ws.append([])
                        ws.append([f"--- Page {idx} ---"])
                    else:
                        ws = ws_default if not used_default else wb.create_sheet(title=f"Page_{idx}")
                        if not used_default:
                            ws.title = f"Page_{idx}"
                            used_default = True
                    if tables:
                        for t_idx, table in enumerate(tables):
                            if t_idx > 0:
                                ws.append([])
                            for row in table:
                                ws.append([cell if cell is not None else "" for cell in row])
                    else:
                        txt = (page.extract_text() or "").splitlines()
                        for line in txt:
                            ws.append([line])
                    job_update(job_id, idx, total, f"Extracted page {idx} of {total}")

            out = io.BytesIO()
            wb.save(out)
            job_finish(job_id, "done", f"Exported {total} pages to Excel")
            return out.getvalue(), total

        try:
            data, total = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(data),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={
                    "Content-Disposition": f'attachment; filename="{stem}.xlsx"',
                    "X-Page-Count": str(total),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 12–14. Word / PowerPoint / Excel to PDF ───────────────────────────────
    async def _handle_office_to_pdf(
        file: UploadFile,
        app_type: str,
        job_id: Optional[str],
        quality: str = "standard",
        orientation: str = "auto",
        fit_columns: bool = False,
    ):
        raw = await file.read()
        orig_name = file.filename or f"document.{app_type}"
        stem = Path(orig_name).stem
        ext = Path(orig_name).suffix or (".docx" if app_type == "word" else ".xlsx" if app_type == "excel" else ".pptx")

        def _run():
            job_init(job_id, total=2, detail=f"Converting {orig_name} via local {app_type.title()} engine...")
            with tempfile.TemporaryDirectory() as tmp_dir:
                in_path = Path(tmp_dir) / f"input{ext}"
                out_path = Path(tmp_dir) / f"{stem}.pdf"
                in_path.write_bytes(raw)
                job_update(job_id, 1, 2, f"Rendering {orig_name} to PDF...")
                _convert_office_to_pdf_sync(
                    in_path,
                    out_path,
                    app_type,
                    quality=quality,
                    orientation=orientation,
                    fit_columns=fit_columns,
                )
                pdf_bytes = out_path.read_bytes()
            job_finish(job_id, "done", f"Converted {orig_name} to PDF")
            return pdf_bytes

        try:
            pdf_bytes = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(pdf_bytes),
                media_type="application/pdf",
                headers={"Content-Disposition": f'attachment; filename="{stem}.pdf"'},
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    @router.post("/api/pdf/word-to-pdf")
    async def word_to_pdf(
        file: UploadFile = File(...),
        quality: str = Form("standard"),
        job_id: Optional[str] = Form(None),
    ):
        return await _handle_office_to_pdf(file, "word", job_id, quality=quality)

    @router.post("/api/pdf/pptx-to-pdf")
    async def pptx_to_pdf(
        file: UploadFile = File(...),
        quality: str = Form("standard"),
        job_id: Optional[str] = Form(None),
    ):
        return await _handle_office_to_pdf(file, "powerpoint", job_id, quality=quality)

    @router.post("/api/pdf/excel-to-pdf")
    async def excel_to_pdf(
        file: UploadFile = File(...),
        orientation: str = Form("auto"),
        fit_columns: bool = Form(True),
        quality: str = Form("standard"),
        job_id: Optional[str] = Form(None),
    ):
        return await _handle_office_to_pdf(
            file, "excel", job_id, quality=quality, orientation=orientation, fit_columns=fit_columns
        )

    # ── 15. HTML / URL to PDF (Edge / Chrome Headless) ────────────────────────
    @router.post("/api/pdf/html-to-pdf")
    async def html_to_pdf(
        file: Optional[UploadFile] = File(None),
        url: str = Form(""),
        orientation: str = Form("portrait"),  # "portrait" | "landscape"
        page_size: str = Form("a4"),          # "a4" | "letter"
        margin: str = Form("small"),          # "none" | "small" | "big"
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read() if file is not None else None
        stem = Path(file.filename).stem if (file is not None and file.filename) else "webpage"
        target_url = url.strip()
        if not raw and not target_url:
            raise HTTPException(status_code=400, detail="Upload an .html file or enter a web URL.")

        def _run():
            browser = _find_browser_binary()
            if not browser:
                raise RuntimeError("Microsoft Edge or Google Chrome is required for local HTML to PDF rendering.")
            job_init(job_id, total=2, detail="Rendering HTML to PDF via local headless browser...")
            with tempfile.TemporaryDirectory() as tmp_dir:
                out_pdf = Path(tmp_dir) / f"{stem}.pdf"
                if raw:
                    css_size = "letter" if page_size.lower() == "letter" else "A4"
                    css_orient = "landscape" if orientation == "landscape" else "portrait"
                    css_margin = "0mm" if margin == "none" else ("10mm" if margin == "small" else "25mm")
                    page_style = f"<style>@page {{ size: {css_size} {css_orient}; margin: {css_margin}; }}</style>"
                    html_str = raw.decode("utf-8", errors="replace")
                    if "</head>" in html_str.lower():
                        idx_h = html_str.lower().index("</head>")
                        html_str = html_str[:idx_h] + page_style + html_str[idx_h:]
                    else:
                        html_str = page_style + html_str
                    in_html = Path(tmp_dir) / "input.html"
                    in_html.write_text(html_str, encoding="utf-8")
                    source_arg = in_html.as_uri()
                else:
                    source_arg = target_url if target_url.startswith(("http://", "https://")) else f"https://{target_url}"

                job_update(job_id, 1, 2, "Printing page to PDF...")
                cmd = [
                    browser,
                    "--headless=new",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--no-pdf-header-footer",
                    f"--print-to-pdf={out_pdf}",
                    source_arg,
                ]
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
                if not out_pdf.exists():
                    raise RuntimeError(res.stderr.strip() or "Headless browser failed to generate PDF.")
                pdf_bytes = out_pdf.read_bytes()
            job_finish(job_id, "done", "Generated PDF from HTML")
            return pdf_bytes

        try:
            pdf_bytes = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(pdf_bytes),
                media_type="application/pdf",
                headers={"Content-Disposition": f'attachment; filename="{stem}.pdf"'},
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 16. PDF to PDF/A (Archival Conversion) ────────────────────────────────
    @router.post("/api/pdf/pdf-to-pdfa")
    async def pdf_to_pdfa(
        file: UploadFile = File(...),
        conformance: str = Form("2b"),  # "1b" | "2b" | "3b"
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem
        part = "1" if conformance == "1b" else ("3" if conformance == "3b" else "2")

        def _run():
            job_init(job_id, total=2, detail=f"Converting to archival PDF/A-{part}b...")
            gs_bin = _find_gs_binary()
            if gs_bin:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    in_p = Path(tmp_dir) / "in.pdf"
                    out_p = Path(tmp_dir) / "out.pdf"
                    in_p.write_bytes(raw)
                    job_update(job_id, 1, 2, f"Running Ghostscript PDF/A-{part}b conversion...")
                    res = subprocess.run(
                        [
                            gs_bin,
                            f"-dPDFA={part}",
                            "-dBATCH",
                            "-dNOPAUSE",
                            "-sColorConversionStrategy=RGB",
                            "-sDEVICE=pdfwrite",
                            "-dPDFACompatibilityPolicy=1",
                            f"-sOutputFile={out_p}",
                            str(in_p),
                        ],
                        capture_output=True,
                        text=True,
                        timeout=180,
                    )
                    if res.returncode == 0 and out_p.exists():
                        job_finish(job_id, "done", f"Converted to PDF/A-{part}b via Ghostscript")
                        return out_p.read_bytes()

            doc = pymupdf.open(stream=raw, filetype="pdf")
            try:
                total = len(doc)
                job_update(job_id, 1, 2, f"Normalizing {total} pages & embedding PDF/A-{part}b XMP metadata...")
                doc.bake(annots=True, widgets=True)
                clean_bytes = doc.tobytes(garbage=4, deflate=True, clean=True)
            finally:
                doc.close()

            with pikepdf.open(io.BytesIO(clean_bytes)) as pdf:
                with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
                    meta["pdfaid:part"] = part
                    meta["pdfaid:conformance"] = "B"
                    meta["dc:title"] = stem
                out_buf = io.BytesIO()
                pdf.save(out_buf, min_version="1.4" if part == "1" else "1.7")
                result_bytes = out_buf.getvalue()

            job_finish(job_id, "done", f"Saved archival PDF/A-{part}b ({total} pages)")
            return result_bytes

        try:
            data = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(data),
                media_type="application/pdf",
                headers={"Content-Disposition": f'attachment; filename="{stem}_PDFA.pdf"'},
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 17. AI Summarizer (Token-by-Token Streaming via Settings LLM) ─────────
    @router.post("/api/pdf/ai-summarize")
    async def pdf_ai_summarize(
        file: UploadFile = File(...),
        style: str = Form("executive"),  # "executive" | "bullets" | "study"
        llm_base_url: str = Form(""),
        llm_model: str = Form(""),
        llm_api_key: str = Form(""),
        stream: bool = Form(True),
        job_id: Optional[str] = Form(None),
    ):
        raw = await file.read()
        stem = Path(file.filename or "document").stem
        if not llm_model.strip():
            raise HTTPException(
                status_code=400,
                detail="Configure an LLM Model & Base URL (e.g. Ollama http://localhost:11434/v1) or API Key in Settings first.",
            )

        def _extract_text():
            doc = pymupdf.open(stream=raw, filetype="pdf")
            extracted: List[str] = []
            try:
                total = len(doc)
                job_init(job_id, total=total + 1, detail="Extracting PDF text for AI summary...")
                for idx, page in enumerate(doc):
                    job_update(job_id, idx, total + 1, f"Reading page {idx + 1} of {total}...")
                    txt = (page.get_text("text") or "").strip()
                    if len(txt) < 20:
                        pix = page.get_pixmap(matrix=pymupdf.Matrix(1.5, 1.5), alpha=False)
                        pil_img = Image.open(io.BytesIO(pix.tobytes("png")))
                        txt, _, _, _ = run_rapidocr_on_image(pil_img, rotate_mode="0")
                    if txt:
                        extracted.append(f"[Page {idx + 1}]\n{txt}")
            finally:
                doc.close()

            full_text = "\n\n".join(extracted).strip()
            if not full_text:
                raise ValueError("No readable text found in this PDF.")
            return full_text[:45000], total

        try:
            clipped, total_pages = await asyncio.to_thread(_extract_text)
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=400 if isinstance(exc, ValueError) else 500, detail=str(exc))

        prompts = {
            "executive": "Write a concise executive summary in Markdown with Key Takeaways, Main Arguments, and Actionable Conclusions.",
            "bullets": "Summarize this document as crisp, well-organized Markdown bullet points grouped by topic.",
            "study": "Create structured study notes in Markdown with definitions, key concepts, and review questions.",
        }
        instruction = prompts.get(style, prompts["executive"])

        if stream:
            def _token_stream():
                from openai import OpenAI

                job_update(job_id, total_pages, total_pages + 1, f"Streaming tokens from {llm_model.strip()}...")
                yield f"# AI Summary — {stem}\n\n".encode("utf-8")
                client = OpenAI(
                    api_key=llm_api_key.strip() or "ollama",
                    base_url=llm_base_url.strip() or None,
                    timeout=90.0,
                )
                try:
                    resp = client.chat.completions.create(
                        model=llm_model.strip(),
                        messages=[
                            {"role": "system", "content": "You are a precise document analyst. Respond in clean Markdown."},
                            {"role": "user", "content": f"{instruction}\n\nDocument ({stem}.pdf):\n\n{clipped}"},
                        ],
                        temperature=0.2,
                        stream=True,
                    )
                    for chunk in resp:
                        job_check_cancel(job_id)
                        delta = chunk.choices[0].delta.content if chunk.choices else None
                        if delta:
                            yield delta.encode("utf-8")
                    job_finish(job_id, "done", f"AI summary streamed ({llm_model.strip()})")
                except job_cancelled_exc:
                    job_finish(job_id, "cancelled", "Cancelled by user.")
                except Exception as exc:
                    job_finish(job_id, "error", str(exc))
                    yield f"\n\n> [Stream error: {exc}]".encode("utf-8")

            return StreamingResponse(
                _token_stream(),
                media_type="text/markdown; charset=utf-8",
                headers={"Content-Disposition": f'attachment; filename="{stem}_summary.md"'},
            )

        def _non_stream():
            from openai import OpenAI

            job_update(job_id, total_pages, total_pages + 1, f"Generating summary via {llm_model.strip()}...")
            client = OpenAI(
                api_key=llm_api_key.strip() or "ollama",
                base_url=llm_base_url.strip() or None,
                timeout=90.0,
            )
            resp = client.chat.completions.create(
                model=llm_model.strip(),
                messages=[
                    {"role": "system", "content": "You are a precise document analyst. Respond in clean Markdown."},
                    {"role": "user", "content": f"{instruction}\n\nDocument ({stem}.pdf):\n\n{clipped}"},
                ],
                temperature=0.2,
            )
            summary_md = f"# AI Summary — {stem}\n\n" + (resp.choices[0].message.content or "").strip()
            job_finish(job_id, "done", f"AI summary generated ({llm_model.strip()})")
            return summary_md

        try:
            summary_md = await asyncio.to_thread(_non_stream)
            return StreamingResponse(
                io.BytesIO(summary_md.encode("utf-8")),
                media_type="text/markdown; charset=utf-8",
                headers={"Content-Disposition": f'attachment; filename="{stem}_summary.md"'},
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=500, detail=str(exc))

    # ── 18. Custom Workflow Builder (Multi-Step PDF Pipeline) ─────────────────
    @router.post("/api/pdf/workflow")
    async def pdf_workflow(
        files: List[UploadFile] = File(...),
        config: str = Form("{}"),
        job_id: Optional[str] = Form(None),
    ):
        if not files:
            raise HTTPException(status_code=400, detail="Upload at least 1 PDF file.")
        raw_files = [(f.filename or "doc.pdf", await f.read()) for f in files]
        try:
            cfg = json.loads(config or "{}")
        except Exception:
            cfg = {}

        def _run():
            steps: List[str] = []
            if len(raw_files) > 1:
                steps.append("merge")
            if cfg.get("rotate_angle"):
                steps.append("rotate")
            if cfg.get("watermark_text"):
                steps.append("watermark")
            if cfg.get("add_page_numbers"):
                steps.append("page_numbers")
            if cfg.get("compress"):
                steps.append("compress")
            if cfg.get("password"):
                steps.append("protect")

            if not steps:
                raise ValueError("Enable at least one workflow step (or upload 2+ PDFs to merge).")

            job_init(job_id, total=len(steps), detail="Starting custom PDF workflow...")
            step_idx = 0

            # Step A: Merge or load first PDF
            doc = pymupdf.open(stream=raw_files[0][1], filetype="pdf")
            try:
                if len(raw_files) > 1:
                    step_idx += 1
                    job_update(job_id, step_idx - 1, len(steps), f"Step {step_idx}/{len(steps)}: Merging {len(raw_files)} PDFs...")
                    for _, extra_bytes in raw_files[1:]:
                        job_check_cancel(job_id)
                        extra_doc = pymupdf.open(stream=extra_bytes, filetype="pdf")
                        doc.insert_pdf(extra_doc)
                        extra_doc.close()
                    job_update(job_id, step_idx, len(steps), f"Merged {len(raw_files)} PDFs")

                # Step B: Rotate
                rot = int(cfg.get("rotate_angle") or 0)
                if rot in (90, 180, 270):
                    step_idx += 1
                    job_update(job_id, step_idx - 1, len(steps), f"Step {step_idx}/{len(steps)}: Rotating {rot}°...")
                    for page in doc:
                        job_check_cancel(job_id)
                        page.set_rotation((page.rotation + rot) % 360)
                    job_update(job_id, step_idx, len(steps), f"Rotated {rot}°")

                # Step C: Watermark
                wm_text = str(cfg.get("watermark_text") or "").strip()
                if wm_text:
                    step_idx += 1
                    job_update(job_id, step_idx - 1, len(steps), f"Step {step_idx}/{len(steps)}: Watermarking...")
                    _apply_watermark_doc(doc, text=wm_text, opacity=0.22, angle=45, font_size=42)
                    job_update(job_id, step_idx, len(steps), "Watermark applied")

                # Step D: Page Numbers
                if cfg.get("add_page_numbers"):
                    step_idx += 1
                    job_update(job_id, step_idx - 1, len(steps), f"Step {step_idx}/{len(steps)}: Adding page numbers...")
                    _apply_page_numbers_doc(doc, position="bottom-center", fmt="page_of")
                    job_update(job_id, step_idx, len(steps), "Page numbers added")

                # Step E: Compress & Save
                compress_flag = bool(cfg.get("compress"))
                if compress_flag:
                    step_idx += 1
                    job_update(job_id, step_idx - 1, len(steps), f"Step {step_idx}/{len(steps)}: Compressing...")
                    out_bytes = doc.tobytes(garbage=4, deflate=True, clean=True)
                    job_update(job_id, step_idx, len(steps), "Compressed PDF")
                else:
                    out_bytes = doc.tobytes(deflate=True)
                total_pages = len(doc)
            finally:
                doc.close()

            # Step F: Protect
            pwd = str(cfg.get("password") or "").strip()
            if pwd:
                step_idx += 1
                job_update(job_id, step_idx - 1, len(steps), f"Step {step_idx}/{len(steps)}: Encrypting PDF...")
                reader = PdfReader(io.BytesIO(out_bytes))
                writer = PdfWriter()
                for p in reader.pages:
                    writer.add_page(p)
                writer.encrypt(pwd)
                pbuf = io.BytesIO()
                writer.write(pbuf)
                out_bytes = pbuf.getvalue()
                job_update(job_id, step_idx, len(steps), "Encrypted PDF")

            job_finish(job_id, "done", f"Workflow completed ({len(steps)} steps, {total_pages} pages)")
            return out_bytes, total_pages

        try:
            out_bytes, total_pages = await asyncio.to_thread(_run)
            return StreamingResponse(
                io.BytesIO(out_bytes),
                media_type="application/pdf",
                headers={
                    "Content-Disposition": 'attachment; filename="local-folio_workflow.pdf"',
                    "X-Page-Count": str(total_pages),
                },
            )
        except job_cancelled_exc as exc:
            job_finish(job_id, "cancelled", str(exc))
            raise HTTPException(status_code=499, detail=str(exc))
        except Exception as exc:
            job_finish(job_id, "error", str(exc))
            raise HTTPException(status_code=400 if isinstance(exc, ValueError) else 500, detail=str(exc))

    return router
