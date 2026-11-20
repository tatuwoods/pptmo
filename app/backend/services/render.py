from __future__ import annotations

import os
import shutil
import subprocess
import threading
import uuid
from pathlib import Path

import pymupdf

_CONVERSION_LOCK = threading.Lock()


class RenderError(RuntimeError):
    pass


class RenderBusy(RenderError):
    """Raised when the render queue is saturated; routes answer 429."""


_queue_slots = threading.BoundedSemaphore(max(1, int(os.getenv("RENDER_QUEUE_MAX", "4"))))
_queue_depth_lock = threading.Lock()
_queue_depth = 0


def _max_pages() -> int:
    try:
        return max(1, int(os.getenv("RENDER_MAX_PAGES", "500")))
    except (TypeError, ValueError):
        return 500


def render_pptx(pptx_path: Path, out_dir: Path, scale: float = 2.0) -> int:
    global _queue_depth
    # bound the wait queue: without this, a burst of requests queues behind a
    # 180s conversion and head-of-line blocks every mutation endpoint
    if not _queue_slots.acquire(blocking=False):
        raise RenderBusy("render queue is full")
    with _queue_depth_lock:
        _queue_depth += 1
    try:
        return _render_pptx_locked(pptx_path, out_dir, scale)
    finally:
        with _queue_depth_lock:
            _queue_depth -= 1
        _queue_slots.release()


def _render_pptx_locked(pptx_path: Path, out_dir: Path, scale: float = 2.0) -> int:
    out_dir.mkdir(parents=True, exist_ok=True)
    for item in out_dir.iterdir():
        if item.is_file():
            item.unlink()

    profile = out_dir / f"_lo-profile-{uuid.uuid4().hex}"
    profile.mkdir(parents=True, exist_ok=True)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RenderError("LibreOffice (soffice) not found")

    cmd = [
        soffice,
        f"-env:UserInstallation=file://{profile}",
        "--headless",
        "--norestore",
        "--nologo",
        "--nodefault",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(pptx_path),
    ]
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"

    try:
        with _CONVERSION_LOCK:
            proc = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=180,
                env=env,
            )
    except subprocess.TimeoutExpired as exc:
        shutil.rmtree(profile, ignore_errors=True)
        raise RenderError("LibreOffice conversion timed out") from exc

    pdfs = sorted(out_dir.glob("*.pdf"))
    if proc.returncode != 0 or not pdfs:
        detail = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace")[-1200:]
        shutil.rmtree(profile, ignore_errors=True)
        raise RenderError(f"LibreOffice conversion failed: {detail}")

    pdf_path = pdfs[0]
    thumb_scale = max(0.2, min(0.5, scale * 0.16))
    try:
        doc = pymupdf.open(pdf_path)
        page_count = doc.page_count
        if page_count > _max_pages():
            raise RenderError(
                f"document has {page_count} pages, above the render cap of {_max_pages()}"
            )
        for index, page in enumerate(doc):
            pix = page.get_pixmap(matrix=pymupdf.Matrix(scale, scale), alpha=False)
            pix.save(out_dir / f"slide-{index + 1:03d}.png")
            thumb = page.get_pixmap(
                matrix=pymupdf.Matrix(thumb_scale, thumb_scale), alpha=False
            )
            thumb.save(out_dir / f"thumb-{index + 1:03d}.png")
        doc.close()
    finally:
        try:
            pdf_path.unlink()
        except OSError:
            pass
        shutil.rmtree(profile, ignore_errors=True)

    return page_count
