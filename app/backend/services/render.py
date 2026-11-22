from __future__ import annotations

import os
import shlex
import shutil
import subprocess  # noqa: F401  保留给非队列路径
import threading
import time
import uuid
from pathlib import Path
from typing import Optional

import pymupdf

_CONVERSION_LOCK = threading.Lock()


class RenderError(RuntimeError):
    pass


class RenderBusy(RenderError):
    """Raised when the render queue is saturated; routes answer 429."""


# v0.3: every render dir keeps one PDF under this name for /export/pdf
EXPORT_PDF_NAME = "export.pdf"


def keep_pdf(out_dir: Path, pdf_path: Path) -> Path:
    """Move the freshly rendered PDF into out_dir under the export name.

    Renamed (not left as the soffice basename) so the wait-loop's *.pdf glob
    only ever sees fresh worker output; out_dir files are cleared at the start
    of each render, so stale copies can't accumulate."""
    target = out_dir / EXPORT_PDF_NAME
    os.replace(pdf_path, target)
    return target


_queue_slots = threading.BoundedSemaphore(max(1, int(os.getenv("RENDER_QUEUE_MAX", "4"))))
_queue_depth_lock = threading.Lock()
_queue_depth = 0


def _max_pages() -> int:
    try:
        return max(1, int(os.getenv("RENDER_MAX_PAGES", "500")))
    except (TypeError, ValueError):
        return 500


def _timeout_s() -> int:
    """Worker wait timeout; env-tunable so tests/ops can shorten it."""
    try:
        return max(1, int(os.getenv("RENDER_TIMEOUT_S", "180")))
    except (TypeError, ValueError):
        return 180


def partial_marker(out_dir: Path) -> Path:
    """Marker written when a render degraded to stale pages (v0.3 E21)."""
    return Path(out_dir) / "_partial"


def _degrade_to_stale(out_dir: Path, stale_dir: Optional[Path]) -> Optional[int]:
    """Timeout fallback (E21): reuse the previous version's page images.

    Returns the page count restored, or None when no stale pages exist (the
    caller then keeps the hard-fail behaviour)."""
    if not stale_dir:
        return None
    pages = sorted(Path(stale_dir).glob("slide-*.png"))
    if not pages:
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    for page in pages:
        shutil.copy2(page, out_dir / page.name)
        thumb = page.with_name(page.name.replace("slide-", "thumb-", 1))
        if thumb.exists():
            shutil.copy2(thumb, out_dir / thumb.name)
    partial_marker(out_dir).write_text("stale", encoding="utf-8")
    return len(pages)


def render_pptx(
    pptx_path: Path, out_dir: Path, scale: float = 2.0, stale_dir: Optional[Path] = None
) -> int:
    global _queue_depth
    # bound the wait queue: without this, a burst of requests queues behind a
    # 180s conversion and head-of-line blocks every mutation endpoint
    if not _queue_slots.acquire(blocking=False):
        raise RenderBusy("render queue is full")
    with _queue_depth_lock:
        _queue_depth += 1
    try:
        return _render_pptx_locked(pptx_path, out_dir, scale, stale_dir=stale_dir)
    finally:
        with _queue_depth_lock:
            _queue_depth -= 1
        _queue_slots.release()


def _render_pptx_locked(
    pptx_path: Path, out_dir: Path, scale: float = 2.0, stale_dir: Optional[Path] = None
) -> int:
    out_dir.mkdir(parents=True, exist_ok=True)
    for item in out_dir.iterdir():
        if item.is_file():
            item.unlink()

    profile = out_dir / f"_lo-profile-{uuid.uuid4().hex}"
    profile.mkdir(parents=True, exist_ok=True)

    # ── 队列模式（v0.2.2）：转换交给 launchd 常驻 worker（Aqua 会话直接子进程）执行。
    # uvicorn(launchd) 直接 spawn 的 soffice 在 launchd 会话链里会丢全部 CJK 字形
    # （只剩 ASCII），而 launchd 直接派生的 bash worker 渲染正常（实测对照）。
    queue_dir = Path(os.getenv(
        "PPTMO_RENDER_QUEUE",
        "/Users/tt06/.agents/skills/pptmo/app/backend/data/render-queue",
    ))
    queue_dir.mkdir(parents=True, exist_ok=True)
    task_id = f"{uuid.uuid4().hex}"
    task_file = queue_dir / f"{task_id}.task"
    done_file = task_file.with_suffix(".task.done")
    task_file.write_text(f"{pptx_path}\t{out_dir}\t{scale}\n")

    # 等待 worker 完成（task 文件被认领移除 + done 文件出现；PDF 就绪即成功）
    deadline = time.monotonic() + _timeout_s()
    worker_failed = None
    while time.monotonic() < deadline:
        if done_file.exists():
            marker = done_file.read_text().strip()
            if marker.startswith("fail"):
                worker_failed = marker
            break
        if not task_file.exists() and any(out_dir.glob("*.pdf")):
            break  # 已被认领且产出就绪
        time.sleep(1)
    if worker_failed:
        raise RenderError(f"render worker failed: {worker_failed}")
    if time.monotonic() >= deadline and not any(out_dir.glob("*.pdf")):
        task_file.unlink(missing_ok=True)
        # E21 超时降级：上一版页图仍可用时不整单失败，返回已生成的页并标记部分渲染
        degraded = _degrade_to_stale(out_dir, stale_dir)
        if degraded is not None:
            return degraded
        raise RenderError("render worker timed out")

    pdfs = sorted(out_dir.glob("*.pdf"))
    if not pdfs:
        raise RenderError("render worker produced no PDF")

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
            keep_pdf(out_dir, pdf_path)
        except OSError:
            pass
        shutil.rmtree(profile, ignore_errors=True)

    return page_count
