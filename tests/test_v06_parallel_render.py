"""v0.6 #9：逐页并发渲染 —— PDF 光栅化按页多进程并行。

LibreOffice 单命令已整档转 PDF，瓶颈在 PyMuPDF 逐页光栅化；多页文档把页
按连续段分给多个进程渲染（每进程独立开 PDF），单页/受限环境/进程池失败
自动回退串行。产出契约不变：slide-NNN.png + thumb-NNN.png 全页成对。
"""
from __future__ import annotations

import shutil
import threading
import time
from pathlib import Path

import pymupdf

from services import render


def _multi_page_pdf(path: Path, pages: int) -> Path:
    doc = pymupdf.open()
    for index in range(pages):
        page = doc.new_page()  # A4 595x842
        page.insert_text((72, 120), f"Page {index + 1} marker", fontsize=24)
    doc.save(str(path))
    doc.close()
    return path


def _names(out_dir: Path, prefix: str):
    return sorted(p.name for p in out_dir.glob(f"{prefix}-*.png"))


# ---------- 分页与配置 ----------


class TestChunking:
    def test_chunks_cover_all_pages_exactly_once(self):
        chunks = render._page_chunks(9, 4)
        flat = [n for chunk in chunks for n in chunk]
        assert flat == list(range(1, 10))
        for chunk in chunks:  # 每段连续（进程内按页顺扫，不随机打散）
            assert list(chunk) == list(range(chunk[0], chunk[-1] + 1))

    def test_chunks_never_exceed_pages(self):
        assert [list(c) for c in render._page_chunks(2, 8)] == [[1], [2]]
        assert [list(c) for c in render._page_chunks(1, 4)] == [[1]]


class TestWorkerConfig:
    def test_default_bounded_by_cpu(self, monkeypatch):
        monkeypatch.delenv("PPTMO_RASTER_WORKERS", raising=False)
        assert 1 <= render.raster_workers() <= 4

    def test_env_override(self, monkeypatch):
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "3")
        assert render.raster_workers() == 3

    def test_invalid_or_zero_falls_back_serial(self, monkeypatch):
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "0")
        assert render.raster_workers() == 1
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "junk")
        assert render.raster_workers() == 1


# ---------- 光栅化产出契约 ----------


class TestRasterizePdf:
    def test_serial_writes_all_pages_and_thumbs(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "1")
        pdf = _multi_page_pdf(tmp_path / "doc.pdf", 5)
        out = tmp_path / "out"
        out.mkdir()
        count = render.rasterize_pdf(pdf, out, scale=2.0)
        assert count == 5
        assert _names(out, "slide") == [f"slide-{i:03d}.png" for i in range(1, 6)]
        assert _names(out, "thumb") == [f"thumb-{i:03d}.png" for i in range(1, 6)]
        slide = pymupdf.Pixmap(str(out / "slide-003.png"))
        assert (slide.width, slide.height) == (1190, 1684)  # A4 × 2.0

    def test_parallel_matches_serial_file_set(self, tmp_path, monkeypatch):
        pdf = _multi_page_pdf(tmp_path / "doc.pdf", 6)
        serial_out, parallel_out = tmp_path / "ser", tmp_path / "par"
        serial_out.mkdir()
        parallel_out.mkdir()
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "1")
        render.rasterize_pdf(pdf, serial_out, scale=1.5)
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "4")
        count = render.rasterize_pdf(pdf, parallel_out, scale=1.5)
        assert count == 6
        assert _names(parallel_out, "slide") == _names(serial_out, "slide")
        assert _names(parallel_out, "thumb") == _names(serial_out, "thumb")

    def test_single_page_stays_serial(self, tmp_path, monkeypatch):
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "4")
        pdf = _multi_page_pdf(tmp_path / "one.pdf", 1)
        out = tmp_path / "out"
        out.mkdir()
        assert render.rasterize_pdf(pdf, out, scale=1.0) == 1
        assert (out / "slide-001.png").exists() and (out / "thumb-001.png").exists()


# ---------- 渲染主链路接线 ----------


class TestRenderPipeline:
    def test_render_pptx_rasterizes_parallel_after_worker(self, tmp_path, monkeypatch):
        """假 worker 产出 PDF → 并发光栅化接管，全页成对产出 + export.pdf 落盘。"""
        monkeypatch.setenv("PPTMO_RASTER_WORKERS", "3")
        monkeypatch.setenv("PPTMO_RENDER_QUEUE", str(tmp_path / "queue"))
        pdf_src = _multi_page_pdf(tmp_path / "fake-worker-out.pdf", 4)
        out = tmp_path / "render"

        def fake_worker():
            queue = tmp_path / "queue"
            deadline = time.time() + 15
            while time.time() < deadline:
                tasks = sorted(queue.glob("*.task"))
                if tasks:
                    time.sleep(0.2)  # 模拟 LibreOffice 转换耗时
                    shutil.copy2(pdf_src, out / "converted.pdf")
                    tasks[0].unlink(missing_ok=True)
                    return
                time.sleep(0.05)

        thread = threading.Thread(target=fake_worker, daemon=True)
        thread.start()
        count = render.render_pptx(tmp_path / "deck.pptx", out, scale=2.0)
        thread.join(timeout=1)
        assert count == 4
        assert _names(out, "slide") == [f"slide-{i:03d}.png" for i in range(1, 5)]
        assert _names(out, "thumb") == [f"thumb-{i:03d}.png" for i in range(1, 5)]
        assert (out / "export.pdf").exists()
