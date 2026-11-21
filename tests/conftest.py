"""Shared fixtures: sys.path setup, generated pptx, TestClient with stubbed rendering."""
from __future__ import annotations

import io
import sys
from pathlib import Path

import pytest

BACKEND_DIR = Path(__file__).resolve().parent.parent / "app" / "backend"
sys.path.insert(0, str(BACKEND_DIR))


def _tiny_png() -> bytes:
    return bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
    )


@pytest.fixture()
def demo_pptx(tmp_path: Path) -> Path:
    """3-slide deck with title/body text, a table with a merged cell, a grouped textbox."""
    from pptx import Presentation
    from pptx.util import Emu, Pt

    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    blank = prs.slide_layouts[6]

    # slide 0: title + body + picture + speaker notes
    s0 = prs.slides.add_slide(blank)
    tb = s0.shapes.add_textbox(Emu(1000000), Emu(500000), Emu(6000000), Emu(1000000))
    tb.text_frame.text = "项目进展汇报"
    for run in tb.text_frame.paragraphs[0].runs:
        run.font.size = Pt(32)
        run.font.bold = True
    body = s0.shapes.add_textbox(Emu(1000000), Emu(2000000), Emu(8000000), Emu(2000000))
    body.text_frame.text = "本季度完成了三个关键里程碑。"
    s0.shapes.add_picture(io.BytesIO(_tiny_png()), Emu(9500000), Emu(500000), Emu(2000000), Emu(1500000))
    s0.notes_slide.notes_text_frame.text = "讲稿：先回顾里程碑。"

    # slide 1: table with merged cells
    s1 = prs.slides.add_slide(blank)
    gfx = s1.shapes.add_table(3, 3, Emu(1000000), Emu(1000000), Emu(8000000), Emu(3000000))
    table = gfx.table
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "金额"
    table.cell(0, 2).text = "状态"
    table.cell(1, 0).text = "基础设施建设"
    table.cell(1, 1).text = "200 万"
    table.cell(2, 0).text = "软件采购"
    table.cell(2, 1).text = "80 万"
    table.cell(2, 2).text = "进行中"
    table.cell(1, 1).merge(table.cell(1, 2))  # merge (1,1)->(1,2)

    # slide 2: group containing a textbox
    s2 = prs.slides.add_slide(blank)
    sub1 = s2.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(2000000), Emu(500000))
    sub1.text_frame.text = "目录"
    grp = s2.shapes.add_group_shape()
    inner = grp.shapes.add_textbox(Emu(3000000), Emu(1000000), Emu(3000000), Emu(500000))
    inner.text_frame.text = "第一章 总体情况"

    path = tmp_path / "demo.pptx"
    prs.save(str(path))
    return path


@pytest.fixture()
def client(monkeypatch, demo_pptx, tmp_path):
    """TestClient with rendering stubbed (no LibreOffice) and isolated data dir."""
    import storage
    from services import render as render_mod
    from fastapi.testclient import TestClient
    import main as app_module

    data_dir = tmp_path / "data"
    monkeypatch.setattr(storage, "DOCS_DIR", data_dir / "docs")

    def fake_render(pptx_path, out_dir, scale=2.0):
        from pptx import Presentation

        out_dir.mkdir(parents=True, exist_ok=True)
        count = len(Presentation(str(pptx_path)).slides)
        for i in range(count):
            (out_dir / f"slide-{i + 1:03d}.png").write_bytes(_tiny_png())
            (out_dir / f"thumb-{i + 1:03d}.png").write_bytes(_tiny_png())
        return count

    monkeypatch.setattr(render_mod, "render_pptx", fake_render)
    monkeypatch.setattr(app_module, "DATA_DIR", data_dir, raising=False)

    # env: no AI key so chat validation-error paths are exercised deterministically
    monkeypatch.setenv("DEEPSEEK_API_KEY", "")
    monkeypatch.delenv("DEEPSEEK_API_BASE", raising=False)

    return TestClient(app_module.app)


@pytest.fixture()
def uploaded_doc(client, demo_pptx):
    with open(demo_pptx, "rb") as fh:
        resp = client.post(
            "/api/docs/upload",
            files={"file": ("demo.pptx", fh.read(), "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        )
    assert resp.status_code == 200, resp.text
    return resp.json()
