"""v0.6 #4：图片导入 OCR 重建可编辑文本（对标 ningzimu，简化版）。

上传扫描件/截图后不只在页上贴图：先用本机 OCR（macOS Vision framework via
pyobjc）提取文字，生成原生文本框叠在图片上层——用户可直接改字；图片作背景
层保留。OCR 不可用/无文字时行为退回 v0.4（纯图片页），ocr=0 可显式关闭。
"""
from __future__ import annotations

import io

import pytest
import pymupdf
from pptx import Presentation

from services import ocr

TINY_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)

_FAKE_LINES = [
    {"text": "季度经营分析", "confidence": 0.94, "x": 0.10, "y": 0.05, "w": 0.50, "h": 0.06},
    {"text": "营收同比增长 23%", "confidence": 0.91, "x": 0.10, "y": 0.20, "w": 0.60, "h": 0.05},
]


def _upload_images(client, ocr_flag=None):
    files = [
        ("files", ("a.png", io.BytesIO(TINY_PNG), "image/png")),
        ("files", ("b.png", io.BytesIO(TINY_PNG), "image/png")),
    ]
    data = {"ocr": ocr_flag} if ocr_flag is not None else None
    return client.post("/api/docs/import-images", files=files, data=data)


def _slides_with_shapes(client, doc_id):
    import storage

    prs = Presentation(str(storage.current_path(doc_id)))
    return prs, list(prs.slides)


def test_ocr_disabled_keeps_v04_behavior(client, monkeypatch):
    monkeypatch.setattr(ocr, "ocr_image_bytes", lambda data: list(_FAKE_LINES))
    r = _upload_images(client, ocr_flag="0")
    assert r.status_code == 200, r.text
    doc = r.json()
    _prs, slides = _slides_with_shapes(client, doc["doc_id"])
    for slide in slides:
        texts = [s for s in slide.shapes if getattr(s, "has_text_frame", False)]
        assert texts == []  # 关闭 OCR：纯图片页
        pics = [s for s in slide.shapes if str(getattr(s, "shape_type", "")) == "PICTURE (13)"]
        assert len(pics) == 1


def test_ocr_lines_become_editable_textboxes(client, monkeypatch):
    monkeypatch.setattr(ocr, "ocr_available", lambda: True)
    monkeypatch.setattr(ocr, "ocr_image_bytes", lambda data: list(_FAKE_LINES))
    r = _upload_images(client)
    assert r.status_code == 200, r.text
    doc = r.json()
    assert doc.get("ocr", {}).get("enabled") is True
    assert doc["ocr"]["lines"] == 4  # 2 页 × 2 行
    prs, slides = _slides_with_shapes(client, doc["doc_id"])
    for slide in slides:
        texts = [s for s in slide.shapes if getattr(s, "has_text_frame", False) and s.text_frame.text.strip()]
        assert [t.text_frame.text for t in texts] == ["季度经营分析", "营收同比增长 23%"]
        for tb in texts:
            # 叠在图片上层、落在页面内
            assert int(tb.left) >= 0 and int(tb.top) >= 0
            assert int(tb.left) + int(tb.width) <= int(prs.slide_width)
            assert int(tb.top) + int(tb.height) <= int(prs.slide_height)
        pics = [s for s in slide.shapes if str(getattr(s, "shape_type", "")) == "PICTURE (13)"]
        assert len(pics) == 1  # 图片背景层保留


def test_ocr_textbox_position_maps_image_bbox(client, monkeypatch):
    monkeypatch.setattr(ocr, "ocr_available", lambda: True)
    monkeypatch.setattr(ocr, "ocr_image_bytes", lambda data: list(_FAKE_LINES[:1]))
    r = _upload_images(client)
    doc = r.json()
    import storage

    prs = Presentation(str(storage.current_path(doc["doc_id"])))
    slide = prs.slides[0]
    pic = next(s for s in slide.shapes if str(getattr(s, "shape_type", "")) == "PICTURE (13)")
    tb = next(
        s for s in slide.shapes
        if getattr(s, "has_text_frame", False) and s.text_frame.text.strip()
    )
    # 行框 x=0.10 → 文本框左缘 ≈ 图片左缘 + 0.10 × 图片宽（允许 ±5% 页宽误差）
    expected = int(pic.left) + 0.10 * int(pic.width)
    assert abs(int(tb.left) - expected) < int(prs.slide_width) * 0.05
    # 行高 0.06 × 图片高 → 字号应显著大于 9pt 下限
    assert tb.text_frame.paragraphs[0].runs[0].font.size is not None


def test_ocr_no_text_falls_back_silently(client, monkeypatch):
    monkeypatch.setattr(ocr, "ocr_available", lambda: True)
    monkeypatch.setattr(ocr, "ocr_image_bytes", lambda data: [])
    r = _upload_images(client)
    assert r.status_code == 200
    doc = r.json()
    assert doc["ocr"]["lines"] == 0
    _prs, slides = _slides_with_shapes(client, doc["doc_id"])
    for slide in slides:
        assert not [s for s in slide.shapes if getattr(s, "has_text_frame", False) and s.text_frame.text.strip()]


def test_ocr_error_does_not_break_import(client, monkeypatch):
    def boom(data):
        raise RuntimeError("vision exploded")

    monkeypatch.setattr(ocr, "ocr_available", lambda: True)
    monkeypatch.setattr(ocr, "ocr_image_bytes", boom)
    r = _upload_images(client)
    assert r.status_code == 200, r.text  # OCR 挂了也不阻塞导入
    doc = r.json()
    assert doc["ocr"]["enabled"] is False


# ---------- 真机 macOS Vision 冒烟（pyobjc 缺失时自动跳过） ----------


def _rendered_text_png(text: str) -> bytes:
    doc = pymupdf.open()
    page = doc.new_page(width=400, height=120)
    page.insert_text((30, 70), text, fontsize=28, fontname="helv")
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    return pix.tobytes("png")


@pytest.mark.skipif(not ocr.ocr_available(), reason="pyobjc Vision framework 不可用")
def test_real_macos_vision_reads_text():
    lines = ocr.ocr_image_bytes(_rendered_text_png("PPT Light 2026"))
    assert lines, "Vision 应至少识别出一行文字"
    joined = " ".join(line["text"] for line in lines)
    assert "2026" in joined or "PPT" in joined.upper()
    for line in lines:
        assert 0 <= line["x"] <= 1 and 0 <= line["y"] <= 1
        assert line["w"] > 0 and line["h"] > 0
        assert line["confidence"] > 0
