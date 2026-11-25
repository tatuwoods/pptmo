"""v0.5 #3：merge 从另一 deck 拉页（跨包克隆 + POST /api/docs/{id}/merge）。

对标 hands-on-deck merge：上传另一 PPTX，指定页码列表，把它的页拉进当前文档
（slide part + 版式/母版/图片等 rels 全量重新家园化）。
"""
from __future__ import annotations

import io

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from services import edits

PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)


def _src_deck(tmp_path):
    """3 页源：带标题/图片/备注。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s0 = prs.slides.add_slide(blank)
    tb = s0.shapes.add_textbox(Emu(500000), Emu(500000), Emu(5000000), Emu(800000))
    tb.text_frame.text = "源第一页"
    s0.notes_slide.notes_text_frame.text = "源页备注"
    s1 = prs.slides.add_slide(blank)
    s1.shapes.add_picture(io.BytesIO(PNG), Emu(500000), Emu(500000), Emu(1500000), Emu(1000000))
    tb1 = s1.shapes.add_textbox(Emu(500000), Emu(2000000), Emu(5000000), Emu(800000))
    tb1.text_frame.text = "源第二页（图）"
    s2 = prs.slides.add_slide(blank)
    tb2 = s2.shapes.add_textbox(Emu(500000), Emu(500000), Emu(5000000), Emu(800000))
    tb2.text_frame.text = "源第三页"
    path = tmp_path / "src.pptx"
    prs.save(str(path))
    return path


def _dst_deck(tmp_path):
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for text in ("目标A", "目标B"):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(5000000), Emu(800000))
        tb.text_frame.text = text
    path = tmp_path / "dst.pptx"
    prs.save(str(path))
    return path


def _slide_texts(prs):
    texts = []
    for slide in prs.slides:
        parts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                parts.append(shape.text.strip())
        texts.append(" | ".join(parts))
    return texts


class TestMergeUnit:
    def test_merge_pages_at_end(self, tmp_path):
        dst, src = _dst_deck(tmp_path), _src_deck(tmp_path)
        out = tmp_path / "out.pptx"
        merged, touched = edits.merge_slides_from_file(
            dst, src, out, pages="2-3", position=None
        )
        assert merged == 2
        prs = Presentation(str(out))
        assert len(prs.slides) == 4
        assert _slide_texts(prs) == ["目标A", "目标B", "源第二页（图）", "源第三页"]
        # 图片完整搬过来（rels 重新家园化）
        pic_slide = prs.slides[2]
        pics = [s for s in pic_slide.shapes if str(s.shape_type).startswith("PICTURE")]
        assert len(pics) == 1
        assert pics[0].image.blob == PNG
        # 版式也复制了：合并页指向 dst 包内自己的 layout part
        layout_part = pic_slide.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
        )
        assert layout_part in [p for p in prs.part.package.iter_parts()]

    def test_merge_at_position(self, tmp_path):
        dst, src = _dst_deck(tmp_path), _src_deck(tmp_path)
        out = tmp_path / "out.pptx"
        merged, _ = edits.merge_slides_from_file(dst, src, out, pages="1", position=1)
        assert merged == 1
        prs = Presentation(str(out))
        assert _slide_texts(prs) == ["目标A", "源第一页", "目标B"]
        # 备注随页搬迁
        assert prs.slides[1].notes_slide.notes_text_frame.text == "源页备注"

    def test_merge_multiple_pages_at_position(self, tmp_path):
        # 多页插入中位：挪页时插入点在删除点之前，待挪页会后移一位
        dst, src = _dst_deck(tmp_path), _src_deck(tmp_path)
        out = tmp_path / "out.pptx"
        merged, _ = edits.merge_slides_from_file(dst, src, out, pages="1-2", position=1)
        assert merged == 2
        prs = Presentation(str(out))
        assert _slide_texts(prs) == [
            "目标A", "源第一页", "源第二页（图）", "目标B",
        ]

    def test_merge_keeps_slide_size_of_target(self, tmp_path):
        dst, src = _dst_deck(tmp_path), _src_deck(tmp_path)
        out = tmp_path / "out.pptx"
        edits.merge_slides_from_file(dst, src, out, pages="1", position=None)
        prs = Presentation(str(out))
        assert int(prs.slide_width) == int(Presentation(str(dst)).slide_width)

    def test_merge_invalid_pages(self, tmp_path):
        dst, src = _dst_deck(tmp_path), _src_deck(tmp_path)
        out = tmp_path / "out.pptx"
        with pytest.raises(edits.PatchError):
            edits.merge_slides_from_file(dst, src, out, pages="9", position=None)
        with pytest.raises(edits.PatchError):
            edits.merge_slides_from_file(dst, src, out, pages="abc", position=None)
        with pytest.raises(edits.PatchError):
            edits.merge_slides_from_file(dst, src, out, pages="1", position=99)

    def test_merge_self_package(self, tmp_path):
        # 同一个文件作为源：自合并（复制自己的页）
        src = _src_deck(tmp_path)
        out = tmp_path / "out.pptx"
        merged, _ = edits.merge_slides_from_file(src, src, out, pages="1", position=0)
        prs = Presentation(str(out))
        assert merged == 1 and len(prs.slides) == 4
        assert _slide_texts(prs)[0] == "源第一页"


class TestMergeApi:
    def test_merge_endpoint(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        src_path = _build_source_bytes()
        resp = client.post(
            f"/api/docs/{doc_id}/merge",
            files={"file": ("source.pptx", src_path, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            data={"pages": "2-3", "position": "1"},
        )
        assert resp.status_code == 200, resp.text
        body = resp.json()
        assert body["slide_count"] == uploaded_doc["slide_count"] + 2
        assert body["version"] == uploaded_doc["version"] + 1
        # 下载验证页序
        data = client.get(f"/api/docs/{doc_id}/download").content
        prs = Presentation(io.BytesIO(data))
        texts = _slide_texts(prs)
        assert texts[1] == "源第二页（图）" and texts[2] == "源第三页"
        # 撤销回滚合并
        undo = client.post(f"/api/docs/{doc_id}/undo")
        assert undo.status_code == 200
        assert undo.json()["slide_count"] == uploaded_doc["slide_count"]

    def test_merge_rejects_non_pptx(self, client, uploaded_doc):
        resp = client.post(
            f"/api/docs/{uploaded_doc['doc_id']}/merge",
            files={"file": ("evil.bin", b"not a pptx", "application/octet-stream")},
            data={"pages": "1"},
        )
        assert resp.status_code == 400

    def test_merge_rejects_bad_pages(self, client, uploaded_doc):
        resp = client.post(
            f"/api/docs/{uploaded_doc['doc_id']}/merge",
            files={"file": ("source.pptx", _build_source_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            data={"pages": "99"},
        )
        assert resp.status_code == 400


def _build_source_bytes() -> bytes:
    buf = io.BytesIO()
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for text in ("源第一页", "源第二页（图）", "源第三页"):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(5000000), Emu(800000))
        tb.text_frame.text = text
    prs.save(buf)
    return buf.getvalue()
