"""v0.5 #5/#10：hide_slide/show_slide 隐藏页 + 导出联动（默认跳过隐藏页）。

对标 hands-on-deck hide/unhide：`show="0"` 挂在 presentation.xml 的 p:sldId 上，
播放/导出自动跳过；导出 PDF/PNG 默认剔除隐藏页，?include_hidden=1 可含。
"""
from __future__ import annotations

import io
import zipfile

import pytest
import pymupdf
from pptx import Presentation

from services import edits
from services import parse


def _hidden_flags(prs) -> list:
    return [el.get("show") == "0" for el in prs.slides._sldIdLst]


class TestHideShowUnit:
    def test_hide_sets_show_attr(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "hide_slide", "payload": {"pages": "2-3"}}],
        )
        assert touched == {1, 2}
        prs = Presentation(str(dst))
        assert _hidden_flags(prs) == [False, True, True]

    def test_show_removes_attr(self, demo_pptx, tmp_path):
        hidden = tmp_path / "hidden.pptx"
        edits.apply_patches_to_file(
            demo_pptx, hidden, [{"action": "hide_slide", "payload": {"pages": "2"}}]
        )
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            hidden, dst, [{"action": "show_slide", "payload": {"pages": "2"}}]
        )
        assert touched == {1}
        prs = Presentation(str(dst))
        assert _hidden_flags(prs) == [False, False, False]
        # show 属性缺省即可见，不残留 show="1"
        assert prs.slides._sldIdLst[1].get("show") is None

    def test_hide_defaults_to_all_pages(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            demo_pptx, dst, [{"action": "hide_slide", "payload": {}}]
        )
        assert touched == {0, 1, 2}
        assert _hidden_flags(Presentation(str(dst))) == [True, True, True]

    def test_hide_idempotent(self, demo_pptx, tmp_path):
        hidden = tmp_path / "hidden.pptx"
        edits.apply_patches_to_file(
            demo_pptx, hidden, [{"action": "hide_slide", "payload": {"pages": "1"}}]
        )
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            hidden, dst, [{"action": "hide_slide", "payload": {"pages": "1"}}]
        )
        assert touched == {0}
        assert _hidden_flags(Presentation(str(dst)))[0] is True

    def test_invalid_pages_raises(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, dst, [{"action": "hide_slide", "payload": {"pages": "9"}}]
            )

    def test_aliases(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst, [{"action": "hide", "payload": {"pages": "1"}}]
        )
        assert _hidden_flags(Presentation(str(dst)))[0] is True
        shown = tmp_path / "shown.pptx"
        edits.apply_patches_to_file(
            dst, shown, [{"action": "unhide", "payload": {"pages": "1"}}]
        )
        assert _hidden_flags(Presentation(str(shown)))[0] is False


class TestParseHidden:
    def test_parse_marks_hidden_slides(self, demo_pptx, tmp_path):
        hidden = tmp_path / "hidden.pptx"
        edits.apply_patches_to_file(
            demo_pptx, hidden, [{"action": "hide_slide", "payload": {"pages": "2"}}]
        )
        _objects, payload = parse.parse_pptx(hidden)
        flags = [s.get("hidden") for s in payload["slides"]]
        assert flags == [False, True, False]


def _hide_page2(client, doc_id) -> dict:
    resp = client.post(
        f"/api/docs/{doc_id}/apply",
        json={
            "label": "隐藏草稿页",
            "patches": [{"action": "hide_slide", "payload": {"pages": "2"}}],
        },
    )
    assert resp.status_code == 200, resp.text
    return resp.json()


class TestHideApi:
    def test_apply_hide_marks_slides_and_undo_restores(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        body = _hide_page2(client, doc_id)
        assert body["slide_count"] == 3  # 隐藏不改页数
        assert [s.get("hidden") for s in body["slides"]] == [False, True, False]
        assert body["touched_slides"] == [1]
        undo = client.post(f"/api/docs/{doc_id}/undo").json()
        assert [s.get("hidden") for s in undo["slides"]] == [False, False, False]

    def test_outline_marks_hidden_page(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        _hide_page2(client, doc_id)
        text = client.get(f"/api/docs/{doc_id}/export/outline").text
        assert "第 2 页（已隐藏）" in text
        assert "第 1 页（已隐藏）" not in text


def _plant_real_pdf(client, doc_id, pages: int = 3) -> None:
    """把当前渲染目录里的 stub export.pdf 换成真实 N 页 PDF（模拟 LO 产出）。"""
    import storage

    doc_dir = storage.doc_dir(doc_id)
    pdfs = sorted(doc_dir.glob("render/*/export.pdf"), key=lambda p: p.stat().st_mtime)
    assert pdfs, "render dir not found"
    doc = pymupdf.open()
    for _ in range(pages):
        doc.new_page(width=960, height=540)
    doc.save(str(pdfs[-1]), garbage=3, deflate=True)
    doc.close()


class TestExportSkipsHidden:
    def test_export_pdf_skips_hidden_by_default(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        _hide_page2(client, doc_id)
        _plant_real_pdf(client, doc_id, pages=3)
        data = client.get(f"/api/docs/{doc_id}/export/pdf").content
        assert pymupdf.open(stream=data, filetype="pdf").page_count == 2

    def test_export_pdf_include_hidden_param(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        _hide_page2(client, doc_id)
        _plant_real_pdf(client, doc_id, pages=3)
        data = client.get(f"/api/docs/{doc_id}/export/pdf?include_hidden=1").content
        assert pymupdf.open(stream=data, filetype="pdf").page_count == 3

    def test_export_zip_skips_hidden_by_default(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        _hide_page2(client, doc_id)
        resp = client.get(f"/api/docs/{doc_id}/export/images.zip")
        assert resp.status_code == 200, resp.text
        names = zipfile.ZipFile(io.BytesIO(resp.content)).namelist()
        assert names == ["slide-001.png", "slide-003.png"]

    def test_export_zip_include_hidden_param(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        _hide_page2(client, doc_id)
        resp = client.get(f"/api/docs/{doc_id}/export/images.zip?include_hidden=1")
        names = zipfile.ZipFile(io.BytesIO(resp.content)).namelist()
        assert len(names) == 3

    def test_export_without_hidden_serves_all(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.get(f"/api/docs/{doc_id}/export/images.zip")
        names = zipfile.ZipFile(io.BytesIO(resp.content)).namelist()
        assert len(names) == 3
