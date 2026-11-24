"""v0.4 B 组：导入导出 — PDF 备注版 / TXT 大纲 / 图片导入转 PPT / 版本对比。"""
from __future__ import annotations

import io
import zipfile

import pytest

TINY_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)


# ---------- 9. /export/pdf-notes ----------

def test_compose_notes_pdf_layout(client, uploaded_doc, tmp_path):
    """幻灯片图 + 备注文字合成讲义 PDF：每页一图一备注。"""
    import pymupdf

    import main as main_mod
    import storage

    doc_id = uploaded_doc["doc_id"]
    slides_pdf = tmp_path / "slides.pdf"
    with pymupdf.open() as src, pymupdf.open() as img_pdf:
        for _ in range(3):
            page = img_pdf.new_page(width=960, height=540)
            page.draw_rect(pymupdf.Rect(0, 0, 960, 540), color=(0.8, 0.8, 0.8))
        img_pdf.save(str(slides_pdf))
    out_pdf = tmp_path / "notes.pdf"
    main_mod._compose_notes_pdf(storage.current_path(doc_id), slides_pdf, out_pdf)
    doc = pymupdf.open(str(out_pdf))
    assert doc.page_count == 3
    page0_text = doc[0].get_text()
    assert "讲稿：先回顾里程碑。" in page0_text  # demo 第 1 页的备注
    assert "第 1 页" in page0_text
    assert "第 2 页" in doc[1].get_text() or "备注" in doc[1].get_text()


def test_export_pdf_notes_endpoint(client, uploaded_doc, monkeypatch):
    import main as main_mod

    doc_id = uploaded_doc["doc_id"]
    calls = {}

    def fake_compose(pptx_path, slides_pdf, out_path):
        calls["pptx"] = str(pptx_path)
        out_path.write_bytes(b"%PDF-1.4 notes-handout-stub\n%%EOF\n")

    monkeypatch.setattr(main_mod, "_compose_notes_pdf", fake_compose)
    r = client.get(f"/api/docs/{doc_id}/export/pdf-notes")
    assert r.status_code == 200, r.text
    assert r.headers["content-type"].startswith("application/pdf")
    assert b"notes-handout-stub" in r.content
    assert "current.pptx" in calls["pptx"]
    assert "notes" in r.headers.get("content-disposition", "")


def test_export_pdf_notes_missing_doc(client, uploaded_doc):
    r = client.get("/api/docs/ffffffffffffffffffffffffffffffff/export/pdf-notes")
    assert r.status_code == 404


# ---------- 10. /export/outline ----------

def test_export_outline_contains_full_text(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.get(f"/api/docs/{doc_id}/export/outline")
    assert r.status_code == 200, r.text
    assert r.headers["content-type"].startswith("text/plain")
    body = r.text
    assert "第 1 页" in body and "第 3 页" in body
    assert "项目进展汇报" in body
    assert "本季度完成了三个关键里程碑。" in body
    assert "基础设施建设" in body  # 表格单元格文字
    assert "讲稿：先回顾里程碑。" in body  # 备注
    assert "outline" in r.headers.get("content-disposition", "")


def test_export_outline_reflects_edits(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": "s0_id3", "after": "全新总结话语。"}]},
    )
    assert r.status_code == 200, r.text
    body = client.get(f"/api/docs/{doc_id}/export/outline").text
    assert "全新总结话语。" in body
    assert "本季度完成了三个关键里程碑。" not in body


# ---------- 11. /api/docs/import-images ----------

def test_import_images_creates_deck(client):
    files = [
        ("files", ("a.png", io.BytesIO(TINY_PNG), "image/png")),
        ("files", ("b.png", io.BytesIO(TINY_PNG), "image/png")),
    ]
    r = client.post("/api/docs/import-images", files=files)
    assert r.status_code == 200, r.text
    doc = r.json()
    assert doc["slide_count"] == 2
    assert doc["file_name"].endswith(".pptx")
    # 每页一张图
    from pptx import Presentation

    import storage

    prs = Presentation(str(storage.current_path(doc["doc_id"])))
    for slide in prs.slides:
        pics = [s for s in slide.shapes if str(getattr(s, "shape_type", "")) == "PICTURE (13)"]
        assert len(pics) == 1
        assert int(pics[0].width) == int(prs.slide_width)  # 满宽


def test_import_images_rejects_non_image(client):
    files = [("files", ("a.txt", io.BytesIO(b"hello"), "text/plain"))]
    r = client.post("/api/docs/import-images", files=files)
    assert r.status_code == 400


def test_import_images_requires_at_least_one(client):
    r = client.post("/api/docs/import-images", files=[])
    assert r.status_code == 400


# ---------- 12. GET /api/docs/{id}/versions/{a}/compare/{b} ----------

def _edit_body(client, doc_id, after):
    return client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": "s0_id3", "after": after}]},
    )


def test_compare_detects_text_change(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert _edit_body(client, doc_id, "新版本正文内容。").status_code == 200
    r = client.get(f"/api/docs/{doc_id}/versions/v0/compare/v1")
    assert r.status_code == 200, r.text
    data = r.json()
    assert data["a"] == "v0" and data["b"] == "v1"
    assert data["pages_a"] == 3 and data["pages_b"] == 3
    changed_pages = {c["page"] for c in data["changed"]}
    assert changed_pages == {1}
    entry = data["changed"][0]
    assert any("新版本正文内容" in t for t in entry["additions"])
    assert any("三个关键里程碑" in t for t in entry["deletions"])
    assert data["unchanged_pages"] == 2


def test_compare_detects_page_removal_and_addition(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert _edit_body(client, doc_id, "第一版修改。").status_code == 200  # v1
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "delete_slide", "slide": 2}]},
    )
    assert r.status_code == 200, r.text  # v2：删除第 3 页
    data = client.get(f"/api/docs/{doc_id}/versions/v0/compare/v2").json()
    removed_pages = {c["page"] for c in data["removed"]}
    assert removed_pages == {3}
    assert data["pages_a"] == 3 and data["pages_b"] == 2

    # 反向对比：v2 → v0 视为「新增」
    data_rev = client.get(f"/api/docs/{doc_id}/versions/v2/compare/v0").json()
    assert {c["page"] for c in data_rev["added"]} == {3}


def test_compare_same_version_no_diff(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    data = client.get(f"/api/docs/{doc_id}/versions/v0/compare/v0").json()
    assert data["changed"] == [] and data["added"] == [] and data["removed"] == []
    assert data["unchanged_pages"] == 3


def test_compare_rejects_unknown_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert client.get(f"/api/docs/{doc_id}/versions/v0/compare/v99").status_code == 404
    assert client.get(f"/api/docs/{doc_id}/versions/v0/compare/junk").status_code == 400


def test_export_outline_ascii_header_with_cjk_title(client, uploaded_doc):
    """中文标题导出 outline：Content-Disposition 必须 ASCII 安全（latin-1 头编码 400 回归）。"""
    from urllib.parse import unquote

    doc_id = uploaded_doc["doc_id"]
    assert (
        client.patch(f"/api/docs/{doc_id}", json={"title": "季度汇报 副本"}).status_code == 200
    )
    resp = client.get(f"/api/docs/{doc_id}/export/outline")
    assert resp.status_code == 200, resp.text
    assert resp.headers["content-type"].startswith("text/plain")
    disp = resp.headers.get("content-disposition", "")
    disp.encode("ascii")  # 头本身必须可 latin-1/ascii 编码
    # RFC 5987：非 ASCII 文件名走 filename*，且能还原出中文
    assert "filename*=UTF-8''" in disp
    encoded = disp.split("filename*=UTF-8''", 1)[1].split(";")[0].strip('"')
    assert "季度汇报" in unquote(encoded)
    assert resp.text  # 正文是大纲文本，不为空
