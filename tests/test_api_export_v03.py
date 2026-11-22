"""v0.3-B: 导入导出 — export/pdf, export/images.zip, POST duplicate."""
from __future__ import annotations

import io
import json
import zipfile

STUB_PDF = b"%PDF-1.4 stub-for-tests\n%%EOF\n"


# ---------- B11: export PDF ----------

def test_export_pdf_streams_kept_pdf(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.get(f"/api/docs/{doc_id}/export/pdf")
    assert r.status_code == 200, r.text
    assert r.headers["content-type"].startswith("application/pdf")
    assert r.content == STUB_PDF
    assert "demo" in r.headers.get("content-disposition", "")


def test_export_pdf_regenerates_when_missing(client, uploaded_doc, monkeypatch):
    import storage

    doc_id = uploaded_doc["doc_id"]
    pdf = storage.render_dir(doc_id, 0) / "export.pdf"
    assert pdf.exists()
    pdf.unlink()
    calls = []
    real_render = client.app  # ensure render path still wired via conftest stub

    import main as app_module
    from services import render as render_mod
    orig = render_mod.render_pptx

    def counting_render(pptx_path, out_dir, scale=2.0):
        calls.append((str(pptx_path), str(out_dir)))
        return orig(pptx_path, out_dir, scale)

    monkeypatch.setattr(app_module.render, "render_pptx", counting_render)
    r = client.get(f"/api/docs/{doc_id}/export/pdf")
    assert r.status_code == 200, r.text
    assert r.content == STUB_PDF
    assert len(calls) == 1  # regenerated exactly once


def test_export_pdf_tracks_current_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    doc = client.get(f"/api/docs/{doc_id}").json()
    target = next(o for o in doc["objects"] if o["kind"] == "text")
    client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": target["shape_id"],
                           "after": "v1 文案", "before": target["text"]}]},
    )
    r = client.get(f"/api/docs/{doc_id}/export/pdf")
    assert r.status_code == 200
    import storage
    assert (storage.render_dir(doc_id, 1) / "export.pdf").exists()


def test_export_pdf_missing_doc_404(client):
    assert client.get(f"/api/docs/{'0' * 32}/export/pdf").status_code == 404


def test_export_pdf_prefers_existing_without_rerender(client, uploaded_doc, monkeypatch):
    import main as app_module
    import storage

    doc_id = uploaded_doc["doc_id"]
    pdf = storage.render_dir(doc_id, 0) / "export.pdf"
    pdf.unlink()
    pdf.write_bytes(b"%PDF-1.4 placed-by-hand")
    calls = []
    orig = app_module.render.render_pptx
    monkeypatch.setattr(
        app_module.render,
        "render_pptx",
        lambda *a, **k: (calls.append(1), orig(*a, **k))[1],
    )
    r = client.get(f"/api/docs/{doc_id}/export/pdf")
    assert r.status_code == 200
    assert r.content == b"%PDF-1.4 placed-by-hand"
    assert not calls  # existing kept PDF must be served as-is


def test_export_pdf_without_pdf_friendly_404(client, uploaded_doc, monkeypatch):
    import main as app_module
    import storage

    doc_id = uploaded_doc["doc_id"]
    (storage.render_dir(doc_id, 0) / "export.pdf").unlink()

    def png_only_render(pptx_path, out_dir, scale=2.0):
        out_dir.mkdir(parents=True, exist_ok=True)
        for i in (1, 2, 3):
            (out_dir / f"slide-{i:03d}.png").write_bytes(b"png")
        return 3  # PNGs but never a PDF

    monkeypatch.setattr(app_module.render, "render_pptx", png_only_render)
    r = client.get(f"/api/docs/{doc_id}/export/pdf")
    assert r.status_code == 404
    assert "PDF" in r.json()["detail"]  # friendly, mentions PDF


def test_render_keeps_pdf_copy(tmp_path):
    # 渲染链渲染完保留一份 export.pdf 供导出复用（v0.2.2 是渲染完即删）
    from services import render as render_mod

    (tmp_path / "out").mkdir()
    (tmp_path / "work.pdf").write_bytes(b"%PDF-keep")
    kept = render_mod.keep_pdf(tmp_path / "out", tmp_path / "work.pdf")
    assert kept == tmp_path / "out" / "export.pdf"
    assert kept.read_bytes() == b"%PDF-keep"
    assert not (tmp_path / "work.pdf").exists()


# ---------- B12: export PNG zip ----------

def test_export_images_zip(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.get(f"/api/docs/{doc_id}/export/images.zip")
    assert r.status_code == 200, r.text
    assert r.headers["content-type"].startswith("application/zip")
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        names = sorted(zf.namelist())
    assert names == [f"slide-{i:03d}.png" for i in (1, 2, 3)]
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        assert zf.read("slide-001.png")  # non-empty payload


def test_export_images_zip_missing_doc_404(client):
    assert client.get(f"/api/docs/{'0' * 32}/export/images.zip").status_code == 404


def test_export_images_zip_empty_doc_404(client, uploaded_doc, monkeypatch):
    import storage
    doc_id = uploaded_doc["doc_id"]
    # wipe rendered pages; regeneration is pdf-oriented, a deck with no pages has nothing to zip
    for png in storage.render_dir(doc_id, 0).glob("slide-*.png"):
        png.unlink()
    monkeypatch.setattr(
        "main.render.render_pptx",
        lambda *a, **k: (_ for _ in ()).throw(FileNotFoundError("no pages")),
    )
    assert client.get(f"/api/docs/{doc_id}/export/images.zip").status_code == 404


# ---------- B13: duplicate doc ----------

def test_duplicate_creates_independent_session(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/duplicate")
    assert r.status_code == 200, r.text
    dup = r.json()
    assert dup["doc_id"] != doc_id
    assert dup["slide_count"] == uploaded_doc["slide_count"]
    assert dup["version"] == 0 and dup["version_id"] == "v0"
    assert "副本" in dup["title"]
    # 副本不带聊天记录，但有一条说明来源的系统备注
    assert len(dup["messages"]) == 1
    assert dup["messages"][0]["role"] == "system"
    assert "副本" in dup["messages"][0]["content"]
    # shows up in the session list
    sessions = client.get("/api/sessions").json()["sessions"]
    assert any(d["doc_id"] == dup["doc_id"] for d in sessions)
    # original untouched
    assert client.get(f"/api/docs/{doc_id}").status_code == 200
    # both downloadable and byte-identical to the source current.pptx
    import storage
    src = storage.current_path(doc_id).read_bytes()
    assert client.get(f"/api/docs/{dup['doc_id']}/download").content == src


def test_duplicate_after_edits_copies_latest(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    doc = client.get(f"/api/docs/{doc_id}").json()
    target = next(o for o in doc["objects"] if o["kind"] == "text")
    client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": target["shape_id"],
                           "after": "编辑后文案", "before": target["text"]}]},
    )
    dup = client.post(f"/api/docs/{doc_id}/duplicate").json()
    assert dup["version"] == 0
    texts = [o["text"] for o in dup["objects"] if o["kind"] == "text"]
    assert any("编辑后文案" in t for t in texts)


def test_duplicate_missing_doc_404(client):
    assert client.post(f"/api/docs/{'0' * 32}/duplicate").status_code == 404


def test_duplicate_leaves_no_orphan_on_render_failure(client, uploaded_doc, monkeypatch):
    import main as app_module
    import storage

    def boom(*a, **k):
        raise RuntimeError("render down")

    monkeypatch.setattr(app_module.render, "render_pptx", boom)
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/duplicate")
    assert r.status_code == 500
    # no half-created session left behind
    ids = {d["doc_id"] for d in client.get("/api/docs").json()["docs"]}
    assert doc_id in ids and all(i != r.json().get("doc_id", None) for i in ids)
    assert len(ids) == 1
