"""v0.4 D 组：性能与稳定性 — 渲染缓存命中 / 后台预热 / 大文档懒加载 / 自动清理。"""
from __future__ import annotations

import os
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest


# ---------- 17. 渲染缓存命中（相同内容不再重渲染） ----------

def test_render_cache_hit_on_restore_same_content(client, uploaded_doc, monkeypatch):
    from services import render as render_mod

    calls = {"n": 0}
    original = render_mod.render_pptx

    def counting(pptx_path, out_dir, scale=2.0, stale_dir=None):
        calls["n"] += 1
        return original(pptx_path, out_dir, scale=scale, stale_dir=stale_dir)

    monkeypatch.setattr(render_mod, "render_pptx", counting)

    import storage

    doc_id = uploaded_doc["doc_id"]
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "v1 内容"}]},
    )
    assert r.status_code == 200
    assert calls["n"] == 1  # 内容变化 → 真渲染

    # 恢复 v0：内容与 v0 完全一致 → 命中缓存，不再渲染
    r = client.post(f"/api/docs/{doc_id}/versions/v0/restore")
    assert r.status_code == 200, r.text
    assert calls["n"] == 1, "相同内容的 restore 应复用已渲染 PNG"

    # 新版本渲染目录有完整页面图 + 与 v0 一致的内容指纹
    doc = client.get(f"/api/docs/{doc_id}").json()
    v2 = doc["version"]
    img = client.get(f"/api/docs/{doc_id}/slides/0/image?v=v{v2}")
    assert img.status_code == 200
    hash_v0 = (storage.render_dir(doc_id, 0) / ".contenthash").read_text().strip()
    hash_v2 = (storage.render_dir(doc_id, v2) / ".contenthash").read_text().strip()
    assert hash_v0 and hash_v0 == hash_v2

    # 再次修改（新内容）→ 缓存未命中，重新渲染
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "v3 内容"}]},
    )
    assert r.status_code == 200
    assert calls["n"] == 2


def test_find_cached_render_skips_partial_and_incomplete(client, uploaded_doc, tmp_path):
    import main as main_mod

    import storage

    doc_id = uploaded_doc["doc_id"]
    base = storage.doc_dir(doc_id) / "render"
    good = storage.render_dir(doc_id, 0)
    assert (good / ".contenthash").exists()
    digest = (good / ".contenthash").read_text().strip()

    # 命中：完整目录
    assert main_mod._find_cached_render(doc_id, digest) == good

    # 降级目录（_partial）不可复用
    partial = base / "v9"
    partial.mkdir(parents=True, exist_ok=True)
    for name in ("slide-001.png", "export.pdf"):
        (partial / name).write_bytes(b"x")
    (partial / ".contenthash").write_text(digest)
    from services import render as render_mod

    render_mod.partial_marker(partial).write_text("stale")
    found = main_mod._find_cached_render(doc_id, digest)
    assert found != partial

    # 缺页目录不可复用
    incomplete = base / "v8"
    incomplete.mkdir(parents=True, exist_ok=True)
    (incomplete / "export.pdf").write_bytes(b"x")
    (incomplete / ".contenthash").write_text(digest)
    assert main_mod._find_cached_render(doc_id, digest) != incomplete

    # 未知指纹 → 无缓存
    assert main_mod._find_cached_render(doc_id, "deadbeef") is None


# ---------- 18. 后台预热：补齐缺失缩略图 ----------

def test_ensure_thumbs_fills_missing(client, uploaded_doc):
    from services import render as render_mod

    import storage

    doc_id = uploaded_doc["doc_id"]
    render_dir = storage.render_dir(doc_id, 0)
    for thumb in render_dir.glob("thumb-*.png"):
        thumb.unlink()
    assert not list(render_dir.glob("thumb-*.png"))
    filled = render_mod.ensure_thumbs(render_dir)
    assert filled == 3
    thumbs = sorted(render_dir.glob("thumb-*.png"))
    assert len(thumbs) == 3
    assert all(t.stat().st_size > 0 for t in thumbs)


def test_prewarm_doc_restores_thumbs(client, uploaded_doc):
    import main as main_mod

    import storage

    doc_id = uploaded_doc["doc_id"]
    render_dir = storage.render_dir(doc_id, int(uploaded_doc["version"] or 0))
    (render_dir / "thumb-002.png").unlink()
    assert not (render_dir / "thumb-002.png").exists()
    main_mod._prewarm_doc(doc_id)
    assert (render_dir / "thumb-002.png").exists()


# ---------- 19. 大文档缩略图懒加载（前端契约） ----------

def test_lazy_load_contract():
    web = Path(__file__).resolve().parent.parent / "app" / "web" / "app.js"
    js = web.read_text(encoding="utf-8")
    # >50 页时缩略图懒加载：IntersectionObserver + data-url 延迟换 src
    assert "IntersectionObserver" in js
    assert "thumbObserver" in js
    assert "LAZY_THUMB_THRESHOLD" in js and "50" in js
    # 懒加载的 img 初始无 src，进入视口才装入真实地址
    lazy_build = js.split("function renderThumbs()", 1)[1][:2400]
    assert "data-url" in lazy_build
    assert "unobserve" in js


# ---------- 20. 自动清理（tmp + 30 天旧版本） ----------

def test_daily_cleanup_prunes_old_versions_and_tmp(client, uploaded_doc):
    import storage

    doc_id = uploaded_doc["doc_id"]
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "v1"}]},
    )
    assert r.status_code == 200
    assert (storage.versions_dir(doc_id) / "v0.pptx").exists()
    assert (storage.versions_dir(doc_id) / "v1.pptx").exists()

    # 把 v0 的创建时间拨回 40 天前（v1 是当前版本，必须保留）
    meta = storage.load_meta(doc_id)
    old_ts = (datetime.now(timezone.utc) - timedelta(days=40)).isoformat()
    for entry in meta["versions"]:
        if entry.get("id") == "v0":
            entry["created_at"] = old_ts
    storage.save_meta(doc_id, meta)

    tmp_dir = storage.DOCS_DIR.parent / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    stale_tmp = tmp_dir / "upload-deadbeef.pptx"
    stale_tmp.write_bytes(b"stale")
    fresh_tmp = tmp_dir / "upload-alive.pptx"
    fresh_tmp.write_bytes(b"fresh")
    old_epoch = time.time() - 40 * 24 * 3600
    os.utime(stale_tmp, (old_epoch, old_epoch))

    removed = storage.run_daily_cleanup(force=True)
    assert removed >= 2  # v0.pptx + 旧 tmp
    assert not (storage.versions_dir(doc_id) / "v0.pptx").exists()
    assert not storage.render_dir(doc_id, 0).exists()
    assert (storage.versions_dir(doc_id) / "v1.pptx").exists()
    assert not stale_tmp.exists()
    assert fresh_tmp.exists()
    meta2 = storage.load_meta(doc_id)
    assert [v["id"] for v in meta2["versions"]] == ["v1"]
    assert int(meta2["version"]) == 1

    # 当天标记：非 force 的二次运行直接跳过
    marker = storage.DOCS_DIR.parent / ".last-cleanup"
    assert marker.exists()
    assert storage.run_daily_cleanup() == 0
