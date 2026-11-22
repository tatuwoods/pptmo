"""v0.3 E21：渲染超时降级 —— worker 超时后不再整单失败。

有上一版页图可用时：复制已生成的页 + 落 _partial 标记（“部分页渲染中”），
apply/上传链路拿到的是降级预览而不是异常；没有旧页图时维持原有报错。
"""
from __future__ import annotations

import shutil

import pytest


PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)


def _stale_pages(stale_dir, count):
    stale_dir.mkdir(parents=True, exist_ok=True)
    for i in range(count):
        (stale_dir / f"slide-{i + 1:03d}.png").write_bytes(PNG)
        (stale_dir / f"thumb-{i + 1:03d}.png").write_bytes(PNG)


def _timeout_env(monkeypatch, tmp_path):
    monkeypatch.setenv("RENDER_TIMEOUT_S", "1")
    monkeypatch.setenv("PPTMO_RENDER_QUEUE", str(tmp_path / "queue"))


def test_timeout_with_stale_pages_degrades_instead_of_failing(demo_pptx, tmp_path, monkeypatch):
    from services import render

    _timeout_env(monkeypatch, tmp_path)
    stale = tmp_path / "stale"
    _stale_pages(stale, 2)
    out = tmp_path / "out"
    count = render.render_pptx(demo_pptx, out, stale_dir=stale)
    assert count == 2
    assert (out / "slide-001.png").read_bytes() == PNG
    assert (out / "thumb-002.png").exists()
    assert render.partial_marker(out).exists()


def test_timeout_without_stale_pages_still_raises(demo_pptx, tmp_path, monkeypatch):
    from services import render

    _timeout_env(monkeypatch, tmp_path)
    with pytest.raises(render.RenderError):
        render.render_pptx(demo_pptx, tmp_path / "out2")


def test_timeout_is_configurable(monkeypatch):
    from services import render

    monkeypatch.setenv("RENDER_TIMEOUT_S", "7")
    assert render._timeout_s() == 7
    monkeypatch.delenv("RENDER_TIMEOUT_S", raising=False)
    assert render._timeout_s() == 180


def test_apply_survives_render_timeout_with_stale_previews(client, uploaded_doc, monkeypatch, tmp_path):
    """端到端：auto 渲染超时但有旧页图 → apply 仍提交成功，响应标记部分渲染。"""
    import main as app_module
    import storage
    from services import render as render_mod

    doc_id = uploaded_doc["doc_id"]
    monkeypatch.setenv("RENDER_TIMEOUT_S", "1")
    monkeypatch.setenv("PPTMO_RENDER_QUEUE", str(tmp_path / "queue"))

    def slow_render(pptx_path, out_dir, scale=2.0, stale_dir=None):
        # 模拟 worker 卡死：复制 stale 页并报超时（走降级路径）
        assert stale_dir is not None
        shutil.rmtree(out_dir, ignore_errors=True)
        out_dir.mkdir(parents=True, exist_ok=True)
        count = 0
        for png in sorted(stale_dir.glob("slide-*.png")):
            shutil.copy2(png, out_dir / png.name)
            count += 1
        render_mod.partial_marker(out_dir).write_text("partial", encoding="utf-8")
        return count

    monkeypatch.setattr(app_module.render, "render_pptx", slow_render)
    target = next(o for o in uploaded_doc["objects"] if o["kind"] == "text")
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": target["shape_id"],
                           "after": "降级渲染下应用", "before": target["text"]}]},
    )
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["version_id"] == "v1"
    assert body.get("render_partial") is True  # 前端可提示“部分页渲染中”
