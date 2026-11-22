"""v0.3 E21：渲染超时降级 —— worker 超时后不再整单失败。

有上一版页图可用时：复制已生成的页 + 落 _partial 标记（“部分页渲染中”），
apply/上传链路拿到的是降级预览而不是异常；没有旧页图时维持原有报错。
"""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

WEB_JS = Path(__file__).resolve().parent.parent / "app" / "web" / "app.js"


def _function_body(name: str) -> str:
    js = WEB_JS.read_text(encoding="utf-8")
    return js.split(f"function {name}(", 1)[1].split("\nfunction ", 1)[0]


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


def test_partial_render_hint_is_shown_in_ui():
    """E21 的另一半：后端给了 render_partial，前端必须把它提示出来。

    超时降级后画布上是上一版的旧页图；不提示的话用户会把旧画面当成最新结果。"""
    # 状态栏常驻提示（aiNotice 在 AI 页签里，文本页签下看不到）
    status = _function_body("updateStatus")
    assert "render_partial" in status
    assert "渲染超时" in status and "上一版" in status
    # 刚进入降级状态时再弹一次 toast，保证用户当场注意到
    apply_doc = _function_body("applyDocument")
    assert "render_partial" in apply_doc and "toast(" in apply_doc


def test_manual_apply_keeps_partial_status_note():
    """常驻提示不能在手动保存路径上被盖掉。

    runApply 是手动改字/原位编辑/拖动/应用 AI 建议的共同入口：它先 applyDocument
    （updateStatus 已按 render_partial 选好状态栏文案），若随后再无条件写回
    「所有修改直接写入原 PPTX」，降级态就只剩一个几秒后消失的 toast。"""
    run_apply = _function_body("runApply")
    assert "applyDocument(doc" in run_apply
    # 状态栏文案只由 updateStatus 按 render_partial 决定；runApply 不再自带一份
    assert "所有修改直接写入原 PPTX" not in run_apply
    assert "setStatus(`已保存" not in run_apply
