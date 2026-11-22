"""v0.3 D 组界面体验契约测试：快捷键 / 拖拽上传 / 缩略图切页 / 深色模式。

前端无 JS 测试基建，以静态契约（关键标识符与行为绑定必须存在）锁住实现，
改坏任意一项都会在这里红。
"""
from __future__ import annotations

from pathlib import Path

WEB = Path(__file__).resolve().parent.parent / "app" / "web"


def _read(name: str) -> str:
    return (WEB / name).read_text(encoding="utf-8")


# ---------- 17. 键盘快捷键 ----------


def test_keyboard_undo_shortcut_wired():
    js = _read("app.js")
    # ⌘Z/Ctrl+Z 触发撤销上一轮应用
    assert "metaKey" in js and "'z'" in js.replace('"z"', "'z'")
    assert "undo" in js
    assert "/undo" in js  # 调用后端撤销端点


def test_keyboard_cmd_enter_send_and_esc_stop():
    js = _read("app.js")
    # ⌘Enter 发送（textarea 的 keydown 里）
    assert "sendChat()" in js
    # Esc 停止生成（streaming 时优先 abort）
    assert "streamController" in js
    assert "abort()" in js


# ---------- 18. 拖拽上传 ----------


def test_drag_upload_overlay_and_friendly_reject():
    html = _read("index.html")
    js = _read("app.js")
    assert "dropOverlay" in html  # 全窗拖拽提示层
    assert "dropOverlay" in js
    # .ppt/.pptm 拖入时前端给出去向清晰的提示（不只是"只支持 .pptx"）
    assert ".pptm" in js


# ---------- 19. 缩略图点击切页联动 ----------


def test_thumb_click_switches_page_and_scope():
    js = _read("app.js")
    assert 'data-slide' in js
    # 点击缩略图：画布切页 + 对话范围同步到该页
    assert "setSlide(" in js


# ---------- 20. 深色模式跟随系统 ----------


def test_dark_mode_follows_system():
    html = _read("index.html")
    css = _read("styles.css")
    assert 'name="color-scheme"' in html and "dark" in html.split('name="color-scheme"')[1][:60]
    assert "@media (prefers-color-scheme: dark)" in css
    # 深色变量组必须覆盖核心表面色，而不是只改一两个变量
    dark_block = css.split("@media (prefers-color-scheme: dark)", 1)[1]
    for var in ("--bg:", "--panel:", "--ink:", "--line:"):
        assert var in dark_block, f"dark block missing {var}"
