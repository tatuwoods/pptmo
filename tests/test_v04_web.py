"""v0.4 前端契约：审计面板 / 只读分享入口 / 密码解锁重试 / 大文档懒加载。"""
from __future__ import annotations

from pathlib import Path

WEB = Path(__file__).resolve().parent.parent / "app" / "web"


def _read(name: str) -> str:
    return (WEB / name).read_text(encoding="utf-8")


def test_audit_panel_wired():
    js = _read("app.js")
    html = _read("index.html")
    assert 'id="auditBtn"' in html
    assert "openAuditPanel" in js
    assert "/api/audit?limit=50" in js  # 拉最近 50 条
    css = _read("styles.css")
    assert ".audit-panel" in css and ".audit-list" in css


def test_share_button_wired():
    js = _read("app.js")
    html = _read("index.html")
    assert 'id="shareBtn"' in html
    assert "createShareLink" in js
    assert "/share" in js  # POST /api/docs/{id}/share
    assert "clipboard" in js  # 复制链接


def test_password_retry_prompt():
    js = _read("app.js")
    composer = js.split("async function api(", 1)[1][:2200]
    assert "X-Doc-Password" in composer
    assert "response.status === 401" in composer
    assert "_retriedAuth" in composer  # 只重试一次，避免死循环
