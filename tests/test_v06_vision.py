"""v0.6 #1：渲染验证闭环 —— 渲染 PNG 喂多模态模型自检（vision_check）。

闭环三段：消息构造（PNG → base64 data URL vision 消息）→ HTTP 端点 →
Agent 工具（读当前版本渲染图，发现溢出/乱码/塌版后自提案修复）。
LLM 调用全部用 _post_chat 桩替换，不发真实请求。
"""
from __future__ import annotations

import base64
import json

import pytest

from services import agent, vision


def _fake_png() -> bytes:
    return bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
    )


# ---------- 消息构造 + 解析 ----------


def test_check_page_png_builds_vision_message(tmp_path, monkeypatch):
    png = tmp_path / "slide-001.png"
    png.write_bytes(_fake_png())
    captured = {}

    def fake_post(messages, max_tokens=1200):
        captured["messages"] = messages
        return {"ok": False, "issues": [{"kind": "garbled", "description": "正文有乱码", "suggestion": "检查符号字体"}]}

    monkeypatch.setattr(vision, "_post_chat", fake_post)
    monkeypatch.setenv("DEEPSEEK_API_KEY", "test-key")
    result = vision.check_page_png(png, page_number=3)

    user = captured["messages"][-1]["content"]
    parts = {p["type"]: p for p in user if isinstance(p, dict)}
    assert "image_url" in parts
    url = parts["image_url"]["image_url"]["url"]
    assert url.startswith("data:image/png;base64,")
    payload = url.split(",", 1)[1]
    assert base64.b64decode(payload) == png.read_bytes()
    assert "第 3 页" in parts["text"]["text"]
    assert result["page"] == 3
    assert result["ok"] is False
    assert result["issues"][0]["kind"] == "garbled"


def test_check_page_png_normalizes_and_caps_issues(tmp_path, monkeypatch):
    png = tmp_path / "slide-001.png"
    png.write_bytes(_fake_png())
    monkeypatch.setenv("DEEPSEEK_API_KEY", "test-key")
    raw = {
        "ok": True,
        "issues": [
            {"kind": "overflow", "description": f"溢出{i}", "suggestion": f"缩字号{i}"}
            for i in range(12)
        ] + [{"kind": "weird-kind", "description": "未知类型", "suggestion": ""}],
    }
    monkeypatch.setattr(vision, "_post_chat", lambda messages, max_tokens=1200: raw)
    result = vision.check_page_png(png, page_number=1)
    assert len(result["issues"]) <= vision.MAX_ISSUES_PER_PAGE
    assert all(i["kind"] in vision.ISSUE_KINDS for i in result["issues"])
    # 模型自报 ok=True 但给出了问题 → 以问题为准
    assert result["ok"] is False


def test_check_page_png_requires_key(tmp_path):
    png = tmp_path / "slide-001.png"
    png.write_bytes(_fake_png())
    import os

    old = os.environ.pop("DEEPSEEK_API_KEY", None)
    try:
        with pytest.raises(vision.VisionError):
            vision.check_page_png(png, page_number=1)
    finally:
        if old is not None:
            os.environ["DEEPSEEK_API_KEY"] = old


# ---------- HTTP 端点 ----------


def test_vision_endpoint_and_tool(uploaded_doc, client, monkeypatch):
    from services import vision as vision_mod

    monkeypatch.setattr(
        vision_mod,
        "_post_chat",
        lambda messages, max_tokens=1200: {
            "ok": False,
            "issues": [{"kind": "garbled", "description": "正文出现豆腐块", "suggestion": "替换 Wingdings 符号"}],
        },
    )
    doc_id = uploaded_doc["doc_id"]
    resp = client.get(f"/api/docs/{doc_id}/vision-check?slides=0")
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["doc_id"] == doc_id
    assert body["pages"][0]["page"] == 1
    assert body["pages"][0]["ok"] is False
    assert body["pages"][0]["issues"][0]["kind"] == "garbled"
    assert body["issue_count"] == 1
    assert body["summary"]

    # 越界页码 → 400
    resp = client.get(f"/api/docs/{doc_id}/vision-check?slides=9")
    assert resp.status_code == 400


def test_vision_endpoint_missing_png(uploaded_doc, client, monkeypatch):
    import storage

    doc_id = uploaded_doc["doc_id"]
    version = int(uploaded_doc["version"])
    png = storage.image_path(doc_id, version, 0)
    png.unlink(missing_ok=True)
    resp = client.get(f"/api/docs/{doc_id}/vision-check?slides=0")
    assert resp.status_code == 404


def test_vision_endpoint_no_key(uploaded_doc, client):
    doc_id = uploaded_doc["doc_id"]
    resp = client.get(f"/api/docs/{doc_id}/vision-check?slides=0")
    assert resp.status_code == 502
    assert "多模态" in resp.json()["detail"] or "AI" in resp.json()["detail"]


def test_apply_response_carries_vision_url(uploaded_doc, client):
    doc_id = uploaded_doc["doc_id"]
    patch = {
        "action": "edit",
        "shape_id": "s0_id2",
        "after": "季度进展汇报（改）",
    }
    resp = client.post(
        f"/api/docs/{doc_id}/apply", json={"patches": [patch], "label": "编辑"}
    )
    assert resp.status_code == 200, resp.text
    body = resp.json()
    assert body["vision_check_url"].startswith(f"/api/docs/{doc_id}/vision-check")


# ---------- Agent 工具 ----------


def test_agent_vision_check_tool(monkeypatch):
    calls = {"vision": None}

    def fake_vision(slides=None):
        calls["vision"] = list(slides or [])
        return {
            "pages": [
                {"page": 1, "ok": False, "issues": [{"kind": "overflow", "description": "文字溢出", "suggestion": "缩字号"}]}
            ],
            "issue_count": 1,
        }

    steps = [
        {"thought": "先自检当前页渲染", "action": "vision_check", "args": {"slides": [0]}},
        {"thought": "发现问题，收尾", "action": "finish", "args": {"reply": "发现 1 处溢出，建议缩字号"}},
    ]
    seq = list(steps)
    monkeypatch.setattr(
        "services.streaming.stream_chat_json",
        lambda messages, on_token=None, temperature=0.2, max_tokens=4000: (
            (lambda p: (p, json.dumps(p, ensure_ascii=False)))(seq.pop(0))
        ),
    )
    events = []
    done = agent.run_stream(
        lambda name, data: events.append((name, data)),
        objects=[{"shape_id": "s0_id2", "slide_index": 0, "kind": "text", "text": "标题",
                  "bbox_emu": {"x": 1, "y": 1, "w": 100, "h": 10}, "grouped": False,
                  "deletable": True, "movable": True, "resizable": True}],
        instruction="自检第一页渲染",
        total_slides=1,
        scope="current",
        slide_index=0,
        vision_check=fake_vision,
    )
    assert calls["vision"] == [0]
    tool_events = [data for name, data in events if name == "tool_call"]
    assert any(e.get("tool") == "vision_check" for e in tool_events)
    assert done["finished"] is True
    # Observation 已把 issues 带回模型（第二条 user 消息里可见 overflow）
    assert "vision_check" in agent.TOOL_NAMES


def test_agent_vision_check_tool_unavailable(monkeypatch):
    steps = [
        {"thought": "自检", "action": "vision_check", "args": {}},
        {"thought": "不可用就说明", "action": "finish", "args": {"reply": "视觉自检不可用"}},
    ]
    seq = list(steps)
    seen = {"obs": []}

    def fake_stream(messages, on_token=None, temperature=0.2, max_tokens=4000):
        seen["obs"].append(messages[-1]["content"])
        parsed = seq.pop(0)
        return parsed, json.dumps(parsed, ensure_ascii=False)

    monkeypatch.setattr("services.streaming.stream_chat_json", fake_stream)
    done = agent.run_stream(
        lambda *_: None,
        objects=[{"shape_id": "s0_id2", "slide_index": 0, "kind": "text", "text": "标题",
                  "bbox_emu": {"x": 1, "y": 1, "w": 100, "h": 10}, "grouped": False,
                  "deletable": True, "movable": True, "resizable": True}],
        instruction="自检",
        total_slides=1,
        scope="current",
        slide_index=0,
    )
    assert done["finished"] is True
    assert "不可用" in seen["obs"][1]
