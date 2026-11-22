"""v0.3 C 组 AI 对话体验：

14. 多步任务规划 —— plan 工具：复杂指令先出计划清单，逐项执行并汇报
15. 操作撤销 —— undo_last 工具 + POST /api/docs/{id}/undo 端点（版本快照回退）
16. 错误自愈 —— AI 调用瞬时失败自动重试一次；工具失败得到换参重试引导
"""
from __future__ import annotations

import json

import pytest

from services import agent, parse, streaming


def _deck(demo_pptx):
    _objs, payload = parse.parse_pptx(demo_pptx)
    return payload


def _collect():
    events = []

    def emit(name, data):
        events.append((name, data))

    return emit, events


def _stub_llm(monkeypatch, responses):
    """Queue up canned LLM JSON responses; records the messages each call saw."""
    seen = []

    def fake_stream(messages, on_token=None, temperature=0.2, max_tokens=4000):
        seen.append(list(messages))
        item = responses.pop(0)
        if isinstance(item, Exception):
            raise item
        return item, json.dumps(item, ensure_ascii=False)

    monkeypatch.setattr(streaming, "stream_chat_json", fake_stream)
    return seen


def _run(emit, payload, **kw):
    return agent.run_stream(
        emit,
        objects=payload.get("objects", []),
        elements=payload.get("elements", []),
        tables=payload.get("tables", []),
        total_slides=int(payload.get("slide_count") or 1),
        slide_size=payload.get("slide_size_emu") or {},
        instruction=kw.pop("instruction", "整份字体统一为微软雅黑，然后全局替换旧词"),
        scope="deck",
        **kw,
    )


# ---------- 14. plan ----------


def test_plan_tool_records_steps_and_emits_event(demo_pptx, monkeypatch):
    payload = _deck(demo_pptx)
    _stub_llm(monkeypatch, [
        {"thought": "任务复杂，先规划", "action": "plan",
         "args": {"steps": ["统一字体为微软雅黑", "全局替换旧项目为新计划", "逐页核对页码"]}},
        {"thought": "按计划执行", "action": "propose_operations",
         "args": {"operations": [{"action": "font_unify", "font_name": "微软雅黑"}]}},
        {"action": "finish", "args": {"summary": "完成 3 项计划中的第 1 项", "reply": "已统一字体"}},
    ])
    emit, events = _collect()
    done = _run(emit, payload)
    assert done["error"] is None
    assert done["plan"] == ["统一字体为微软雅黑", "全局替换旧项目为新计划", "逐页核对页码"]
    # 独立 plan 事件：前端可展示“我将：1… 2… 3…”
    plan_events = [d for name, d in events if name == "plan"]
    assert plan_events and plan_events[0]["steps"][0] == "统一字体为微软雅黑"
    assert any(t.get("action") == "plan" for t in done["trace"])
    # 计划后继续执行了操作
    assert done["operations"] and done["operations"][0]["action"] == "font_unify"


def test_plan_tool_rejects_bad_payload(demo_pptx, monkeypatch):
    payload = _deck(demo_pptx)
    seen = _stub_llm(monkeypatch, [
        {"action": "plan", "args": {"steps": ["唯一一步"]}},   # 少于 2 步 → 错误观察
        {"action": "finish", "args": {"reply": "放弃规划"}},
    ])
    emit, events = _collect()
    done = _run(emit, payload)
    assert done["error"] is None
    # 失败的工具结果作为 Observation 回给模型（带重试引导），不直接抛给用户
    obs = [m for m in seen[-1] if m.get("role") == "user"]
    assert any("plan" in str(m.get("content")) for m in obs)
    assert done.get("plan") is None


def test_system_prompt_mentions_plan_and_undo():
    text = agent.SYSTEM_PROMPT
    assert "plan" in text and "undo_last" in text
    # 复杂任务先规划、逐项执行并汇报
    assert "逐项" in text or "计划" in text


# ---------- 15. undo_last (agent tool) ----------


def test_undo_last_tool_invokes_callback(demo_pptx, monkeypatch):
    payload = _deck(demo_pptx)
    calls = []

    def undo(base_version):
        calls.append(base_version)
        return {"restored": "v1", "new_version": "v3", "label": "AI 对话改稿 · 2 项"}

    _stub_llm(monkeypatch, [
        {"thought": "用户要撤销", "action": "undo_last", "args": {}},
        {"action": "finish", "args": {"reply": "已撤销上一轮修改，回到 v1"}},
    ])
    emit, events = _collect()
    done = _run(emit, payload, instruction="撤销上一步", undo_last=undo)
    assert calls == [done["base_version"]]
    assert any(t.get("action") == "undo_last" for t in done["trace"])
    tool_results = [d for name, d in events if name == "tool_result" and d.get("action") == "undo_last"]
    assert tool_results and tool_results[0]["result"]["restored"] == "v1"


def test_undo_last_without_callback_friendly_error(demo_pptx, monkeypatch):
    payload = _deck(demo_pptx)
    _stub_llm(monkeypatch, [
        {"action": "undo_last", "args": {}},
        {"action": "finish", "args": {"reply": "无法撤销"}},
    ])
    emit, events = _collect()
    done = _run(emit, payload, instruction="撤销")
    assert done["error"] is None  # 工具失败不是会话失败
    undo_results = [d for name, d in events if name == "tool_result" and d.get("action") == "undo_last"]
    assert undo_results and "error" in undo_results[0]["result"]


# ---------- 15. undo API ----------


def test_undo_endpoint_restores_previous_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    target = next(o for o in uploaded_doc["objects"] if o["kind"] == "text")
    before_text = target["text"]
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": target["shape_id"],
                           "after": "撤销我", "before": before_text}]},
    )
    assert r.status_code == 200, r.text
    assert r.json()["version_id"] == "v1"

    undo = client.post(f"/api/docs/{doc_id}/undo")
    assert undo.status_code == 200, undo.text
    body = undo.json()
    assert body["version_id"] == "v2"
    assert body["restored_from"] == "v0"
    texts = [o.get("text") for o in body["objects"] if o.get("kind") == "text"]
    assert any(t == before_text for t in texts)  # 文本回到撤销前


def test_undo_endpoint_rejects_version_zero(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/undo")
    assert r.status_code == 400  # v0 无可撤销：用户错误而非版本冲突（与 test_undo_v03 统一）
    assert "撤销" in r.json()["detail"]


def test_undo_endpoint_respects_base_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    target = next(o for o in uploaded_doc["objects"] if o["kind"] == "text")
    client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": target["shape_id"],
                           "after": "x", "before": target["text"]}]},
    )
    r = client.post(f"/api/docs/{doc_id}/undo", json={"base_version": 0})
    assert r.status_code == 409


# ---------- 16. self-heal ----------


def test_transient_ai_error_retried_once(demo_pptx, monkeypatch):
    from services.ai import AIServiceError

    payload = _deck(demo_pptx)
    _stub_llm(monkeypatch, [
        AIServiceError("connection reset by peer"),
        {"action": "finish", "args": {"reply": "重试后成功"}},
    ])
    emit, events = _collect()
    done = _run(emit, payload)
    assert done["error"] is None
    assert done["reply"] == "重试后成功"
    assert done["steps"] == 1  # 重试不消耗步数
    retry_events = [d for name, d in events if name == "retry"]
    assert retry_events and retry_events[0]["attempt"] == 2


def test_ai_error_fails_after_second_try(demo_pptx, monkeypatch):
    from services.ai import AIServiceError

    payload = _deck(demo_pptx)
    _stub_llm(monkeypatch, [AIServiceError("down"), AIServiceError("still down")])
    emit, events = _collect()
    done = _run(emit, payload)
    assert done["error"] and "AI" in done["error"]


def test_failed_tool_gets_retry_guidance(demo_pptx, monkeypatch):
    """工具失败 → Observation 里带调整参数重试的引导，而不是直接失败给用户。"""
    payload = _deck(demo_pptx)
    seen = _stub_llm(monkeypatch, [
        {"action": "read_slides", "args": {"slide_indices": [99]}},   # 越界 → 工具报错
        {"action": "finish", "args": {"reply": "已修正参数"}},
    ])
    emit, events = _collect()
    done = _run(emit, payload)
    assert done["error"] is None
    observations = [m["content"] for m in seen[-1] if m.get("role") == "user"]
    assert any("重试" in obs or "调整" in obs for obs in observations)
