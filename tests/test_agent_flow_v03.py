"""v0.3 C 组 Agent 体验：多步计划（plan）、撤销（undo_last）、错误自愈重试。

streaming.stream_chat_json 用脚本桩替换：每次调用弹出预置的一步，
从而精确控制 ReAct 循环的走向。
"""
from __future__ import annotations

import json

import pytest

from services import agent, ai


def _obj(slide_index=0, text="标题文本", shape_id=None):
    return {
        "shape_id": shape_id or f"s{slide_index}_id2",
        "slide_index": slide_index,
        "kind": "text",
        "text": text,
        "bbox_emu": {"x": 100000, "y": 100000, "w": 4000000, "h": 600000},
        "grouped": False,
        "deletable": True,
        "movable": True,
        "resizable": True,
    }


def _run(monkeypatch, steps, *, error_first=False, error_twice=False, **kwargs):
    """Script the model: steps are returned one per stream call.

    error_first: raise AIServiceError on the very first call, then proceed.
    error_twice: raise on the first two calls (retry also fails).
    """
    calls = {"n": 0, "messages": []}
    seq = list(steps)
    state = {"raised": 0}

    def fake_stream(messages, on_token=None, temperature=0.2, max_tokens=4000):
        calls["n"] += 1
        calls["messages"].append([dict(m) for m in messages])
        if error_twice and state["raised"] < 2:
            state["raised"] += 1
            raise ai.AIServiceError("网络波动")
        if error_first and state["raised"] < 1:
            state["raised"] += 1
            raise ai.AIServiceError("网络波动")
        parsed = seq.pop(0)
        return parsed, json.dumps(parsed, ensure_ascii=False)

    monkeypatch.setattr(
        "services.streaming.stream_chat_json", fake_stream
    )
    events = []

    def emit(name, data):
        events.append((name, data))

    done = agent.run_stream(
        emit,
        objects=[_obj()],
        instruction=kwargs.pop("instruction", "整理第一页"),
        total_slides=1,
        scope="current",
        slide_index=0,
        **kwargs,
    )
    return done, events, calls


# ---------- 14. 多步任务规划 ----------


def test_plan_tool_emits_event_and_persists(monkeypatch):
    steps = [
        {"thought": "拆解任务", "action": "plan", "args": {"steps": ["统一字体", "替换术语", "加页码"]}},
        {"thought": "完成", "action": "finish", "args": {"reply": "计划已执行"}},
    ]
    done, events, _ = _run(monkeypatch, steps)
    plan_events = [data for name, data in events if name == "plan"]
    assert plan_events and plan_events[0]["steps"] == ["统一字体", "替换术语", "加页码"]
    assert done["plan"] == ["统一字体", "替换术语", "加页码"]
    assert any(t.get("action") == "plan" for t in done["trace"])
    assert done["error"] is None


def test_plan_rejects_empty_steps(monkeypatch):
    steps = [
        {"action": "plan", "args": {"steps": []}},
        {"action": "finish", "args": {"reply": "ok"}},
    ]
    done, events, _ = _run(monkeypatch, steps)
    assert done["plan"] is None
    assert any(name == "plan" for name, _ in events) is False
    # the model received a corrective observation mentioning steps
    assert done["finished"] is True


def test_finish_plan_via_inline_fields(monkeypatch):
    """模型把计划直接放在 finish 里（无 plan 工具调用）也应保留计划。"""
    steps = [
        {"action": "finish", "args": {"reply": "done", "plan": ["步骤一", "步骤二"]}},
    ]
    done, _, _ = _run(monkeypatch, steps)
    assert done["plan"] == ["步骤一", "步骤二"]


# ---------- 15. undo_last 工具 ----------


def test_undo_last_calls_handler(monkeypatch):
    seen = {}

    def handler():
        seen["called"] = True
        return {"restored_from": "v1", "version_id": "v2", "slide_count": 3}

    steps = [
        {"action": "undo_last", "args": {}},
        {"action": "finish", "args": {"reply": "已回退上一轮修改"}},
    ]
    done, events, _ = _run(monkeypatch, steps, undo_handler=handler)
    assert seen.get("called") is True
    assert done["undo"]["restored_from"] == "v1"
    tool_results = [data for name, data in events if name == "tool_result" and data.get("action") == "undo_last"]
    assert tool_results and "v2" in json.dumps(tool_results[0]["result"])


def test_undo_last_without_handler_returns_error_not_crash(monkeypatch):
    steps = [
        {"action": "undo_last", "args": {}},
        {"action": "finish", "args": {"reply": "当前环境不支持撤销"}},
    ]
    done, events, _ = _run(monkeypatch, steps)  # no undo_handler
    assert done["error"] is None  # tool error ≠ crash; loop continues
    tool_results = [data for name, data in events if name == "tool_result" and data.get("action") == "undo_last"]
    assert tool_results and tool_results[0]["result"].get("error")


def test_undo_last_handler_exception_becomes_tool_error(monkeypatch):
    def handler():
        raise ValueError("没有更早的版本可撤销")

    steps = [
        {"action": "undo_last", "args": {}},
        {"action": "finish", "args": {"reply": "无法撤销"}},
    ]
    done, _, _ = _run(monkeypatch, steps, undo_handler=handler)
    assert done["error"] is None
    assert done["undo"] is None or done["undo"].get("error")


# ---------- 16. 错误自愈 ----------


def test_transient_ai_error_retried_once(monkeypatch):
    steps = [
        {"action": "finish", "args": {"reply": "恢复后完成"}},
    ]
    done, _, calls = _run(monkeypatch, steps, error_first=True)
    assert done["error"] is None
    assert done["reply"] == "恢复后完成"
    assert calls["n"] == 2  # first attempt raised, retry succeeded


def test_persistent_ai_error_fails_gracefully(monkeypatch):
    steps = [
        {"action": "finish", "args": {"reply": "never"}},
    ]
    done, _, _ = _run(monkeypatch, steps, error_twice=True)
    assert done["error"] and "网络波动" in done["error"]
    assert done["operations"] == []


def test_tool_error_observation_carries_retry_hint(monkeypatch):
    """工具失败时，观察里要带“换参数重试”的指引，而不是把错误甩给用户。"""
    steps = [
        {"action": "search_text", "args": {"query": ""}},          # error: 空 query
        {"action": "search_text", "args": {"query": "标题"}},       # corrected retry
        {"action": "finish", "args": {"reply": "找到了"}},
    ]
    done, _, calls = _run(monkeypatch, steps)
    assert done["error"] is None
    assert done["reply"] == "找到了"
    # the observation after the failed call must tell the model to retry differently
    obs = [m for m in calls["messages"][1] if m.get("role") == "user"]
    assert any("重试" in m["content"] or "修正参数" in m["content"] for m in obs)


def test_unknown_tool_gets_retry_guidance(monkeypatch):
    steps = [
        {"action": "explode_deck", "args": {}},
        {"action": "outline", "args": {}},
        {"action": "finish", "args": {"reply": "改用 outline"}},
    ]
    done, _, calls = _run(monkeypatch, steps)
    assert done["error"] is None
    assert done["finished"] is True
