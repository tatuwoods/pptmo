"""v0.4 E 组：AI 提示词升级 — 21 风格模仿 / 22 中英双语汇报。"""
from __future__ import annotations

import json

import pytest

from services import agent, ai


# ---------- 21. 风格模仿：系统提示词 ----------

def test_style_imitation_in_system_prompt():
    text = agent.SYSTEM_PROMPT
    # 用户给出参考页时要分析并复用其风格规律
    assert "参考页" in text
    for keyword in ("字体", "配色", "版式"):
        assert keyword in text
    assert "复用" in text
    # 提示先读取参考页再动手，而不是凭空发明风格
    assert "read_slides" in text


# ---------- 22. 中英双语：语言检测 ----------

def test_detect_language_basic():
    assert agent.detect_language("把第3页标题改成蓝色") == "zh"
    assert agent.detect_language("请统一全文字体并加页码") == "zh"
    assert agent.detect_language("照第 2 页的风格做一页总结") == "zh"
    assert agent.detect_language("make the title bold and blue") == "en"
    assert agent.detect_language("Please summarize this deck on one page") == "en"
    # 中英混合以中文为主导（出现中文即中文）
    assert agent.detect_language("PPT 第3页 title 改成 Overview") == "zh"
    # 无从判断（空/纯符号/零散字母）→ 默认中文
    assert agent.detect_language("") == "zh"
    assert agent.detect_language("!!! ??? 123") == "zh"
    assert agent.detect_language("PPT") == "zh"


def _run_script(monkeypatch, steps, instruction):
    calls = {"messages": []}

    def fake_stream(messages, on_token=None, temperature=0.2, max_tokens=4000):
        calls["messages"].append([dict(m) for m in messages])
        parsed = steps.pop(0)
        return parsed, json.dumps(parsed, ensure_ascii=False)

    monkeypatch.setattr("services.streaming.stream_chat_json", fake_stream)
    done = agent.run_stream(
        lambda name, data: None,
        objects=[
            {
                "shape_id": "s0_id2",
                "slide_index": 0,
                "kind": "text",
                "text": "标题文本",
                "bbox_emu": {"x": 1, "y": 1, "w": 100, "h": 10},
            }
        ],
        instruction=instruction,
        total_slides=1,
        scope="current",
        slide_index=0,
    )
    return done, calls


def test_language_directive_default_chinese(monkeypatch):
    _, calls = _run_script(
        monkeypatch,
        [{"thought": "done", "action": "finish", "args": {"reply": "好的"}}],
        "把标题改成蓝色",
    )
    system = calls["messages"][0][0]["content"]
    assert "中文" in system
    assert "English" not in system.split("当前日期")[0]  # 未触发英文切换


def test_language_directive_switches_to_english(monkeypatch):
    _, calls = _run_script(
        monkeypatch,
        [{"thought": "done", "action": "finish", "args": {"reply": "Sure"}}],
        "Please make the title blue and bold",
    )
    system = calls["messages"][0][0]["content"]
    # 英文提问 → 系统提示切换为英文回复指令
    assert "English" in system
    assert "summary" in system and "reply" in system
    # 幻灯片内容语言不随之切换（遵循内容本身）
    assert "slide content" in system or "幻灯片内容" in system
