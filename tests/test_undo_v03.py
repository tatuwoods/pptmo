"""v0.3 C 组：撤销端点 /api/docs/{id}/undo（版本快照回退，AI undo_last 工具与 ⌘Z 共用）。"""
from __future__ import annotations

import re
from pathlib import Path

import storage


def _apply(client, doc_id, patches, **kw):
    r = client.post(f"/api/docs/{doc_id}/apply", json={"patches": patches, **kw})
    assert r.status_code == 200, r.text
    return r.json()


# ---------- 15. undo endpoint ----------


def test_undo_restores_previous_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    _apply(client, doc_id, [{"action": "duplicate_slide", "source_slide": 0}])  # v1: 4 slides
    r = client.post(f"/api/docs/{doc_id}/undo")
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["version_id"] == "v2"
    assert body["slide_count"] == 3  # back to v0 content
    assert any("撤销" in str(m.get("content", "")) for m in body["messages"][-3:])
    # the undo'd content is a real rollback: v2.pptx == v0 bytes
    assert (storage.versions_dir(doc_id) / "v2.pptx").read_bytes() == \
        (storage.versions_dir(doc_id) / "v0.pptx").read_bytes()
    # v1 snapshot still recoverable in history
    versions = client.get(f"/api/docs/{doc_id}/versions").json()["versions"]
    assert {v["id"] for v in versions} >= {"v0", "v1", "v2"}


def test_undo_twice_walks_back_two_versions(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    _apply(client, doc_id, [{"action": "duplicate_slide", "source_slide": 0}])  # v1
    _apply(client, doc_id, [{"action": "delete_slide", "slide": 0}])            # v2
    r = client.post(f"/api/docs/{doc_id}/undo")  # → v3 == v1 (4 slides)
    assert r.status_code == 200
    assert r.json()["slide_count"] == 4
    r2 = client.post(f"/api/docs/{doc_id}/undo")  # → v4 == v0 (3 slides)
    assert r2.status_code == 200
    assert r2.json()["slide_count"] == 3


def test_undo_at_v0_returns_friendly_400(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/undo")
    # v0 无可撤销：用户操作错误 → 400（与 test_agent_c_v03 统一；并发冲突才是 409）
    assert r.status_code == 400
    assert "撤销" in r.json()["detail"] or "版本" in r.json()["detail"]


def test_undo_base_version_conflict_409(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    _apply(client, doc_id, [{"action": "duplicate_slide", "source_slide": 0}])
    r = client.post(f"/api/docs/{doc_id}/undo", json={"base_version": 99})
    assert r.status_code == 409


def test_undo_unknown_doc_404(client):
    r = client.post(f"/api/docs/{'0' * 32}/undo")
    assert r.status_code == 404


# ---------- wiring contracts ----------


def test_chat_passes_undo_callback_to_agent():
    """chat worker 必须把版本回退器注入 agent（undo_last 工具的执行后端）。"""
    source = Path(__file__).resolve().parent.parent / "app" / "backend" / "main.py"
    text = source.read_text(encoding="utf-8")
    assert "run_stream" in text
    assert re.search(r"undo_last\s*=", text) or "undo_last=" in text


def test_agent_prompt_documents_plan_and_undo():
    source = Path(__file__).resolve().parent.parent / "app" / "backend" / "services" / "agent.py"
    text = source.read_text(encoding="utf-8")
    prompt = text.split("SYSTEM_PROMPT", 1)[1]
    assert "plan" in prompt and "undo_last" in prompt
    assert "重试" in prompt or "修正参数" in prompt  # 自愈指引进了系统提示词
