"""API integration tests against TestClient with stubbed rendering."""
from __future__ import annotations

import pytest


def _edit_patch(doc):
    target = next(o for o in doc["objects"] if o["kind"] == "text" and "项目进展汇报" in o["text"])
    return {"action": "edit", "shape_id": target["shape_id"], "after": "季度总结汇报",
            "before": target["text"]}


# ---------- health & basics ----------

def test_health(client):
    r = client.get("/api/health")
    assert r.status_code == 200
    body = r.json()
    assert body["ok"] is True and body["service"] == "ppt-light"


def test_sessions_list(client, uploaded_doc):
    r = client.get("/api/sessions")
    assert r.status_code == 200
    docs = r.json()["sessions"]
    assert any(d["doc_id"] == uploaded_doc["doc_id"] for d in docs)


# ---------- upload validation ----------

def test_upload_rejects_non_pptx(client):
    r = client.post("/api/docs/upload", files={"file": ("x.txt", b"hello", "text/plain")})
    assert r.status_code == 400


def test_upload_rejects_empty(client):
    r = client.post(
        "/api/docs/upload",
        files={"file": ("empty.pptx", b"", "application/octet-stream")},
    )
    assert r.status_code == 400


def test_upload_rejects_corrupt_pptx(client):
    r = client.post(
        "/api/docs/upload",
        files={"file": ("bad.pptx", b"not a zip file at all" * 10, "application/octet-stream")},
    )
    assert r.status_code == 400


# ---------- document shape ----------

def test_get_doc_payload(client, uploaded_doc):
    r = client.get(f"/api/docs/{uploaded_doc['doc_id']}")
    assert r.status_code == 200
    doc = r.json()
    assert doc["slide_count"] == 3
    assert doc["version"] == 0
    assert doc["ai_configured"] is False
    kinds = {o["kind"] for o in doc["objects"]}
    assert "text" in kinds and "table_cell" in kinds
    assert doc["messages"] == []


def test_get_missing_doc_404(client):
    assert client.get("/api/docs/" + "0" * 32).status_code == 404


def test_get_invalid_doc_id_rejected(client):
    # URL-encoded traversal must survive httpx normalization and hit the server guard
    assert client.get("/api/docs/%2e%2e/%2e%2e/etc/passwd").status_code in (400, 404)
    assert client.get("/api/docs/short").status_code in (400, 404)
    assert client.get("/api/docs/" + "z" * 32).status_code in (400, 404)


# ---------- rename ----------

def test_rename(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.patch(f"/api/docs/{doc_id}", json={"title": "改名测试"})
    assert r.status_code == 200 and r.json()["title"] == "改名测试"
    assert client.patch(f"/api/docs/{doc_id}", json={"title": ""}).status_code == 422


# ---------- apply / versions / restore / download ----------

def test_apply_creates_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={"patches": [_edit_patch(uploaded_doc)],
                                                       "label": "测试编辑", "base_version": 0})
    assert r.status_code == 200
    body = r.json()
    assert body["version_id"] == "v1"
    assert body["touched_slides"] == [0]
    # applied message recorded
    msgs = client.get(f"/api/docs/{doc_id}/messages").json()["messages"]
    assert any(m["kind"] == "apply" for m in msgs)


def test_apply_base_version_conflict(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [_edit_patch(uploaded_doc)], "base_version": 5})
    assert r.status_code == 409


def test_apply_empty_patches(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply", json={"patches": []})
    assert r.status_code == 422


def test_apply_invalid_shape(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [{"action": "edit", "shape_id": "s0_id999", "after": "x"}]})
    assert r.status_code in (400, 409)


def test_versions_and_restore(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    client.post(f"/api/docs/{doc_id}/apply", json={"patches": [_edit_patch(uploaded_doc)], "base_version": 0})
    versions = client.get(f"/api/docs/{doc_id}/versions").json()
    assert versions["current"] == "v1"
    assert [v["id"] for v in versions["versions"]] == ["v0", "v1"]
    r = client.post(f"/api/docs/{doc_id}/versions/v0/restore")
    assert r.status_code == 200
    assert r.json()["version_id"] == "v2"  # restore creates a NEW version
    doc = client.get(f"/api/docs/{doc_id}").json()
    assert any("项目进展汇报" in o["text"] for o in doc["objects"])


def test_restore_invalid_version(client, uploaded_doc):
    base = f"/api/docs/{uploaded_doc['doc_id']}/versions"
    # traversal variants are rejected either at routing (405) or by the v\d+ regex (400)
    assert client.post(f"{base}/%2e%2e/restore").status_code != 200
    assert client.post(f"{base}/..%2fsecret%2fpptx/restore").status_code != 200
    assert client.post(f"{base}/v99/restore").status_code == 404
    assert client.post(f"{base}/nope/restore").status_code in (400, 404)


def test_download(client, uploaded_doc):
    r = client.get(f"/api/docs/{uploaded_doc['doc_id']}/download")
    assert r.status_code == 200
    data = r.content
    assert data[:2] == b"PK"  # zip magic
    assert "demo.pptx" in r.headers.get("content-disposition", "")


# ---------- slide images ----------

def test_slide_image(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert client.get(f"/api/docs/{doc_id}/slides/0/image").status_code == 200
    assert client.get(f"/api/docs/{doc_id}/slides/2/image?thumb=true").status_code == 200
    assert client.get(f"/api/docs/{doc_id}/slides/3/image").status_code == 404
    assert client.get(f"/api/docs/{doc_id}/slides/-1/image").status_code == 404
    assert client.get(f"/api/docs/{doc_id}/slides/0/image?v=v9").status_code == 404
    assert client.get(f"/api/docs/{doc_id}/slides/0/image?v=zzz").status_code == 400


# ---------- delete ----------

def test_delete_doc(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert client.delete(f"/api/docs/{doc_id}").status_code == 200
    assert client.get(f"/api/docs/{doc_id}").status_code == 404
    assert client.delete(f"/api/docs/{doc_id}").status_code == 404


# ---------- chat validation (no AI key configured) ----------

def test_chat_rejects_empty_message(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/chat", json={"message": "  "})
    assert r.status_code == 400


def test_chat_rejects_bad_scope(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/chat",
                    json={"message": "改标题", "scope": "galaxy"})
    assert r.status_code == 400


def test_chat_rejects_slide_out_of_range(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/chat",
                    json={"message": "改标题", "scope": "current", "slide_index": 99})
    assert r.status_code == 400


def test_chat_base_version_conflict(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/chat",
                    json={"message": "改标题", "base_version": 42})
    assert r.status_code == 409


def test_chat_without_key_streams_error(client, uploaded_doc):
    """Without a key the agent errors after start; the SSE stream must end with done/error."""
    doc_id = uploaded_doc["doc_id"]
    with client.stream("POST", f"/api/docs/{doc_id}/chat",
                       json={"message": "把标题改短", "scope": "current", "slide_index": 0}) as resp:
        assert resp.status_code == 200
        events = []
        for line in resp.iter_lines():
            if line.startswith("event: "):
                events.append(line[7:].strip())
    assert events[0] == "start"
    assert "error" in events
    assert events[-1] == "done"


# ---------- secrets never leak ----------

def test_no_key_in_responses(client, uploaded_doc, monkeypatch):
    monkeypatch.setenv("DEEPSEEK_API_KEY", "super-secret-key-value")
    doc_id = uploaded_doc["doc_id"]
    for resp in (
        client.get("/api/health"),
        client.get(f"/api/docs/{doc_id}"),
        client.get(f"/api/docs/{doc_id}/versions"),
        client.get("/api/sessions"),
        client.get("/api/docs-ui"),
    ):
        assert "super-secret-key-value" not in resp.text


# ---------- v0.1.1: precise applied-marking, partial applies, page guards ----------

def _stage_assistant_message(doc_id, ops, base_version=0):
    import storage
    import uuid
    from datetime import datetime, timezone
    message = {
        "id": uuid.uuid4().hex[:12],
        "role": "assistant",
        "content": "已收集操作",
        "at": datetime.now(timezone.utc).isoformat(),
        "kind": "chat",
        "operations": ops,
        "base_version": base_version,
    }
    storage.append_messages(doc_id, [message])
    return message


def _op_for_text(doc, needle):
    target = next(o for o in doc["objects"] if o["kind"] == "text" and needle in o["text"])
    return {"action": "edit", "shape_id": target["shape_id"], "after": "改后文本",
            "before": target["text"]}


def test_apply_marks_only_targeted_message(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    doc = client.get(f"/api/docs/{doc_id}").json()
    op_a = _op_for_text(doc, "项目进展汇报")
    op_b = _op_for_text(doc, "本季度完成")
    msg_a = _stage_assistant_message(doc_id, [op_a])
    msg_b = _stage_assistant_message(doc_id, [op_b])
    # apply msg_b's op with explicit message id
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [op_b], "label": "AI 改稿 · 1 项",
        "base_version": 0, "assistant_message_id": msg_b["id"],
    })
    assert r.status_code == 200
    messages = client.get(f"/api/docs/{doc_id}/messages").json()["messages"]
    by_id = {m["id"]: m for m in messages if m.get("id")}
    assert by_id[msg_b["id"]].get("applied_version") == "v1"
    assert "applied_version" not in by_id[msg_a["id"]]


def test_apply_partial_records_partial_applies(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    doc = client.get(f"/api/docs/{doc_id}").json()
    op_a = _op_for_text(doc, "项目进展汇报")
    op_b = _op_for_text(doc, "本季度完成")
    msg = _stage_assistant_message(doc_id, [op_a, op_b])
    # apply only one of the two ops
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [op_b], "label": "AI 改稿 · 1 项",
        "base_version": 0, "assistant_message_id": msg["id"],
    })
    assert r.status_code == 200
    messages = client.get(f"/api/docs/{doc_id}/messages").json()["messages"]
    marked = next(m for m in messages if m.get("id") == msg["id"])
    assert "applied_version" not in marked  # not fully applied
    partial = marked.get("partial_applies") or []
    assert len(partial) == 1
    assert op_b["shape_id"] in partial[0]["applied_shape_ids"]
    assert op_a["shape_id"] not in partial[0]["applied_shape_ids"]


def test_chat_out_of_range_page_rejected(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/chat",
                    json={"message": "帮我改一下第 99 页的标题", "scope": "current",
                          "slide_index": 0, "base_version": 0})
    assert r.status_code == 400
    assert "超出范围" in r.json()["detail"]


def test_chat_disconnect_still_persists(client, uploaded_doc):
    """Abandoning the SSE stream must not lose the assistant message (v0.1.1 fix)."""
    import time
    doc_id = uploaded_doc["doc_id"]
    with client.stream("POST", f"/api/docs/{doc_id}/chat",
                       json={"message": "把标题改短一点", "scope": "current",
                             "slide_index": 0, "base_version": 0}) as resp:
        assert resp.status_code == 200
        for _ in resp.iter_lines():
            break  # read one line then abandon the stream
    # worker thread finalizes regardless; poll for the assistant message
    deadline = time.time() + 10
    found = False
    while time.time() < deadline and not found:
        messages = client.get(f"/api/docs/{doc_id}/messages").json()["messages"]
        found = any(m.get("role") == "assistant" for m in messages)
        if not found:
            time.sleep(0.3)
    assert found, "assistant message lost after client disconnect"


# ---------- v0.1.2: request guard, body caps, patch caps, auto-apply guard ----------

def test_cross_site_post_rejected(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [_edit_patch(uploaded_doc)]},
                    headers={"Sec-Fetch-Site": "cross-site"})
    assert r.status_code == 403


def test_foreign_origin_post_rejected(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [_edit_patch(uploaded_doc)]},
                    headers={"Origin": "https://evil.example"})
    assert r.status_code == 403


def test_same_origin_post_allowed(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [_edit_patch(uploaded_doc)], "base_version": 0},
                    headers={"Origin": "http://testserver", "Sec-Fetch-Site": "same-origin"})
    assert r.status_code == 200


def test_keyless_client_still_allowed(client, uploaded_doc):
    # curl / native API clients send no Origin/Sec-Fetch-Site
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [_edit_patch(uploaded_doc)], "base_version": 0})
    assert r.status_code == 200


def test_rebinding_host_rejected(client):
    r = client.get("/api/health", headers={"Host": "attacker.example"})
    assert r.status_code == 421


def test_json_body_too_large_rejected(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [_edit_patch(uploaded_doc)],
                          "label": "x" * (3 * 1024 * 1024)})
    assert r.status_code == 413


def test_multipart_too_large_rejected(client, demo_pptx, monkeypatch):
    import main as app_module
    monkeypatch.setattr(app_module, "MAX_UPLOAD_BYTES", 100 * 1024)
    monkeypatch.setattr(app_module, "MULTIPART_SLACK_BYTES", 1000)
    payload = demo_pptx.read_bytes() + b"\0" * 200 * 1024
    r = client.post("/api/docs/upload",
                    files={"file": ("big.pptx", payload, "application/octet-stream")})
    assert r.status_code == 413


def test_apply_patch_string_cap(client, uploaded_doc):
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/apply",
                    json={"patches": [{"action": "edit",
                                       "shape_id": uploaded_doc["objects"][0]["shape_id"],
                                       "after": "长" * 25000}]})
    assert r.status_code == 400
    assert "too long" in r.json()["detail"]


def test_destructive_auto_apply_requires_confirmation(client, uploaded_doc, monkeypatch):
    import main as app_module

    doc_id = uploaded_doc["doc_id"]
    doc = client.get(f"/api/docs/{doc_id}").json()
    target = next(o for o in doc["objects"] if o["kind"] == "text")
    ops = [{"action": "clear", "shape_id": target["shape_id"], "slide_index": target["slide_index"]}]

    # default: destructive ops force the review path, nothing is applied
    monkeypatch.setattr(app_module, "AUTO_APPLY_DESTRUCTIVE", False)
    applied_event, apply_error, fields = app_module._finalize_chat(
        doc_id, auto_apply=True, start_version=0,
        done_data={"operations": ops, "summary": "s", "reply": "r"},
    )
    assert applied_event is None
    assert apply_error and "手动确认" in apply_error
    assert client.get(f"/api/docs/{doc_id}").json()["version"] == 0

    # opt-in: auto-apply proceeds
    monkeypatch.setattr(app_module, "AUTO_APPLY_DESTRUCTIVE", True)
    applied_event, apply_error, fields = app_module._finalize_chat(
        doc_id, auto_apply=True, start_version=0,
        done_data={"operations": ops, "summary": "s", "reply": "r"},
    )
    assert applied_event is not None and applied_event.get("version_id")


def test_error_details_sanitized(client, uploaded_doc, tmp_path):
    import main as app_module
    # unit: the sanitizer masks filesystem prefixes
    sanitized = app_module._sanitize_detail(f"boom at {tmp_path}/data/docs/x/current.pptx")
    assert str(tmp_path) not in sanitized
    assert "<path>" in sanitized
    # end-to-end: missing file yields a clean 404 with no server paths
    doc_id = uploaded_doc["doc_id"]
    current = tmp_path / "data" / "docs" / doc_id / "current.pptx"
    current.unlink()
    r = client.get(f"/api/docs/{doc_id}/download")
    assert r.status_code == 404
    assert str(tmp_path) not in r.text


# ---------- v0.1.3: resource caps, pruning, concurrency bounds ----------

def test_version_pruning_keeps_recent(client, uploaded_doc, monkeypatch):
    import storage as storage_mod
    doc_id = uploaded_doc["doc_id"]
    monkeypatch.setenv("MAX_VERSIONS", "2")
    for i in range(3):
        doc = client.get(f"/api/docs/{doc_id}").json()
        target = next(o for o in doc["objects"] if o["kind"] == "text")
        r = client.post(f"/api/docs/{doc_id}/apply", json={
            "patches": [{"action": "style", "shape_id": target["shape_id"],
                         "style": {"bold": bool(i % 2)}}],
            "base_version": doc["version"],
        })
        assert r.status_code == 200, r.text
    versions = client.get(f"/api/docs/{doc_id}/versions").json()["versions"]
    assert len(versions) <= 2
    assert [v["id"] for v in versions] == ["v2", "v3"]
    assert not (storage_mod.versions_dir(doc_id) / "v0.pptx").exists()
    assert (storage_mod.versions_dir(doc_id) / "v3.pptx").exists()
    # pruned render dirs are gone too
    assert not (storage_mod.doc_dir(doc_id) / "render" / "v0").exists()


def test_parse_slide_cap_rejected_on_upload(client, demo_pptx, monkeypatch):
    monkeypatch.setenv("PARSE_MAX_SLIDES", "2")
    with open(demo_pptx, "rb") as fh:
        r = client.post("/api/docs/upload",
                        files={"file": ("demo.pptx", fh.read(), "application/octet-stream")})
    assert r.status_code == 413
    assert "上限" in r.json()["detail"]


def test_parse_truncation_flag_direct(demo_pptx, monkeypatch):
    from services import parse as parse_mod
    monkeypatch.setenv("PARSE_MAX_SLIDES", "2")
    objects, payload = parse_mod.parse_pptx(str(demo_pptx))
    assert payload["slide_count"] == 2
    assert payload["truncated_slides"] is True


def test_render_busy_maps_to_429(client, demo_pptx, monkeypatch):
    from services import render as render_mod

    def busy_render(*args, **kwargs):
        raise render_mod.RenderBusy("render queue is full")

    monkeypatch.setattr(render_mod, "render_pptx", busy_render)
    with open(demo_pptx, "rb") as fh:
        r = client.post("/api/docs/upload",
                        files={"file": ("demo.pptx", fh.read(), "application/octet-stream")})
    assert r.status_code == 429
    assert "渲染队列已满" in r.json()["detail"]


def test_render_pptx_busy_semantics(tmp_path):
    import threading
    from services import render as render_mod
    import pytest as _pytest
    original = render_mod._queue_slots
    render_mod._queue_slots = threading.BoundedSemaphore(0)
    try:
        with _pytest.raises(render_mod.RenderBusy):
            render_mod.render_pptx(tmp_path / "x.pptx", tmp_path / "out")
    finally:
        render_mod._queue_slots = original


def test_chat_per_doc_guard(client, uploaded_doc, monkeypatch):
    import main as app_module
    doc_id = uploaded_doc["doc_id"]
    monkeypatch.setenv("DEEPSEEK_API_KEY", "k")  # pass the ai_configured gate client-side irrelevant; server doesn't check
    with app_module._chat_inflight_lock:
        app_module._chat_inflight[doc_id] = 1
    try:
        r = client.post(f"/api/docs/{doc_id}/chat", json={"message": "改标题"})
        assert r.status_code == 429
        assert "已有对话进行中" in r.json()["detail"]
    finally:
        with app_module._chat_inflight_lock:
            app_module._chat_inflight.pop(doc_id, None)


def test_chat_slot_exhaustion_429(client, uploaded_doc, monkeypatch):
    import threading
    import main as app_module
    monkeypatch.setattr(app_module, "_chat_slots", threading.BoundedSemaphore(0))
    r = client.post(f"/api/docs/{uploaded_doc['doc_id']}/chat", json={"message": "改标题"})
    assert r.status_code == 429
    assert "AI 并发已满" in r.json()["detail"]

import pytest

from services import edits


def _objects_of(demo_pptx):
    from services import parse as parse_mod

    objects, _payload = parse_mod.parse_pptx(str(demo_pptx))
    return objects


# ---------- v0.1.4: atomicity, reconciliation, pending ops, robustness ----------

def test_reconcile_meta_adopts_orphaned_version(client, uploaded_doc):
    import storage as storage_mod
    doc_id = uploaded_doc["doc_id"]
    # simulate a crash after the version file write but before the meta save
    current = storage_mod.current_path(doc_id)
    storage_mod.versions_dir(doc_id).joinpath("v9.pptx.tmp").unlink(missing_ok=True)
    (storage_mod.versions_dir(doc_id) / "v9.pptx").write_bytes(current.read_bytes())
    meta = storage_mod.reconcile_meta(doc_id)
    assert meta["version"] == 9
    assert [v["id"] for v in meta["versions"]][-1] == "v9"
    assert any(v["label"] == "recovered" for v in meta["versions"])
    # restored doc is consistent afterwards
    doc = client.get(f"/api/docs/{doc_id}").json()
    assert doc["version"] == 9


def test_pending_operations_survives_message_window(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    doc = client.get(f"/api/docs/{doc_id}").json()
    op = _op_for_text(doc, "本季度完成")
    message = _stage_assistant_message(doc_id, [op])
    # flood the trail past the 60-message API window
    import storage as storage_mod
    fillers = [
        {"id": f"f{i:03}", "role": "user", "content": f"填充消息 {i}", "kind": "chat"}
        for i in range(80)
    ]
    storage_mod.append_messages(doc_id, fillers)
    fresh = client.get(f"/api/docs/{doc_id}").json()
    # ops no longer in the trimmed message list, but pending_operations survives
    assert all(m.get("id") != message["id"] for m in fresh["messages"])
    pending = fresh["pending_operations"]
    assert pending and pending["message_id"] == message["id"]
    assert pending["operations"][0]["shape_id"] == op["shape_id"]


def test_restore_base_version_conflict(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/versions/v0/restore", json={"base_version": 42})
    assert r.status_code == 409
    r = client.post(f"/api/docs/{doc_id}/versions/v0/restore", json={"base_version": 0})
    assert r.status_code == 200


def test_bbox_zero_width_rejected(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    with pytest.raises(edits.PatchError, match="width must be positive"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "geometry", "shape_id": target["shape_id"],
              "bbox_emu": {"x": 1000000, "y": 1000000, "w": 0, "h": 500000}}],
        )


def test_slide_index_mismatch_rejected(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    with pytest.raises(edits.PatchError, match="does not match"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "edit", "shape_id": target["shape_id"], "after": "x",
              "slide_index": 1}],  # target lives on slide 0
        )


def test_trim_messages_keeps_json_valid():
    import json as _json
    from services import agent as agent_mod
    observation = {"slides": [{"objects": [{"text": "长" * 9000}, {"text": "y"}]}] * 3}
    long_content = _json.dumps(observation, ensure_ascii=False)
    trimmed = agent_mod._trim_messages([
        {"role": "system", "content": "sys"},
        {"role": "user", "content": "task"},
        {"role": "user", "content": long_content},
    ])
    content = trimmed[-1]["content"]
    assert len(content) <= 12000
    # clipped content must still be valid JSON the model can read
    parsed = _json.loads(content)
    assert isinstance(parsed, dict)


def test_streaming_done_sentinel():
    from services import streaming as streaming_mod

    class FakeResponse:
        headers = {"content-type": "text/event-stream"}
        def __init__(self, lines):
            self._lines = lines
        def iter_lines(self, decode_unicode=False):
            return iter(self._lines)
        def close(self):
            pass

    lines = [
        b'data: {"choices":[{"delta":{"content":"A"}}]}',
        b"",
        b"data: [DONE]",  # OpenAI-style sentinel (was previously unmatched)
        b'data: {"choices":[{"delta":{"content":"never"}}]}',  # must not arrive
    ]
    got = []
    content = streaming_mod._read_sse(FakeResponse(lines), on_token=got.append)
    assert content == "A"
