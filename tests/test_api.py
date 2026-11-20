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
