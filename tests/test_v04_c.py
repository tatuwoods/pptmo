"""v0.4 C 组：协作与安全 — 只读分享 / 密码保护 / 审计日志 / 宏检测。"""
from __future__ import annotations

import io
import zipfile

import pytest


# ---------- 13. 只读分享链接 ----------

def test_share_lifecycle(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/share")
    assert r.status_code == 200, r.text
    share = r.json()
    token = share["token"]
    assert share["share_url"].endswith(token)

    # 只读页可访问
    page = client.get(f"/share/{token}")
    assert page.status_code == 200
    assert page.headers["content-type"].startswith("text/html")

    data = client.get(f"/share/{token}/data").json()
    assert data["slide_count"] == 3
    assert len(data["pages"]) == 3
    assert "objects" not in data  # 只读：不泄漏可编辑对象清单

    img = client.get(data["pages"][0]["thumb"])
    assert img.status_code == 200
    assert img.headers["content-type"] == "image/png"

    dl = client.get(f"/share/{token}/download")
    assert dl.status_code == 200
    assert dl.content[:2] == b"PK"  # pptx zip

    # 列表 + 撤销
    shares = client.get(f"/api/docs/{doc_id}/shares").json()["shares"]
    assert any(s["token"] == token and not s["revoked"] for s in shares)
    assert client.delete(f"/api/docs/{doc_id}/share/{token}").status_code == 200
    assert client.get(f"/share/{token}").status_code == 404
    assert client.get(f"/share/{token}/data").status_code == 404
    assert client.get(f"/share/{token}/download").status_code == 404


def test_share_reflects_latest_version(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    token = client.post(f"/api/docs/{doc_id}/share").json()["token"]
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "delete_slide", "slide": 2}]},
    )
    assert r.status_code == 200
    data = client.get(f"/share/{token}/data").json()
    assert data["slide_count"] == 2


def test_share_unknown_token_404(client, uploaded_doc):
    assert client.get("/share/nonexistenttoken123").status_code == 404


# ---------- 14. 文档密码保护 ----------

def test_password_lock_and_unlock(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert client.post(f"/api/docs/{doc_id}/password", json={"password": "s3cret"}).status_code == 200

    # 锁定后未带密码 → 401
    assert client.get(f"/api/docs/{doc_id}").status_code == 401
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "x"}]},
    )
    assert r.status_code == 401

    # unlock 校验
    assert client.post(f"/api/docs/{doc_id}/unlock", json={"password": "wrong"}).status_code == 401
    assert client.post(f"/api/docs/{doc_id}/unlock", json={"password": "s3cret"}).status_code == 200

    # 带密码头可访问、可编辑
    headers = {"X-Doc-Password": "s3cret"}
    assert client.get(f"/api/docs/{doc_id}", headers=headers).status_code == 200
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "带密码可编辑"}]},
        headers=headers,
    )
    assert r.status_code == 200, r.text

    # meta 里只有 hash，没有明文
    import json as _json

    import storage

    meta_raw = storage.meta_path(doc_id).read_text(encoding="utf-8")
    assert "s3cret" not in meta_raw
    meta = _json.loads(meta_raw)
    assert meta.get("password_hash")

    # 清除密码 → 开放
    assert client.post(f"/api/docs/{doc_id}/password", json={"password": None}, headers=headers).status_code == 200
    assert client.get(f"/api/docs/{doc_id}").status_code == 200


def test_password_doc_still_listed_and_shareable(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    client.post(f"/api/docs/{doc_id}/password", json={"password": "pw123"})
    docs = client.get("/api/docs").json()["docs"]
    assert any(d["doc_id"] == doc_id for d in docs)
    # 分享链接自带凭据：无需文档密码即可看
    token = client.post(
        f"/api/docs/{doc_id}/share", headers={"X-Doc-Password": "pw123"}
    ).json()["token"]
    assert client.get(f"/share/{token}/data").status_code == 200


# ---------- 15. 操作审计日志 ----------

def test_audit_logs_applies(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "审计测试"}],
              "label": "备注修改"},
    )
    assert r.status_code == 200
    entries = client.get("/api/audit").json()["entries"]
    assert entries, "audit.jsonl 应有记录"
    entry = entries[0]
    assert entry["doc_id"] == doc_id
    assert entry["op_count"] == 1
    assert entry["ok"] is True
    assert "备注修改" in entry["summary"]
    assert entry.get("at") and entry.get("action") == "apply"

    # 失败的 apply 也留痕
    client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit", "shape_id": "s9_id9", "after": "x"}]},
    )
    entries = client.get("/api/audit").json()["entries"]
    failed = [e for e in entries if e.get("ok") is False]
    assert failed and failed[0]["op_count"] == 1


def test_audit_limit_and_doc_filter(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    for i in range(3):
        client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": f"n{i}"}]},
        )
    limited = client.get("/api/audit?limit=2").json()["entries"]
    assert len(limited) == 2
    filtered = client.get(f"/api/audit?doc_id={doc_id}").json()["entries"]
    assert filtered and all(e["doc_id"] == doc_id for e in filtered)
    fresh_first = [e["at"] for e in filtered]
    assert fresh_first == sorted(fresh_first, reverse=True)  # 新的在前


def test_undo_and_restore_are_audited(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "v1"}]},
    )
    assert client.post(f"/api/docs/{doc_id}/undo").status_code == 200
    assert client.post(f"/api/docs/{doc_id}/versions/v0/restore").status_code == 200
    actions = [e["action"] for e in client.get("/api/audit").json()["entries"]]
    assert "undo" in actions
    assert "restore" in actions


# ---------- 16. 上传宏检测（.pptx 内嵌 vbaProject） ----------

def _macro_pptx_bytes(demo_pptx) -> bytes:
    with zipfile.ZipFile(demo_pptx) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}
    items["ppt/vbaProject.bin"] = b"\x00\x01\x02MACRO-STUB"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    return buf.getvalue()


def test_macro_upload_becomes_readonly(client, demo_pptx):
    r = client.post(
        "/api/docs/upload",
        files={"file": ("macro.pptx", _macro_pptx_bytes(demo_pptx),
                        "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
    )
    assert r.status_code == 200, r.text
    doc = r.json()
    assert doc["has_macro"] is True
    assert "只读" in doc["macro_warning"]

    doc_id = doc["doc_id"]
    detail = client.get(f"/api/docs/{doc_id}").json()
    assert detail["has_macro"] is True

    # 拒绝编辑
    r = client.post(
        f"/api/docs/{doc_id}/apply",
        json={"patches": [{"action": "edit_notes", "slide_index": 0, "text": "x"}]},
    )
    assert r.status_code == 400
    assert "宏" in r.json()["detail"]

    # 只读能力保留：下载 / 分享 / 渲染
    assert client.get(f"/api/docs/{doc_id}/download").status_code == 200
    token = client.post(f"/api/docs/{doc_id}/share").json()["token"]
    assert client.get(f"/share/{token}/data").status_code == 200


def test_clean_upload_has_no_macro_flag(client, uploaded_doc):
    assert uploaded_doc.get("has_macro") is not True
