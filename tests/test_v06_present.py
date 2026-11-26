"""v0.6 #5：演讲者模式 + 排练计时（对标 guizang 第 1 名的核心体验）。

`GET /api/docs/{id}/present` 全屏演示页：渲染 PNG 直接放（保真 100%，
不用浏览器重排），←→/空格翻页、演讲者视图（当前页 + 下一页预览 + 演讲者
备注 + 排练计时器：总时长 + 本页时长、暂停/继续、重置）。
"""
from __future__ import annotations


def test_present_page_renders(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    resp = client.get(f"/api/docs/{doc_id}/present")
    assert resp.status_code == 200
    body = resp.text
    assert doc_id in body
    # 演示画面 + 键盘翻页
    assert "present" in body
    # 演讲者视图三要素：下一页预览、备注面板、计时器
    assert "next" in body
    assert "note" in body
    assert "timer" in body.lower()
    # 渲染 PNG 直连（保真 100%）
    assert "/slides/" in body and "/image" in body


def test_present_page_has_rehearsal_controls(client, uploaded_doc):
    body = client.get(f"/api/docs/{uploaded_doc['doc_id']}/present").text
    for marker in ("keydown", "Fullscreen", "toggleSpeaker"):
        assert marker in body


def test_present_page_missing_doc_404(client):
    resp = client.get("/api/docs/ffffffffffffffffffffffffffffffff/present")
    assert resp.status_code == 404


def test_present_page_password_protected(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    assert client.post(f"/api/docs/{doc_id}/password", json={"password": "pw123"}).status_code == 200
    locked = client.get(f"/api/docs/{doc_id}/present")
    assert locked.status_code == 401
    ok = client.get(f"/api/docs/{doc_id}/present", headers={"X-Doc-Password": "pw123"})
    assert ok.status_code == 200


def test_notes_available_for_speaker_view(client, uploaded_doc):
    """演讲者视图的数据源：/api/docs/{id} 序列化里带每页备注。"""
    doc = client.get(f"/api/docs/{uploaded_doc['doc_id']}").json()
    assert isinstance(doc.get("notes"), list) and doc["notes"]
    assert set(doc["notes"][0].keys()) >= {"slide_index", "text"}
    assert "讲稿" in doc["notes"][0]["text"]
