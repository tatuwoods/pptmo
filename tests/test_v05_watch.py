"""v0.5 #9: watch 实时预览 — SSE 事件推送 + 浏览器 watch 页。

对标 OfficeCLI watch：文档每次版本变更（apply/chat/undo/restore/duplicate）
向订阅者推送 changed 事件，watch 页据此自动刷新页面渲染图。
"""
from __future__ import annotations

import queue
import threading

import pytest


def _patch_for(payload, after="改动后的标题"):
    obj = next(o for o in payload["objects"] if o["kind"] == "text" and o["text"])
    return {
        "patches": [{"action": "edit", "shape_id": obj["shape_id"], "after": after, "before": obj["text"]}],
        "label": "v05-9 watch",
    }


# ---------- broker 单元 ----------

def test_broker_subscribe_publish():
    from services import broker

    q = broker.subscribe("doc1")
    try:
        sent = broker.publish("doc1", {"type": "changed", "version": 1})
        assert sent == 1
        assert q.get(timeout=1)["type"] == "changed"
    finally:
        broker.unsubscribe("doc1", q)


def test_broker_unsubscribe_stops_delivery():
    from services import broker

    q = broker.subscribe("doc1")
    broker.unsubscribe("doc1", q)
    assert broker.publish("doc1", {"type": "changed", "version": 2}) == 0


def test_broker_slow_subscriber_does_not_block_publisher():
    """队列有界：塞满后发布方不阻塞、不崩溃（丢最旧保最新）。"""
    from services import broker

    q = broker.subscribe("doc2", maxsize=2)
    try:
        for n in range(5):
            broker.publish("doc2", {"type": "changed", "version": n})
        assert q.qsize() == 2
        # 最旧的两条被挤掉，剩最后两条
        assert q.get_nowait()["version"] == 3
        assert q.get_nowait()["version"] == 4
    finally:
        broker.unsubscribe("doc2", q)


def test_broker_isolated_per_doc():
    from services import broker

    q = broker.subscribe("docA")
    try:
        assert broker.publish("docB", {"type": "changed"}) == 0
        assert q.empty()
    finally:
        broker.unsubscribe("docA", q)


# ---------- API：/events SSE ----------

def test_events_hello_and_changed(client, uploaded_doc):
    from services import broker

    doc_id = uploaded_doc["doc_id"]
    # 手动发布一条 changed 事件：必须在流读取过程中发生（异步定时器），
    # 同步 publish 在 _read_events 返回后执行是收不到的
    publisher = threading.Timer(
        0.5,
        lambda: broker.publish(
            doc_id, {"type": "changed", "version": 1, "version_id": "v1", "touched_slides": [0]}
        ),
    )
    publisher.start()
    try:
        # max_events=2：hello+changed 后服务端发 bye 收尾（无限流在 TestClient 下无法断开）
        with client.stream("GET", f"/api/docs/{doc_id}/events?max_events=2") as resp:
            assert resp.status_code == 200
            assert "text/event-stream" in resp.headers["content-type"]
            events = _read_events(resp.iter_lines(), count=3, timeout=8.0)
    finally:
        publisher.cancel()
    assert events[0]["event"] == "hello"
    assert events[0]["data"]["doc_id"] == doc_id
    assert events[0]["data"]["version"] == 0
    assert events[1]["event"] == "changed"
    assert events[1]["data"]["version"] == 1
    assert events[2]["event"] == "bye"


def test_events_unknown_doc_404(client):
    assert client.get("/api/docs/" + "f" * 32 + "/events").status_code == 404


def test_events_gated_by_password(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    client.post(f"/api/docs/{doc_id}/password", json={"password": "pw123"})
    resp = client.get(f"/api/docs/{doc_id}/events")
    assert resp.status_code == 401
    # SSE 是无限流：用 stream 消费 + max_events 让流自然收尾
    with client.stream(
        "GET",
        f"/api/docs/{doc_id}/events?max_events=1",
        headers={"X-Doc-Password": "pw123"},
    ) as resp:
        assert resp.status_code == 200
        assert "text/event-stream" in resp.headers["content-type"]


def test_apply_publishes_changed_event(client, uploaded_doc):
    """apply 提交新版本后向订阅者推送 changed（版本号/触碰页/页数）。"""
    from services import broker

    doc_id = uploaded_doc["doc_id"]
    q = broker.subscribe(doc_id)
    try:
        resp = client.post(f"/api/docs/{doc_id}/apply", json=_patch_for(uploaded_doc))
        assert resp.status_code == 200, resp.text
        event = q.get(timeout=5)
        assert event["type"] == "changed"
        assert event["version"] == 1
        assert event["version_id"] == "v1"
        assert event["touched_slides"] == [0]
        assert event["slide_count"] == 3
    finally:
        broker.unsubscribe(doc_id, q)


def test_undo_publishes_changed_event(client, uploaded_doc):
    from services import broker

    doc_id = uploaded_doc["doc_id"]
    client.post(f"/api/docs/{doc_id}/apply", json=_patch_for(uploaded_doc))
    q = broker.subscribe(doc_id)
    try:
        resp = client.post(f"/api/docs/{doc_id}/undo", json={"base_version": 1})
        assert resp.status_code == 200, resp.text
        event = q.get(timeout=5)
        assert event["type"] == "changed"
        assert event["version"] == 2
    finally:
        broker.unsubscribe(doc_id, q)


# ---------- API：/watch 预览页 ----------

def test_watch_page_served(client, uploaded_doc):
    resp = client.get(f"/api/docs/{uploaded_doc['doc_id']}/watch")
    assert resp.status_code == 200
    assert "text/html" in resp.headers["content-type"]
    assert "EventSource" in resp.text
    assert "/api/docs/" in resp.text


def test_watch_page_unknown_doc_404(client):
    assert client.get("/api/docs/" + "f" * 32 + "/watch").status_code == 404


# ---------- helpers ----------

def _read_events(lines, count: int, timeout: float):
    """从 SSE 行流解析 event/data 对，读到 count 条为止（带总超时）。"""
    import json as _json
    import time

    events = []
    deadline = time.monotonic() + timeout
    current = {}
    for line in lines:
        if time.monotonic() > deadline:
            break
        if line.startswith("event:"):
            current = {"event": line[6:].strip()}
        elif line.startswith("data:") and current.get("event") is not None:
            current["data"] = _json.loads(line[5:].strip())
            events.append(current)
            current = {}
        if len(events) >= count:
            break
    assert len(events) == count, f"expected {count} events, got {len(events)}: {events}"
    return events
