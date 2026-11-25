"""v0.5 #9: 进程内文档事件广播（watch 实时预览的推送底座）。

每个订阅者一条有界队列：发布方永不阻塞（满则丢最旧保最新），
单文档订阅者归零时清理登记，避免长驻文档累积死队列。
"""
from __future__ import annotations

import queue
import threading
from typing import Any, Dict

_subs: Dict[str, set] = {}
_lock = threading.Lock()

DEFAULT_QUEUE_SIZE = 64


def subscribe(doc_id: str, maxsize: int = DEFAULT_QUEUE_SIZE) -> "queue.Queue":
    q: "queue.Queue" = queue.Queue(maxsize=max(1, int(maxsize)))
    with _lock:
        _subs.setdefault(doc_id, set()).add(q)
    return q


def unsubscribe(doc_id: str, q: "queue.Queue") -> None:
    with _lock:
        subscribers = _subs.get(doc_id)
        if subscribers is not None:
            subscribers.discard(q)
            if not subscribers:
                _subs.pop(doc_id, None)


def publish(doc_id: str, event: Dict[str, Any]) -> int:
    """向 doc_id 的全部订阅者推送事件；返回成功入队的订阅者数。

    队列满：丢最旧一条再入队（保最新状态，watch 页只需最后一次变更）。"""
    with _lock:
        subscribers = list(_subs.get(doc_id, ()))
    sent = 0
    for q in subscribers:
        try:
            q.put_nowait(event)
            sent += 1
        except queue.Full:
            try:
                q.get_nowait()
                q.put_nowait(event)
                sent += 1
            except (queue.Empty, queue.Full):
                pass
    return sent


def subscriber_count(doc_id: str) -> int:
    with _lock:
        return len(_subs.get(doc_id, ()))
