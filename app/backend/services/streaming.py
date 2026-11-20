from __future__ import annotations

import json
import os
from typing import Any, Callable, Dict, List, Optional, Tuple

import requests

from . import ai

def stream_chat_json(
    messages: List[Dict[str, str]],
    on_token: Optional[Callable[[str], None]] = None,
    temperature: float = 0.25,
    max_tokens: int = 4000,
) -> Tuple[Dict[str, Any], str]:
    """Call an OpenAI-compatible endpoint with streaming SSE, then parse JSON."""
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise ai.AIServiceError("未配置 DEEPSEEK_API_KEY")
    base = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com/v1").rstrip("/")
    model = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
    payload: Dict[str, Any] = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": True,
        "response_format": {"type": "json_object"},
    }
    last_error = "AI 调用失败"
    for attempt in range(2):
        body = dict(payload)
        if attempt == 1:
            body.pop("response_format", None)
        try:
            response = requests.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=body,
                stream=True,
                timeout=(15, 300),
            )
        except requests.RequestException as exc:
            last_error = f"AI 请求失败：{exc}"
            continue
        if response.status_code >= 400:
            text = response.text[:400]
            response.close()
            last_error = f"AI 返回 {response.status_code}：{text}"
            if response.status_code in (400, 422) and attempt == 0:
                continue
            break
        content_type = str(response.headers.get("content-type") or "").lower()
        try:
            if "text/event-stream" not in content_type:
                data = response.json()
                content = data["choices"][0]["message"]["content"]
                if on_token and content:
                    on_token(str(content))
                return _parse(content), str(content)
            content = _read_sse(response, on_token)
        finally:
            response.close()
        if not content.strip():
            last_error = "AI 流式响应为空"
            if attempt == 0:
                continue
            break
        try:
            return _parse(content), content
        except ai.AIServiceError as exc:
            last_error = str(exc)
            if attempt == 0:
                continue
    raise ai.AIServiceError(last_error)

def _parse(content):
    try:
        parsed = json.loads(ai._strip_json(content))
    except json.JSONDecodeError as exc:
        raise ai.AIServiceError(f"AI 返回的 JSON 无法解析：{exc}") from exc
    if not isinstance(parsed, dict):
        raise ai.AIServiceError("AI 必须返回 JSON 对象")
    return parsed
def _read_sse(response, on_token):
    chunks = []
    for raw in response.iter_lines(decode_unicode=False):
        if not raw:
            continue
        line = raw.decode("utf-8", "replace").strip()
        if not line.startswith("data:"):
            continue
        data = line[5:].strip()
        if data == "DONE":
            break
        try:
            obj = json.loads(data)
            choice = (obj.get("choices") or [{}])[0]
            piece = (choice.get("delta") or {}).get("content")
        except Exception:
            continue
        if piece:
            chunks.append(piece)
            if on_token:
                on_token(piece)
    return "".join(chunks)

