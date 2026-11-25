"""v0.6 #1：渲染验证闭环 —— 把渲染 PNG 喂给多模态模型自检。

对标 OfficeCLI 的 render→look→fix 与 ppt-master 的「实测」哲学：apply 之后
不能只信几何启发式 lint，还要让多模态模型真正「看」一眼渲染结果——文字是否
溢出、符号是否乱码（豆腐块）、布局是否塌。发现问题以结构化 issues 返回，
Agent 拿到后可自动提案修复（fix_issues / edit / style），形成闭环。

复用 DEEPSEEK_* 环境配置（任意 OpenAI 兼容端点）；模型需支持视觉输入，
可用 VISION_MODEL 单独指定（默认沿用 DEEPSEEK_MODEL）。
"""
from __future__ import annotations

import base64
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

from .ai import AIServiceError

# 每页最多回报的问题条数（防模型絮叨撑爆 observation）
MAX_ISSUES_PER_PAGE = 8
# 单次端点调用最多自检页数（token/时长护栏）
MAX_PAGES_PER_CALL = 6
# 超过该像素宽度先降采样再上传（请求体护栏，base64 后约 ×1.37）
MAX_IMAGE_WIDTH_PX = 1800

ISSUE_KINDS = ("overflow", "garbled", "layout", "contrast", "other")
_KIND_ALIASES = {
    "text_overflow": "overflow",
    "溢出": "overflow",
    "乱码": "garbled",
    "tofu": "garbled",
    "塌版": "layout",
    "布局": "layout",
    "对比度": "contrast",
}

_SYSTEM_PROMPT = (
    "你是 PPT 渲染质检员。你会收到一页幻灯片的渲染图，请像设计师审稿一样检查：\n"
    "1. 文字是否溢出文本框/画布边缘、是否被截断；\n"
    "2. 是否有乱码、豆腐块（□□）、错位符号；\n"
    "3. 布局是否塌陷（元素重叠成一团、明显错位、大面积空白失衡）；\n"
    "4. 文字与背景对比度是否低到难以阅读。\n"
    "只输出一个 JSON 对象（不要 Markdown）：\n"
    '{"ok": true/false, "issues": [{"kind": "overflow|garbled|layout|contrast|other",'
    '"description": "具体在哪、什么问题", "suggestion": "设计师会怎么改（一句话）"}]}\n'
    "没有问题就 ok=true 且 issues 为空数组。不要报无关紧要的审美偏好。"
)


class VisionError(AIServiceError):
    """视觉自检失败（未配置 key / 模型不支持视觉 / 请求失败）。"""


def _model() -> str:
    return (os.getenv("VISION_MODEL") or os.getenv("DEEPSEEK_MODEL", "deepseek-chat")).strip()


def _strip_json(raw: str) -> str:
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start, end = text.find("{"), text.rfind("}")
    return text[start : end + 1] if start >= 0 and end > start else text


def _shrink_png(data: bytes) -> bytes:
    """超宽 PNG 先降采样（base64 前的请求体护栏）；失败原样返回。"""
    if len(data) <= 2 * 1024 * 1024:
        return data
    try:
        import pymupdf

        src = pymupdf.Pixmap(data)
        if src.width <= MAX_IMAGE_WIDTH_PX:
            return data
        zoom = MAX_IMAGE_WIDTH_PX / src.width
        out = pymupdf.Pixmap(src, pymupdf.Matrix(zoom, zoom)) if src.n < 4 else pymupdf.Pixmap(
            src, pymupdf.Matrix(zoom, zoom), src.alpha
        )
        return out.tobytes("png")
    except Exception:
        return data


def build_check_messages(png_bytes: bytes, page_number: int, context: Optional[str] = None) -> List[Dict[str, Any]]:
    """PNG → OpenAI 视觉消息（base64 data URL content parts）。"""
    payload = _shrink_png(png_bytes)
    text = (
        f"这是渲染后的第 {page_number} 页幻灯片（渲染器为 LibreOffice，字体可能与 PowerPoint 略有差异，"
        "这不算问题）。请按系统要求只输出 JSON。"
    )
    if context:
        text += f"\n本轮改动背景：{context[:400]}"
    return [
        {"role": "system", "content": _SYSTEM_PROMPT},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": text},
                {
                    "type": "image_url",
                    "image_url": {"url": "data:image/png;base64," + base64.b64encode(payload).decode("ascii")},
                },
            ],
        },
    ]


def _post_chat(messages: List[Dict[str, Any]], max_tokens: int = 1200) -> Dict[str, Any]:
    """调多模态 chat/completions；与 ai.py 同款容错（response_format 被拒时重试一次）。"""
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise VisionError("未配置多模态模型 API key（DEEPSEEK_API_KEY）")
    base = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com/v1").rstrip("/")
    body: Dict[str, Any] = {
        "model": _model(),
        "messages": messages,
        "temperature": 0.1,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
    }
    last_error: Optional[str] = None
    for attempt in range(2):
        payload = dict(body)
        if attempt == 1:
            payload.pop("response_format", None)
        try:
            response = requests.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=payload,
                timeout=180,
            )
        except requests.RequestException as exc:
            last_error = f"视觉自检请求失败：{exc}"
            continue
        if response.status_code >= 400:
            last_error = f"视觉自检模型返回 {response.status_code}：{response.text[:300]}"
            if response.status_code in (400, 422) and attempt == 0:
                continue
            break
        try:
            data = response.json()
            content = data["choices"][0]["message"]["content"]
            return json.loads(_strip_json(content))
        except (ValueError, KeyError, IndexError, TypeError) as exc:
            last_error = f"视觉自检响应无法解析：{exc}"
    raise VisionError(last_error or "视觉自检调用失败")


def _normalize_kind(raw: Any) -> str:
    text = str(raw or "").strip().lower()
    if text in ISSUE_KINDS:
        return text
    return _KIND_ALIASES.get(text, "other")


def check_page_png(
    png_path: Path, page_number: int, context: Optional[str] = None
) -> Dict[str, Any]:
    """自检一页渲染图；返回 {page, ok, issues[], model}。"""
    png_path = Path(png_path)
    if not png_path.exists():
        raise FileNotFoundError(f"render page not found: {png_path.name}")
    data = _post_chat(build_check_messages(png_path.read_bytes(), page_number, context))
    raw_issues = data.get("issues") if isinstance(data, dict) else None
    issues: List[Dict[str, Any]] = []
    if isinstance(raw_issues, list):
        for item in raw_issues[:MAX_ISSUES_PER_PAGE]:
            if not isinstance(item, dict):
                continue
            description = str(item.get("description") or "").strip()
            if not description:
                continue
            issues.append(
                {
                    "kind": _normalize_kind(item.get("kind")),
                    "description": description[:300],
                    "suggestion": str(item.get("suggestion") or "").strip()[:300],
                }
            )
    return {
        "page": int(page_number),
        "ok": not issues,
        "issues": issues,
        "model": _model(),
    }


def check_pages(
    png_paths: List[Path], start_page: int = 1, context: Optional[str] = None
) -> List[Dict[str, Any]]:
    """顺序自检多页（每页一次模型调用）。"""
    results = []
    for offset, path in enumerate(png_paths):
        results.append(check_page_png(path, start_page + offset, context))
    return results


def summarize(pages: List[Dict[str, Any]]) -> str:
    total = sum(len(p.get("issues") or []) for p in pages)
    if not total:
        return f"已自检 {len(pages)} 页渲染图，未发现问题"
    by_kind: Dict[str, int] = {}
    for page in pages:
        for issue in page.get("issues") or []:
            by_kind[issue.get("kind", "other")] = by_kind.get(issue.get("kind", "other"), 0) + 1
    parts = "、".join(f"{kind}×{count}" for kind, count in sorted(by_kind.items()))
    bad = [str(p["page"]) for p in pages if p.get("issues")]
    return f"已自检 {len(pages)} 页，发现问题 {total} 处（{parts}）：第 {'、'.join(bad)} 页"
