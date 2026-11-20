from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests


class AIServiceError(RuntimeError):
    pass


MAX_PATCHES = 40
CHUNK_CHARS = 14000


def _clean_text(value: str, limit: int = 700) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _region(obj: Dict[str, Any]) -> str:
    try:
        bbox = obj.get("bbox_emu") or {}
        size = obj.get("slide_size_emu") or {}
        height = max(1, int(size.get("height") or 1))
        y = int(bbox.get("y") or 0) / height
        h = int(bbox.get("h") or 0) / height
        font = float(obj.get("font_size_pt") or 0)
        if y + h / 2 < 0.2 and font >= 22:
            return "标题区"
        if y > 0.86 or (font and font < 11 and h < 0.06):
            return "页脚"
        return "正文"
    except Exception:
        return "正文"


def _object_line(obj: Dict[str, Any]) -> str:
    text = _clean_text(obj.get("text", ""))
    font = f"{float(obj['font_size_pt']):g}pt" if obj.get("font_size_pt") else "-"
    alignment = str(obj.get("alignment") or "-")
    parts = [
        str(obj.get("slide_index")),
        str(obj.get("shape_id")),
        _region(obj),
        font,
        alignment,
    ]
    return " | ".join(parts) + " | " + text


def _chunks(objects: List[Dict[str, Any]]) -> Iterable[List[Dict[str, Any]]]:
    current: List[Dict[str, Any]] = []
    current_chars = 0
    for obj in objects:
        line_len = len(_object_line(obj)) + 1
        if current and current_chars + line_len > CHUNK_CHARS:
            yield current
            current = []
            current_chars = 0
        current.append(obj)
        current_chars += line_len
    if current:
        yield current


def _system_prompt() -> str:
    return (
        "你是一位资深中文演示文稿（PPT）文案编辑。你只输出一个 JSON 对象，不要输出 Markdown 或解释。\n"
        "工作规则：\n"
        "1. 只修改用户点名的范围；用户没有点名时，优先改明显啰嗦、口语化、逻辑不清或与页面主题不符的句子。\n"
        "2. 保持原意与事实，不得编造数据、案例、人名、机构或结论。\n"
        "3. 保留原文中的数字、单位、专有名词、产品名、英文缩写与关键标签。\n"
        "4. 不改变版式：不新增文本框、不移动文本框、不改变字体字号；只改写已有文字或删除纯冗余文本。\n"
        "5. 除非用户明确要求，不要修改页码、页脚、版权声明、日期、目录序号。\n"
        "6. 标题要短而具体；正文要书面、简洁、可直接放进 PPT；避免排比堆砌和空话。\n"
        "7. 如果原文已经足够好，不要返回这一条；宁缺毋滥。\n"
        "8. before 必须与给出的原文完全一致（含标点与空格）；after 是改写后的完整文本。\n"
        "9. 删除冗余文本框时 action 用 delete_textbox，after 用空字符串。\n"
        "返回格式：\n"
        '{"patches":[{"slide_index":0,"shape_id":"s0_id2","before":"原文",'
        '"after":"改写后文本","description":"一句话说明改了什么","action":"edit"}]}\n'
        "最多返回 8 条，按重要性排序。没有建议时返回 {\"patches\":[]}。"
    )


def _user_prompt(
    objects: List[Dict[str, Any]], instruction: str, total_slides: int
) -> str:
    slides = sorted({int(o["slide_index"]) for o in objects})
    scope = (
        f"第 {slides[0] + 1} 页"
        if len(slides) == 1
        else f"第 {slides[0] + 1}-{slides[-1] + 1} 页"
    )
    lines = [
        f"演示文稿共 {total_slides} 页。本次只处理：{scope}。",
        "对象格式：slide_index | shape_id | 区域 | 字号 | 对齐 | 原文",
        "（slide_index 从 0 开始，必须原样返回；shape_id 必须原样返回。）",
        "",
    ]
    lines.extend(_object_line(obj) for obj in objects)
    lines.extend(
        [
            "",
            f"用户要求：{instruction.strip()}",
            "",
            "请返回 JSON。只对确实需要修改的对象给出 patch；不要返回未列出的 shape_id。",
        ]
    )
    return "\n".join(lines)


def _strip_json(raw: str) -> str:
    text = (raw or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        return text[start : end + 1]
    return text


def _call_deepseek(
    messages: List[Dict[str, str]], temperature: float = 0.2, max_tokens: int = 2500
) -> Dict[str, Any]:
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise AIServiceError("未配置 DEEPSEEK_API_KEY")
    base = os.getenv("DEEPSEEK_API_BASE", "https://api.deepseek.com/v1").rstrip("/")
    model = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
    payload: Dict[str, Any] = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "response_format": {"type": "json_object"},
    }
    last_error: Optional[str] = None
    for attempt in range(2):
        body = dict(payload)
        if attempt == 1:
            body.pop("response_format", None)
        try:
            response = requests.post(
                f"{base}/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                },
                json=body,
                timeout=120,
            )
        except requests.RequestException as exc:
            last_error = f"AI 请求失败：{exc}"
            continue
        if response.status_code >= 400:
            last_error = f"AI 返回 {response.status_code}：{response.text[:300]}"
            if response.status_code in (400, 422) and attempt == 0:
                continue
            break
        try:
            return response.json()
        except ValueError:
            last_error = "AI 返回的不是有效 JSON"
    raise AIServiceError(last_error or "AI 调用失败")


def request_json(
    messages: List[Dict[str, str]], temperature: float = 0.2, max_tokens: int = 3000
) -> Dict[str, Any]:
    data = _call_deepseek(messages, temperature=temperature, max_tokens=max_tokens)
    try:
        content = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise AIServiceError("AI 响应结构异常") from exc
    try:
        parsed = json.loads(_strip_json(content))
    except json.JSONDecodeError as exc:
        raise AIServiceError(f"AI 返回的 JSON 无法解析：{exc}") from exc
    if not isinstance(parsed, dict):
        raise AIServiceError("AI 必须返回 JSON 对象")
    return parsed


def locate_requested_slide(instruction: str, slide_count: int) -> Tuple[str, Optional[int]]:
    """Classify an explicit page mention in the instruction.

    Returns one of:
      ("ok", index)          — a valid, in-range page was requested (0-based)
      ("out_of_range", n)    — a page number parsed but exceeds the deck (1-based n)
      ("none", None)         — no explicit page mention
    Distinguishing out-of-range from absent prevents silently editing whatever
    page happens to be open when the user names a page that doesn't exist.
    """
    text = (instruction or "").strip()
    if not text:
        return "none", None
    lower = text.lower()
    if "首页" in text or "第一页" in text or "first page" in lower:
        return ("ok", 0) if slide_count > 0 else ("none", None)

    def _cn_number(token: str) -> Optional[int]:
        token = token.strip()
        if token.isdigit():
            return int(token)
        digits = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
        if token in digits:
            return digits[token]
        if token == "十":
            return 10
        if "十" in token:
            left, _, right = token.partition("十")
            tens = digits.get(left, 1) if left else 1
            ones = digits.get(right, 0) if right else 0
            if tens is None or ones is None:
                return None
            return tens * 10 + ones
        return None

    match = re.search(r"第\s*([0-9零一二两三四五六七八九十]+)\s*[页張张]", text)
    if match:
        number = _cn_number(match.group(1))
        if number is not None:
            index = number - 1
            if 0 <= index < slide_count:
                return "ok", index
            return "out_of_range", number
    match = re.search(r"\b(?:page|slide)\s*([0-9]+)\b", lower)
    if match:
        number = int(match.group(1))
        index = number - 1
        if 0 <= index < slide_count:
            return "ok", index
        return "out_of_range", number
    return "none", None


def extract_requested_slide(instruction: str, slide_count: int) -> Optional[int]:
    state_, index = locate_requested_slide(instruction, slide_count)
    return index if state_ == "ok" else None


def _validate(
    raw_patches: Any,
    objects: List[Dict[str, Any]],
    limit: int,
    instruction: str,
) -> Tuple[List[Dict[str, Any]], int]:
    if not isinstance(raw_patches, list):
        return [], 0
    by_id = {str(obj["shape_id"]): obj for obj in objects}
    by_slide_text: Dict[Tuple[int, str], List[Dict[str, Any]]] = {}
    for obj in objects:
        key = (int(obj["slide_index"]), re.sub(r"\s+", "", str(obj.get("text") or "")))
        by_slide_text.setdefault(key, []).append(obj)

    result: List[Dict[str, Any]] = []
    seen = set()
    dropped = 0
    for item in raw_patches:
        if not isinstance(item, dict):
            dropped += 1
            continue
        action = str(item.get("action") or "edit").strip().lower()
        if action in ("rewrite", "replace", "update", ""):
            action = "edit"
        if action not in ("edit", "delete_textbox"):
            dropped += 1
            continue
        shape_id = str(item.get("shape_id") or "").strip()
        obj = by_id.get(shape_id)
        if obj is None:
            try:
                slide_index = int(item.get("slide_index"))
            except (TypeError, ValueError):
                dropped += 1
                continue
            key = (slide_index, re.sub(r"\s+", "", str(item.get("before") or "")))
            candidates = by_slide_text.get(key) or []
            if len(candidates) == 1:
                obj = candidates[0]
                shape_id = str(obj["shape_id"])
        if obj is None or shape_id in seen:
            dropped += 1
            continue
        before = str(obj.get("text") or "")
        allow_footer = bool(
            re.search(r"页脚|页码|footer|日期|版权|copyright", instruction, re.IGNORECASE)
        )
        if not allow_footer and (
            re.fullmatch(r"\d{1,4}", before.strip()) or _region(obj) == "页脚"
        ):
            dropped += 1
            continue
        after = item.get("after")
        after = "" if after is None else str(after)
        if action == "edit":
            if not after.strip():
                dropped += 1
                continue
            if re.sub(r"\s+", "", after) == re.sub(r"\s+", "", before):
                dropped += 1
                continue
            if len(after) > max(1200, len(before) * 4):
                dropped += 1
                continue
        seen.add(shape_id)
        result.append(
            {
                "id": f"p{len(result) + 1}",
                "slide_index": int(obj["slide_index"]),
                "shape_id": shape_id,
                "shape_index": obj.get("shape_index"),
                "shape_name": obj.get("shape_name"),
                "action": action,
                "before": before,
                "after": after,
                "description": _clean_text(str(item.get("description") or ""), 200)
                or ("删除冗余文本框" if action == "delete_textbox" else "优化这段文案"),
                "bbox_emu": obj.get("bbox_emu"),
                "font_size_pt": obj.get("font_size_pt"),
                "alignment": obj.get("alignment"),
            }
        )
        if len(result) >= limit:
            break
    return result, dropped


def validate_patches(
    raw_patches: Any,
    objects: List[Dict[str, Any]],
    limit: int,
    instruction: str,
) -> Tuple[List[Dict[str, Any]], int]:
    return _validate(raw_patches, objects, limit, instruction)


def object_line(obj: Dict[str, Any]) -> str:
    return _object_line(obj)


def region_of(obj: Dict[str, Any]) -> str:
    return _region(obj)


def clean_text(value: Any, limit: int = 700) -> str:
    return _clean_text(value, limit)


def suggest(
    objects: List[Dict[str, Any]],
    instruction: str,
    total_slides: int,
    slide_indices: Optional[List[int]] = None,
) -> Dict[str, Any]:
    instruction = (instruction or "").strip()
    if not instruction:
        raise AIServiceError("请先输入修改要求")
    if len(instruction) > 2000:
        instruction = instruction[:2000]
    if slide_indices:
        wanted = set(int(i) for i in slide_indices)
        objects = [obj for obj in objects if int(obj["slide_index"]) in wanted]
    if not objects:
        raise AIServiceError("当前范围没有可编辑的文字")

    all_patches: List[Dict[str, Any]] = []
    dropped = 0
    calls = 0
    for chunk in _chunks(objects):
        messages = [
            {"role": "system", "content": _system_prompt()},
            {"role": "user", "content": _user_prompt(chunk, instruction, total_slides)},
        ]
        data = _call_deepseek(messages)
        calls += 1
        try:
            content = data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError):
            raise AIServiceError("AI 响应结构异常")
        try:
            parsed = json.loads(_strip_json(content))
        except json.JSONDecodeError as exc:
            raise AIServiceError(f"AI 返回的 JSON 无法解析：{exc}") from exc
        raw = parsed.get("patches") if isinstance(parsed, dict) else parsed
        patches, chunk_dropped = _validate(
            raw, chunk, limit=MAX_PATCHES - len(all_patches), instruction=instruction
        )
        all_patches.extend(patches)
        dropped += chunk_dropped
        if len(all_patches) >= MAX_PATCHES:
            break

    for index, patch in enumerate(all_patches, start=1):
        patch["id"] = f"p{index}"
    return {
        "patches": all_patches,
        "dropped": dropped,
        "calls": calls,
        "model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
    }
