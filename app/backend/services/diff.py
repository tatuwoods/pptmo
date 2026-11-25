"""v0.5 #8：结构 diff 增强 — 对象级结构化 changelog（JSON 给 AI 消费）。

对标 hands-on-deck diff：在文字 diff 之外，按「页位对齐 + shape_id 匹配」
输出对象级变更：新增/删除了哪些对象、移动（origin 平移 ≥ 阈值）、
几何变化（x/y/w/h 哪个轴）、样式变化（字号/字体/加粗/颜色）。

对象取顶层叶子（grouped=False 且非 group 容器）——表格/组合降维到容器级，
不把 table_cell / smartart_node 子对象灌进 changelog。页间移动的对象无法
可靠关联（shape_id 含页号），如实报告为 removed + added。
"""
from __future__ import annotations

from typing import Any, Dict, List

from . import edits

# origin 平移小于该值不算移动（1pt=12700；2pt 内视为排版抖动）
MOVE_THRESHOLD_EMU = 25400
_TEXT_SNIPPET = 120
_MAX_PER_SECTION = 50

_SECTIONS = ("added", "removed", "moved", "text_changed", "geometry_changed", "style_changed")
# parse._font_info 的源字段 → changelog 对外字段名（font_size_pt → font_size）
_STYLE_FIELDS = {"font_name": "font_name", "font_size_pt": "font_size", "bold": "bold", "color": "color"}


def _top_level_objects(payload: Dict[str, Any]) -> Dict[int, Dict[str, Dict[str, Any]]]:
    """payload.objects + payload.elements 按 (页, shape_id) 归并；只留顶层叶子。"""
    by_page: Dict[int, Dict[str, Dict[str, Any]]] = {}
    merged: Dict[str, Dict[str, Any]] = {}
    for obj in list(payload.get("objects") or []) + list(payload.get("elements") or []):
        shape_id = str(obj.get("shape_id") or "")
        if not shape_id:
            continue
        if obj.get("grouped") or obj.get("depth"):
            continue
        kind = str(obj.get("kind") or "")
        if kind == "group":
            continue
        merged.setdefault(shape_id, obj)
    for shape_id, obj in merged.items():
        try:
            page = int(obj.get("slide_index"))
        except (TypeError, ValueError):
            continue
        by_page.setdefault(page, {})[shape_id] = obj
    return by_page


def _snippet(text: Any) -> str:
    return str(text or "")[:_TEXT_SNIPPET]


def _norm_style(value: Any) -> Any:
    if isinstance(value, float):
        return int(value)
    if isinstance(value, str) and value.startswith("#"):
        return value[1:].upper()
    if isinstance(value, str) and len(value) == 6:
        return value.upper()
    return value


def _style_changes(a: Dict[str, Any], b: Dict[str, Any]) -> List[Dict[str, Any]]:
    changes = []
    for src, field in _STYLE_FIELDS.items():
        before, after = _norm_style(a.get(src)), _norm_style(b.get(src))
        if before != after:
            changes.append({"field": field, "before": before, "after": after})
    return changes


def _bbox_axes(a: Dict[str, Any], b: Dict[str, Any]) -> List[str]:
    if not isinstance(a, dict) or not isinstance(b, dict):
        return []
    return [k for k in ("x", "y", "w", "h") if a.get(k) != b.get(k)]


def structural_changelog(
    payload_a: Dict[str, Any],
    payload_b: Dict[str, Any],
    *,
    move_threshold_emu: int = MOVE_THRESHOLD_EMU,
) -> Dict[str, Any]:
    """两版 payload 的对象级 changelog；只输出有变化的页。"""
    pages_a = _top_level_objects(payload_a)
    pages_b = _top_level_objects(payload_b)
    n_a = max(list(pages_a) + [-1]) + 1 if pages_a else 0
    n_b = max(list(pages_b) + [-1]) + 1 if pages_b else 0
    shared = min(n_a, n_b)

    slides: List[Dict[str, Any]] = []
    totals = {"added": 0, "removed": 0, "moved": 0, "text": 0, "geometry": 0, "style": 0}
    for page in range(max(n_a, n_b)):
        objs_a = pages_a.get(page, {})
        objs_b = pages_b.get(page, {})
        entry: Dict[str, Any] = {"page": page + 1}
        added = [objs_b[k] for k in sorted(set(objs_b) - set(objs_a))]
        removed = [objs_a[k] for k in sorted(set(objs_a) - set(objs_b))]
        moved, geometry, text_changed, style_changed = [], [], [], []
        for shape_id in sorted(set(objs_a) & set(objs_b)):
            obj_a, obj_b = objs_a[shape_id], objs_b[shape_id]
            axes = _bbox_axes(obj_a.get("bbox_emu"), obj_b.get("bbox_emu"))
            if axes:
                geometry.append(
                    {
                        "shape_id": shape_id,
                        "name": obj_b.get("name") or obj_a.get("name"),
                        "kind": obj_b.get("kind"),
                        "axes": axes,
                        "from": obj_a.get("bbox_emu"),
                        "to": obj_b.get("bbox_emu"),
                    }
                )
            if any(f"{axis}" in ("x", "y") for axis in axes):
                dx = int(obj_b["bbox_emu"].get("x") or 0) - int(obj_a["bbox_emu"].get("x") or 0)
                dy = int(obj_b["bbox_emu"].get("y") or 0) - int(obj_a["bbox_emu"].get("y") or 0)
                if abs(dx) >= move_threshold_emu or abs(dy) >= move_threshold_emu:
                    moved.append(
                        {
                            "shape_id": shape_id,
                            "name": obj_b.get("name") or obj_a.get("name"),
                            "kind": obj_b.get("kind"),
                            "from": obj_a.get("bbox_emu"),
                            "to": obj_b.get("bbox_emu"),
                        }
                    )
            text_a, text_b = obj_a.get("text"), obj_b.get("text")
            if text_a is not None and text_b is not None and text_a != text_b:
                if edits.normalize_text(str(text_a)) != edits.normalize_text(str(text_b)):
                    text_changed.append(
                        {
                            "shape_id": shape_id,
                            "name": obj_b.get("name") or obj_a.get("name"),
                            "before": _snippet(text_a),
                            "after": _snippet(text_b),
                        }
                    )
            changes = _style_changes(obj_a, obj_b)
            if changes:
                style_changed.append(
                    {"shape_id": shape_id, "name": obj_b.get("name") or obj_a.get("name"), "changes": changes}
                )
        if added:
            entry["added"] = [
                {"shape_id": o["shape_id"], "name": o.get("name"), "kind": o.get("kind"), "text": _snippet(o.get("text"))}
                for o in added[:_MAX_PER_SECTION]
            ]
        if removed:
            entry["removed"] = [
                {"shape_id": o["shape_id"], "name": o.get("name"), "kind": o.get("kind"), "text": _snippet(o.get("text"))}
                for o in removed[:_MAX_PER_SECTION]
            ]
        if moved:
            entry["moved"] = moved[:_MAX_PER_SECTION]
        if text_changed:
            entry["text_changed"] = text_changed[:_MAX_PER_SECTION]
        if geometry:
            entry["geometry_changed"] = geometry[:_MAX_PER_SECTION]
        if style_changed:
            entry["style_changed"] = style_changed[:_MAX_PER_SECTION]
        if len(entry) > 1:
            slides.append(entry)
        totals["added"] += len(added)
        totals["removed"] += len(removed)
        totals["moved"] += len(moved)
        totals["text"] += len(text_changed)
        totals["geometry"] += len(geometry)
        totals["style"] += len(style_changed)

    pages_added = max(0, n_b - shared)
    pages_removed = max(0, n_a - shared)
    changed_pages = len(slides)
    summary = {
        "pages_changed": changed_pages,
        "pages_added": pages_added,
        "pages_removed": pages_removed,
        "objects_added": totals["added"],
        "objects_removed": totals["removed"],
        "moved": totals["moved"],
        "text_changed": totals["text"],
        "geometry_changed": totals["geometry"],
        "style_changed": totals["style"],
    }
    parts = [
        f"对象增 {summary['objects_added']}/删 {summary['objects_removed']}",
        f"移动 {summary['moved']}",
        f"文字改 {summary['text_changed']}",
        f"几何改 {summary['geometry_changed']}",
        f"样式改 {summary['style_changed']}",
    ]
    summary["text"] = (
        f"{changed_pages} 页有对象级变更，页面增 {pages_added}/删 {pages_removed}；" + "、".join(parts)
    )
    return {"slides": slides, "summary": summary}
