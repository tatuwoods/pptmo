from __future__ import annotations

import difflib
import json
import os
import re
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

from . import ai, edits, streaming

EventEmitter = Callable[[str, Dict[str, Any]], None]

EMU_PER_CM = 360000
MAX_OPS_PER_CALL = 6
MAX_TOOL_RESULT_CHARS = 12000
DEFAULT_SLIDE_WIDTH = 12192000
DEFAULT_SLIDE_HEIGHT = 6858000


def _env_int(name: str, default: int, low: int, high: int) -> int:
    try:
        value = int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        value = default
    return max(low, min(high, value))


MAX_STEPS = _env_int("AI_MAX_STEPS", 10, 2, 16)
MAX_OPS = _env_int("AI_MAX_OPS", 24, 4, 60)

STYLE_KEYS = ("font_size_pt", "font_name", "bold", "italic", "color", "alignment")
TEXT_ACTIONS = {"edit", "clear", "delete_textbox"}
ALIGNMENTS = {"LEFT", "CENTER", "RIGHT", "JUSTIFY", "DISTRIBUTE"}
ADD_ALIGNMENTS = {"LEFT", "CENTER", "RIGHT"}

ACTION_ALIASES = {
    "outline": "outline",
    "get_outline": "outline",
    "deck_outline": "outline",
    "read_slides": "read_slides",
    "read_slide": "read_slides",
    "get_slides": "read_slides",
    "get_slide": "read_slides",
    "read_pages": "read_slides",
    "search_text": "search_text",
    "search": "search_text",
    "find_text": "search_text",
    "read_table": "read_table",
    "get_table": "read_table",
    "inspect_shapes": "inspect_shapes",
    "inspect_shape": "inspect_shapes",
    "inspect": "inspect_shapes",
    "get_shape": "inspect_shapes",
    "read_shape": "inspect_shapes",
    "propose_operations": "propose_operations",
    "propose_operation": "propose_operations",
    "propose_ops": "propose_operations",
    "propose": "propose_operations",
    "submit_operations": "propose_operations",
    "submit_ops": "propose_operations",
    "propose_patches": "propose_operations",
    "submit_patches": "propose_operations",
    "finish": "finish",
    "final": "finish",
    "done": "finish",
    "stop": "finish",
    "answer": "finish",
    "ask": "finish",
    "duplicate_slide": "duplicate_slide",
    "copy_slide": "duplicate_slide",
    "clone_slide": "duplicate_slide",
    "duplicate": "duplicate_slide",
    "add_slide": "add_slide",
    "new_slide": "add_slide",
    "insert_slide": "add_slide",
    "delete_slide": "delete_slide",
    "remove_slide": "delete_slide",
    "move_slide": "move_slide",
    "reorder_slide": "move_slide",
    # v0.3 批量/增强编辑操作（与 edits 层别名保持一致）
    "unify_font": "font_unify",
    "unify_fonts": "font_unify",
    "global_replace": "replace_text",
    "replace_all": "replace_text",
    "scale_font": "font_scale",
    "page_numbers": "add_page_numbers",
    "add_page_number": "add_page_numbers",
    "align_shapes": "align_objects",
    "distribute": "align_objects",
    "distribute_objects": "align_objects",
    "copy_object": "duplicate_object",
    "clone_object": "duplicate_object",
    "zorder": "z_order",
    "reorder_z": "z_order",
    "layer": "z_order",
    "modify_table": "table_mod",
    "table_row": "table_mod",
    "table_col": "table_mod",
    "swap_image": "replace_image",
    "change_image": "replace_image",
    "background": "slide_background",
    "set_background": "slide_background",
    # v0.3 C 组：多步规划 / 撤销（工具，不是编辑操作）
    "plan": "plan",
    "make_plan": "plan",
    "plan_steps": "plan",
    "undo_last": "undo_last",
    "undo": "undo_last",
    "revert_last": "undo_last",
}

# slide-level operations submitted through propose_operations like any other op
SLIDE_ACTIONS = {"duplicate_slide", "add_slide", "delete_slide", "move_slide"}

# v0.3 batch / enhanced edit operations routed to edits.apply_patches_to_file
V03_ACTIONS = {
    "font_unify",
    "replace_text",
    "align_objects",
    "duplicate_object",
    "z_order",
    "table_mod",
    "replace_image",
    "slide_background",
    "add_page_numbers",
    "font_scale",
}

TOOL_NAMES = {
    "outline",
    "read_slides",
    "search_text",
    "read_table",
    "inspect_shapes",
    "propose_operations",
    "plan",
    "undo_last",
    "finish",
}

FOOTER_HINT_RE = re.compile(
    r"页脚|頁腳|页码|頁碼|页眉|頁眉|footer|page\s*number|header|版权|版權|copyright|"
    r"logo|标志|標誌|日期",
    re.IGNORECASE,
)
# 指令里明确要求整份生效（全局替换/统一字体等）时，自动把处理范围扩成整份，
# 否则 deck 级批量操作会被范围校验挡下
DECK_INTENT_RE = re.compile(
    r"全局|整份|全文|通篇|整套|所有页|所有幻灯片|所有页面|全部页|全部幻灯片|全部页面|"
    r"每一页|每一页|各页|跨页|全篇|所有文字|全部文字",
)
NUMBER_ONLY_RE = re.compile(r"^\d{1,4}$")
COLOR_RE = re.compile(r"^[0-9A-F]{6}$")
# 只给走得通的路：指令一旦点名了页面，自动扩范围不会生效，
# 所以不再建议「在指令里写全局/整份后重发」
_DECK_SCOPE_HINT = "该操作作用于整份文稿，需要整份范围；请让用户把范围切到「整份文稿」后重发"


def _as_int(value: Any) -> Optional[int]:
    if isinstance(value, bool) or value is None:
        return None
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return None


def _slide_index_of(entry: Dict[str, Any]) -> int:
    """0-based slide_index, -1 when missing. ``or -1`` would also map slide 0
    (the first page) to -1, breaking every operation that targets it."""
    try:
        return int(entry.get("slide_index"))
    except (TypeError, ValueError):
        return -1


def _as_float(value: Any) -> Optional[float]:
    if isinstance(value, bool) or value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_bool(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in ("1", "true", "yes", "y", "是", "加粗"):
        return True
    if text in ("0", "false", "no", "n", "否"):
        return False
    return None


def _clip(value: Any, limit: int) -> str:
    text = str(value if value is not None else "")
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)] + "…"


def _normalize_text(value: Any) -> str:
    return edits.normalize_text(value)


def _alignment_name(value: Any) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip().upper().replace("PP_ALIGN.", "")
    token = re.split(r"[\s(（]", text, maxsplit=1)[0]
    aliases = {
        "左对齐": "LEFT",
        "居中": "CENTER",
        "右对齐": "RIGHT",
        "两端对齐": "JUSTIFY",
        "分散对齐": "DISTRIBUTE",
    }
    if token in ALIGNMENTS:
        return token
    return aliases.get(str(value).strip())


def _normalize_style(raw: Any) -> Dict[str, Any]:
    if not isinstance(raw, dict):
        return {}
    style: Dict[str, Any] = {}
    size = _as_float(
        raw.get("font_size_pt") if "font_size_pt" in raw else raw.get("font_size", raw.get("size"))
    )
    if size is not None and 1 <= size <= 400:
        style["font_size_pt"] = int(size) if float(size).is_integer() else round(size, 2)
    name = raw.get("font_name", raw.get("font"))
    if isinstance(name, str) and name.strip():
        style["font_name"] = name.strip()[:100]
    for key in ("bold", "italic"):
        if key in raw:
            flag = _as_bool(raw.get(key))
            if flag is not None:
                style[key] = flag
    color = raw.get("color", raw.get("font_color"))
    if isinstance(color, str) and color.strip():
        value = color.strip().lstrip("#").upper()
        if len(value) == 3 and all(ch in "0123456789ABCDEF" for ch in value):
            value = "".join(ch * 2 for ch in value)
        if COLOR_RE.fullmatch(value):
            style["color"] = value
    alignment = _alignment_name(raw.get("alignment", raw.get("align")))
    if alignment:
        style["alignment"] = alignment
    return style


def _has_style_fields(raw: Dict[str, Any]) -> bool:
    return any(key in raw for key in STYLE_KEYS)


def _safe_emit(emit: EventEmitter, event: str, data: Dict[str, Any]) -> None:
    try:
        emit(event, data)
    except Exception:
        pass


def _cm_bbox(value: Any) -> Optional[Dict[str, float]]:
    if not isinstance(value, dict):
        return None
    try:
        return {
            "x": round(float(value.get("x") or 0) / EMU_PER_CM, 2),
            "y": round(float(value.get("y") or 0) / EMU_PER_CM, 2),
            "w": round(float(value.get("w") or 0) / EMU_PER_CM, 2),
            "h": round(float(value.get("h") or 0) / EMU_PER_CM, 2),
        }
    except (TypeError, ValueError):
        return None


def _slide_title(objects: List[Dict[str, Any]]) -> str:
    candidates: List[Tuple[Tuple[Any, ...], str]] = []
    for obj in objects:
        if obj.get("kind") != "text":
            continue
        raw = str(obj.get("text") or "").strip()
        if not raw:
            continue
        first = raw.split("\n", 1)[0].strip()
        if not first or NUMBER_ONLY_RE.fullmatch(first):
            continue
        bbox = obj.get("bbox_emu") or {}
        size = obj.get("slide_size_emu") or {}
        try:
            top_ratio = float(bbox.get("y") or 0) / max(1, int(size.get("height") or 1))
        except (TypeError, ValueError):
            top_ratio = 1.0
        font = _as_float(obj.get("font_size_pt")) or 0.0
        score = (0 if top_ratio < 0.3 else 1, 1 if obj.get("grouped") else 0, -font)
        candidates.append((score, first))
    if not candidates:
        return "（无文本）"
    candidates.sort(key=lambda item: item[0])
    return ai.clean_text(candidates[0][1], 60) or "（无文本）"


def _style_fields(obj: Dict[str, Any]) -> Dict[str, Any]:
    values: Dict[str, Any] = {}
    for key in STYLE_KEYS:
        if obj.get(key) is not None:
            values[key] = obj[key]
    if values.get("alignment") is not None:
        values["alignment"] = _alignment_name(values["alignment"]) or str(values["alignment"])
    return values


def _object_brief(obj: Dict[str, Any]) -> Dict[str, Any]:
    kind = "table_cell" if obj.get("kind") == "table_cell" else str(obj.get("kind") or "shape")
    brief: Dict[str, Any] = {
        "shape_id": obj.get("shape_id"),
        "kind": kind,
        "text": _clip(obj.get("text"), 420),
        "bbox_cm": _cm_bbox(obj.get("bbox_emu")),
        "grouped": bool(obj.get("grouped")),
    }
    if obj.get("shape_name") and kind == "text":
        brief["shape_name"] = _clip(obj.get("shape_name"), 50)
    if obj.get("parent_shape_id"):
        brief["parent_shape_id"] = obj.get("parent_shape_id")
    style = _style_fields(obj)
    if style:
        brief.update(style)
    if kind == "table_cell":
        for key in ("table_shape_id", "row", "col", "row_span", "col_span"):
            if obj.get(key) is not None:
                brief[key] = obj.get(key)
    return brief


def _object_detail(obj: Dict[str, Any]) -> Dict[str, Any]:
    detail = _object_brief(obj)
    detail["text"] = _clip(obj.get("text"), 2000)
    detail["bbox_emu"] = obj.get("bbox_emu")
    detail["deletable"] = bool(obj.get("deletable"))
    detail["movable"] = bool(obj.get("movable"))
    detail["resizable"] = bool(obj.get("resizable"))
    if obj.get("group_path"):
        detail["group_path"] = obj.get("group_path")
    if obj.get("empty_cell"):
        detail["empty_cell"] = True
    return detail


def _table_summary(table: Dict[str, Any]) -> Dict[str, Any]:
    slide_index = (_slide_index_of(table) if table.get("slide_index") is not None else 0)
    return {
        "shape_id": table.get("shape_id"),
        "slide_index": slide_index,
        "page": slide_index + 1,
        "name": _clip(table.get("name"), 50),
        "rows": table.get("rows"),
        "cols": table.get("cols"),
        "bbox_cm": _cm_bbox(table.get("bbox_emu")),
        "cell_count": table.get("cell_count"),
        "editable_cells": table.get("editable_cells"),
        "grouped": bool(table.get("grouped")),
    }


class _DeckView:
    def __init__(
        self,
        objects: List[Dict[str, Any]],
        elements: List[Dict[str, Any]],
        tables: List[Dict[str, Any]],
        total_slides: int,
        scope_slides: List[int],
        slide_size: Optional[Dict[str, Any]],
    ) -> None:
        self.objects = [obj for obj in objects if obj.get("shape_id") is not None]
        self.elements = [elem for elem in elements if elem.get("shape_id") is not None]
        self.tables = [table for table in tables if table.get("shape_id") is not None]
        self.total_slides = max(0, int(total_slides))
        self.scope_slides = sorted({int(index) for index in scope_slides})
        self.scope_set = set(self.scope_slides)
        self.default_slide = self.scope_slides[0] if len(self.scope_slides) == 1 else None
        size = slide_size if isinstance(slide_size, dict) else {}
        self.slide_w = int(size.get("width") or DEFAULT_SLIDE_WIDTH)
        self.slide_h = int(size.get("height") or DEFAULT_SLIDE_HEIGHT)

        self.object_by_id = {str(obj["shape_id"]): obj for obj in self.objects}
        self.element_by_id = {str(elem["shape_id"]): elem for elem in self.elements}
        self.table_by_id = {str(table["shape_id"]): table for table in self.tables}
        self.objects_by_slide: Dict[int, List[Dict[str, Any]]] = {}
        self.elements_by_slide: Dict[int, List[Dict[str, Any]]] = {}
        self.tables_by_slide: Dict[int, List[Dict[str, Any]]] = {}
        for obj in self.objects:
            try:
                self.objects_by_slide.setdefault(int(obj["slide_index"]), []).append(obj)
            except (KeyError, TypeError, ValueError):
                continue
        for elem in self.elements:
            try:
                self.elements_by_slide.setdefault(int(elem["slide_index"]), []).append(elem)
            except (KeyError, TypeError, ValueError):
                continue
        for table in self.tables:
            try:
                self.tables_by_slide.setdefault(int(table["slide_index"]), []).append(table)
            except (KeyError, TypeError, ValueError):
                continue

        self.child_counts: Dict[str, int] = {}
        for elem in self.elements:
            parent = elem.get("parent_shape_id")
            if parent:
                self.child_counts[str(parent)] = self.child_counts.get(str(parent), 0) + 1

        self.covered: Dict[str, set] = {}
        for obj in self.objects:
            if obj.get("kind") != "table_cell":
                continue
            table_id = str(obj.get("table_shape_id") or "")
            if not table_id:
                continue
            try:
                row = int(obj.get("row") or 0)
                col = int(obj.get("col") or 0)
                row_span = max(1, int(obj.get("row_span") or 1))
                col_span = max(1, int(obj.get("col_span") or 1))
            except (TypeError, ValueError):
                continue
            covered = self.covered.setdefault(table_id, set())
            for rr in range(row, row + row_span):
                for cc in range(col, col + col_span):
                    covered.add((rr, cc))

    def is_scoped(self, slide_index: int) -> bool:
        return int(slide_index) in self.scope_set

    def resolve_text_target(self, shape_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        cell_ref = edits.parse_cell_id(shape_id)
        if cell_ref:
            table_id, row, col = cell_ref
            table = self.table_by_id.get(table_id)
            if table is None:
                return None, "表格不存在，请先 read_table 确认"
            try:
                slide_index = int(table.get("slide_index"))
                rows = int(table.get("rows") or 0)
                cols = int(table.get("cols") or 0)
            except (TypeError, ValueError):
                return None, "表格元数据异常"
            if not self.is_scoped(slide_index):
                return None, "表格不在当前处理范围"
            if row < 0 or col < 0 or row >= rows or col >= cols:
                return None, "单元格行列超出表格范围"
            obj = self.object_by_id.get(shape_id)
            if obj is not None:
                if not self.is_scoped(_slide_index_of(obj)):
                    return None, "对象不在当前处理范围"
                return obj, None
            if (row, col) in self.covered.get(table_id, set()):
                return None, "该坐标属于合并单元格，请使用合并起始单元格"
            pseudo = {
                "shape_id": shape_id,
                "slide_index": slide_index,
                "kind": "table_cell",
                "text": "",
                "bbox_emu": None,
                "table_shape_id": table_id,
                "row": row,
                "col": col,
                "row_span": 1,
                "col_span": 1,
                "table_rows": rows,
                "table_cols": cols,
                "shape_name": f"{table.get('name') or '表格'} / R{row + 1}C{col + 1}",
                "grouped": bool(table.get("grouped")),
                "deletable": False,
                "movable": False,
                "resizable": False,
                "empty_cell": True,
            }
            return pseudo, None
        obj = self.object_by_id.get(shape_id)
        if obj is None:
            return None, "对象不存在，请先读取页面或使用 search_text 查找"
        if not self.is_scoped(_slide_index_of(obj)):
            return None, "对象不在当前处理范围"
        if obj.get("kind") not in ("text", "table_cell"):
            return None, "该对象没有可编辑文本"
        return obj, None

    def resolve_geometry_target(self, shape_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        if edits.parse_cell_id(shape_id):
            return None, "表格单元格不支持几何调整"
        obj = self.object_by_id.get(shape_id)
        if obj is not None:
            if not self.is_scoped(_slide_index_of(obj)):
                return None, "对象不在当前处理范围"
            if obj.get("kind") == "table_cell":
                return None, "表格单元格不支持几何调整"
            if obj.get("grouped"):
                return None, "分组内的对象不能移动或缩放"
            return obj, None
        elem = self.element_by_id.get(shape_id)
        if elem is not None:
            if not self.is_scoped(_slide_index_of(elem)):
                return None, "对象不在当前处理范围"
            if elem.get("grouped"):
                return None, "分组内的对象不能移动或缩放"
            return elem, None
        return None, "对象不存在"

    def resolve_delete_target(self, shape_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        if edits.parse_cell_id(shape_id):
            return None, "表格单元格不能删除"
        obj = self.object_by_id.get(shape_id)
        if obj is None:
            return None, "对象不存在"
        if not self.is_scoped(_slide_index_of(obj)):
            return None, "对象不在当前处理范围"
        if obj.get("kind") != "text" or obj.get("grouped") or not obj.get("deletable"):
            return None, "只能删除非分组的普通文本对象"
        return obj, None

    def inspect(self, shape_id: str) -> Dict[str, Any]:
        obj = self.object_by_id.get(shape_id)
        if obj is not None:
            if not self.is_scoped(_slide_index_of(obj)):
                return {"shape_id": shape_id, "error": "对象不在当前处理范围"}
            detail = _object_detail(obj)
            if obj.get("kind") == "table_cell":
                table = self.table_by_id.get(str(obj.get("table_shape_id") or ""))
                if table is not None:
                    detail["table"] = _table_summary(table)
            return detail
        elem = self.element_by_id.get(shape_id)
        if elem is not None:
            if not self.is_scoped(_slide_index_of(elem)):
                return {"shape_id": shape_id, "error": "对象不在当前处理范围"}
            slide_index = (_slide_index_of(elem) if elem.get("slide_index") is not None else 0)
            detail = {
                "shape_id": elem.get("shape_id"),
                "kind": elem.get("kind"),
                "slide_index": slide_index,
                "page": slide_index + 1,
                "name": _clip(elem.get("name"), 60),
                "bbox_emu": elem.get("bbox_emu"),
                "bbox_cm": _cm_bbox(elem.get("bbox_emu")),
                "grouped": bool(elem.get("grouped")),
                "parent_shape_id": elem.get("parent_shape_id"),
                "group_path": elem.get("group_path") or [],
                "child_count": self.child_counts.get(shape_id, 0),
            }
            table = self.table_by_id.get(shape_id)
            if table is not None:
                detail["table"] = _table_summary(table)
                detail["hint"] = "表格内容请用 read_table 读取"
            return detail
        table = self.table_by_id.get(shape_id)
        if table is not None:
            if not self.is_scoped(_slide_index_of(table)):
                return {"shape_id": shape_id, "error": "表格不在当前处理范围"}
            return {"shape_id": shape_id, **_table_summary(table)}
        target, error = self.resolve_text_target(shape_id)
        if target is not None and target.get("empty_cell"):
            return _object_detail(target)
        return {"shape_id": shape_id, "error": error or "对象不存在"}


def _tool_outline(view: _DeckView) -> Dict[str, Any]:
    slides: List[Dict[str, Any]] = []
    for index in range(view.total_slides):
        objs = view.objects_by_slide.get(index, [])
        texts = [obj for obj in objs if obj.get("kind") == "text"]
        table_cells = [obj for obj in objs if obj.get("kind") == "table_cell"]
        table_metas = view.tables_by_slide.get(index, [])
        visuals = [
            elem
            for elem in view.elements_by_slide.get(index, [])
            if elem.get("kind") in ("picture", "chart", "group", "shape") and not elem.get("grouped")
        ]
        slides.append(
            {
                "slide_index": index,
                "page": index + 1,
                "title": _slide_title(texts),
                "text_count": len(texts),
                "table_cell_count": len(table_cells),
                "table_count": len(table_metas),
                "visual_count": len(visuals),
                "in_scope": index in view.scope_set,
            }
        )
    return {"total_slides": view.total_slides, "scope_slides": view.scope_slides, "slides": slides}


def _tool_read_slides(view: _DeckView, args: Dict[str, Any]) -> Dict[str, Any]:
    raw = args.get("slide_indices")
    if raw is None:
        raw = args.get("slide_index")
    if raw is None:
        raw = args.get("pages")
    if raw is None:
        raw = view.scope_slides[:4]
    if not isinstance(raw, list):
        raw = [raw]
    requested: List[int] = []
    ignored: List[Any] = []
    for value in raw:
        index = _as_int(value)
        if index is None:
            ignored.append(value)
            continue
        if index in view.scope_set:
            if index not in requested:
                requested.append(index)
        else:
            ignored.append(index)
    if not requested:
        return {"error": "slide_indices 不在当前处理范围", "available": view.scope_slides}
    requested = requested[:4]
    budget = MAX_TOOL_RESULT_CHARS
    slides: List[Dict[str, Any]] = []
    omitted = 0
    for index in requested:
        objs = view.objects_by_slide.get(index, [])
        texts = [obj for obj in objs if obj.get("kind") == "text"]
        section: Dict[str, Any] = {
            "slide_index": index,
            "page": index + 1,
            "title": _slide_title(texts),
            "objects": [],
            "tables": [],
            "visuals": [],
        }
        for obj in objs:
            brief = _object_brief(obj)
            cost = len(json.dumps(brief, ensure_ascii=False))
            if budget - cost < 700:
                omitted += 1
                continue
            section["objects"].append(brief)
            budget -= cost
        for table in view.tables_by_slide.get(index, [])[:8]:
            summary = _table_summary(table)
            summary.pop("page", None)
            summary.pop("grouped", None)
            section["tables"].append(summary)
        visuals = [
            elem
            for elem in view.elements_by_slide.get(index, [])
            if elem.get("kind") in ("picture", "chart", "group", "shape") and not elem.get("grouped")
        ]
        for elem in visuals[:12]:
            item = {
                "shape_id": elem.get("shape_id"),
                "kind": elem.get("kind"),
                "name": _clip(elem.get("name"), 40),
                "bbox_cm": _cm_bbox(elem.get("bbox_emu")),
            }
            if elem.get("kind") == "group":
                item["children"] = view.child_counts.get(str(elem.get("shape_id")), 0)
            section["visuals"].append(item)
        slides.append(section)
    result: Dict[str, Any] = {
        "slides": slides,
        "scope_slides": view.scope_slides,
        "ignored_slides": ignored,
    }
    if omitted:
        result["omitted_objects"] = omitted
        result["note"] = "对象过多已截断，可缩小页码范围或用 search_text/inspect_shapes 精查"
    return result


def _tool_search_text(view: _DeckView, args: Dict[str, Any]) -> Dict[str, Any]:
    query = str(args.get("query") or args.get("text") or args.get("keyword") or "").strip()
    if not query:
        return {"error": "query 不能为空"}
    needle = query.casefold()
    compact_needle = _normalize_text(query)
    matches: List[Dict[str, Any]] = []
    total = 0
    for obj in view.objects:
        if not view.is_scoped(_slide_index_of(obj)):
            continue
        text = str(obj.get("text") or "")
        lower = text.casefold()
        compact = _normalize_text(text)
        if needle not in lower and (not compact_needle or compact_needle not in compact):
            continue
        total += 1
        if len(matches) >= 30:
            continue
        item: Dict[str, Any] = {
            "shape_id": obj.get("shape_id"),
            "slide_index": obj.get("slide_index"),
            "page": (_slide_index_of(obj) if obj.get("slide_index") is not None else 0) + 1,
            "kind": "table_cell" if obj.get("kind") == "table_cell" else obj.get("kind"),
            "text": _clip(text, 180),
        }
        if obj.get("table_shape_id"):
            item["table_shape_id"] = obj.get("table_shape_id")
            item["row"] = obj.get("row")
            item["col"] = obj.get("col")
        matches.append(item)
    return {
        "query": query,
        "count": total,
        "matches": matches,
        "truncated": total > len(matches),
    }


def _tool_read_table(view: _DeckView, args: Dict[str, Any]) -> Dict[str, Any]:
    table_id = str(
        args.get("table_shape_id") or args.get("shape_id") or args.get("table_id") or ""
    ).strip()
    if not table_id:
        return {"error": "缺少 table_shape_id"}
    table = view.table_by_id.get(table_id)
    if table is None:
        return {"error": "表格不存在，请先 read_slides 或 outline 确认"}
    if not view.is_scoped(_slide_index_of(table)):
        return {"error": "表格不在当前处理范围"}
    cells = [
        obj
        for obj in view.objects
        if obj.get("kind") == "table_cell" and str(obj.get("table_shape_id") or "") == table_id
    ]
    cells.sort(key=lambda obj: (int(obj.get("row") or 0), int(obj.get("col") or 0)))
    out_cells: List[Dict[str, Any]] = []
    budget = MAX_TOOL_RESULT_CHARS
    for obj in cells:
        item: Dict[str, Any] = {
            "shape_id": obj.get("shape_id"),
            "row": obj.get("row"),
            "col": obj.get("col"),
            "row_span": obj.get("row_span"),
            "col_span": obj.get("col_span"),
            "text": _clip(obj.get("text"), 600),
        }
        for key in ("font_size_pt", "bold", "italic", "alignment", "color"):
            if obj.get(key) is not None:
                item[key] = _alignment_name(obj[key]) if key == "alignment" else obj[key]
        cost = len(json.dumps(item, ensure_ascii=False))
        if budget - cost < 500:
            break
        out_cells.append(item)
        budget -= cost
    summary = _table_summary(table)
    return {
        "table_shape_id": table_id,
        "slide_index": summary["slide_index"],
        "page": summary["page"],
        "name": summary["name"],
        "rows": summary["rows"],
        "cols": summary["cols"],
        "bbox_cm": summary["bbox_cm"],
        "editable_cells": summary["editable_cells"],
        "cells": out_cells,
        "omitted_cells": max(0, len(cells) - len(out_cells)),
        "note": "只列出非空单元格；合并区域用起始单元格 shape_id",
    }


def _tool_inspect_shapes(view: _DeckView, args: Dict[str, Any]) -> Dict[str, Any]:
    raw = args.get("shape_ids")
    if raw is None:
        raw = args.get("shape_id")
    if raw is None:
        raw = args.get("ids")
    if raw is None:
        return {"error": "缺少 shape_ids"}
    if not isinstance(raw, list):
        raw = [raw]
    shape_ids: List[str] = []
    for value in raw:
        shape_id = str(value or "").strip()
        if shape_id and shape_id not in shape_ids:
            shape_ids.append(shape_id)
    if not shape_ids:
        return {"error": "shape_ids 不能为空"}
    truncated = len(shape_ids) > 8
    shapes = [view.inspect(shape_id) for shape_id in shape_ids[:8]]
    result: Dict[str, Any] = {"shapes": shapes}
    if truncated:
        result["truncated"] = True
    return result


def _dispatch_tool(action: str, args: Dict[str, Any], view: _DeckView) -> Dict[str, Any]:
    if action == "outline":
        return _tool_outline(view)
    if action == "read_slides":
        return _tool_read_slides(view, args)
    if action == "search_text":
        return _tool_search_text(view, args)
    if action == "read_table":
        return _tool_read_table(view, args)
    if action == "inspect_shapes":
        return _tool_inspect_shapes(view, args)
    return {"error": f"未知工具：{action}"}


def _normalize_plan(raw: Any, min_steps: int = 2) -> Tuple[Optional[List[str]], Optional[str]]:
    """Validate a plan step list; returns (steps, error) — exactly one set."""
    if not isinstance(raw, list):
        return None, 'plan 需要 steps 数组，例如 {"action":"plan","args":{"steps":["第一步","第二步"]}}'
    steps = [_clip(str(item or "").strip(), 120) for item in raw]
    steps = [s for s in steps if s]
    if len(steps) < min_steps:
        return None, f"plan 至少需要 {min_steps} 个非空步骤描述（收到 {len(steps)} 个）"
    return steps[:12], None


def _observation_content(detail: Dict[str, Any]) -> str:
    """Observation message for the model; tool errors carry a retry hint so the
    loop self-heals (换参数重试) instead of surfacing the failure to the user."""
    text = "Observation:\n" + json.dumps(detail, ensure_ascii=False)
    if isinstance(detail, dict) and detail.get("error"):
        text += (
            "\n该调用失败：请修正参数后重试（换参数或换一个等价工具），"
            "不要因此结束任务；确实无法完成时再用 finish 说明原因。"
        )
    return text


class _OperationStore:
    def __init__(self, view: _DeckView, instruction: str, limit: int) -> None:
        self.view = view
        self.instruction = instruction
        self.limit = max(1, int(limit))
        self.ops: List[Dict[str, Any]] = []
        self.dropped = 0
        self.text_targets: set = set()
        self.style_targets: set = set()
        self.geometry_targets: set = set()
        self.add_keys: set = set()
        self.slide_keys: set = set()
        self.deleted_slides: set = set()
        self.v03_keys: set = set()

    def add_batch(self, raw: Any) -> Dict[str, Any]:
        if not isinstance(raw, list):
            return {"error": "operations 必须是数组"}
        if not raw:
            return {"error": "operations 为空"}
        accepted: List[Dict[str, Any]] = []
        reasons: List[Dict[str, Any]] = []
        batch_dropped = 0
        for index, item in enumerate(raw):
            if index >= MAX_OPS_PER_CALL:
                extra = len(raw) - MAX_OPS_PER_CALL
                batch_dropped += extra
                reasons.append(
                    {"index": index, "reason": f"单次最多 {MAX_OPS_PER_CALL} 条，其余已忽略"}
                )
                break
            if len(self.ops) >= self.limit:
                batch_dropped += 1
                reasons.append({"index": index, "reason": "已达到本次操作总数上限"})
                continue
            try:
                op, reason = self._normalize(item)
            except Exception as exc:
                op, reason = None, f"操作校验失败：{exc}"
            if op is None:
                batch_dropped += 1
                reasons.append({"index": index, "reason": reason or "无效操作"})
                continue
            self.ops.append(op)
            accepted.append(op)
            self._register(op)
        self.dropped += batch_dropped
        return {
            "accepted": len(accepted),
            "dropped": batch_dropped,
            "dropped_total": self.dropped,
            "accepted_operations": [self._op_summary(op) for op in accepted],
            "reasons": reasons[:10],
            "total": len(self.ops),
            "remaining": max(0, self.limit - len(self.ops)),
        }

    def _normalize(self, item: Any) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        if not isinstance(item, dict):
            return None, "操作必须是对象"
        payload = item.get("payload") if isinstance(item.get("payload"), dict) else None
        raw = {**payload, **item} if payload else item
        shape_id = str(
            raw.get("shape_id") or raw.get("target_id") or raw.get("id") or ""
        ).strip()
        action = _canonical_action(raw.get("action") or raw.get("type") or raw.get("op"))
        after = raw.get("after")
        if after is None:
            after = raw.get("text")
        if after is None:
            after = raw.get("new_text")
        if after is None:
            after = raw.get("content")
        if not action:
            if shape_id and after is not None:
                action = "edit"
            elif shape_id and _has_style_fields(raw):
                action = "style"
            elif shape_id and isinstance(raw.get("bbox_emu"), dict):
                action = "geometry"
            elif raw.get("slide_index") is not None and after is not None and isinstance(
                raw.get("bbox_emu"), dict
            ):
                action = "add_textbox"
            else:
                return None, "缺少 action"
        if action in SLIDE_ACTIONS:
            return self._slide_op(raw, action)
        if action in V03_ACTIONS:
            return self._v03_op(raw, action)
        if action in TEXT_ACTIONS or action in ("style", "geometry"):
            if not shape_id:
                return None, "缺少 shape_id"
            if action == "edit":
                return self._edit(raw, shape_id, after)
            if action == "style":
                return self._style(raw, shape_id)
            if action == "clear":
                return self._clear(raw, shape_id)
            if action == "geometry":
                return self._geometry(raw, shape_id)
            if action == "delete_textbox":
                return self._delete(raw, shape_id)
            return None, f"不支持的操作：{action}"
        if action == "add_textbox":
            return self._add(raw, after)
        return None, f"不支持的操作：{action}"

    def _before_ok(self, raw: Dict[str, Any], current: str) -> bool:
        before = raw.get("before")
        if before is None or not str(before).strip():
            return True
        if _normalize_text(before) == _normalize_text(current):
            return True
        ratio = difflib.SequenceMatcher(
            None, _normalize_text(before), _normalize_text(current)
        ).ratio()
        return ratio >= 0.9

    def _footer_requested(self) -> bool:
        return bool(FOOTER_HINT_RE.search(self.instruction))

    def _protected_text(self, obj: Dict[str, Any]) -> bool:
        if self._footer_requested() or obj.get("kind") != "text":
            return False
        text = str(obj.get("text") or "").strip()
        if NUMBER_ONLY_RE.fullmatch(text):
            return True
        if ai.region_of(obj) == "页脚":
            return True
        name = str(obj.get("shape_name") or "").casefold()
        return any(
            token in name
            for token in ("logo", "标志", "標誌", "页脚", "頁腳", "footer", "页码", "頁碼")
        )

    def _protected_visual(self, target: Dict[str, Any]) -> bool:
        if self._footer_requested():
            return False
        name = str(target.get("shape_name") or target.get("name") or "").casefold()
        return any(
            token in name
            for token in ("logo", "标志", "標誌", "页脚", "頁腳", "footer", "页码", "頁碼")
        )

    def _description(self, raw: Dict[str, Any], default: str) -> str:
        value = raw.get("description") or raw.get("reason") or raw.get("note") or default
        return ai.clean_text(value, 200) or default

    def _extract_style(self, raw: Dict[str, Any]) -> Dict[str, Any]:
        style: Dict[str, Any] = {}
        if isinstance(raw.get("style"), dict):
            style.update(_normalize_style(raw.get("style")))
        if _has_style_fields(raw):
            direct = {key: raw.get(key) for key in STYLE_KEYS if key in raw}
            style.update(_normalize_style(direct))
        return style

    def _target_meta(self, obj: Dict[str, Any]) -> Dict[str, Any]:
        meta: Dict[str, Any] = {
            "slide_index": (_slide_index_of(obj) if obj.get("slide_index") is not None else 0),
            "kind": obj.get("kind"),
            "grouped": bool(obj.get("grouped")),
        }
        for key in (
            "bbox_emu",
            "font_name",
            "font_size_pt",
            "bold",
            "italic",
            "color",
            "alignment",
            "table_shape_id",
            "row",
            "col",
            "row_span",
            "col_span",
        ):
            if obj.get(key) is not None:
                meta[key] = obj.get(key)
        if meta.get("alignment") is not None:
            meta["alignment"] = _alignment_name(meta["alignment"]) or str(meta["alignment"])
        if obj.get("empty_cell"):
            meta["empty_cell"] = True
        return meta

    def _build_text_op(
        self,
        action: str,
        obj: Dict[str, Any],
        shape_id: str,
        before: str,
        after: Optional[str],
        style: Dict[str, Any],
        raw: Dict[str, Any],
        default_description: str,
    ) -> Dict[str, Any]:
        op: Dict[str, Any] = {
            "id": f"op{len(self.ops) + 1}",
            "source": "agent",
            "action": action,
            "slide_index": (_slide_index_of(obj) if obj.get("slide_index") is not None else 0),
            "shape_id": shape_id,
            "shape_name": obj.get("shape_name"),
            "kind": obj.get("kind"),
            "before": before,
            "description": self._description(raw, default_description),
            "meta": self._target_meta(obj),
        }
        if after is not None:
            op["after"] = after
        if style:
            op["style"] = style
        if obj.get("kind") == "table_cell":
            op["table_shape_id"] = obj.get("table_shape_id")
            op["row"] = obj.get("row")
            op["col"] = obj.get("col")
        return op

    def _edit(
        self, raw: Dict[str, Any], shape_id: str, after: Any
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        obj, error = self.view.resolve_text_target(shape_id)
        if obj is None:
            return None, error
        if shape_id in self.text_targets:
            return None, "同一对象已有文本修改操作"
        current = str(obj.get("text") or "")
        if not self._before_ok(raw, current):
            return None, "before 与当前原文不一致，请重新读取后再改"
        if self._protected_text(obj):
            return None, "该对象疑似页码/页脚/logo，未获明确授权"
        style = self._extract_style(raw)
        if after is None:
            if style:
                return self._style(raw, shape_id)
            return None, "edit 缺少 after"
        text = str(after)
        if not text.strip():
            if not current.strip():
                return None, "文本已经为空"
            if obj.get("grouped") and obj.get("kind") == "text":
                return None, "分组内文字不能清空；请改写为具体文本"
            return self._build_text_op(
                "clear", obj, shape_id, current, "", {}, raw, "清空文本"
            ), None
        if _normalize_text(text) == _normalize_text(current) and not style:
            return None, "新文本与原文相同"
        if len(text) > max(1200, len(current) * 4):
            return None, "新文本过长，请精简后再提交"
        return self._build_text_op(
            "edit", obj, shape_id, current, text, style, raw, "修改文案"
        ), None

    def _style(
        self, raw: Dict[str, Any], shape_id: str
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        obj, error = self.view.resolve_text_target(shape_id)
        if obj is None:
            return None, error
        if shape_id in self.style_targets:
            return None, "同一对象已有样式修改操作"
        if self._protected_text(obj):
            return None, "该对象疑似页码/页脚/logo，未获明确授权"
        style = self._extract_style(raw)
        if not style:
            return None, "style 缺少有效字段"
        current = str(obj.get("text") or "")
        if not current.strip() and obj.get("kind") == "table_cell":
            return None, "空单元格没有文本，无法调整样式"
        return self._build_text_op(
            "style", obj, shape_id, current, None, style, raw, "调整样式"
        ), None

    def _clear(
        self, raw: Dict[str, Any], shape_id: str
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        obj, error = self.view.resolve_text_target(shape_id)
        if obj is None:
            return None, error
        if shape_id in self.text_targets:
            return None, "同一对象已有文本修改操作"
        current = str(obj.get("text") or "")
        if not self._before_ok(raw, current):
            return None, "before 与当前原文不一致，请重新读取后再改"
        if self._protected_text(obj):
            return None, "该对象疑似页码/页脚/logo，未获明确授权"
        if obj.get("grouped") and obj.get("kind") == "text":
            return None, "分组内文字不能清空；请改写为具体文本"
        if not current.strip():
            return None, "文本已经为空"
        return self._build_text_op(
            "clear", obj, shape_id, current, "", {}, raw, "清空文本"
        ), None

    def _delete(
        self, raw: Dict[str, Any], shape_id: str
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        obj, error = self.view.resolve_delete_target(shape_id)
        if obj is None:
            return None, error
        if shape_id in self.text_targets:
            return None, "同一对象已有文本修改操作"
        if not self._before_ok(raw, str(obj.get("text") or "")):
            return None, "before 与当前原文不一致，请重新读取后再改"
        if self._protected_text(obj):
            return None, "该对象疑似页码/页脚/logo，未获明确授权"
        return self._build_text_op(
            "delete_textbox",
            obj,
            shape_id,
            str(obj.get("text") or ""),
            "",
            {},
            raw,
            "删除文本框",
        ), None

    def _norm_bbox(
        self, raw: Any, current: Optional[Dict[str, Any]]
    ) -> Tuple[Optional[Dict[str, int]], Optional[str]]:
        if not isinstance(raw, dict) or not raw:
            return None, "缺少 bbox_emu"
        pairs = {
            "x": ("x", "left"),
            "y": ("y", "top"),
            "w": ("w", "width"),
            "h": ("h", "height"),
        }
        values: Dict[str, int] = {}
        for key, names in pairs.items():
            value = None
            for name in names:
                if name in raw:
                    value = _as_int(raw.get(name))
                    if value is not None:
                        break
            if value is None and isinstance(current, dict):
                for name in names:
                    if name in current:
                        value = _as_int(current.get(name))
                        if value is not None:
                            break
            if value is None:
                return None, f"bbox_emu 缺少 {key}"
            values[key] = value
        x, y, w, h = values["x"], values["y"], values["w"], values["h"]
        if w <= 0 or h <= 0:
            return None, "bbox_emu 的宽高必须大于 0"
        slide_w = max(1, self.view.slide_w)
        slide_h = max(1, self.view.slide_h)
        x = max(0, min(x, slide_w - 1))
        y = max(0, min(y, slide_h - 1))
        w = max(1, min(w, slide_w - x))
        h = max(1, min(h, slide_h - y))
        return {"x": x, "y": y, "w": w, "h": h}, None

    def _geometry(
        self, raw: Dict[str, Any], shape_id: str
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        target, error = self.view.resolve_geometry_target(shape_id)
        if target is None:
            return None, error
        if shape_id in self.geometry_targets:
            return None, "同一对象已有几何调整操作"
        if self._protected_visual(target):
            return None, "该对象疑似页码/页脚/logo，未获明确授权"
        current_bbox = (
            target.get("bbox_emu") if isinstance(target.get("bbox_emu"), dict) else None
        )
        bbox, bbox_error = self._norm_bbox(
            raw.get("bbox_emu") if raw.get("bbox_emu") is not None else raw.get("bbox"),
            current_bbox,
        )
        if bbox is None:
            return None, bbox_error
        op: Dict[str, Any] = {
            "id": f"op{len(self.ops) + 1}",
            "source": "agent",
            "action": "geometry",
            "slide_index": (_slide_index_of(target) if target.get("slide_index") is not None else 0),
            "shape_id": shape_id,
            "kind": target.get("kind"),
            "name": target.get("name") or target.get("shape_name"),
            "bbox_emu": bbox,
            "description": self._description(raw, "调整位置或大小"),
            "meta": self._target_meta(target),
        }
        return op, None

    def _add(
        self, raw: Dict[str, Any], after: Any
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        if after is None:
            return None, "add_textbox 缺少文本"
        text = str(after)
        if not text.strip():
            return None, "新增文本框的文本不能为空"
        if len(text) > 2000:
            return None, "新增文本过长"
        slide_raw = raw.get("slide_index")
        if slide_raw is None:
            if self.view.default_slide is None:
                return None, "add_textbox 缺少 slide_index"
            slide_index = self.view.default_slide
        else:
            slide_index = _as_int(slide_raw)
            if slide_index is None:
                return None, "slide_index 无效"
        if slide_index not in self.view.scope_set:
            return None, "slide_index 不在当前处理范围"
        bbox, bbox_error = self._norm_bbox(raw.get("bbox_emu"), None)
        if bbox is None:
            return None, bbox_error
        key = (
            slide_index,
            _normalize_text(text),
            bbox["x"],
            bbox["y"],
            bbox["w"],
            bbox["h"],
        )
        if key in self.add_keys:
            return None, "重复的新增文本框操作"
        style = self._extract_style(raw)
        alignment = style.get("alignment")
        if alignment and alignment not in ADD_ALIGNMENTS:
            style.pop("alignment", None)
        op: Dict[str, Any] = {
            "id": f"op{len(self.ops) + 1}",
            "source": "agent",
            "action": "add_textbox",
            "slide_index": slide_index,
            "after": text,
            "bbox_emu": bbox,
            "description": self._description(raw, "新增文本框"),
            "meta": {
                "slide_index": slide_index,
                "kind": "textbox",
                "bbox_emu": bbox,
                **style,
            },
        }
        for key_name, value in style.items():
            op[key_name] = value
        return op, None

    def _slide_op(
        self, raw: Dict[str, Any], action: str
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        total = self.view.total_slides

        def _slide_number(*names: str) -> Optional[int]:
            for name in names:
                if raw.get(name) is not None:
                    return _as_int(raw.get(name))
            return None

        def _deleted_reason(slide: Optional[int]) -> Optional[str]:
            if slide is not None and slide in self.deleted_slides:
                return f"第 {slide + 1} 页已在本批操作中删除（已删除）"
            return None

        if action == "duplicate_slide":
            source = _slide_number("source_slide", "source", "slide_index", "slide")
            if source is None:
                return None, "duplicate_slide 缺少 source_slide"
            if not self.view.is_scoped(source):
                return None, "source_slide 不在当前处理范围"
            deleted_reason = _deleted_reason(source)
            if deleted_reason:
                return None, deleted_reason
            position = _slide_number("position", "to", "insert_at")
            if position is None:
                position = source + 1
            if position < 0 or position > total:
                return None, f"position 必须在 0..{total}"
            defaults = ("复制页面",)
            op = {
                "action": action,
                "slide_index": source,
                "source_slide": source,
                "position": position,
            }
        elif action == "add_slide":
            after = _slide_number("after_slide")
            if after is not None:
                if after < 0 or after >= total:
                    return None, f"after_slide 必须在 0..{total - 1}"
                if not self.view.is_scoped(after):
                    return None, "after_slide 不在当前处理范围"
                deleted_reason = _deleted_reason(after)
                if deleted_reason:
                    return None, deleted_reason
            position = _slide_number("position", "insert_at")
            if position is not None and (position < 0 or position > total):
                return None, f"position 必须在 0..{total}"
            layout_index = _slide_number("layout_index", "layout")
            if layout_index is not None and layout_index < 0:
                return None, "layout_index 不能为负"
            anchor = after + 1 if after is not None else (position if position is not None else total)
            op = {
                "action": action,
                "slide_index": anchor,
                "after_slide": after,
                "position": position,
                "layout_index": layout_index,
            }
            defaults = ("新增空白页",)
        elif action == "delete_slide":
            slide = _slide_number("slide", "slide_index", "source_slide")
            if slide is None:
                return None, "delete_slide 缺少 slide"
            if not self.view.is_scoped(slide):
                return None, "该页不在当前处理范围"
            if total <= 1:
                return None, "演示文稿只剩唯一一页，不能删除"
            deleted_reason = _deleted_reason(slide)
            if deleted_reason:
                return None, deleted_reason
            op = {"action": action, "slide_index": slide, "slide": slide}
            defaults = ("删除整页",)
        else:  # move_slide
            slide = _slide_number("slide", "slide_index", "from", "source_slide")
            if slide is None:
                return None, "move_slide 缺少 slide"
            if not self.view.is_scoped(slide):
                return None, "该页不在当前处理范围"
            deleted_reason = _deleted_reason(slide)
            if deleted_reason:
                return None, deleted_reason
            to = _slide_number("to", "position", "to_index")
            if to is None:
                return None, "move_slide 缺少 to"
            if to < 0 or to >= total:
                return None, f"to 必须在 0..{total - 1}"
            op = {"action": action, "slide_index": slide, "slide": slide, "to": to}
            defaults = ("调整页序",)

        key = (action, op.get("source_slide"), op.get("position"), op.get("after_slide"),
               op.get("layout_index"), op.get("slide"), op.get("to"))
        if key in self.slide_keys:
            return None, "已提交过相同的幻灯片级操作"
        if action == "delete_slide":
            self.deleted_slides.add(op["slide"])
        op.update(
            {
                "id": f"op{len(self.ops) + 1}",
                "source": "agent",
                "before": None,
                "description": self._description(raw, defaults[0]),
                "meta": {"slide_index": op["slide_index"], "kind": "slide"},
            }
        )
        return op, None

    # ---------- v0.3 批量/增强操作 ----------

    def _v03_merged(self, raw: Dict[str, Any]) -> Dict[str, Any]:
        payload = raw.get("payload") if isinstance(raw.get("payload"), dict) else None
        return {**payload, **raw} if payload else raw

    def _v03_scope(self, merged: Dict[str, Any], deck_only: bool = False):
        """Resolve slide_index / deck eligibility. Returns (slide_index, scope_label, error)."""
        explicit = str(merged.get("scope") or "").strip().lower()
        slide_raw = merged.get("slide_index")
        if slide_raw is not None:
            index = _as_int(slide_raw)
            if index is None:
                return None, None, "slide_index 无效"
            if not self.view.is_scoped(index):
                return None, None, "slide_index 不在当前处理范围"
            if explicit in ("deck", "all", "document", "whole"):
                return None, None, "slide_index 与 scope=deck 冲突，二选一"
            return index, "slide", None
        if explicit in ("slide", "current", "page"):
            if self.view.default_slide is not None:
                return self.view.default_slide, "slide", None
            return None, None, "scope=slide 需要 slide_index"
        if not deck_only and explicit == "" and self.view.default_slide is not None:
            # 当前页范围且指令没写范围：默认只改这一页，全局需明说
            return self.view.default_slide, "slide", None
        if deck_only or explicit in ("deck", "all", "document", "whole"):
            if len(self.view.scope_slides) >= self.view.total_slides:
                return None, "deck", None  # full-deck scope: fine
            return None, None, _DECK_SCOPE_HINT
        if len(self.view.scope_slides) < self.view.total_slides:
            return None, None, _DECK_SCOPE_HINT
        return None, "deck", None

    def _v03_resolve(self, shape_id: str, kinds=None):
        """Find an object/element entry; returns (entry, slide_index, error)."""
        sid = str(shape_id or "").strip()
        if not sid:
            return None, None, "缺少 shape_id"
        for entry in (self.view.object_by_id.get(sid), self.view.element_by_id.get(sid)):
            if entry is None:
                continue
            raw_index = entry.get("slide_index")
            # int(x or -1) turns slide 0 into -1 — the falsy trap
            if raw_index is None:
                return None, None, "对象缺少 slide_index"
            slide_index = int(raw_index)
            if not self.view.is_scoped(slide_index):
                return None, None, "对象不在当前处理范围"
            return entry, slide_index, None
        return None, None, "对象不存在，请先 read_slides / inspect_shapes 确认"

    def _v03_op(
        self, raw: Dict[str, Any], action: str
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        merged = self._v03_merged(raw)
        defaults = {
            "font_unify": "统一字体",
            "replace_text": "全局替换文字",
            "align_objects": "对齐/分布对象",
            "duplicate_object": "复制对象",
            "z_order": "调整层级",
            "table_mod": "修改表格结构",
            "replace_image": "替换图片",
            "slide_background": "设置页面背景",
            "add_page_numbers": "批量添加页码",
            "font_scale": "统一调整字号",
        }
        op: Dict[str, Any] = {
            "id": f"op{len(self.ops) + 1}",
            "source": "agent",
            "action": action,
            "before": None,
            "description": self._description(raw, defaults[action]),
        }

        if action == "replace_image":
            shape_id = str(raw.get("shape_id") or raw.get("target_id") or "").strip()
            # source check first: a chat-side replace_image without any image
            # source can never work, so lead with that message regardless of
            # whether the target id resolves
            b64 = merged.get("image_base64")
            path = merged.get("image_path")
            if b64:
                if not isinstance(b64, str) or not b64.strip():
                    return None, "image_base64 必须是非空 base64 字符串"
                if len(b64) > 12_000_000:
                    return None, "image_base64 过大（解码前约 9 MB 上限）"
                op["image_base64"] = b64
            elif path:
                if not isinstance(path, str) or not path.strip():
                    return None, "image_path 必须是非空路径"
                op["image_path"] = path
            else:
                return None, "replace_image 需要 image_base64 或 image_path（服务器本地 png/jpg 路径）"
            entry, slide_index, error = self._v03_resolve(shape_id, kinds=("picture",))
            if error:
                return None, error
            if str(entry.get("kind")) != "picture":
                return None, "replace_image 只能作用于图片对象；请用 inspect_shapes 确认 shape_id"
            op["shape_id"] = shape_id
            op["slide_index"] = slide_index
            op["meta"] = {"slide_index": slide_index, "kind": "picture"}
            return self._finish_v03_op(op, raw, merged)

        if action in ("font_unify", "replace_text", "font_scale"):
            slide_index, scope_label, scope_error = self._v03_scope(merged)
            if scope_error:
                return None, scope_error
            op["slide_index"] = slide_index
            op["scope"] = scope_label
            if action == "font_unify":
                font_name = str(merged.get("font_name") or merged.get("font") or "").strip()
                if not font_name:
                    return None, "font_unify 缺少 font_name"
                if len(font_name) > 100:
                    return None, "font_name 过长（最多 100 字符）"
                op["font_name"] = font_name
                op["meta"] = {"kind": "font", "font_name": font_name}
            elif action == "replace_text":
                find = str(merged.get("find") or "").strip()
                replace = merged.get("replace")
                if not find:
                    return None, "replace_text 缺少 find"
                if len(find) > 200:
                    return None, "find 过长（最多 200 字符）"
                if replace is None:
                    return None, "replace_text 缺少 replace"
                replace = str(replace)
                if len(replace) > 2000:
                    return None, "replace 过长（最多 2000 字符）"
                op["find"] = find
                op["replace"] = replace
                op["meta"] = {"kind": "replace", "find": find, "replace": replace}
            else:  # font_scale
                delta = _as_float(
                    merged.get("delta_pt", merged.get("delta", merged.get("by")))
                )
                if delta is None:
                    return None, "font_scale 缺少 delta_pt（±50 以内）"
                if not (-50 <= delta <= 50):
                    return None, "delta_pt 超出 ±50 范围"
                op["delta_pt"] = int(delta) if float(delta).is_integer() else round(delta, 1)
                op["meta"] = {"kind": "font_scale", "delta_pt": op["delta_pt"]}
            return self._finish_v03_op(op, raw, merged)

        if action == "add_page_numbers":
            _slide_index, _scope_label, scope_error = self._v03_scope(merged, deck_only=True)
            if scope_error:
                return None, scope_error
            skip_first = _as_bool(merged.get("skip_first", merged.get("skip_cover")))
            op["skip_first"] = bool(skip_first)  # always explicit; False is meaningful
            position = str(merged.get("position") or "bottom-right").strip().lower()
            if position not in edits.PAGE_NUMBER_POSITIONS:
                return None, "position 仅支持 bottom-left/bottom-center/bottom-right"
            op["position"] = position
            start = _as_int(merged.get("start", 1))
            if start is None or not (0 <= start <= 9999):
                return None, "start 必须是 0..9999 的整数"
            op["start"] = start
            op["meta"] = {"kind": "page_numbers", "position": position, "skip_first": op["skip_first"]}
            return self._finish_v03_op(op, raw, merged)

        if action == "slide_background":
            slide_index, scope_error = self._v03_scope_deck_required(merged)
            if scope_error:
                return None, scope_error
            gradient = merged.get("gradient")
            if isinstance(gradient, dict):
                try:
                    edits._bg_color(gradient.get("from", gradient.get("start")), "gradient.from")
                    edits._bg_color(gradient.get("to", gradient.get("end")), "gradient.to")
                except edits.PatchError as exc:
                    return None, str(exc)
                op["gradient"] = gradient
            else:
                try:
                    color = edits._bg_color(merged.get("color"), "color")
                except edits.PatchError as exc:
                    return None, str(exc)
                op["color"] = color
            op["slide_index"] = slide_index
            op["meta"] = {"slide_index": slide_index, "kind": "background"}
            return self._finish_v03_op(op, raw, merged)

        if action == "align_objects":
            raw_ids = merged.get("shape_ids")
            if not isinstance(raw_ids, list) or len(raw_ids) < 2:
                return None, "align_objects 需要至少 2 个 shape_ids"
            if len(raw_ids) > 12:
                return None, "align_objects 单次最多 12 个对象"
            align = str(merged.get("align") or merged.get("mode") or "").strip().lower()
            align = edits._ALIGN_ALIASES.get(align, align)
            if align not in edits.ALIGN_MODES:
                return None, "align 仅支持 left/right/top/bottom/hcenter/vcenter/distribute_h/distribute_v"
            if align.startswith("distribute") and len(raw_ids) < 3:
                return None, "distribute_h/distribute_v 需要至少 3 个对象"
            slide_indices = set()
            shape_ids: List[str] = []
            for value in raw_ids:
                sid = str(value or "").strip()
                entry, slide_index, error = self._v03_resolve(sid)
                if error:
                    return None, f"{sid}：{error}"
                if entry.get("grouped"):
                    return None, f"{sid} 在分组内，不能参与对齐/分布"
                if entry.get("kind") == "table_cell":
                    return None, "表格单元格不能参与对齐/分布，请使用表格 shape_id"
                slide_indices.add(slide_index)
                shape_ids.append(sid)
            if len(slide_indices) != 1:
                return None, "align_objects 只支持同一页内的对象"
            slide_index = slide_indices.pop()
            op["slide_index"] = slide_index
            op["shape_ids"] = shape_ids
            op["align"] = align
            op["meta"] = {"slide_index": slide_index, "kind": "align", "align": align}
            return self._finish_v03_op(op, raw, merged)

        if action in ("duplicate_object", "z_order"):
            shape_id = str(raw.get("shape_id") or raw.get("target_id") or "").strip()
            entry, slide_index, error = self._v03_resolve(shape_id)
            if error:
                return None, error
            op["shape_id"] = shape_id
            op["slide_index"] = slide_index
            if action == "duplicate_object":
                offset = merged.get("offset_emu")
                if offset is not None:
                    if not isinstance(offset, dict):
                        return None, "offset_emu 必须是 {x, y} 对象"
                    x = _as_int(offset.get("x", 304800))
                    y = _as_int(offset.get("y", 304800))
                    if x is None or y is None or x < 0 or y < 0:
                        return None, "offset_emu.x/y 必须是非负整数（EMU）"
                    op["offset_emu"] = {"x": x, "y": y}
                op["meta"] = {"slide_index": slide_index, "kind": entry.get("kind")}
            else:
                mode = str(merged.get("mode") or "").strip().lower()
                mode = edits._Z_MODE_ALIASES.get(mode, mode)
                if mode not in edits._Z_ORDER_MODES:
                    return None, "z_order mode 仅支持 front/back/forward/backward（置顶/置底/上移/下移）"
                op["mode"] = mode
                op["meta"] = {"slide_index": slide_index, "kind": "z_order", "mode": mode}
            return self._finish_v03_op(op, raw, merged)

        if action == "table_mod":
            shape_id = str(raw.get("shape_id") or raw.get("table_shape_id") or "").strip()
            table = self.view.table_by_id.get(shape_id)
            if table is None:
                return None, "shape_id 不是表格；请先用 read_slides/read_table 找到表格 shape_id"
            slide_index = _slide_index_of(table)
            if not self.view.is_scoped(slide_index):
                return None, "表格不在当前处理范围"
            op_name = str(merged.get("op") or merged.get("table_op") or "").strip().lower()
            if op_name not in edits.TABLE_MOD_OPS:
                return None, "op 仅支持 add_row/add_col/delete_row/delete_col/merge_cells/split_cells"
            rows = int(table.get("rows") or 0)
            cols = int(table.get("cols") or 0)

            def _cell_index(*names: str, label: str) -> Optional[int]:
                for name in names:
                    if merged.get(name) is not None:
                        value = _as_int(merged.get(name))
                        if value is None:
                            raise edits.PatchError(f"{label} 无效")
                        return value
                return None

            try:
                if op_name == "delete_row":
                    index = _cell_index("index", "row", label="index")
                    if index is None or not (0 <= index < rows):
                        return None, f"index 必须在 0..{rows - 1}"
                    op["index"] = index
                elif op_name == "delete_col":
                    index = _cell_index("index", "col", label="index")
                    if index is None or not (0 <= index < cols):
                        return None, f"index 必须在 0..{cols - 1}"
                    op["index"] = index
                elif op_name == "merge_cells":
                    for name, high in (("row", rows), ("col", cols), ("row2", rows), ("col2", cols)):
                        value = _as_int(merged.get(name))
                        if value is None or not (0 <= value < high):
                            return None, f"{name} 必须在 0..{high - 1}"
                        op[name] = value
                elif op_name == "split_cells":
                    for name, high in (("row", rows), ("col", cols)):
                        value = _as_int(merged.get(name))
                        if value is None or not (0 <= value < high):
                            return None, f"{name} 必须在 0..{high - 1}"
                        op[name] = value
            except edits.PatchError as exc:
                return None, str(exc)
            op["shape_id"] = shape_id
            op["slide_index"] = slide_index
            op["op"] = op_name
            op["meta"] = {"slide_index": slide_index, "kind": "table_mod", "op": op_name}
            return self._finish_v03_op(op, raw, merged)

        return None, f"不支持的操作：{action}"

    def _v03_scope_deck_required(self, merged: Dict[str, Any]):
        """slide_background needs an explicit, in-scope slide_index."""
        slide_raw = merged.get("slide_index")
        index = _as_int(slide_raw)
        if index is None:
            return None, "slide_background 需要 slide_index"
        if not self.view.is_scoped(index):
            return None, "slide_index 不在当前处理范围"
        return index, None

    def _finish_v03_op(
        self, op: Dict[str, Any], raw: Dict[str, Any], merged: Dict[str, Any]
    ) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
        op.setdefault("meta", {})
        if "slide_index" in op:
            op["meta"].setdefault("slide_index", op["slide_index"])
        key = (
            op.get("action"),
            op.get("slide_index"),
            op.get("shape_id"),
            tuple(op.get("shape_ids") or []),
            op.get("font_name"),
            op.get("find"),
            op.get("replace"),
            op.get("delta_pt"),
            op.get("align"),
            op.get("mode"),
            op.get("op"),
            op.get("position"),
            op.get("color"),
            json.dumps(op.get("gradient"), ensure_ascii=False, sort_keys=True) if op.get("gradient") else None,
        )
        if key in self.v03_keys:
            return None, "已提交过相同的批量操作"
        self.v03_keys.add(key)
        return op, None

    def _register(self, op: Dict[str, Any]) -> None:
        action = op.get("action")
        shape_id = str(op.get("shape_id") or "")
        if action in TEXT_ACTIONS and shape_id:
            self.text_targets.add(shape_id)
        elif action == "style" and shape_id:
            self.style_targets.add(shape_id)
        elif action == "geometry" and shape_id:
            self.geometry_targets.add(shape_id)
        elif action == "add_textbox":
            bbox = op.get("bbox_emu") or {}
            self.add_keys.add(
                (
                    op.get("slide_index"),
                    _normalize_text(op.get("after")),
                    bbox.get("x"),
                    bbox.get("y"),
                    bbox.get("w"),
                    bbox.get("h"),
                )
            )
        elif action in SLIDE_ACTIONS:
            self.slide_keys.add(
                (
                    action,
                    op.get("source_slide"),
                    op.get("position"),
                    op.get("after_slide"),
                    op.get("layout_index"),
                    op.get("slide"),
                    op.get("to"),
                )
            )

    def _op_summary(self, op: Dict[str, Any]) -> Dict[str, Any]:
        item = {
            "id": op.get("id"),
            "action": op.get("action"),
            "slide_index": op.get("slide_index"),
            "description": op.get("description"),
        }
        if op.get("shape_id"):
            item["shape_id"] = op.get("shape_id")
        if op.get("scope"):
            item["scope"] = op.get("scope")
        if op.get("action") == "add_textbox":
            item["text"] = _clip(op.get("after"), 80)
        if op.get("action") in SLIDE_ACTIONS:
            for key in ("source_slide", "position", "after_slide", "layout_index", "slide", "to"):
                if op.get(key) is not None:
                    item[key] = op.get(key)
        if op.get("action") in V03_ACTIONS:
            for key in (
                "font_name", "find", "replace", "delta_pt", "align", "mode", "op",
                "shape_ids", "skip_first", "position", "start", "color", "gradient",
                "offset_emu",
            ):
                if op.get(key) is not None:
                    item[key] = op.get(key)
        return item


def _canonical_action(value: Any) -> str:
    key = str(value or "").strip().lower()
    if not key:
        return ""
    return ACTION_ALIASES.get(key, key)


def _extract_args(parsed: Dict[str, Any]) -> Dict[str, Any]:
    for key in ("args", "arguments", "input", "params"):
        value = parsed.get(key)
        if isinstance(value, dict):
            return value
    return {}


def _args_with_fallback(args: Dict[str, Any], parsed: Dict[str, Any]) -> Dict[str, Any]:
    if args:
        return args
    reserved = {
        "thought",
        "reasoning",
        "action",
        "tool",
        "name",
        "args",
        "arguments",
        "input",
        "params",
    }
    return {key: value for key, value in parsed.items() if key not in reserved}


def _extract_operations(parsed: Dict[str, Any], args: Dict[str, Any]) -> Any:
    for source in (args, parsed):
        if isinstance(source, dict):
            for key in ("operations", "patches", "ops"):
                if isinstance(source.get(key), list):
                    return source.get(key)
    return None


def _trace_detail(action: str, args: Dict[str, Any], detail: Dict[str, Any]) -> str:
    if detail.get("error"):
        return _clip(detail.get("error"), 200)
    if action == "outline":
        return f"读取大纲（{detail.get('total_slides', 0)} 页）"
    if action == "read_slides":
        pages = []
        for section in detail.get("slides") or []:
            if isinstance(section, dict):
                pages.append(str(section.get("page") or ((_slide_index_of(section) if section.get("slide_index") is not None else 0) + 1)))
        return f"读取第 {', '.join(pages)} 页" if pages else "读取页面"
    if action == "search_text":
        return f"搜索「{detail.get('query', '')}」命中 {detail.get('count', 0)} 处"
    if action == "read_table":
        return (
            f"读取表格 {detail.get('table_shape_id', '')}"
            f"（{detail.get('rows', 0)}×{detail.get('cols', 0)}）"
        )
    if action == "inspect_shapes":
        shapes = detail.get("shapes") or []
        return f"查看 {len(shapes)} 个对象"
    if action == "propose_operations":
        text = f"提交 {detail.get('accepted', 0)} 条，累计 {detail.get('total', 0)} 条"
        if detail.get("dropped"):
            text += f"，丢弃 {detail.get('dropped')} 条"
        return text
    if action == "finish":
        return _clip(detail.get("summary") or detail.get("reply") or "完成", 200)
    return _clip(str(detail), 160)


def _clip_content(content: str, limit: int = 12000) -> str:
    """Clip a conversation turn for the context window. Tool observations are
    JSON; clipping the raw string mid-value feeds the model malformed JSON, so
    clip per-field when the content parses and fall back to raw truncation."""
    if len(content) <= limit:
        return content
    try:
        data = json.loads(content)
    except (ValueError, TypeError):
        return content[: limit - 1] + "…"
    if isinstance(data, dict):
        field_limit = max(200, limit // 4)

        def clip_value(value: Any) -> Any:
            if isinstance(value, str) and len(value) > field_limit:
                return value[: field_limit - 1] + "…"
            if isinstance(value, list):
                return [clip_value(item) for item in value[:40]]
            if isinstance(value, dict):
                return {key: clip_value(item) for key, item in value.items()}
            return value

        out = json.dumps({key: clip_value(item) for key, item in data.items()}, ensure_ascii=False)
        if len(out) <= limit:
            return out
    return content[: limit - 1] + "…"


def _trim_messages(messages: List[Dict[str, str]]) -> List[Dict[str, str]]:
    if len(messages) <= 2:
        return messages
    head = messages[:2]
    rest = messages[2:]
    if len(rest) > 8:
        rest = rest[-8:]
    trimmed: List[Dict[str, str]] = []
    for message in rest:
        content = str(message.get("content") or "")
        if len(content) > 12000:
            content = _clip_content(content)
        trimmed.append({**message, "content": content})
    return head + trimmed


SYSTEM_PROMPT = (
    "你是一个 AI 主导的 PPT 精准编辑 Agent。你不能直接修改文件，只能调用工具读取内容并提交操作；"
    "操作由系统在用户确认后应用（调用方也可以按配置自动应用）。\n\n"
    "每一步只输出一个 JSON 对象，不要 Markdown、不要代码块、不要额外解释。格式：\n"
    '{"thought":"一句话判断","action":"工具名","args":{...}}\n\n'
    "可用工具：\n"
    "1. outline {}：整份文稿大纲，含页码、标题、文本/表格/视觉元素数量。\n"
    '2. read_slides {"slide_indices":[0,1]}：读取范围内最多 4 页的详细对象。文本与单元格包含 '
    "shape_id、类型、分组、bbox_cm、字体字号/对齐；单元格含行列与合并跨度；另附表格与图片/图表/组合摘要。\n"
    '3. search_text {"query":"关键词"}：在范围内搜索原文。\n'
    '4. read_table {"table_shape_id":"s3_id4"}：读取表格行列与非空单元格（cell shape_id、行列、合并跨度、文本）。\n'
    '5. inspect_shapes {"shape_ids":["s3_id4_c0_0"]}：查看指定文本/单元格/视觉对象的精确信息，'
    "含完整文本与 bbox_emu。\n"
    '6. propose_operations {"operations":[...]}：提交精确操作，系统会逐条校验。\n'
    '7. plan {"steps":["第一步","第二步","第三步"]}：复杂任务（≥2 步）先制定计划清单，'
    "系统会把「我将：1… 2… 3…」展示给用户；之后逐步执行、逐项汇报，完成后在 finish 里总结各项结果。\n"
    '8. undo_last {}：撤销上一轮已应用的修改（回退到上一版本快照）。用户要求“撤销/回退上一步”时调用，'
    "撤销结果以 Observation 返回，不要重复撤销。\n"
    '9. finish {"summary":"...","reply":"给用户的回复","question":"可选反问"}：结束。\n\n'
    "操作 schema：\n"
    '- edit：{"action":"edit","shape_id":"...","after":"新文本","before":"可选原文",'
    '"description":"...","style":{"font_size_pt":18,"font_name":"微软雅黑","bold":true,'
    '"italic":false,"color":"RRGGBB","alignment":"LEFT"}}\n'
    '- style：{"action":"style","shape_id":"...","style":{...},"description":"..."}\n'
    '- clear：{"action":"clear","shape_id":"...","description":"..."} 只清空文本，不删除对象。\n'
    '- geometry：{"action":"geometry","shape_id":"...","bbox_emu":{"x":0,"y":0,"w":100,"h":100},'
    '"description":"..."} 仅限非分组对象/元素；不能用于表格单元格；1 厘米 = 360000 EMU。\n'
    '- add_textbox：{"action":"add_textbox","slide_index":3,"after":"文本",'
    '"bbox_emu":{"x":0,"y":0,"w":100,"h":100},"font_size_pt":18,"bold":false,'
    '"color":"1F4E79","alignment":"LEFT"}\n'
    '- delete_textbox：{"action":"delete_textbox","shape_id":"..."} 仅限非分组的普通文本对象，'
    "不能删除分组内文字或表格单元格。\n"
    "幻灯片级操作（页码一律用 0-based slide_index，与 shape_id 前缀一致）：\n"
    '- duplicate_slide：{"action":"duplicate_slide","source_slide":0,"position":1,"description":"复制第1页"} '
    "复制整页（文字、样式、图片、表格、位置、备注全保留），position 默认插在源页后面。\n"
    '- add_slide：{"action":"add_slide","after_slide":3,"layout_index":6} 新增空白页，默认加到末尾；'
    "layout_index 可选。\n"
    '- delete_slide：{"action":"delete_slide","slide":2} 删除整页（唯一一页时不允许）。\n'
    '- move_slide：{"action":"move_slide","slide":4,"to":0} 把某页移动到指定位置。\n'
    "幻灯片级操作规则：同一批里页内编辑（edit/style/geometry/…）按操作前的页码解析，"
    "幻灯片级操作在其后按提交顺序执行；应用后所有 s{n}_id{m} 按新页码重编号，"
    "要编辑新插入页的内容，请等本批应用后下一轮重新 read_slides 再提交。\n"
    "典型用法：「复制第 1 页生成第 2 页并改标题」→ 本轮只提交 duplicate_slide，"
    "应用后下一轮再 edit 新页的标题对象。\n"
    "v0.3 批量/增强操作：\n"
    '- font_unify：{"action":"font_unify","font_name":"微软雅黑","slide_index":可选} '
    "统一字体（保留字号/加粗差异）；不填 slide_index 作用于整份。\n"
    '- replace_text：{"action":"replace_text","find":"旧词","replace":"新词"} 跨页全量替换，'
    "建议先 search_text 确认命中位置；找不到会报错。\n"
    '- align_objects：{"action":"align_objects","shape_ids":["s0_id2","s0_id3"],'
    '"align":"left|right|top|bottom|hcenter|vcenter|distribute_h|distribute_v"} '
    "同页非分组对象对齐/等间距分布（分布需 ≥3 个）。\n"
    '- duplicate_object：{"action":"duplicate_object","shape_id":"...","offset_emu":{"x":300000,"y":300000}} '
    "页内复制文本框/图片等（默认向右下偏移约 0.85cm）。\n"
    '- z_order：{"action":"z_order","shape_id":"...","mode":"front|back|forward|backward"} 置顶/置底/上移/下移。\n'
    '- table_mod：{"action":"table_mod","shape_id":"表格id","op":"add_row|add_col|delete_row|delete_col|'
    'merge_cells|split_cells","index/row/col/row2/col2":...} 表格结构修改（行列增删、合并/拆分）。\n'
    '- slide_background：{"action":"slide_background","slide_index":0,"color":"1F4E79"}，'
    '或 {"gradient":{"from":"0F172A","to":"2563EB","angle":45}} 纯色/渐变背景。\n'
    '- add_page_numbers：{"action":"add_page_numbers","skip_first":true,"position":"bottom-right"} '
    "批量加页码（skip_first 可跳过封面）。\n"
    '- font_scale：{"action":"font_scale","delta_pt":2} 整页/整份字号统一 +N/-N 磅。\n'
    "批量操作规则：font_unify/replace_text/font_scale 不填 slide_index、以及 add_page_numbers，"
    "都会作用于整份文稿，需要整份范围（当前页范围下，指令含「全局/整份/所有页」且没有点名具体页面时，"
    "系统会自动扩大范围；否则需用户把范围切到「整份文稿」）；"
    "replace_image 需要图片来源（image_base64 或服务器可读的 image_path，png/jpg）；"
    "对话里拿不到图片来源时，用 finish 引导用户通过 /api/docs/{id}/apply 提供图片。\n\n"
    f"约束：单次 propose 最多 {MAX_OPS_PER_CALL} 条，本次对话最多 {MAX_OPS} 条。\n\n"
    "工作方式：\n"
    "- 先理解意图：用户点名字段或页面时优先处理该对象/页面；模糊需求先 outline/read_slides 判断。\n"
    "- 复杂任务（涉及多页或多种操作组合）先用 plan 拆解成计划清单，再逐步执行、逐项汇报进度。\n"
    "- 先看再改：shape_id 必须来自工具结果；表格单元格必须用 s3_id4_c0_2 这类 ID，"
    "合并区域使用合并起始单元格。\n"
    "- 只提交有必要、可执行的少量高价值操作；不需要每页都读。整份优化时先 outline，"
    "再按页 read_slides（每次最多 4 页）。\n"
    "- 保持原意与事实，不编造数据、案例、人名、机构或结论；保留数字、单位、专有名词、英文缩写。\n"
    "- 保持版式：默认只改文字；只有用户明确要求时才改字体、字号、颜色、位置或尺寸。"
    "新增文本框必须给精确 EMU 坐标。\n"
    "- 未经明确点名，不要修改 logo、页码、页脚、页眉、版权、日期。\n"
    "- before 可省略，系统会用真实原文校验；若填写必须与原文一致。after 为空表示清空，"
    "系统会按 clear 处理。\n"
    "- 信息不足、目标不明确或需要在多个方案中取舍时，用 finish 的 question 反问，不要猜测。\n"
    "- 如果用户只是询问内容，直接用 finish 的 reply 回答，不提交操作。\n"
    "- 结束前必须调用 finish；没有更多修改时也要 finish。每次只输出一个合法 JSON 对象。"
)


def _history_lines(history: Optional[List[Dict[str, Any]]]) -> List[str]:
    lines: List[str] = []
    for item in (history or [])[-10:]:
        if not isinstance(item, dict):
            continue
        role = str(item.get("role") or "")
        if role not in ("user", "assistant"):
            continue
        content = item.get("content") or item.get("reply") or item.get("summary") or ""
        content = ai.clean_text(content, 320)
        if not content:
            continue
        lines.append(f"{'用户' if role == 'user' else '助手'}：{content}")
    return lines[-8:]


def _prompt_safe_detail(detail: Dict[str, Any]) -> Dict[str, Any]:
    safe = dict(detail)
    if "text" in safe:
        safe["text"] = _clip(safe.get("text"), 500)
    return safe


def _initial_prompt(
    view: _DeckView,
    instruction: str,
    scope: str,
    selected_shape_id: Optional[str],
    history: Optional[List[Dict[str, Any]]],
) -> str:
    if len(view.scope_slides) == 1:
        index = view.scope_slides[0]
        scope_desc = f"第 {index + 1} 页（slide_index={index}）"
    elif scope == "deck":
        scope_desc = f"整份文稿（{view.total_slides} 页）"
    elif view.scope_slides:
        scope_desc = f"第 {view.scope_slides[0] + 1}–{view.scope_slides[-1] + 1} 页"
    else:
        scope_desc = "无可用页面"

    outline_lines: List[str] = []
    for index in range(view.total_slides):
        title = _slide_title(view.objects_by_slide.get(index, []))
        marker = "*" if index in view.scope_set else " "
        outline_lines.append(f"{marker} 第{index + 1}页 slide_index={index}：{title}")

    parts = [
        f"演示文稿共 {view.total_slides} 页；本次处理范围：{scope_desc}。",
        f"可编辑 slide_index：{view.scope_slides}。范围外的页面不能修改。",
    ]

    if selected_shape_id:
        selection = view.inspect(str(selected_shape_id))
        if selection.get("error"):
            selection = {"shape_id": selected_shape_id, "error": selection.get("error")}
        elif not view.is_scoped(_slide_index_of(selection)):
            selection = {
                "shape_id": selected_shape_id,
                "note": "已选对象不在本次处理范围，仅作参考",
            }
        else:
            selection = _prompt_safe_detail(selection)
        parts.append("当前选中对象（仅供参考）：\n" + json.dumps(selection, ensure_ascii=False))

    history_lines = _history_lines(history)
    if history_lines:
        parts.append("最近对话：\n" + "\n".join(history_lines))

    parts.append("页面标题大纲（* 表示本次范围）：\n" + "\n".join(outline_lines))
    parts.append(f"用户最新要求：{instruction}")
    parts.append(
        "请先判断用户意图，选择下一步工具。需要精确信息时先读取；可以直接 propose_operations；"
        "完成后必须 finish。只输出一个 JSON 对象。"
    )
    return "\n\n".join(parts)


def run_stream(
    emit: EventEmitter,
    *,
    objects: List[Dict[str, Any]],
    instruction: str,
    total_slides: int,
    slide_size: Optional[Dict[str, Any]] = None,
    elements: Optional[List[Dict[str, Any]]] = None,
    tables: Optional[List[Dict[str, Any]]] = None,
    scope: str = "current",
    slide_index: int = 0,
    selected_shape_id: Optional[str] = None,
    history: Optional[List[Dict[str, Any]]] = None,
    base_version: Optional[int] = None,
    auto_apply: bool = False,
    undo_last: Optional[Callable[[Optional[int]], Dict[str, Any]]] = None,
    undo_handler: Optional[Callable[[], Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Run the streaming AI-led tool loop and return the done payload.

    Events emitted synchronously through ``emit``:
    start, step, token, thought, tool_call, tool_result, operations, plan,
    retry, summary, error, done.
    """
    instruction = str(instruction or "").strip()
    if not instruction:
        raise ValueError("请先输入修改要求")
    instruction = instruction[:4000]

    objects = [obj for obj in (objects or []) if isinstance(obj, dict)]
    elements = [elem for elem in (elements or []) if isinstance(elem, dict)]
    tables = [table for table in (tables or []) if isinstance(table, dict)]

    max_index = -1
    for item in objects + elements:
        value = _as_int(item.get("slide_index"))
        if value is not None:
            max_index = max(max_index, value)
    try:
        total_slides = int(total_slides)
    except (TypeError, ValueError):
        total_slides = 0
    total_slides = max(total_slides, max_index + 1)
    if total_slides <= 0:
        raise ValueError("演示文稿没有幻灯片")

    raw_scope = str(scope or "").strip().lower()
    scope = "deck" if raw_scope in ("deck", "all", "document", "whole") else "current"
    requested = ai.extract_requested_slide(instruction, total_slides)
    if requested is not None and scope != "deck":
        # 点名页面只收窄「当前页」范围。用户明确选了整份时不降级：多步指令常常
        # 既点名某页又包含整份步骤，降级会让整份步骤全被范围校验挡下
        scope = "current"
        target_slide = requested
    else:
        value = _as_int(slide_index)
        target_slide = 0 if value is None else value
        target_slide = max(0, min(target_slide, total_slides - 1))
        if scope == "current" and DECK_INTENT_RE.search(instruction):
            # 用户点名「全局/整份/所有页…」→ 自动扩大范围，允许 deck 级批量操作
            scope = "deck"
    scope_slides = [target_slide] if scope == "current" else list(range(total_slides))

    view = _DeckView(objects, elements, tables, total_slides, scope_slides, slide_size)
    store = _OperationStore(view, instruction, MAX_OPS)
    model = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")

    # undo_last 工具的后端回调：优先带 base_version 的版本（API 层注入），
    # 兼容无参的 undo_handler（测试/嵌入场景）。
    undo_callback: Optional[Callable[[], Dict[str, Any]]] = None
    if undo_last is not None:
        undo_callback = lambda: undo_last(base_version)  # noqa: E731
    elif undo_handler is not None:
        undo_callback = undo_handler

    today = datetime.now().astimezone().strftime("%Y年%m月%d日")
    conversation: List[Dict[str, str]] = [
        {
            "role": "system",
            "content": (
                f"{SYSTEM_PROMPT}\n\n当前日期：{today}。"
                "用户说“今天/明天/本月”等相对日期时，以该日期为准换算成具体日期。"
            ),
        },
        {
            "role": "user",
            "content": _initial_prompt(view, instruction, scope, selected_shape_id, history),
        },
    ]

    summary = ""
    reply = ""
    question = ""
    error: Optional[str] = None
    finished = False
    steps_run = 0
    trace: List[Dict[str, Any]] = []
    plan: List[str] = []
    undo_result: Optional[Dict[str, Any]] = None

    _safe_emit(
        emit,
        "start",
        {
            "mode": "agent",
            "model": model,
            "scope": scope,
            "scope_slides": scope_slides,
            "total_slides": total_slides,
            "max_steps": MAX_STEPS,
            "max_operations": MAX_OPS,
            "base_version": base_version,
            "auto_apply": bool(auto_apply),
        },
    )

    for step in range(1, MAX_STEPS + 1):
        steps_run = step
        _safe_emit(emit, "step", {"step": step, "max_steps": MAX_STEPS})
        parsed = None
        raw_content = ""
        # 自愈：瞬时 AI 服务错误（网络波动/限流）自动重试一次，而不是把
        # 错误直接甩给用户；重试仍失败才中断本轮对话
        for attempt in (1, 2):
            try:
                parsed, raw_content = streaming.stream_chat_json(
                    _trim_messages(conversation),
                    on_token=lambda piece, current_step=step: _safe_emit(
                        emit, "token", {"step": current_step, "delta": piece, "text": piece}
                    ),
                    temperature=0.2,
                    max_tokens=4000,
                )
                break
            except ai.AIServiceError as exc:
                if attempt == 1:
                    _safe_emit(
                        emit,
                        "retry",
                        {
                            "step": step,
                            "attempt": 2,
                            "message": f"AI 服务波动，自动重试一次：{exc}",
                        },
                    )
                    trace.append(
                        {"step": step, "thought": "AI 服务波动，自动重试", "action": "retry", "detail": _clip(str(exc), 200)}
                    )
                    continue
                error = f"AI 服务异常：{exc}"
                _safe_emit(emit, "error", {"step": step, "message": error})
                trace.append(
                    {"step": step, "thought": "", "action": "error", "detail": _clip(error, 200)}
                )
                break
            except Exception as exc:
                error = f"AI 服务异常：{exc}"
                _safe_emit(emit, "error", {"step": step, "message": error})
                trace.append(
                    {"step": step, "thought": "", "action": "error", "detail": _clip(error, 200)}
                )
                break
        if parsed is None and error:
            break

        assistant_content = str(raw_content or "").strip()
        if not assistant_content:
            assistant_content = json.dumps(parsed, ensure_ascii=False)
        conversation.append({"role": "assistant", "content": assistant_content})

        thought = ai.clean_text(parsed.get("thought") or parsed.get("reasoning") or "", 300)
        if thought:
            _safe_emit(emit, "thought", {"step": step, "thought": thought, "text": thought})

        raw_action = parsed.get("action")
        if raw_action is None:
            raw_action = parsed.get("tool")
        if raw_action is None:
            raw_action = parsed.get("name")
        action = _canonical_action(raw_action)
        args = _extract_args(parsed)
        operations_raw = _extract_operations(parsed, args)
        if not action:
            if operations_raw is not None:
                action = "propose_operations"
            elif any(key in parsed for key in ("summary", "reply", "question")):
                action = "finish"
            else:
                action = "unknown"
        args = _args_with_fallback(args, parsed)
        if action == "propose_operations" and operations_raw is None:
            operations_raw = _extract_operations(parsed, args)

        _safe_emit(
            emit,
            "tool_call",
            {
                "step": step,
                "action": action or "unknown",
                "tool": action or "unknown",
                "args": args,
            },
        )

        if action == "plan":
            # 多步任务规划：先亮出清单再逐步执行，UI 以 plan 事件渲染进度
            raw_steps = args.get("steps") if isinstance(args.get("steps"), list) else None
            if raw_steps is None and isinstance(parsed.get("steps"), list):
                raw_steps = parsed.get("steps")
            cleaned: List[str] = []
            for item in raw_steps or []:
                text = ai.clean_text(item, 200)
                if text and text not in cleaned:
                    cleaned.append(text)
                if len(cleaned) >= 12:
                    break
            if len(cleaned) < 2:
                # 单步/空清单不算计划：直接执行即可，别为凑数拆步骤
                detail = {"error": "plan 需要 steps（至少 2 条非空字符串）；单步任务请直接调用对应工具"}
            else:
                plan = cleaned
                detail = {
                    "plan": plan,
                    "note": "计划已记录，请逐步执行并逐项汇报；完成后 finish 里总结每项结果",
                }
                _safe_emit(emit, "plan", {"step": step, "steps": plan})
            _safe_emit(
                emit,
                "tool_result",
                {
                    "step": step,
                    "action": "plan",
                    "tool": "plan",
                    "detail": f"制定 {len(plan)} 步计划" if plan else "计划为空",
                    "result": detail,
                },
            )
            trace.append(
                {"step": step, "thought": thought, "action": "plan", "detail": f"制定 {len(plan)} 步计划"}
            )
            observation = "Observation:\n" + json.dumps(detail, ensure_ascii=False)
            if detail.get("error"):
                observation += "\n请修正参数后重试（steps 需为非空字符串数组）。"
            conversation.append({"role": "user", "content": observation})
            continue

        if action == "undo_last":
            # 撤销上一轮已应用的操作：由调用方注入基于版本快照的回退器
            if undo_callback is None:
                detail = {"error": "当前环境不支持撤销（未接入版本回退）"}
            else:
                try:
                    detail = undo_callback() or {}
                    undo_result = detail if not detail.get("error") else None
                except Exception as exc:  # noqa: BLE001 — handler errors become tool errors
                    detail = {"error": str(exc) or "撤销失败"}
            _safe_emit(
                emit,
                "tool_result",
                {
                    "step": step,
                    "action": "undo_last",
                    "tool": "undo_last",
                    "detail": _trace_detail("undo_last", args, detail),
                    "result": detail,
                },
            )
            trace.append(
                {
                    "step": step,
                    "thought": thought,
                    "action": "undo_last",
                    "detail": _trace_detail("undo_last", args, detail),
                }
            )
            observation = "Observation:\n" + json.dumps(detail, ensure_ascii=False)
            if detail.get("error"):
                observation += "\n撤销未执行；请向用户说明原因，不要重复尝试。"
            conversation.append({"role": "user", "content": observation})
            continue

        if action == "finish":
            raw_plan = args.get("plan") if isinstance(args.get("plan"), list) else parsed.get("plan")
            if isinstance(raw_plan, list) and not plan:
                plan = [ai.clean_text(item, 200) for item in raw_plan if ai.clean_text(item, 200)][:12]
            summary = ai.clean_text(args.get("summary") or parsed.get("summary") or "", 500)
            reply = ai.clean_text(
                args.get("reply") or args.get("message") or parsed.get("reply") or "", 800
            )
            question = ai.clean_text(args.get("question") or parsed.get("question") or "", 500)
            detail = {
                "summary": summary,
                "reply": reply,
                "question": question,
                "operations_total": len(store.ops),
            }
            _safe_emit(
                emit,
                "tool_result",
                {
                    "step": step,
                    "action": "finish",
                    "tool": "finish",
                    "detail": _trace_detail("finish", args, detail),
                    "result": detail,
                },
            )
            trace.append(
                {
                    "step": step,
                    "thought": thought,
                    "action": "finish",
                    "detail": _trace_detail("finish", args, detail),
                }
            )
            finished = True
            break

        if action not in TOOL_NAMES:
            detail = {"error": "未知工具", "allowed": sorted(TOOL_NAMES)}
            _safe_emit(
                emit,
                "tool_result",
                {
                    "step": step,
                    "action": action or "unknown",
                    "tool": action or "unknown",
                    "detail": "未知工具",
                    "result": detail,
                },
            )
            trace.append(
                {
                    "step": step,
                    "thought": thought,
                    "action": action or "unknown",
                    "detail": "未知工具",
                }
            )
            conversation.append({"role": "user", "content": _observation_content(detail)})
            continue

        before_count = len(store.ops)
        try:
            if action == "propose_operations":
                detail = store.add_batch(operations_raw)
            else:
                detail = _dispatch_tool(action, args, view)
        except Exception as exc:
            detail = {"error": f"工具执行失败：{exc}"}
        new_ops = store.ops[before_count:] if action == "propose_operations" else []
        if new_ops:
            _safe_emit(
                emit,
                "operations",
                {"step": step, "operations": new_ops, "total": len(store.ops)},
            )
        _safe_emit(
            emit,
            "tool_result",
            {
                "step": step,
                "action": action,
                "tool": action,
                "detail": _trace_detail(action, args, detail),
                "result": detail,
            },
        )
        trace.append(
            {
                "step": step,
                "thought": thought,
                "action": action,
                "detail": _trace_detail(action, args, detail),
            }
        )
        observation = "Observation:\n" + json.dumps(detail, ensure_ascii=False)
        if isinstance(detail, dict) and detail.get("error"):
            # 自愈提示：让模型换参数重试一次，而不是把错误原样抛给用户
            observation += (
                "\n该调用失败：请修正参数后重试一次（换一种工具或参数也可以）；"
                "若确实无法完成，用 finish 向用户说明，不要编造结果。"
            )
        conversation.append({"role": "user", "content": observation})

    operations = store.ops
    if not summary:
        if steps_run >= MAX_STEPS and not finished:
            summary = (
                f"已达到最大步数，已收集 {len(operations)} 条待应用操作。"
                if operations
                else "已达到最大步数，未生成修改操作。"
            )
        elif operations:
            summary = f"已收集 {len(operations)} 条待应用的修改操作。"
        elif question:
            summary = "AI 需要补充信息后再继续。"
        elif error:
            summary = "AI 调用中断，未生成修改操作。"
        else:
            summary = "已检查当前范围，没有需要执行的修改。"
    if not reply:
        if question:
            reply = question
        elif operations:
            reply = f"已准备 {len(operations)} 条修改操作，等待确认应用。"
        elif error:
            reply = f"处理中断：{error}"
        else:
            reply = "已检查当前范围，没有需要执行的修改。如需调整，请说明目标对象或期望效果。"
    summary = ai.clean_text(summary, 500)
    reply = ai.clean_text(reply, 800)
    question = ai.clean_text(question, 500)

    _safe_emit(
        emit,
        "summary",
        {
            "summary": summary,
            "reply": reply,
            "question": question,
            "operations": len(operations),
        },
    )
    done: Dict[str, Any] = {
        "mode": "agent",
        "model": model,
        "scope": scope,
        "scope_slides": scope_slides,
        "total_slides": total_slides,
        "summary": summary,
        "reply": reply,
        "question": question,
        "operations": operations,
        "trace": trace,
        "steps": steps_run,
        "max_steps": MAX_STEPS,
        "finished": finished,
        "dropped": store.dropped,
        "error": error,
        "base_version": base_version,
        "auto_apply": bool(auto_apply),
        "plan": plan or None,
        "undo": undo_result,
    }
    _safe_emit(emit, "done", done)
    return done
