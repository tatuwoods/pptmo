from __future__ import annotations

import copy
import difflib
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt


class PatchError(ValueError):
    pass


class PatchConflict(PatchError):
    pass


CELL_ID_RE = re.compile(r"^(s\d+_id\d+)_c(\d+)_(\d+)$")


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def parse_shape_id(shape_id: str) -> Optional[Tuple[int, str, int]]:
    match = re.fullmatch(r"s(\d+)_id(\d+)", str(shape_id or ""))
    if match:
        return int(match.group(1)), "id", int(match.group(2))
    match = re.fullmatch(r"s(\d+)_sh(\d+)", str(shape_id or ""))
    if match:
        return int(match.group(1)), "index", int(match.group(2))
    return None


def parse_cell_id(shape_id: str) -> Optional[Tuple[str, int, int]]:
    match = CELL_ID_RE.fullmatch(str(shape_id or ""))
    if match:
        return match.group(1), int(match.group(2)), int(match.group(3))
    return None


def _iter_shapes(shapes):
    for shape in shapes:
        yield shape
        try:
            if getattr(shape, "shape_type", None) == 6:
                yield from _iter_shapes(shape.shapes)
        except Exception:
            continue


def find_shape(slide, shape_id: str):
    parsed = parse_shape_id(shape_id)
    if not parsed:
        return None
    _, kind, value = parsed
    if kind == "index":
        shapes = list(slide.shapes)
        if 0 <= value < len(shapes):
            return shapes[value]
        return None
    for shape in _iter_shapes(slide.shapes):
        try:
            if int(getattr(shape, "shape_id", -1)) == value:
                return shape
        except (TypeError, ValueError):
            continue
    return None


def find_table_cell(slide, shape_id: str):
    parsed = parse_cell_id(shape_id)
    if not parsed:
        return None
    table_shape_id, row, col = parsed
    table_shape = find_shape(slide, table_shape_id)
    if table_shape is None or not getattr(table_shape, "has_table", False):
        return None
    try:
        table = table_shape.table
        if row < 0 or row >= len(table.rows) or col < 0 or col >= len(table.columns):
            return None
        return table.cell(row, col)
    except Exception:
        return None


def _font_style(shape) -> Dict[str, Any]:
    style: Dict[str, Any] = {}
    try:
        tf = shape.text_frame
        paragraph = tf.paragraphs[0]
        font = paragraph.runs[0].font if paragraph.runs else paragraph.font
        style["name"] = font.name
        style["size"] = font.size
        style["bold"] = font.bold
        style["italic"] = font.italic
        try:
            if font.color is not None and font.color.rgb is not None:
                style["rgb"] = font.color.rgb
        except Exception:
            pass
    except Exception:
        pass
    return style


def _apply_run_style(run, style: Dict[str, Any]) -> None:
    try:
        font = run.font
        if style.get("name"):
            font.name = style["name"]
        if style.get("size"):
            font.size = style["size"]
        if style.get("bold") is not None:
            font.bold = style["bold"]
        if style.get("italic") is not None:
            font.italic = style["italic"]
        if style.get("rgb") is not None:
            font.color.rgb = style["rgb"]
    except Exception:
        pass


def set_text_preserving_style(shape_or_cell, text: Any) -> None:
    """Replace text while keeping the first paragraph/run formatting.

    Limitation: mixed formatting inside one paragraph collapses to the first run.
    """
    value = str(text if text is not None else "").replace("\r\n", "\n").replace("\v", "\n")
    lines = value.split("\n")
    tf = shape_or_cell.text_frame
    paragraphs = list(tf.paragraphs)
    first = paragraphs[0]

    p_pr = copy.deepcopy(first._p.pPr) if first._p.pPr is not None else None
    end_pr = copy.deepcopy(first._p.endParaRPr) if first._p.endParaRPr is not None else None
    first_run = first.runs[0] if first.runs else None
    r_pr = copy.deepcopy(first_run._r.rPr) if first_run is not None and first_run._r.rPr is not None else None
    fallback = _font_style(shape_or_cell)

    for paragraph in paragraphs[1:]:
        parent = paragraph._p.getparent()
        if parent is not None:
            parent.remove(paragraph._p)

    for child in list(first._p):
        first._p.remove(child)
    if p_pr is not None:
        first._p.append(p_pr)

    last_paragraph = first
    for index, line in enumerate(lines):
        if index > 0:
            last_paragraph = tf.add_paragraph()
            if p_pr is not None:
                last_paragraph._p.insert(0, copy.deepcopy(p_pr))
        run = last_paragraph.add_run()
        run.text = line
        if r_pr is not None:
            run._r.insert(0, copy.deepcopy(r_pr))
        else:
            _apply_run_style(run, fallback)
        if index == 0 and end_pr is not None:
            last_paragraph._p.append(end_pr)


def apply_text_style(shape_or_cell, payload: Dict[str, Any]) -> None:
    """Apply explicit character/paragraph style to all runs in a text frame."""
    tf = getattr(shape_or_cell, "text_frame", None)
    if tf is None:
        return
    alignment = str(payload.get("alignment") or "").upper()
    align_value = None
    if alignment:
        if alignment in PP_ALIGN.__members__:
            align_value = PP_ALIGN[alignment]
        elif alignment == "CENTER":
            align_value = PP_ALIGN.CENTER
    for paragraph in tf.paragraphs:
        if align_value is not None:
            try:
                paragraph.alignment = align_value
            except Exception:
                pass
        runs = list(paragraph.runs)
        if not runs:
            runs = []
        for run in runs:
            try:
                font = run.font
                if payload.get("font_name"):
                    font.name = str(payload["font_name"])
                if payload.get("font_size_pt") is not None:
                    font.size = Pt(float(payload["font_size_pt"]))
                if payload.get("bold") is not None:
                    font.bold = bool(payload["bold"])
                if payload.get("italic") is not None:
                    font.italic = bool(payload["italic"])
                if payload.get("color"):
                    font.color.rgb = RGBColor.from_string(str(payload["color"]).lstrip("#").upper())
            except Exception:
                continue


def _clamp_bbox(bbox: Dict[str, Any], slide_w: int, slide_h: int) -> Dict[str, int]:
    def value(*names: str) -> Optional[int]:
        for name in names:
            if bbox.get(name) is not None:
                try:
                    return int(round(float(bbox[name])))
                except (TypeError, ValueError):
                    return None
        return None

    x = value("x", "left") or 0
    y = value("y", "top") or 0
    w = value("w", "width") or int(slide_w * 0.7)
    h = value("h", "height") or int(slide_h * 0.1)
    w = max(1, w)
    h = max(1, h)
    x = max(0, min(x, slide_w - 1))
    y = max(0, min(y, slide_h - 1))
    w = max(1, min(w, slide_w - x))
    h = max(1, min(h, slide_h - y))
    return {"x": x, "y": y, "w": w, "h": h}


def _slide_body_style(slide) -> Dict[str, Any]:
    best = None
    best_len = -1
    for shape in slide.shapes:
        try:
            if not getattr(shape, "has_text_frame", False) or not shape.text.strip():
                continue
            length = len(shape.text)
            if length > best_len:
                best = shape
                best_len = length
        except Exception:
            continue
    return _font_style(best) if best is not None else {}


def _style_from_payload(payload: Dict[str, Any], slide) -> Dict[str, Any]:
    style = _slide_body_style(slide)
    if payload.get("font_name"):
        style["name"] = str(payload["font_name"])
    if payload.get("font_size_pt") is not None:
        try:
            style["size"] = Pt(float(payload["font_size_pt"]))
        except (TypeError, ValueError):
            pass
    if payload.get("bold") is not None:
        style["bold"] = bool(payload["bold"])
    if payload.get("italic") is not None:
        style["italic"] = bool(payload["italic"])
    if payload.get("color"):
        try:
            style["rgb"] = RGBColor.from_string(str(payload["color"]).lstrip("#").upper())
        except Exception:
            pass
    return style


def _add_textbox(slide, slide_w: int, slide_h: int, patch: Dict[str, Any]) -> Any:
    payload = patch.get("payload") or {}
    merged = {**patch, **payload}
    bbox = _clamp_bbox(merged.get("bbox_emu") or {}, slide_w, slide_h)
    text = patch.get("after")
    if text is None:
        text = payload.get("text", "")
    shape = slide.shapes.add_textbox(
        Emu(bbox["x"]), Emu(bbox["y"]), Emu(bbox["w"]), Emu(bbox["h"])
    )
    tf = shape.text_frame
    tf.word_wrap = True
    try:
        tf.auto_size = None
    except Exception:
        pass
    tf.text = ""
    first = tf.paragraphs[0]
    alignment = str(merged.get("alignment") or "").upper()
    try:
        if alignment in PP_ALIGN.__members__:
            first.alignment = PP_ALIGN[alignment]
    except Exception:
        pass
    style = _style_from_payload(merged, slide)
    if style.get("size") is None:
        style["size"] = Pt(18)
    last = first
    for index, line in enumerate(str(text or "").split("\n")):
        if index > 0:
            last = tf.add_paragraph()
        run = last.add_run()
        run.text = line
        _apply_run_style(run, style)
    return shape


def _remove_shape(shape) -> None:
    element = shape.element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _validate_before(patch: Dict[str, Any], current_text: str, shape_id: str) -> None:
    before = patch.get("before")
    if before is None or not str(before).strip():
        return
    if normalize_text(before) == normalize_text(current_text):
        return
    ratio = difflib.SequenceMatcher(None, normalize_text(before), normalize_text(current_text)).ratio()
    if ratio < 0.9:
        raise PatchConflict(
            f"shape {shape_id} changed on the server; refresh and retry"
        )


def _apply_geometry(shape, bbox: Dict[str, Any], slide_w: int, slide_h: int) -> None:
    clamped = _clamp_bbox(bbox, slide_w, slide_h)
    if any(bbox.get(k) is not None for k in ("x", "left")):
        shape.left = Emu(clamped["x"])
    if any(bbox.get(k) is not None for k in ("y", "top")):
        shape.top = Emu(clamped["y"])
    if any(bbox.get(k) is not None for k in ("w", "width")):
        shape.width = Emu(clamped["w"])
    if any(bbox.get(k) is not None for k in ("h", "height")):
        shape.height = Emu(clamped["h"])


def _apply_one(prs, patch: Dict[str, Any], seen: Set[str]) -> int:
    action = str(patch.get("action") or "edit").strip().lower()
    if action in ("update", "rewrite", "replace", "edit_text", "set_text"):
        action = "edit"
    if action in ("move", "resize", "layout", "set_geometry"):
        action = "geometry"
    if action in ("format", "format_text"):
        action = "style"
    if action in ("clear_text",):
        action = "clear"
    if action in ("delete_cell_text",):
        action = "clear"

    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)

    if action == "add_textbox":
        try:
            slide_index = int(patch.get("slide_index"))
        except (TypeError, ValueError):
            raise PatchError("slide_index is required for add_textbox")
        if slide_index < 0 or slide_index >= len(prs.slides):
            raise PatchError(f"slide_index out of range: {slide_index}")
        _add_textbox(prs.slides[slide_index], slide_w, slide_h, patch)
        return slide_index

    shape_id = str(patch.get("shape_id") or "")
    cell_ref = parse_cell_id(shape_id)
    parsed = parse_shape_id(shape_id)
    if cell_ref:
        parsed = parse_shape_id(cell_ref[0])
    if not parsed:
        raise PatchError(f"invalid shape_id: {shape_id!r}")
    slide_index = parsed[0]
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise PatchError(f"slide_index out of range: {slide_index}")
    text_action = action in ("edit", "clear", "delete_textbox")
    if text_action and shape_id in seen:
        raise PatchError(f"duplicate patch for {shape_id}")
    if text_action:
        seen.add(shape_id)

    slide = prs.slides[slide_index]
    payload = patch.get("payload") or {}
    merged = {**patch, **payload}
    style_payload = merged.get("style") if isinstance(merged.get("style"), dict) else merged

    if cell_ref:
        cell = find_table_cell(slide, shape_id)
        if cell is None:
            raise PatchConflict(f"table cell not found: {shape_id}")
        current = cell.text or ""
        if action == "clear":
            _validate_before(patch, current, shape_id)
            set_text_preserving_style(cell, "")
            return slide_index
        if action not in ("edit", "style"):
            raise PatchError(f"unsupported action for table cell: {action}")
        _validate_before(patch, current, shape_id)
        text = patch.get("after")
        if text is None:
            text = payload.get("text")
        if text is not None:
            set_text_preserving_style(cell, text)
        apply_text_style(cell, style_payload)
        return slide_index

    shape = find_shape(slide, shape_id)
    if shape is None:
        raise PatchConflict(f"shape not found: {shape_id}")
    grouped = bool(getattr(shape, "shape_type", None) == 6)
    # A nested shape is found recursively; moving nested shapes in group coordinate
    # space is unsafe, so only text/style actions are allowed for groups.
    is_grouped = False
    try:
        parent = shape._element.getparent()
        is_grouped = parent is not None and parent.tag.endswith("grpSp")
    except Exception:
        pass
    if action == "clear":
        if not getattr(shape, "has_text_frame", False):
            raise PatchError(f"shape has no text frame: {shape_id}")
        _validate_before(patch, shape.text or "", shape_id)
        if is_grouped:
            raise PatchError("cannot clear text inside a group")
        set_text_preserving_style(shape, "")
        return slide_index
    if action == "delete_textbox":
        if not getattr(shape, "has_text_frame", False):
            raise PatchError(f"shape has no text frame: {shape_id}")
        if is_grouped:
            raise PatchError("cannot delete a shape inside a group")
        _validate_before(patch, shape.text or "", shape_id)
        _remove_shape(shape)
        return slide_index
    if action not in ("edit", "style", "geometry"):
        raise PatchError(f"unsupported action: {action}")

    has_text = getattr(shape, "has_text_frame", False)
    if action in ("edit", "style") and not has_text:
        raise PatchError(f"shape has no text frame: {shape_id}")
    if has_text:
        _validate_before(patch, shape.text or "", shape_id)
    if action == "geometry" and is_grouped:
        raise PatchError("cannot move or resize a shape inside a group")
    if not has_text and action != "geometry":
        raise PatchError(f"shape has no text frame: {shape_id}")

    if action in ("edit", "style"):
        text = patch.get("after")
        if text is None:
            text = payload.get("text")
        if text is not None:
            set_text_preserving_style(shape, text)
        apply_text_style(shape, style_payload)
    bbox = merged.get("bbox_emu")
    if isinstance(bbox, dict) and bbox:
        _apply_geometry(shape, bbox, slide_w, slide_h)
    return slide_index


def apply_patches_to_file(
    src: Path, dst: Path, patches: List[Dict[str, Any]]
) -> Set[int]:
    if not patches:
        raise PatchError("no patches supplied")
    if len(patches) > 100:
        raise PatchError("too many patches (max 100)")
    prs = Presentation(str(src))
    touched: Set[int] = set()
    seen: Set[str] = set()
    for patch in patches:
        if not isinstance(patch, dict):
            raise PatchError("patch must be an object")
        touched.add(_apply_one(prs, patch, seen))
    dst.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dst))
    return touched
