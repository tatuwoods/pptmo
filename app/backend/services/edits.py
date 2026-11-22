from __future__ import annotations

import copy
import difflib
import re
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
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


def _set_paragraph_text_preserving_style(paragraph, text: Any) -> None:
    """Rewrite one paragraph's text on its first run's formatting.

    Paragraph-scoped sibling of set_text_preserving_style, so the other
    paragraphs of the text frame keep their own formatting. Same limitation:
    mixed formatting inside the paragraph collapses to the first run.
    """
    value = str(text if text is not None else "").replace("\r\n", "\n").replace("\v", "\n")
    lines = value.split("\n")
    p = paragraph._p
    p_pr = copy.deepcopy(p.pPr) if p.pPr is not None else None
    end_pr = copy.deepcopy(p.endParaRPr) if p.endParaRPr is not None else None
    first_run = paragraph.runs[0] if paragraph.runs else None
    r_pr = copy.deepcopy(first_run._r.rPr) if first_run is not None and first_run._r.rPr is not None else None

    for child in list(p):
        p.remove(child)
    if p_pr is not None:
        p.append(p_pr)
    for index, line in enumerate(lines):
        if index > 0:
            paragraph.add_line_break()
        run = paragraph.add_run()
        run.text = line
        if r_pr is not None:
            run._r.insert(0, copy.deepcopy(r_pr))
    if end_pr is not None:
        p.append(end_pr)


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

    x = value("x", "left")
    y = value("y", "top")
    w = value("w", "width")
    h = value("h", "height")
    # defaults only for ABSENT values; an explicit 0/negative w/h used to be
    # silently replaced by 70%/10% of the slide instead of being rejected
    if w is not None and w <= 0:
        raise PatchError("bbox width must be positive")
    if h is not None and h <= 0:
        raise PatchError("bbox height must be positive")
    x = 0 if x is None else x
    y = 0 if y is None else y
    w = int(slide_w * 0.7) if w is None else w
    h = int(slide_h * 0.1) if h is None else h
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


MAX_PATCH_TEXT_CHARS = 20000
MAX_PATCH_IMAGE_B64_CHARS = 41 * 1024 * 1024  # ≈30 MB binary after base64
FUZZY_BEFORE_MAX_CHARS = 10000


def _validate_patch_sizes(patch: Dict[str, Any]) -> None:
    """Direct /apply uploads bypass the agent's own caps; bound the strings so a
    hostile patch list can't force quadratic CPU or huge text into the PPTX."""
    def _check(value: Any, label: str) -> None:
        if not isinstance(value, str):
            return
        if label.endswith("image_base64"):
            if len(value) > MAX_PATCH_IMAGE_B64_CHARS:
                raise PatchError("payload.image_base64 too large (max ~30 MB)")
            return
        if len(value) > MAX_PATCH_TEXT_CHARS:
            raise PatchError(f"{label} too long (max {MAX_PATCH_TEXT_CHARS} chars)")

    for key in ("before", "after", "text", "find", "replace"):
        _check(patch.get(key), key)
    payload = patch.get("payload")
    if isinstance(payload, dict):
        for key, value in payload.items():
            _check(value, f"payload.{key}")


def _validate_before(patch: Dict[str, Any], current_text: str, shape_id: str) -> None:
    before = patch.get("before")
    if before is None or not str(before).strip():
        return
    if normalize_text(before) == normalize_text(current_text):
        return
    # SequenceMatcher is quadratic; only worth it for realistically-sized texts
    if len(before) <= FUZZY_BEFORE_MAX_CHARS:
        ratio = difflib.SequenceMatcher(None, normalize_text(before), normalize_text(current_text)).ratio()
        if ratio < 0.9:
            raise PatchConflict(
                f"shape {shape_id} changed on the server; refresh and retry"
                f"（对象 {shape_id} 的原文已变化，请刷新后重试）"
            )
    else:
        raise PatchConflict(
            f"shape {shape_id} changed on the server; refresh and retry"
            f"（对象 {shape_id} 的原文已变化，请刷新后重试）"
        )


def _own_xfrm(element):
    """The shape's own `a:xfrm`; None when its box is inherited from the layout."""
    for container_tag in ("p:spPr", "p:grpSpPr"):
        container = element.find(qn(container_tag))
        if container is not None:
            xfrm = container.find(qn("a:xfrm"))
            if xfrm is not None:
                return xfrm
    return element.find(qn("p:xfrm"))  # graphicFrame


def _pin_inherited_box(shape) -> None:
    """Write a shape's inherited box as an explicit a:xfrm before a coordinate is set.

    Placeholders usually carry no a:xfrm and inherit position and size from
    their layout. python-pptx resolves those for reads, but assigning e.g.
    `.left` creates an a:xfrm holding just that value, so the shape snaps to
    y=0 (or x=0). Writing all four resolved values first keeps the axes that
    are not being edited where they render.
    """
    xfrm = _own_xfrm(shape._element)
    if xfrm is not None and xfrm.find(qn("a:off")) is not None and xfrm.find(qn("a:ext")) is not None:
        return
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if None in (left, top, width, height):
        return
    box = (Emu(int(left)), Emu(int(top)), Emu(int(width)), Emu(int(height)))
    shape.left, shape.top, shape.width, shape.height = box


def _apply_geometry(shape, bbox: Dict[str, Any], slide_w: int, slide_h: int) -> None:
    # a placeholder inherits its box; pin it before the first coordinate write
    _pin_inherited_box(shape)
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
    action = _canonical_action(patch)

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

    if action == "slide_background":
        payload = patch.get("payload") or {}
        merged = {**patch, **payload}
        try:
            slide_index = int(merged.get("slide_index"))
        except (TypeError, ValueError):
            raise PatchError("slide_background requires slide_index")
        if slide_index < 0 or slide_index >= len(prs.slides):
            raise PatchError(f"slide_index out of range: {slide_index}")
        _apply_slide_background(prs.slides[slide_index], merged)
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
    payload_slide = patch.get("slide_index")
    try:
        payload_num = int(payload_slide) if payload_slide is not None else None
    except (TypeError, ValueError):
        payload_num = None
    if payload_num is not None and payload_num != slide_index:
        # PatchError subclasses ValueError — the comparison must live outside
        # the int-parse try/except or it would swallow itself
        raise PatchError(
            f"slide_index {payload_slide} does not match shape_id target slide {slide_index}"
        )
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
    if action == "replace_image":
        _apply_replace_image(slide, shape, merged)
        return slide_index
    if action == "table_mod":
        if not getattr(shape, "has_table", False):
            raise PatchError(f"shape is not a table: {shape_id}")
        _apply_table_mod(shape.table, merged)
        return slide_index
    if action == "duplicate_object":
        _apply_duplicate_object(slide, shape, merged.get("offset_emu"), slide_w, slide_h)
        return slide_index
    if action == "z_order":
        _apply_z_order(slide, shape, merged.get("mode"))
        return slide_index
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


# ---------- slide-level operations (v0.2) ----------
# python-pptx has no duplicate/delete/move API, so these work directly on the
# sldIdLst element and the OPC relationship graph.

SLIDE_ACTIONS = {"duplicate_slide", "add_slide", "delete_slide", "move_slide"}

# deck-wide bulk ops (v0.3): run after shape patches but BEFORE slide-level ops
# so slide_index parameters still refer to the base deck the caller read
DECK_ACTIONS = {"font_unify", "replace_text", "font_scale", "add_page_numbers"}

_ACTION_ALIASES = {
    "update": "edit",
    "rewrite": "edit",
    "replace": "edit",
    "edit_text": "edit",
    "set_text": "edit",
    "move": "geometry",
    "resize": "geometry",
    "layout": "geometry",
    "set_geometry": "geometry",
    "format": "style",
    "format_text": "style",
    "clear_text": "clear",
    "delete_cell_text": "clear",
    "duplicate": "duplicate_slide",
    "copy_slide": "duplicate_slide",
    "clone_slide": "duplicate_slide",
    "new_slide": "add_slide",
    "insert_slide": "add_slide",
    "remove_slide": "delete_slide",
    "reorder_slide": "move_slide",
    "unify_font": "font_unify",
    "font_unify": "font_unify",
    "unify_fonts": "font_unify",
    "replace_text": "replace_text",
    "global_replace": "replace_text",
    "replace_all": "replace_text",
    "font_scale": "font_scale",
    "scale_font": "font_scale",
    "page_numbers": "add_page_numbers",
    "add_page_number": "add_page_numbers",
    "add_page_numbers": "add_page_numbers",
    "align_objects": "align_objects",
    "align_shapes": "align_objects",
    "distribute": "align_objects",
    "distribute_objects": "align_objects",
    "duplicate_object": "duplicate_object",
    "copy_object": "duplicate_object",
    "clone_object": "duplicate_object",
    "z_order": "z_order",
    "zorder": "z_order",
    "reorder_z": "z_order",
    "layer": "z_order",
    "table_mod": "table_mod",
    "modify_table": "table_mod",
    "table_row": "table_mod",
    "table_col": "table_mod",
    "replace_image": "replace_image",
    "swap_image": "replace_image",
    "change_image": "replace_image",
    "slide_background": "slide_background",
    "background": "slide_background",
    "set_background": "slide_background",
}


def _canonical_action(patch: Dict[str, Any]) -> str:
    action = str(patch.get("action") or patch.get("type") or "edit").strip().lower()
    return _ACTION_ALIASES.get(action, action)


_UNSET = object()


def _get_slide_param(patch: Dict[str, Any], names: Tuple[str, ...], label: str, default: Any = _UNSET) -> Optional[int]:
    for name in names:
        value = patch.get(name)
        if value is None:
            continue
        if isinstance(value, bool):
            raise PatchError(f"{label} must be an integer: {value!r}")
        try:
            return int(value)
        except (TypeError, ValueError):
            raise PatchError(f"{label} must be an integer: {value!r}")
    if default is not _UNSET:
        return default
    raise PatchError(f"{label} is required")


def _remap_rids(element, remap: Dict[str, str]) -> None:
    """Rewrite rId references inside copied slide XML (r:embed / r:id / r:link ...).
    Two-phase via unique tokens so remaps like rId2→rId3 + rId3→rId2 can't collide."""
    tokens = {old: f"@@rmap{index}@@" for index, old in enumerate(remap)}
    for node in element.iter():
        for attr, value in list(node.attrib.items()):
            if value in tokens:
                node.set(attr, tokens[value])
    reverse = {token: remap[old] for old, token in tokens.items()}
    for node in element.iter():
        for attr, value in list(node.attrib.items()):
            if value in reverse:
                node.set(attr, reverse[value])


def _clone_slide(prs, source) -> Any:
    """Deep-copy a slide into a new part reusing its layout: shapes, background,
    relationships (images/charts/... re-bound to the same target parts) and notes."""
    dest = prs.slides.add_slide(source.slide_layout)
    _dedupe_slide_partname(dest)
    # add_slide clones layout placeholders; the source's own shapes replace them
    for shape in list(dest.shapes):
        _remove_shape(shape)
    dst_tree = dest.shapes._spTree
    for child in source.shapes._spTree:
        if child.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        copied = copy.deepcopy(child)
        _renew_field_ids(copied)
        dst_tree.append(copied)
    src_cSld = source._element.cSld
    dst_cSld = dest._element.cSld
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is not None:
        dst_bg = dst_cSld.find(qn("p:bg"))
        if dst_bg is not None:
            dst_cSld.remove(dst_bg)
        dst_cSld.insert(0, copy.deepcopy(src_bg))
    # transition / timing live on the sld root after clrMapOvr
    for tag in ("p:transition", "p:timing"):
        src_el = source._element.find(qn(tag))
        if src_el is not None:
            dest._element.append(copy.deepcopy(src_el))
    remap: Dict[str, str] = {}
    for rId, rel in list(source.part.rels.items()):
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue  # layout already bound by add_slide; notes are 1:1, copied below
        if rel.is_external:
            new_rId = dest.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rId = dest.part.relate_to(rel.target_part, rel.reltype)
        remap[rId] = new_rId
    if remap:
        _remap_rids(dest._element, remap)
    if source.has_notes_slide:
        try:
            dest.notes_slide.notes_text_frame.text = source.notes_slide.notes_text_frame.text
        except Exception:
            pass
    return dest


def _sld_ids(prs) -> List[Any]:
    return list(prs.slides._sldIdLst)


def _dedupe_slide_partname(slide) -> None:
    """python-pptx derives new slide partnames from len(sldIdLst)+1, so after a
    delete (or in decks with non-contiguous numbering) the next add/copy can
    collide with a live part; the package then saves with duplicate zip entries
    and PowerPoint reads the wrong content for one of the two slides. Renaming
    the fresh part to a free slot is safe: rels reference parts as objects and
    [Content_Types].xml is rebuilt from partnames at save time."""
    part = slide.part
    others = {
        str(p.partname) for p in part.package.iter_parts() if p is not part
    }
    name = str(part.partname)
    if name not in others:
        return
    index = 1
    while f"/ppt/slides/slide{index}.xml" in others:
        index += 1
    part._partname = PackURI(f"/ppt/slides/slide{index}.xml")


def _move_sld_id(prs, from_index: int, to_index: int) -> None:
    sldIdLst = prs.slides._sldIdLst
    element = _sld_ids(prs)[from_index]
    sldIdLst.remove(element)
    sldIdLst.insert(to_index, element)


def _require_range(value: int, low: int, high: int, label: str) -> None:
    if value < low or value > high:
        raise PatchError(f"{label} out of range: {value} (valid {low}..{high})")


def _add_blank_slide_at(prs, layout_index: Optional[int], new_index: int) -> None:
    layouts = prs.slide_layouts
    if layout_index is None:
        # 6 is "Blank" in the standard template; fall back to the first layout
        layout_index = 6 if len(layouts) > 6 else 0
    if layout_index < 0 or layout_index >= len(layouts):
        raise PatchError(
            f"layout_index out of range: {layout_index} (deck has {len(layouts)} layouts)"
        )
    prs.slides.add_slide(layouts[layout_index])
    _dedupe_slide_partname(prs.slides[len(prs.slides) - 1])
    last = len(prs.slides) - 1  # appended at the end first
    if new_index != last:
        _move_sld_id(prs, last, new_index)


def _delete_slide_at(prs, index: int) -> None:
    if len(prs.slides) <= 1:
        raise PatchError("cannot delete the only slide in the deck")
    element = _sld_ids(prs)[index]
    prs.part.drop_rel(element.get(qn("r:id")))
    prs.slides._sldIdLst.remove(element)


def _apply_slide_patches(
    prs, patches: List[Dict[str, Any]], n_orig: int
) -> Set[int]:
    """Apply slide-level patches in submission order.

    Every slide number in the batch refers to the BASE deck — the state the
    caller read, matching objects.json and the shape patches above. An anchor
    list tracks where each original slide currently lives: entries are current
    deck order, values are original coordinates, and inserted/moved slides take
    ``position - 0.5`` so later ops still address base positions unambiguously.
    """
    anchors: List[float] = [float(i) for i in range(n_orig)]
    touched: Set[int] = set()

    def current_index(orig: int, label: str) -> int:
        try:
            return anchors.index(float(orig))
        except ValueError:
            raise PatchError(
                f"{label} {orig} already deleted earlier in this batch"
                f"（该页已在本批操作中删除）"
            ) from None

    def insertion_index(position: int) -> int:
        return sum(1 for anchor in anchors if anchor < float(position))

    for patch in patches:
        action = _canonical_action(patch)
        if action == "duplicate_slide":
            source = _get_slide_param(
                patch, ("source_slide", "source", "slide_index", "slide"), "source_slide"
            )
            position = _get_slide_param(
                patch, ("position", "to", "insert_at"), "position", default=source + 1
            )
            _require_range(source, 0, n_orig - 1, "source_slide")
            _require_range(position, 0, n_orig, "position")
            source_index = current_index(source, "source_slide")
            _clone_slide(prs, prs.slides[source_index])
            new_index = insertion_index(position)
            _move_sld_id(prs, len(anchors), new_index)  # clone was appended at the end
            anchors.insert(new_index, position - 0.5)
            touched.update({source_index, new_index})
        elif action == "add_slide":
            after = _get_slide_param(patch, ("after_slide",), "after_slide", default=None)
            position = _get_slide_param(
                patch, ("position", "insert_at"), "position", default=None
            )
            layout_index = _get_slide_param(
                patch, ("layout_index", "layout"), "layout_index", default=None
            )
            if after is not None:
                _require_range(after, 0, n_orig - 1, "after_slide")
                if position is None:
                    position = after + 1
            if position is None:
                position = n_orig
            _require_range(position, 0, n_orig, "position")
            new_index = insertion_index(position)
            _add_blank_slide_at(prs, layout_index, new_index)
            anchors.insert(new_index, position - 0.5)
            touched.add(new_index)
        elif action == "delete_slide":
            slide = _get_slide_param(patch, ("slide", "slide_index", "source_slide"), "slide")
            _require_range(slide, 0, n_orig - 1, "slide")
            if len(anchors) <= 1:
                raise PatchError("cannot delete the only slide（不能删除唯一一页）")
            index = current_index(slide, "slide")
            _delete_slide_at(prs, index)
            anchors.pop(index)
            touched.add(min(index, len(anchors) - 1))
        elif action == "move_slide":
            slide = _get_slide_param(
                patch, ("slide", "slide_index", "from", "source_slide"), "slide"
            )
            to = _get_slide_param(patch, ("to", "to_index", "position"), "to")
            _require_range(slide, 0, n_orig - 1, "slide")
            _require_range(to, 0, n_orig - 1, "to")
            from_index = current_index(slide, "slide")
            _move_sld_id(prs, from_index, to)  # remove-then-insert → lands at `to`
            anchors.pop(from_index)
            anchors.insert(to, float(to) - 0.5)
            touched.update({from_index, to})
        else:
            raise PatchError(f"unsupported slide action: {action}")
    return touched


# ---------- multi-shape operations (v0.3) ----------

ALIGN_MODES = {
    "left", "right", "top", "bottom",
    "hcenter", "vcenter", "center_h", "center_v",
    "distribute_h", "distribute_v",
}
_ALIGN_ALIASES = {"center": "hcenter", "middle": "vcenter", "center_x": "hcenter", "center_y": "vcenter"}


_SHAPE_TAGS = ("}sp", "}pic", "}graphicFrame", "}grpSp", "}cxnSp")


def _is_shape_element(element) -> bool:
    return any(element.tag.endswith(suffix) for suffix in _SHAPE_TAGS)


def _max_shape_id(slide) -> int:
    """Largest cNvPr id anywhere in the slide tree (groups included)."""
    best = 1
    for node in slide.shapes._spTree.iter():
        if node.tag.endswith("}cNvPr"):
            try:
                best = max(best, int(node.get("id")))
            except (TypeError, ValueError):
                continue
    return best


def _apply_duplicate_object(slide, shape, offset_raw: Any, slide_w: int, slide_h: int) -> None:
    if shape._element.getparent() is None:
        raise PatchError("shape has no parent")
    offset = offset_raw if isinstance(offset_raw, dict) else {}
    try:
        dx = int(round(float(offset.get("x", 304800))))
        dy = int(round(float(offset.get("y", 304800))))
    except (TypeError, ValueError):
        raise PatchError("offset_emu.x/y must be numbers")
    if dx < 0 or dy < 0 or dx > slide_w or dy > slide_h:
        raise PatchError("offset_emu 超出画布范围")
    new_el = copy.deepcopy(shape._element)
    next_id = _max_shape_id(slide) + 1
    for node in new_el.iter():
        if node.tag.endswith("}cNvPr"):
            node.set("id", str(next_id))
            next_id += 1
    # shift the copy before insertion: works the same at top level and inside
    # groups (offsets live in the parent's coordinate space either way)
    new_left = max(0, int(shape.left or 0) + dx)
    new_top = max(0, int(shape.top or 0) + dy)
    off = next(new_el.iter(qn("a:off")), None)
    if off is not None:
        off.set("x", str(new_left))
        off.set("y", str(new_top))
    else:
        # the source inherits its box (placeholder without a:xfrm): a copy
        # without one of its own stacks exactly under the original
        width, height = shape.width, shape.height
        sp_pr = new_el.find(qn("p:spPr"))
        if None in (width, height) or sp_pr is None:
            raise PatchError("shape has no box of its own to copy")
        xfrm = parse_xml(
            f'<a:xfrm {nsdecls("a")}><a:off x="{new_left}" y="{new_top}"/>'
            f'<a:ext cx="{int(width)}" cy="{int(height)}"/></a:xfrm>'
        )
        sp_pr.insert(0, xfrm)
    _renew_field_ids(new_el)
    shape._element.addnext(new_el)
    # same slide part → copied rId references stay valid; no rel rebinding needed


_Z_ORDER_MODES = {"front", "back", "forward", "backward"}
_Z_MODE_ALIASES = {
    "bring_to_front": "front",
    "bring_front": "front",
    "top": "front",
    "置顶": "front",
    "send_to_back": "back",
    "send_back": "back",
    "bottom": "back",
    "置底": "back",
    "up": "forward",
    "上移": "forward",
    "down": "backward",
    "下移": "backward",
}


def _apply_z_order(slide, shape, mode_raw: Any) -> None:  # noqa: ARG001 (slide kept for symmetry)
    mode = str(mode_raw or "").strip().lower()
    mode = _Z_MODE_ALIASES.get(mode, mode)
    if mode not in _Z_ORDER_MODES:
        raise PatchError(
            f"unsupported z_order mode: {mode_raw!r} (支持 front/back/forward/backward)"
        )
    element = shape._element
    parent = element.getparent()
    if parent is None:
        raise PatchError("shape has no parent")
    children = list(parent)
    if mode == "front":
        parent.remove(element)
        # p:extLst closes spTree / grpSp; a shape after it breaks the sequence
        ext_lst = next((c for c in parent if c.tag == qn("p:extLst")), None)
        if ext_lst is None:
            parent.append(element)
        else:
            ext_lst.addprevious(element)
    elif mode == "back":
        parent.remove(element)
        insert_at = 0
        for index, child in enumerate(children):
            if _is_shape_element(child):
                insert_at = index
                break
        else:
            insert_at = len(children)
        # insert where the first shape was (children captured before removal)
        parent.insert(insert_at, element)
    elif mode == "forward":
        idx = children.index(element)
        nxt = next((c for c in children[idx + 1:] if _is_shape_element(c)), None)
        if nxt is not None:
            parent.remove(element)
            nxt.addnext(element)
    else:  # backward
        idx = children.index(element)
        prev = next((c for c in reversed(children[:idx]) if _is_shape_element(c)), None)
        if prev is not None:
            parent.remove(element)
            prev.addprevious(element)


# ---------- table row/col/merge operations (v0.3) ----------

TABLE_MOD_OPS = {"add_row", "add_col", "delete_row", "delete_col", "merge_cells", "split_cells"}
TABLE_MAX_ROWS = 60
TABLE_MAX_COLS = 30


def _tc_attr(tc, name: str) -> int:
    try:
        return int(tc.get(name) or 1)
    except (TypeError, ValueError):
        return 1


def _strip_merge_flags(tr) -> None:
    """Make a cloned row a plain row: drop merge attrs so the copy can't
    dangle references to merges that don't span it."""
    for tc in tr.findall(qn("a:tc")):
        for attr in ("rowSpan", "gridSpan", "hMerge", "vMerge"):
            if tc.get(attr) is not None:
                del tc.attrib[attr]


def _clear_tr_text(tr) -> None:
    """Empty every cell of a cloned row, keeping pPr/tcPr (styling) intact."""
    for tc in tr.findall(qn("a:tc")):
        _clear_tc_text(tc)


def _clear_tc_text(tc) -> None:
    txBody = tc.find(qn("a:txBody"))
    if txBody is None:
        return
    paragraphs = txBody.findall(qn("a:p"))
    for paragraph in paragraphs[1:]:
        txBody.remove(paragraph)
    if paragraphs:
        for child in list(paragraphs[0]):
            if child.tag != qn("a:pPr"):
                paragraphs[0].remove(child)


def _merge_origins(tbl) -> List[Tuple[int, int, int, int]]:
    """(origin_row, origin_col, rowSpan, gridSpan) for every merged range."""
    origins = []
    for r, tr in enumerate(tbl.tr_lst):
        for c, tc in enumerate(tr.findall(qn("a:tc"))):
            row_span = _tc_attr(tc, "rowSpan")
            grid_span = _tc_attr(tc, "gridSpan")
            if row_span > 1 or grid_span > 1:
                origins.append((r, c, row_span, grid_span))
    return origins


def _require_index(value: Any, high: int, label: str) -> int:
    if value is None:
        raise PatchError(f"table_mod.{label} is required")
    try:
        index = int(value)
    except (TypeError, ValueError):
        raise PatchError(f"table_mod.{label} must be an integer: {value!r}")
    if index < 0 or index >= high:
        raise PatchError(f"table_mod.{label} out of range: {index} (0..{high - 1})")
    return index


def _sync_frame_ext(table) -> None:
    """Keep the graphicFrame size equal to the grid after add/delete row/col.

    python-pptx only touches tblGrid and tr@h, so the frame keeps its old ext.
    LibreOffice still lays the table out by the grid, but parse._table_cells
    scales column widths to the frame width — every cell hit-box and the
    bbox_cm handed to the agent end up misplaced."""
    tbl = table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    widths = [int(gc.get("w") or 0) for gc in grid.findall(qn("a:gridCol"))] if grid is not None else []
    if not widths:
        return
    heights = [int(tr.get("h") or 0) for tr in tbl.tr_lst]
    frame = tbl.getparent().getparent().getparent()  # a:tbl → a:graphicData → a:graphic → p:graphicFrame
    xfrm = frame.find(qn("p:xfrm")) if frame is not None else None
    ext = xfrm.find(qn("a:ext")) if xfrm is not None else None
    if ext is None:
        return
    ext.set("cx", str(sum(widths)))
    ext.set("cy", str(sum(heights)))


def _apply_table_mod(table, merged: Dict[str, Any]) -> None:
    op = str(merged.get("op") or merged.get("table_op") or "").strip().lower()
    if op not in TABLE_MOD_OPS:
        raise PatchError(f"unsupported table_mod op: {op!r} (支持 {'/'.join(sorted(TABLE_MOD_OPS))})")
    tbl = table._tbl
    rows = len(table.rows)
    cols = len(table.columns)

    if op == "add_row":
        if rows >= TABLE_MAX_ROWS:
            raise PatchError(f"表格行数已达上限（{TABLE_MAX_ROWS}）")
        tr_new = copy.deepcopy(tbl.tr_lst[-1])
        _strip_merge_flags(tr_new)
        _clear_tr_text(tr_new)
        tbl.append(tr_new)
        _sync_frame_ext(table)
        return
    if op == "add_col":
        if cols >= TABLE_MAX_COLS:
            raise PatchError(f"表格列数已达上限（{TABLE_MAX_COLS}）")
        grid = tbl.find(qn("a:tblGrid"))
        if grid is None:
            raise PatchError("table grid not found")
        widths = [int(gc.get("w") or 0) for gc in grid.findall(qn("a:gridCol"))]
        width = max(360000, int(sum(widths) / max(1, len(widths))) if widths else 720000)
        grid.append(grid.makeelement(qn("a:gridCol"), {"w": str(width)}))
        for tr in tbl.tr_lst:
            tcs = tr.findall(qn("a:tc"))
            template = tcs[-1] if tcs else None
            if template is None:
                raise PatchError("table row has no cells")
            tc_new = copy.deepcopy(template)
            # strip merge attrs on the CLONE only — the copy is a plain cell
            for attr in ("rowSpan", "gridSpan", "hMerge", "vMerge"):
                if tc_new.get(attr) is not None:
                    del tc_new.attrib[attr]
            _clear_tc_text(tc_new)
            tr.append(tc_new)
        _sync_frame_ext(table)
        return
    if op == "delete_row":
        index = _require_index(merged.get("index") if merged.get("index") is not None else merged.get("row"), rows, "index")
        for r0, c0, row_span, grid_span in _merge_origins(tbl):
            if r0 <= index < r0 + max(1, row_span):
                raise PatchError(
                    f"第 {index} 行与合并单元格（起点 R{r0 + 1}C{c0 + 1}）相交，请先拆分再删除"
                )
        tbl.remove(tbl.tr_lst[index])
        _sync_frame_ext(table)
        return
    if op == "delete_col":
        index = _require_index(merged.get("index") if merged.get("index") is not None else merged.get("col"), cols, "index")
        for r0, c0, row_span, grid_span in _merge_origins(tbl):
            if c0 <= index < c0 + max(1, grid_span):
                raise PatchError(
                    f"第 {index} 列与合并单元格（起点 R{r0 + 1}C{c0 + 1}）相交，请先拆分再删除"
                )
        grid = tbl.find(qn("a:tblGrid"))
        grid.remove(grid.findall(qn("a:gridCol"))[index])
        for tr in tbl.tr_lst:
            tcs = tr.findall(qn("a:tc"))
            tr.remove(tcs[index])
        _sync_frame_ext(table)
        return
    if op == "merge_cells":
        row = _require_index(merged.get("row"), rows, "row")
        col = _require_index(merged.get("col"), cols, "col")
        row2 = _require_index(merged.get("row2"), rows, "row2")
        col2 = _require_index(merged.get("col2"), cols, "col2")
        # either corner may come first; the region is the same and its origin
        # is the top-left cell. python-pptx moves every cell's text into the
        # origin as paragraphs, like PowerPoint's own merge — nothing is lost
        top, bottom = min(row, row2), max(row, row2)
        left, right = min(col, col2), max(col, col2)
        try:
            table.cell(top, left).merge(table.cell(bottom, right))
        except ValueError as exc:
            raise PatchError(f"merge_cells 失败：{exc}") from exc
        return
    # split_cells
    row = _require_index(merged.get("row"), rows, "row")
    col = _require_index(merged.get("col"), cols, "col")
    cell = table.cell(row, col)
    if not cell.is_merge_origin:
        raise PatchError(f"R{row + 1}C{col + 1} 不是合并单元格的起始格，无法拆分")
    try:
        cell.split()
    except ValueError as exc:
        raise PatchError(f"split_cells 失败：{exc}") from exc


def _apply_align_objects(prs, patch: Dict[str, Any]) -> int:
    payload = patch.get("payload") or {}
    merged = {**patch, **payload}
    raw_ids = merged.get("shape_ids")
    if not isinstance(raw_ids, list) or len(raw_ids) < 2:
        raise PatchError("align_objects 需要至少 2 个 shape_ids")
    if len(raw_ids) > 12:
        raise PatchError("align_objects 单次最多 12 个对象")
    align = str(merged.get("align") or merged.get("mode") or "").strip().lower()
    align = _ALIGN_ALIASES.get(align, align)
    if align not in ALIGN_MODES:
        raise PatchError(
            f"unsupported align mode: {align!r} (支持 {'/'.join(sorted(ALIGN_MODES))})"
        )
    parsed = [parse_shape_id(str(sid)) for sid in raw_ids]
    if any(p is None for p in parsed):
        raise PatchError("align_objects 的 shape_ids 含无效 ID")
    slides = {p[0] for p in parsed}
    if len(slides) != 1:
        raise PatchError("align_objects 只支持同一页内的对象")
    slide_index = slides.pop()
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise PatchError(f"slide_index out of range: {slide_index}")
    slide = prs.slides[slide_index]
    shapes = []
    for sid in raw_ids:
        shape = find_shape(slide, str(sid))
        if shape is None:
            raise PatchConflict(f"shape not found: {sid}")
        parent = shape._element.getparent()
        if parent is not None and parent.tag.endswith("grpSp"):
            raise PatchError(f"对象 {sid} 在分组内，不能参与对齐/分布")
        if shape not in shapes:
            shapes.append(shape)
    if len(shapes) < 2:
        raise PatchError("align_objects 需要至少 2 个不同对象")
    # a placeholder inherits its box; pin it before the first coordinate write
    for shape in shapes:
        _pin_inherited_box(shape)

    rects = [
        (int(s.left or 0), int(s.top or 0), int(s.width or 0), int(s.height or 0))
        for s in shapes
    ]
    if align == "left":
        target = min(r[0] for r in rects)
        for s in shapes:
            s.left = target
    elif align == "right":
        target = max(r[0] + r[2] for r in rects)
        for s in shapes:
            s.left = target - int(s.width or 0)
    elif align == "top":
        target = min(r[1] for r in rects)
        for s in shapes:
            s.top = target
    elif align == "bottom":
        target = max(r[1] + r[3] for r in rects)
        for s in shapes:
            s.top = target - int(s.height or 0)
    elif align in ("hcenter", "center_h"):
        target = round(sum(r[0] + r[2] / 2 for r in rects) / len(rects))
        for s in shapes:
            s.left = round(target - int(s.width or 0) / 2)
    elif align in ("vcenter", "center_v"):
        target = round(sum(r[1] + r[3] / 2 for r in rects) / len(rects))
        for s in shapes:
            s.top = round(target - int(s.height or 0) / 2)
    elif align in ("distribute_h", "distribute_v"):
        if len(shapes) < 3:
            raise PatchError("distribute_h/distribute_v 需要至少 3 个对象")
        # first and last stay fixed; middle shapes get equal gaps
        order = sorted(range(len(shapes)), key=lambda i: rects[i][0] if align == "distribute_h" else rects[i][1])
        if len(order) > 2:
            if align == "distribute_h":
                start = rects[order[0]][0]
                end = rects[order[-1]][0] + rects[order[-1]][2]
                total_w = sum(rects[i][2] for i in order)
                gap = (end - start - total_w) / (len(order) - 1)
                cursor = start
                for i in order:
                    shapes[i].left = round(cursor)
                    cursor += rects[i][2] + gap
            else:
                start = rects[order[0]][1]
                end = rects[order[-1]][1] + rects[order[-1]][3]
                total_h = sum(rects[i][3] for i in order)
                gap = (end - start - total_h) / (len(order) - 1)
                cursor = start
                for i in order:
                    shapes[i].top = round(cursor)
                    cursor += rects[i][3] + gap
    return slide_index





def _iter_text_targets(container):
    """Yield every text-bearing shape / table cell on a slide (groups included)."""
    def walk(shapes):
        for shape in shapes:
            try:
                if getattr(shape, "shape_type", None) == 6:
                    yield from walk(shape.shapes)
                    continue
            except Exception:
                pass
            if getattr(shape, "has_table", False):
                try:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            yield cell
                    continue
                except Exception:
                    continue
            if getattr(shape, "has_text_frame", False):
                yield shape

    try:
        yield from walk(container.shapes)
    except AttributeError:
        yield container  # a bare cell from a previous walk level


def _set_typefaces_on(rPr, name: str, create: bool = True) -> bool:
    """Set latin+ea+cs typefaces on an rPr-like element (rPr/defRPr/endParaRPr).

    With create=False only pre-existing latin declarations are rewritten —
    used for paragraph defaults, where inventing new defRPr nodes would be noise."""
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        if not create:
            return False
        latin = rPr.makeelement(qn("a:latin"), {})
        rPr.append(latin)
    latin.set("typeface", name)
    anchor = latin
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            if not create:
                continue
            el = rPr.makeelement(qn(tag), {})
            anchor.addnext(el)
        el.set("typeface", name)
        anchor = el
    return True


def _set_run_font_name(run, name: str) -> None:
    """Set latin + east-asian + complex-script typefaces (font.name only writes latin)."""
    run.font.name = name
    _set_typefaces_on(run._r.get_or_add_rPr(), name)


# Symbol/icon fonts must survive font_unify: their glyphs (check marks, arrows,
# PUA icons) render as garbage letters under a normal text face.
SYMBOL_FONT_NAMES = {
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "bookshelf symbol 7", "segoe mdl2 assets", "segoe fluent icons",
    "segoe ui symbol", "segoe ui emoji", "outlook", "marlett",
}


def _is_icon_text(text: str) -> bool:
    """True when every char lives in the Unicode private-use/icon planes."""
    stripped = (text or "").replace(" ", "").replace("\u00a0", "")
    if not stripped:
        return False
    for ch in stripped:
        code = ord(ch)
        if not (0xE000 <= code <= 0xF8FF or 0xF0000 <= code <= 0x1FFFFD):
            return False
    return True


def _run_uses_symbol_font(run) -> bool:
    rPr = run._r.find(qn("a:rPr"))
    latin = rPr.find(qn("a:latin")) if rPr is not None else None
    if latin is not None and str(latin.get("typeface") or "").strip().lower() in SYMBOL_FONT_NAMES:
        return True
    return _is_icon_text(run.text or "")


def _resolve_bulk_slides(prs, merged: Dict[str, Any], n_orig: int) -> List[Any]:
    """Resolve the target slide list for a bulk op against base numbering.

    An explicit scope=deck wins over slide_index (agents attach a card-locator
    slide_index to deck-wide ops — the anchor must not narrow the blast radius);
    scope must be "deck" or "slide"; the default for bulk ops is the whole deck."""
    scope = str(merged.get("scope") or "").strip().lower()
    slide_index = merged.get("slide_index")
    if scope in ("deck", "all", "document", "whole"):
        return list(prs.slides)
    if slide_index is not None:
        try:
            index = int(slide_index)
        except (TypeError, ValueError):
            raise PatchError(f"slide_index must be an integer: {slide_index!r}")
        if index < 0 or index >= n_orig:
            raise PatchError(f"slide_index out of range: {index} (deck has {n_orig})")
        return [prs.slides[index]]
    if scope in ("slide", "current", "page"):
        raise PatchError("scope=slide 需要 slide_index")
    if scope:
        raise PatchError(f"invalid scope: {scope}")
    return list(prs.slides)


def _unify_paragraph_defaults(text_frame, font_name: str) -> bool:
    """Rewrite fonts declared at paragraph/lstStyle level (defRPr / endParaRPr).

    Without this, text typed after a unified run still comes out in the old
    face — the defaults keep steering new runs."""
    changed = False
    tx_body = text_frame._txBody
    lst_style = tx_body.find(qn("a:lstStyle"))
    if lst_style is not None:
        for def_rpr in lst_style.iter(qn("a:defRPr")):
            changed |= _set_typefaces_on(def_rpr, font_name, create=False)
    for paragraph in text_frame.paragraphs:
        p = paragraph._p
        pPr = p.find(qn("a:pPr"))
        if pPr is not None:
            def_rpr = pPr.find(qn("a:defRPr"))
            if def_rpr is not None:
                changed |= _set_typefaces_on(def_rpr, font_name, create=False)
        end_para = p.find(qn("a:endParaRPr"))
        if end_para is not None:
            changed |= _set_typefaces_on(end_para, font_name, create=False)
    return changed


def _unify_master_tx_styles(prs, font_name: str) -> None:
    """Deck scope also covers master/layout text: shapes plus the master's
    titleStyle/bodyStyle/otherStyle declarations (footers, master titles)."""
    containers = list(prs.slide_masters)
    for master in prs.slide_masters:
        containers.extend(master.slide_layouts)
    for container in containers:
        for target in _iter_text_targets(container):
            for paragraph in target.text_frame.paragraphs:
                for run in paragraph.runs:
                    if not _run_uses_symbol_font(run):
                        _set_run_font_name(run, font_name)
            _unify_paragraph_defaults(target.text_frame, font_name)
    for master in prs.slide_masters:
        tx_styles = master._element.find(qn("p:txStyles"))
        if tx_styles is None:
            continue
        for style_tag in ("p:titleStyle", "p:bodyStyle", "p:otherStyle"):
            style = tx_styles.find(qn(style_tag))
            if style is None:
                continue
            for def_rpr in style.iter(qn("a:defRPr")):
                _set_typefaces_on(def_rpr, font_name, create=True)


def _op_font_unify(prs, merged: Dict[str, Any], n_orig: int) -> Set[int]:
    font_name = str(merged.get("font_name") or merged.get("font") or "").strip()
    if not font_name:
        raise PatchError("font_unify requires font_name")
    if len(font_name) > 100:
        raise PatchError("font_name too long (max 100 chars)")
    slides = _resolve_bulk_slides(prs, merged, n_orig)
    deck_wide = len(slides) >= n_orig
    touched: Set[int] = set()
    all_slides = list(prs.slides)
    for slide in slides:
        changed = False
        for target in _iter_text_targets(slide):
            for paragraph in target.text_frame.paragraphs:
                for run in paragraph.runs:
                    if _run_uses_symbol_font(run):
                        continue
                    _set_run_font_name(run, font_name)
                    changed = True
            changed |= _unify_paragraph_defaults(target.text_frame, font_name)
        if changed:
            touched.add(all_slides.index(slide))
    if deck_wide:
        _unify_master_tx_styles(prs, font_name)
    return touched


def _replace_in_paragraph(paragraph, find: str, replace: str) -> bool:
    """Replace every `find` in one paragraph; True when the paragraph changed.

    Occurrences that sit inside a single run are rewritten in that run, so the
    paragraph's other runs keep their own style. As soon as one occurrence
    spans run boundaries the whole paragraph is rebuilt from its original text
    on the first run's style — never from the already-rewritten text, which
    replaced twice whenever the replacement contained the needle (10 → 100
    came out as 1000).
    """
    text = paragraph.text
    if find not in text:
        return False
    run_hits = sum((run.text or "").count(find) for run in paragraph.runs)
    if text.count(find) > run_hits:
        _set_paragraph_text_preserving_style(paragraph, text.replace(find, replace))
        return True
    changed = False
    for run in paragraph.runs:
        if find in (run.text or ""):
            run.text = run.text.replace(find, replace)
            changed = True
    return changed


def _op_replace_text(prs, merged: Dict[str, Any], n_orig: int) -> Set[int]:
    find = str(merged.get("find") or "").strip()
    replace = str(merged.get("replace") if merged.get("replace") is not None else "")
    if not find:
        raise PatchError("replace_text requires a non-empty find")
    if len(find) > 500:
        raise PatchError("find too long (max 500 chars)")
    if len(replace) > 5000:
        raise PatchError("replace too long (max 5000 chars)")
    slides = _resolve_bulk_slides(prs, merged, n_orig)
    all_slides = list(prs.slides)
    touched: Set[int] = set()
    total_replaced = 0
    for slide in slides:
        changed = False
        for target in _iter_text_targets(slide):
            for paragraph in target.text_frame.paragraphs:
                if _replace_in_paragraph(paragraph, find, replace):
                    total_replaced += 1
                    changed = True
        if changed:
            touched.add(all_slides.index(slide))
    if total_replaced == 0:
        raise PatchError(f"未找到要替换的内容：「{find}」（请先 search_text 确认原文）")
    return touched


def _para_default_size_pt(paragraph) -> Optional[float]:
    """Paragraph-level default size (a:pPr/a:defRPr or a:endParaRPr @sz)."""
    pPr = paragraph._p.find(qn("a:pPr"))
    if pPr is not None:
        defRPr = pPr.find(qn("a:defRPr"))
        if defRPr is not None and defRPr.get("sz"):
            try:
                value = int(defRPr.get("sz")) / 100.0
                if value > 0:
                    return value
            except (TypeError, ValueError):
                pass
    end = paragraph._p.find(qn("a:endParaRPr"))
    if end is not None and end.get("sz"):
        try:
            value = int(end.get("sz")) / 100.0
            if value > 0:
                return value
        except (TypeError, ValueError):
            pass
    return None


def _sz_pt(rpr) -> Optional[float]:
    """`sz` of an rPr-like element (rPr/defRPr/endParaRPr) in points; None when absent."""
    if rpr is None or not rpr.get("sz"):
        return None
    try:
        value = int(rpr.get("sz")) / 100.0
    except (TypeError, ValueError):
        return None
    return value if value > 0 else None


def _lst_style_size_pt(tx_body, level: int) -> Optional[float]:
    """Size declared for a paragraph level (0-based) in a text body's a:lstStyle."""
    lst_style = tx_body.find(qn("a:lstStyle")) if tx_body is not None else None
    if lst_style is None:
        return None
    lvl = lst_style.find(qn(f"a:lvl{level + 1}pPr"))
    return _sz_pt(lvl.find(qn("a:defRPr"))) if lvl is not None else None


_TITLE_PH_TYPES = frozenset({PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE})
_BODY_PH_TYPES = frozenset({PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.SUBTITLE})


def _placeholder_size_pt(shape, level: int) -> Optional[float]:
    """Size a placeholder run inherits: the layout placeholder's list style, the
    master placeholder's, then the master text style for the placeholder type."""
    try:
        fmt = shape.placeholder_format
        layout = shape.part.slide_layout
        master = layout.slide_master
        layout_base = layout.placeholders.get(idx=fmt.idx)
    except (AttributeError, KeyError, ValueError):
        return None
    try:
        master_base = master.placeholders.get(fmt.type)  # keyed by type, not idx
    except Exception:
        master_base = None
    for base in (layout_base, master_base):
        if base is not None:
            size = _lst_style_size_pt(base._element.find(qn("p:txBody")), level)
            if size is not None:
                return size
    if fmt.type in _TITLE_PH_TYPES:
        style_tag = "p:titleStyle"
    elif fmt.type in _BODY_PH_TYPES:
        style_tag = "p:bodyStyle"
    else:
        style_tag = "p:otherStyle"
    tx_styles = master._element.find(qn("p:txStyles"))
    style = tx_styles.find(qn(style_tag)) if tx_styles is not None else None
    lvl = style.find(qn(f"a:lvl{level + 1}pPr")) if style is not None else None
    return _sz_pt(lvl.find(qn("a:defRPr"))) if lvl is not None else None


def _effective_size_pt(shape, paragraph, run) -> float:
    """Font size a run renders at, following where PowerPoint inherits it from:
    the run, the paragraph defaults, the shape's list style, then — for
    placeholders — `_placeholder_size_pt`. 18pt when nothing declares one."""
    if run.font.size is not None:
        return run.font.size.pt
    level = int(paragraph.level or 0)
    size = _para_default_size_pt(paragraph)
    if size is None:
        size = _lst_style_size_pt(shape.text_frame._txBody, level)
    if size is None and getattr(shape, "is_placeholder", False):
        size = _placeholder_size_pt(shape, level)
    return size if size is not None else 18.0


def _op_font_scale(prs, merged: Dict[str, Any], n_orig: int) -> Set[int]:
    raw_delta = merged.get("delta_pt", merged.get("delta", merged.get("by")))
    try:
        delta = float(raw_delta)
    except (TypeError, ValueError):
        raise PatchError("font_scale requires a numeric delta_pt")
    if not (-50 <= delta <= 50):
        raise PatchError("delta_pt must be within ±50 pt")
    slides = _resolve_bulk_slides(prs, merged, n_orig)
    all_slides = list(prs.slides)
    touched: Set[int] = set()
    for slide in slides:
        changed = False
        for target in _iter_text_targets(slide):
            for paragraph in target.text_frame.paragraphs:
                for run in paragraph.runs:
                    base = _effective_size_pt(target, paragraph, run)
                    run.font.size = Pt(max(6.0, min(400.0, base + delta)))
                    changed = True
        if changed:
            touched.add(all_slides.index(slide))
    return touched


PAGE_NUMBER_POSITIONS = {"bottom-left", "bottom-center", "bottom-right"}
# name stamped on the boxes add_page_numbers creates; the repeat guard keys on it
PAGE_NUMBER_NAME_PREFIX = "页码 "


def _field_id() -> str:
    return "{%s}" % str(uuid.uuid4()).upper()


def _renew_field_ids(element) -> None:
    """Give every a:fld in a copied element a fresh id — field ids are meant to
    be unique, and a copied slide or shape would otherwise share them."""
    for fld in element.iter(qn("a:fld")):
        fld.set("id", _field_id())


def _op_add_page_numbers(prs, merged: Dict[str, Any]) -> Set[int]:
    skip_first = bool(merged.get("skip_first") or merged.get("skip_cover") or False)
    position = str(merged.get("position") or "bottom-right").strip().lower()
    if position not in PAGE_NUMBER_POSITIONS:
        raise PatchError(f"position must be one of {'/'.join(sorted(PAGE_NUMBER_POSITIONS))}")
    presentation = prs.part._element
    explicit_start = merged.get("start")
    try:
        # without `start`, keep numbering from where the deck already does
        start = int(explicit_start) if explicit_start is not None else int(presentation.get("firstSlideNum") or 1)
    except (TypeError, ValueError):
        raise PatchError("start must be an integer")
    if not (0 <= start <= 9999):
        raise PatchError("start out of range (0..9999)")
    # A slidenum field is live but ignores the deck's "number slides from" in
    # LibreOffice — the renderer behind the preview and the PDF export — so a
    # start other than 1 can only be shown the same way everywhere as plain
    # digits (which then don't follow pages that are added, removed or moved).
    live_fields = start == 1
    try:
        size = float(merged.get("font_size_pt", 12))
    except (TypeError, ValueError):
        raise PatchError("font_size_pt must be numeric")
    if not (6 <= size <= 72):
        raise PatchError("font_size_pt out of range (6..72)")
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    box_w = min(1200000, slide_w // 8)
    box_h = min(420000, slide_h // 12)
    margin = min(300000, slide_w // 40)
    touched: Set[int] = set()
    for index, slide in enumerate(prs.slides):
        # idempotency guard: a second run over an already numbered deck is a
        # user error (duplicate numbers), not a silent no-op
        for shape in slide.shapes:
            # a slidenum field from a template numbers the deck just as well
            if any(fld.get("type") == "slidenum" for fld in shape._element.iter(qn("a:fld"))):
                raise PatchError(f"第 {index + 1} 页已有页码；如需重加请先删除旧页码")
            # keyed on the name this op stamps on its own boxes: matching on
            # "text equals the page number" also caught unrelated number tiles
            if str(getattr(shape, "name", "") or "").startswith(PAGE_NUMBER_NAME_PREFIX):
                raise PatchError(f"第 {index + 1} 页似乎已有页码；如需重加请先删除旧页码")
        if skip_first and index == 0:
            continue
        number = str(index + start)
        if position == "bottom-left":
            x, align = margin, PP_ALIGN.LEFT
        elif position == "bottom-center":
            x, align = (slide_w - box_w) // 2, PP_ALIGN.CENTER
        else:
            x, align = slide_w - box_w - margin, PP_ALIGN.RIGHT
        box = slide.shapes.add_textbox(Emu(x), Emu(slide_h - box_h - margin), Emu(box_w), Emu(box_h))
        box.name = f"{PAGE_NUMBER_NAME_PREFIX}{number}"
        tf = box.text_frame
        tf.word_wrap = False
        try:
            tf.auto_size = None
        except Exception:
            pass
        paragraph = tf.paragraphs[0]
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = number
        run.font.size = Pt(size)
        if live_fields:
            # a real slide-number field rather than the digit itself, so the
            # number follows its page when slides are added, removed or reordered
            fld = parse_xml(f'<a:fld {nsdecls("a")} id="{_field_id()}" type="slidenum"/>')
            for tag in ("a:rPr", "a:t"):
                child = run._r.find(qn(tag))
                if child is not None:
                    fld.append(child)
            paragraph._p.replace(run._r, fld)
        touched.add(index)
    return touched


_IMAGE_MAGIC_PNG = b"\x89PNG\r\n\x1a\n"


def _decode_image_payload(merged: Dict[str, Any]) -> bytes:
    b64 = merged.get("image_base64")
    path = merged.get("image_path")
    if b64:
        import base64
        import binascii

        if not isinstance(b64, str):
            raise PatchError("image_base64 must be a base64 string")
        try:
            data = base64.b64decode(b64, validate=True)
        except (binascii.Error, ValueError):
            raise PatchError("image_base64 解码失败：不是合法的 base64")
    elif path:
        source = Path(str(path))
        if not source.is_file():
            raise PatchError(f"image_path 不存在或不可读：{source.name}")
        if source.stat().st_size > 30 * 1024 * 1024:
            raise PatchError("图片超过 30 MB 限制")
        data = source.read_bytes()
    else:
        raise PatchError("replace_image 需要 image_base64 或 image_path")
    if len(data) > 30 * 1024 * 1024:
        raise PatchError("图片超过 30 MB 限制")
    if not (data.startswith(_IMAGE_MAGIC_PNG) or data.startswith(b"\xff\xd8\xff")):
        raise PatchError("只支持 png / jpg 图片（检测到其他格式）")
    return data


def _apply_replace_image(slide, shape, merged: Dict[str, Any]) -> None:
    import io

    if str(getattr(shape, "shape_type", "")).find("PICTURE") < 0:
        # shape_type for pictures renders as "PICTURE (13)"
        raise PatchError(f"shape 不是图片对象：{merged.get('shape_id')}")
    data = _decode_image_payload(merged)
    image_part, rId = slide.part.get_or_add_image_part(io.BytesIO(data))
    blip = shape._element.blipFill.blip
    blip.set(qn("r:embed"), rId)


_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def _bg_color(value: Any, label: str) -> str:
    text = str(value or "").strip().lstrip("#")
    if not _COLOR_RE.fullmatch(text):
        raise PatchError(f"{label} must be a 6-digit hex color (RRGGBB)")
    return text.upper()


def _apply_slide_background(slide, merged: Dict[str, Any]) -> None:
    gradient = merged.get("gradient")
    cSld = slide._element.cSld
    old_bg = cSld.find(qn("p:bg"))
    if old_bg is not None:
        cSld.remove(old_bg)
    if isinstance(gradient, dict):
        start = _bg_color(gradient.get("from", gradient.get("start")), "gradient.from")
        end = _bg_color(gradient.get("to", gradient.get("end")), "gradient.to")
        try:
            angle = float(gradient.get("angle", 90))
        except (TypeError, ValueError):
            angle = 90.0
        angle = max(0.0, min(359.9, angle))
        fill_xml = (
            f'<a:gradFill {nsdecls("a")} rotWithShape="1"><a:gsLst>'
            f'<a:gs pos="0"><a:srgbClr val="{start}"/></a:gs>'
            f'<a:gs pos="100000"><a:srgbClr val="{end}"/></a:gs>'
            f'</a:gsLst><a:lin ang="{int(angle * 60000)}" scaled="1"/></a:gradFill>'
        )
    else:
        color = _bg_color(merged.get("color"), "color")
        fill_xml = f'<a:solidFill {nsdecls("a")}><a:srgbClr val="{color}"/></a:solidFill>'
    bg = parse_xml(f'<p:bg {nsdecls("p", "a")}><p:bgPr>{fill_xml}<a:effectLst/></p:bgPr></p:bg>')
    cSld.insert(0, bg)


def _apply_deck_patches(
    prs, patches: List[Dict[str, Any]], n_orig: int
) -> Set[int]:
    touched: Set[int] = set()
    for patch in patches:
        action = _canonical_action(patch)
        payload = patch.get("payload") or {}
        merged = {**patch, **payload}
        if action == "font_unify":
            touched.update(_op_font_unify(prs, merged, n_orig))
        elif action == "replace_text":
            touched.update(_op_replace_text(prs, merged, n_orig))
        elif action == "font_scale":
            touched.update(_op_font_scale(prs, merged, n_orig))
        elif action == "add_page_numbers":
            touched.update(_op_add_page_numbers(prs, merged))
        else:
            raise PatchError(f"unsupported bulk action: {action}")
    return touched


def apply_patches_to_file(
    src: Path, dst: Path, patches: List[Dict[str, Any]]
) -> Set[int]:
    if not patches:
        raise PatchError("no patches supplied")
    if len(patches) > 100:
        raise PatchError("too many patches (max 100)")
    prs = Presentation(str(src))
    n_orig = len(prs.slides)
    touched: Set[int] = set()
    seen: Set[str] = set()
    # Mixed batches: shape-level ops resolve against the ORIGINAL numbering the
    # caller saw (matching objects.json), then slide-level ops mutate the deck
    # in submission order. Editing content on a just-inserted slide must wait
    # for the next round, after objects are re-parsed.
    shape_patches: List[Dict[str, Any]] = []
    slide_patches: List[Dict[str, Any]] = []
    deck_patches: List[Dict[str, Any]] = []
    for patch in patches:
        if not isinstance(patch, dict):
            raise PatchError("patch must be an object")
        _validate_patch_sizes(patch)
        action = _canonical_action(patch)
        if action in SLIDE_ACTIONS:
            slide_patches.append(patch)
        elif action in DECK_ACTIONS:
            deck_patches.append(patch)
        else:
            shape_patches.append(patch)
    for patch in shape_patches:
        if _canonical_action(patch) == "align_objects":
            touched.add(_apply_align_objects(prs, patch))
        else:
            touched.add(_apply_one(prs, patch, seen))
    if deck_patches:
        # deck ops still resolve against base numbering → must run before
        # slide-level ops mutate the deck
        touched.update(_apply_deck_patches(prs, deck_patches, n_orig))
    if slide_patches:
        touched.update(_apply_slide_patches(prs, slide_patches, n_orig))
    dst.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dst))
    return touched
