from __future__ import annotations

import copy
import difflib
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
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
FUZZY_BEFORE_MAX_CHARS = 10000


def _validate_patch_sizes(patch: Dict[str, Any]) -> None:
    """Direct /apply uploads bypass the agent's own caps; bound the strings so a
    hostile patch list can't force quadratic CPU or huge text into the PPTX."""
    def _check(value: Any, label: str) -> None:
        if isinstance(value, str) and len(value) > MAX_PATCH_TEXT_CHARS:
            raise PatchError(f"{label} too long (max {MAX_PATCH_TEXT_CHARS} chars)")

    for key in ("before", "after", "text"):
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
    # add_slide clones layout placeholders; the source's own shapes replace them
    for shape in list(dest.shapes):
        _remove_shape(shape)
    dst_tree = dest.shapes._spTree
    for child in source.shapes._spTree:
        if child.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        dst_tree.append(copy.deepcopy(child))
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
    for patch in patches:
        if not isinstance(patch, dict):
            raise PatchError("patch must be an object")
        _validate_patch_sizes(patch)
        if _canonical_action(patch) in SLIDE_ACTIONS:
            slide_patches.append(patch)
        else:
            shape_patches.append(patch)
    for patch in shape_patches:
        touched.add(_apply_one(prs, patch, seen))
    if slide_patches:
        touched.update(_apply_slide_patches(prs, slide_patches, n_orig))
    dst.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dst))
    return touched
