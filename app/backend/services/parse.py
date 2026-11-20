from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn


def _int(v: Any) -> int:
    try:
        return int(v) if v is not None else 0
    except (TypeError, ValueError):
        return 0


def _xml_font_size_pt(element) -> Optional[float]:
    for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
        try:
            for node in element.iter(qn(tag)):
                raw = node.get("sz")
                if raw:
                    value = float(int(raw)) / 100.0
                    if value > 0:
                        return value
        except Exception:
            continue
    return None


def _layout_props(slide, shape) -> Dict[str, Any]:
    """Read font size/alignment inherited from matching layout/master placeholders."""
    props: Dict[str, Any] = {}
    if not getattr(shape, "is_placeholder", False):
        return props
    try:
        idx = int(shape.placeholder_format.idx)
    except Exception:
        return props
    containers = []
    try:
        layout = slide.slide_layout
        containers.append(layout)
        containers.append(layout.slide_master)
    except Exception:
        pass
    for container in containers:
        try:
            placeholders = list(container.placeholders)
        except Exception:
            continue
        for placeholder in placeholders:
            try:
                if int(placeholder.placeholder_format.idx) != idx:
                    continue
                if "font_size_pt" not in props:
                    size = _xml_font_size_pt(placeholder._element)
                    if size:
                        props["font_size_pt"] = size
                if "alignment" not in props and placeholder.has_text_frame:
                    for para in placeholder.text_frame.paragraphs:
                        if para.alignment is not None:
                            props["alignment"] = (
                                para.alignment.name
                                if isinstance(para.alignment, PP_ALIGN)
                                else str(para.alignment)
                            )
                            break
                if "font_size_pt" in props and "alignment" in props:
                    return props
            except Exception:
                continue
    return props


def _font_info(shape, slide=None) -> Dict[str, Any]:
    info: Dict[str, Any] = {
        "font_name": None,
        "font_size_pt": None,
        "bold": None,
        "italic": None,
        "color": None,
        "alignment": None,
    }
    try:
        tf = shape.text_frame
        p = tf.paragraphs[0] if tf.paragraphs else None
        if p is None:
            return info
        if p.alignment is not None:
            info["alignment"] = p.alignment.name if isinstance(p.alignment, PP_ALIGN) else str(p.alignment)
        font = p.runs[0].font if p.runs else p.font
        info["font_name"] = font.name
        info["font_size_pt"] = float(font.size.pt) if font.size is not None else None
        info["bold"] = font.bold
        info["italic"] = font.italic
        try:
            if font.color is not None and font.color.rgb is not None:
                info["color"] = str(font.color.rgb)
        except Exception:
            pass
        if info["font_size_pt"] is None:
            info["font_size_pt"] = _xml_font_size_pt(shape._element)
    except Exception:
        pass
    if slide is not None and (info["font_size_pt"] is None or info["alignment"] is None):
        inherited = _layout_props(slide, shape)
        if info["font_size_pt"] is None:
            info["font_size_pt"] = inherited.get("font_size_pt")
        if info["alignment"] is None:
            info["alignment"] = inherited.get("alignment")
    return info


def _shape_id(slide_index: int, shape, fallback: int) -> str:
    xml_id = getattr(shape, "shape_id", None)
    if xml_id is not None:
        try:
            return f"s{slide_index}_id{int(xml_id)}"
        except (TypeError, ValueError):
            pass
    return f"s{slide_index}_sh{fallback}"


def _shape_kind(shape) -> str:
    try:
        if getattr(shape, "has_table", False):
            return "table"
        if getattr(shape, "has_chart", False):
            return "chart"
        if getattr(shape, "shape_type", None) == 6 or shape.__class__.__name__.lower().endswith("group"):
            return "group"
        if getattr(shape, "shape_type", None) is not None and str(shape.shape_type) == "PICTURE (13)":
            return "picture"
        if getattr(shape, "has_text_frame", False):
            return "text"
    except Exception:
        pass
    return "shape"


def _compose(parent: Callable[[float, float], Tuple[float, float]], group) -> Callable[[float, float], Tuple[float, float]]:
    """Map group child coordinates -> slide coordinates."""
    try:
        xfrm = group._element.grpSpPr.xfrm
    except Exception:
        return parent
    try:
        off_x, off_y = int(xfrm.off.x), int(xfrm.off.y)
        ext_cx, ext_cy = int(xfrm.ext.cx), int(xfrm.ext.cy)
        ch_off_x, ch_off_y = int(xfrm.chOff.x), int(xfrm.chOff.y)
        ch_ext_cx, ch_ext_cy = int(xfrm.chExt.cx), int(xfrm.chExt.cy)
    except Exception:
        return parent
    if ch_ext_cx <= 0 or ch_ext_cy <= 0:
        return parent
    sx = ext_cx / ch_ext_cx
    sy = ext_cy / ch_ext_cy

    def mapped(x: float, y: float) -> Tuple[float, float]:
        px = off_x + (x - ch_off_x) * sx
        py = off_y + (y - ch_off_y) * sy
        return parent(px, py)

    return mapped


def _identity(x: float, y: float) -> Tuple[float, float]:
    return x, y


def _map_rect(transform, left, top, width, height) -> Dict[str, int]:
    x1, y1 = transform(float(left), float(top))
    x2, y2 = transform(float(left + width), float(top + height))
    x = int(round(min(x1, x2)))
    y = int(round(min(y1, y2)))
    w = max(1, int(round(abs(x2 - x1))))
    h = max(1, int(round(abs(y2 - y1))))
    return {"x": x, "y": y, "w": w, "h": h}


def _sizes_with_fallback(sizes: List[int], total: int) -> List[int]:
    values = [max(0, int(v or 0)) for v in sizes]
    if not values:
        return []
    zero_count = sum(1 for v in values if v <= 0)
    known = sum(v for v in values if v > 0)
    remaining = max(0, int(total) - known)
    if zero_count:
        fill = remaining // zero_count if remaining > 0 else max(1, int(total) // len(values))
        values = [fill if v <= 0 else v for v in values]
    current = sum(values)
    if current <= 0:
        base = max(1, int(total) // len(values))
        values = [base] * len(values)
    elif abs(current - int(total)) > 2:
        scale = float(total) / float(current)
        values = [int(round(v * scale)) for v in values]
    return values


def _cell_span(cell) -> Tuple[int, int, bool]:
    try:
        if bool(getattr(cell, "is_spanned", False)) and not bool(getattr(cell, "is_merge_origin", False)):
            return 0, 0, True
    except Exception:
        pass
    try:
        if bool(getattr(cell, "is_merge_origin", False)):
            return int(getattr(cell, "span_width", 1) or 1), int(getattr(cell, "span_height", 1) or 1), False
    except Exception:
        pass
    return 1, 1, False


def _cell_text(cell) -> str:
    try:
        return str(cell.text or "").replace("\v", "\n")
    except Exception:
        return ""


def _table_cells(shape, slide_index: int, shape_id: str, bbox: Dict[str, int], slide_size: Dict[str, int]) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    table = shape.table
    rows = len(table.rows)
    cols = len(table.columns)
    col_widths = _sizes_with_fallback([_int(c.width) for c in table.columns], bbox["w"])
    row_heights = _sizes_with_fallback([_int(r.height) for r in table.rows], bbox["h"])
    cells: List[Dict[str, Any]] = []
    for r in range(rows):
        for c in range(cols):
            try:
                cell = table.cell(r, c)
            except Exception:
                continue
            col_span, row_span, spanned = _cell_span(cell)
            if spanned:
                continue
            text = _cell_text(cell)
            if not text.strip():
                continue
            span_w = max(1, col_span)
            span_h = max(1, row_span)
            cx = bbox["x"] + sum(col_widths[:c])
            cy = bbox["y"] + sum(row_heights[:r])
            cw = sum(col_widths[c : c + span_w])
            ch = sum(row_heights[r : r + span_h])
            if cw <= 0 or ch <= 0:
                continue
            info = _font_info(cell)
            alignment = info.get("alignment")
            try:
                anchor = cell.vertical_anchor
                if anchor is not None:
                    info["vertical_anchor"] = anchor.name
            except Exception:
                pass
            cells.append(
                {
                    "slide_index": slide_index,
                    "shape_index": 0,
                    "shape_id": f"{shape_id}_c{r}_{c}",
                    "shape_name": f"{getattr(shape, 'name', '表格')} / R{r + 1}C{c + 1}",
                    "kind": "table_cell",
                    "text": text,
                    "bbox_emu": {"x": int(cx), "y": int(cy), "w": int(cw), "h": int(ch)},
                    "slide_size_emu": slide_size,
                    "table_shape_id": shape_id,
                    "row": r,
                    "col": c,
                    "row_span": span_h,
                    "col_span": span_w,
                    "table_rows": rows,
                    "table_cols": cols,
                    "grouped": False,
                    "deletable": False,
                    "movable": False,
                    "resizable": False,
                    "alignment": alignment,
                    **{k: v for k, v in info.items() if k != "alignment"},
                }
            )
    meta = {
        "shape_id": shape_id,
        "slide_index": slide_index,
        "name": getattr(shape, "name", "表格"),
        "rows": rows,
        "cols": cols,
        "bbox_emu": bbox,
        "cell_count": rows * cols,
        "editable_cells": len(cells),
    }
    return cells, meta


def parse_pptx(pptx_path: Path | str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    path = Path(pptx_path)
    prs = Presentation(str(path))
    slide_w = _int(prs.slide_width)
    slide_h = _int(prs.slide_height)
    slide_size = {"width": slide_w, "height": slide_h}
    slides: List[Dict[str, Any]] = []
    objects: List[Dict[str, Any]] = []
    elements: List[Dict[str, Any]] = []
    tables: List[Dict[str, Any]] = []
    seq = 0

    def walk(shape, transform, slide_index: int, grouped: bool, parent_shape_id: Optional[str], group_path: List[str], depth: int = 0) -> None:
        nonlocal seq
        seq += 1
        current_seq = seq
        left = _int(getattr(shape, "left", 0))
        top = _int(getattr(shape, "top", 0))
        width = _int(getattr(shape, "width", 0))
        height = _int(getattr(shape, "height", 0))
        bbox = _map_rect(transform, left, top, width, height)
        shape_id = _shape_id(slide_index, shape, current_seq)
        kind = _shape_kind(shape)
        name = getattr(shape, "name", None)
        elements.append(
            {
                "slide_index": slide_index,
                "shape_id": shape_id,
                "shape_index": current_seq,
                "name": name,
                "kind": kind,
                "bbox_emu": bbox,
                "grouped": grouped,
                "parent_shape_id": parent_shape_id,
                "group_path": list(group_path),
                "depth": depth,
            }
        )
        if kind == "group":
            child_transform = _compose(transform, shape)
            for child in list(getattr(shape, "shapes", [])):
                walk(child, child_transform, slide_index, True, shape_id, group_path + [shape_id], depth + 1)
            return
        if kind == "table":
            try:
                table_cells, table_meta = _table_cells(shape, slide_index, shape_id, bbox, slide_size)
            except Exception:
                return
            for cell in table_cells:
                cell["shape_index"] = current_seq
                cell["parent_shape_id"] = shape_id
                cell["grouped"] = grouped
                cell["group_path"] = list(group_path)
                objects.append(cell)
            table_meta["grouped"] = grouped
            tables.append(table_meta)
            return
        if not getattr(shape, "has_text_frame", False):
            return
        try:
            text = str(shape.text or "").replace("\v", "\n")
        except Exception:
            return
        if not text.strip():
            return
        info = _font_info(shape, None if grouped else prs.slides[slide_index])
        objects.append(
            {
                "slide_index": slide_index,
                "shape_index": current_seq,
                "shape_id": shape_id,
                "shape_name": name,
                "kind": "text",
                "text": text,
                "bbox_emu": bbox,
                "slide_size_emu": slide_size,
                "grouped": grouped,
                "parent_shape_id": parent_shape_id,
                "group_path": list(group_path),
                "deletable": not grouped,
                "movable": not grouped,
                "resizable": not grouped,
                **info,
            }
        )

    for slide_index, slide in enumerate(prs.slides):
        slides.append({"slide_index": slide_index})
        for shape in slide.shapes:
            walk(shape, _identity, slide_index, False, None, [])

    payload = {
        "slide_count": len(slides),
        "slide_size_emu": slide_size,
        "slides": slides,
        "objects": objects,
        "elements": elements,
        "tables": tables,
    }
    return objects, payload
