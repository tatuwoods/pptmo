"""SmartArt (DrawingML diagram) helpers shared by parse and edits.

python-pptx has no SmartArt API, so this works on the OPC parts directly:
the data model part (``dgm:dataModel``) is the authoritative text source, and
the cached drawing part (``dsp:drawing``) is what LibreOffice actually paints,
so text edits must land in both or the preview keeps showing the old words.
"""
from __future__ import annotations

import re
from typing import Iterator, List, Optional, Tuple

from lxml import etree
from pptx.oxml.ns import qn

DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_DGM = DIAGRAM_URI
# pptx.oxml.ns.qn() has no "dgm" prefix — build Clark-notation names directly
_DGM_RELIDS = f"{{{_DGM}}}relIds"
_DGM_PT = f"{{{_DGM}}}pt"
_DGM_T = f"{{{_DGM}}}t"


def _rel_ids(graphic_data) -> dict:
    """r:dm/r:lo/r:qs/r:cs → the rIds a dgm:relIds element declares.

    ``drawing`` carries the rId from the MS extension ``dsp:dataModelExt``
    inside the same graphicData (real decks); fixture decks relate the cached
    drawing straight off the slide, which ``find_frame_data`` falls back to.
    """
    out = {}
    for child in graphic_data:
        if child.tag == _DGM_RELIDS:
            for attr in ("dm", "lo", "qs", "cs"):
                value = child.get(qn(f"r:{attr}"))
                if value:
                    out[attr] = value
    for node in graphic_data.iter():
        if etree.QName(node).localname == "dataModelExt":
            for attr, value in node.attrib.items():
                if etree.QName(attr).localname == "relId" and value:
                    out["drawing"] = value
    return out


def find_frame_data(shape, slide) -> Optional[dict]:
    """Resolve a SmartArt graphicFrame shape to its diagram parts.

    Returns ``{"shape_id": ..., "data_part": ..., "drawing_part": ...}`` where
    the parts are python-pptx opc parts (either may be None when absent).
    """
    element = getattr(shape, "_element", None)
    if element is None:
        return None
    for graphic_data in element.iter(qn("a:graphicData")):
        if str(graphic_data.get("uri") or "") != DIAGRAM_URI:
            continue
        rels = slide.part.rels
        data_part = drawing_part = None
        ids = _rel_ids(graphic_data)
        dm = ids.get("dm")
        if dm and dm in rels:
            data_part = rels[dm].target_part
        drawing_rid = ids.get("drawing")
        if drawing_rid and drawing_rid in rels:
            drawing_part = rels[drawing_rid].target_part
        else:
            # fixture decks relate the cached drawing straight off the slide
            # (diagramDrawing type) without the extension-list indirection
            for rel in rels.values():
                if rel.reltype.endswith("/diagramDrawing"):
                    drawing_part = rel.target_part
        return {"shape": shape, "data_part": data_part, "drawing_part": drawing_part}
    return None


def iter_smartart(slide) -> Iterator[dict]:
    for shape in slide.shapes:
        info = find_frame_data(shape, slide)
        if info is not None:
            yield info


def _parse(blob: bytes):
    return etree.fromstring(blob)


def node_texts(blob: bytes) -> List[str]:
    """Text of every dgm:pt that carries a dgm:t, in document order."""
    if not blob:
        return []
    try:
        root = _parse(blob)
    except etree.XMLSyntaxError:
        return []
    texts: List[str] = []
    for pt in root.iter(_DGM_PT):
        t = pt.find(_DGM_T)
        if t is None:
            continue
        parts = [node.text or "" for node in t.iter(qn("a:t"))]
        joined = "".join(parts).strip()
        if joined:
            texts.append(joined)
    return texts


def set_node_text(blob: bytes, node_index: int, text: str) -> Optional[bytes]:
    """Rewrite one text-bearing node (0-based over ``node_texts`` order).

    The first a:t of the node gets the full new text; any further runs are
    emptied so no stale characters survive. Returns the new blob, or None when
    the index is out of range / the part is unparseable.
    """
    try:
        root = _parse(blob)
    except etree.XMLSyntaxError:
        return None
    index = 0
    for pt in root.iter(_DGM_PT):
        t = pt.find(_DGM_T)
        if t is None:
            continue
        runs = list(t.iter(qn("a:t")))
        joined = "".join(node.text or "" for node in runs).strip()
        if not joined:
            continue
        if index == node_index:
            if not runs:  # stripped text came from somewhere else; nothing to write
                return None
            runs[0].text = str(text)
            for extra in runs[1:]:
                extra.text = ""
            return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        index += 1
    return None


def replace_text(blob: bytes, find: str, replace: str) -> Tuple[Optional[bytes], int]:
    """Replace every occurrence of ``find`` inside a:t runs; (new blob, count)."""
    try:
        root = _parse(blob)
    except etree.XMLSyntaxError:
        return None, 0
    count = 0
    for node in root.iter(qn("a:t")):
        current = node.text or ""
        if find in current:
            count += current.count(find)
            node.text = current.replace(find, replace)
    if not count:
        return None, 0
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), count


def write_part(part, blob: bytes) -> None:
    """Swap a generic opc part's payload in place."""
    part._blob = blob
