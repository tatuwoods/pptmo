"""v0.6 #2：OOXML passthrough 保真 —— 未触碰的 slide part 原字节保留。

python-pptx 保存时整包重序列化：即使某页没被编辑，它的 slideN.xml 也可能
因为属性顺序/命名空间前缀/自闭合标签差异而字节变化；更糟的是部分未被
schema 映射的内容（动画 timing 细节、第三方扩展元素）在极端情况下丢属性。

策略（对标 CasualOffice 的未映射 XML 透传）：apply 保存后，把「原索引未被
触碰」的 slide part（连同它的 .rels 与 notesSlide part）用源包里的原始字节
写回。被编辑的 part 走 python-pptx 新序列化结果。映射按 part 名而不是目标
位置——move_slide 重排后 part 名跟着内容走，按名还原不会错页。
"""
from __future__ import annotations

import posixpath
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Set

_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_SLIDE_REL_TYPE = f"{_NS_R}/notesSlide"


def _zip_name(part_name: str) -> str:
    return part_name.lstrip("/")


def _normalize_target(base_zip_name: str, target: str) -> str:
    """rels 里的相对 target（如 ../notesSlides/notesSlide1.xml）→ zip 内绝对路径。"""
    if target.startswith("/"):
        return target.lstrip("/")
    base_dir = posixpath.dirname(base_zip_name)
    return posixpath.normpath(posixpath.join(base_dir, target))


def _rels_for(slide_zip: str) -> str:
    # ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels
    parent = posixpath.dirname(slide_zip)
    name = posixpath.basename(slide_zip)
    return posixpath.join(parent, "_rels", f"{name}.rels")


def _parse_xml(data: bytes):
    import xml.etree.ElementTree as ET

    return ET.fromstring(data)


def slide_part_order(pptx_path: Path | str) -> List[str]:
    """presentation.xml 的 sldIdLst 顺序 → slide part zip 路径列表。"""
    with zipfile.ZipFile(pptx_path) as zf:
        pres = _parse_xml(zf.read("ppt/presentation.xml"))
        rels = _parse_xml(zf.read("ppt/_rels/presentation.xml.rels"))
    rid_target: Dict[str, str] = {}
    for rel in rels.findall(f"{{{_REL_NS}}}Relationship"):
        rid_target[rel.get("Id")] = rel.get("Target")
    order: List[str] = []
    for sld_id in pres.findall(f".//{{{_NS_P}}}sldIdLst/{{{_NS_P}}}sldId"):
        rid = sld_id.get(f"{{{_NS_R}}}id")
        target = rid_target.get(rid)
        if target:
            order.append(_normalize_target("ppt/presentation.xml", target))
    return order


def _notes_part_for(zf: zipfile.ZipFile, slide_zip: str) -> Optional[str]:
    rels_name = _rels_for(slide_zip)
    try:
        rels = _parse_xml(zf.read(rels_name))
    except KeyError:
        return None
    for rel in rels.findall(f"{{{_REL_NS}}}Relationship"):
        if rel.get("Type") == _SLIDE_REL_TYPE:
            return _normalize_target(slide_zip, rel.get("Target") or "")
    return None


def preserve_untouched_slides(src: Path | str, dst: Path | str, touched: Set[int]) -> int:
    """把 src 中「未被触碰页」的 slide part（+rels+notes）原字节写回 dst。

    touched 是 0-based 原始页索引（apply 流程的返回值）。返回还原的 slide
    part 数量。dst 中不存在的 part（已删除）跳过；src 中不存在的（新增页）
    不会被还原。
    """
    src, dst = Path(src), Path(dst)
    if not src.exists() or not dst.exists():
        return 0
    try:
        src_order = slide_part_order(src)
    except (KeyError, zipfile.BadZipFile):
        return 0

    restore: Set[str] = set()
    with zipfile.ZipFile(src) as zs:
        for index, slide_zip in enumerate(src_order):
            if index in touched:
                continue
            if not slide_zip:
                continue
            restore.add(slide_zip)
            restore.add(_rels_for(slide_zip))
            notes = _notes_part_for(zs, slide_zip)
            if notes:
                restore.add(notes)
                restore.add(_rels_for(notes))
    restore.discard("")

    with zipfile.ZipFile(src) as zs:
        src_bytes = {name: zs.read(name) for name in restore if name in zs.namelist()}

    restored_slides = 0
    fd, tmp_name = tempfile.mkstemp(dir=str(dst.parent), suffix=".passthrough")
    import os

    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(dst) as zd, zipfile.ZipFile(tmp_path, "w") as zt:
            for info in zd.infolist():
                data = zd.read(info.filename)
                if info.filename in src_bytes:
                    candidate = src_bytes[info.filename]
                    if candidate:
                        if info.filename in {
                            name
                            for name in src_order
                            if name and name in src_bytes
                        }:
                            restored_slides += 1
                        data = candidate
                new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                new_info.compress_type = info.compress_type
                new_info.external_attr = info.external_attr
                zt.writestr(new_info, data)
        shutil.move(str(tmp_path), str(dst))
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise
    return restored_slides


def passthrough_enabled() -> bool:
    import os

    return os.getenv("PPTMO_PASSTHROUGH", "1").strip().lower() not in ("0", "false", "no", "off")
