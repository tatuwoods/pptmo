"""上传/渲染前的 Wingdings 兼容修复（macOS 无 Wingdings 字体导致 LibreOffice 预览乱码）。

根因：Office 项目符号常用 buFont=Wingdings + buChar=私用区字符（U+F020-U+F0FF，
如 n=■、ü=✓）。macOS 没有 Wingdings 字体，LibreOffice 映射不了这些私用码，
预览 PDF/PNG 显示为豆腐或直接输出 U+F0xx 乱码字符。

修法：改符号字体声明为通用字体，并把私用码换成 Unicode 等价符号。
不改文字语义，只动符号。
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

# Wingdings 常用码位 → Unicode 等价符号
_WINGDINGS_MAP = {
    0x21: "\u270f",  # ✏
    0x22: "\u2702",  # ✂
    0x26: "\U0001F4DA",  # 📚
    0x28: "\u260e",  # ☎
    0x29: "\u2709",  # ✉
    0x2A: "\u270e",  # ✎
    0x2B: "\u270d",  # ✍
    0x37: "\u231b",  # ⌛
    0x3E: "\u2712",  # ✒
    0x4D: "\u2b1b",  # ⬛
    0x50: "\U0001F4E6",  # 📦
    0x51: "\U0001F4CB",  # 📋
    0x54: "\U0001F4C1",  # 📁
    0x5C: "\U0001F5D1",  # 🗑
    0x6E: "\u25a0",  # ■
    0x6F: "\u25a1",  # □
    0x70: "\u2610",  # ☐
    0x71: "\u2b1c",  # ⬜
    0x72: "\u2b1b",  # ⬛
    0x73: "\u25aa",  # ▪
    0x74: "\u25ab",  # ▫
    0x75: "\u25fe",  # ◾
    0x76: "\u25fd",  # ◽
    0xA8: "\u2610",  # ☐
    0xFB: "\u2714",  # ✔
    0xFC: "\u2713",  # ✓
    0xFD: "\u2717",  # ✗
    0xFE: "\u2718",  # ✘
    0xB6: "\u25c6",  # ◆
    0xB7: "\u2756",  # ❖
    0xD8: "\u27a1",  # ➡
}

_FALLBACK = "\u25cf"  # ●

# run 文本里的 F0xx → Unicode
_TEXT_MAP = {}
for _c, _u in _WINGDINGS_MAP.items():
    _TEXT_MAP[0xF000 + _c] = _u


def _fix_bullets_on_text_frame(tf) -> int:
    fixed = 0
    for para in tf.paragraphs:
        pPr = para._pPr
        if pPr is None:
            continue
        buFont = pPr.find(qn("a:buFont"))
        buChar = pPr.find(qn("a:buChar"))
        if buFont is None or buChar is None:
            continue
        typeface = (buFont.get("typeface") or "").lower()
        if "wingdings" not in typeface and "webdings" not in typeface and typeface != "symbol":
            continue
        ch = buChar.get("char") or ""
        try:
            code = ord(ch[0]) if ch else 0
        except TypeError:
            continue
        if 0xE000 <= code <= 0xF8FF:
            uni = _WINGDINGS_MAP.get(code - 0xF000)
            buChar.set("char", uni if uni else _FALLBACK)
            buFont.set("typeface", "Helvetica")
            fixed += 1
    return fixed


def _fix_text_runs(tf) -> int:
    fixed = 0
    for para in tf.paragraphs:
        for run in para.runs:
            t = run.text
            if not t:
                continue
            if not any(0xE000 <= ord(c) <= 0xF8FF for c in t):
                continue
            new = []
            for c in t:
                code = ord(c)
                if 0xE000 <= code <= 0xF8FF:
                    new.append(_TEXT_MAP.get(code, _FALLBACK))
                    fixed += 1
                else:
                    new.append(c)
            run.text = "".join(new)
    return fixed


def sanitize_wingdings(pptx_path: str | Path) -> tuple[int, int]:
    """就地修复 PPTX 的 Wingdings 乱码隐患。返回 (修复段落符号数, 修复文本字符数)。"""
    prs = Presentation(str(pptx_path))
    bullet_fixed = 0
    text_fixed = 0

    def walk(shapes):
        nonlocal bullet_fixed, text_fixed
        for shape in shapes:
            if shape.has_text_frame:
                bullet_fixed += _fix_bullets_on_text_frame(shape.text_frame)
                text_fixed += _fix_text_runs(shape.text_frame)
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        bullet_fixed += _fix_bullets_on_text_frame(cell.text_frame)
                        text_fixed += _fix_text_runs(cell.text_frame)
            if shape.shape_type == 6:  # GROUP
                walk(getattr(shape, "shapes", []))

    for slide in prs.slides:
        walk(slide.shapes)
    prs.save(str(pptx_path))
    return bullet_fixed, text_fixed
