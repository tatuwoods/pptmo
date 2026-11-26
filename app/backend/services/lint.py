"""v0.5 #7 + v0.6 #3：后置 lint → 设计建议引擎。

v0.5（对标 hands-on-deck）：apply 后自动跑一遍，把「文字溢出框外 / 字号超界 /
对象越画布 / 重叠」四类问题以结构化 issues 返回给 AI；fix_issues（edits.py）
按同一套判定做确定性修复（缩字号/收回画布），重叠只告警——挪谁都不一定对，
交给 AI 决策。

v0.6（对标 Kition 的 designer 建议）：升级为设计建议引擎——新增字号层级混乱
（同页标题 < 正文）、对比度过低（文字 vs 形状填充/页面背景）、图片拉伸变形
（非等比缩放）三类检测，且每条 issue 携带「设计师会怎么改」的自然语言
suggestion，让 AI 拿到的不是干巴巴的错误码而是可执行的改法。

检查是几何启发式（无渲染回流），判定都留了容忍度：宁可少报不误报。
"""
from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Pt

MIN_FONT_PT = 8.0
MAX_FONT_PT = 96.0
EMU_PER_PT = 12700
DEFAULT_FONT_PT = 18.0
LINE_HEIGHT_FACTOR = 1.22     # 行高 ≈ 字号 × 1.22（含行距余量）
FIT_TOLERANCE = 1.15          # 溢出判定容忍：估算误差 ≥15% 才报
OVERLAP_RATIO = 0.5           # 重叠面积 / 较小形状面积 ≥ 50% 才告警
MAX_ISSUES = 200
CONTRAST_MIN_RATIO = 3.0      # WCAG 大字（≥18pt）可读下限 3:1
DISTORTION_TOLERANCE = 0.15   # 显示长宽比偏离原图 ≥15% 才算拉伸变形
MIN_NATIVE_PX = 8             # 原图像素低于该值视为占位图，不做变形检查

FIXABLE_KINDS = ("text_overflow", "font_size_out_of_range", "shape_off_canvas")

# v0.6 #3：每类问题的「设计师会怎么改」一句话建议
SUGGESTIONS: Dict[str, str] = {
    "text_overflow": "缩小字号、加宽文本框或精简文案，让文字完整落在框内",
    "font_size_out_of_range": "把字号拉回 8–96pt 常规区间（正文建议 12–24pt、标题 24–40pt）",
    "shape_off_canvas": "把对象拖回画布内或删除多余元素，至少保证完整可见",
    "shape_overlap": "错开两个对象的位置或缩小其中一个，避免关键内容互相遮挡",
    "font_hierarchy": "放大标题或缩小正文，保持同页「标题 > 正文」的字号层级",
    "low_contrast": "把文字或背景加深/提亮一档，对比度至少 3:1（WCAG 大字可读标准）",
    "image_distorted": "按原图长宽比等比缩放（或先裁剪再放置），避免画面变形",
    "lint_truncated": "问题过多，先集中修复再复查",
}


def _with_suggestion(issue: Dict[str, Any]) -> Dict[str, Any]:
    issue.setdefault("suggestion", SUGGESTIONS.get(str(issue.get("kind")), ""))
    return issue

_CJK_RE = re.compile(r"[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]")


def _int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _group_transform(group) -> Optional[Callable[[int, int, int, int], Tuple[int, int, int, int]]]:
    """组内子形状坐标 → 画布绝对坐标（grpSpPr xfrm 的 off/ext/chOff/chExt 映射）。"""
    try:
        xfrm = group._element.grpSpPr.xfrm
        if xfrm is None:
            return None
        off_x, off_y = _int(xfrm.off.x), _int(xfrm.off.y)
        ext_cx, ext_cy = _int(xfrm.ext.cx), _int(xfrm.ext.cy)
        ch_x, ch_y = _int(xfrm.chOff.x), _int(xfrm.chOff.y)
        ch_cx, ch_cy = _int(xfrm.chExt.cx), _int(xfrm.chExt.cy)
        if ch_cx <= 0 or ch_cy <= 0 or ext_cx <= 0 or ext_cy <= 0:
            return None
        sx, sy = ext_cx / ch_cx, ext_cy / ch_cy

        def mapped(left: int, top: int, width: int, height: int):
            return (
                off_x + int((_int(left) - ch_x) * sx),
                off_y + int((_int(top) - ch_y) * sy),
                int(_int(width) * sx),
                int(_int(height) * sy),
            )

        return mapped
    except Exception:
        return None


def _iter_shapes(shapes, transform=None) -> Iterator[Tuple[Any, Tuple[int, int, int, int], bool]]:
    """递归走形状树：yield (shape, 绝对 bbox (l,t,w,h), is_group)。"""
    for shape in shapes:
        try:
            left, top = _int(shape.left), _int(shape.top)
            width, height = _int(shape.width), _int(shape.height)
        except Exception:
            continue
        bbox = (left, top, width, height)
        if transform is not None:
            bbox = transform(*bbox)
        is_group = str(getattr(shape, "shape_type", "")) == "GROUP (6)" or getattr(shape, "shape_type", None) == 6
        yield shape, bbox, is_group
        if is_group:
            child_transform = _group_transform(shape)
            if child_transform is not None:
                yield from _iter_shapes(shape.shapes, child_transform)
            else:
                # 拿不到映射时按父级坐标近似，总比漏检好
                yield from _iter_shapes(shape.shapes, transform)


def _run_sizes(shape) -> List[float]:
    """显式设置过的字号列表（pt）；继承字号不在 XML 里，算不进超界检查。"""
    sizes: List[float] = []
    if not getattr(shape, "has_text_frame", False):
        return sizes
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                sizes.append(float(run.font.size.pt))
    return sizes


def _max_font_pt(shape) -> float:
    sizes = _run_sizes(shape)
    return max(sizes) if sizes else DEFAULT_FONT_PT


def _effective_sizes(shape) -> List[float]:
    """用于溢出估算的字号：显式字号优先，无显式则按默认 18pt 估。"""
    sizes = _run_sizes(shape)
    if sizes:
        return sizes
    text = _shape_text(shape)
    return [DEFAULT_FONT_PT] if text else []


def _shape_text(shape) -> str:
    if not getattr(shape, "has_text_frame", False):
        return ""
    try:
        return str(shape.text or "")
    except Exception:
        return ""


def _estimate_lines(text: str, width_pt: float, size_pt: float) -> int:
    """按字宽估算换行后的行数：CJK 全宽、拉丁 0.6 宽。"""
    if not text:
        return 0
    cjk = len(_CJK_RE.findall(text))
    ratio = cjk / max(1, len(text))
    char_w = max(0.1, size_pt * (ratio * 1.0 + (1 - ratio) * 0.6))
    per_line = max(1.0, width_pt / char_w)
    lines = 0
    for para in text.split("\n"):
        stripped = para.strip()
        lines += max(1, math.ceil(len(stripped) / per_line)) if stripped else 1
    return lines


def text_fits(shape, bbox: Tuple[int, int, int, int]) -> bool:
    """文本是否装得下（含容忍度）；无文本视为装得下。"""
    text = _shape_text(shape)
    if not text.strip():
        return True
    _l, _t, width, height = bbox
    width_pt, height_pt = width / EMU_PER_PT, height / EMU_PER_PT
    if width_pt <= 0 or height_pt <= 0:
        return False
    size = _max_font_pt(shape)
    needed = _estimate_lines(text, width_pt, size)
    available = height_pt / (size * LINE_HEIGHT_FACTOR)
    return needed <= available * FIT_TOLERANCE


def _slide_bounds_issues(slide, slide_w: int, slide_h: int, slide_index: int) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    leaf_boxes: List[Tuple[str, Tuple[int, int, int, int]]] = []
    for shape, bbox, is_group in _iter_shapes(slide.shapes):
        left, top, width, height = bbox
        shape_id = f"s{slide_index}_id{_int(getattr(shape, 'shape_id', 0))}"
        name = str(getattr(shape, "name", "") or "")
        if width > 0 and height > 0 and (
            left < 0 or top < 0 or left + width > slide_w or top + height > slide_h
        ):
            issues.append(_with_suggestion({
                "kind": "shape_off_canvas",
                "slide_index": slide_index,
                "shape_id": shape_id,
                "shape_name": name,
                "severity": "error",
                "fixable": True,
                "message": f"对象「{name}」超出画布边界（{left},{top},{width}x{height}）",
            }))
        if not is_group and width > 0 and height > 0 and _shape_text(shape).strip():
            leaf_boxes.append((shape_id, bbox))
        if len(issues) >= MAX_ISSUES:
            return issues
    # 重叠：只对「有文字的叶子形状」两两检查，面积过半才算
    for i in range(len(leaf_boxes)):
        for j in range(i + 1, len(leaf_boxes)):
            id_a, (l1, t1, w1, h1) = leaf_boxes[i]
            id_b, (l2, t2, w2, h2) = leaf_boxes[j]
            inter_w = min(l1 + w1, l2 + w2) - max(l1, l2)
            inter_h = min(t1 + h1, t2 + h2) - max(t1, t2)
            if inter_w <= 0 or inter_h <= 0:
                continue
            inter = inter_w * inter_h
            if inter / min(w1 * h1, w2 * h2) < OVERLAP_RATIO:
                continue
            issues.append(_with_suggestion({
                "kind": "shape_overlap",
                "slide_index": slide_index,
                "shape_id": id_a,
                "other_shape_id": id_b,
                "severity": "warning",
                "fixable": False,
                "message": f"第 {slide_index + 1} 页对象 {id_a} 与 {id_b} 重叠过半（自动挪动有风险，请人工确认）",
            }))
            if len(issues) >= MAX_ISSUES:
                return issues
    return issues


# ---------- v0.6 #3：设计建议三类新检查 ----------


def _solid_fill_rgb(shape) -> Optional[str]:
    """形状 solidFill 的 RGB hex；非纯色填充/主题色/读取失败返回 None。"""
    try:
        fill = shape.fill
        if fill.type is None or int(fill.type) != 1:  # MSO_FILL.SOLID == 1
            return None
        rgb = fill.fore_color.rgb  # 主题色此处抛异常
        return str(rgb)
    except Exception:
        return None


def _slide_background_rgb(slide) -> Optional[str]:
    """页面背景纯色填充的 RGB hex；无/非纯色返回 None。"""
    try:
        fill = slide.background.fill
        if fill.type is None or int(fill.type) != 1:
            return None
        return str(fill.fore_color.rgb)
    except Exception:
        return None


def _run_font_rgb(shape, theme_parsed: Optional[Dict[str, Dict[str, str]]] = None) -> Optional[str]:
    """首个显式 run 字体色 hex：显式 RGB 优先，主题色经 theme+clrMap 解析；都没有返回 None。"""
    if not getattr(shape, "has_text_frame", False):
        return None
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            try:
                color = run.font.color
                try:
                    rgb = color.rgb  # SCHEME 色在此抛 TypeError，走主题解析分支
                except Exception:
                    rgb = None
                if rgb is not None:
                    return str(rgb)
                if theme_parsed is not None:
                    if getattr(color, "brightness", 0) or 0:
                        continue  # 带亮度修饰（tint/shade）的主题色还原不准，不猜
                    resolved = _resolve_theme_hex(color.theme_color, theme_parsed)
                    if resolved:
                        return resolved
            except Exception:
                continue
    return None


# ---------- 主题色解析：schemeClr（如 bg1/tx1）→ 具体 hex ----------

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_SYS_COLOR_FALLBACK = {"window": "FFFFFF", "windowText": "000000"}
# clrMap 缺失时的常规缺省（浅色主题惯例：bg1→lt1、tx1→dk1）
_CLR_MAP_DEFAULT = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}

_THEME_COLOR_TO_SLOT: Dict[Any, str] = {
    MSO_THEME_COLOR.ACCENT_1: "accent1",
    MSO_THEME_COLOR.ACCENT_2: "accent2",
    MSO_THEME_COLOR.ACCENT_3: "accent3",
    MSO_THEME_COLOR.ACCENT_4: "accent4",
    MSO_THEME_COLOR.ACCENT_5: "accent5",
    MSO_THEME_COLOR.ACCENT_6: "accent6",
    MSO_THEME_COLOR.BACKGROUND_1: "bg1",
    MSO_THEME_COLOR.BACKGROUND_2: "bg2",
    MSO_THEME_COLOR.TEXT_1: "tx1",
    MSO_THEME_COLOR.TEXT_2: "tx2",
    MSO_THEME_COLOR.DARK_1: "dk1",
    MSO_THEME_COLOR.DARK_2: "dk2",
    MSO_THEME_COLOR.LIGHT_1: "lt1",
    MSO_THEME_COLOR.LIGHT_2: "lt2",
    MSO_THEME_COLOR.HYPERLINK: "hlink",
    MSO_THEME_COLOR.FOLLOWED_HYPERLINK: "folHlink",
}


def _theme_parsed(slide, cache: Dict[str, Dict[str, Dict[str, str]]]) -> Dict[str, Dict[str, str]]:
    """该页所属 master 的 {"scheme": {slot: hex}, "map": {逻辑槽: scheme 槽}}；解析失败返回空结构。

    同一 master 只解析一次（按 partname 缓存在本次 lint 调用内）。
    """
    try:
        master = slide.slide_layout.slide_master
    except Exception:
        return {"scheme": {}, "map": {}}
    key = str(master.part.partname)
    if key not in cache:
        scheme: Dict[str, str] = {}
        clr_map: Dict[str, str] = {}
        try:
            theme_part = master.part.part_related_by(RT.THEME)
            root = etree.fromstring(theme_part.blob)
            scheme_el = root.find(f".//{{{_A_NS}}}clrScheme")
            if scheme_el is not None:
                for slot_el in scheme_el:
                    slot = etree.QName(slot_el).localname
                    srgb = slot_el.find(f"{{{_A_NS}}}srgbClr")
                    if srgb is not None and srgb.get("val"):
                        scheme[slot] = str(srgb.get("val"))
                        continue
                    sys_el = slot_el.find(f"{{{_A_NS}}}sysClr")
                    if sys_el is not None:
                        fallback = sys_el.get("lastClr") or _SYS_COLOR_FALLBACK.get(str(sys_el.get("val") or ""))
                        if fallback:
                            scheme[slot] = fallback
            clr_map_el = master.element.find(qn("p:clrMap"))
            if clr_map_el is not None:
                clr_map = {str(k): str(v) for k, v in clr_map_el.attrib.items()}
        except Exception:
            pass
        cache[key] = {"scheme": scheme, "map": clr_map}
    return cache[key]


def _resolve_theme_hex(theme_color, parsed: Dict[str, Dict[str, str]]) -> Optional[str]:
    """MSO_THEME_COLOR → hex：逻辑槽（bg1/tx1/…）先经 master clrMap 映射到 scheme 槽。"""
    slot = _THEME_COLOR_TO_SLOT.get(theme_color)
    if not slot:
        return None
    if slot in _CLR_MAP_DEFAULT:
        slot = parsed["map"].get(slot) or _CLR_MAP_DEFAULT[slot]
    return parsed["scheme"].get(slot) or None


def _rel_luminance(hex_rgb: str) -> float:
    """sRGB 相对亮度（WCAG 公式）。"""
    value = hex_rgb.strip().lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    if len(value) != 6:
        return 0.0
    channels = []
    for index in (0, 2, 4):
        raw = int(value[index : index + 2], 16) / 255.0
        channels.append(raw / 12.92 if raw <= 0.04045 else ((raw + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def _contrast_ratio(fg: str, bg: str) -> float:
    l1, l2 = _rel_luminance(fg), _rel_luminance(bg)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _is_title_placeholder(shape) -> bool:
    try:
        if not shape.is_placeholder:
            return False
        return shape.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
    except Exception:
        return False


def _slide_design_issues(
    slide,
    slide_index: int,
    theme_cache: Optional[Dict[str, Dict[str, Dict[str, str]]]] = None,
) -> List[Dict[str, Any]]:
    """字号层级混乱 / 对比度过低 / 图片拉伸变形（全部 warn-only，交给 AI 决策）。"""
    issues: List[Dict[str, Any]] = []
    theme_cache = theme_cache if theme_cache is not None else {}
    theme_parsed = _theme_parsed(slide, theme_cache)

    # 1) 字号层级：标题占位符与正文都有显式字号时才比较（继承字号不猜）
    title_max: Optional[float] = None
    body_max: Optional[float] = None
    for shape, _bbox, is_group in _iter_shapes(slide.shapes):
        sizes = _run_sizes(shape)
        if not sizes or not _shape_text(shape).strip():
            continue
        biggest = max(sizes)
        if _is_title_placeholder(shape):
            title_max = biggest if title_max is None else max(title_max, biggest)
        elif not is_group:
            body_max = biggest if body_max is None else max(body_max, biggest)
    if title_max is not None and body_max is not None and title_max < body_max:
        issues.append(_with_suggestion({
            "kind": "font_hierarchy",
            "slide_index": slide_index,
            "shape_id": f"s{slide_index}_title",
            "severity": "warning",
            "fixable": False,
            "title_size_pt": title_max,
            "body_size_pt": body_max,
            "message": (
                f"第 {slide_index + 1} 页标题字号（{title_max:g}pt）小于正文（{body_max:g}pt），层级混乱"
            ),
        }))

    # 2) 对比度：文字色（显式 RGB 或可解析的主题色）vs 形状填充，无填充时看页面背景
    background = _slide_background_rgb(slide)
    for shape, _bbox, _is_group in _iter_shapes(slide.shapes):
        if len(issues) >= MAX_ISSUES:
            break
        font_rgb = _run_font_rgb(shape, theme_parsed)
        if not font_rgb or not _shape_text(shape).strip():
            continue
        backdrop = _solid_fill_rgb(shape) or background
        if not backdrop:
            continue
        ratio = _contrast_ratio(font_rgb, backdrop)
        if ratio < CONTRAST_MIN_RATIO:
            shape_id = f"s{slide_index}_id{_int(getattr(shape, 'shape_id', 0))}"
            name = str(getattr(shape, "name", "") or "")
            issues.append(_with_suggestion({
                "kind": "low_contrast",
                "slide_index": slide_index,
                "shape_id": shape_id,
                "shape_name": name,
                "severity": "warning",
                "fixable": False,
                "contrast_ratio": round(ratio, 2),
                "message": (
                    f"第 {slide_index + 1} 页「{name}」文字与背景对比度仅 {ratio:.1f}:1"
                    f"（低于 {CONTRAST_MIN_RATIO:g}:1），难以阅读"
                ),
            }))

    # 3) 图片拉伸变形：显示长宽比 vs 原图（扣除裁剪区）长宽比
    for shape, bbox, _is_group in _iter_shapes(slide.shapes):
        if len(issues) >= MAX_ISSUES:
            break
        if getattr(shape, "shape_type", None) != MSO_SHAPE_TYPE.PICTURE:
            continue
        _l, _t, width, height = bbox
        if width <= 0 or height <= 0:
            continue
        try:
            native_w, native_h = shape.image.size
        except Exception:
            continue  # 链接图/读取失败不猜
        if min(native_w, native_h) < MIN_NATIVE_PX:
            continue  # 占位小图（如 1x1）无可见变形
        try:
            crop_l = max(0.0, min(0.9, float(shape.crop_left or 0)))
            crop_r = max(0.0, min(0.9, float(shape.crop_right or 0)))
            crop_t = max(0.0, min(0.9, float(shape.crop_top or 0)))
            crop_b = max(0.0, min(0.9, float(shape.crop_bottom or 0)))
        except Exception:
            crop_l = crop_r = crop_t = crop_b = 0.0
        if any((crop_l, crop_r, crop_t, crop_b)):
            continue  # 裁剪过的图片宽高比本来就该不同，不按变形报（宁缺毋滥）
        native_aspect = native_w / native_h
        shown_aspect = width / height
        if abs(shown_aspect - native_aspect) / native_aspect > DISTORTION_TOLERANCE:
            shape_id = f"s{slide_index}_id{_int(getattr(shape, 'shape_id', 0))}"
            name = str(getattr(shape, "name", "") or "")
            issues.append(_with_suggestion({
                "kind": "image_distorted",
                "slide_index": slide_index,
                "shape_id": shape_id,
                "shape_name": name,
                "severity": "warning",
                "fixable": False,
                "native_aspect": round(native_aspect, 3),
                "shown_aspect": round(shown_aspect, 3),
                "message": (
                    f"第 {slide_index + 1} 页图片「{name}」被非等比拉伸"
                    f"（原图比例 {native_aspect:.2f}，显示比例 {shown_aspect:.2f}）"
                ),
            }))
    return issues


def lint_pptx(pptx_path: Path | str) -> List[Dict[str, Any]]:
    """全稿检查：文字溢出 / 字号超界 / 越画布 / 重叠。上限 MAX_ISSUES 条。"""
    prs = Presentation(str(pptx_path))
    slide_w, slide_h = _int(prs.slide_width), _int(prs.slide_height)
    issues: List[Dict[str, Any]] = []
    theme_cache: Dict[str, Dict[str, Dict[str, str]]] = {}
    for slide_index, slide in enumerate(prs.slides):
        for shape, bbox, _is_group in _iter_shapes(slide.shapes):
            if len(issues) >= MAX_ISSUES:
                issues.append({"kind": "lint_truncated", "severity": "warning",
                               "message": f"问题过多，只保留前 {MAX_ISSUES} 条"})
                return issues
            shape_id = f"s{slide_index}_id{_int(getattr(shape, 'shape_id', 0))}"
            name = str(getattr(shape, "name", "") or "")
            for size in _run_sizes(shape):
                if size < MIN_FONT_PT or size > MAX_FONT_PT:
                    issues.append(_with_suggestion({
                        "kind": "font_size_out_of_range",
                        "slide_index": slide_index,
                        "shape_id": shape_id,
                        "shape_name": name,
                        "severity": "error",
                        "fixable": True,
                        "font_size_pt": size,
                        "message": f"「{name}」字号 {size:g}pt 超出 [{MIN_FONT_PT:g}, {MAX_FONT_PT:g}] 界",
                    }))
            if not text_fits(shape, bbox):
                issues.append(_with_suggestion({
                    "kind": "text_overflow",
                    "slide_index": slide_index,
                    "shape_id": shape_id,
                    "shape_name": name,
                    "severity": "error",
                    "fixable": True,
                    "message": f"「{name}」文字装不进文本框（估算溢出 ≥{int((FIT_TOLERANCE - 1) * 100)}%）",
                }))
        issues.extend(_slide_bounds_issues(slide, slide_w, slide_h, slide_index))
        issues.extend(_slide_design_issues(slide, slide_index, theme_cache))
        if len(issues) >= MAX_ISSUES:
            issues = issues[:MAX_ISSUES]
            issues.append({"kind": "lint_truncated", "severity": "warning",
                           "message": f"问题过多，只保留前 {MAX_ISSUES} 条"})
            return issues
    return issues
