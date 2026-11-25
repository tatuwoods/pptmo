"""v0.5 #7：后置 lint（确定性检查）+ fix_issues 自修依据。

对标 hands-on-deck 的 render→look→fix 闭环：apply 后自动跑一遍，把
「文字溢出框外 / 字号超界 / 对象越画布 / 重叠」四类问题以结构化 issues 返回
给 AI；fix_issues（edits.py）按同一套判定做确定性修复（缩字号/收回画布），
重叠只告警——挪谁都不一定对，交给 AI 决策。

检查是几何启发式（无渲染回流），判定都留了容忍度：宁可少报不误报。
"""
from __future__ import annotations

import math
import re
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Pt

MIN_FONT_PT = 8.0
MAX_FONT_PT = 96.0
EMU_PER_PT = 12700
DEFAULT_FONT_PT = 18.0
LINE_HEIGHT_FACTOR = 1.22     # 行高 ≈ 字号 × 1.22（含行距余量）
FIT_TOLERANCE = 1.15          # 溢出判定容忍：估算误差 ≥15% 才报
OVERLAP_RATIO = 0.5           # 重叠面积 / 较小形状面积 ≥ 50% 才告警
MAX_ISSUES = 200

FIXABLE_KINDS = ("text_overflow", "font_size_out_of_range", "shape_off_canvas")

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
            issues.append({
                "kind": "shape_off_canvas",
                "slide_index": slide_index,
                "shape_id": shape_id,
                "shape_name": name,
                "severity": "error",
                "fixable": True,
                "message": f"对象「{name}」超出画布边界（{left},{top},{width}x{height}）",
            })
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
            issues.append({
                "kind": "shape_overlap",
                "slide_index": slide_index,
                "shape_id": id_a,
                "other_shape_id": id_b,
                "severity": "warning",
                "fixable": False,
                "message": f"第 {slide_index + 1} 页对象 {id_a} 与 {id_b} 重叠过半（自动挪动有风险，请人工确认）",
            })
            if len(issues) >= MAX_ISSUES:
                return issues
    return issues


def lint_pptx(pptx_path: Path | str) -> List[Dict[str, Any]]:
    """全稿检查：文字溢出 / 字号超界 / 越画布 / 重叠。上限 MAX_ISSUES 条。"""
    prs = Presentation(str(pptx_path))
    slide_w, slide_h = _int(prs.slide_width), _int(prs.slide_height)
    issues: List[Dict[str, Any]] = []
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
                    issues.append({
                        "kind": "font_size_out_of_range",
                        "slide_index": slide_index,
                        "shape_id": shape_id,
                        "shape_name": name,
                        "severity": "error",
                        "fixable": True,
                        "font_size_pt": size,
                        "message": f"「{name}」字号 {size:g}pt 超出 [{MIN_FONT_PT:g}, {MAX_FONT_PT:g}] 界",
                    })
            if not text_fits(shape, bbox):
                issues.append({
                    "kind": "text_overflow",
                    "slide_index": slide_index,
                    "shape_id": shape_id,
                    "shape_name": name,
                    "severity": "error",
                    "fixable": True,
                    "message": f"「{name}」文字装不进文本框（估算溢出 ≥{int((FIT_TOLERANCE - 1) * 100)}%）",
                })
        issues.extend(_slide_bounds_issues(slide, slide_w, slide_h, slide_index))
        if len(issues) >= MAX_ISSUES:
            issues = issues[:MAX_ISSUES]
            issues.append({"kind": "lint_truncated", "severity": "warning",
                           "message": f"问题过多，只保留前 {MAX_ISSUES} 条"})
            return issues
    return issues
