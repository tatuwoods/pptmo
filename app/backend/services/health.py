"""v0.6 #10：交付健康报告 —— 文档级质量评分卡（对标 awesome 榜「rework 量」度量）。

`health_report(pptx_path)` 纯函数可单测；main.py 的 GET /api/docs/{id}/health
调它。评分从 100 起扣：lint 错误/警告、乱码符号残留（PUA/替换符，Wingdings 类
字体在无字体环境渲染变豆腐的前兆）、字体非统一度。扣分明细在 checks 里逐项
可解释——AI 改完一轮引用它告诉用户「还差多少分」，而不是只给一个黑盒数字。

判定全部确定性（同 v0.5 lint 哲学）：宁可少扣不误扣。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

from pptx import Presentation

from services import lint

# 乱码嫌疑字符：私用区（Wingdings/Symbol 等符号字体的代码位）+ 替换符
_GARBLED_RE = re.compile(r"[\uE000-\uF8FF\uFFFD]")
MAX_GARBLED_SAMPLES = 8

# 评分扣分上限：单项扣穿后不再累加，避免一页脏稿把 100 分打穿。
# 权重对标「rework 量」：乱码是交付即废级（每字符 8 分、单项最多打穿到 55），
# lint 错误是返工级（每个 3 分），字体混用是打磨级。
LINT_ERROR_CAP = 20
LINT_ERROR_PER_ISSUE = 3
LINT_WARNING_CAP = 10
GARBLED_CAP = 45
GARBLED_PER_CHAR = 8
FONT_CAP = 15

GRADE_CUTS = (("A", 90), ("B", 75), ("C", 60))


def _int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def _grade(score: int) -> str:
    for grade, floor in GRADE_CUTS:
        if score >= floor:
            return grade
    return "D"


def _collect_text(prs) -> List[str]:
    """全稿文本（含表格单元格、组合内、备注），用于乱码扫描。"""
    texts: List[str] = []

    def walk(shapes) -> None:
        for shape in shapes:
            if getattr(shape, "shape_type", None) == 6:
                try:
                    walk(shape.shapes)
                except Exception:
                    pass
                continue
            if getattr(shape, "has_table", False):
                try:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text:
                                texts.append(str(cell.text))
                except Exception:
                    pass
                continue
            if getattr(shape, "has_text_frame", False):
                try:
                    if shape.text:
                        texts.append(str(shape.text))
                except Exception:
                    pass

    for slide in prs.slides:
        walk(slide.shapes)
        if getattr(slide, "has_notes_slide", False):
            try:
                note = slide.notes_slide.notes_text_frame.text
                if note:
                    texts.append(str(note))
            except Exception:
                pass
    return texts


def _font_stats(prs) -> Dict[str, Any]:
    """显式设置的 run 字体名统计：count/usage/dominant/uniformity。

    继承字体不落 XML，统计不到的 run 不计入分母（与 lint 字号检查同一口径）；
    全稿没有显式字体时视为完全统一，不扣分。"""
    usage: Dict[str, int] = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    name = str(getattr(run.font, "name", "") or "").strip()
                    if name:
                        usage[name] = usage.get(name, 0) + 1
    total = sum(usage.values())
    if not usage:
        return {"count": 0, "usage": {}, "dominant": None, "uniformity": 1.0, "scored_runs": 0}
    dominant, dominant_runs = max(usage.items(), key=lambda kv: kv[1])
    return {
        "count": len(usage),
        "usage": usage,
        "dominant": dominant,
        "uniformity": round(dominant_runs / total, 4),
        "scored_runs": total,
    }


def health_report(pptx_path: Path | str) -> Dict[str, Any]:
    """文档质量评分卡：score/grade + stats/lint/garbled/fonts + checks 扣分明细。"""
    prs = Presentation(str(pptx_path))
    issues = lint.lint_pptx(pptx_path)

    texts = _collect_text(prs)
    garbled_hits: List[str] = []
    garbled_count = 0
    for text in texts:
        found = _GARBLED_RE.findall(text)
        if found:
            garbled_count += len(found)
            garbled_hits.extend(found[:MAX_GARBLED_SAMPLES])
    garbled_hits = garbled_hits[:MAX_GARBLED_SAMPLES]

    fonts = _font_stats(prs)

    by_kind: Dict[str, int] = {}
    for issue in issues:
        kind = str(issue.get("kind") or "other")
        by_kind[kind] = by_kind.get(kind, 0) + 1
    error_count = sum(
        1 for issue in issues if issue.get("severity") == "error"
    )
    warning_count = sum(
        1 for issue in issues if issue.get("severity") == "warning"
    )

    checks: List[Dict[str, Any]] = []
    deduction_lint_errors = min(LINT_ERROR_CAP, error_count * LINT_ERROR_PER_ISSUE)
    checks.append({
        "name": "lint_errors",
        "score": 100 - deduction_lint_errors,
        "deduction": deduction_lint_errors,
        "detail": f"{error_count} 个错误级 lint 问题（每个扣 {LINT_ERROR_PER_ISSUE} 分，封顶 {LINT_ERROR_CAP}）",
    })
    deduction_lint_warnings = min(LINT_WARNING_CAP, warning_count)
    checks.append({
        "name": "lint_warnings",
        "score": 100 - deduction_lint_warnings,
        "deduction": deduction_lint_warnings,
        "detail": f"{warning_count} 个警告级 lint 问题（每个扣 1 分，封顶 {LINT_WARNING_CAP}）",
    })
    deduction_garbled = min(GARBLED_CAP, garbled_count * GARBLED_PER_CHAR)
    checks.append({
        "name": "garbled",
        "score": 100 - deduction_garbled,
        "deduction": deduction_garbled,
        "detail": (
            f"{garbled_count} 个乱码嫌疑字符（私用区/替换符，每个扣 {GARBLED_PER_CHAR} 分，"
            f"封顶 {GARBLED_CAP}）"
        ),
    })
    deduction_fonts = min(FONT_CAP, int(round((1.0 - fonts["uniformity"]) * FONT_CAP)))
    checks.append({
        "name": "fonts",
        "score": 100 - deduction_fonts,
        "deduction": deduction_fonts,
        "detail": (
            f"{fonts['count']} 种字体（主导 {fonts['dominant'] or '-'}，"
            f"统一度 {fonts['uniformity']:.0%}）"
        ),
    })

    score = max(0, 100 - sum(int(check["deduction"]) for check in checks))
    stats = {
        "slides": len(prs.slides),
        "text_shapes": sum(1 for t in texts if t.strip()),
        "images": sum(
            1
            for slide in prs.slides
            for shape in slide.shapes
            if getattr(shape, "shape_type", None) == 13
        ),
        "tables": sum(
            1
            for slide in prs.slides
            for shape in slide.shapes
            if getattr(shape, "has_table", False)
        ),
        "charts": sum(
            1
            for slide in prs.slides
            for shape in slide.shapes
            if getattr(shape, "has_chart", False)
        ),
    }
    return {
        "score": score,
        "grade": _grade(score),
        "stats": stats,
        "lint": {"count": len(issues), "by_kind": by_kind, "issues": issues},
        "garbled": {"count": garbled_count, "samples": garbled_hits},
        "fonts": fonts,
        "checks": checks,
    }
