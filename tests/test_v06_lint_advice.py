"""v0.6 #3：lint 升级为设计建议引擎（对标 Kition 的 designer 建议）。

新增三类检测：字号层级混乱（同页标题 < 正文）、对比度过低（文字 vs 形状
填充/页面背景）、图片拉伸变形（非等比缩放）；所有 issue 携带「设计师会怎么
改」的自然语言 suggestion。既有四类检查不变（fix_issues 语义不受影响）。
"""
from __future__ import annotations

import io

import pymupdf
import pytest
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

from services import lint

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)


def _issue(issues, kind, slide_index=0):
    return next(
        (i for i in issues if i["kind"] == kind and i["slide_index"] == slide_index),
        None,
    )


def _png_bytes(width: int, height: int) -> bytes:
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, width, height))
    pix.set_rect(pix.irect, (120, 140, 160))
    return pix.tobytes("png")


def _title_body_deck(tmp_path, title_pt: float, body_pt: float):
    """带真标题占位符（layout 5 = Title Only 有标题位）+ 正文的稿子。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    layout = prs.slide_layouts[5]  # Title Only
    slide = prs.slides.add_slide(layout)
    title = slide.shapes.title
    title.text = "季度汇报"
    for run in title.text_frame.paragraphs[0].runs:
        run.font.size = Pt(title_pt)
    body = slide.shapes.add_textbox(Emu(1000000), Emu(2500000), Emu(8000000), Emu(1500000))
    body.text_frame.text = "正文内容若干"
    for run in body.text_frame.paragraphs[0].runs:
        run.font.size = Pt(body_pt)
    path = tmp_path / "hierarchy.pptx"
    prs.save(str(path))
    return path


def _filled_textbox(slide, text: str, fill_rgb: str, font_rgb: str, x=Emu(1000000), y=Emu(1000000)):
    tb = slide.shapes.add_textbox(x, y, Emu(6000000), Emu(1000000))
    tb.text_frame.text = text
    for run in tb.text_frame.paragraphs[0].runs:
        run.font.size = Pt(20)
        run.font.color.rgb = RGBColor.from_string(font_rgb)
    tb.fill.solid()
    tb.fill.fore_color.rgb = RGBColor.from_string(fill_rgb)
    return tb


class TestFontHierarchy:
    def test_title_smaller_than_body_flagged(self, tmp_path):
        issues = lint.lint_pptx(_title_body_deck(tmp_path, title_pt=18, body_pt=32))
        issue = _issue(issues, "font_hierarchy")
        assert issue is not None
        assert issue["severity"] == "warning"
        assert issue.get("suggestion")

    def test_title_larger_than_body_ok(self, tmp_path):
        assert _issue(lint.lint_pptx(_title_body_deck(tmp_path, 40, 18)), "font_hierarchy") is None

    def test_no_explicit_sizes_not_flagged(self, tmp_path):
        # 标题/正文都没写死字号（继承母版）→ 不猜，不报
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "只有默认字号"
        slide.shapes.add_textbox(Emu(1000000), Emu(2500000), Emu(8000000), Emu(1500000)).text_frame.text = "正文"
        path = tmp_path / "inherit.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "font_hierarchy") is None

    def test_no_title_placeholder_skips(self, tmp_path):
        # 没有标题占位符的页面不做层级比较（宁缺毋滥）
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb1 = slide.shapes.add_textbox(Emu(1000000), Emu(500000), Emu(6000000), Emu(1000000))
        tb1.text_frame.text = "没有标题占位符"
        for run in tb1.text_frame.paragraphs[0].runs:
            run.font.size = Pt(18)
        tb2 = slide.shapes.add_textbox(Emu(1000000), Emu(2500000), Emu(8000000), Emu(1500000))
        tb2.text_frame.text = "更大的正文"
        for run in tb2.text_frame.paragraphs[0].runs:
            run.font.size = Pt(36)
        path = tmp_path / "notitle.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "font_hierarchy") is None


class TestLowContrast:
    def test_dark_text_on_dark_fill(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _filled_textbox(slide, "深色底深色字", "1A1A1A", "333333")
        path = tmp_path / "contrast.pptx"
        prs.save(str(path))
        issue = _issue(lint.lint_pptx(path), "low_contrast")
        assert issue is not None
        assert issue["severity"] == "warning"
        assert issue.get("suggestion")

    def test_dark_text_on_white_fill_ok(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _filled_textbox(slide, "白底黑字", "FFFFFF", "1A1A1A")
        path = tmp_path / "contrast-ok.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "low_contrast") is None

    def test_white_text_on_white_slide_background(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
        tb = slide.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(1000000))
        tb.text_frame.text = "白底白字（无形状填充）"
        for run in tb.text_frame.paragraphs[0].runs:
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor.from_string("F5F5F5")
        path = tmp_path / "contrast-bg.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "low_contrast") is not None

    def test_no_explicit_color_skipped(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(1000000)).text_frame.text = "继承主题色"
        path = tmp_path / "contrast-inherit.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "low_contrast") is None


class TestImageDistorted:
    def test_stretched_image_flagged(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 原图 2:1（120x60），摆成 1:1 → 明显非等比拉伸
        slide.shapes.add_picture(
            io.BytesIO(_png_bytes(120, 60)), Emu(1000000), Emu(1000000), Emu(3000000), Emu(3000000)
        )
        path = tmp_path / "stretch.pptx"
        prs.save(str(path))
        issue = _issue(lint.lint_pptx(path), "image_distorted")
        assert issue is not None
        assert issue["severity"] == "warning"
        assert issue.get("suggestion")

    def test_proportional_image_ok(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(
            io.BytesIO(_png_bytes(120, 60)), Emu(1000000), Emu(1000000), Emu(4000000), Emu(2000000)
        )
        path = tmp_path / "proportional.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "image_distorted") is None

    def test_cropped_picture_skips(self, tmp_path):
        # 裁剪过的图片宽高比本来就该不同，不按变形报
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        pic = slide.shapes.add_picture(
            io.BytesIO(_png_bytes(120, 60)), Emu(1000000), Emu(1000000), Emu(4000000), Emu(2000000)
        )
        pic.crop_left = 0.25
        pic.crop_right = 0.25
        path = tmp_path / "cropped.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "image_distorted") is None

    def test_tiny_placeholder_image_skipped(self, tmp_path):
        # 1x1 占位图（conftest demo 同款）拉伸不报：像素太少无可见变形
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tiny = bytes.fromhex(
            "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
            "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
        )
        slide.shapes.add_picture(io.BytesIO(tiny), Emu(1000000), Emu(1000000), Emu(3000000), Emu(3000000))
        path = tmp_path / "tiny.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "image_distorted") is None


class TestSuggestions:
    def test_every_issue_carries_suggestion(self, tmp_path):
        from pathlib import Path

        # 复用 v0.5 的四问题稿 + 本版三类新问题稿，逐一断言 suggestion 非空
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(3000000), Emu(1100000))
        tb.text_frame.text = "这是一段很长的正文，" * 12
        for run in tb.text_frame.paragraphs[0].runs:
            run.font.size = Pt(20)
        _filled_textbox(slide, "低对比文本", "111111", "222222", y=Emu(2200000))
        slide.shapes.add_picture(
            io.BytesIO(_png_bytes(120, 60)), Emu(1000000), Emu(3600000), Emu(3000000), Emu(3000000)
        )
        path = Path(tmp_path / "mixed.pptx")
        prs.save(str(path))
        issues = lint.lint_pptx(path)
        assert issues, "混合稿至少应报出一个问题"
        for issue in issues:
            assert isinstance(issue.get("suggestion"), str) and issue["suggestion"].strip()

    def test_demo_deck_still_clean(self, demo_pptx):
        # v0.5 契约：demo 稿零问题（新检查不得误报占位小图/无填充文本框）
        assert lint.lint_pptx(demo_pptx) == []
