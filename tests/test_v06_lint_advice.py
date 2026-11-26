"""v0.6 #3：lint 升级为设计建议引擎（对标 Kition 的 designer 建议）。

新增三类检测：字号层级混乱（同页标题 < 正文）、对比度过低（文字 vs 形状
填充/页面背景，主题色文字也要解析出 hex 再比）、图片拉伸变形（非等比缩放，
裁剪过/占位小图不查）；所有 issue 携带「设计师会怎么改」的自然语言
suggestion。既有四类检查不变（fix_issues 语义不受影响，新类型不参与自修）。
"""
from __future__ import annotations

import io

import pymupdf
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Emu, Pt

from services import edits
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
    """带真标题占位符（layout 5 = Title Only）+ 正文的稿子，字号都显式。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    slide = prs.slides.add_slide(prs.slide_layouts[5])
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


def _white_bg_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
    return slide


def _picture_deck(tmp_path, px_w: int, px_h: int, emu_w: int, emu_h: int, crop=None):
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    pic = slide.shapes.add_picture(
        io.BytesIO(_png_bytes(px_w, px_h)), Emu(1000000), Emu(1000000), Emu(emu_w), Emu(emu_h)
    )
    if crop:
        pic.crop_left, pic.crop_right = crop
    path = tmp_path / "pic.pptx"
    prs.save(str(path))
    return path


# ---------- 字号层级混乱 ----------


class TestFontHierarchy:
    def test_title_smaller_than_body_flagged(self, tmp_path):
        issues = lint.lint_pptx(_title_body_deck(tmp_path, title_pt=14, body_pt=28))
        issue = _issue(issues, "font_hierarchy")
        assert issue is not None
        assert issue["severity"] == "warning"
        assert issue["fixable"] is False
        assert "14" in issue["message"] and "28" in issue["message"]
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
        slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(6000000), Emu(1200000)).text_frame.text = "没有标题的页面"
        slide.shapes.add_textbox(Emu(500000), Emu(2000000), Emu(6000000), Emu(1200000)).text_frame.text = "更大的正文"
        path = tmp_path / "notitle.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "font_hierarchy") is None


# ---------- 对比度过低 ----------


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
        slide = _white_bg_slide(prs)
        tb = slide.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(1000000))
        tb.text_frame.text = "白底白字（无形状填充）"
        for run in tb.text_frame.paragraphs[0].runs:
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor.from_string("F5F5F5")
        path = tmp_path / "contrast-bg.pptx"
        prs.save(str(path))
        issue = _issue(lint.lint_pptx(path), "low_contrast")
        assert issue is not None
        assert issue.get("contrast_ratio", 99) < 2.0

    def test_scheme_text_color_resolved(self, tmp_path):
        """主题色文字（如 bg1 引用）也要经 theme+clrMap 解析出 hex 再比对比度。"""
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = _white_bg_slide(prs)
        tb = slide.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(1000000))
        tb.text_frame.text = "主题色文字"
        for run in tb.text_frame.paragraphs[0].runs:
            run.font.size = Pt(24)
            run.font.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1  # 默认主题解析为白
        path = tmp_path / "contrast-scheme.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "low_contrast") is not None

    def test_no_explicit_bg_skips(self, tmp_path):
        # 无显式背景色（继承主题）时跳过，避免对图片背景误报
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(1000000))
        tb.text_frame.text = "背景未知不查"
        for run in tb.text_frame.paragraphs[0].runs:
            run.font.size = Pt(24)
            run.font.color.rgb = RGBColor.from_string("FFFFFF")
        path = tmp_path / "nobg.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "low_contrast") is None

    def test_no_explicit_color_skipped(self, tmp_path):
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(1000000)).text_frame.text = "继承主题色"
        path = tmp_path / "contrast-inherit.pptx"
        prs.save(str(path))
        assert _issue(lint.lint_pptx(path), "low_contrast") is None


# ---------- 图片拉伸变形 ----------


class TestImageDistorted:
    def test_stretched_image_flagged(self, tmp_path):
        # 原图 2:1（120x60），摆成 3:1 → 明显非等比拉伸
        issues = lint.lint_pptx(_picture_deck(tmp_path, 120, 60, 3000000, 1000000))
        issue = _issue(issues, "image_distorted")
        assert issue is not None
        assert issue["severity"] == "warning"
        assert issue.get("suggestion")

    def test_proportional_image_ok(self, tmp_path):
        issues = lint.lint_pptx(_picture_deck(tmp_path, 120, 60, 4000000, 2000000))
        assert _issue(issues, "image_distorted") is None

    def test_cropped_picture_skips(self, tmp_path):
        # 裁剪过的图片宽高比本来就该不同（裁剪语义拿不准），不按变形报
        issues = lint.lint_pptx(_picture_deck(tmp_path, 120, 120, 4000000, 2000000, crop=(0.25, 0.25)))
        assert _issue(issues, "image_distorted") is None

    def test_tiny_placeholder_image_skipped(self, tmp_path):
        # 1x1 占位图（conftest demo 同款）拉伸不报：像素太少无可见变形
        issues = lint.lint_pptx(_picture_deck(tmp_path, 1, 1, 3000000, 3000000))
        assert _issue(issues, "image_distorted") is None


# ---------- suggestion 全覆盖 + 回归 ----------


class TestSuggestions:
    def test_every_new_issue_carries_suggestion(self, tmp_path):
        # 溢出 + 低对比 + 变形三类混合稿，逐一断言 suggestion 非空
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
        path = tmp_path / "mixed.pptx"
        prs.save(str(path))
        issues = lint.lint_pptx(path)
        assert issues, "混合稿至少应报出一个问题"
        for issue in issues:
            assert isinstance(issue.get("suggestion"), str) and issue["suggestion"].strip()

    def test_legacy_kinds_carry_suggestion(self, tmp_path):
        # 一稿四问题（v0.5 同款构造：溢出/超界/越画布/重叠），旧检查也要带建议
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(3000000), Emu(1100000))
        tb.text_frame.text = "这是一段很长的正文，" * 12
        for run in tb.text_frame.paragraphs[0].runs:
            run.font.size = Pt(20)
        big = slide.shapes.add_textbox(Emu(500000), Emu(2000000), Emu(6000000), Emu(2200000))
        big.text_frame.text = "巨大标题"
        for run in big.text_frame.paragraphs[0].runs:
            run.font.size = Pt(128)
        slide.shapes.add_textbox(
            Emu(11000000), Emu(6400000), Emu(3000000), Emu(1500000)
        ).text_frame.text = "角落"
        slide.shapes.add_textbox(Emu(7000000), Emu(3000000), Emu(4000000), Emu(3000000)).text_frame.text = "甲"
        slide.shapes.add_textbox(Emu(7400000), Emu(3300000), Emu(4000000), Emu(3000000)).text_frame.text = "乙"
        path = tmp_path / "legacy.pptx"
        prs.save(str(path))

        issues = lint.lint_pptx(path)
        assert issues
        for issue in issues:
            assert issue.get("suggestion"), f"{issue['kind']} 缺 suggestion"

    def test_demo_deck_still_clean(self, demo_pptx):
        # v0.5 契约：demo 稿零问题（新检查不得误报占位小图/无填充文本框）
        assert lint.lint_pptx(demo_pptx) == []

    def test_new_kinds_not_fixable_by_fix_issues(self, tmp_path):
        # fix_issues 只修 FIXABLE_KINDS，新设计类问题不参与确定性自修
        deck = _title_body_deck(tmp_path, title_pt=14, body_pt=28)
        dst = tmp_path / "fixed.pptx"
        edits.apply_patches_to_file(deck, dst, [{"action": "fix_issues", "payload": {}}])
        assert _issue(lint.lint_pptx(dst), "font_hierarchy") is not None

    def test_design_check_constants(self):
        assert lint.DISTORTION_TOLERANCE > 0
        assert 1.0 < lint.CONTRAST_MIN_RATIO < 7.0


# ---------- 端点 ----------


class TestLintEndpoint:
    def test_endpoint_returns_designer_issues(self, client, tmp_path):
        data = _picture_deck(tmp_path, 120, 60, 4000000, 1000000).read_bytes()
        resp = client.post(
            "/api/docs/upload",
            files={"file": ("pic.pptx", data, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        )
        assert resp.status_code == 200
        doc_id = resp.json()["doc_id"]
        lint_resp = client.get(f"/api/docs/{doc_id}/lint")
        assert lint_resp.status_code == 200
        payload = lint_resp.json()
        kinds = {i["kind"] for i in payload["issues"]}
        assert "image_distorted" in kinds
        distorted = next(i for i in payload["issues"] if i["kind"] == "image_distorted")
        assert distorted.get("suggestion")
