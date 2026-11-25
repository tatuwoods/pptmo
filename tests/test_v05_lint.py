"""v0.5 #7：lint 后置检查 + fix_issues 确定性自修。

对标 hands-on-deck fix：apply 后自动 lint（文字溢出框外/字号超界/对象越画布/
重叠告警）返回 issues；fix_issues 缩字号/收回画布确定性修复；重叠只告警不修。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
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


def _overflow_deck(tmp_path):
    """小框塞长文 + 超大字号 + 越画布形状 + 大面积重叠，一稿四问题。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    # 1) 文字溢出：小框塞长文（规模控制在缩字号可修复的范围内）
    tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(3000000), Emu(1100000))
    tf = tb.text_frame
    tf.text = "这是一段很长的正文，" * 12
    for run in tf.paragraphs[0].runs:
        run.font.size = Pt(20)
    # 2) 字号超界：128pt
    big = slide.shapes.add_textbox(Emu(500000), Emu(2000000), Emu(6000000), Emu(2200000))
    big.text_frame.text = "巨大标题"
    for run in big.text_frame.paragraphs[0].runs:
        run.font.size = Pt(128)
    # 3) 越画布：右下探出画布
    slide.shapes.add_textbox(
        Emu(11000000), Emu(6400000), Emu(3000000), Emu(1500000)
    ).text_frame.text = "角落"
    # 4) 重叠：两个大框叠 80%
    slide.shapes.add_textbox(Emu(7000000), Emu(3000000), Emu(4000000), Emu(3000000)).text_frame.text = "甲"
    slide.shapes.add_textbox(Emu(7400000), Emu(3300000), Emu(4000000), Emu(3000000)).text_frame.text = "乙"
    path = tmp_path / "lint.pptx"
    prs.save(str(path))
    return path


class TestLintUnit:
    def test_detects_all_four_kinds(self, tmp_path):
        issues = lint.lint_pptx(_overflow_deck(tmp_path))
        kinds = {i["kind"] for i in issues}
        assert "text_overflow" in kinds
        assert "font_size_out_of_range" in kinds
        assert "shape_off_canvas" in kinds
        assert "shape_overlap" in kinds

    def test_issue_fields(self, tmp_path):
        issues = lint.lint_pptx(_overflow_deck(tmp_path))
        for issue in issues:
            assert isinstance(issue["slide_index"], int)
            assert issue.get("shape_id")
            assert issue.get("message")
            assert issue.get("severity") in ("warning", "error")
        # 溢出/超界/越画布可修，重叠只告警
        assert _issue(issues, "text_overflow")["fixable"] is True
        assert _issue(issues, "font_size_out_of_range")["fixable"] is True
        assert _issue(issues, "shape_off_canvas")["fixable"] is True
        assert _issue(issues, "shape_overlap")["fixable"] is False

    def test_clean_deck_has_no_issues(self, demo_pptx):
        assert lint.lint_pptx(demo_pptx) == []

    def test_font_bounds(self):
        assert lint.MIN_FONT_PT == 8 and lint.MAX_FONT_PT == 96


class TestFixUnit:
    def test_fix_shrinks_overflow_font(self, tmp_path):
        deck = _overflow_deck(tmp_path)
        dst = tmp_path / "fixed.pptx"
        touched = edits.apply_patches_to_file(
            deck, dst, [{"action": "fix_issues", "payload": {}}]
        )
        assert 0 in touched
        issues = lint.lint_pptx(dst)
        assert _issue(issues, "text_overflow") is None
        assert _issue(issues, "font_size_out_of_range") is None
        # 修完字号回到界内
        prs = Presentation(str(dst))
        sizes = [
            r.font.size.pt
            for s in prs.slides[0].shapes
            for p in s.text_frame.paragraphs
            for r in p.runs
            if r.font.size is not None
        ]
        assert all(lint.MIN_FONT_PT <= size <= lint.MAX_FONT_PT for size in sizes)

    def test_fix_pulls_shape_back_into_canvas(self, tmp_path):
        deck = _overflow_deck(tmp_path)
        dst = tmp_path / "fixed.pptx"
        edits.apply_patches_to_file(deck, dst, [{"action": "fix_issues", "payload": {}}])
        assert _issue(lint.lint_pptx(dst), "shape_off_canvas") is None
        prs = Presentation(str(dst))
        w, h = int(prs.slide_width), int(prs.slide_height)
        for shape in prs.slides[0].shapes:
            assert shape.left >= 0 and shape.top >= 0
            assert shape.left + shape.width <= w
            assert shape.top + shape.height <= h

    def test_fix_leaves_overlap_alone(self, tmp_path):
        deck = _overflow_deck(tmp_path)
        dst = tmp_path / "fixed.pptx"
        edits.apply_patches_to_file(deck, dst, [{"action": "fix_issues", "payload": {}}])
        # 重叠只告警：修复后仍应在（且不再有其他问题）
        assert _issue(lint.lint_pptx(dst), "shape_overlap") is not None

    def test_fix_scoped_to_kinds(self, tmp_path):
        deck = _overflow_deck(tmp_path)
        dst = tmp_path / "fixed.pptx"
        edits.apply_patches_to_file(
            deck, dst,
            [{"action": "fix_issues", "payload": {"kinds": ["shape_off_canvas"]}}],
        )
        issues = lint.lint_pptx(dst)
        assert _issue(issues, "shape_off_canvas") is None
        assert _issue(issues, "text_overflow") is not None  # 未要求修的保留

    def test_alias_fix(self, tmp_path):
        deck = _overflow_deck(tmp_path)
        dst = tmp_path / "fixed.pptx"
        edits.apply_patches_to_file(deck, dst, [{"action": "fix", "payload": {}}])
        assert _issue(lint.lint_pptx(dst), "text_overflow") is None


class TestLintApi:
    def test_apply_response_carries_lint_issues(self, client, uploaded_doc, tmp_path):
        # 上传一个带问题的稿子（复用 overflow fixture 走完整链路）
        import io as _io

        buf = _io.BytesIO()
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        blank = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank)
        big = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(3000000), Emu(600000))
        big.text_frame.text = "超界字号测试"
        for run in big.text_frame.paragraphs[0].runs:
            run.font.size = Pt(128)
        prs.save(buf)
        resp = client.post(
            "/api/docs/upload",
            files={"file": ("bad.pptx", buf.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
        )
        assert resp.status_code == 200, resp.text
        doc_id = resp.json()["doc_id"]
        # 无害编辑触发 apply，响应带 lint_issues
        apply_resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"label": "改字", "patches": [
                {"action": "fix_issues", "payload": {}}
            ]},
        )
        assert apply_resp.status_code == 200, apply_resp.text
        # fix_issues 修完字号问题，lint_issues 里不再有 font_size_out_of_range
        body = apply_resp.json()
        assert "lint_issues" in body
        assert _issue(body["lint_issues"], "font_size_out_of_range") is None

    def test_lint_endpoint(self, client, uploaded_doc):
        resp = client.get(f"/api/docs/{uploaded_doc['doc_id']}/lint")
        assert resp.status_code == 200
        body = resp.json()
        assert body["issues"] == []
        assert body["count"] == 0
