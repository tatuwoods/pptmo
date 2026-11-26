"""v0.6 #10：交付健康报告 —— 文档级质量评分卡。

对标 awesome-codex-ppt-skills 的「rework 量」度量：lint 问题数、乱码符号残留、
字体统一度、页数/对象统计，输出 0-100 分 + A/B/C/D 档。AI 每轮改完可引用，
用户随时看「这份稿子还差多少分」。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from services import health

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)


def _clean_deck(tmp_path):
    """干净稿：统一字体、无乱码、无 lint 问题。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    tb = slide.shapes.add_textbox(Emu(1000000), Emu(500000), Emu(6000000), Emu(1000000))
    tb.text_frame.text = "项目进展汇报"
    for run in tb.text_frame.paragraphs[0].runs:
        run.font.name = "Microsoft YaHei"
        run.font.size = Pt(32)
    body = slide.shapes.add_textbox(Emu(1000000), Emu(2000000), Emu(8000000), Emu(1500000))
    body.text_frame.text = "本季度完成了三个关键里程碑。"
    for run in body.text_frame.paragraphs[0].runs:
        run.font.name = "Microsoft YaHei"
        run.font.size = Pt(18)
    path = tmp_path / "clean.pptx"
    prs.save(str(path))
    return path


def _dirty_deck(tmp_path):
    """脏稿：乱码 PUA 字符 + 三种字体混用 + 溢出 + 越画布。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    # 1) 乱码：私用区符号（Wingdings 转字前的典型残留）+ 替换符
    tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(3000000), Emu(800000))
    tb.text_frame.text = "要点\uE012\uE013 与损坏字符\uFFFD"
    for run in tb.text_frame.paragraphs[0].runs:
        run.font.name = "Wingdings"
        run.font.size = Pt(18)
    # 2) 溢出：小框塞长文
    over = slide.shapes.add_textbox(Emu(500000), Emu(1600000), Emu(2400000), Emu(700000))
    over.text_frame.text = "这段正文特别长，塞进小文本框必然溢出，" * 10
    for run in over.text_frame.paragraphs[0].runs:
        run.font.name = "SimSun"
        run.font.size = Pt(16)
    # 3) 字号超界 + 第三种字体
    big = slide.shapes.add_textbox(Emu(500000), Emu(2600000), Emu(6000000), Emu(1200000))
    big.text_frame.text = "巨大标题"
    for run in big.text_frame.paragraphs[0].runs:
        run.font.name = "Arial Black"
        run.font.size = Pt(120)
    # 4) 越画布
    slide.shapes.add_textbox(
        Emu(11500000), Emu(6400000), Emu(2000000), Emu(1200000)
    ).text_frame.text = "角落"
    path = tmp_path / "dirty.pptx"
    prs.save(str(path))
    return path


class TestHealthUnit:
    def test_clean_deck_scores_high(self, tmp_path):
        report = health.health_report(_clean_deck(tmp_path))
        assert report["score"] >= 90
        assert report["grade"] == "A"
        assert report["garbled"]["count"] == 0
        assert report["fonts"]["count"] == 1
        assert report["fonts"]["uniformity"] == pytest.approx(1.0)
        assert report["lint"]["count"] == 0
        assert report["stats"]["slides"] == 1
        assert report["stats"]["text_shapes"] >= 2

    def test_dirty_deck_flags_everything(self, tmp_path):
        report = health.health_report(_dirty_deck(tmp_path))
        assert report["score"] < 60
        assert report["grade"] in ("C", "D")
        assert report["garbled"]["count"] >= 2  # 两个 PUA + 1 个替换符
        assert report["fonts"]["count"] == 3
        assert report["fonts"]["uniformity"] < 0.6
        kinds = set(report["lint"]["by_kind"])
        assert "text_overflow" in kinds
        assert "shape_off_canvas" in kinds
        assert "font_size_out_of_range" in kinds
        # 扣分明细可解释：每项检查带扣分与说明
        for check in report["checks"]:
            assert check["name"] and isinstance(check["score"], int)
        assert any(c["name"] == "garbled" and c["deduction"] > 0 for c in report["checks"])
        assert any(c["name"] == "fonts" and c["deduction"] > 0 for c in report["checks"])

    def test_report_shape(self, tmp_path):
        report = health.health_report(_clean_deck(tmp_path))
        assert 0 <= report["score"] <= 100
        assert report["grade"] in ("A", "B", "C", "D")
        assert set(report) >= {"score", "grade", "stats", "lint", "garbled", "fonts", "checks"}
        assert report["stats"]["images"] == 0
        assert report["stats"]["tables"] == 0
        assert report["stats"]["charts"] == 0

    def test_score_never_negative(self, tmp_path):
        # 两页都脏，扣分项远超 100 也不应为负
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        blank = prs.slide_layouts[6]
        for _ in range(2):
            slide = prs.slides.add_slide(blank)
            tb = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(1000000), Emu(400000))
            tb.text_frame.text = "\uE000\uE001\uE002\uE003\uE004\uE005" * 5
        path = tmp_path / "awful.pptx"
        prs.save(str(path))
        report = health.health_report(path)
        assert 0 <= report["score"] <= 100
        assert report["grade"] == "D"


class TestHealthAPI:
    def test_endpoint_report(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.get(f"/api/docs/{doc_id}/health")
        assert resp.status_code == 200, resp.text
        data = resp.json()
        assert data["doc_id"] == doc_id
        assert data["version_id"] == "v0"
        assert 0 <= data["score"] <= 100
        assert data["grade"] in ("A", "B", "C", "D")
        assert data["stats"]["slides"] == uploaded_doc["slide_count"]

    def test_endpoint_unknown_doc(self, client):
        # 格式合法但不存在（格式非法走全局 400，与其它端点一致）
        assert client.get("/api/docs/deadbeefdeadbeefdeadbeefdeadbeef/health").status_code == 404
