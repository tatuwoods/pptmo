"""v0.3 验收补测（A10 font_scale）：继承来的字号必须按“实际生效值”起算。

v0.3.0 的 font_scale 对没有显式 sz 的 run 一律按 18pt 起算。但占位符里的标题/正文
通常不带 sz，字号继承自 版式 → 母版占位符 → 母版 txStyles；在 samples/demo.pptx 上
整份 +2 磅会把 109 个 run 全部压成 20pt（52pt 的标题、9pt 的页脚都不例外）。

期望值来源：母版 txStyles 的 sz（python-pptx 默认模板 44/32/28）以及 LibreOffice 实际
渲染 samples/demo.pptx 后从 PDF 抽出的 span 字号（52/28/20/22/16/18/9）。
"""
from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from services import edits

SAMPLE = Path(__file__).resolve().parent.parent / "app" / "samples" / "demo.pptx"


def _size_pt(run):
    return run.font.size.pt if run.font.size is not None else None


def _runs_of(slide):
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    yield shape, paragraph, run


def _shape_by_name(prs, slide_index, name):
    return next(s for s in prs.slides[slide_index].shapes if s.name == name)


def _first_run_size(prs, slide_index, name):
    return _size_pt(_shape_by_name(prs, slide_index, name).text_frame.paragraphs[0].runs[0])


@pytest.fixture()
def placeholder_deck(tmp_path):
    """python-pptx 默认模板：标题 44pt、副标题/一级正文 32pt、二级正文 28pt，全部继承、run 上无 sz。"""
    prs = Presentation()
    cover = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide
    cover.shapes.title.text_frame.text = "封面标题"
    cover.placeholders[1].text_frame.text = "副标题"
    content = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
    content.shapes.title.text_frame.text = "内容页标题"
    body = content.placeholders[1]
    body.text_frame.paragraphs[0].text = "一级要点"
    second = body.text_frame.add_paragraph()
    second.text = "二级要点"
    second.level = 1
    path = tmp_path / "placeholders.pptx"
    prs.save(str(path))
    return path


def test_placeholder_runs_really_have_no_explicit_size(placeholder_deck):
    """前提：这份夹具的字号确实是继承的，否则下面的断言测不到继承解析。"""
    prs = Presentation(str(placeholder_deck))
    sizes = [_size_pt(run) for _s, _p, run in _runs_of(prs.slides[0])] + [
        _size_pt(run) for _s, _p, run in _runs_of(prs.slides[1])
    ]
    assert sizes and all(size is None for size in sizes)


def test_font_scale_starts_from_layout_and_master_inherited_size(placeholder_deck, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(placeholder_deck, dst, [{"action": "font_scale", "delta_pt": 2}])
    prs = Presentation(str(dst))
    sizes = {(index, run.text): _size_pt(run) for index in (0, 1) for _s, _p, run in _runs_of(prs.slides[index])}
    assert sizes == {
        (0, "封面标题"): 46,  # master titleStyle 44 + 2
        (0, "副标题"): 34,  # master bodyStyle lvl1 32 + 2
        (1, "内容页标题"): 46,
        (1, "一级要点"): 34,
        (1, "二级要点"): 30,  # master bodyStyle lvl2 28 + 2
    }


def test_font_scale_negative_delta_keeps_inherited_hierarchy(placeholder_deck, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(placeholder_deck, dst, [{"action": "font_scale", "delta_pt": -4}])
    prs = Presentation(str(dst))
    assert _size_pt(next(_runs_of(prs.slides[0]))[2]) == 40  # 44 - 4
    second = [run for _s, p, run in _runs_of(prs.slides[1]) if p.level == 1]
    assert [ _size_pt(run) for run in second ] == [24]  # 28 - 4


def test_font_scale_sample_deck_shifts_each_inherited_size_by_delta(tmp_path):
    """samples/demo.pptx（冒烟用样张）上没有一个 run 带显式字号；+2 后每处都只大 2 磅。"""
    if not SAMPLE.exists():
        pytest.skip("sample deck not available")
    dst = tmp_path / "demo.pptx"
    edits.apply_patches_to_file(SAMPLE, dst, [{"action": "font_scale", "delta_pt": 2}])
    prs = Presentation(str(dst))
    # 期望 = LibreOffice 渲染原稿的实际字号 + 2
    assert _first_run_size(prs, 1, "标题 1") == 54  # master titleStyle 52
    assert _first_run_size(prs, 4, "标题 1") == 54
    assert _first_run_size(prs, 1, "内容占位符 2") == 30  # layout lstStyle 28
    assert _first_run_size(prs, 2, "内容占位符 2") == 22  # layout lstStyle 20
    assert _first_run_size(prs, 5, "文本占位符 3") == 24  # layout lstStyle 22
    assert _first_run_size(prs, 6, "文本占位符 15") == 18  # layout lstStyle 16
    assert _first_run_size(prs, 1, "页脚占位符 16") == 11  # master footer placeholder 9
    # 页码是 a:fld 域而不是 run：要和旁边的页脚一起变大，否则一行里两种字号
    number_field = _shape_by_name(prs, 1, "灯片编号占位符 17")._element.find(".//" + qn("a:fld"))
    assert number_field.get("type") == "slidenum"
    assert number_field.find(qn("a:rPr")).get("sz") == "1100"
    table = _shape_by_name(prs, 4, "表 4").table
    cell_run = table.cell(0, 1).text_frame.paragraphs[0].runs[0]
    assert _size_pt(cell_run) == 20  # 表格默认 18
    # 原来那种“一律 20pt”的压平结果不允许再出现
    all_sizes = {_size_pt(run) for slide in prs.slides for _s, _p, run in _runs_of(slide)}
    assert all_sizes != {20.0}


def test_font_scale_keeps_slide_scope(placeholder_deck, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        placeholder_deck, dst, [{"action": "font_scale", "delta_pt": 2, "slide_index": 0}]
    )
    prs = Presentation(str(dst))
    assert _size_pt(next(_runs_of(prs.slides[0]))[2]) == 46
    assert all(_size_pt(run) is None for _s, _p, run in _runs_of(prs.slides[1]))


def test_font_scale_moves_explicit_paragraph_defaults_with_the_text(tmp_path):
    """段落上显式写的字号（defRPr / endParaRPr）跟文字一起挪，未显式写的不凭空发明。"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(6000000), Emu(3000000))
    first = box.text_frame.paragraphs[0]
    run = first.add_run()
    run.text = "显式默认"
    p_pr = first._p.get_or_add_pPr()
    def_rpr = p_pr.makeelement(qn("a:defRPr"), {"sz": "2000"})
    p_pr.append(def_rpr)
    first._p.append(first._p.makeelement(qn("a:endParaRPr"), {"sz": "2000"}))
    second = box.text_frame.add_paragraph()
    second.add_run().text = "无显式默认"
    second._p.append(second._p.makeelement(qn("a:endParaRPr"), {"lang": "zh-CN"}))  # 继承，不写 sz
    src = tmp_path / "defaults.pptx"
    prs.save(str(src))
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(src, dst, [{"action": "font_scale", "delta_pt": 2}])
    prs2 = Presentation(str(dst))
    paras = prs2.slides[0].shapes[0].text_frame.paragraphs
    assert _size_pt(paras[0].runs[0]) == 22  # defRPr 20 + 2
    assert paras[0]._p.find(qn("a:pPr")).find(qn("a:defRPr")).get("sz") == "2200"
    assert paras[0]._p.find(qn("a:endParaRPr")).get("sz") == "2200"
    assert _size_pt(paras[1].runs[0]) == 20  # 什么都没声明 → 默认 18 + 2
    assert paras[1]._p.find(qn("a:endParaRPr")).get("sz") is None


def test_font_scale_stays_clamped_at_bounds(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(6000000), Emu(1000000))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "极小字号"
    run.font.size = Pt(7)
    src = tmp_path / "tiny.pptx"
    prs.save(str(src))
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(src, dst, [{"action": "font_scale", "delta_pt": -4}])
    assert _size_pt(Presentation(str(dst)).slides[0].shapes[0].text_frame.paragraphs[0].runs[0]) == 6
