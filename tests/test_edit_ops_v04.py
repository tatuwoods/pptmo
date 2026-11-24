"""v0.4-A: 编辑与 AI 能力 — batch_slide_ops / apply_layout / generate_summary /
edit_chart / edit_smartart / edit_notes / set_hyperlink / highlight_text（edits 层）。"""
from __future__ import annotations

import copy
import io
import shutil
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Emu, Pt

from services import edits, parse


PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)


@pytest.fixture()
def multi_deck(tmp_path) -> Path:
    """5 页文稿：每页一个标题文本框（shape id=2，spTree 自身占 id=1）。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for index in range(5):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Emu(1000000), Emu(500000), Emu(6000000), Emu(900000))
        tb.text_frame.text = f"第{index + 1}页标题"
    path = tmp_path / "multi.pptx"
    prs.save(str(path))
    return path


@pytest.fixture()
def chart_deck(tmp_path) -> Path:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    data = CategoryChartData()
    data.categories = ["一月", "二月", "三月"]
    data.add_series("销量", (10.0, 20.0, 30.0))
    data.add_series("成本", (5.0, 8.0, 12.0))
    slide.shapes.add_chart(1, Emu(1000000), Emu(1000000), Emu(6000000), Emu(4000000), data)
    path = tmp_path / "chart.pptx"
    prs.save(str(path))
    return path


# ---------- SmartArt 测试夹具：zip 级注入最小 SmartArt ----------

SMARTART_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DATA_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<dgm:dataModel xmlns:dgm="{SMARTART_NS}" xmlns:a="{A_NS}"><dgm:ptLst>'
    '<dgm:pt modelId="1"><dgm:prSet phldr="1"/><dgm:spPr/><dgm:t><a:bodyPr/>'
    '<a:p><a:r><a:rPr lang="zh-CN"/><a:t>开始</a:t></a:r></a:p></dgm:t></dgm:pt>'
    '<dgm:pt modelId="2"><dgm:spPr/><dgm:t><a:bodyPr/>'
    '<a:p><a:r><a:rPr lang="zh-CN"/><a:t>过程审核</a:t></a:r></a:p></dgm:t></dgm:pt>'
    '</dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>'
)
DRAWING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" '
    f'xmlns:a="{A_NS}"><dsp:spTree>'
    '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>'
    '<dsp:grpSpPr/>'
    '<dsp:sp modelId="1"><dsp:nvSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>'
    '<dsp:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="600000"/></a:xfrm>'
    '<a:prstGeom prst="rect"/></dsp:spPr>'
    '<dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>开始</a:t></a:r></a:p></dsp:txBody>'
    '</dsp:sp>'
    '</dsp:spTree></dsp:drawing>'
)
LAYOUT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<dgm:layout xmlns:dgm="{SMARTART_NS}" xmlns:a="{A_NS}"><dgm:title val=""/>'
    '<dgm:layoutNode name="0"><dgm:var type="node"/></dgm:layoutNode></dgm:layout>'
)
QUICKSTYLE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<dgm:quickStyle xmlns:dgm="{SMARTART_NS}" xmlns:a="{A_NS}"><dgm:title val=""/></dgm:quickStyle>'
)
COLORS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<dgm:colors xmlns:dgm="{SMARTART_NS}" xmlns:a="{A_NS}"><a:schemeClr val="accent1"/></dgm:colors>'
)
GRAPHIC_FRAME_XML = (
    '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    f'xmlns:a="{A_NS}" xmlns:r="{R_NS}">'
    '<p:nvGraphicFramePr><p:cNvPr id="99" name="SmartArt 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    '<p:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    f'<dgm:relIds xmlns:dgm="{SMARTART_NS}" xmlns:r="{R_NS}" '
    'r:dm="rId90" r:lo="rId91" r:qs="rId92" r:cs="rId93"/>'
    '</a:graphicData></a:graphic></p:graphicFrame>'
)


def _build_smartart_deck(src: Path, dst: Path) -> None:
    """在 src 的 slide1 里注入一个最小 SmartArt（数据模型 + 缓存 drawing）。"""
    with zipfile.ZipFile(src) as zin:
        items = {name: zin.read(name) for name in zin.namelist()}
    slide_name = "ppt/slides/slide1.xml"
    slide = items[slide_name].decode("utf-8")
    slide = slide.replace("</p:spTree>", GRAPHIC_FRAME_XML + "</p:spTree>")
    items[slide_name] = slide.encode("utf-8")

    rels_name = "ppt/slides/_rels/slide1.xml.rels"
    rels = items[rels_name].decode("utf-8")
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    add = "".join(
        f'<Relationship Id="{rid}" Type="{rel_type}/{rtype}" Target="{target}"/>'
        for rid, rtype, target in [
            ("rId90", "diagramData", "../diagrams/data1.xml"),
            ("rId91", "diagramDrawing", "../diagrams/drawing1.xml"),
            ("rId92", "diagramLayout", "../diagrams/layout1.xml"),
            ("rId93", "diagramQuickStyle", "../diagrams/quickStyle1.xml"),
            ("rId94", "diagramColors", "../diagrams/colors1.xml"),
        ]
    )
    items[rels_name] = rels.replace("</Relationships>", add + "</Relationships>").encode("utf-8")

    items["ppt/diagrams/data1.xml"] = DATA_XML.encode("utf-8")
    items["ppt/diagrams/drawing1.xml"] = DRAWING_XML.encode("utf-8")
    items["ppt/diagrams/layout1.xml"] = LAYOUT_XML.encode("utf-8")
    items["ppt/diagrams/quickStyle1.xml"] = QUICKSTYLE_XML.encode("utf-8")
    items["ppt/diagrams/colors1.xml"] = COLORS_XML.encode("utf-8")

    content_types = items["[Content_Types].xml"].decode("utf-8")
    ct_type = "application/vnd.openxmlformats-officedocument.drawingml"
    overrides = "".join(
        f'<Override PartName="/ppt/diagrams/{name}.xml" '
        f'ContentType="{ct_type}.{ctype}+xml"/>'
        for name, ctype in [
            ("data1", "diagramData"),
            ("drawing1", "diagramDrawing"),
            ("layout1", "diagramLayoutDefinition"),
            ("quickStyle1", "diagramQuickStyleDefinition"),
            ("colors1", "diagramColors"),
        ]
    )
    items["[Content_Types].xml"] = content_types.replace(
        "</Types>", overrides + "</Types>"
    ).encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)


@pytest.fixture()
def smartart_deck(tmp_path) -> Path:
    src = tmp_path / "plain.pptx"
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(str(src))
    dst = tmp_path / "smartart.pptx"
    _build_smartart_deck(src, dst)
    return dst


# ---------- 1. batch_slide_ops ----------

def test_batch_duplicate_range_expands(multi_deck, tmp_path):
    out = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        multi_deck,
        out,
        [{"action": "batch_slide_ops", "pages": "2-4",
          "op": {"action": "duplicate_slide", "source_slide": "$page"}}],
    )
    assert touched
    prs = Presentation(str(out))
    assert len(prs.slides) == 8  # 5 + 3 duplicates


def test_batch_page_list_and_mixed_spec(multi_deck, tmp_path):
    out = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "batch_slide_ops", "pages": "1,3", "op": {"action": "duplicate_slide", "source_slide": "$page"}}],
    )
    prs = Presentation(str(out))
    assert len(prs.slides) == 7


def test_batch_page_spec_out_of_range(multi_deck, tmp_path):
    out = tmp_path / "out.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "batch_slide_ops", "pages": "2-9", "op": {"action": "duplicate_slide", "source_slide": "$page"}}],
        )


def test_batch_delete_slides(multi_deck, tmp_path):
    out = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "batch_slide_ops", "pages": "1,5", "op": {"action": "delete_slide", "slide": "$page"}}],
    )
    prs = Presentation(str(out))
    assert len(prs.slides) == 3


def test_batch_requires_template(multi_deck, tmp_path):
    out = tmp_path / "out.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "batch_slide_ops", "pages": "1-2"}],
        )


# ---------- 2. apply_layout ----------

def test_apply_layout_title_preset(multi_deck):
    out = multi_deck.with_name("layout.pptx")
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "apply_layout", "slide_index": 0, "layout": "title",
          "title": "季度汇报", "subtitle": "2026 Q4"}],
    )
    prs = Presentation(str(out))
    slide = prs.slides[0]
    texts = [s.text for s in slide.shapes if getattr(s, "has_text_frame", False)]
    assert "季度汇报" in texts and "2026 Q4" in texts


def test_apply_layout_two_col(multi_deck):
    out = multi_deck.with_name("layout2.pptx")
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "apply_layout", "slide_index": 1, "layout": "two_col",
          "title": "对比", "left": "方案 A", "right": "方案 B"}],
    )
    prs = Presentation(str(out))
    texts = [s.text for s in prs.slides[1].shapes if getattr(s, "has_text_frame", False)]
    assert "方案 A" in texts and "方案 B" in texts


def test_apply_layout_unknown_name_rejected(multi_deck, tmp_path):
    out = tmp_path / "bad.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "apply_layout", "slide_index": 0, "layout": "nonexistent"}],
        )


def test_apply_layout_image_text_with_image(multi_deck):
    out = multi_deck.with_name("layout3.pptx")
    import base64

    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "apply_layout", "slide_index": 2, "layout": "image_text",
          "title": "图文页", "text": "右侧说明文字",
          "image_base64": base64.b64encode(PNG).decode()}],
    )
    prs = Presentation(str(out))
    kinds = [str(getattr(s, "shape_type", "")) for s in prs.slides[2].shapes]
    assert any("PICTURE" in k for k in kinds)


def test_apply_layout_data_tiles(multi_deck):
    out = multi_deck.with_name("layout4.pptx")
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "apply_layout", "slide_index": 3, "layout": "data",
          "title": "核心指标",
          "data": [["收入", "1.2亿"], ["用户", "45万"], ["留存", "38%"], ["NPS", "62"]]}],
    )
    prs = Presentation(str(out))
    texts = [s.text for s in prs.slides[3].shapes if getattr(s, "has_text_frame", False)]
    assert "收入" in texts and "62" in texts


# ---------- 3. generate_summary ----------

def test_generate_summary_inserts_toc(multi_deck, tmp_path):
    out = tmp_path / "toc.pptx"
    touched = edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "generate_summary", "position": 1, "mode": "toc"}],
    )
    prs = Presentation(str(out))
    assert len(prs.slides) == 6
    toc_texts = [s.text for s in prs.slides[1].shapes if getattr(s, "has_text_frame", False)]
    joined = "\n".join(toc_texts)
    assert "第1页标题" in joined and "第5页标题" in joined


def test_generate_summary_custom_items(multi_deck, tmp_path):
    out = tmp_path / "toc2.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "generate_summary", "position": 0, "mode": "summary",
          "title": "本章要点", "items": ["要点一", "要点二"]}],
    )
    prs = Presentation(str(out))
    texts = [s.text for s in prs.slides[0].shapes if getattr(s, "has_text_frame", False)]
    assert "要点一" in "\n".join(texts)


def test_generate_summary_out_of_range_position(multi_deck, tmp_path):
    out = tmp_path / "toc3.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "generate_summary", "position": 9}],
        )


# ---------- 4. edit_chart ----------

def test_edit_chart_replace_series_values(chart_deck, tmp_path):
    out = tmp_path / "chart_out.pptx"
    edits.apply_patches_to_file(
        chart_deck, out,
        [{"action": "edit_chart", "shape_id": "s0_id2",
          "payload": {"categories": ["Q1", "Q2", "Q3"],
                      "series": [{"name": "销量", "values": [11, 22, 33]}]}}],
    )
    prs = Presentation(str(out))
    chart = next(s for s in prs.slides[0].shapes if getattr(s, "has_chart", False)).chart
    assert list(chart.plots[0].categories) == ["Q1", "Q2", "Q3"]
    assert list(chart.plots[0].series[0].values) == [11.0, 22.0, 33.0]


def test_edit_chart_partial_keeps_other_series(chart_deck, tmp_path):
    out = tmp_path / "chart_out2.pptx"
    edits.apply_patches_to_file(
        chart_deck, out,
        [{"action": "edit_chart", "shape_id": "s0_id2",
          "payload": {"series": [{"name": "成本", "values": [6, 9, 15]}]}}],
    )
    prs = Presentation(str(out))
    chart = next(s for s in prs.slides[0].shapes if getattr(s, "has_chart", False)).chart
    assert list(chart.plots[0].series[0].values) == [10.0, 20.0, 30.0]  # 销量保留
    assert list(chart.plots[0].series[1].values) == [6.0, 9.0, 15.0]


def test_edit_chart_requires_chart_shape(multi_deck, tmp_path):
    out = tmp_path / "nochart.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "edit_chart", "shape_id": "s0_id2",
              "payload": {"series": [{"name": "x", "values": [1]}]}}],
        )


# ---------- 5. SmartArt ----------

def test_smartart_parse_detected(smartart_deck):
    objects, payload = parse.parse_pptx(str(smartart_deck))
    arts = [e for e in payload["elements"] if e.get("kind") == "smartart"]
    assert len(arts) == 1
    obj = next(o for o in objects if o.get("kind") == "smartart")
    assert "开始" in obj["text"] and "过程审核" in obj["text"]


def test_edit_smartart_replaces_text(smartart_deck, tmp_path):
    out = tmp_path / "smartart_out.pptx"
    edits.apply_patches_to_file(
        smartart_deck, out,
        [{"action": "edit_smartart", "shape_id": "s0_id99",
          "find": "开始", "replace": "启动"}],
    )
    with zipfile.ZipFile(out) as zf:
        data_xml = zf.read("ppt/diagrams/data1.xml").decode("utf-8")
        drawing_xml = zf.read("ppt/diagrams/drawing1.xml").decode("utf-8")
    assert "启动" in data_xml and "开始" not in data_xml
    assert "启动" in drawing_xml


def test_edit_smartart_missing_text_rejected(smartart_deck, tmp_path):
    out = tmp_path / "smartart_bad.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            smartart_deck, out,
            [{"action": "edit_smartart", "shape_id": "s0_id99",
              "find": "不存在", "replace": "x"}],
        )


# ---------- 6. edit_notes ----------

def test_edit_notes_set_one_page(multi_deck, tmp_path):
    out = tmp_path / "notes.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "edit_notes", "slide_index": 2, "text": "这页讲交付节奏"}],
    )
    prs = Presentation(str(out))
    assert prs.slides[2].notes_slide.notes_text_frame.text == "这页讲交付节奏"


def test_edit_notes_batch_items(multi_deck, tmp_path):
    out = tmp_path / "notes2.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "edit_notes", "items": [
            {"slide_index": 0, "text": "开场"},
            {"slide_index": 1, "text": "背景"},
        ]}],
    )
    prs = Presentation(str(out))
    assert prs.slides[0].notes_slide.notes_text_frame.text == "开场"
    assert prs.slides[1].notes_slide.notes_text_frame.text == "背景"


def test_edit_notes_out_of_range(multi_deck, tmp_path):
    out = tmp_path / "notes3.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "edit_notes", "slide_index": 7, "text": "x"}],
        )


def test_parse_includes_notes(multi_deck):
    prs_path = multi_deck.with_name("with_notes.pptx")
    shutil.copy(multi_deck, prs_path)
    prs = Presentation(str(prs_path))
    prs.slides[1].notes_slide.notes_text_frame.text = "备注内容"
    prs.save(str(prs_path))
    _, payload = parse.parse_pptx(str(prs_path))
    assert payload["slides"][1]["notes"] == "备注内容"


# ---------- 7. set_hyperlink / remove_hyperlink ----------

def test_set_hyperlink_whole_shape(multi_deck, tmp_path):
    out = tmp_path / "link.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "set_hyperlink", "shape_id": "s0_id2", "url": "https://example.com"}],
    )
    prs = Presentation(str(out))
    runs = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert all(r.hyperlink.address == "https://example.com" for r in runs)


def test_set_hyperlink_matching_runs_only(multi_deck, tmp_path):
    prs_path = multi_deck.with_name("mixed.pptx")
    shutil.copy(multi_deck, prs_path)
    prs = Presentation(str(prs_path))
    tf = prs.slides[0].shapes[0].text_frame
    tf.paragraphs[0].add_run().text = " 详情见官网"
    prs.save(str(prs_path))
    out = prs_path.with_name("link2.pptx")
    edits.apply_patches_to_file(
        prs_path, out,
        [{"action": "set_hyperlink", "shape_id": "s0_id2",
          "find": "第1页标题", "url": "https://example.com/x"}],
    )
    prs = Presentation(str(out))
    runs = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    linked = [r.hyperlink.address for r in runs]
    assert "https://example.com/x" in linked
    assert None in linked  # 未匹配的 run 不加链接


def test_set_hyperlink_invalid_url(multi_deck, tmp_path):
    out = tmp_path / "link3.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "set_hyperlink", "shape_id": "s0_id2", "url": "javascript:alert(1)"}],
        )


def test_remove_hyperlink(multi_deck, tmp_path):
    linked = multi_deck.with_name("linked.pptx")
    edits.apply_patches_to_file(
        multi_deck, linked,
        [{"action": "set_hyperlink", "shape_id": "s0_id2", "url": "https://example.com"}],
    )
    out = linked.with_name("unlinked.pptx")
    edits.apply_patches_to_file(
        linked, out,
        [{"action": "remove_hyperlink", "shape_id": "s0_id2"}],
    )
    prs = Presentation(str(out))
    runs = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    assert all(r.hyperlink.address is None for r in runs)


# ---------- 8. highlight_text ----------

def test_highlight_whole_shape(multi_deck, tmp_path):
    out = tmp_path / "hl.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "highlight_text", "shape_id": "s0_id2"}],
    )
    prs = Presentation(str(out))
    run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    hl = run._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr").find(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}highlight"
    )
    assert hl is not None
    assert hl.find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").get("val") == "FFFF00"


def test_highlight_matching_substring_splits_run(multi_deck, tmp_path):
    out = tmp_path / "hl2.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "highlight_text", "slide_index": 0, "find": "标题"}],
    )
    prs = Presentation(str(out))
    runs = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs
    a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    highlighted_texts = [
        r.text for r in runs
        if r._r.find(f"{a_ns}rPr") is not None
        and r._r.find(f"{a_ns}rPr").find(f"{a_ns}highlight") is not None
    ]
    assert highlighted_texts == ["标题"]  # 精确切分，只高亮匹配片段
    assert "".join(r.text for r in runs) == "第1页标题"


def test_highlight_find_no_match_rejected(multi_deck, tmp_path):
    out = tmp_path / "hl3.pptx"
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            multi_deck, out,
            [{"action": "highlight_text", "find": "不存在的词"}],
        )


def test_highlight_custom_color(multi_deck, tmp_path):
    out = tmp_path / "hl4.pptx"
    edits.apply_patches_to_file(
        multi_deck, out,
        [{"action": "highlight_text", "shape_id": "s0_id2", "color": "00FF00"}],
    )
    prs = Presentation(str(out))
    run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    hl = run._r.find(f"{a_ns}rPr").find(f"{a_ns}highlight")
    assert hl.find(f"{a_ns}srgbClr").get("val") == "00FF00"
