"""v0.4 A 组：编辑与 AI 能力（8 项）。

1 batch_slide_ops / 2 apply_layout / 3 insert_summary / 4 edit_chart
5 SmartArt 读取与编辑 / 6 edit_notes / 7 超链接 / 8 highlight_text
"""
from __future__ import annotations

import io
import re
import shutil
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Pt

from services import edits


# ---------- helpers ----------


@pytest.fixture()
def chart_pptx(tmp_path: Path) -> Path:
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    data = CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("营收", (10, 20))
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(1000000), Emu(1000000), Emu(8000000), Emu(4000000), data
    )
    tb = slide.shapes.add_textbox(Emu(100000), Emu(100000), Emu(4000000), Emu(500000))
    tb.text_frame.text = "季度营收"
    path = tmp_path / "chart.pptx"
    prs.save(str(path))
    return path


SMARTART_DATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    "<dgm:ptLst>"
    '<dgm:pt modelId="ns1:1" type="doc"/>'
    '<dgm:pt modelId="ns1:2" type="node"><dgm:t><a:bodyPr/><a:lstStyle/>'
    "<a:p><a:r><a:rPr lang=\"zh-CN\"/><a:t>要点一</a:t></a:r></a:p></dgm:t></dgm:pt>"
    '<dgm:pt modelId="ns1:3" type="node"><dgm:t><a:bodyPr/><a:lstStyle/>'
    "<a:p><a:r><a:rPr lang=\"zh-CN\"/><a:t>要点二</a:t></a:r></a:p></dgm:t></dgm:pt>"
    "</dgm:ptLst>"
    "<dgm:cxnLst>"
    '<dgm:cxn modelId="ns1:4" srcId="ns1:1" destId="ns1:2" type="parOf"/>'
    '<dgm:cxn modelId="ns1:5" srcId="ns1:1" destId="ns1:3" type="parOf"/>'
    "</dgm:cxnLst>"
    "<dgm:bg/><dgm:whole/>"
    "</dgm:dataModel>"
)

_FRAME_XML = (
    '<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<p:nvGraphicFramePr><p:cNvPr id="90" name="SmartArt 90"/>'
    '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    '<p:xfrm><a:off x="600000" y="600000"/><a:ext cx="9000000" cy="4000000"/></p:xfrm>'
    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">"
    '<dgm:relIds r:dm="{dm}" r:lo="{lo}" r:qs="{qs}" r:cs="{cs}"/>'
    "</a:graphicData></a:graphic></p:graphicFrame>"
)

_MINIMAL_PARTS = {
    "lo": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/default"/>'
    ),
    "qs": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"/>'
    ),
    "cs": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" '
        'uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"/>'
    ),
}


def inject_smartart(src: Path, dst: Path, slide_part: str = "ppt/slides/slide1.xml") -> None:
    """Zip surgery: add a SmartArt graphicFrame + diagram parts into a slide."""
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        slide_xml = zin.read(slide_part).decode("utf-8")
        rels_name = f"ppt/slides/_rels/{Path(slide_part).name}.rels"
        rels_xml = zin.read(rels_name).decode("utf-8") if rels_name in names else (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )
        rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
        base = max(rids, default=0) + 1
        dm_rid, lo_rid, qs_rid, cs_rid = (f"rId{base + i}" for i in range(4))
        rels_xml = rels_xml.replace(
            "</Relationships>",
            "".join(
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                f'officeDocument/2006/relationships/{kind}" Target="../diagrams/{name}"/>'
                for rid, kind, name in (
                    (dm_rid, "diagramData", "data1.xml"),
                    (lo_rid, "diagramLayout", "layout1.xml"),
                    (qs_rid, "diagramQuickStyle", "quickStyle1.xml"),
                    (cs_rid, "diagramColors", "colors1.xml"),
                )
            )
            + "</Relationships>",
        )
        frame = _FRAME_XML.format(dm=dm_rid, lo=lo_rid, qs=qs_rid, cs=cs_rid)
        slide_xml = slide_xml.replace("</p:spTree>", frame + "</p:spTree>")

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == slide_part:
                    zout.writestr(item, slide_xml)
                elif item.filename == rels_name:
                    zout.writestr(item, rels_xml)
                elif item.filename == "[Content_Types].xml":
                    ct = item and zin.read(item.filename).decode("utf-8")
                    ct = ct.replace(
                        "</Types>",
                        '<Override PartName="/ppt/diagrams/data1.xml" ContentType='
                        '"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>'
                        '<Override PartName="/ppt/diagrams/layout1.xml" ContentType='
                        '"application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>'
                        '<Override PartName="/ppt/diagrams/quickStyle1.xml" ContentType='
                        '"application/vnd.openxmlformats-officedocument.drawingml.diagramQuickStyle+xml"/>'
                        '<Override PartName="/ppt/diagrams/colors1.xml" ContentType='
                        '"application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>'
                        "</Types>",
                    )
                    zout.writestr(item, ct)
                else:
                    zout.writestr(item, zin.read(item.filename))
            zout.writestr("ppt/diagrams/data1.xml", SMARTART_DATA)
            for key, xml in _MINIMAL_PARTS.items():
                zout.writestr(f"ppt/diagrams/{ {'lo': 'layout1', 'qs': 'quickStyle1', 'cs': 'colors1'}[key] }.xml", xml)


def _apply(src: Path, dst: Path, patches: list) -> set:
    return edits.apply_patches_to_file(src, dst, patches)


# ---------- 1. batch_slide_ops ----------


def test_batch_duplicate_pages_range_and_list(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    touched = _apply(demo_pptx, out, [
        {"action": "batch_slide_ops", "op": "duplicate_slide", "pages": "1-2"},
    ])
    prs = Presentation(str(out))
    assert len(prs.slides) == 5  # 3 + 2 clones
    assert touched  # touched pages reported


def test_batch_delete_pages_list(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [
        {"action": "batch_slide_ops", "op": "delete_slide", "pages": [1, 2]},
    ])
    prs = Presentation(str(out))
    assert len(prs.slides) == 1


def test_batch_pages_syntax_variants(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [
        {"action": "batch_slide_ops", "op": "duplicate_slide", "pages": "1, 3"},
    ])
    assert len(Presentation(str(out)).slides) == 5


def test_batch_rejects_bad_op_or_out_of_range(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, out, [{"action": "batch_slide_ops", "op": "explode", "pages": [0]}])
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, out, [{"action": "batch_slide_ops", "op": "delete_slide", "pages": [9]}])
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, out, [{"action": "batch_slide_ops", "op": "delete_slide", "pages": "x-y"}])


def test_batch_move_shifts_each_page(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [
        {"action": "batch_slide_ops", "op": "move_slide", "pages": [1], "shift": 1},
    ])
    prs = Presentation(str(out))
    # slide 0 moved one position later: the table slide is now first
    assert any(getattr(sh, "has_table", False) for sh in prs.slides[0].shapes)


# ---------- 2. apply_layout ----------


def test_apply_layout_title_preset(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [{"action": "apply_layout", "slide_index": 0, "layout": "title"}])
    prs = Presentation(str(out))
    slide = prs.slides[0]
    shapes = sorted(
        (sh for sh in slide.shapes if getattr(sh, "has_text_frame", False) and sh.text.strip()),
        key=lambda sh: -int(sh.height or 0),
    )
    title = shapes[0]
    # title box horizontally centered on the slide
    center = int(title.left) + int(title.width) / 2
    assert abs(center - int(prs.slide_width) / 2) < int(prs.slide_width) * 0.05


def test_apply_layout_two_column_splits_bodies(tmp_path):
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide.shapes.add_textbox(Emu(100000), Emu(100000), Emu(4000000), Emu(800000))
    title.text_frame.text = "标题"
    for run in title.text_frame.paragraphs[0].runs:
        run.font.size = Pt(36)
    left = slide.shapes.add_textbox(Emu(100000), Emu(2000000), Emu(3000000), Emu(800000))
    left.text_frame.text = "左栏内容"
    right = slide.shapes.add_textbox(Emu(100000), Emu(3000000), Emu(3000000), Emu(800000))
    right.text_frame.text = "右栏内容"
    src = tmp_path / "src.pptx"
    prs.save(str(src))

    out = tmp_path / "out.pptx"
    _apply(src, out, [{"action": "apply_layout", "slide_index": 0, "layout": "two_column"}])
    prs2 = Presentation(str(out))
    boxes = {
        sh.text_frame.text: (int(sh.left), int(sh.width))
        for sh in prs2.slides[0].shapes
        if getattr(sh, "has_text_frame", False) and sh.text.strip()
    }
    lx, lw = boxes["左栏内容"]
    rx, rw = boxes["右栏内容"]
    slide_w = 12192000
    assert lx + lw < slide_w * 0.5 < rx  # left column ends before midline, right starts after
    assert boxes["标题"][0] < slide_w * 0.1  # title spans from the left edge band


def test_apply_layout_unknown_preset_rejected(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, tmp_path / "out.pptx", [
            {"action": "apply_layout", "slide_index": 0, "layout": "portal"},
        ])


# ---------- 3. insert_summary ----------


def test_insert_summary_creates_toc_slide(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [{
        "action": "insert_summary", "position": 1,
        "title": "目录", "items": ["项目进展汇报", "数据表格", "分组说明"],
    }])
    prs = Presentation(str(out))
    assert len(prs.slides) == 4
    texts = []
    for sh in prs.slides[1].shapes:
        if getattr(sh, "has_text_frame", False):
            texts.append(sh.text_frame.text)
    joined = "\n".join(texts)
    assert "目录" in joined
    for item in ("项目进展汇报", "数据表格", "分组说明"):
        assert item in joined


def test_insert_summary_requires_title_and_items(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, tmp_path / "o.pptx", [{"action": "insert_summary", "position": 0}])


# ---------- 4. edit_chart ----------


def test_edit_chart_categories_series_and_name(chart_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(chart_pptx, out, [{
        "action": "edit_chart", "shape_id": "s0_id2",
        "categories": ["一季度", "二季度"],
        "series": [{"name": "营收（万元）", "values": [12, 24]}],
    }])
    prs = Presentation(str(out))
    chart = next(sh for sh in prs.slides[0].shapes if getattr(sh, "has_chart", False)).chart
    assert list(chart.plots[0].categories) == ["一季度", "二季度"]
    series = chart.series[0]
    assert series.name == "营收（万元）"
    assert list(series.values) == [12.0, 24.0]


def test_edit_chart_partial_update_keeps_other_values(chart_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(chart_pptx, out, [{
        "action": "edit_chart", "shape_id": "s0_id2", "series": [{"values": [7, 8]}],
    }])
    prs = Presentation(str(out))
    chart = next(sh for sh in prs.slides[0].shapes if getattr(sh, "has_chart", False)).chart
    assert chart.series[0].name == "营收"  # untouched
    assert list(chart.plots[0].categories) == ["Q1", "Q2"]
    assert list(chart.series[0].values) == [7.0, 8.0]


def test_edit_chart_shape_must_be_chart(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, tmp_path / "o.pptx", [
            {"action": "edit_chart", "shape_id": "s0_id2", "series": [{"values": [1]}]},
        ])


# ---------- 5. SmartArt ----------


@pytest.fixture()
def smartart_pptx(demo_pptx, tmp_path):
    dst = tmp_path / "smart.pptx"
    inject_smartart(demo_pptx, dst)
    return dst


def test_smartart_text_read_into_objects(smartart_pptx):
    from services.parse import parse_pptx

    _objs, payload = parse_pptx(smartart_pptx)
    smart = [o for o in payload["objects"] if o["kind"] == "smartart_node"]
    assert {o["text"] for o in smart} == {"要点一", "要点二"}
    assert all(o["shape_id"].startswith("s0_id90_n") for o in smart)


def test_smartart_node_edit(smartart_pptx, tmp_path):
    from services.parse import parse_pptx

    out = tmp_path / "out.pptx"
    _apply(smartart_pptx, out, [
        {"action": "edit", "shape_id": "s0_id90_n1", "after": "改写后的第二点"},
    ])
    _objs, payload = parse_pptx(out)
    smart = {o["text"] for o in payload["objects"] if o["kind"] == "smartart_node"}
    assert smart == {"要点一", "改写后的第二点"}


def test_smartart_node_edit_validates_before(smartart_pptx, tmp_path):
    with pytest.raises(edits.PatchConflict):
        _apply(smartart_pptx, tmp_path / "o.pptx", [
            {"action": "edit", "shape_id": "s0_id90_n0", "before": "完全无关的原文", "after": "x"},
        ])


def test_smartart_node_out_of_range(smartart_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        _apply(smartart_pptx, tmp_path / "o.pptx", [
            {"action": "edit", "shape_id": "s0_id90_n9", "after": "x"},
        ])


# ---------- 6. edit_notes ----------


def test_edit_notes_single_and_clear(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [{"action": "edit_notes", "slide_index": 1, "text": "本页讲表格"}])
    prs = Presentation(str(out))
    assert prs.slides[1].notes_slide.notes_text_frame.text == "本页讲表格"
    assert prs.slides[0].notes_slide.notes_text_frame.text == "讲稿：先回顾里程碑。"  # untouched

    out2 = tmp_path / "out2.pptx"
    _apply(out, out2, [{"action": "edit_notes", "slide_index": 1, "text": ""}])
    prs2 = Presentation(str(out2))
    assert prs2.slides[1].notes_slide.notes_text_frame.text == ""


def test_edit_notes_batch(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [{
        "action": "edit_notes",
        "notes": [
            {"slide_index": 0, "text": "备注零"},
            {"slide_index": 1, "text": "备注一"},
            {"slide_index": 2, "text": "备注二"},
        ],
    }])
    prs = Presentation(str(out))
    assert [prs.slides[i].notes_slide.notes_text_frame.text for i in range(3)] == [
        "备注零", "备注一", "备注二",
    ]


def test_edit_notes_out_of_range(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, tmp_path / "o.pptx", [{"action": "edit_notes", "slide_index": 5, "text": "x"}])


def test_parse_payload_exposes_notes(demo_pptx):
    from services.parse import parse_pptx

    _objs, payload = parse_pptx(demo_pptx)
    notes = [s.get("notes") or "" for s in payload["slides"]]
    assert notes[0] == "讲稿：先回顾里程碑。"
    assert notes[1] == ""


# ---------- 7. set/remove hyperlink ----------


def test_set_hyperlink_whole_shape(demo_pptx, tmp_path):
    out = tmp_path / "out.pptx"
    _apply(demo_pptx, out, [
        {"action": "set_hyperlink", "shape_id": "s0_id2", "url": "https://example.com/report"},
    ])
    prs = Presentation(str(out))
    shape = next(sh for sh in prs.slides[0].shapes if sh.shape_id == 2)
    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.hyperlink.address == "https://example.com/report"


def test_set_hyperlink_find_scopes_runs(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    p = tb.text_frame.paragraphs[0]
    for piece in ("访问", "官网", "了解更多"):
        r = p.add_run()
        r.text = piece
    src = tmp_path / "s.pptx"
    prs.save(str(src))
    out = tmp_path / "o.pptx"
    _apply(src, out, [{"action": "set_hyperlink", "shape_id": "s0_id2", "url": "https://a.b", "find": "官网"}])
    prs2 = Presentation(str(out))
    runs = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs
    links = [bool(r.hyperlink.address) for r in runs]
    assert links == [False, True, False]


def test_set_hyperlink_validates_url(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, tmp_path / "o.pptx", [
            {"action": "set_hyperlink", "shape_id": "s0_id2", "url": "javascript:alert(1)"},
        ])
    with pytest.raises(edits.PatchError):
        _apply(demo_pptx, tmp_path / "o.pptx", [{"action": "set_hyperlink", "shape_id": "s0_id2"}])


def test_remove_hyperlink(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    tb.text_frame.text = "链接文字"
    src = tmp_path / "s.pptx"
    prs.save(str(src))
    out = tmp_path / "o.pptx"
    _apply(src, out, [{"action": "set_hyperlink", "shape_id": "s0_id2", "url": "https://x.y"}])
    out2 = tmp_path / "o2.pptx"
    _apply(out, out2, [{"action": "remove_hyperlink", "shape_id": "s0_id2"}])
    prs2 = Presentation(str(out2))
    run = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert not run.hyperlink.address


# ---------- 8. highlight_text ----------


def _highlight_runs(prs, slide_index=0):
    from pptx.oxml.ns import qn

    out = []
    for shape in prs.slides[slide_index].shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                rpr = r._r.find(qn("a:rPr"))
                hl = rpr.find(qn("a:highlight")) if rpr is not None else None
                if hl is not None:
                    color = hl.find(qn("a:srgbClr"))
                    out.append((r.text, color.get("val") if color is not None else None))
    return out


def test_highlight_text_default_yellow(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    tb.text_frame.text = "本季度完成三个关键里程碑"
    src = tmp_path / "s.pptx"
    prs.save(str(src))

    out = tmp_path / "o.pptx"
    _apply(src, out, [{"action": "highlight_text", "find": "里程碑"}])
    hits = _highlight_runs(Presentation(str(out)))
    assert hits == [("里程碑", "FFFF00")]  # 精确高亮命中片段，周边文字保持无高亮


def test_highlight_text_custom_color_and_scope(tmp_path):
    prs = Presentation()
    s0 = prs.slides.add_slide(prs.slide_layouts[6])
    tb0 = s0.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    tb0.text_frame.text = "页面一包含关键词"
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb1 = s1.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    tb1.text_frame.text = "页面二没有词"
    src = tmp_path / "s.pptx"
    prs.save(str(src))

    out = tmp_path / "o.pptx"
    _apply(src, out, [{"action": "highlight_text", "find": "关键词", "color": "FFC000", "slide_index": 0}])
    prs2 = Presentation(str(out))
    assert _highlight_runs(prs2, 0) == [("关键词", "FFC000")]
    assert _highlight_runs(prs2, 1) == []


def test_highlight_text_no_match_rejected(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    tb.text_frame.text = "无关内容"
    src = tmp_path / "s.pptx"
    prs.save(str(src))
    with pytest.raises(edits.PatchError):
        _apply(src, tmp_path / "o.pptx", [{"action": "highlight_text", "find": "不存在"}])


def test_highlight_text_bad_color_rejected(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(5000000), Emu(500000))
    tb.text_frame.text = "内容"
    src = tmp_path / "s.pptx"
    prs.save(str(src))
    with pytest.raises(edits.PatchError):
        _apply(src, tmp_path / "o.pptx", [{"action": "highlight_text", "find": "内容", "color": "ZZZ"}])
