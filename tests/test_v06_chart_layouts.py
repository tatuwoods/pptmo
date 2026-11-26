"""v0.6 #6 原生图表创建 + #7 母版版式强化。

#6：add_chart 操作——python-pptx 原生 chart（column/bar/bar_h/line/pie/area/doughnut
+ 中文别名），数据来自 AI 整理的 categories/series。真图表：PowerPoint 里可双击改数据，
序列化后 parse 出 kind="chart" 对象，后续 edit_chart 可继续改数。

#7：读取文档自带 slide layouts（含多母版）列表给 AI；add_slide 缺省不再硬编码
空白版式而是用全稿最常用版式（品牌一致），支持 layout_name 按名套用。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu

from services import edits

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)

CHART_PATCH = {
    "action": "add_chart",
    "slide_index": 0,
    "chart_type": "column",
    "categories": ["Q1", "Q2", "Q3"],
    "series": [{"name": "营收", "values": [12, 15, 18]}],
    "title": "季度营收（万元）",
}


def _deck(path, pages: int = 2, layout_index: int = 6) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    for _ in range(pages):
        prs.slides.add_slide(prs.slide_layouts[layout_index])
    prs.save(str(path))
    return path


def _chart_of(slide):
    frames = [sh for sh in slide.shapes if getattr(sh, "has_chart", False)]
    assert frames, "页面上应有一个图表 graphicFrame"
    return frames[0].chart


# ---------- #6 add_chart：服务层 ----------


class TestAddChartService:
    def test_column_chart_created_with_data(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        touched = edits.apply_patches_to_file(src, dst, [CHART_PATCH])
        assert touched == {0}
        chart = _chart_of(Presentation(str(dst)).slides[0])
        assert chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED
        assert [str(c) for c in chart.plots[0].categories] == ["Q1", "Q2", "Q3"]
        series = list(chart.series)
        assert len(series) == 1 and series[0].name == "营收"
        assert [float(v) for v in series[0].values] == [12.0, 15.0, 18.0]
        assert chart.has_title and chart.chart_title.text_frame.text == "季度营收（万元）"

    @pytest.mark.parametrize(
        "alias, expected",
        [
            ("bar", XL_CHART_TYPE.COLUMN_CLUSTERED),
            ("柱状图", XL_CHART_TYPE.COLUMN_CLUSTERED),
            ("柱形图", XL_CHART_TYPE.COLUMN_CLUSTERED),
            ("bar_h", XL_CHART_TYPE.BAR_CLUSTERED),
            ("条形图", XL_CHART_TYPE.BAR_CLUSTERED),
            ("line", XL_CHART_TYPE.LINE_MARKERS),
            ("折线图", XL_CHART_TYPE.LINE_MARKERS),
            ("pie", XL_CHART_TYPE.PIE),
            ("饼图", XL_CHART_TYPE.PIE),
            ("area", XL_CHART_TYPE.AREA),
            ("面积图", XL_CHART_TYPE.AREA),
            ("doughnut", XL_CHART_TYPE.DOUGHNUT),
            ("环形图", XL_CHART_TYPE.DOUGHNUT),
        ],
    )
    def test_type_aliases(self, tmp_path, alias, expected):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(
            src, dst, [{**CHART_PATCH, "chart_type": alias}]
        )
        assert _chart_of(Presentation(str(dst)).slides[0]).chart_type == expected

    def test_unknown_type_rejected(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        with pytest.raises(edits.PatchError, match="chart_type"):
            edits.apply_patches_to_file(
                src, dst, [{**CHART_PATCH, "chart_type": "三维炫彩"}]
            )

    def test_bbox_and_multi_series(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        bbox = {"x": 914400, "y": 1828800, "w": 7315200, "h": 3657600}
        edits.apply_patches_to_file(
            src, dst,
            [{
                **CHART_PATCH,
                "chart_type": "line",
                "series": [
                    {"name": "营收", "values": [1, 2, 3]},
                    {"name": "成本", "values": [3, 2, 1]},
                ],
                "bbox_emu": bbox,
            }],
        )
        prs = Presentation(str(dst))
        frame = [sh for sh in prs.slides[0].shapes if getattr(sh, "has_chart", False)][0]
        assert (frame.left, frame.top, frame.width, frame.height) == (
            Emu(bbox["x"]), Emu(bbox["y"]), Emu(bbox["w"]), Emu(bbox["h"])
        )
        assert len(list(_chart_of(prs.slides[0]).series)) == 2

    def test_bad_series_rejected(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        with pytest.raises(edits.PatchError, match="series"):
            edits.apply_patches_to_file(src, dst, [{**CHART_PATCH, "series": []}])
        with pytest.raises(edits.PatchError, match="数值"):
            edits.apply_patches_to_file(
                src, dst,
                [{**CHART_PATCH, "series": [{"name": "x", "values": ["a", "b"]}]}],
            )
        with pytest.raises(edits.PatchError, match="slide_index"):
            edits.apply_patches_to_file(src, dst, [{**CHART_PATCH, "slide_index": 9}])

    def test_kind_alias_key(self, tmp_path):
        """chart_type 也接受 kind 键（AI 词汇容错）。"""
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(
            src, dst,
            [{"action": "add_chart", "slide_index": 0, "kind": "pie",
              "categories": ["甲", "乙"], "series": [{"name": "s", "values": [1, 2]}]}],
        )
        assert _chart_of(Presentation(str(dst)).slides[0]).chart_type == XL_CHART_TYPE.PIE

    def test_alias_create_chart_action(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        touched = edits.apply_patches_to_file(src, dst, [{**CHART_PATCH, "action": "create_chart"}])
        assert touched == {0}
        assert getattr(_chart_of(Presentation(str(dst)).slides[0]), "chart_type", None) is not None

    def test_default_bbox_within_canvas(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(src, dst, [CHART_PATCH])
        prs = Presentation(str(dst))
        frame = [sh for sh in prs.slides[0].shapes if getattr(sh, "has_chart", False)][0]
        assert int(frame.left) >= 0 and int(frame.top) >= 0
        assert int(frame.left) + int(frame.width) <= int(prs.slide_width)
        assert int(frame.top) + int(frame.height) <= int(prs.slide_height)

    def test_values_length_must_match_categories(self, tmp_path):
        """与 edit_chart 同语义：系列值数量与类目数量不一致直接拒绝。"""
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        with pytest.raises(edits.PatchError, match="不一致"):
            edits.apply_patches_to_file(
                src, dst,
                [{**CHART_PATCH, "series": [{"name": "s", "values": [1, 2]}]}],
            )

    def test_pie_accepts_single_series_only(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        with pytest.raises(edits.PatchError, match="饼图"):
            edits.apply_patches_to_file(
                src, dst,
                [{**CHART_PATCH, "chart_type": "pie", "series": [
                    {"name": "a", "values": [1, 2, 3]},
                    {"name": "b", "values": [3, 2, 1]},
                ]}],
            )
        # 单系列饼图 OK
        edits.apply_patches_to_file(
            src, dst,
            [{**CHART_PATCH, "chart_type": "pie", "categories": ["甲", "乙", "丙"],
              "series": [{"name": "占比", "values": [5, 3, 2]}]}],
        )
        assert _chart_of(Presentation(str(dst)).slides[0]).chart_type == XL_CHART_TYPE.PIE

    def test_edit_chart_roundtrip_on_created_chart(self, tmp_path):
        """add_chart 造的图表能被既有 edit_chart 继续改数（真图表闭环）。"""
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(src, dst, [CHART_PATCH])
        prs = Presentation(str(dst))
        frame = [sh for sh in prs.slides[0].shapes if getattr(sh, "has_chart", False)][0]
        edits.apply_patches_to_file(
            dst, tmp_path / "c.pptx",
            [{"action": "edit_chart", "shape_id": f"s0_id{frame.shape_id}",
              "series": [{"name": "营收", "values": [99, 98, 97]}]}],
        )
        series = list(_chart_of(Presentation(str(tmp_path / "c.pptx")).slides[0]).series)
        assert [float(v) for v in series[0].values] == [99.0, 98.0, 97.0]


# ---------- #6 add_chart：端点 ----------


class TestAddChartEndpoint:
    def test_apply_add_chart_then_objects_show_chart(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"label": "加图表", "patches": [CHART_PATCH]},
        )
        assert resp.status_code == 200, resp.text
        doc = client.get(f"/api/docs/{doc_id}").json()
        charts = [o for o in doc["objects"] if o.get("kind") == "chart"]
        assert charts, "序列化对象里应出现 kind=chart"
        assert charts[0]["slide_index"] == 0

    def test_apply_add_chart_bad_payload_4xx(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"label": "坏图表", "patches": [{**CHART_PATCH, "chart_type": "nope"}]},
        )
        assert 400 <= resp.status_code < 500


# ---------- #7 母版版式：服务层 ----------


class TestDefaultLayoutService:
    def _branded_deck(self, path, content_pages: int = 2):
        """2 页用「标题和内容」版式 + 1 页空白 → 最常用版式是标题和内容。"""
        prs = Presentation()
        prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
        for _ in range(content_pages):
            prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
        prs.slides.add_slide(prs.slide_layouts[6])  # Blank
        prs.save(str(path))
        return path

    def test_add_slide_defaults_to_most_used_layout(self, tmp_path):
        src = self._branded_deck(tmp_path / "brand.pptx")
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(src, dst, [{"action": "add_slide"}])
        prs = Presentation(str(dst))
        assert len(prs.slides) == 4
        assert prs.slides[3].slide_layout.name == "Title and Content"

    def test_all_blank_deck_stays_blank(self, tmp_path):
        """全空白菜式稿子：最常用仍是空白，行为不回退到版式 0。"""
        src, dst = _deck(tmp_path / "a.pptx", pages=3, layout_index=6), tmp_path / "out.pptx"
        edits.apply_patches_to_file(src, dst, [{"action": "add_slide"}])
        prs = Presentation(str(dst))
        assert prs.slides[3].slide_layout.name == "Blank"

    def test_layout_name_selection(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(
            src, dst, [{"action": "add_slide", "layout_name": "标题和内容"}]
        )
        prs = Presentation(str(dst))
        # python-pptx 默认模板版式名是英文，中文名应模糊匹配到 Title and Content
        assert prs.slides[2].slide_layout.name == "Title and Content"

    def test_layout_name_english_exact(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(
            src, dst, [{"action": "add_slide", "layout_name": "Two Content"}]
        )
        assert Presentation(str(dst)).slides[2].slide_layout.name == "Two Content"

    def test_unknown_layout_name_rejected(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        with pytest.raises(edits.PatchError, match="layout_name"):
            edits.apply_patches_to_file(
                src, dst, [{"action": "add_slide", "layout_name": "不存在的版式"}]
            )

    def test_explicit_layout_index_still_wins(self, tmp_path):
        src, dst = _deck(tmp_path / "a.pptx"), tmp_path / "b.pptx"
        edits.apply_patches_to_file(src, dst, [{"action": "add_slide", "layout_index": 0}])
        assert Presentation(str(dst)).slides[2].slide_layout.name == "Title Slide"


# ---------- #7 母版版式：端点 ----------


class TestLayoutsEndpoint:
    def test_layouts_listed_with_usage(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.get(f"/api/docs/{doc_id}/layouts")
        assert resp.status_code == 200, resp.text
        payload = resp.json()
        layouts = payload["layouts"]
        assert len(layouts) >= 5, "默认模板有 11 个版式"
        names = {item["name"] for item in layouts}
        assert "Blank" in names and "Title and Content" in names
        # demo 稿 3 页全用空白版式 → used_by 统计正确
        blank = next(item for item in layouts if item["name"] == "Blank")
        assert blank["used_by"] == 3
        assert payload["slide_count"] == 3
        # 每个版式都带占位符类型清单（默认模板空白版式仅日期/页脚/页码）
        assert isinstance(blank["placeholders"], list)
        assert all(ph["type"] in ("date", "footer", "slide_number") for ph in blank["placeholders"])

    def test_doc_payload_contains_layouts(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        doc = client.get(f"/api/docs/{doc_id}").json()
        assert isinstance(doc.get("layouts"), list) and doc["layouts"]

    def test_add_slide_with_layout_name_via_apply(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"label": "按版式加页",
                  "patches": [{"action": "add_slide", "layout_name": "标题和内容"}]},
        )
        assert resp.status_code == 200, resp.text
        doc = client.get(f"/api/docs/{doc_id}").json()
        assert doc["slide_count"] == 4
