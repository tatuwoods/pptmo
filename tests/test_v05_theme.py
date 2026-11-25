"""v0.5 F1/F2：theme_rebrand 整套重主题 / recolor 调色板映射。"""
from __future__ import annotations

import pytest
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from services import edits


def _theme_roots(prs):
    roots = []
    for part in prs.part.package.iter_parts():
        name = str(getattr(part, "partname", ""))
        if name.startswith("/ppt/theme/theme") and name.endswith(".xml"):
            roots.append(etree.fromstring(part.blob))
    return roots


def _scheme(root, tag):
    for clr_scheme in root.iter(qn("a:clrScheme")):
        holder = clr_scheme.find(qn(f"a:{tag}"))
        if holder is not None:
            srgb = holder.find(qn("a:srgbClr"))
            return srgb.get("val") if srgb is not None else None
    return None


def _typeface(root, font_tag):
    for font_scheme in root.iter(qn("a:fontScheme")):
        holder = font_scheme.find(qn(f"a:{font_tag}"))
        if holder is not None:
            latin = holder.find(qn("a:latin"))
            return latin.get("typeface") if latin is not None else None
    return None


# ---------- F1: theme_rebrand ----------

class TestThemeRebrand:
    def test_palette_replaces_scheme_colors(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            demo_pptx,
            dst,
            [{"action": "theme_rebrand", "payload": {
                "palette": {"accent1": "1A1A2E", "accent2": "16213E"},
            }}],
        )
        assert touched == {0, 1, 2}
        root = _theme_roots(Presentation(str(dst)))[0]
        assert _scheme(root, "accent1") == "1A1A2E"
        assert _scheme(root, "accent2") == "16213E"
        # 未指定的槽位保持原值
        assert _scheme(root, "accent3") is not None
        assert _scheme(root, "accent3") != "1A1A2E"

    def test_palette_aliases_text1_background1(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "theme_rebrand", "payload": {
                "palette": {"text1": "0F0F0F", "background1": "F5F5F5"},
            }}],
        )
        root = _theme_roots(Presentation(str(dst)))[0]
        assert _scheme(root, "dk1") == "0F0F0F"
        assert _scheme(root, "lt1") == "F5F5F5"

    def test_fonts_major_minor(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "theme_rebrand", "payload": {
                "fonts": {"major": "Georgia", "minor": "Verdana"},
            }}],
        )
        root = _theme_roots(Presentation(str(dst)))[0]
        assert _typeface(root, "majorFont") == "Georgia"
        assert _typeface(root, "minorFont") == "Verdana"

    def test_palette_and_fonts_together(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "theme_rebrand", "payload": {
                "palette": {"accent1": "E94560"},
                "fonts": {"major": "Impact", "minor": "Menlo"},
            }}],
        )
        root = _theme_roots(Presentation(str(dst)))[0]
        assert _scheme(root, "accent1") == "E94560"
        assert _typeface(root, "majorFont") == "Impact"

    def test_scheme_color_shape_follows_theme(self, tmp_path):
        # 引用 schemeClr accent1 的形状随主题一起变色（theme 引用而非硬编码）
        from pptx.enum.dml import MSO_THEME_COLOR

        src = tmp_path / "src.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(3000000), Emu(600000))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = "品牌色文字"
        run.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            src, dst,
            [{"action": "theme_rebrand", "payload": {"palette": {"accent1": "0066CC"}}}],
        )
        prs2 = Presentation(str(dst))
        shape = prs2.slides[0].shapes[0]
        # XML 层仍是 schemeClr 引用（主题生效的标志）
        scheme_ref = shape.text_frame.paragraphs[0].runs[0]._r.find(
            qn("a:rPr")
        ).find(qn("a:solidFill")).find(qn("a:schemeClr"))
        assert scheme_ref is not None and scheme_ref.get("val") == "accent1"
        assert _scheme(_theme_roots(prs2)[0], "accent1") == "0066CC"

    def test_invalid_hex_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "theme_rebrand", "payload": {"palette": {"accent1": "RED"}}}],
            )

    def test_unknown_slot_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "theme_rebrand", "payload": {"palette": {"accent9": "112233"}}}],
            )

    def test_empty_payload_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "theme_rebrand", "payload": {}}],
            )

    def test_alias_replace_color(self, demo_pptx, tmp_path):
        # 动作名别名：set_theme
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_theme", "payload": {"palette": {"accent4": "ABCDEF"}}}],
        )
        root = _theme_roots(Presentation(str(dst)))[0]
        assert _scheme(root, "accent4") == "ABCDEF"


# ---------- F2: recolor ----------

def _add_filled_shape(slide, color_hex, x=Emu(100000)):
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Emu(100000), Emu(1000000), Emu(500000))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color_hex)
    return shape


class TestRecolor:
    def test_solid_fill_replaced(self, tmp_path):
        src = tmp_path / "src.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_filled_shape(slide, "FF0000")
        _add_filled_shape(slide, "00AA00", x=Emu(1500000))
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            src, dst, [{"action": "recolor", "payload": {"mapping": {"FF0000": "0066CC"}}}],
        )
        assert touched == {0}
        prs2 = Presentation(str(dst))
        colors = [
            s.fill.fore_color.rgb for s in prs2.slides[0].shapes if s.shape_type != 6
        ]
        assert str(colors[0]) == "0066CC"
        assert str(colors[1]) == "00AA00"  # 未映射的颜色不动

    def test_mapping_keys_case_insensitive(self, tmp_path):
        src = tmp_path / "src.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_filled_shape(slide, "FF0000")
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            src, dst, [{"action": "recolor", "payload": {"mapping": {"ff0000": "112233"}}}],
        )
        prs2 = Presentation(str(dst))
        shape = prs2.slides[0].shapes[0]
        assert str(shape.fill.fore_color.rgb) == "112233"

    def test_gradient_stop_replaced(self, tmp_path):
        # 渐变 stop 也是 srgbClr，必须一并替换
        src = tmp_path / "src.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = _add_filled_shape(slide, "FFFFFF")
        grad = etree.SubElement(
            shape._element.find(qn("p:spPr")),
            qn("a:gradFill"),
        )
        gs_lst = etree.SubElement(grad, qn("a:gsLst"))
        for pos, color in (("0", "FF0000"), ("100000", "0000FF")):
            gs = etree.SubElement(gs_lst, qn("a:gs"))
            gs.set("pos", pos)
            etree.SubElement(gs, qn("a:srgbClr")).set("val", color)
        # solidFill 移除（已被 gradFill 替代语义）
        solid = shape._element.find(qn("p:spPr")).find(qn("a:solidFill"))
        if solid is not None:
            shape._element.find(qn("p:spPr")).remove(solid)
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            src, dst,
            [{"action": "recolor", "payload": {"mapping": {"FF0000": "0066CC", "0000FF": "FFFFFF"}}}],
        )
        prs2 = Presentation(str(dst))
        sp_pr = prs2.slides[0].shapes[0]._element.find(qn("p:spPr"))
        vals = [n.get("val") for n in sp_pr.iter(qn("a:srgbClr"))]
        assert "0066CC" in vals and "FFFFFF" in vals
        assert "FF0000" not in vals and "0000FF" not in vals

    def test_line_color_replaced(self, tmp_path):
        src = tmp_path / "src.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Emu(100000), Emu(100000), Emu(2000000), Emu(500000))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = "红字"
        run.font.color.rgb = RGBColor.from_string("CC0000")
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            src, dst, [{"action": "recolor", "payload": {"mapping": {"CC0000": "336699"}}}],
        )
        prs2 = Presentation(str(dst))
        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert str(run2.font.color.rgb) == "336699"

    def test_master_layout_background_covered(self, tmp_path):
        # 母版/版式上的硬编码颜色也要替换（全稿语义）。
        # LayoutShapes 没有 add_textbox API，用裸 XML 往版式 spTree 塞一个带色文本框。
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        src = tmp_path / "src.pptx"
        prs = Presentation()
        layout = prs.slide_masters[0].slide_layouts[0]
        sp_xml = (
            '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'{nsdecls("a")}><p:nvSpPr><p:cNvPr id="99" name="母版文字"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>'
            '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="zh-CN">'
            '<a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:rPr>'
            '<a:t>母版文字</a:t></a:r></a:p></p:txBody></p:sp>'
        )
        layout.shapes._spTree.append(parse_xml(sp_xml))
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            src, dst, [{"action": "recolor", "payload": {"mapping": {"123456": "654321"}}}],
        )
        prs2 = Presentation(str(dst))
        tree2 = prs2.slide_masters[0].slide_layouts[0].shapes._spTree
        vals = [n.get("val") for n in tree2.iter(qn("a:srgbClr"))]
        assert "654321" in vals and "123456" not in vals

    def test_empty_mapping_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "recolor", "payload": {"mapping": {}}}],
            )

    def test_invalid_value_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "recolor", "payload": {"mapping": {"FF0000": "BLUE"}}}],
            )

    def test_alias_replace_color(self, tmp_path):
        src = tmp_path / "src.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_filled_shape(slide, "FF0000")
        prs.save(str(src))
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            src, dst, [{"action": "replace_color", "payload": {"mapping": {"FF0000": "0066CC"}}}],
        )
        prs2 = Presentation(str(dst))
        assert str(prs2.slides[0].shapes[0].fill.fore_color.rgb) == "0066CC"


# ---------- API 集成 ----------

class TestThemeApi:
    def test_apply_theme_rebrand_via_api(self, client, uploaded_doc, tmp_path):
        doc_id = uploaded_doc["doc_id"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "theme_rebrand", "payload": {
                "palette": {"accent1": "1A1A2E"},
                "fonts": {"major": "Georgia"},
            }}], "label": "重主题"},
        )
        assert resp.status_code == 200, resp.text
        data = resp.json()
        assert data["version"] == 1
        assert data["touched_slides"] == [0, 1, 2]
        # 下载并验证主题已写入
        with client.stream("GET", f"/api/docs/{doc_id}/download") as dl:
            body = b"".join(dl.iter_bytes())
        prs = Presentation(__import__("io").BytesIO(body))
        root = _theme_roots(prs)[0]
        assert _scheme(root, "accent1") == "1A1A2E"
        assert _typeface(root, "majorFont") == "Georgia"

    def test_apply_recolor_via_api(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        # demo deck 没有硬编码色：先给一个文本对象上色，再做映射替换
        text_obj = next(o for o in uploaded_doc["objects"] if o.get("kind") == "text")
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "style", "shape_id": text_obj["shape_id"],
                               "style": {"color": "CC1100"}}], "label": "上色"},
        )
        assert resp.status_code == 200, resp.text
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "recolor", "payload": {
                "mapping": {"CC1100": "334455"},
            }}], "label": "调色"},
        )
        assert resp.status_code == 200, resp.text
        assert resp.json()["touched_slides"] == [text_obj["slide_index"]]
        # 换完的颜色可以直接从下载的文件里验证
        import io as _io

        data = client.get(f"/api/docs/{doc_id}/download").content
        prs = Presentation(_io.BytesIO(data))
        colors = set()
        for slide in prs.slides:
            for node in slide._element.iter(qn("a:srgbClr")):
                colors.add(node.get("val"))
        assert "334455" in colors and "CC1100" not in colors


# ---------- agent 校验（propose_operations 侧） ----------

from services import agent  # noqa: E402


def _store(total=3, scope=None):
    view = agent._DeckView(
        objects=[], elements=[], tables=[], total_slides=total,
        scope_slides=scope if scope is not None else list(range(total)),
        slide_size=None,
    )
    return agent._OperationStore(view, "换品牌色", 24)


class TestAgentThemeOps:
    def test_theme_rebrand_accepted(self):
        store = _store()
        result = store.add_batch([
            {"action": "theme_rebrand", "palette": {"accent1": "1A1A2E"}, "fonts": {"major": "Georgia"}},
            {"action": "theme_rebrand", "palette": {"accent1": "bad!"}},  # 非法色
            {"action": "theme_rebrand"},  # 缺参数
            {"action": "theme_rebrand", "palette": {"accent9": "112233"}},  # 未知槽位
        ])
        assert result["accepted"] == 1 and result["dropped"] == 3
        op = store.ops[0]
        assert op["palette"] == {"accent1": "1A1A2E"}
        assert op["fonts"] == {"major": "Georgia"}
        assert op["scope"] == "deck"

    def test_theme_rebrand_needs_deck_scope(self):
        store = _store(total=3, scope=[0])  # 当前页范围
        result = store.add_batch([{"action": "theme_rebrand", "palette": {"accent1": "1A1A2E"}}])
        assert result["accepted"] == 0
        assert "整份" in result["reasons"][0]["reason"]

    def test_recolor_accepted_and_validated(self):
        store = _store()
        result = store.add_batch([
            {"action": "recolor", "mapping": {"FF0000": "0066CC"}},
            {"action": "recolor"},  # 缺 mapping
            {"action": "recolor", "mapping": {"FF0000": "short"}},
            {"action": "replace_color", "mapping": {"FF0000": "112233"}},  # 别名
        ])
        assert result["accepted"] == 2 and result["dropped"] == 2
        assert store.ops[0]["mapping"] == {"FF0000": "0066CC"}
        assert store.ops[1]["action"] == "recolor"
