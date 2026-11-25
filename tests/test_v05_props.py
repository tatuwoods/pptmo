"""v0.5 #6：edit_props 文档元数据 + set_alt_text 对象替代文本。

edit_props 改 core properties（标题/作者/主题/关键词等），模板重品牌化后
一并改署名；set_alt_text 写 cNvPr@descr（无障碍 + AI 理解图片内容）。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn

from services import edits, parse


def _shape_id_at(demo_pptx, slide_index: int, predicate):
    prs = Presentation(str(demo_pptx))
    for shape in prs.slides[slide_index].shapes:
        if predicate(shape):
            return f"s{slide_index}_id{shape.shape_id}"
    raise AssertionError("fixture shape not found")


def _is_picture(shape):
    return str(getattr(shape, "shape_type", "")).startswith("PICTURE")


def _cnvpr_of(prs, slide_index: int, shape_id_num: int):
    for shape in prs.slides[slide_index].shapes:
        if int(shape.shape_id) == shape_id_num:
            for child in shape._element:
                if child.tag.endswith("}nvSpPr") or child.tag.endswith("}nvPicPr") \
                        or child.tag.endswith("}nvGraphicFramePr") or child.tag.endswith("}nvGrpSpPr"):
                    return child.find(qn("p:cNvPr"))
    raise AssertionError("cNvPr not found")


class TestEditProps:
    def test_sets_core_properties(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "edit_props", "payload": {
                "title": "季度汇报", "author": "张三", "subject": "Q3 复盘",
                "keywords": "销售, 复盘", "category": "内部", "comments": "仅供评审",
            }}],
        )
        props = Presentation(str(dst)).core_properties
        assert props.title == "季度汇报"
        assert props.author == "张三"
        assert props.subject == "Q3 复盘"
        assert props.keywords == "销售, 复盘"
        assert props.category == "内部"
        assert props.comments == "仅供评审"

    def test_partial_update_keeps_others(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "edit_props", "payload": {"title": "A", "author": "B"}}],
        )
        edits.apply_patches_to_file(
            dst, dst,
            [{"action": "edit_props", "payload": {"title": "C"}}],
        )
        props = Presentation(str(dst)).core_properties
        assert props.title == "C"
        assert props.author == "B"  # 未给的槽位不动

    def test_keywords_list_joined(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "edit_props", "payload": {"keywords": ["销售", "复盘", "增长"]}}],
        )
        assert Presentation(str(dst)).core_properties.keywords == "销售, 复盘, 增长"

    def test_value_coerced_to_str(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "edit_props", "payload": {"title": 2024}}],
        )
        assert Presentation(str(dst)).core_properties.title == "2024"

    def test_empty_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "edit_props", "payload": {}}],
            )

    def test_unknown_field_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "edit_props", "payload": {"foo": "bar"}}],
            )

    def test_aliases(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_props", "payload": {"title": "别名一号"}}],
        )
        edits.apply_patches_to_file(
            dst, dst,
            [{"action": "props", "payload": {"author": "别名二号"}}],
        )
        props = Presentation(str(dst)).core_properties
        assert props.title == "别名一号"
        assert props.author == "别名二号"


class TestAltText:
    def test_set_on_textbox(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, lambda s: s.has_text_frame)
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "页面主标题"}],
        )
        assert touched == {0}
        num = int(shape_id.split("id")[1])
        assert _cnvpr_of(Presentation(str(dst)), 0, num).get("descr") == "页面主标题"

    def test_set_on_picture(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, _is_picture)
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "柱状图：Q3 销售额"}],
        )
        num = int(shape_id.split("id")[1])
        assert _cnvpr_of(Presentation(str(dst)), 0, num).get("descr") == "柱状图：Q3 销售额"

    def test_overwrite_existing(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, lambda s: s.has_text_frame)
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "第一版"}],
        )
        edits.apply_patches_to_file(
            dst, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "第二版"}],
        )
        num = int(shape_id.split("id")[1])
        assert _cnvpr_of(Presentation(str(dst)), 0, num).get("descr") == "第二版"

    def test_empty_string_clears(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, lambda s: s.has_text_frame)
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "临时"}],
        )
        edits.apply_patches_to_file(
            dst, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": ""}],
        )
        num = int(shape_id.split("id")[1])
        assert _cnvpr_of(Presentation(str(dst)), 0, num).get("descr") is None

    def test_missing_alt_text_rejected(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, lambda s: s.has_text_frame)
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_alt_text", "shape_id": shape_id}],
            )

    def test_too_long_rejected(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, lambda s: s.has_text_frame)
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "长" * 1001}],
            )

    def test_unknown_shape_conflict(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_alt_text", "shape_id": "s0_id999", "alt_text": "x"}],
            )

    def test_alias_alt(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, lambda s: s.has_text_frame)
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "alt_text", "shape_id": shape_id, "alt_text": "别名路径"}],
        )
        num = int(shape_id.split("id")[1])
        assert _cnvpr_of(Presentation(str(dst)), 0, num).get("descr") == "别名路径"

    def test_parse_exposes_alt_text(self, demo_pptx, tmp_path):
        shape_id = _shape_id_at(demo_pptx, 0, _is_picture)
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_alt_text", "shape_id": shape_id, "alt_text": "给 AI 看的图说"}],
        )
        _objects, payload = parse.parse_pptx(dst)
        obj = next(o for o in payload["objects"] if o["shape_id"] == shape_id)
        assert obj["alt_text"] == "给 AI 看的图说"

    def test_parse_clean_deck_has_no_alt_field(self, demo_pptx):
        _objects, payload = parse.parse_pptx(demo_pptx)
        assert all("alt_text" not in o or not o["alt_text"] for o in payload["objects"])


class TestPropsApi:
    def test_apply_props_via_api(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "edit_props", "payload": {"title": "接口标题", "author": "接口作者"}}],
                  "label": "改元数据"},
        )
        assert resp.status_code == 200, resp.text
        assert resp.json()["version"] == 1

    def test_apply_alt_text_via_api(self, client, uploaded_doc):
        objects = client.get(f"/api/docs/{uploaded_doc['doc_id']}/objects").json()["objects"]
        target = next(o for o in objects if o["slide_index"] == 0 and o["kind"] == "picture")
        resp = client.post(
            f"/api/docs/{uploaded_doc['doc_id']}/apply",
            json={"patches": [{"action": "set_alt_text", "shape_id": target["shape_id"],
                               "alt_text": "API 图说"}],
                  "label": "补图说"},
        )
        assert resp.status_code == 200, resp.text
