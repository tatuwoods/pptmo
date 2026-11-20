"""Unit tests for backend/services: edits, parse, agent validators, ai helpers."""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from services import agent, ai, edits, parse


# ---------- edits.parse_shape_id / parse_cell_id ----------

def test_parse_shape_id_variants():
    assert edits.parse_shape_id("s3_id12") == (3, "id", 12)
    assert edits.parse_shape_id("s0_sh5") == (0, "index", 5)
    assert edits.parse_shape_id("bogus") is None
    assert edits.parse_shape_id("") is None
    assert edits.parse_shape_id(None) is None


def test_parse_cell_id_variants():
    assert edits.parse_cell_id("s3_id4_c0_1") == ("s3_id4", 0, 1)
    assert edits.parse_cell_id("s3_id4_c12_3") == ("s3_id4", 12, 3)
    assert edits.parse_cell_id("s3_id4") is None
    assert edits.parse_cell_id("s3_id4_c0") is None
    assert edits.parse_cell_id("s3_id4_c0_1_extra") is None


# ---------- set_text_preserving_style ----------

def _find_textbox(slide, text):
    for shape in slide.shapes:
        if shape.has_text_frame and text in shape.text:
            return shape
    return None


def test_set_text_preserves_style(demo_pptx):
    prs = Presentation(str(demo_pptx))
    shape = _find_textbox(prs.slides[0], "项目进展汇报")
    edits.set_text_preserving_style(shape, "新标题文本")
    assert shape.text == "新标题文本"
    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.font.size == Pt(32)
    assert run.font.bold is True


def test_set_text_multiline(demo_pptx):
    prs = Presentation(str(demo_pptx))
    shape = _find_textbox(prs.slides[0], "项目进展汇报")
    edits.set_text_preserving_style(shape, "第一行\n第二行")
    paras = shape.text_frame.paragraphs
    assert len(paras) == 2
    assert paras[0].runs[0].text == "第一行"
    assert paras[1].runs[0].text == "第二行"


def test_set_text_empty_string(demo_pptx):
    prs = Presentation(str(demo_pptx))
    shape = _find_textbox(prs.slides[0], "项目进展汇报")
    edits.set_text_preserving_style(shape, "")
    assert shape.text == ""


# ---------- apply_patches_to_file ----------

def _objects_of(demo_pptx):
    objects, _payload = parse.parse_pptx(str(demo_pptx))
    return objects


def test_apply_edit_patch(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "项目进展汇报" in o["text"])
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "edit", "shape_id": target["shape_id"], "after": "季度总结汇报", "before": target["text"]}],
    )
    assert touched == {0}
    prs = Presentation(str(dst))
    assert "季度总结汇报" in prs.slides[0].shapes[0].text or any(
        "季度总结汇报" in s.text for s in prs.slides[0].shapes if s.has_text_frame
    )


def test_apply_table_cell_edit_preserves_style(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    cell = next(o for o in objects if o["text"] == "200 万")
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "edit", "shape_id": cell["shape_id"], "after": "180 万", "before": "200 万"}],
    )
    prs = Presentation(str(dst))
    table = next(s for s in prs.slides[1].shapes if s.has_table).table
    texts = [table.cell(r, c).text for r in range(3) for c in range(3)]
    assert "180 万" in texts and "200 万" not in texts


def test_apply_style_patch(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "style", "shape_id": target["shape_id"], "style": {"bold": True, "color": "FF0000", "font_size_pt": 20}}],
    )
    prs = Presentation(str(dst))
    shape = _find_textbox(prs.slides[0], "本季度完成")
    run = shape.text_frame.paragraphs[0].runs[0]
    assert run.font.bold is True
    assert str(run.font.color.rgb) == "FF0000"
    assert run.font.size == Pt(20)


def test_apply_add_and_delete_textbox(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [
            {"action": "add_textbox", "slide_index": 0, "after": "新增备注",
             "bbox_emu": {"x": 1000000, "y": 5000000, "w": 3000000, "h": 500000}},
            {"action": "delete_textbox", "shape_id": target["shape_id"]},
        ],
    )
    prs = Presentation(str(dst))
    texts = [s.text for s in prs.slides[0].shapes if s.has_text_frame]
    assert any("新增备注" in t for t in texts)
    assert not any("本季度完成" in t for t in texts)


def test_apply_geometry_patch(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    dst = tmp_path / "out.pptx"
    new_bbox = {"x": 2000000, "y": 2500000, "w": 7000000, "h": 1500000}
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "geometry", "shape_id": target["shape_id"], "bbox_emu": new_bbox}],
    )
    prs = Presentation(str(dst))
    shape = _find_textbox(prs.slides[0], "本季度完成")
    assert (shape.left, shape.top, shape.width, shape.height) == (
        new_bbox["x"], new_bbox["y"], new_bbox["w"], new_bbox["h"]
    )


def test_apply_duplicate_patch_rejected(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    patch = {"action": "edit", "shape_id": target["shape_id"], "after": "x"}
    with pytest.raises(edits.PatchError, match="duplicate"):
        edits.apply_patches_to_file(demo_pptx, tmp_path / "o.pptx", [patch, dict(patch)])


def test_apply_before_mismatch_rejected(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    target = next(o for o in objects if "本季度完成" in o["text"])
    with pytest.raises(edits.PatchConflict):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "edit", "shape_id": target["shape_id"], "after": "x", "before": "完全不同的原文"}],
        )


def test_apply_too_many_patches_rejected(demo_pptx, tmp_path):
    patch = {"action": "add_textbox", "slide_index": 0, "after": "x",
             "bbox_emu": {"x": 0, "y": 0, "w": 100000, "h": 100000}}
    with pytest.raises(edits.PatchError, match="too many"):
        edits.apply_patches_to_file(demo_pptx, tmp_path / "o.pptx", [patch] * 101)


def test_apply_bad_shape_id_rejected(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "edit", "shape_id": "s99_id999", "after": "x"}],
        )


def test_apply_geometry_on_group_member_rejected(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    grouped = next(o for o in objects if o.get("grouped") and o["kind"] == "text")
    with pytest.raises(edits.PatchError, match="group"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "geometry", "shape_id": grouped["shape_id"],
              "bbox_emu": {"x": 0, "y": 0, "w": 1000000, "h": 500000}}],
        )


def test_apply_group_text_edit_allowed(demo_pptx, tmp_path):
    objects = _objects_of(demo_pptx)
    grouped = next(o for o in objects if o.get("grouped") and "总体情况" in o["text"])
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "edit", "shape_id": grouped["shape_id"], "after": "第一章 综合情况"}],
    )
    parsed, _ = parse.parse_pptx(str(dst))
    assert any("综合情况" in o["text"] for o in parsed)


# ---------- ai.extract_requested_slide ----------

@pytest.mark.parametrize("text,total,expected", [
    ("只改第 4 页的标题", 20, 3),
    ("第4页要加粗", 20, 3),
    ("第十二页呢", 20, 11),
    ("首页标题改成X", 5, 0),
    ("第一页", 5, 0),
    ("please fix page 7", 10, 6),
    ("slide 3 title", 10, 2),
    ("第 99 页", 20, None),   # out of range -> None
    ("没有页码的普通指令", 10, None),
])
def test_extract_requested_slide(text, total, expected):
    assert ai.extract_requested_slide(text, total) == expected


# ---------- ai.validate_patches ----------

def _fake_objects():
    return [
        {"slide_index": 0, "shape_id": "s0_id2", "shape_index": 1, "text": "正文内容",
         "bbox_emu": {"x": 0, "y": 3000000, "w": 5000000, "h": 1000000},
         "slide_size_emu": {"width": 12192000, "height": 6858000}, "font_size_pt": 18,
         "alignment": "LEFT"},
        {"slide_index": 0, "shape_id": "s0_id3", "shape_index": 2, "text": "12",
         "bbox_emu": {"x": 0, "y": 6500000, "w": 1000000, "h": 300000},
         "slide_size_emu": {"width": 12192000, "height": 6858000}, "font_size_pt": 10,
         "alignment": "CENTER"},
    ]


def test_validate_accepts_good_edit():
    patches, dropped = ai.validate_patches(
        [{"action": "edit", "shape_id": "s0_id2", "after": "新内容", "before": "正文内容"}],
        _fake_objects(), limit=10, instruction="改一下",
    )
    assert len(patches) == 1 and dropped == 0
    assert patches[0]["shape_id"] == "s0_id2"


def test_validate_protects_page_number():
    patches, dropped = ai.validate_patches(
        [{"action": "edit", "shape_id": "s0_id3", "after": "99"}],
        _fake_objects(), limit=10, instruction="改一下",
    )
    assert patches == [] and dropped == 1


def test_validate_allows_footer_when_instruction_names_it():
    patches, _ = ai.validate_patches(
        [{"action": "edit", "shape_id": "s0_id3", "after": "99"}],
        _fake_objects(), limit=10, instruction="把页码改成99",
    )
    assert len(patches) == 1


def test_validate_drops_empty_and_identical_after():
    patches, dropped = ai.validate_patches(
        [
            {"action": "edit", "shape_id": "s0_id2", "after": ""},
            {"action": "edit", "shape_id": "s0_id2", "after": "正文内容"},
        ],
        _fake_objects(), limit=10, instruction="改一下",
    )
    assert patches == [] and dropped == 2


# ---------- agent._normalize_style / _alignment_name ----------

def test_normalize_style_full():
    style = agent._normalize_style({"font_size_pt": "24", "font_name": " 黑体 ", "bold": "是",
                                    "italic": False, "color": "#ff00aa", "alignment": "居中"})
    assert style == {"font_size_pt": 24, "font_name": "黑体", "bold": True,
                     "italic": False, "color": "FF00AA", "alignment": "CENTER"}


def test_normalize_style_rejects_bad_values():
    style = agent._normalize_style({"font_size_pt": 999, "color": "red", "alignment": "斜的"})
    assert style == {}


def test_alignment_name_aliases():
    assert agent._alignment_name("居中") == "CENTER"
    assert agent._alignment_name("left") == "LEFT"
    assert agent._alignment_name("PP_ALIGN.RIGHT") == "RIGHT"
    assert agent._alignment_name(None) is None


# ---------- parse ----------

def test_parse_pptx_structure(demo_pptx):
    objects, payload = parse.parse_pptx(str(demo_pptx))
    assert payload["slide_count"] == 3
    kinds = [o["kind"] for o in objects]
    assert "text" in kinds and "table_cell" in kinds
    # merged-away cell (1,2) must not appear; origin cell carries span info
    cells = [o for o in objects if o["kind"] == "table_cell"]
    assert not any(o["row"] == 1 and o["col"] == 2 for o in cells)
    merged = next(o for o in cells if o["row"] == 1 and o["col"] == 1)
    assert merged["col_span"] == 2
    # group child present and flagged
    grouped = [o for o in objects if o.get("grouped")]
    assert any("总体情况" in o["text"] for o in grouped)
    # every bbox inside slide bounds
    for o in objects:
        bb = o["bbox_emu"]
        assert bb["x"] >= 0 and bb["y"] >= 0
        assert bb["x"] + bb["w"] <= payload["slide_size_emu"]["width"] + 1000


def test_parse_shape_ids_unique(demo_pptx):
    objects, _ = parse.parse_pptx(str(demo_pptx))
    ids = [o["shape_id"] for o in objects]
    assert len(ids) == len(set(ids))
