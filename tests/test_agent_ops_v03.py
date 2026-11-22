"""v0.3 agent 操作存储：10 个新编辑操作（font_unify/replace_text/align_objects/
duplicate_object/z_order/table_mod/replace_image/slide_background/
add_page_numbers/font_scale）经 propose_operations 的校验与默认值。"""
from __future__ import annotations

import pytest
from pptx import Presentation

from services import agent, parse


def _store_from(pptx_path, instruction="批量整理", scope=None):
    objects, payload = parse.parse_pptx(str(pptx_path))
    total = payload["slide_count"]
    view = agent._DeckView(
        objects=objects,
        elements=payload["elements"],
        tables=payload["tables"],
        total_slides=total,
        scope_slides=scope if scope is not None else list(range(total)),
        slide_size=payload["slide_size_emu"],
    )
    return agent._OperationStore(view, instruction, 24)


def _store(total=3, scope=None):
    view = agent._DeckView(
        objects=[], elements=[], tables=[], total_slides=total,
        scope_slides=scope if scope is not None else list(range(total)),
        slide_size=None,
    )
    return agent._OperationStore(view, "整理", 24)


@pytest.fixture()
def pic_deck(tmp_path):
    import io

    from pptx.util import Emu

    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    png = bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
    )
    slide.shapes.add_picture(io.BytesIO(png), Emu(1000000), Emu(1000000), Emu(2000000), Emu(1500000))
    tb = slide.shapes.add_textbox(Emu(1000000), Emu(4000000), Emu(4000000), Emu(800000))
    tb.text_frame.text = "配图说明"
    path = tmp_path / "pic.pptx"
    prs.save(str(path))
    return path


# ---------- deck 批量操作 ----------


def test_agent_font_unify_defaults_to_deck():
    store = _store()
    result = store.add_batch([{"action": "font_unify", "font_name": "微软雅黑"}])
    assert result["accepted"] == 1
    op = store.ops[0]
    assert op["action"] == "font_unify"
    assert op["font_name"] == "微软雅黑"
    assert op["scope"] == "deck"
    assert op["slide_index"] is None


def test_agent_font_unify_slide_scope_and_validation():
    store = _store()
    result = store.add_batch([
        {"action": "font_unify", "font_name": "黑体", "slide_index": 1},
        {"action": "font_unify"},                       # 缺 font_name
        {"action": "font_unify", "font_name": "X", "slide_index": 9},  # 越界
        {"action": "font_unify", "font_name": "X", "slide_index": 2, "scope": "deck"},  # 冲突
    ])
    assert result["accepted"] == 1 and result["dropped"] == 3
    assert store.ops[0]["slide_index"] == 1 and store.ops[0]["scope"] == "slide"


def test_agent_replace_text_requires_find():
    store = _store()
    result = store.add_batch([
        {"action": "replace_text", "find": "旧词", "replace": "新词"},
        {"action": "replace_text", "replace": "新词"},
        {"action": "replace_text", "find": "旧词", "replace": "新词", "slide_index": 1},
    ])
    assert result["accepted"] == 2 and result["dropped"] == 1
    assert store.ops[0]["scope"] == "deck"
    assert store.ops[1]["slide_index"] == 1


def test_agent_font_scale_delta_bounds():
    store = _store()
    result = store.add_batch([
        {"action": "font_scale", "delta_pt": 2},
        {"action": "font_scale"},                # 缺 delta
        {"action": "font_scale", "delta_pt": 99},  # 超界
        {"action": "font_scale", "delta_pt": -3, "slide_index": 0},
    ])
    assert result["accepted"] == 2 and result["dropped"] == 2


def test_agent_add_page_numbers_options():
    store = _store()
    result = store.add_batch([
        {"action": "add_page_numbers"},
        {"action": "add_page_numbers", "skip_first": True, "position": "bottom-center"},
        {"action": "add_page_numbers", "position": "middle-right"},  # 非法位置
    ])
    assert result["accepted"] == 2 and result["dropped"] == 1
    assert store.ops[0]["skip_first"] is False
    assert store.ops[1]["skip_first"] is True


# ---------- 需要真实对象视图的操作 ----------


def test_agent_align_objects_from_real_deck(demo_pptx):
    store = _store_from(demo_pptx)
    ids = [
        o["shape_id"] for o in store.view.objects
        if o["slide_index"] == 0 and o["kind"] == "text"
    ][:2]
    assert len(ids) == 2
    result = store.add_batch([
        {"action": "align_objects", "shape_ids": ids, "align": "left"},
        {"action": "align_objects", "shape_ids": ids[:1], "align": "left"},     # <2
        {"action": "align_objects", "shape_ids": ["s9_id1", "s9_id2"], "align": "left"},  # 不存在
        {"action": "align_objects", "shape_ids": ids, "align": "wat"},          # 非法模式
    ])
    assert result["accepted"] == 1 and result["dropped"] == 3


def test_agent_table_mod_targets_table(demo_pptx):
    store = _store_from(demo_pptx)
    table_id = store.view.tables[0]["shape_id"]
    result = store.add_batch([
        {"action": "table_mod", "shape_id": table_id, "op": "add_row"},
        {"action": "table_mod", "shape_id": table_id, "op": "merge_cells",
         "row": 0, "col": 0, "row2": 0, "col2": 1},
        {"action": "table_mod", "shape_id": store.view.objects[0]["shape_id"], "op": "add_row"},  # 非表格
        {"action": "table_mod", "shape_id": table_id, "op": "wat"},  # 非法 op
        {"action": "table_mod", "shape_id": table_id, "op": "delete_row"},  # 缺 index
    ])
    assert result["accepted"] == 2 and result["dropped"] == 3


def test_agent_z_order_and_duplicate_object(demo_pptx):
    store = _store_from(demo_pptx)
    text_id = store.view.objects[0]["shape_id"]
    pic_id = next(
        e["shape_id"] for e in store.view.elements if e["kind"] == "picture"
    )
    result = store.add_batch([
        {"action": "z_order", "shape_id": text_id, "mode": "front"},
        {"action": "z_order", "shape_id": text_id, "mode": "spin"},   # 非法模式
        {"action": "duplicate_object", "shape_id": pic_id, "offset_emu": {"x": 100000, "y": 100000}},
        {"action": "duplicate_object", "shape_id": "s0_id777"},       # 不存在
    ])
    assert result["accepted"] == 2 and result["dropped"] == 2


def test_agent_replace_image_requires_picture(pic_deck, tmp_path):
    store = _store_from(pic_deck)
    pic_id = next(e["shape_id"] for e in store.view.elements if e["kind"] == "picture")
    text_id = store.view.objects[0]["shape_id"]
    result = store.add_batch([
        {"action": "replace_image", "shape_id": pic_id, "image_path": "/tmp/x.png"},
        {"action": "replace_image", "shape_id": text_id, "image_path": "/tmp/x.png"},  # 非图片
        {"action": "replace_image", "shape_id": pic_id},  # 缺来源
    ])
    assert result["accepted"] == 1 and result["dropped"] == 2


def test_agent_slide_background_requires_index():
    store = _store()
    result = store.add_batch([
        {"action": "slide_background", "slide_index": 0, "color": "1F4E79"},
        {"action": "slide_background", "color": "1F4E79"},  # 缺 slide_index
        {"action": "slide_background", "slide_index": 0, "color": "nothex"},  # 非法色
        {"action": "slide_background", "slide_index": 0,
         "gradient": {"from": "0F172A", "to": "2563EB", "angle": 45}},
    ])
    assert result["accepted"] == 2 and result["dropped"] == 2


def test_agent_new_ops_flow_through_apply(demo_pptx, tmp_path):
    """store 产出的 op 原样可被 edits 层应用（agent→apply 闭环）。"""
    store = _store_from(demo_pptx)
    store.add_batch([
        {"action": "replace_text", "find": "项目", "replace": "工程", "scope": "deck"},
        {"action": "font_unify", "font_name": "微软雅黑"},
    ])
    patches = [
        {k: v for k, v in op.items() if k not in ("id", "source", "meta", "before", "shape_name", "kind", "description")}
        for op in store.ops
    ]
    touched = __import__("services.edits", fromlist=["edits"]).apply_patches_to_file(
        demo_pptx, tmp_path / "out.pptx", patches
    )
    assert touched == {0, 1, 2}
    texts = [o["text"] for o in parse.parse_pptx(str(tmp_path / "out.pptx"))[0]]
    assert not any("项目" in t for t in texts)
