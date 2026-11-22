"""v0.3 edit operations (A 组): font_unify / replace_text / align_objects /
duplicate_object / z_order / table_mod / replace_image / slide_background /
add_page_numbers / font_scale.

edits 层（apply_patches_to_file）+ agent 操作存储（_OperationStore）双层覆盖。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from services import agent, edits, parse

# ---------- helpers ----------


@pytest.fixture()
def mixed_font_pptx(tmp_path):
    """2 页：第 1 页两种字体（字号/加粗不同）+ 表格；第 2 页一种字体。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s0 = prs.slides.add_slide(blank)
    t1 = s0.shapes.add_textbox(Emu(500000), Emu(400000), Emu(6000000), Emu(800000))
    t1.text_frame.text = "标题宋体"
    r1 = t1.text_frame.paragraphs[0].runs[0]
    r1.font.name = "宋体"
    r1.font.size = Pt(32)
    r1.font.bold = True
    t2 = s0.shapes.add_textbox(Emu(500000), Emu(1500000), Emu(6000000), Emu(800000))
    t2.text_frame.text = "正文黑体"
    r2 = t2.text_frame.paragraphs[0].runs[0]
    r2.font.name = "黑体"
    r2.font.size = Pt(18)
    gfx = s0.shapes.add_table(2, 2, Emu(500000), Emu(2600000), Emu(6000000), Emu(1800000))
    gfx.table.cell(0, 0).text = "表头楷体"
    cell_run = gfx.table.cell(0, 0).text_frame.paragraphs[0].runs[0]
    cell_run.font.name = "楷体"
    cell_run.font.size = Pt(14)
    s1 = prs.slides.add_slide(blank)
    t3 = s1.shapes.add_textbox(Emu(500000), Emu(400000), Emu(6000000), Emu(800000))
    t3.text_frame.text = "第二页仿宋"
    r3 = t3.text_frame.paragraphs[0].runs[0]
    r3.font.name = "仿宋"
    r3.font.size = Pt(24)
    path = tmp_path / "mixed.pptx"
    prs.save(str(path))
    return path


def _all_runs(prs):
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        yield run
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            for run in para.runs:
                                yield run


def _find_run(prs, needle):
    for run in _all_runs(prs):
        if needle in (run.text or ""):
            return run
    return None


# ---------- 1. font_unify ----------


def test_font_unify_deck_keeps_size_and_bold(mixed_font_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        mixed_font_pptx, dst, [{"action": "font_unify", "font_name": "微软雅黑"}]
    )
    prs = Presentation(str(dst))
    runs = {run.text: run for run in _all_runs(prs) if run.text}
    assert all(run.font.name == "微软雅黑" for run in runs.values())
    # 字号与加粗差异保留
    assert runs["标题宋体"].font.size == Pt(32)
    assert runs["标题宋体"].font.bold is True
    assert runs["正文黑体"].font.size == Pt(18)
    assert runs["表头楷体"].font.size == Pt(14)
    assert touched == {0, 1}  # deck 范围两页都被触及


def test_font_unify_sets_east_asian_typeface(mixed_font_pptx, tmp_path):
    """font.name 只写 a:latin；中文正文必须同步 a:ea 才真正生效。"""
    from pptx.oxml.ns import qn

    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        mixed_font_pptx, dst, [{"action": "font_unify", "font_name": "微软雅黑"}]
    )
    prs = Presentation(str(dst))
    rPr = _find_run(prs, "标题宋体")._r.find(qn("a:rPr"))
    ea = rPr.find(qn("a:ea"))
    assert ea is not None and ea.get("typeface") == "微软雅黑"


def test_font_unify_slide_scope(mixed_font_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        mixed_font_pptx,
        dst,
        [{"action": "font_unify", "font_name": "微软雅黑", "slide_index": 0}],
    )
    prs = Presentation(str(dst))
    assert _find_run(prs, "标题宋体").font.name == "微软雅黑"
    assert _find_run(prs, "第二页仿宋").font.name == "仿宋"  # 范围外不动


def test_font_unify_requires_font_name(mixed_font_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            mixed_font_pptx, tmp_path / "x.pptx", [{"action": "font_unify"}]
        )


# ---------- 2. replace_text ----------


@pytest.fixture()
def replace_pptx(tmp_path):
    """3 页：目标词出现在 4 处（含表格与同段两次出现），第 3 页无关词。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s0 = prs.slides.add_slide(blank)
    t = s0.shapes.add_textbox(Emu(500000), Emu(400000), Emu(8000000), Emu(800000))
    t.text_frame.text = "旧项目启动会"
    s1 = prs.slides.add_slide(blank)
    t2 = s1.shapes.add_textbox(Emu(500000), Emu(400000), Emu(8000000), Emu(800000))
    t2.text_frame.text = "旧项目进度与旧项目风险"
    gfx = s1.shapes.add_table(2, 1, Emu(500000), Emu(1500000), Emu(6000000), Emu(1500000))
    gfx.table.cell(0, 0).text = "表格里的旧项目"
    s2 = prs.slides.add_slide(blank)
    t3 = s2.shapes.add_textbox(Emu(500000), Emu(400000), Emu(8000000), Emu(800000))
    t3.text_frame.text = "完全无关的内容"
    path = tmp_path / "replace.pptx"
    prs.save(str(path))
    return path


def test_replace_text_deck_wide(replace_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        replace_pptx,
        dst,
        [{"action": "replace_text", "find": "旧项目", "replace": "新计划", "scope": "deck"}],
    )
    prs = Presentation(str(dst))
    texts = [run.text for run in _all_runs(prs) if run.text]
    assert "新计划启动会" in texts
    assert "新计划进度与新计划风险" in texts  # 同段两处都替换
    assert "表格里的新计划" in texts  # 表格单元格也替换
    assert "完全无关的内容" in texts
    assert not any("旧项目" in t for t in texts)
    assert touched == {0, 1}


def test_replace_text_slide_scope_and_case(replace_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        replace_pptx,
        dst,
        [{"action": "replace_text", "find": "旧项目", "replace": "新计划", "slide_index": 0}],
    )
    prs = Presentation(str(dst))
    texts = [run.text for run in _all_runs(prs) if run.text]
    assert "新计划启动会" in texts
    assert "旧项目进度与旧项目风险" in texts  # 其他页不动


def test_replace_text_keeps_run_style(replace_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        replace_pptx, dst, [{"action": "replace_text", "find": "旧项目", "replace": "X"}]
    )
    # 逐 run 替换而非整段重写：其余 run 的样式不受影响
    prs = Presentation(str(dst))
    assert _find_run(prs, "启动会") is not None


def test_replace_text_validation(replace_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            replace_pptx, tmp_path / "a.pptx", [{"action": "replace_text", "replace": "x"}]
        )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            replace_pptx, tmp_path / "b.pptx",
            [{"action": "replace_text", "find": "旧项目", "replace": "新计划", "scope": "slide"}],
        )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            replace_pptx, tmp_path / "c.pptx",
            [{"action": "replace_text", "find": "旧项目", "replace": "新计划", "slide_index": 9}],
        )


# ---------- 3. align_objects ----------


@pytest.fixture()
def align_pptx(tmp_path):
    """1 页 3 个错落文本框 + 1 个组内文本框（不可对齐）。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    boxes = [(100000, 200000, 2000000, 500000), (500000, 1000000, 2000000, 500000),
             (900000, 1800000, 2000000, 500000)]
    for i, (x, y, w, h) in enumerate(boxes):
        t = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
        t.text_frame.text = f"框{i}"
    grp = slide.shapes.add_group_shape()
    inner = grp.shapes.add_textbox(Emu(4000000), Emu(3000000), Emu(2000000), Emu(500000))
    inner.text_frame.text = "组内"
    path = tmp_path / "align.pptx"
    prs.save(str(path))
    return path


def _top_level_boxes(prs, slide_index=0):
    out = []
    for shape in prs.slides[slide_index].shapes:
        if getattr(shape, "has_text_frame", False) and (shape.text or "").startswith("框"):
            out.append(shape)
    return out


ALIGN_TEST_CASES = {
    "left": lambda s: s.left,
    "right": lambda s: s.left + s.width,
    "top": lambda s: s.top,
    "bottom": lambda s: s.top + s.height,
    "hcenter": lambda s: s.left + s.width // 2,
    "vcenter": lambda s: s.top + s.height // 2,
}


@pytest.mark.parametrize("mode", sorted(ALIGN_TEST_CASES))
def test_align_objects_modes(align_pptx, tmp_path, mode):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        align_pptx,
        dst,
        [{"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3", "s0_id4"],
          "align": mode}],
    )
    prs = Presentation(str(dst))
    values = [ALIGN_TEST_CASES[mode](s) for s in _top_level_boxes(prs)]
    assert len(set(values)) == 1, f"{mode}: {values}"


def test_align_objects_distribute_h(align_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        align_pptx,
        dst,
        [{"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3", "s0_id4"],
          "align": "distribute_h"}],
    )
    prs = Presentation(str(dst))
    boxes = sorted(_top_level_boxes(prs), key=lambda s: s.left)
    gaps = [boxes[1].left - (boxes[0].left + boxes[0].width),
            boxes[2].left - (boxes[1].left + boxes[1].width)]
    assert abs(gaps[0] - gaps[1]) <= 1  # 等间距（EMU 取整容差 1）


def test_align_objects_distribute_v(align_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        align_pptx,
        dst,
        [{"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3", "s0_id4"],
          "align": "distribute_v"}],
    )
    prs = Presentation(str(dst))
    boxes = sorted(_top_level_boxes(prs), key=lambda s: s.top)
    gaps = [boxes[1].top - (boxes[0].top + boxes[0].height),
            boxes[2].top - (boxes[1].top + boxes[1].height)]
    assert abs(gaps[0] - gaps[1]) <= 1


def test_align_objects_validation(align_pptx, tmp_path):
    # 少于 2 个对象
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            align_pptx, tmp_path / "a.pptx",
            [{"action": "align_objects", "shape_ids": ["s0_id2"], "align": "left"}],
        )
    # 跨页对象
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            align_pptx, tmp_path / "b.pptx",
            [{"action": "align_objects", "shape_ids": ["s0_id2", "s1_id9"], "align": "left"}],
        )
    # 非法对齐模式
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            align_pptx, tmp_path / "c.pptx",
            [{"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3"], "align": "diagonal"}],
        )
    # 分组内对象不可对齐
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            align_pptx, tmp_path / "d.pptx",
            [{"action": "align_objects", "shape_ids": ["s0_id2", "s0_id6"], "align": "left"}],
        )


# ---------- 4. duplicate_object ----------

PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)


@pytest.fixture()
def dup_pptx(tmp_path):
    """1 页：文本框（带样式）+ 图片。"""
    import io

    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    t = slide.shapes.add_textbox(Emu(600000), Emu(500000), Emu(4000000), Emu(800000))
    t.text_frame.text = "被复制"
    run = t.text_frame.paragraphs[0].runs[0]
    run.font.size = Pt(20)
    run.font.bold = True
    slide.shapes.add_picture(io.BytesIO(PNG_BYTES), Emu(600000), Emu(2000000), Emu(600000), Emu(400000))
    path = tmp_path / "dup.pptx"
    prs.save(str(path))
    return path


def test_duplicate_object_textbox_with_offset(dup_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        dup_pptx,
        dst,
        [{"action": "duplicate_object", "shape_id": "s0_id2",
          "offset_emu": {"x": 300000, "y": 200000}}],
    )
    assert touched == {0}
    prs = Presentation(str(dst))  # reopen: broken rels would explode here
    boxes = [s for s in prs.slides[0].shapes if getattr(s, "has_text_frame", False) and s.text == "被复制"]
    assert len(boxes) == 2
    orig, copy = sorted(boxes, key=lambda s: s.left)
    assert copy.left == orig.left + 300000
    assert copy.top == orig.top + 200000
    assert copy.width == orig.width and copy.height == orig.height
    copy_run = copy.text_frame.paragraphs[0].runs[0]
    assert copy_run.font.size == Pt(20) and copy_run.font.bold is True
    assert orig.shape_id != copy.shape_id  # 新对象拿到不同的 shape_id


def test_duplicate_object_picture_rebinds_image(dup_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        dup_pptx, dst, [{"action": "duplicate_object", "shape_id": "s0_id3"}]
    )
    prs = Presentation(str(dst))
    pics = [s for s in prs.slides[0].shapes if s.shape_type == 13]
    assert len(pics) == 2
    assert pics[0].image.blob == PNG_BYTES
    assert pics[1].image.blob == PNG_BYTES
    # 默认偏移避免完全重叠
    assert pics[1].left != pics[0].left or pics[1].top != pics[0].top


def test_duplicate_object_defaults_and_errors(dup_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        dup_pptx, dst, [{"action": "copy_object", "shape_id": "s0_id2"}]
    )
    prs = Presentation(str(dst))
    boxes = [s for s in prs.slides[0].shapes if getattr(s, "has_text_frame", False) and s.text == "被复制"]
    assert len(boxes) == 2  # 默认偏移也必须可辨
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            dup_pptx, tmp_path / "e.pptx", [{"action": "duplicate_object", "shape_id": "s0_id99"}]
        )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            dup_pptx, tmp_path / "f.pptx",
            [{"action": "duplicate_object", "shape_id": "s0_id2", "offset_emu": {"x": -5, "y": 0}}],
        )


# ---------- 5. z_order ----------


@pytest.fixture()
def z_pptx(tmp_path):
    """1 页 3 个形状 A/B/C（XML 顺序即 z 顺序：A 最底）。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    for label in ("A", "B", "C"):
        t = slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(2000000), Emu(600000))
        t.text_frame.text = label
    path = tmp_path / "z.pptx"
    prs.save(str(path))
    return path


def _shape_texts_in_z_order(prs):
    return [s.text for s in prs.slides[0].shapes if getattr(s, "has_text_frame", False)]


@pytest.mark.parametrize(
    "shape_id,mode,expected",
    [
        ("s0_id2", "front", ["B", "C", "A"]),      # A 置顶
        ("s0_id4", "back", ["C", "A", "B"]),       # C 置底
        ("s0_id2", "forward", ["B", "A", "C"]),    # A 上移一层
        ("s0_id4", "backward", ["A", "C", "B"]),   # C 下移一层
    ],
)
def test_z_order_modes(z_pptx, tmp_path, shape_id, mode, expected):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        z_pptx, dst, [{"action": "z_order", "shape_id": shape_id, "mode": mode}]
    )
    prs = Presentation(str(dst))
    assert _shape_texts_in_z_order(prs) == expected


def test_z_order_validation(z_pptx, tmp_path):
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            z_pptx, tmp_path / "a.pptx",
            [{"action": "z_order", "shape_id": "s0_id2", "mode": "sideways"}],
        )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            z_pptx, tmp_path / "b.pptx", [{"action": "z_order", "shape_id": "s0_id77", "mode": "front"}]
        )


# ---------- 6. table_mod ----------


@pytest.fixture()
def table_pptx(tmp_path):
    """1 页 3×3 表格，文本有内容。"""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    gfx = slide.shapes.add_table(3, 3, Emu(500000), Emu(500000), Emu(9000000), Emu(3600000))
    table = gfx.table
    for r in range(3):
        for c in range(3):
            table.cell(r, c).text = f"R{r}C{c}"
    path = tmp_path / "table.pptx"
    prs.save(str(path))
    return path


def _table_of(prs, slide_index=0):
    for shape in prs.slides[slide_index].shapes:
        if getattr(shape, "has_table", False):
            return shape.table
    raise AssertionError("no table")


def test_table_mod_add_delete_row(table_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        table_pptx, dst, [{"action": "table_mod", "shape_id": "s0_id2", "op": "add_row"}]
    )
    prs = Presentation(str(dst))
    table = _table_of(prs)
    assert len(table.rows) == 4 and len(table.columns) == 3
    assert table.cell(0, 0).text == "R0C0"  # 原内容不动
    assert table.cell(3, 0).text == ""  # 新行可写、为空
    table.cell(3, 1).text = "新行写入"
    assert table.cell(3, 1).text == "新行写入"

    dst2 = tmp_path / "out2.pptx"
    edits.apply_patches_to_file(
        dst, dst2, [{"action": "table_mod", "shape_id": "s0_id2", "op": "delete_row", "index": 0}]
    )
    prs2 = Presentation(str(dst2))
    table2 = _table_of(prs2)
    assert len(table2.rows) == 3
    texts = [table2.cell(r, 0).text for r in range(3)]
    assert texts == ["R1C0", "R2C0", ""]


def test_table_mod_add_delete_col(table_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        table_pptx, dst, [{"action": "table_mod", "shape_id": "s0_id2", "op": "add_col"}]
    )
    prs = Presentation(str(dst))
    table = _table_of(prs)
    assert len(table.columns) == 4 and len(table.rows) == 3
    assert table.cell(1, 0).text == "R1C0"
    assert table.cell(1, 3).text == ""

    dst2 = tmp_path / "out2.pptx"
    edits.apply_patches_to_file(
        dst, dst2, [{"action": "table_mod", "shape_id": "s0_id2", "op": "delete_col", "index": 1}]
    )
    prs2 = Presentation(str(dst2))
    table2 = _table_of(prs2)
    assert len(table2.columns) == 3
    assert [table2.cell(0, c).text for c in range(3)] == ["R0C0", "R0C2", ""]


def test_table_mod_merge_and_split(table_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        table_pptx,
        dst,
        [{"action": "table_mod", "shape_id": "s0_id2", "op": "merge_cells",
          "row": 0, "col": 0, "row2": 0, "col2": 1}],
    )
    prs = Presentation(str(dst))
    table = _table_of(prs)
    assert table.cell(0, 0).is_merge_origin
    assert table.cell(0, 0).span_width == 2
    assert table.cell(0, 1).is_spanned
    assert table.cell(0, 0).text == "R0C0"

    dst2 = tmp_path / "out2.pptx"
    edits.apply_patches_to_file(
        dst, dst2,
        [{"action": "table_mod", "shape_id": "s0_id2", "op": "split_cells", "row": 0, "col": 0}],
    )
    prs2 = Presentation(str(dst2))
    table2 = _table_of(prs2)
    assert not table2.cell(0, 0).is_merge_origin
    assert not table2.cell(0, 1).is_spanned


def test_table_mod_merge_guards(table_pptx, tmp_path):
    # 先合并 (0,0)-(0,1)，再合并与之重叠的区域 → 报错
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        table_pptx, dst,
        [{"action": "table_mod", "shape_id": "s0_id2", "op": "merge_cells",
          "row": 0, "col": 0, "row2": 0, "col2": 1}],
    )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            dst, tmp_path / "x.pptx",
            [{"action": "table_mod", "shape_id": "s0_id2", "op": "merge_cells",
              "row": 0, "col": 1, "row2": 0, "col2": 2}],
        )
    # 合并区域跨列时不能删除该列
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            dst, tmp_path / "y.pptx",
            [{"action": "table_mod", "shape_id": "s0_id2", "op": "delete_col", "index": 0}],
        )


def test_table_mod_validation(table_pptx, tmp_path):
    # 非表格对象
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            table_pptx, tmp_path / "a.pptx",
            [{"action": "table_mod", "shape_id": "s0_id2", "op": "wat"}],
        )
    # 行号越界
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            table_pptx, tmp_path / "b.pptx",
            [{"action": "table_mod", "shape_id": "s0_id2", "op": "delete_row", "index": 9}],
        )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            table_pptx, tmp_path / "c.pptx",
            [{"action": "table_mod", "shape_id": "s0_id2", "op": "delete_col", "index": -1}],
        )
    # split 非合并起点
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            table_pptx, tmp_path / "d.pptx",
            [{"action": "table_mod", "shape_id": "s0_id2", "op": "split_cells", "row": 1, "col": 1}],
        )


# ---------- 2b. replace_text 边界 ----------


def test_replace_text_no_match_rejected(replace_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="未找到"):
        edits.apply_patches_to_file(
            replace_pptx, tmp_path / "o.pptx",
            [{"action": "replace_text", "find": "根本不存在的词", "replace": "x"}],
        )


def test_replace_text_spanning_runs(replace_pptx, tmp_path):
    """关键词被 run 边界切开时，退化为整段替换（保留首 run 样式）。"""
    prs = Presentation(str(replace_pptx))
    tb = next(s for s in prs.slides[0].shapes if getattr(s, "has_text_frame", False))
    p = tb.text_frame.paragraphs[0]
    for child in list(p._p):
        p._p.remove(child)
    r1 = p.add_run(); r1.text = "旧项"
    r2 = p.add_run(); r2.text = "目启动会"
    src = tmp_path / "spans.pptx"
    prs.save(str(src))
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        src, dst, [{"action": "replace_text", "find": "旧项目", "replace": "新计划"}]
    )
    texts = [run.text for run in _all_runs(Presentation(str(dst))) if run.text]
    assert any("新计划启动会" in t for t in texts)


# ---------- 3b. align_objects 边界 ----------


def test_align_objects_distribute_needs_three(align_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="3"):
        edits.apply_patches_to_file(
            align_pptx, tmp_path / "o.pptx",
            [{"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3"],
              "align": "distribute_h"}],
        )


# ---------- 4b. duplicate_object 分组内 ----------


def test_duplicate_object_inside_group(align_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        align_pptx, dst,
        [{"action": "duplicate_object", "shape_id": "s0_id6",
          "offset_emu": {"x": 100000, "y": 100000}}],
    )
    prs = Presentation(str(dst))
    grp = next(s for s in prs.slides[0].shapes if getattr(s, "shape_type", None) == 6)
    texts = [s.text for s in grp.shapes if getattr(s, "has_text_frame", False)]
    assert texts.count("组内") == 2  # 复制进了同一个组


# ---------- 7. replace_image ----------

TINY_PNG2 = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000e49444154789c626060f8cfc00000003b0106ac04d7f2a3b80000000049454e44ae426082"
)


@pytest.fixture()
def image_pptx(tmp_path):
    import io

    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    png = bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
    )
    slide.shapes.add_picture(io.BytesIO(png), Emu(1000000), Emu(1000000), Emu(2000000), Emu(1500000))
    t = slide.shapes.add_textbox(Emu(1000000), Emu(4000000), Emu(4000000), Emu(800000))
    t.text_frame.text = "配图说明"
    path = tmp_path / "img.pptx"
    prs.save(str(path))
    return path


def _pic_of(prs):
    return next(s for s in prs.slides[0].shapes if s.shape_type == 13)


def test_replace_image_by_base64(image_pptx, tmp_path):
    import base64

    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        image_pptx, dst,
        [{"action": "replace_image", "shape_id": "s0_id2",
          "image_base64": base64.b64encode(TINY_PNG2).decode("ascii")}],
    )
    prs = Presentation(str(dst))
    pic = _pic_of(prs)
    assert pic.image.blob == TINY_PNG2  # 内容换了
    assert (pic.left, pic.top, pic.width, pic.height) == (1000000, 1000000, 2000000, 1500000)


def test_replace_image_by_path(image_pptx, tmp_path):
    img = tmp_path / "new.png"
    img.write_bytes(TINY_PNG2)
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        image_pptx, dst,
        [{"action": "replace_image", "shape_id": "s0_id2", "image_path": str(img)}],
    )
    assert _pic_of(Presentation(str(dst))).image.blob == TINY_PNG2


def test_replace_image_validation(image_pptx, tmp_path):
    import base64

    # 非图片对象
    with pytest.raises(edits.PatchError, match="图片"):
        edits.apply_patches_to_file(
            image_pptx, tmp_path / "a.pptx",
            [{"action": "replace_image", "shape_id": "s0_id3",
              "image_base64": base64.b64encode(TINY_PNG2).decode()}],
        )
    # 不是 png/jpg
    with pytest.raises(edits.PatchError, match="png"):
        edits.apply_patches_to_file(
            image_pptx, tmp_path / "b.pptx",
            [{"action": "replace_image", "shape_id": "s0_id2",
              "image_base64": base64.b64encode(b"not-an-image-at-all").decode()}],
        )
    # 既无 base64 也无 path
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            image_pptx, tmp_path / "c.pptx",
            [{"action": "replace_image", "shape_id": "s0_id2"}],
        )
    # path 不存在
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            image_pptx, tmp_path / "d.pptx",
            [{"action": "replace_image", "shape_id": "s0_id2",
              "image_path": str(tmp_path / "missing.png")}],
        )


# ---------- 8. slide_background ----------


def _bg_fill_of(prs, slide_index):
    from pptx.oxml.ns import qn

    bg = prs.slides[slide_index]._element.cSld.find(qn("p:bg"))
    assert bg is not None, "no p:bg element"
    return bg


def test_slide_background_solid(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "slide_background", "slide_index": 1, "color": "1F4E79"}]
    )
    assert touched == {1}
    prs = Presentation(str(dst))
    from pptx.oxml.ns import qn

    bg = _bg_fill_of(prs, 1)
    srgb = bg.find(".//" + qn("a:srgbClr"))
    assert srgb is not None and srgb.get("val") == "1F4E79"
    # 其他页无背景
    assert prs.slides[0]._element.cSld.find(qn("p:bg")) is None


def test_slide_background_gradient(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "slide_background", "slide_index": 0,
          "gradient": {"from": "0F172A", "to": "2563EB", "angle": 45}}],
    )
    from pptx.oxml.ns import qn

    bg = _bg_fill_of(Presentation(str(dst)), 0)
    grad = bg.find(".//" + qn("a:gradFill"))
    assert grad is not None
    stops = grad.findall(".//" + qn("a:srgbClr"))
    assert [s.get("val") for s in stops] == ["0F172A", "2563EB"]


def test_slide_background_overwrites_existing(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "slide_background", "slide_index": 0, "color": "111111"}]
    )
    with pytest.raises(edits.PatchError, match="color"):
        edits.apply_patches_to_file(
            dst, tmp_path / "o.pptx",
            [{"action": "slide_background", "slide_index": 0, "color": "nothex"}],
        )
    # 重新设置成功覆盖，不会出现两个 bg
    dst2 = tmp_path / "out2.pptx"
    edits.apply_patches_to_file(
        dst, dst2, [{"action": "slide_background", "slide_index": 0, "color": "222222"}]
    )
    from pptx.oxml.ns import qn

    cSld = Presentation(str(dst2)).slides[0]._element.cSld
    assert len(cSld.findall(qn("p:bg"))) == 1


def test_slide_background_requires_slide_index(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="slide_index"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "slide_background", "color": "112233"}],
        )


# ---------- 9. add_page_numbers ----------


def test_add_page_numbers_all_pages(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "add_page_numbers"}]
    )
    assert touched == {0, 1, 2}
    prs = Presentation(str(dst))
    for index, slide in enumerate(prs.slides):
        numbers = [
            s.text for s in slide.shapes
            if getattr(s, "has_text_frame", False) and s.text.strip() == str(index + 1)
        ]
        assert numbers, f"slide {index} missing page number"


def test_add_page_numbers_skip_first(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "add_page_numbers", "skip_first": True}]
    )
    prs = Presentation(str(dst))
    texts0 = [s.text.strip() for s in prs.slides[0].shapes if getattr(s, "has_text_frame", False)]
    assert "1" not in texts0
    texts2 = [s.text.strip() for s in prs.slides[2].shapes if getattr(s, "has_text_frame", False)]
    assert "3" in texts2  # 页码仍是真实页码


def test_add_page_numbers_position_bottom_left(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "add_page_numbers", "position": "bottom-left", "start": 1}],
    )
    prs = Presentation(str(dst))
    prs.slide_width  # 12192000
    num = next(
        s for s in prs.slides[1].shapes
        if getattr(s, "has_text_frame", False) and s.text.strip() == "2"
    )
    assert num.left < 12192000 / 2  # 靠左


def test_add_page_numbers_twice_rejected(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(demo_pptx, dst, [{"action": "add_page_numbers"}])
    with pytest.raises(edits.PatchError, match="页码"):
        edits.apply_patches_to_file(dst, tmp_path / "o.pptx", [{"action": "add_page_numbers"}])


# ---------- 10. font_scale ----------


def test_font_scale_deck(mixed_font_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        mixed_font_pptx, dst, [{"action": "font_scale", "delta_pt": 2}]
    )
    assert touched == {0, 1}
    prs = Presentation(str(dst))
    sizes = {run.text: run.font.size.pt for run in _all_runs(prs) if run.text}
    assert sizes["标题宋体"] == 34
    assert sizes["正文黑体"] == 20
    assert sizes["表头楷体"] == 16
    assert sizes["第二页仿宋"] == 26


def test_font_scale_negative_and_slide_scope(mixed_font_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        mixed_font_pptx, dst, [{"action": "font_scale", "delta_pt": -2, "slide_index": 0}]
    )
    prs = Presentation(str(dst))
    sizes = {run.text: run.font.size.pt for run in _all_runs(prs) if run.text}
    assert sizes["标题宋体"] == 30
    assert sizes["第二页仿宋"] == 24  # 范围外不动


def test_font_scale_inherited_size_gets_explicit(mixed_font_pptx, tmp_path):
    """无显式字号的 run 用默认 18pt 起算，缩放后有显式字号。"""
    prs = Presentation(str(mixed_font_pptx))
    tb = prs.slides[0].shapes.add_textbox(Emu(500000), Emu(4000000), Emu(5000000), Emu(800000))
    tb.text_frame.text = "无字号文本"
    src = tmp_path / "nosize.pptx"
    prs.save(str(src))
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(src, dst, [{"action": "font_scale", "delta_pt": 4, "slide_index": 0}])
    sizes = {run.text: run.font.size.pt for run in _all_runs(Presentation(str(dst))) if run.text}
    assert sizes["无字号文本"] == 22


def test_font_scale_validation(mixed_font_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="delta"):
        edits.apply_patches_to_file(
            mixed_font_pptx, tmp_path / "a.pptx", [{"action": "font_scale"}]
        )
    with pytest.raises(edits.PatchError, match="delta"):
        edits.apply_patches_to_file(
            mixed_font_pptx, tmp_path / "b.pptx", [{"action": "font_scale", "delta_pt": 500}]
        )


# ---------- agent 操作存储（_OperationStore）----------

def _view_of(path, scope_slides=None):
    _objects, payload = parse.parse_pptx(str(path))
    total = payload["slide_count"]
    if scope_slides is None:
        scope_slides = list(range(total))
    return agent._DeckView(
        payload["objects"], payload["elements"], payload["tables"], total,
        scope_slides, payload["slide_size_emu"],
    )


def _store_of(path, instruction="整理一下", scope_slides=None):
    view = _view_of(path, scope_slides)
    return agent._OperationStore(view, instruction, 24)


def _accept(store, ops):
    batch = ops if isinstance(ops, list) else [ops]
    result = store.add_batch(batch)
    assert result["accepted"] == len(batch) >= 1, result
    assert len(result["accepted_operations"]) == len(batch), result
    return result


def _drop_reason(store, ops):
    result = store.add_batch(ops if isinstance(ops, list) else [ops])
    assert result["dropped"] >= 1 and result["accepted"] == 0, result
    return result["reasons"][0]["reason"]


def test_agent_font_unify_deck_scope(demo_pptx):
    store = _store_of(demo_pptx, "把整份字体统一")
    result = _accept(store, [{"action": "font_unify", "font_name": "微软雅黑"}])
    op = result["accepted_operations"][0]
    assert op["font_name"] == "微软雅黑"


def test_agent_font_unify_slide_scope_needs_index_in_scope(demo_pptx):
    store = _store_of(demo_pptx, scope_slides=[0])  # 当前页范围
    # v0.3 复核后：当前页范围 + 未写范围 → 默认只改这一页（不再要求整份）
    result = _accept(store, [{"action": "font_unify", "font_name": "微软雅黑"}])
    op = result["accepted_operations"][0]
    assert op["scope"] == "slide" and op["slide_index"] == 0
    # 明确要整份、但当前范围只有一页 → 拒绝并提示怎么切换
    reason = _drop_reason(store, [{"action": "font_unify", "font_name": "微软雅黑", "scope": "deck"}])
    assert "整份" in reason
    # 显式 slide_index 指到范围内页面仍然可用（与默认等价，幂等去重由 store 保证）
    store2 = _store_of(demo_pptx, scope_slides=[0])
    result = _accept(store2, [{"action": "font_unify", "font_name": "微软雅黑", "slide_index": 0}])
    assert result["accepted_operations"][0]["slide_index"] == 0


def test_agent_replace_text_validation(demo_pptx):
    store = _store_of(demo_pptx, "全局替换")
    result = _accept(store, [{"action": "replace_text", "find": "项目", "replace": "工程"}])
    op = result["accepted_operations"][0]
    assert op["find"] == "项目" and op["replace"] == "工程"
    reason = _drop_reason(store, [{"action": "replace_text", "find": "", "replace": "x"}])
    assert "find" in reason


def test_agent_align_objects_accepts_and_drops(demo_pptx):
    store = _store_of(demo_pptx, "把两个文本框左对齐")
    result = _accept(store, [
        {"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3"], "align": "left"},
    ])
    assert result["accepted_operations"][0]["shape_ids"] == ["s0_id2", "s0_id3"]
    reason = _drop_reason(store, [
        {"action": "align_objects", "shape_ids": ["s0_id2", "s0_id3"], "align": "diagonal"},
    ])
    assert "align" in reason or "mode" in reason


def test_agent_duplicate_and_z_order(demo_pptx):
    store = _store_of(demo_pptx, "复制一个文本框并置顶")
    result = _accept(store, [
        {"action": "duplicate_object", "shape_id": "s0_id2", "offset_emu": {"x": 300000, "y": 300000}},
        {"action": "z_order", "shape_id": "s0_id2", "mode": "front"},
    ])
    actions = [op["action"] for op in result["accepted_operations"]]
    assert actions == ["duplicate_object", "z_order"]
    reason = _drop_reason(store, [{"action": "z_order", "shape_id": "s0_id2", "mode": "wat"}])
    assert "mode" in reason


def test_agent_table_mod(demo_pptx):
    store = _store_of(demo_pptx, "表格加一行")
    table_id = _view_of(demo_pptx).tables[0]["shape_id"]
    result = _accept(store, [{"action": "table_mod", "shape_id": table_id, "op": "add_row"}])
    assert result["accepted_operations"][0]["shape_id"] == table_id
    reason = _drop_reason(store, [{"action": "table_mod", "shape_id": "s0_id2", "op": "add_row"}])
    assert "表格" in reason


def test_agent_slide_background(demo_pptx):
    store = _store_of(demo_pptx, "第 1 页换个背景")
    result = _accept(store, [{"action": "slide_background", "slide_index": 0, "color": "1F4E79"}])
    assert result["accepted_operations"][0]["color"] == "1F4E79"
    reason = _drop_reason(store, [{"action": "slide_background", "slide_index": 0, "color": "nothex"}])
    assert "color" in reason


def test_agent_page_numbers_and_font_scale(demo_pptx):
    store = _store_of(demo_pptx, "加页码并字号统一加两磅")
    result = _accept(store, [
        {"action": "add_page_numbers", "skip_first": True},
        {"action": "font_scale", "delta_pt": 2},
    ])
    actions = [op["action"] for op in result["accepted_operations"]]
    assert actions == ["add_page_numbers", "font_scale"]
    reason = _drop_reason(store, [{"action": "font_scale", "delta_pt": 999}])
    assert "delta" in reason


def test_agent_replace_image_rejected_in_chat(demo_pptx):
    store = _store_of(demo_pptx, "换掉图片")
    reason = _drop_reason(store, [{"action": "replace_image", "shape_id": "s0_id5"}])
    assert "图片" in reason or "image" in reason
