"""v0.6 #2：OOXML passthrough 保真 —— part 级原字节保留。

python-pptx 整包重序列化会动到它没映射的 XML（动画 timing、自定义扩展元素
的属性顺序/命名空间前缀）。apply 后把「未被触碰的 slide part」用原包字节
原样写回，改一处不再丢别处。探针：在 slide XML 里注入动画/自定义元素，
编辑另一页后验证本页字节完全一致。
"""
from __future__ import annotations

import re
import zipfile

from services import edits, passthrough

# 注入到未编辑页 slide XML 的探针：动画 timing + 自定义扩展元素
_PROBE = (
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never"'
    ' nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
    '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"/></p:seq>'
    "</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"
)


def _read_part(path, zip_name: str) -> bytes:
    with zipfile.ZipFile(path) as zf:
        return zf.read(zip_name)


def _slide_part_names(path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        return sorted(n for n in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))


def _inject_probe(demo_pptx, tmp_path, slide_zip_offset: int = 1):
    """把探针 XML 注入第 (offset+1) 张 slide part（在 </p:sld> 前插入）。"""
    target = tmp_path / "probed.pptx"
    names = _slide_part_names(demo_pptx)
    victim = names[slide_zip_offset]
    with zipfile.ZipFile(demo_pptx) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == victim:
                text = data.decode("utf-8")
                assert "</p:sld>" in text
                text = text.replace("</p:sld>", _PROBE + "</p:sld>")
                data = text.encode("utf-8")
            zout.writestr(item, data)
    return target, victim


def test_untouched_slide_part_bytes_identical(demo_pptx, tmp_path):
    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    original_bytes = _read_part(src, victim)
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        src, dst,
        [{"action": "edit", "shape_id": "s0_id2", "after": "改过的标题"}],
    )
    assert _read_part(dst, victim) == original_bytes
    # 而被编辑的页确实变了
    from pptx import Presentation

    texts = []
    for shape in Presentation(str(dst)).slides[0].shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            texts.append(shape.text_frame.text)
    assert any("改过的标题" in t for t in texts)


def test_touched_slide_part_not_restored(demo_pptx, tmp_path):
    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    dst = tmp_path / "out.pptx"
    # 编辑 slide 1（victim 是第二张 slide 的 part；conftest 中 slide 索引 1 是表格页）
    edits.apply_patches_to_file(
        src, dst,
        [{"action": "edit", "shape_id": "s1_id2_c0_0", "after": "新的标题文字"}],
    )
    # 被触碰的页走 python-pptx 新序列化（文字已改 → 字节必变），不回填
    assert _read_part(dst, victim) != _read_part(src, victim)
    # 语义上探针内容仍在（round-trip 不丢）
    assert b'nodeType="tmRoot"' in _read_part(dst, victim) or b"tmRoot" in _read_part(dst, victim)


def test_deck_op_touches_all_prevents_restore(demo_pptx, tmp_path):
    """deck 级 font_scale 触碰所有页 → 全部走新序列化，不回填。"""
    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(src, dst, [{"action": "font_scale", "delta_pt": 2}])
    assert _read_part(dst, victim) != _read_part(src, victim)


def test_merge_keeps_existing_slide_bytes(demo_pptx, tmp_path):
    """merge：目标文稿已有页未触碰 → 原字节；新拉入页为新 part 不受影响。"""
    from pptx import Presentation
    from pptx.util import Emu

    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    other = tmp_path / "other.pptx"
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(1000000), Emu(500000), Emu(6000000), Emu(1000000))
    box.text_frame.text = "外来页"
    prs.save(str(other))

    merged = tmp_path / "merged.pptx"
    count, _touched = edits.merge_slides_from_file(src, other, merged, "1", None)
    assert count == 1
    assert len(Presentation(str(merged)).slides._sldIdLst) == 4
    assert _read_part(merged, victim) == _read_part(src, victim)


def test_passthrough_disabled_by_env(demo_pptx, tmp_path, monkeypatch):
    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    monkeypatch.setenv("PPTMO_PASSTHROUGH", "0")
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        src, dst,
        [{"action": "edit", "shape_id": "s0_id2", "after": "再改一次"}],
    )
    # 关闭后走 python-pptx 序列化（此处仅验证不炸 + 编辑生效）
    from pptx import Presentation

    slide0_texts = [s.text_frame.text for s in Presentation(str(dst)).slides[0].shapes if s.has_text_frame]
    assert any("再改一次" in t for t in slide0_texts)


def test_preserve_function_direct(demo_pptx, tmp_path):
    """直接调用 preserve_untouched_slides：只还原未触碰页的 part 字节。"""
    import shutil

    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    first = _slide_part_names(src)[0]
    dst = tmp_path / "work.pptx"
    edits.apply_patches_to_file(
        src, dst,
        [{"action": "edit", "shape_id": "s0_id2", "after": "编辑页0"}],
    )
    # 人为破坏 passthrough（模拟关闭），再手动还原 untouched={1}
    monkey_src = tmp_path / "src-copy.pptx"
    shutil.copy2(src, monkey_src)
    restored = passthrough.preserve_untouched_slides(monkey_src, dst, touched={0})
    assert restored >= 1
    assert _read_part(dst, victim) == _read_part(src, victim)
    assert _read_part(dst, first) != _read_part(src, first)  # 页0被编辑过，不还原
    # 还原后的文件仍可被 python-pptx 打开
    from pptx import Presentation

    Presentation(str(dst))


def test_edit_notes_survives_passthrough(demo_pptx, tmp_path):
    """edit_notes 只改 notesSlide part——该页必须视为已触碰，否则批量模式下
    只有最后一个 slide_index 进 touched，前面页的 notes 被原字节冲掉。"""
    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        src, dst,
        [{"action": "edit_notes", "items": [
            {"slide_index": 0, "text": "开场"},          # 已有 notes part → 会被还原冲掉
            {"slide_index": 2, "text": "结尾"},          # 无 notes part → 新建，不受影响
        ]}],
    )
    from pptx import Presentation

    prs = Presentation(str(dst))
    assert prs.slides[0].notes_slide.notes_text_frame.text == "开场"
    assert prs.slides[2].notes_slide.notes_text_frame.text == "结尾"
    # 未触碰页（探针页，原始索引 1）字节仍保真
    assert _read_part(dst, victim) == _read_part(src, victim)


def test_preserve_after_slide_reorder(demo_pptx, tmp_path):
    """move_slide 后：按 part 名映射仍把未编辑页还原成原字节。"""
    src, victim = _inject_probe(demo_pptx, tmp_path, slide_zip_offset=1)
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        src, dst,
        [{"action": "move_slide", "slide": 2, "to": 0}],
    )
    from pptx import Presentation

    order_texts = []
    for slide in Presentation(str(dst)).slides:
        first_text = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                first_text = shape.text_frame.text.strip()
                break
        order_texts.append(first_text)
    # 页序已变（被移动的组页「目录」排到最前）
    assert "目录" in order_texts[0]
    # 未编辑页的字节保真（探针页）
    assert _read_part(dst, victim) == _read_part(src, victim)
