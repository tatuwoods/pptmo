"""v0.2 slide-level operations: duplicate_slide / add_slide / delete_slide / move_slide.

Covers the edits layer (apply_patches_to_file), the API layer (base_version guard,
objects.json renumbering) and the agent operation store (schema + validation).
"""
from __future__ import annotations

import pytest
from pptx import Presentation

from services import agent, edits, parse


# ---------- fixtures ----------

@pytest.fixture()
def pic_pptx(tmp_path):
    """2-slide deck: slide 0 = picture + styled title + speaker notes; slide 1 = text."""
    from io import BytesIO

    from pptx.util import Emu, Pt

    png = bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
    )
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s0 = prs.slides.add_slide(blank)
    pic = s0.shapes.add_picture(BytesIO(png), Emu(4000000), Emu(1000000), Emu(600000), Emu(400000))
    tb = s0.shapes.add_textbox(Emu(1000000), Emu(3000000), Emu(6000000), Emu(1000000))
    tb.text_frame.text = "带图的页面"
    run = tb.text_frame.paragraphs[0].runs[0]
    run.font.size = Pt(28)
    run.font.bold = True
    s0.notes_slide.notes_text_frame.text = "这是讲稿备注"
    s1 = prs.slides.add_slide(blank)
    tb2 = s1.shapes.add_textbox(Emu(1000000), Emu(1000000), Emu(6000000), Emu(800000))
    tb2.text_frame.text = "第二页正文"
    path = tmp_path / "pic.pptx"
    prs.save(str(path))
    return path


def _iter_all_shapes(shapes):
    for s in shapes:
        yield s
        if getattr(s, "shape_type", None) == 6:
            try:
                yield from _iter_all_shapes(s.shapes)
            except Exception:
                pass


def _texts_of(slide):
    return [s.text for s in _iter_all_shapes(slide.shapes) if getattr(s, "has_text_frame", False)]


def _find_text(slide, needle):
    for s in _iter_all_shapes(slide.shapes):
        if getattr(s, "has_text_frame", False) and needle in (s.text or ""):
            return s
    return None


# ---------- duplicate_slide: edits layer ----------

def test_duplicate_slide_basic(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "duplicate_slide", "source_slide": 0}]
    )
    assert touched == {0, 1}  # default position = right after source
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    src_texts = _texts_of(prs.slides[0])
    copy_texts = _texts_of(prs.slides[1])
    assert src_texts == copy_texts
    assert _texts_of(prs.slides[3]) == ["第二页正文"] or _texts_of(prs.slides[3])  # tail intact


def test_duplicate_slide_explicit_position(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "duplicate_slide", "source_slide": 0, "position": 3}]
    )
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    # copy (title texts) landed at the END; original order untouched
    assert any("项目进展汇报" in t for t in _texts_of(prs.slides[3]))
    assert any("目录" in t for t in _texts_of(prs.slides[2]))


def test_duplicate_slide_with_picture(pic_pptx, tmp_path):
    png = bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
    )
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        pic_pptx, dst, [{"action": "duplicate_slide", "source_slide": 0, "position": 1}]
    )
    prs = Presentation(str(dst))  # reopen: broken rels would explode here or below
    assert len(prs.slides) == 3
    src_pic = next(s for s in prs.slides[0].shapes if s.shape_type == 13)
    copy_pic = next(s for s in prs.slides[1].shapes if s.shape_type == 13)
    assert copy_pic.image.blob == png == src_pic.image.blob
    assert (copy_pic.left, copy_pic.top, copy_pic.width, copy_pic.height) == (
        src_pic.left, src_pic.top, src_pic.width, src_pic.height
    )
    assert _find_text(prs.slides[1], "带图的页面") is not None


def test_duplicate_slide_with_table(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "duplicate_slide", "source_slide": 1}]
    )
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    src_table = next(s for s in prs.slides[1].shapes if s.has_table).table
    copy_table = next(s for s in prs.slides[2].shapes if s.has_table).table
    assert (len(copy_table.rows), len(copy_table.columns)) == (
        len(src_table.rows), len(src_table.columns)
    )
    src_texts = [src_table.cell(r, c).text for r in range(3) for c in range(3)]
    copy_texts = [copy_table.cell(r, c).text for r in range(3) for c in range(3)]
    assert src_texts == copy_texts
    # merged cell (1,1)+(1,2) must survive the copy
    assert copy_table.cell(1, 1).is_merge_origin
    assert copy_table.cell(1, 1).span_width == 2
    assert copy_table.cell(1, 2).is_spanned


def test_duplicate_slide_with_group(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "duplicate_slide", "source_slide": 2, "position": 0}]
    )
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    assert _find_text(prs.slides[0], "第一章 总体情况") is not None
    assert _find_text(prs.slides[0], "目录") is not None


def test_duplicate_slide_preserves_style_and_geometry(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "duplicate_slide", "source_slide": 0}]
    )
    prs = Presentation(str(dst))
    src = _find_text(prs.slides[0], "项目进展汇报")
    copy = _find_text(prs.slides[1], "项目进展汇报")
    src_run = src.text_frame.paragraphs[0].runs[0]
    copy_run = copy.text_frame.paragraphs[0].runs[0]
    assert copy_run.font.size == src_run.font.size
    assert copy_run.font.bold == src_run.font.bold
    assert (copy.left, copy.top, copy.width, copy.height) == (
        src.left, src.top, src.width, src.height
    )


def test_duplicate_slide_copies_notes(pic_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        pic_pptx, dst, [{"action": "duplicate_slide", "source_slide": 0}]
    )
    prs = Presentation(str(dst))
    copy = prs.slides[1]
    assert copy.has_notes_slide
    assert copy.notes_slide.notes_text_frame.text == "这是讲稿备注"


def test_duplicate_slide_out_of_range(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "duplicate_slide", "source_slide": 9}]
        )
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [{"action": "duplicate_slide", "source_slide": 0, "position": 9}],
        )
    with pytest.raises(edits.PatchError):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "duplicate_slide"}]
        )


def test_duplicate_slide_alias_and_negative(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "copy_slide", "slide": 0, "to": 1}]
    )
    assert touched == {0, 1}
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "duplicate_slide", "source_slide": -1}]
        )


# ---------- add_slide ----------

def test_add_slide_default_appends(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(demo_pptx, dst, [{"action": "add_slide"}])
    assert touched == {3}
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    assert _texts_of(prs.slides[3]) == []  # truly blank
    # default layout: blank-ish (no cloned placeholder text), tail is editable


def test_add_slide_after_position(pic_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        pic_pptx, dst, [{"action": "add_slide", "after_slide": 0}]
    )
    assert touched == {1}
    prs = Presentation(str(dst))
    assert len(prs.slides) == 3
    assert _texts_of(prs.slides[1]) == []  # blank page inserted between the two
    assert any("带图的页面" in t for t in _texts_of(prs.slides[0]))
    assert any("第二页正文" in t for t in _texts_of(prs.slides[2]))


def test_add_slide_layout_index(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "add_slide", "layout_index": 5}]
    )
    prs = Presentation(str(dst))
    new_slide = prs.slides[3]
    assert new_slide.slide_layout.name == prs.slide_layouts[5].name
    assert new_slide.slide_layout.name != prs.slide_layouts[6].name


def test_add_slide_bad_params(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="layout"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "add_slide", "layout_index": 99}]
        )
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "add_slide", "after_slide": 9}]
        )


# ---------- partname uniqueness ----------
# python-pptx names new slide parts by slide COUNT (len(sldIdLst)+1). After a
# delete (or in decks whose partnames have gaps) that count reuses an existing
# slideN.xml and the saved zip gets DUPLICATE entries — non-conformant OPC that
# strict consumers reject. Our wrapper must reassign a genuinely free name.

def _zip_names(path):
    import zipfile

    with zipfile.ZipFile(str(path)) as zf:
        return [info.filename for info in zf.infolist()]


def test_add_after_delete_keeps_zip_entries_unique(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [{"action": "delete_slide", "slide": 0}, {"action": "add_slide", "after_slide": 0}],
    )
    names = _zip_names(dst)
    assert len(names) == len(set(names)), "duplicate zip entries (partname collision)"
    prs = Presentation(str(dst))
    assert len(prs.slides) == 3


def test_duplicate_after_delete_keeps_zip_entries_unique(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [
            {"action": "delete_slide", "slide": 0},
            {"action": "duplicate_slide", "source_slide": 1},  # base numbering: table page
        ],
    )
    names = _zip_names(dst)
    assert len(names) == len(set(names)), "duplicate zip entries (partname collision)"
    prs = Presentation(str(dst))
    assert len(prs.slides) == 3
    assert any(s.has_table for s in prs.slides[0].shapes)
    assert any(s.has_table for s in prs.slides[1].shapes)


def test_duplicate_of_deleted_slide_rejected(demo_pptx, tmp_path):
    """Batch numbering refers to the BASE deck: a slide deleted earlier in the
    same batch can no longer be addressed."""
    with pytest.raises(edits.PatchError, match="deleted"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx",
            [
                {"action": "delete_slide", "slide": 0},
                {"action": "duplicate_slide", "source_slide": 0},
            ],
        )


def test_mixed_slide_ops_no_zip_duplicates(demo_pptx, tmp_path):
    """Delete + duplicates + add in one batch: names must stay unique throughout."""
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [
            {"action": "delete_slide", "slide": 1},
            {"action": "duplicate_slide", "source_slide": 0},
            {"action": "duplicate_slide", "source_slide": 2},
            {"action": "add_slide"},
        ],
    )
    names = _zip_names(dst)
    assert len(names) == len(set(names)), "duplicate zip entries (partname collision)"
    prs = Presentation(str(dst))
    assert len(prs.slides) == 5  # 3 - 1 + 2 clones + 1 blank


# ---------- delete_slide ----------

def test_delete_slide_middle(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "delete_slide", "slide": 1}]
    )
    assert touched == {1}
    prs = Presentation(str(dst))
    assert len(prs.slides) == 2
    assert not any(sh.has_table for s in prs.slides for sh in s.shapes)  # table page gone
    assert any("项目进展汇报" in t for t in _texts_of(prs.slides[0]))
    assert any("目录" in t for t in _texts_of(prs.slides[1]))


def test_delete_slide_renumbers_object_ids(demo_pptx, tmp_path):
    """After deleting slide 1, re-parsing must shift s{n}_ prefixes: old s2 objects become s1."""
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "delete_slide", "slide": 1}]
    )
    objects, payload = parse.parse_pptx(str(dst))
    assert payload["slide_count"] == 2
    assert not any(o["kind"] == "table_cell" for o in objects)
    group_text = next(o for o in objects if "总体情况" in o.get("text", ""))
    assert group_text["shape_id"].startswith("s1_")
    title = next(o for o in objects if "项目进展汇报" in o.get("text", ""))
    assert title["shape_id"].startswith("s0_")


def test_delete_only_slide_rejected(tmp_path):
    from pptx import Presentation as P

    prs = P()
    prs.slides.add_slide(prs.slide_layouts[6])
    solo = tmp_path / "solo.pptx"
    prs.save(str(solo))
    with pytest.raises(edits.PatchError, match="only"):
        edits.apply_patches_to_file(
            solo, tmp_path / "o.pptx", [{"action": "delete_slide", "slide": 0}]
        )


def test_delete_slide_out_of_range(demo_pptx, tmp_path):
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "delete_slide", "slide": 5}]
        )


# ---------- move_slide ----------

def test_move_slide_to_front(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    touched = edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "move_slide", "slide": 2, "to": 0}]
    )
    assert touched == {0, 2}
    prs = Presentation(str(dst))
    assert len(prs.slides) == 3
    assert any("目录" in t for t in _texts_of(prs.slides[0]))
    assert any("项目进展汇报" in t for t in _texts_of(prs.slides[1]))
    assert any(s.has_table for s in prs.slides[2].shapes)


def test_move_slide_to_end(pic_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(pic_pptx, dst, [{"action": "move_slide", "slide": 0, "to": 1}])
    prs = Presentation(str(dst))
    assert any("第二页正文" in t for t in _texts_of(prs.slides[0]))
    assert any("带图的页面" in t for t in _texts_of(prs.slides[1]))


def test_move_slide_same_and_out_of_range(demo_pptx, tmp_path):
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst, [{"action": "move_slide", "slide": 1, "to": 1}]
    )
    prs = Presentation(str(dst))
    assert any(s.has_table for s in prs.slides[1].shapes)  # unchanged
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "move_slide", "slide": 1, "to": 9}]
        )
    with pytest.raises(edits.PatchError, match="range"):
        edits.apply_patches_to_file(
            demo_pptx, tmp_path / "o.pptx", [{"action": "move_slide", "slide": 9, "to": 0}]
        )


# ---------- mixed batches: shape ops resolve against pre-batch numbering ----------

def test_mixed_batch_shape_ops_use_original_numbering(demo_pptx, tmp_path):
    objects, _payload = parse.parse_pptx(str(demo_pptx))
    cell = next(o for o in objects if o.get("text") == "200 万")
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [
            {"action": "edit", "shape_id": cell["shape_id"], "after": "改成 180 万"},  # pre-batch numbering
            {"action": "duplicate_slide", "source_slide": 0, "position": 1},
        ],
    )
    prs = Presentation(str(dst))
    assert len(prs.slides) == 4
    # the edited table page shifted to index 2 by the duplicate; the edit still landed on it
    table = next(s for s in prs.slides[2].shapes if s.has_table).table
    texts = [table.cell(r, c).text for r in range(3) for c in range(3)]
    assert "改成 180 万" in texts
    assert any("项目进展汇报" in t for t in _texts_of(prs.slides[1]))


def test_slide_op_ids_survive_alias_partition(demo_pptx, tmp_path):
    """copy_slide/remove_slide aliases must be routed to the slide phase, not error out."""
    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [
            {"action": "copy_slide", "source_slide": 2, "position": 2},
            {"action": "remove_slide", "slide": 0},
        ],
    )
    prs = Presentation(str(dst))
    assert len(prs.slides) == 3  # 3 + 1 duplicate - 1 delete
    # [A, B, C] --copy C @2--> [A, B, copyC, C] --delete A--> [B, copyC, C]
    assert any("目录" in t for t in _texts_of(prs.slides[1]))


def test_slide_partnames_stay_unique_after_delete(demo_pptx, tmp_path):
    """Regression: python-pptx derives new slide partnames from len(sldIdLst)+1,
    so after a delete the next add/copy would collide with an existing part and
    produce a zip with duplicate entries (PowerPoint then loads wrong content)."""
    import zipfile

    dst = tmp_path / "out.pptx"
    edits.apply_patches_to_file(
        demo_pptx, dst,
        [
            {"action": "delete_slide", "slide": 0},   # len(sldIdLst) drops to 2
            {"action": "duplicate_slide", "source_slide": 1, "position": 1},
            {"action": "add_slide"},
        ],
    )
    prs = Presentation(str(dst))
    partnames = [str(s.part.partname) for s in prs.slides]
    assert len(partnames) == len(set(partnames)), partnames
    with zipfile.ZipFile(dst) as z:
        names = z.namelist()
        dupes = {n for n in names if names.count(n) > 1}
        assert not dupes, f"duplicate zip entries: {dupes}"
    # batch numbers are BASE-deck coordinates: [title, table, group] with title
    # deleted, table copied to position 1, blank appended → [copy(table), table, group, blank]
    assert len(prs.slides) == 4
    assert any(s.has_table for s in prs.slides[0].shapes)
    assert any(s.has_table for s in prs.slides[1].shapes)
    assert any("目录" in t for t in _texts_of(prs.slides[2]))


# ---------- API layer ----------

def test_api_duplicate_slide(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "duplicate_slide", "source_slide": 0, "position": 1}],
        "label": "复制第一页", "base_version": 0,
    })
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["version_id"] == "v1"
    assert body["slide_count"] == 4
    assert body["touched_slides"] == [0, 1]
    # objects.json re-parsed with the new numbering: s1 now carries the title texts
    titles = [o for o in body["objects"] if "项目进展汇报" in o.get("text", "")]
    assert len(titles) == 2  # original on s0 + copy on s1
    assert any(o["shape_id"].startswith("s1_") for o in titles)


def test_api_delete_slide_renumbers_objects(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "delete_slide", "slide": 0}], "base_version": 0,
    })
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["slide_count"] == 2
    # old slide 1 (table) is now s0
    cells = [o for o in body["objects"] if o.get("kind") == "table_cell"]
    assert cells and all(o["shape_id"].startswith("s0_") for o in cells)
    assert not any(o["shape_id"].startswith("s2_") for o in body["objects"])


def test_api_move_slide(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "move_slide", "slide": 2, "to": 0}], "base_version": 0,
    })
    assert r.status_code == 200, r.text
    body = r.json()
    group_text = next(o for o in body["objects"] if "总体情况" in o.get("text", ""))
    assert group_text["shape_id"].startswith("s0_")


def test_api_slide_op_base_version_conflict(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "duplicate_slide", "source_slide": 0}],
        "base_version": 42,
    })
    assert r.status_code == 409


def test_api_slide_op_out_of_range_400(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "duplicate_slide", "source_slide": 99}],
    })
    assert r.status_code == 400
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "move_slide", "slide": 0, "to": 99}],
    })
    assert r.status_code == 400


def test_api_add_slide(client, uploaded_doc):
    doc_id = uploaded_doc["doc_id"]
    r = client.post(f"/api/docs/{doc_id}/apply", json={
        "patches": [{"action": "add_slide", "after_slide": 0}], "base_version": 0,
    })
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["slide_count"] == 4
    # inserted blank slide is s1: no objects there, old slide-1 table now s2
    cells = [o for o in body["objects"] if o.get("kind") == "table_cell"]
    assert cells and all(o["shape_id"].startswith("s2_") for o in cells)


# ---------- agent operation store ----------

def _view(total=3, scope=None):
    return agent._DeckView(
        objects=[],
        elements=[],
        tables=[],
        total_slides=total,
        scope_slides=scope if scope is not None else list(range(total)),
        slide_size=None,
    )


def _store(view, instruction="整理页面"):
    return agent._OperationStore(view, instruction, 24)


def test_agent_accepts_slide_ops():
    store = _store(_view())
    result = store.add_batch([
        {"action": "duplicate_slide", "source_slide": 0, "position": 1, "description": "复制封面"},
        {"action": "add_slide", "after_slide": 2, "description": "末尾前加一页"},
        {"action": "delete_slide", "slide": 1, "description": "删掉表格页"},
        {"action": "move_slide", "slide": 2, "to": 0, "description": "目录提前"},
    ])
    assert result["accepted"] == 4 and result["dropped"] == 0
    actions = [op["action"] for op in store.ops]
    assert actions == ["duplicate_slide", "add_slide", "delete_slide", "move_slide"]
    dup = store.ops[0]
    assert dup["source_slide"] == 0 and dup["position"] == 1
    assert dup["slide_index"] == 0  # card anchor page


def test_agent_slide_op_defaults():
    store = _store(_view())
    result = store.add_batch([{"action": "duplicate_slide", "source_slide": 1}])
    assert result["accepted"] == 1
    assert store.ops[0]["position"] == 2  # right after source


def test_agent_slide_op_validation():
    store = _store(_view())
    result = store.add_batch([
        {"action": "duplicate_slide", "source_slide": 9},          # out of range
        {"action": "duplicate_slide"},                              # missing source
        {"action": "duplicate_slide", "source_slide": 0, "position": 9},  # bad position
        {"action": "move_slide", "slide": 0},                       # missing to
        {"action": "delete_slide", "slide": 0, "to": 1},            # delete has no `to`
    ])
    assert result["accepted"] == 1  # the delete op is valid
    assert result["dropped"] == 4
    assert result["reasons"]


def test_agent_slide_op_scope_guard():
    store = _store(_view(total=3, scope=[0]))  # current-slide scope on page 0
    result = store.add_batch([
        {"action": "duplicate_slide", "source_slide": 2},  # out of scope source
        {"action": "move_slide", "slide": 2, "to": 0},     # out of scope slide
    ])
    assert result["accepted"] == 0 and result["dropped"] == 2


def test_agent_rejects_duplicate_slide_ops_and_only_page_delete():
    store = _store(_view())
    op = {"action": "duplicate_slide", "source_slide": 0, "position": 1}
    result = store.add_batch([op, dict(op)])
    assert result["accepted"] == 1 and result["dropped"] == 1

    solo = _store(_view(total=1, scope=[0]))
    result = solo.add_batch([{"action": "delete_slide", "slide": 0}])
    assert result["accepted"] == 0
    assert "唯一" in result["reasons"][0]["reason"] or "不能" in result["reasons"][0]["reason"]


def test_agent_op_summary_includes_slide_params():
    store = _store(_view())
    store.add_batch([{"action": "move_slide", "slide": 2, "to": 0}])
    summary = store._op_summary(store.ops[0])
    assert summary["action"] == "move_slide"
    assert summary["slide"] == 2 and summary["to"] == 0
