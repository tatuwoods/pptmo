"""v0.5 #8：结构 diff 增强 — 对象级结构化 changelog（JSON 给 AI 消费）。

对标 hands-on-deck diff：不只文字 diff，还输出新增/删除/移动了哪些对象、
几何变化（哪个轴）、样式变化（字号/加粗/颜色/字体）——按页位对齐 + shape_id
匹配。页间移动的对象无法可靠关联，如实报告为 removed + added。
"""
from __future__ import annotations

import io
import json

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Emu, Pt

from services import diff

EMU_PT = 12700


def _textbox(slide, text, x=Emu(1000000), y=Emu(1000000), w=Emu(4000000), h=Emu(800000)):
    box = slide.shapes.add_textbox(x, y, w, h)
    box.text_frame.text = text
    return box


def _build(spec) -> bytes:
    """spec: [[(text, x, y, w, h, size_pt, bold) | dict overrides], ...] 每项一页。"""
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    blank = prs.slide_layouts[6]
    for entries in spec:
        slide = prs.slides.add_slide(blank)
        for entry in entries:
            if isinstance(entry, dict):
                box = _textbox(slide, entry["text"])
                run = box.text_frame.paragraphs[0].runs[0]
                if entry.get("size"):
                    run.font.size = Pt(entry["size"])
                if entry.get("bold") is not None:
                    run.font.bold = entry["bold"]
                if entry.get("color"):
                    run.font.color.rgb = RGBColor.from_string(entry["color"])
                if entry.get("x") is not None:
                    box.left = Emu(entry["x"])
                    box.top = Emu(entry.get("y", 1000000))
                    box.width = Emu(entry.get("w", 4000000))
                    box.height = Emu(entry.get("h", 800000))
            else:
                text, x, y, w, h, size, bold = entry
                box = _textbox(slide, text, Emu(x), Emu(y), Emu(w), Emu(h))
                run = box.text_frame.paragraphs[0].runs[0]
                if size:
                    run.font.size = Pt(size)
                if bold is not None:
                    run.font.bold = bold
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _payload(data: bytes):
    from services import parse

    path = __import__("pathlib").Path(__import__("tempfile").mkdtemp()) / "deck.pptx"
    path.write_bytes(data)
    _objs, payload = parse.parse_pptx(path)
    return payload


def _by_id(payload, page):
    return {
        obj["shape_id"]: obj
        for obj in payload["objects"]
        if obj.get("kind") == "text" and int(obj.get("slide_index")) == page
    }


def _entry(changelog, page, section):
    slides = {s["page"]: s for s in changelog["slides"]}
    return slides.get(page, {}).get(section) or []


# ---------- 单元：services.diff.structural_changelog ----------

class TestStructuralDiff:
    def test_identical_decks_empty(self):
        data = _build([ [("标题", 1000000, 1000000, 4000000, 800000, 32, True)] ])
        changelog = diff.structural_changelog(_payload(data), _payload(data))
        assert changelog["summary"]["pages_changed"] == 0
        assert all(
            not changelog["slides"][i][section]
            for i in range(len(changelog["slides"]))
            for section in ("added", "removed", "moved", "text_changed", "geometry_changed", "style_changed")
        )

    def test_added_and_removed_objects(self):
        a = _build([ [("保留", 1000000, 1000000, 4000000, 800000, 18, None)] ])
        b = _build([ [
            ("保留", 1000000, 1000000, 4000000, 800000, 18, None),
            ("新增框", 1000000, 3000000, 4000000, 800000, 18, None),
        ] ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        added = _entry(changelog, 1, "added")
        assert len(added) == 1
        assert added[0]["text"] == "新增框"
        assert added[0]["shape_id"]
        assert not _entry(changelog, 1, "removed")
        # 反向：B 缺一个对象 → removed
        changelog_rev = diff.structural_changelog(_payload(b), _payload(a))
        removed = _entry(changelog_rev, 1, "removed")
        assert len(removed) == 1
        assert removed[0]["text"] == "新增框"

    def test_text_changed_same_shape_id(self):
        a = _build([ [("旧标题", 1000000, 1000000, 4000000, 800000, 32, None)] ])
        b = _build([ [("新标题", 1000000, 1000000, 4000000, 800000, 32, None)] ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        changed = _entry(changelog, 1, "text_changed")
        assert len(changed) == 1
        assert changed[0]["before"] == "旧标题"
        assert changed[0]["after"] == "新标题"
        # 同 shape_id 才算"改"而不是"删+加"
        assert not _entry(changelog, 1, "added") and not _entry(changelog, 1, "removed")

    def test_geometry_changed_axes_and_moved(self):
        a = _build([ [("正文", 1000000, 1000000, 4000000, 800000, 18, None)] ])
        # 平移 400000 EMU（> 2pt 阈值）→ moved + geometry_changed(x/y)
        b = _build([ [("正文", 1400000, 1000000, 4000000, 800000, 18, None)] ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        moved = _entry(changelog, 1, "moved")
        assert len(moved) == 1
        assert moved[0]["from"]["x"] == 1000000
        assert moved[0]["to"]["x"] == 1400000
        geo = _entry(changelog, 1, "geometry_changed")
        assert len(geo) == 1
        assert set(geo[0]["axes"]) == {"x"}
        assert not _entry(changelog, 1, "text_changed")
        # 缩放（w 变化）→ geometry_changed(w)，不算 moved
        c = _build([ [("正文", 1000000, 1000000, 5000000, 800000, 18, None)] ])
        changelog_c = diff.structural_changelog(_payload(a), _payload(c))
        assert not _entry(changelog_c, 1, "moved")
        assert set(_entry(changelog_c, 1, "geometry_changed")[0]["axes"]) == {"w"}

    def test_style_changed_fields(self):
        a = _build([ [{"text": "要点", "size": 18, "bold": None, "color": None, "x": 1000000}] ])
        b = _build([ [{"text": "要点", "size": 24, "bold": True, "color": "CC0000", "x": 1000000}] ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        styles = _entry(changelog, 1, "style_changed")
        assert len(styles) == 1
        fields = {c["field"]: c for c in styles[0]["changes"]}
        assert fields["font_size"]["before"] == 18 and fields["font_size"]["after"] == 24
        assert fields["bold"]["after"] is True
        assert fields["color"]["after"] == "CC0000"

    def test_small_jitter_not_moved(self):
        # 12700 EMU（1pt）内的抖动不算移动（阈值 2pt）
        a = _build([ [("框", 1000000, 1000000, 4000000, 800000, 18, None)] ])
        b = _build([ [("框", 1000000 + EMU_PT, 1000000, 4000000, 800000, 18, None)] ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        assert not _entry(changelog, 1, "moved")

    def test_added_page_objects(self):
        a = _build([ [("第一页", 1000000, 1000000, 4000000, 800000, 18, None)] ])
        b = _build([
            [("第一页", 1000000, 1000000, 4000000, 800000, 18, None)],
            [("新页对象", 1000000, 1000000, 4000000, 800000, 18, None)],
        ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        added = _entry(changelog, 2, "added")
        assert len(added) == 1
        assert changelog["summary"]["pages_added"] == 1

    def test_changelog_json_serializable(self):
        a = _build([ [("标题", 1000000, 1000000, 4000000, 800000, 32, True)] ])
        b = _build([ [
            ("改了", 1200000, 1000000, 4000000, 800000, 40, None),
            ("多出来的", 1000000, 3000000, 4000000, 800000, 18, None),
        ] ])
        changelog = diff.structural_changelog(_payload(a), _payload(b))
        text = json.dumps(changelog, ensure_ascii=False)
        assert '"text_changed"' in text and '"style_changed"' in text


# ---------- API：compare 响应升级（向后兼容 + struct 字段） ----------

class TestCompareApiStruct:
    def test_compare_returns_struct_changelog(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        obj = next(o for o in uploaded_doc["objects"] if o["kind"] == "text" and o["text"])
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "edit", "shape_id": obj["shape_id"],
                               "after": "结构diff后的新标题", "before": obj["text"]}],
                  "label": "v05-8 diff"},
        )
        assert resp.status_code == 200, resp.text
        cmp = client.get(f"/api/docs/{doc_id}/versions/v0/compare/v1")
        assert cmp.status_code == 200, cmp.text
        body = cmp.json()
        # 旧字段仍在（向后兼容）
        assert "changed" in body and "added" in body and "removed" in body
        # 新增结构化 changelog
        struct = body["struct"]
        assert struct["a"] == "v0" and struct["b"] == "v1"
        text_changed = [c for s in struct["slides"] for c in s["text_changed"]]
        assert any("结构diff后的新标题" in c["after"] for c in text_changed)

    def test_compare_struct_geometry_after_move(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        obj = next(o for o in uploaded_doc["objects"] if o.get("kind") == "text" and o.get("movable", True))
        bbox = obj["bbox_emu"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "geometry", "shape_id": obj["shape_id"],
                               "bbox_emu": {"x": bbox["x"] + 600000, "y": bbox["y"],
                                            "w": bbox["w"], "h": bbox["h"]}}],
                  "label": "v05-8 move"},
        )
        assert resp.status_code == 200, resp.text
        cmp = client.get(f"/api/docs/{doc_id}/versions/v0/compare/v1")
        struct = cmp.json()["struct"]
        moved = [m for s in struct["slides"] for m in s["moved"]]
        assert any(m["to"]["x"] == bbox["x"] + 600000 for m in moved)
