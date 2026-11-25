"""v0.5 F4：set_transition 页面切换动画（fade/push/wipe 等 + 速度 + 自动换片）。

python-pptx 无切换 API，手工组 <p:transition> XML（CT_Slide 序列位：
cSld → clrMapOvr → transition → timing）。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn

from services import edits


def _transition_of(prs, index):
    return prs.slides[index]._element.find(qn("p:transition"))


class TestSetTransition:
    def test_fade_with_speed(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"pages": "1", "effect": "fade", "speed": "slow"}}],
        )
        assert touched == {0}
        prs = Presentation(str(dst))
        trans = _transition_of(prs, 0)
        assert trans is not None
        assert trans.find(qn("p:fade")) is not None
        assert trans.get("spd") == "slow"

    def test_push_with_direction(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"pages": "2-3", "effect": "push", "direction": "l"}}],
        )
        prs = Presentation(str(dst))
        for index in (1, 2):
            trans = _transition_of(prs, index)
            assert trans is not None
            push = trans.find(qn("p:push"))
            assert push is not None
            assert push.get("dir") == "l"

    def test_advance_after_and_click(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {
                "pages": "1", "effect": "wipe",
                "advance_after": 5, "advance_on_click": False,
            }}],
        )
        prs = Presentation(str(dst))
        trans = _transition_of(prs, 0)
        assert trans.get("advTm") == "5000"
        assert trans.get("advClick") == "0"

    def test_none_removes_transition(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"pages": "1", "effect": "fade"}}],
        )
        edits.apply_patches_to_file(
            dst, dst,
            [{"action": "set_transition", "payload": {"pages": "1", "effect": "none"}}],
        )
        prs = Presentation(str(dst))
        assert _transition_of(prs, 0) is None

    def test_replaces_existing_transition(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"pages": "1", "effect": "fade"}}],
        )
        edits.apply_patches_to_file(
            dst, dst,
            [{"action": "set_transition", "payload": {"pages": "1", "effect": "cut", "speed": "fast"}}],
        )
        prs = Presentation(str(dst))
        trans = _transition_of(prs, 0)
        assert trans.find(qn("p:fade")) is None
        assert trans.find(qn("p:cut")) is not None
        assert trans.get("spd") == "fast"

    def test_deck_scope_default(self, demo_pptx, tmp_path):
        # 不给 pages/slide_index → 默认整套
        dst = tmp_path / "out.pptx"
        touched = edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"effect": "dissolve"}}],
        )
        assert touched == {0, 1, 2}
        prs = Presentation(str(dst))
        for index in range(3):
            assert _transition_of(prs, index).find(qn("p:dissolve")) is not None

    def test_slide_index_0based(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"slide_index": 2, "effect": "random"}}],
        )
        prs = Presentation(str(dst))
        assert _transition_of(prs, 2) is not None
        assert _transition_of(prs, 0) is None

    def test_transition_keeps_schema_position_before_timing(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "set_transition", "payload": {"pages": "1", "effect": "fade"}}],
        )
        prs = Presentation(str(dst))
        children = [c.tag.split("}")[-1] for c in prs.slides[0]._element]
        assert "transition" in children
        # transition 必须在 cSld 之后、timing 之前（timing 存在时）
        pos = children.index("transition")
        assert pos > children.index("cSld")
        if "timing" in children:
            assert pos < children.index("timing")

    def test_invalid_effect_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_transition", "payload": {"pages": "1", "effect": "explode"}}],
            )

    def test_missing_effect_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_transition", "payload": {"pages": "1"}}],
            )

    def test_invalid_speed_rejected(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_transition", "payload": {"pages": "1", "effect": "fade", "speed": "instant"}}],
            )

    def test_pages_out_of_range(self, demo_pptx, tmp_path):
        with pytest.raises(edits.PatchError):
            edits.apply_patches_to_file(
                demo_pptx, tmp_path / "o.pptx",
                [{"action": "set_transition", "payload": {"pages": "9", "effect": "fade"}}],
            )

    def test_alias_transition(self, demo_pptx, tmp_path):
        dst = tmp_path / "out.pptx"
        edits.apply_patches_to_file(
            demo_pptx, dst,
            [{"action": "transition", "payload": {"pages": "1", "effect": "cover", "direction": "r"}}],
        )
        prs = Presentation(str(dst))
        assert _transition_of(prs, 0).find(qn("p:cover")) is not None


class TestTransitionApi:
    def test_apply_via_api(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        resp = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"patches": [{"action": "set_transition", "payload": {"effect": "fade", "speed": "slow"}}],
                  "label": "统一切换动画"},
        )
        assert resp.status_code == 200, resp.text
        assert resp.json()["version"] == 1
