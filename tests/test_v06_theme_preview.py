"""v0.6 #8：整稿重主题预览安全网 —— 大改动先看效果再落稿。

theme_rebrand/recolor 加预览模式：POST /api/docs/{id}/theme-preview 在副本上
应用 patches 并渲染，返回 before/after 页图 URL 给用户确认；真文档一字节
不动（版本号不变），确认后照常 POST /apply。服务层可注入假渲染函数，
不依赖 LibreOffice。
"""
from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Emu

from services import theme_preview

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)

REBRAND_PATCH = {"action": "theme_rebrand", "payload": {"palette": {"accent1": "0066CC"}}}
TINY_PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c626001000000ffff03000006000557bfabd40000000049454e44ae426082"
)


def _deck(path: Path, pages: int = 4) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H
    for index in range(pages):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Emu(500000), Emu(500000), Emu(4000000), Emu(800000)).text_frame.text = f"第{index + 1}页"
    prs.save(str(path))
    return path


def _fake_render(pptx_path, out_dir, scale=2.0, stale_dir=None):
    """测试替身：按页数写占位 PNG（内容与真渲染契约一致：slide-NNN.png）。"""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    count = len(Presentation(str(pptx_path)).slides)
    for index in range(count):
        (out_dir / f"slide-{index + 1:03d}.png").write_bytes(TINY_PNG)
        (out_dir / f"thumb-{index + 1:03d}.png").write_bytes(TINY_PNG)
    return count


# ---------- 服务层 ----------


class TestBuildPreview:
    def test_copy_rendered_source_untouched(self, tmp_path):
        deck = _deck(tmp_path / "deck.pptx", pages=4)
        original = deck.read_bytes()
        out = tmp_path / "preview"
        result = theme_preview.build_theme_preview(
            deck, [REBRAND_PATCH], out,
            before_dir=_fake_render_to(tmp_path / "before", 4),
            render_fn=_fake_render,
        )
        assert deck.read_bytes() == original  # 真文档一字节不动
        assert result["page_count"] == 4
        # 缺省取前 3 页，每页 before/after 成对存在
        assert [p["slide_index"] for p in result["pages"]] == [0, 1, 2]
        for page in result["pages"]:
            assert Path(page["before"]).exists() and Path(page["after"]).exists()
        assert (out / "after-004.png").exists() is False  # 只导出选中的页

    def test_explicit_pages_selection(self, tmp_path):
        deck = _deck(tmp_path / "deck.pptx", pages=5)
        result = theme_preview.build_theme_preview(
            deck, [REBRAND_PATCH], tmp_path / "preview",
            pages=[2, 4], before_dir=_fake_render_to(tmp_path / "before", 5),
            render_fn=_fake_render,
        )
        assert [p["slide_index"] for p in result["pages"]] == [2, 4]

    def test_out_of_range_page_rejected(self, tmp_path):
        deck = _deck(tmp_path / "deck.pptx", pages=3)
        with pytest.raises(theme_preview.PreviewError, match="超出范围"):
            theme_preview.build_theme_preview(
                deck, [REBRAND_PATCH], tmp_path / "preview",
                pages=[3], before_dir=_fake_render_to(tmp_path / "before", 3),
                render_fn=_fake_render,
            )

    def test_non_theme_patch_rejected(self, tmp_path):
        deck = _deck(tmp_path / "deck.pptx")
        with pytest.raises(theme_preview.PreviewError, match="只支持整稿重主题"):
            theme_preview.build_theme_preview(
                deck, [{"action": "replace_text", "payload": {}}],
                tmp_path / "preview", render_fn=_fake_render,
            )

    def test_before_missing_renders_source(self, tmp_path):
        """当前版本页图缺失时现场渲染来源稿，before 侧不缺页。"""
        deck = _deck(tmp_path / "deck.pptx", pages=2)
        empty_before = tmp_path / "no-before"
        empty_before.mkdir()
        calls = []

        def counting_render(pptx_path, out_dir, scale=2.0, stale_dir=None):
            calls.append(Path(pptx_path).name)
            return _fake_render(pptx_path, out_dir, scale, stale_dir)

        result = theme_preview.build_theme_preview(
            deck, [REBRAND_PATCH], tmp_path / "preview",
            before_dir=empty_before, render_fn=counting_render,
        )
        assert sorted(calls) == ["copy-deck.pptx", "deck.pptx"]  # 副本 + 来源补渲
        for page in result["pages"]:
            assert Path(page["before"]).exists()

    def test_recolor_alias_accepted(self, tmp_path):
        deck = _deck(tmp_path / "deck.pptx", pages=2)
        result = theme_preview.build_theme_preview(
            deck, [{"action": "rebrand", "payload": {"palette": {"accent1": "112233"}}}],
            tmp_path / "preview", render_fn=_fake_render,
        )
        assert result["page_count"] == 2


def _fake_render_to(path: Path, pages: int) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    for index in range(pages):
        (path / f"slide-{index + 1:03d}.png").write_bytes(TINY_PNG)
    return path


# ---------- 端点 ----------


class TestThemePreviewEndpoint:
    def test_preview_does_not_touch_current_version(self, client, uploaded_doc):
        from storage import current_path, load_meta

        doc_id = uploaded_doc["doc_id"]
        before_bytes = current_path(doc_id).read_bytes()
        before_version = int(load_meta(doc_id)["version"])

        response = client.post(
            f"/api/docs/{doc_id}/theme-preview",
            json={"patches": [REBRAND_PATCH]},
        )
        assert response.status_code == 200, response.text
        payload = response.json()
        assert payload["page_count"] >= 1
        assert payload["pages"], "demo 稿至少预览第 1 页"
        first = payload["pages"][0]
        assert first["slide_index"] == 0
        assert first["before_url"] and first["after_url"]
        # 真文档一字节未动、版本号不变
        assert current_path(doc_id).read_bytes() == before_bytes
        assert int(load_meta(doc_id)["version"]) == before_version
        # 对比图可访问
        for key in ("before_url", "after_url"):
            img = client.get(first[key])
            assert img.status_code == 200, first[key]
            assert img.content

    def test_apply_after_preview_still_works(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        preview = client.post(
            f"/api/docs/{doc_id}/theme-preview",
            json={"patches": [REBRAND_PATCH], "slides": [0]},
        )
        assert preview.status_code == 200
        applied = client.post(
            f"/api/docs/{doc_id}/apply",
            json={"label": "重主题", "patches": [REBRAND_PATCH]},
        )
        assert applied.status_code == 200, applied.text
        new_version = int(str(applied.json().get("version", "v0")).lstrip("v"))
        old_version = int(str(uploaded_doc["version_id"]).lstrip("v"))
        assert new_version > old_version

    def test_bad_patch_rejected_4xx(self, client, uploaded_doc):
        doc_id = uploaded_doc["doc_id"]
        response = client.post(
            f"/api/docs/{doc_id}/theme-preview",
            json={"patches": [{"action": "delete_slide", "payload": {}}]},
        )
        assert 400 <= response.status_code < 500
        assert "整稿重主题" in response.json()["detail"]
