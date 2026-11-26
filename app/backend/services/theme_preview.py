"""v0.6 #8：整稿重主题预览安全网（对标 CasualOffice round-trip 测试）。

theme_rebrand / recolor 这类「一改全稿」的操作风险高：直接 apply 到真文档，
效果不满意要靠 undo 兜底。预览模式在副本上应用同样的 patches 并渲染，
产出 before/after 逐页对比图——用户先看效果，确认后照常 POST /apply。

安全契约：
- 真文档一字节不动（版本号、current.pptx、消息流都不碰）；
- 只接受整稿重主题类操作（theme_rebrand/recolor 及别名），其它 action 拒绝；
- 预览产物落在 doc_dir/theme-previews/<token>/，URL 按文件名白名单分发。
"""
from __future__ import annotations

import re
import shutil
import uuid
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from pptx import Presentation

from services import edits

# 允许预览的整稿重主题类操作（含别名 → 规范名后的集合）
THEME_PREVIEW_ACTIONS = {"theme_rebrand", "recolor"}
DEFAULT_PREVIEW_PAGES = 3   # 缺省预览前 3 页
MAX_PREVIEW_PAGES = 8       # 单次预览上限（对比图别把盘塞满）

PREVIEW_FILE_RE = re.compile(r"^(before|after)-\d{3}\.png$")


class PreviewError(ValueError):
    """预览构造失败（patches 非法/页码越界），端点转 4xx。"""


def _normalize(patches: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not patches:
        raise PreviewError("patches 不能为空")
    normalized: List[Dict[str, Any]] = []
    for patch in patches:
        action = str((patch or {}).get("action") or "").strip()
        canonical = edits._ACTION_ALIASES.get(action, action)
        if canonical not in THEME_PREVIEW_ACTIONS:
            raise PreviewError(
                f"theme-preview 只支持整稿重主题操作（theme_rebrand/recolor），收到：{action or '空'}"
            )
        normalized.append({"action": canonical, "payload": (patch or {}).get("payload") or {}})
    return normalized


def _select_pages(
    slide_count: int, pages: Optional[List[int]]
) -> List[int]:
    wanted: List[int] = list(range(min(DEFAULT_PREVIEW_PAGES, slide_count))) if pages is None else pages
    if not wanted:
        raise PreviewError("pages 为空列表：至少选一页")
    if len(wanted) > MAX_PREVIEW_PAGES:
        raise PreviewError(f"一次最多预览 {MAX_PREVIEW_PAGES} 页（收到 {len(wanted)}）")
    for index in wanted:
        if not isinstance(index, int) or isinstance(index, bool) or index < 0 or index >= slide_count:
            raise PreviewError(f"页码超出范围：{index}（文档共 {slide_count} 页，0-based）")
    return sorted(set(wanted))


def build_theme_preview(
    src_pptx: Path,
    patches: List[Dict[str, Any]],
    out_dir: Path,
    before_dir: Optional[Path] = None,
    pages: Optional[List[int]] = None,
    render_fn: Optional[Callable[..., int]] = None,
) -> Dict[str, Any]:
    """在副本上应用重主题 patches 并渲染，产出前后对比页图清单。

    - src_pptx 只读（副本落 out_dir/copy-*.pptx，用完即删）；
    - before 优先复用 before_dir（当前版本渲染页），缺失时现场补渲来源稿；
    - render_fn 缺省用 render.render_pptx（测试注入假实现绕开 LibreOffice）。
    返回 {"token", "page_count", "pages": [{"slide_index", "before", "after"}]}。
    """
    if render_fn is None:
        from services import render as render_mod

        render_fn = render_mod.render_pptx
    normalized = _normalize(patches)

    slide_count = len(Presentation(str(src_pptx)).slides)
    if slide_count <= 0:
        raise PreviewError("文档没有页面，无法预览")
    wanted = _select_pages(slide_count, pages)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    token = out_dir.name or uuid.uuid4().hex

    # 1) 副本上应用同样的 patches（apply_patches_to_file 天然只碰副本）
    copy_path = out_dir / f"copy-{src_pptx.stem}.pptx"
    shutil.copy2(src_pptx, copy_path)
    try:
        edits.apply_patches_to_file(copy_path, copy_path, normalized)
    except edits.PatchError as exc:
        try:
            copy_path.unlink()
        except OSError:
            pass
        raise PreviewError(f"patches 应用失败：{exc}") from exc

    # 2) 渲染副本（after 侧）
    try:
        render_dir = out_dir / "after-render"
        render_fn(copy_path, render_dir)
        for index in wanted:
            src_page = render_dir / f"slide-{index + 1:03d}.png"
            if not src_page.exists():
                raise PreviewError(f"预览渲染缺页：slide-{index + 1:03d}.png")
            shutil.copy2(src_page, out_dir / f"after-{index + 1:03d}.png")
    finally:
        try:
            copy_path.unlink()
        except OSError:
            pass
        shutil.rmtree(out_dir / "after-render", ignore_errors=True)

    # 3) before 侧：优先复用当前版本渲染页，缺失则现场补渲来源稿
    fallback_dir: Optional[Path] = None
    for index in wanted:
        name = f"before-{index + 1:03d}.png"
        source = Path(before_dir or "") / f"slide-{index + 1:03d}.png"
        if source.exists():
            shutil.copy2(source, out_dir / name)
            continue
        if fallback_dir is None:
            fallback_dir = out_dir / "before-render"
            render_fn(src_pptx, fallback_dir)
        source = fallback_dir / f"slide-{index + 1:03d}.png"
        if not source.exists():
            raise PreviewError(f"来源稿渲染缺页：slide-{index + 1:03d}.png")
        shutil.copy2(source, out_dir / name)
    if fallback_dir is not None:
        shutil.rmtree(fallback_dir, ignore_errors=True)

    return {
        "token": token,
        "page_count": slide_count,
        "pages": [
            {
                "slide_index": index,
                "before": str(out_dir / f"before-{index + 1:03d}.png"),
                "after": str(out_dir / f"after-{index + 1:03d}.png"),
            }
            for index in wanted
        ],
    }
