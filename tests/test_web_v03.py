"""v0.3 D 组界面体验契约测试：快捷键 / 拖拽上传 / 缩略图切页 / 深色模式。

前端无 JS 测试基建，以静态契约（关键标识符与行为绑定必须存在）锁住实现，
改坏任意一项都会在这里红。
"""
from __future__ import annotations

import re
from pathlib import Path

WEB = Path(__file__).resolve().parent.parent / "app" / "web"


def _read(name: str) -> str:
    return (WEB / name).read_text(encoding="utf-8")


# ---------- 17. 键盘快捷键 ----------


def test_keyboard_undo_shortcut_wired():
    js = _read("app.js")
    # ⌘Z/Ctrl+Z 触发撤销上一轮应用
    assert "metaKey" in js and "'z'" in js.replace('"z"', "'z'")
    assert "undo" in js
    assert "/undo" in js  # 调用后端撤销端点


def test_keyboard_cmd_enter_send_and_esc_stop():
    js = _read("app.js")
    # ⌘/Ctrl+Enter 发送：必须是聊天输入框 keydown 里「带修饰键的 Enter」分支。
    # 裸 Enter 发送是 v0.2 就有的，只断言 sendChat() 存在锁不住这个功能
    composer = js.split('dom.chatInput.addEventListener("keydown"', 1)[1][:600]
    assert '(event.metaKey || event.ctrlKey) && event.key === "Enter"' in composer
    assert "sendChat()" in composer
    # Esc 停止生成：全局 keydown 里 streaming 状态下 Esc 先 abort 流，
    # 且排在 busy 判断之前（处理中也要能停）
    global_keys = js.split('document.addEventListener("keydown"', 1)[1]
    esc_stop = global_keys.split("if (state.busy", 1)[0]
    assert 'event.key === "Escape" && state.streaming' in esc_stop
    assert "streamController.abort()" in esc_stop


# ---------- 18. 拖拽上传 ----------


def test_drag_upload_overlay_and_friendly_reject():
    html = _read("index.html")
    js = _read("app.js")
    assert "dropOverlay" in html  # 全窗拖拽提示层
    assert "dropOverlay" in js
    # .ppt/.pptm 拖入时前端给出去向清晰的提示（不只是"只支持 .pptx"）
    assert ".pptm" in js


# ---------- 19. 缩略图点击切页联动 ----------


def test_thumb_click_switches_page_and_scope():
    js = _read("app.js")
    # 点击缩略图：同一个 click 处理器里先切页、再把对话范围同步为「当前页」。
    # 只切页不联动范围是 v0.2 的旧行为，所以必须断言 setScope 也在这个处理器里
    thumbs = js.split("function renderThumbs()", 1)[1].split("\nfunction ", 1)[0]
    assert re.search(
        r'button\.addEventListener\("click", \(\) => \{'
        r'[\s\S]*?setSlide\(Number\(button\.dataset\.slide\)\);'
        r'[\s\S]*?if \(state\.scope !== "current"\) setScope\("current"\);',
        thumbs,
    )


# ---------- 20. 深色模式跟随系统 ----------


def test_dark_mode_follows_system():
    html = _read("index.html")
    css = _read("styles.css")
    assert 'name="color-scheme"' in html and "dark" in html.split('name="color-scheme"')[1][:60]
    assert "@media (prefers-color-scheme: dark)" in css
    # 深色变量组必须覆盖核心表面色，而不是只改一两个变量
    dark_block = css.split("@media (prefers-color-scheme: dark)", 1)[1]
    for var in ("--bg:", "--panel:", "--ink:", "--line:"):
        assert var in dark_block, f"dark block missing {var}"


def _dark_block(css: str) -> str:
    return css.split("@media (prefers-color-scheme: dark)", 1)[1].split("\n}", 1)[0]


def _outside_theme_blocks(css: str) -> str:
    """CSS with :root {...} and the dark media block removed — what is left must be mode-independent."""
    kept, depth, skipping = [], 0, False
    for line in css.splitlines():
        stripped = line.strip()
        if depth == 0 and (stripped.startswith(":root") or stripped.startswith("@media (prefers-color-scheme: dark)")):
            skipping = True
        depth += line.count("{") - line.count("}")
        if not skipping:
            kept.append(line)
        if skipping and depth == 0:
            skipping = False
    return "\n".join(kept)


# 深色模式只靠变量组切换，所以主题块之外不能再写死随主题变化的文字色/底色——
# 写死的值在深色下不会跟着变。v0.3.0 实测深色下 AI 回复正文、顶部按钮、文档名的
# 对比度只有 1.7:1，范围切换器、对话工具栏等还是浅色底。
_THEME_TEXT = {
    "#334155", "#475569", "#64748b", "#94a3b8", "#3f453f",  # 中性文字
    "#0f172a", "#0b1220", "#262b27",                         # 标题
    "#1d4ed8", "#1e3a8a",                                    # 浅强调底上的文字
    "#9a3412", "#92400e",                                    # 提示底上的文字
    "#175c32", "#a52a12", "#0c5c2d",                         # 修改前后对比的文字
}
_THEME_SURFACE = {
    "#e9eef7", "#eef2f8", "#f8fafc", "#f7f9fc", "#fbfcff", "#f8fbff",  # 中性浅底
    "#eef2ff", "#eff6ff", "#fcfffd", "#fff6f6",                         # 着色浅底
    "#fffbeb", "#faf8f4", "#f0fbf4", "#ffe3dc", "#c9f2d8",
}


def test_dark_mode_has_no_theme_dependent_hardcoded_colors():
    body = _outside_theme_blocks(_read("styles.css"))
    offenders = []
    for prop, value in re.findall(r"([\w-]+)\s*:\s*([^;{}]+)", body):
        palette = _THEME_TEXT if prop == "color" else _THEME_SURFACE if prop in ("background", "background-color") else None
        if palette is None:
            continue
        offenders += [f"{prop}: {value.strip()}" for token in re.findall(r"#[0-9a-fA-F]{3,8}\b", value) if token.lower() in palette]
    assert offenders == [], f"写死的随主题变化颜色（深色下不跟随）: {offenders}"


def _luminance(hex_color: str) -> float:
    channels = [int(hex_color[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    linear = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in channels]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def _contrast(fg: str, bg: str) -> float:
    hi, lo = sorted((_luminance(fg), _luminance(bg)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


def test_dark_canvas_backdrop_stays_dark():
    """--dark 是幻灯片四周的画布衬底，不是文字色：深色模式下也必须保持深色。"""
    dark_block = _dark_block(_read("styles.css"))
    value = re.search(r"--dark:\s*(#[0-9a-fA-F]{6})", dark_block).group(1)
    channels = [int(value[i:i + 2], 16) for i in (1, 3, 5)]
    assert max(channels) <= 0x20, f"深色模式下画布衬底应为深色，实际 {value}"


def test_dark_text_tokens_are_readable_on_dark_surfaces():
    """深色文字色压在深色表面上要达到 4.5:1（正文可读的最低线）。"""
    css = _read("styles.css")
    tokens = dict(re.findall(r"(--[\w-]+)\s*:\s*(#[0-9a-fA-F]{6})", _dark_block(css)))
    pairs = [
        ("--ink", "--panel"), ("--ink-2", "--panel"), ("--ink-3", "--panel-2"),
        ("--muted", "--panel-2"), ("--muted-2", "--panel-2"), ("--muted", "--well"),
        ("--accent-strong", "--chip-soft"), ("--amber", "--amber-soft"),
        ("--ink-accent-soft", "--accent-soft"),
    ]
    unreadable = [
        f"{fg} on {bg} = {_contrast(tokens[fg], tokens[bg]):.2f}:1"
        for fg, bg in pairs if fg in tokens and bg in tokens and _contrast(tokens[fg], tokens[bg]) < 4.5
    ]
    missing = [name for pair in pairs for name in pair if name not in tokens]
    assert not missing and not unreadable, f"缺少深色变量 {missing}；对比度不足 {unreadable}"


def test_dark_mode_keeps_white_on_accent_text_readable():
    """深色的 --accent 更亮，白字压上去只剩 3:1：按钮/头像的底色在深色块里要压回深一档。"""
    dark_block = _dark_block(_read("styles.css"))
    for selector in (".btn\\.primary", ".chat-avatar"):
        rule = re.search(rf"{selector}\s*\{{[^}}]*\}}", dark_block)
        assert rule and "background" in rule.group(0) and "#2563eb" in rule.group(0), f"深色块缺少 {selector} 的底色覆盖"


# ---------- 14b / 16b：后端发的 plan / retry 事件，界面上必须接得住 ----------


def test_plan_event_is_rendered_and_kept_in_chat_log():
    js = _read("app.js")
    css = _read("styles.css")
    handler = js.split("function handleChatEvent(", 1)[1].split("\nfunction ", 1)[0]
    assert 'event === "plan"' in handler
    plan_branch = handler.split('event === "plan"', 1)[1].split("} else if", 1)[0]
    assert "live.plan" in plan_branch
    # 刷新页面后仍可见：会话记录里的 assistant 消息要渲染 message.plan
    message_html = js.split("function chatMessageHtml(", 1)[1].split("\nfunction ", 1)[0]
    assert "message.plan" in message_html
    # 计划是 AI 的：只对 assistant 消息渲染，用户消息即使带了 plan 字段也不显示
    assert 'const plan = role === "assistant" ? planListHtml(message.plan) : "";' in message_html
    assert "我将" in js and "plan-list" in js and ".plan-list" in css


def test_retry_event_is_surfaced():
    js = _read("app.js")
    handler = js.split("function handleChatEvent(", 1)[1].split("\nfunction ", 1)[0]
    assert 'event === "retry"' in handler
    retry_branch = handler.split('event === "retry"', 1)[1].split("} else if", 1)[0]
    # 上一次尝试流出的半截 token 必须丢掉，否则会和重试输出拼成一段乱码思考
    assert 'live.raw = ""' in retry_branch
    assert "data.message" in retry_branch


# ---------- 15b：AI 撤销后画布要跟上 ----------


def test_ai_undo_refreshes_document():
    js = _read("app.js")
    finalize = js.split("function finalizeChat(", 1)[1].split("\nfunction ", 1)[0]
    match = re.search(r"if \(done\.undo\) \{([\s\S]*?)\n  \}", finalize)
    assert match, "finalizeChat 没有处理 AI 撤销（done.undo）"
    # 服务端已提交回退版本：必须重新拉文档上屏，只刷新 messages 会让画布停在旧版本
    assert "applyDocument(" in match.group(1)
