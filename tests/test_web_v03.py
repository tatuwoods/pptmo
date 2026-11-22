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
