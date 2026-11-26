from __future__ import annotations

import hashlib
import hmac
import html
import io
import json
import logging
import os
import queue
import re
import shutil
import tempfile
import threading
import uuid
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, Header, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.datastructures import Headers as StarletteHeaders

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

import storage  # noqa: E402
from services import agent, ai, broker, diff, edits, parse, render, vision  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("ppt-light")

APP_DIR = BASE_DIR.parent
WEB_DIR = APP_DIR / "web"
MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_MB", "60")) * 1024 * 1024
RENDER_SCALE = float(os.getenv("RENDER_SCALE", "2"))
PPTX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# --- request guard configuration (v0.1.2 security hardening) ---
# TRUSTED_HOSTS / TRUSTED_ORIGINS: extra Host/Origin values to accept, comma
# separated. Needed when the app is reached through a reverse proxy (e.g.
# `tailscale serve`), where the browser origin differs from the loopback Host
# the proxy forwards. Hostnames appearing in TRUSTED_ORIGINS are trusted
# automatically. Loopback names are always trusted.
TRUSTED_HOSTS = [h.strip().lower() for h in os.getenv("TRUSTED_HOSTS", "").split(",") if h.strip()]
TRUSTED_ORIGINS = [o.strip().rstrip("/") for o in os.getenv("TRUSTED_ORIGINS", "").split(",") if o.strip()]
MAX_JSON_BODY_BYTES = int(os.getenv("MAX_JSON_BODY_MB", "2")) * 1024 * 1024
MULTIPART_SLACK_BYTES = 8 * 1024 * 1024  # multipart framing overhead beyond the file itself
AUTO_APPLY_DESTRUCTIVE = os.getenv("AUTO_APPLY_DESTRUCTIVE", "1").strip().lower() in ("1", "true", "yes")
DESTRUCTIVE_ACTIONS = {"delete_textbox", "clear", "delete_slide"}
_LOCAL_HOST_NAMES = {"localhost", "127.0.0.1", "::1", "[::1]", "0.0.0.0", "testserver", "testclient"}

# bound concurrent agent chats: each one holds LLM connections for minutes
AI_CONCURRENCY = max(1, int(os.getenv("AI_CONCURRENCY", "2")))
_chat_slots = threading.BoundedSemaphore(AI_CONCURRENCY)
_chat_inflight_lock = threading.Lock()
_chat_inflight: Dict[str, int] = {}
_MUTATING_METHODS = {"POST", "PUT", "PATCH", "DELETE"}


def _origin_hostname(origin: str) -> str:
    match = re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://([^/?#]+)", origin or "")
    if not match:
        return ""
    return match.group(1).rsplit(":", 1)[0].strip("[]").lower()


_TRUSTED_ORIGIN_HOSTS = {_origin_hostname(o) for o in TRUSTED_ORIGINS}
_ALLOWED_HOST_NAMES = _LOCAL_HOST_NAMES | set(TRUSTED_HOSTS) | {h for h in _TRUSTED_ORIGIN_HOSTS if h}


def _host_allowed(request_headers: StarletteHeaders) -> bool:
    host = request_headers.get("host", "")
    hostname = host.rsplit(":", 1)[0].strip("[]").lower()
    # forwarded hosts (reverse proxy) are acceptable if trusted
    forwarded = request_headers.get("x-forwarded-host", "")
    forwarded_names = {value.strip().rsplit(":", 1)[0].strip("[]").lower() for value in forwarded.split(",") if value.strip()}
    if not hostname:
        return True  # HTTP/1.0 style or internal client; Origin check still applies
    return hostname in _ALLOWED_HOST_NAMES or bool(forwarded_names & {h for h in _ALLOWED_HOST_NAMES if h})


def _origin_allowed(request_headers: StarletteHeaders) -> bool:
    origin = request_headers.get("origin", "").rstrip("/")
    if not origin:
        return True  # curl / native clients / same-origin GET navigations
    if origin in TRUSTED_ORIGINS:
        return True
    host = request_headers.get("host", "")
    match = re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://([^/?#]+)", origin)
    if match and match.group(1).lower() == host.lower():
        return True  # same host:port → same-origin request
    return False


def _body_cap(request_headers: StarletteHeaders) -> int:
    content_type = request_headers.get("content-type", "").lower()
    if content_type.startswith("multipart/form-data"):
        return MAX_UPLOAD_BYTES + MULTIPART_SLACK_BYTES
    return MAX_JSON_BODY_BYTES


class RequestGuardMiddleware:
    """Blocks browser drive-by attacks (cross-site POSTs) and DNS rebinding,
    and caps request body sizes before they are buffered."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        request_headers = StarletteHeaders(scope=scope)
        method = (scope.get("method") or "GET").upper()
        path = scope.get("path") or ""

        if not _host_allowed(request_headers):
            log.warning("guard: rejected host %r for %s %s", request_headers.get("host"), method, path)
            response = JSONResponse({"detail": "host not allowed"}, status_code=421)
            await response(scope, receive, send)
            return

        if method in _MUTATING_METHODS and path.startswith("/api/"):
            sec_fetch_site = request_headers.get("sec-fetch-site")
            if sec_fetch_site == "cross-site":
                log.warning("guard: rejected cross-site %s %s", method, path)
                response = JSONResponse({"detail": "cross-site request not allowed"}, status_code=403)
                await response(scope, receive, send)
                return
            if not _origin_allowed(request_headers):
                log.warning("guard: rejected origin %r for %s %s", request_headers.get("origin"), method, path)
                response = JSONResponse({"detail": "origin not allowed"}, status_code=403)
                await response(scope, receive, send)
                return

        cap = _body_cap(request_headers)
        content_length = request_headers.get("content-length")
        if content_length and content_length.isdigit() and int(content_length) > cap:
            response = JSONResponse({"detail": "request body too large"}, status_code=413)
            await response(scope, receive, send)
            return

        if method in ("POST", "PUT", "PATCH") and not content_length:
            # chunked / unknown length: count bytes as they stream through and
            # reject mid-body if the cap is exceeded
            received = 0

            async def counting_receive():
                nonlocal received
                message = await receive()
                if message["type"] == "http.request":
                    received += len(message.get("body") or b"")
                    if received > cap:
                        raise _BodyTooLargeError()
                return message

            try:
                await self.app(scope, counting_receive, send)
            except _BodyTooLargeError:
                response = JSONResponse({"detail": "request body too large"}, status_code=413)
                await response(scope, receive, send)
            return

        await self.app(scope, receive, send)


class _BodyTooLargeError(Exception):
    pass


def _sanitize_detail(text: Any) -> str:
    """Strip server filesystem layout from error details before they reach clients."""
    clean = str(text or "")
    for prefix in {str(storage.DATA_DIR), str(BASE_DIR), str(APP_DIR), tempfile.gettempdir()}:
        if prefix and prefix != "/":
            clean = clean.replace(prefix, "<path>")
    return clean


app = FastAPI(title="PPT Light", version="0.5.0", docs_url="/api/docs-ui", openapi_url="/api/openapi.json")
app.add_middleware(GZipMiddleware, minimum_size=1500)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)
# added last → outermost, so the guard sees (and can reject) requests before
# CORSMiddleware answers preflights directly
app.add_middleware(RequestGuardMiddleware)


class ApplyIn(BaseModel):
    patches: List[Dict[str, Any]] = Field(min_length=1, max_length=100)
    label: Optional[str] = Field(default=None, max_length=80)
    base_version: Optional[int] = None
    assistant_message_id: Optional[str] = Field(default=None, max_length=32)


class ThemePreviewIn(BaseModel):
    patches: List[Dict[str, Any]] = Field(min_length=1, max_length=100)
    slides: Optional[List[int]] = None  # 0-based 页码，缺省前 3 页


class ChatIn(BaseModel):
    message: str = Field(min_length=1, max_length=4000)
    scope: Optional[str] = Field(default=None, max_length=20)
    slide_index: Optional[int] = None
    selected_shape_id: Optional[str] = Field(default=None, max_length=100)
    auto_apply: bool = False
    history: Optional[List[Dict[str, Any]]] = None
    base_version: Optional[int] = None


class RenameIn(BaseModel):
    title: str = Field(min_length=1, max_length=100)


class RestoreIn(BaseModel):
    base_version: Optional[int] = None


class PasswordIn(BaseModel):
    password: Optional[str] = Field(default=None, max_length=100)


class UnlockIn(BaseModel):
    password: str = Field(min_length=1, max_length=100)


@app.on_event("startup")
def _startup_reconcile() -> None:
    """Adopt orphaned version files left by a crash between pptx and meta writes."""
    try:
        repaired = storage.reconcile_all()
        if repaired:
            log.warning("startup: reconciled %d doc(s) with orphaned version files", repaired)
    except Exception:
        log.exception("startup reconciliation failed")

    def _daily_maintenance() -> None:
        try:
            removed = storage.run_daily_cleanup()  # 每日首次启动跑一次（D20）
            if removed:
                log.info("startup cleanup: removed %d stale file(s)", removed)
        except Exception:
            log.exception("startup cleanup failed")

    threading.Thread(target=_daily_maintenance, daemon=True, name="daily-cleanup").start()


@app.exception_handler(edits.PatchConflict)
async def _conflict_handler(_request, exc: edits.PatchConflict):
    return JSONResponse(status_code=409, content={"detail": _sanitize_detail(exc)})


@app.exception_handler(edits.PatchError)
async def _patch_error_handler(_request, exc: edits.PatchError):
    return JSONResponse(status_code=400, content={"detail": _sanitize_detail(exc)})


@app.exception_handler(ai.AIServiceError)
async def _ai_error_handler(_request, exc: ai.AIServiceError):
    return JSONResponse(status_code=502, content={"detail": _sanitize_detail(exc)})


@app.exception_handler(render.RenderError)
async def _render_error_handler(_request, exc: render.RenderError):
    log.exception("render failed")
    return JSONResponse(status_code=500, content={"detail": f"渲染失败：{_sanitize_detail(exc)}"})


@app.exception_handler(render.RenderBusy)
async def _render_busy_handler(_request, exc: render.RenderBusy):
    return JSONResponse(status_code=429, content={"detail": "渲染队列已满，请稍后重试"})


@app.exception_handler(FileNotFoundError)
async def _not_found_handler(_request, exc: FileNotFoundError):
    # OS-level errors can embed absolute server paths; keep those server-side
    log.info("not found: %s", exc)
    return JSONResponse(status_code=404, content={"detail": _sanitize_detail(exc) or "not found"})


@app.exception_handler(ValueError)
async def _value_error_handler(_request, exc: ValueError):
    return JSONResponse(status_code=400, content={"detail": _sanitize_detail(exc)})


def _version_id(meta: Dict[str, Any]) -> str:
    return f"v{int(meta.get('version', 0))}"


def _pending_from_messages(messages: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Newest un-applied assistant operations, minus any already applied through
    partial applies. Served explicitly so staged ops survive message trimming —
    they used to silently disappear once chat history scrolled past the window."""
    for message in reversed(messages):
        if message.get("role") != "assistant":
            continue
        operations = message.get("operations")
        if not isinstance(operations, list) or not operations:
            continue
        if message.get("applied_version"):
            return None
        applied_ids = {
            str(sid)
            for apply_record in (message.get("partial_applies") or [])
            if isinstance(apply_record, dict)
            for sid in (apply_record.get("applied_shape_ids") or [])
        }
        remaining = [
            op
            for op in operations
            if not (isinstance(op, dict) and op.get("shape_id") and str(op["shape_id"]) in applied_ids)
        ]
        if not remaining:
            return None
        return {
            "message_id": message.get("id"),
            "base_version": message.get("base_version"),
            "operations": remaining,
        }
    return None


def _serialize(meta: Dict[str, Any], objects_payload: Dict[str, Any]) -> Dict[str, Any]:
    doc_id = meta["doc_id"]
    version = int(meta.get("version", 0))
    messages = storage.load_messages(doc_id)
    return {
        "doc_id": doc_id,
        "title": meta.get("title") or Path(str(meta.get("file_name") or "presentation.pptx")).stem,
        "file_name": meta.get("file_name", "presentation.pptx"),
        "slide_count": int(meta.get("slide_count", objects_payload.get("slide_count", 0))),
        "slide_size_emu": objects_payload.get("slide_size_emu") or meta.get("slide_size_emu"),
        "version": version,
        "version_id": f"v{version}",
        "created_at": meta.get("created_at"),
        "updated_at": meta.get("updated_at"),
        "versions": meta.get("versions", []),
        "slides": objects_payload.get("slides", []),
        "objects": objects_payload.get("objects", []),
        "elements": objects_payload.get("elements", []),
        "tables": objects_payload.get("tables", []),
        "image_url_base": f"/api/docs/{doc_id}/slides",
        "download_url": f"/api/docs/{doc_id}/download",
        "messages": messages[-60:],
        "message_count": len(messages),
        "notes": objects_payload.get("notes", []),
        "pending_operations": _pending_from_messages(messages),
        "ai_configured": bool(os.getenv("DEEPSEEK_API_KEY", "").strip()),
        "ai_model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
        "render_scale": RENDER_SCALE,
        "max_upload_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
        # E21：本版渲染是否降级为旧页图（部分页渲染中）
        "render_partial": bool(objects_payload.get("render_partial")),
        # v0.4：宏文档只读标记
        "has_macro": bool(meta.get("has_macro")),
        "macro_warning": (
            "含宏的文件已转为只读预览（检测到 vbaProject，防宏病毒）"
            if meta.get("has_macro") else None
        ),
        "password_protected": bool(meta.get("password_hash")),
    }


def _load_doc(doc_id: str):
    meta = storage.load_meta(doc_id)
    objects_payload = storage.load_objects(doc_id)
    return meta, objects_payload


def _sse(event: str, data: Dict[str, Any]) -> str:
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _content_hash(path: Path) -> str:
    import hashlib

    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _find_cached_render(doc_id: str, digest: str) -> Optional[Path]:
    """内容指纹相同的完整渲染目录（无降级标记、有 PDF 与页面图）。"""
    base = storage.doc_dir(doc_id) / "render"
    if not base.exists():
        return None
    for item in sorted(base.iterdir()):
        if not item.is_dir() or not re.fullmatch(r"v\d+", item.name):
            continue
        marker = item / ".contenthash"
        if not marker.exists() or marker.read_text(encoding="utf-8").strip() != digest:
            continue
        if render.partial_marker(item).exists():
            continue  # 降级目录：页图属于上一版，不能复用
        if not (item / render.EXPORT_PDF_NAME).exists():
            continue
        if not list(item.glob("slide-*.png")):
            continue
        return item
    return None


def _render_with_cache(
    doc_id: str, work_path: Path, render_dir: Path, stale_dir: Optional[Path] = None
) -> int:
    """v0.4 D17：相同内容（restore/重复操作）直接复用已渲染 PNG，不再排队渲染。"""
    digest = _content_hash(work_path)
    cached = _find_cached_render(doc_id, digest)
    if cached is not None and cached != render_dir:
        render_dir.mkdir(parents=True, exist_ok=True)
        for item in cached.iterdir():
            if item.is_file() and item.name != ".contenthash":
                shutil.copy2(item, render_dir / item.name)
        (render_dir / ".contenthash").write_text(digest, encoding="utf-8")
        log.info("render cache hit: %s reuses %s (%.12s)", render_dir.name, cached.name, digest)
        return len(list(render_dir.glob("slide-*.png")))
    page_count = render.render_pptx(work_path, render_dir, scale=RENDER_SCALE, stale_dir=stale_dir)
    (render_dir / ".contenthash").write_text(digest, encoding="utf-8")
    return page_count


def _prewarm_doc(doc_id: str, version: Optional[int] = None) -> int:
    """v0.4 D18：后台补齐当前版本缺失的缩略图（上传后/启动时自愈）。"""
    try:
        meta = storage.load_meta(doc_id)
        if version is None:
            version = int(meta.get("version", 0))
        filled = render.ensure_thumbs(storage.render_dir(doc_id, version))
        if filled:
            log.info("prewarm: filled %d missing thumb(s) for %s v%d", filled, doc_id, version)
        return filled
    except Exception:
        log.exception("prewarm failed for %s", doc_id)
        return 0


def _prewarm_doc_async(doc_id: str) -> None:
    threading.Thread(
        target=_prewarm_doc, args=(doc_id,), daemon=True, name=f"prewarm-{doc_id[:8]}"
    ).start()


def _prepare_commit(doc_id: str, work_path: Path, new_version: int):
    """Render + parse the work file WITHOUT holding the doc lock.

    LibreOffice can take up to 180s per document; doing this inside the lock
    blocked delete/rename/apply/chat on the doc (and the global conversion lock
    blocked every other doc). Returns the prepared payloads for the commit.
    """
    render_dir = storage.render_dir(doc_id, new_version)
    try:
        # E21 超时降级：把上一版渲染目录作为 stale 来源——worker 卡死时
        # 复用旧页图继续提交，而不是整单失败
        try:
            current_version = int(storage.load_meta(doc_id).get("version", 0))
        except (OSError, ValueError):
            current_version = new_version - 1
        stale_dir = (
            storage.render_dir(doc_id, current_version)
            if current_version >= 0 and current_version != new_version
            else None
        )
        page_count = _render_with_cache(doc_id, work_path, render_dir, stale_dir=stale_dir)
        _objs, objects_payload = parse.parse_pptx(work_path)
        del _objs
        if page_count != objects_payload.get("slide_count"):
            page_count = objects_payload.get("slide_count", page_count)
        if render.partial_marker(render_dir).exists():
            objects_payload["render_partial"] = True
        return render_dir, objects_payload, int(page_count)
    except Exception:
        shutil.rmtree(render_dir, ignore_errors=True)
        raise  # caller owns work_path cleanup


def _discard_prepared(work_path: Path, render_dir: Optional[Path]) -> None:
    try:
        work_path.unlink()
    except OSError:
        pass
    if render_dir is not None:
        shutil.rmtree(render_dir, ignore_errors=True)


def _commit_result(committed: Dict[str, Any], objects_payload: Dict[str, Any], touched: Optional[List[int]] = None) -> Dict[str, Any]:
    result = _serialize(committed, objects_payload)
    if touched is not None:
        result["touched_slides"] = sorted(set(touched))
    _publish_changed(result)
    return result


def _publish_changed(result: Dict[str, Any]) -> None:
    """向 watch 订阅者广播版本变更。

    所有提交路径（apply / chat 自动应用 / undo / restore / merge）都经
    _commit_result 汇合，在这里发一次即可全覆盖；推送失败不影响提交。"""
    doc_id = str(result.get("doc_id") or "")
    if not doc_id:
        return
    event: Dict[str, Any] = {
        "type": "changed",
        "version": int(result.get("version") or 0),
        "version_id": result.get("version_id"),
    }
    if result.get("touched_slides") is not None:
        event["touched_slides"] = result["touched_slides"]
    if result.get("slide_count") is not None:
        event["slide_count"] = int(result["slide_count"])
    try:
        broker.publish(doc_id, event)
    except Exception:
        log.warning("watch event publish failed for %s", doc_id, exc_info=True)


def _finalize_chat(
    doc_id: str,
    auto_apply: bool,
    start_version: int,
    done_data: Dict[str, Any],
) -> Tuple[Optional[Dict[str, Any]], Optional[str], Dict[str, Any]]:
    """Persist the assistant message, optionally apply operations, and build done fields."""
    operations = done_data.get("operations")
    if not isinstance(operations, list):
        operations = []
    trace = done_data.get("trace")
    if not isinstance(trace, list):
        trace = []
    plan = done_data.get("plan")
    if isinstance(plan, list):
        plan = [str(step)[:200] for step in plan][:12]
    else:
        plan = []
    summary = str(done_data.get("summary") or "")
    reply = str(done_data.get("reply") or "")
    question = str(done_data.get("question") or "")
    error = done_data.get("error")
    model = str(done_data.get("model") or os.getenv("DEEPSEEK_MODEL", "deepseek-chat"))
    content = reply or summary or question or "本轮处理完成。"
    if question and question not in content:
        content = f"{content}\n{question}" if content else question

    assistant_id = uuid.uuid4().hex[:12]
    assistant_at = storage.now_iso()
    applied_event: Optional[Dict[str, Any]] = None
    apply_error: Optional[str] = None
    applied = False

    # ---- phase 1 (short lock): persist the message, decide eligibility ----
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        assistant_message = {
            "id": assistant_id,
            "role": "assistant",
            "content": content,
            "at": assistant_at,
            "kind": "chat",
            "mode": "agent",
            "model": model,
            "summary": summary,
            "reply": reply,
            "question": question,
            "trace": trace,
            "operations": operations,
            "auto_apply": bool(auto_apply),
            "version": f"v{int(start_version)}",
            "base_version": int(start_version),
            "steps": int(done_data.get("steps") or 0),
            "finished": bool(done_data.get("finished")),
            "error": str(error) if error else None,
            # 多步计划（plan 工具）随消息落盘：刷新页面后聊天区仍能看到「我将：1… 2… 3…」
            "plan": plan or None,
        }
        storage.append_messages(doc_id, [assistant_message])

        eligible = False
        if auto_apply and operations:
            if int(meta.get("version", 0)) != int(start_version):
                apply_error = "文档在对话期间已更新，操作已暂存，请刷新后重试"
            elif not AUTO_APPLY_DESTRUCTIVE and any(
                str(op.get("action")) in DESTRUCTIVE_ACTIONS for op in operations
            ):
                # deletion/clear can't be undone by eye on the canvas; force the
                # review path unless explicitly enabled in .env
                apply_error = "包含删除/清空类操作，已转为手动确认；请在操作卡片中检查后应用"
            else:
                eligible = True

    # ---- phase 2 (no lock): apply edits to the work file, render, parse ----
    prepared: Optional[Tuple[int, Path, Path, Dict[str, Any], int, List[int]]] = None
    if eligible:
        new_version = int(start_version) + 1
        work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
        render_dir: Optional[Path] = None
        try:
            touched = edits.apply_patches_to_file(
                storage.current_path(doc_id), work_path, operations
            )
            render_dir, objects_payload, page_count = _prepare_commit(doc_id, work_path, new_version)
            prepared = (new_version, work_path, render_dir, objects_payload, page_count, sorted(touched))
        except Exception as exc:
            try:
                work_path.unlink()
            except OSError:
                pass
            if isinstance(exc, render.RenderBusy):
                apply_error = "渲染队列已满，操作已暂存，请稍后手动应用"
            else:
                apply_error = str(exc) or "应用操作失败"

    # ---- phase 3 (short lock): re-check version, commit, mark applied ----
    if prepared is not None:
        new_version, work_path, render_dir, objects_payload, page_count, touched = prepared
        with storage.lock_for(doc_id):
            meta = storage.load_meta(doc_id)
            if int(meta.get("version", 0)) != int(start_version):
                _discard_prepared(work_path, render_dir)
                apply_error = "文档在对话期间已更新，操作已暂存，请刷新后重试"
            else:
                committed = storage.commit_work(
                    doc_id=doc_id,
                    meta=meta,
                    work_path=work_path,
                    new_version=new_version,
                    label=f"AI 对话改稿 · {len(operations)} 项",
                    slide_count=int(page_count),
                    slide_size=objects_payload.get("slide_size_emu")
                    or meta.get("slide_size_emu"),
                    objects_payload=objects_payload,
                )
                applied = True
                applied_doc = _commit_result(committed, objects_payload, touched)
                apply_message = {
                    "id": uuid.uuid4().hex[:12],
                    "role": "system",
                    "content": (
                        f"AI 对话改稿 · 已应用 {len(operations)} 条修改 · "
                        f"保存为 {applied_doc['version_id']}"
                    ),
                    "at": storage.now_iso(),
                    "kind": "apply",
                    "version": applied_doc["version_id"],
                    "patch_count": len(operations),
                    "mode": "agent",
                }
                try:
                    messages = storage.load_messages(doc_id)
                    for item in messages:
                        if item.get("id") == assistant_id:
                            item["applied_version"] = applied_doc["version_id"]
                    messages.append(apply_message)
                    storage.save_messages(doc_id, messages)
                    messages = messages[-120:]
                    applied_doc["messages"] = messages
                    applied_doc["message_count"] = len(messages)
                except Exception:
                    log.exception("append apply message failed for %s", doc_id)
                storage.append_audit(
                    action="chat_apply", doc_id=doc_id,
                    summary=f"AI 对话改稿 · {len(operations)} 项",
                    op_count=len(operations), ok=True,
                    version=applied_doc.get("version_id"),
                )
                applied_event = {
                    "version_id": applied_doc.get("version_id"),
                    "version": applied_doc.get("version"),
                    "operation_count": len(operations),
                    "operations_applied": len(operations),
                    "touched_slides": applied_doc.get("touched_slides") or [],
                    "doc": applied_doc,
                    "message_id": apply_message["id"],
                    "message_count": int(applied_doc.get("message_count") or 0),
                    "assistant_message_id": assistant_id,
                }

    meta_after = storage.load_meta(doc_id)
    messages_after = storage.load_messages(doc_id)
    version_id = _version_id(meta_after)
    fields: Dict[str, Any] = {
        "message_id": assistant_id,
        "at": assistant_at,
        "content": content,
        "summary": summary,
        "reply": reply,
        "question": question,
        "trace": trace,
        "operations": operations,
        "version_id": version_id,
        "base_version": int(start_version),
        "message_count": len(messages_after),
        "applied": applied,
        "error": error,
        "apply_error": apply_error,
        "mode": "agent",
        "model": model,
        "auto_apply": bool(auto_apply),
        "steps": int(done_data.get("steps") or 0),
        "finished": bool(done_data.get("finished")),
    }
    if applied_event is not None:
        fields["applied_version_id"] = applied_event["version_id"]
        fields["touched_slides"] = applied_event["touched_slides"]
    return applied_event, apply_error, fields


@app.get("/api/health")
def health():
    return {
        "ok": True,
        "service": "ppt-light",
        "version": "0.5.0",
        "ai_configured": bool(os.getenv("DEEPSEEK_API_KEY", "").strip()),
        "ai_model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
        "libreoffice": bool(shutil.which("soffice") or shutil.which("libreoffice")),
        "docs": len(storage.list_docs()),
        "max_upload_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
    }


@app.get("/api/docs")
def list_documents():
    return {"docs": storage.list_docs()}


@app.get("/api/sessions")
def list_sessions():
    return {"sessions": storage.list_docs()}


@app.post("/api/docs/upload")
def upload_document(file: UploadFile = File(...)):
    original = Path(file.filename or "presentation.pptx").name
    safe_name = re.sub(r"[\x00-\x1f]", "", original).strip() or "presentation.pptx"
    # E22：旧版二进制格式给出明确去向，而不是笼统的“只支持 .pptx”
    suffix = safe_name.lower().rsplit(".", 1)[-1] if "." in safe_name else ""
    if suffix in {"ppt", "pptm", "pps", "ppsm"}:
        raise HTTPException(
            status_code=400,
            detail=f".{suffix} 是旧版/宏格式：请先在 PowerPoint 中转存为 .pptx 再上传",
        )
    if not safe_name.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="只支持 .pptx 文件")
    data = file.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=(
                f"文件超过 {MAX_UPLOAD_BYTES // (1024 * 1024)} MB 限制，"
                "请压缩图片/媒体后重试"
            ),
        )
    if len(data) == 0:
        raise HTTPException(status_code=400, detail="上传的文件是空的，请检查后重新选择")
    if len(data) < 100:
        raise HTTPException(
            status_code=400, detail="文件过小或已损坏：请确认导出的是有效 .pptx"
        )

    storage.ensure_dirs()
    tmp_dir = storage.DATA_DIR / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = tmp_dir / f"{uuid.uuid4().hex}.pptx"
    tmp_path.write_bytes(data)
    doc_id: Optional[str] = None
    try:
        try:
            _, parsed = parse.parse_pptx(tmp_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"无法解析 PPTX：{exc}") from exc
        if not parsed.get("slide_count"):
            raise HTTPException(status_code=400, detail="PPTX 中没有幻灯片")
        if parsed.get("truncated_slides"):
            raise HTTPException(
                status_code=413,
                detail=f"页数超过解析上限（{os.getenv('PARSE_MAX_SLIDES', '500')} 页），已拒绝",
            )
        if parsed.get("truncated_shapes"):
            raise HTTPException(
                status_code=413,
                detail=f"对象数超过解析上限（{os.getenv('PARSE_MAX_SHAPES', '5000')} 个），已拒绝",
            )
        has_macro = False
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                has_macro = any(
                    name == "ppt/vbaProject.bin" or name.endswith("/vbaProject.bin")
                    for name in zf.namelist()
                )
        except zipfile.BadZipFile:
            pass
        meta = storage.create_doc(
            file_name=safe_name,
            content=data,
            slide_count=int(parsed["slide_count"]),
            slide_size=parsed["slide_size_emu"],
        )
        if has_macro:
            # 宏文档：保留预览/下载，编辑入口由 _assert_editable 拒绝
            meta["has_macro"] = True
            storage.save_meta(meta["doc_id"], meta)
            log.warning("macro detected in %s (%s) — read-only preview", safe_name, meta["doc_id"])
        doc_id = meta["doc_id"]
        storage.save_objects(doc_id, parsed)
        # Wingdings 乱码防护：macOS 无 Wingdings 字体，私用区符号在 LibreOffice
        # 渲染里变豆腐/乱码。上传即把 buFont/buChar 换成 Unicode 等价符号。
        try:
            from services.garble_fix import sanitize_wingdings
            b_fixed, t_fixed = sanitize_wingdings(storage.current_path(doc_id))
            if b_fixed or t_fixed:
                log.info("wingdings sanitized: %d bullets, %d chars", b_fixed, t_fixed)
        except Exception as exc:  # 修复失败不阻塞上传
            log.warning("wingdings sanitize skipped: %s", exc)
        _render_with_cache(doc_id, storage.current_path(doc_id), storage.render_dir(doc_id, 0))
        # v0 的缓存指纹以版本快照为准：sanitize_wingdings 会无条件重存 current.pptx，
        # 而 restore 的来源是 versions/v0.pptx，两者字节可能不同
        (storage.render_dir(doc_id, 0) / ".contenthash").write_text(
            _content_hash(storage.versions_dir(doc_id) / "v0.pptx"), encoding="utf-8"
        )
        _prewarm_doc_async(doc_id)  # D18：响应即刻返回，缩略图缺口后台补齐
        return _serialize(meta, parsed)
    except Exception:
        if doc_id:
            shutil.rmtree(storage.doc_dir(doc_id), ignore_errors=True)
            storage.release_lock(doc_id)
        raise
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass


@app.get("/api/docs/{doc_id}")
def get_document(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    meta, objects_payload = _load_doc(doc_id)
    _assert_doc_open(meta, x_doc_password)
    return _serialize(meta, objects_payload)


@app.get("/api/docs/{doc_id}/objects")
def get_objects(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    """轻量对象清单：AI 补 alt_text/查 shape_id 时不必拉全量 doc 序列化。"""
    meta, objects_payload = _load_doc(doc_id)
    _assert_doc_open(meta, x_doc_password)
    return {
        "objects": objects_payload.get("objects", []),
        "elements": objects_payload.get("elements", []),
        "tables": objects_payload.get("tables", []),
        "slide_count": int(meta.get("slide_count", 0)),
        "version_id": _version_id(meta),
    }


@app.patch("/api/docs/{doc_id}")
def rename_document(doc_id: str, body: RenameIn):
    with storage.lock_for(doc_id):
        meta = storage.touch_doc(doc_id, title=body.title)
        objects_payload = storage.load_objects(doc_id)
        return _serialize(meta, objects_payload)


@app.get("/api/docs/{doc_id}/messages")
def get_messages(doc_id: str):
    storage.load_meta(doc_id)
    return {"messages": storage.load_messages(doc_id)}


@app.delete("/api/docs/{doc_id}")
def delete_document(doc_id: str):
    with storage.lock_for(doc_id):
        storage.delete_doc(doc_id)
    return {"ok": True}


@app.post("/api/docs/{doc_id}/chat")
def chat(doc_id: str, body: ChatIn, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    message = body.message.strip()
    if not message:
        raise HTTPException(status_code=400, detail="请输入修改要求")
    meta, objects_payload = _load_doc(doc_id)
    _assert_doc_open(meta, x_doc_password)
    _assert_editable(meta)
    total_slides = int(meta.get("slide_count") or objects_payload.get("slide_count") or 0)
    if total_slides <= 0:
        raise HTTPException(status_code=400, detail="文档没有可编辑页面")

    raw_scope = (body.scope or "current").strip().lower()
    if raw_scope in ("current", "slide"):
        scope = "current"
    elif raw_scope in ("deck", "all", "document", "whole"):
        scope = "deck"
    else:
        raise HTTPException(status_code=400, detail="scope 仅支持 current 或 deck")

    slide_index = int(body.slide_index) if body.slide_index is not None else 0
    if slide_index < 0 or slide_index >= total_slides:
        raise HTTPException(status_code=400, detail="页码超出范围")
    selected_shape_id = (body.selected_shape_id or "").strip() or None
    requested_state, requested_number = ai.locate_requested_slide(message, total_slides)
    if requested_state == "out_of_range":
        raise HTTPException(
            status_code=400,
            detail=f"指令提到的页码（第 {requested_number} 页）超出范围：文档共 {total_slides} 页",
        )
    if requested_state == "ok" and scope != "deck":
        # 点名页面只收窄「当前页」范围；显式的整份范围不降级（与 agent.run_stream 一致）
        scope = "current"
        slide_index = requested_number
    scope_label = f"第 {slide_index + 1} 页" if scope == "current" else "整份文稿"

    # concurrency gates before anything is persisted
    if not _chat_slots.acquire(blocking=False):
        raise HTTPException(status_code=429, detail="AI 并发已满，请稍后再试")
    with _chat_inflight_lock:
        if _chat_inflight.get(doc_id, 0) >= 1:
            _chat_slots.release()
            raise HTTPException(status_code=429, detail="该文档已有对话进行中，请等待完成")
        _chat_inflight[doc_id] = _chat_inflight.get(doc_id, 0) + 1

    user_id = uuid.uuid4().hex[:12]
    user_at = storage.now_iso()
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        objects_payload = storage.load_objects(doc_id)
        total_slides = int(
            meta.get("slide_count") or objects_payload.get("slide_count") or total_slides
        )
        start_version = int(meta.get("version", 0))
        if body.base_version is not None and int(body.base_version) != start_version:
            raise HTTPException(status_code=409, detail="文档已经更新，请刷新后再发送")
        prior_messages = storage.load_messages(doc_id)
        history = [
            item
            for item in prior_messages
            if item.get("role") in ("user", "assistant")
            and str(item.get("content") or "").strip()
        ][-12:]
        user_message = {
            "id": user_id,
            "role": "user",
            "content": message,
            "at": user_at,
            "kind": "chat",
            "scope": scope_label,
            "slide_index": slide_index,
            "selected_shape_id": selected_shape_id,
            "auto_apply": bool(body.auto_apply),
        }
        storage.append_messages(doc_id, [user_message])
        meta = storage.touch_doc(doc_id)

    events: "queue.Queue" = queue.Queue()
    state: Dict[str, Any] = {}

    def emit(name: str, data: Dict[str, Any]) -> None:
        events.put((name, data))

    def _chat_worker() -> None:
        try:
            state["result"] = agent.run_stream(
                emit,
                objects=objects_payload.get("objects", []),
                elements=objects_payload.get("elements", []),
                tables=objects_payload.get("tables", []),
                notes=objects_payload.get("notes"),
                total_slides=total_slides,
                slide_size=objects_payload.get("slide_size_emu")
                or meta.get("slide_size_emu")
                or {},
                instruction=message,
                scope=scope,
                slide_index=slide_index,
                selected_shape_id=selected_shape_id,
                history=history,
                base_version=start_version,
                auto_apply=bool(body.auto_apply),
                undo_last=lambda base_version: _agent_undo(doc_id, base_version),
                vision_check=lambda slides: _agent_vision_check(doc_id, slides),
            )
        except Exception as exc:
            message_text = str(exc) or "Agent 执行失败"
            log.exception("agent stream failed for %s", doc_id)
            emit("error", {"message": message_text})
            state["result"] = {
                "mode": "agent",
                "model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
                "scope": scope,
                "scope_slides": [slide_index] if scope == "current" else list(range(total_slides)),
                "total_slides": total_slides,
                "summary": "AI 执行中断",
                "reply": message_text,
                "question": "",
                "operations": [],
                "trace": [],
                "steps": 0,
                "finished": False,
                "dropped": 0,
                "error": message_text,
                "base_version": start_version,
                "auto_apply": bool(body.auto_apply),
            }
        # Finalize inside the worker thread: a client disconnect (Stop button,
        # tab close, proxy cut) abandons the response generator, and any logic
        # living in the generator after the read loop would never run. Doing it
        # here guarantees the assistant message — and with auto_apply the
        # already-validated operations — are always persisted.
        result = state.get("result")
        if not isinstance(result, dict):
            result = {}
        try:
            if body.auto_apply and result.get("operations"):
                emit(
                    "applying",
                    {
                        "count": len(result.get("operations") or []),
                        "version_id": f"v{start_version + 1}",
                    },
                )
            applied_event, apply_error, final_fields = _finalize_chat(
                doc_id, bool(body.auto_apply), start_version, result
            )
            if applied_event is not None:
                emit("applied", applied_event)
            elif apply_error:
                emit(
                    "apply_error",
                    {
                        "message": apply_error,
                        "staged": True,
                        "version_id": final_fields["version_id"],
                    },
                )
            result.update(final_fields)
            log.info(
                "chat doc=%s steps=%s ops=%s applied=%s",
                doc_id,
                result.get("steps"),
                len(result.get("operations") or []),
                result.get("applied"),
            )
            emit("done", result)
        except Exception as exc:
            log.exception("chat finalize failed for %s", doc_id)
            message_text = f"保存对话失败：{exc}"
            emit("error", {"message": message_text})
            result.update(
                {
                    "message_id": None,
                    "content": result.get("reply") or result.get("summary") or message_text,
                    "version_id": f"v{start_version}",
                    "message_count": 0,
                    "applied": False,
                    "error": result.get("error") or message_text,
                    "apply_error": None,
                    "mode": "agent",
                    "model": result.get("model")
                    or os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
                    "auto_apply": bool(body.auto_apply),
                }
            )
            emit("done", result)
        finally:
            events.put(None)
            with _chat_inflight_lock:
                remaining = _chat_inflight.get(doc_id, 1) - 1
                if remaining <= 0:
                    _chat_inflight.pop(doc_id, None)
                else:
                    _chat_inflight[doc_id] = remaining
            _chat_slots.release()

    threading.Thread(target=_chat_worker, daemon=True, name=f"agent-{doc_id[:8]}").start()

    def event_stream():
        # Pure relay: all state transitions (applying/applied/done) are emitted
        # by the worker, so a disconnect can't skip persistence work.
        while True:
            item = events.get()
            if item is None:
                break
            name, data = item
            yield _sse(name, data)

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
            "Content-Encoding": "identity",
        },
    )


@app.post("/api/docs/{doc_id}/apply")
def apply(doc_id: str, body: ApplyIn, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    label = (body.label or "编辑").strip()[:80] or "编辑"
    _assert_doc_open(storage.load_meta(doc_id), x_doc_password)
    _assert_editable(storage.load_meta(doc_id))
    try:
        result = _apply_locked(doc_id, body, label)
    except Exception as exc:
        storage.append_audit(
            action="apply", doc_id=doc_id, summary=label,
            op_count=len(body.patches), ok=False,
            error=str(exc)[:200],
        )
        raise
    storage.append_audit(
        action="apply", doc_id=doc_id, summary=label,
        op_count=len(body.patches), ok=True, version=result.get("version_id"),
    )
    return result


def _apply_locked(doc_id: str, body: ApplyIn, label: str):
    # ---- phase 1 (short lock): snapshot the optimistic base ----
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        if body.base_version is not None and int(body.base_version) != int(meta.get("version", 0)):
            raise edits.PatchConflict("文档已经更新，请刷新后再试")
        entry_version = int(meta.get("version", 0))
        new_version, _ = storage.next_version_id(meta)

    # ---- phase 2 (no lock): apply patches, render, parse ----
    work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
    render_dir: Optional[Path] = None
    try:
        touched = edits.apply_patches_to_file(
            storage.current_path(doc_id), work_path, body.patches
        )
        render_dir, objects_payload, page_count = _prepare_commit(doc_id, work_path, new_version)
    except Exception:
        try:
            work_path.unlink()
        except OSError:
            pass
        raise

    # ---- phase 3 (short lock): re-check base, commit, record messages ----
    try:
        with storage.lock_for(doc_id):
            committed, objects_payload = _commit_prepared_locked(
                doc_id, new_version, label, work_path, render_dir, objects_payload,
                page_count, expected_version=entry_version,
            )
            result = _commit_result(committed, objects_payload, sorted(touched))
            # v0.5 #7：apply 后自动后置 lint，AI 拿 issues 决定是否再发 fix_issues
            from services import lint as lint_mod

            result["lint_issues"] = lint_mod.lint_pptx(storage.current_path(doc_id))
            # v0.6 #1：渲染验证闭环入口——触达页的视觉自检 URL（多模态模型看渲染图）
            result["vision_check_url"] = (
                f"/api/docs/{doc_id}/vision-check?slides="
                + ",".join(str(i) for i in sorted(set(touched)))
            )
            # v0.6 #10：交付健康报告引用——AI 每轮改完看「还差多少分」
            try:
                from services import health as health_mod

                report = health_mod.health_report(storage.current_path(doc_id))
                result["health"] = {
                    "score": report["score"],
                    "grade": report["grade"],
                    "url": f"/api/docs/{doc_id}/health",
                }
            except Exception:  # noqa: BLE001 — 报告失败不阻塞 apply
                pass
            # Mark exactly the message the patches came from. The legacy
            # label-prefix scan marked the newest un-applied assistant message even
            # when the user applied a different (or partial) set, silently
            # un-restoring ops that were never applied.
            mark_id = (body.assistant_message_id or "").strip()
            if mark_id or label.startswith("AI 改稿"):
                old_messages = storage.load_messages(doc_id)
                target = None
                if mark_id:
                    candidate = next(
                        (m for m in old_messages if m.get("id") == mark_id), None
                    )
                    if candidate is not None and candidate.get("role") == "assistant":
                        target = candidate
                else:
                    target = next(
                        (
                            m
                            for m in reversed(old_messages)
                            if m.get("role") == "assistant"
                            and m.get("operations")
                            and not m.get("applied_version")
                        ),
                        None,
                    )
                if target is not None and not target.get("applied_version"):
                    op_ids = {
                        str(op.get("shape_id"))
                        for op in (target.get("operations") or [])
                        if isinstance(op, dict) and op.get("shape_id")
                    }
                    submitted_ids = {
                        str(p.get("shape_id"))
                        for p in body.patches
                        if isinstance(p, dict) and p.get("shape_id")
                    }
                    if op_ids and op_ids <= submitted_ids:
                        target["applied_version"] = result["version_id"]
                        target["content"] = (
                            f"{target.get('content') or ''}（已应用为 {result['version_id']}）"
                        )
                    else:
                        target.setdefault("partial_applies", []).append(
                            {
                                "version_id": result["version_id"],
                                "applied_shape_ids": sorted(op_ids & submitted_ids),
                            }
                        )
                    storage.save_messages(doc_id, old_messages)
            message = {
                "id": uuid.uuid4().hex[:12],
                "role": "assistant" if label.startswith("AI 改稿") else "system",
                "content": f"{label} · 已应用 {len(body.patches)} 条修改 · 保存为 {result['version_id']}",
                "at": storage.now_iso(),
                "kind": "apply",
                "version": result["version_id"],
                "patch_count": len(body.patches),
            }
            result["messages"] = storage.append_messages(doc_id, [message])
            result["message_count"] = len(result["messages"])
    except edits.PatchConflict:
        _discard_prepared(work_path, render_dir)
        raise
    log.info("applied %d patches to %s -> v%d", len(body.patches), doc_id, new_version)
    return result


def _commit_prepared_locked(
    doc_id: str,
    new_version: int,
    label: str,
    work_path: Path,
    render_dir: Path,
    objects_payload: Dict[str, Any],
    page_count: int,
    expected_version: Optional[int] = None,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Commit prepared artifacts; caller already holds the doc lock.
    Raises PatchConflict on version drift (caller cleans up via _discard_prepared)."""
    meta = storage.load_meta(doc_id)
    if expected_version is not None and int(meta.get("version", 0)) != int(expected_version):
        raise edits.PatchConflict("文档已经更新，请刷新后再试")
    committed = storage.commit_work(
        doc_id=doc_id,
        meta=meta,
        work_path=work_path,
        new_version=new_version,
        label=label,
        slide_count=int(page_count),
        slide_size=objects_payload.get("slide_size_emu") or meta.get("slide_size_emu"),
        objects_payload=objects_payload,
    )
    return committed, objects_payload


@app.get("/api/docs/{doc_id}/versions")
def list_versions(doc_id: str):
    meta = storage.load_meta(doc_id)
    current = _version_id(meta)
    items = []
    for item in meta.get("versions", []):
        version_id = str(item.get("id"))
        match = re.fullmatch(r"v(\d+)", version_id)
        if not match:
            continue
        items.append(
            {
                **item,
                "is_current": version_id == current,
                "slide_count": meta.get("slide_count", 0),
                "image_url_base": f"/api/docs/{doc_id}/slides",
            }
        )
    return {"current": current, "versions": items}


@app.post("/api/docs/{doc_id}/versions/{version_id}/restore")
def restore_version(doc_id: str, version_id: str, body: Optional[RestoreIn] = None, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    _assert_doc_open(storage.load_meta(doc_id), x_doc_password)
    _assert_editable(storage.load_meta(doc_id))
    if not re.fullmatch(r"v\d+", version_id or ""):
        raise HTTPException(status_code=400, detail="invalid version id")
    # ---- phase 1 (short lock): allocate the version id ----
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        if body is not None and body.base_version is not None and int(body.base_version) != int(meta.get("version", 0)):
            raise edits.PatchConflict("文档已经更新，请刷新后再试")
        source = storage.version_file(doc_id, version_id)
        new_version, _ = storage.next_version_id(meta)

    # ---- phase 2 (no lock): copy + render + parse ----
    work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
    render_dir: Optional[Path] = None
    try:
        shutil.copy2(source, work_path)
        render_dir, objects_payload, page_count = _prepare_commit(doc_id, work_path, new_version)
    except Exception:
        try:
            work_path.unlink()
        except OSError:
            pass
        raise

    # ---- phase 3 (short lock): commit + record ----
    try:
        with storage.lock_for(doc_id):
            committed, objects_payload = _commit_prepared_locked(
                doc_id, new_version, f"恢复 {version_id}", work_path, render_dir,
                objects_payload, page_count,
            )
            result = _commit_result(committed, objects_payload)
            message = {
                "id": uuid.uuid4().hex[:12],
                "role": "system",
                "content": f"恢复了 {version_id}，并保存为 {result['version_id']}",
                "at": storage.now_iso(),
                "kind": "restore",
                "version": result["version_id"],
            }
            result["messages"] = storage.append_messages(doc_id, [message])
            result["message_count"] = len(result["messages"])
    except edits.PatchConflict:
        _discard_prepared(work_path, render_dir)
        raise
    storage.append_audit(
        action="restore", doc_id=doc_id, summary=f"恢复 {version_id}",
        op_count=1, ok=True, version=result.get("version_id"),
    )
    return result


def _previous_version_id(doc_id: str, meta: Dict[str, Any], current: int) -> Optional[str]:
    """Highest version strictly below ``current`` whose pptx still exists."""
    numbers: List[int] = []
    for item in meta.get("versions", []):
        match = re.fullmatch(r"v(\d+)", str(item.get("id") or ""))
        if match and int(match.group(1)) < current:
            numbers.append(int(match.group(1)))
    for number in sorted(numbers, reverse=True):
        if (storage.versions_dir(doc_id) / f"v{number}.pptx").exists():
            return f"v{number}"
    return None


def _perform_undo(doc_id: str, base_version: Optional[int] = None) -> Dict[str, Any]:
    """撤销上一次应用：把上一版本内容提交为新版本（历史只前进，可连续撤销）。

    端点与 AI 的 undo_last 工具共用这一条链路。异常语义：
    ValueError → 400（v0 无可撤销）、PatchConflict → 409（版本已被并发更新）。
    """
    # ---- phase 1 (short lock): validate + allocate version id ----
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        current = int(meta.get("version", 0))
        if current <= 0:
            raise ValueError("已在最初版本（v0），没有可撤销的操作")
        if base_version is not None and int(base_version) != current:
            raise edits.PatchConflict("文档已经更新，请刷新后再试")
        # 连续撤销走 undo 血统：当前版本本身是撤销产物时，锚点用它恢复到的
        # 那个版本，继续向前回退（undo-undo ≠ redo）
        anchor = current
        current_item = next(
            (it for it in meta.get("versions", []) if str(it.get("id")) == f"v{current}"),
            None,
        )
        lineage = str((current_item or {}).get("restored_from") or "")
        if re.fullmatch(r"v\d+", lineage):
            anchor = int(lineage[1:])
        prev_id = _previous_version_id(doc_id, meta, anchor)
        if prev_id is None:
            raise ValueError("找不到可回退的历史版本，无法撤销")
        source = storage.version_file(doc_id, prev_id)
        new_version, _ = storage.next_version_id(meta)

    # ---- phase 2 (no lock): copy + render + parse ----
    work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
    render_dir: Optional[Path] = None
    try:
        shutil.copy2(source, work_path)
        render_dir, objects_payload, page_count = _prepare_commit(doc_id, work_path, new_version)
    except Exception:
        try:
            work_path.unlink()
        except OSError:
            pass
        raise

    # ---- phase 3 (short lock): commit + record ----
    try:
        with storage.lock_for(doc_id):
            committed, objects_payload = _commit_prepared_locked(
                doc_id, new_version, f"撤销 v{current}（回到 {prev_id}）",
                work_path, render_dir, objects_payload, page_count,
                expected_version=current,
            )
            # 血统标记：undo 产出的版本记录 restored_from，供下一次连续撤销定位锚点
            for item in committed.get("versions", []):
                if str(item.get("id")) == f"v{new_version}":
                    item["restored_from"] = prev_id
                    storage.save_meta(doc_id, committed)
                    break
            result = _commit_result(committed, objects_payload)
            result["restored_from"] = prev_id
            message = {
                "id": uuid.uuid4().hex[:12],
                "role": "system",
                "content": f"已撤销最近一次修改，回到 {prev_id}，保存为 v{new_version}",
                "at": storage.now_iso(),
                "kind": "undo",
                "version": result["version_id"],
            }
            result["messages"] = storage.append_messages(doc_id, [message])
            result["message_count"] = len(result["messages"])
    except edits.PatchConflict:
        _discard_prepared(work_path, render_dir)
        raise
    return result


@app.post("/api/docs/{doc_id}/undo")
def undo_last(doc_id: str, body: Optional[RestoreIn] = None, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    """撤销上一轮已应用的修改（v0 无可撤销 → 400；并发版本冲突 → 409）。"""
    _assert_doc_open(storage.load_meta(doc_id), x_doc_password)
    _assert_editable(storage.load_meta(doc_id))
    try:
        result = _perform_undo(doc_id, body.base_version if body else None)
        storage.append_audit(
            action="undo", doc_id=doc_id, summary=f"撤销 → {result.get('restored_from')}",
            op_count=1, ok=True, version=result.get("version_id"),
        )
        return result
    except edits.PatchConflict as exc:  # PatchConflict 继承 ValueError，必须先接
        raise HTTPException(status_code=409, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def _agent_undo(doc_id: str, base_version: Optional[int]) -> Dict[str, Any]:
    """Compact undo result for the agent's undo_last tool observation.

    The full commit payload (objects/messages) would bloat the model's
    observation; the tool only needs to report what it rolled back to."""
    result = _perform_undo(doc_id, base_version)
    return {
        "restored": result.get("restored_from"),
        "new_version": result.get("version_id"),
        "label": f"撤销上一轮 · 回到 {result.get('restored_from')}",
    }


def _resolve_slides_for_vision(doc_id: str, meta: Dict[str, Any], slides: Optional[List[int]]) -> List[int]:
    """v0.6 #1：归一化自检页码（0-based，去重排序，封顶 MAX_PAGES_PER_CALL）。"""
    total = int(meta.get("slide_count", 0))
    if slides:
        wanted = []
        for item in slides:
            try:
                index = int(item)
            except (TypeError, ValueError):
                continue
            if 0 <= index < total and index not in wanted:
                wanted.append(index)
        return sorted(wanted)[: vision.MAX_PAGES_PER_CALL]
    return list(range(min(total, vision.MAX_PAGES_PER_CALL)))


def _vision_check_pages(doc_id: str, meta: Dict[str, Any], slides: Optional[List[int]]) -> Dict[str, Any]:
    """对当前（或指定）版本的渲染 PNG 逐页跑多模态自检。

    FileNotFoundError（页图未渲染）与 VisionError（模型不可用）原样上抛，
    由全局 exception handler 转 404 / 502。"""
    version = int(meta.get("version", 0))
    wanted = _resolve_slides_for_vision(doc_id, meta, slides)
    if not wanted:
        raise HTTPException(status_code=400, detail="文档没有可自检的页面")
    pages = [
        vision.check_page_png(storage.image_path(doc_id, version, index), index + 1)
        for index in wanted
    ]
    return {
        "doc_id": doc_id,
        "version_id": f"v{version}",
        "model": pages[0].get("model") if pages else vision._model(),
        "pages": pages,
        "issue_count": sum(len(p.get("issues") or []) for p in pages),
        "summary": vision.summarize(pages),
    }


def _agent_vision_check(doc_id: str, slides: Optional[List[int]]) -> Dict[str, Any]:
    """vision_check 工具的后端回调：出错时返回 error 字段（observation 可读）。"""
    try:
        meta = storage.load_meta(doc_id)
        return _vision_check_pages(doc_id, meta, slides)
    except HTTPException as exc:
        return {"error": str(exc.detail)}
    except vision.VisionError as exc:
        return {"error": str(exc)}
    except FileNotFoundError:
        return {"error": "渲染图尚未生成，请稍后重试"}


@app.get("/api/docs/{doc_id}/vision-check")
def vision_check_document(
    doc_id: str,
    slides: Optional[str] = None,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    """v0.6 #1：渲染验证闭环——把渲染 PNG 喂多模态模型自检（溢出/乱码/塌版/对比度）。

    slides：0-based 页码逗号列表（如 0,1,2）；缺省自检前 6 页。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    indices: Optional[List[int]] = None
    if slides:
        indices = []
        for token in str(slides).split(","):
            token = token.strip()
            if not token:
                continue
            try:
                index = int(token)
            except ValueError:
                raise HTTPException(status_code=400, detail=f"无效页码：{token}") from None
            if index < 0 or index >= int(meta.get("slide_count", 0)):
                raise HTTPException(status_code=400, detail=f"页码超出范围：{index}（文档共 {meta.get('slide_count')} 页）")
            indices.append(index)
    return _vision_check_pages(doc_id, meta, indices)


@app.get("/api/docs/{doc_id}/health")
def doc_health(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    """v0.6 #10：交付健康报告——文档级质量评分卡（lint/乱码/字体统一度 + 统计）。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    from services import health as health_mod

    report = health_mod.health_report(storage.current_path(doc_id))
    return {"doc_id": doc_id, "version_id": _version_id(meta), **report}


# ---------- v0.6 #8：整稿重主题预览安全网 ----------


@app.post("/api/docs/{doc_id}/theme-preview")
def theme_preview_doc(
    doc_id: str,
    body: ThemePreviewIn,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    """v0.6 #8：重主题大改动先看效果——副本上应用 patches 渲染前后对比图。

    真文档一字节不动（版本号/current.pptx 都不碰），确认后照常 POST /apply。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    from services import theme_preview as theme_preview_mod

    previews_root = storage.doc_dir(doc_id) / "theme-previews"
    token = uuid.uuid4().hex[:12]
    try:
        result = theme_preview_mod.build_theme_preview(
            storage.current_path(doc_id),
            body.patches,
            previews_root / token,
            before_dir=storage.render_dir(doc_id, int(meta.get("version", 0))),
            pages=body.slides,
        )
    except theme_preview_mod.PreviewError:
        raise  # ValueError handler → 400
    # 预览产物可再生，只留最近 3 份防盘膨胀
    try:
        stale = sorted(previews_root.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True)[3:]
        for item in stale:
            shutil.rmtree(item, ignore_errors=True)
    except OSError:
        pass
    base = f"/api/docs/{doc_id}/theme-previews/{token}"
    return {
        "doc_id": doc_id,
        "version_id": _version_id(meta),  # 未变——真文档没动过
        "token": token,
        "page_count": result["page_count"],
        "pages": [
            {
                "slide_index": page["slide_index"],
                "before_url": f"{base}/before-{page['slide_index'] + 1:03d}.png",
                "after_url": f"{base}/after-{page['slide_index'] + 1:03d}.png",
            }
            for page in result["pages"]
        ],
        "hint": "效果确认后照常 POST /api/docs/%s/apply 应用到真文档" % doc_id,
    }


@app.get("/api/docs/{doc_id}/theme-previews/{token}/{name}")
def theme_preview_image(
    doc_id: str,
    token: str,
    name: str,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    """v0.6 #8：对比图分发——文件名白名单 + token 格式校验，防路径穿越。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    from services import theme_preview as theme_preview_mod

    if not re.fullmatch(r"[0-9a-f]{8,32}", token) or not theme_preview_mod.PREVIEW_FILE_RE.match(name):
        raise HTTPException(status_code=404, detail="not found")
    path = storage.doc_dir(doc_id) / "theme-previews" / token / name
    if not path.is_file():
        raise HTTPException(status_code=404, detail="not found")
    return FileResponse(path, media_type="image/png", headers={"Cache-Control": "no-store"})


@app.get("/api/docs/{doc_id}/download")
def download(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    path = storage.current_path(doc_id)
    if not path.exists():
        raise FileNotFoundError("document file not found")
    name = meta.get("file_name") or "presentation.pptx"
    if not name.lower().endswith(".pptx"):
        name = f"{name}.pptx"
    return FileResponse(
        path,
        media_type=PPTX_MEDIA_TYPE,
        filename=name,
        headers={"Cache-Control": "no-store"},
    )


def _export_base_name(meta: Dict[str, Any]) -> str:
    """Header-safe download basename: prefer the (possibly renamed) title."""
    name = str(
        meta.get("title")
        or Path(str(meta.get("file_name") or "presentation.pptx")).stem
    )
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "-", name).strip()
    return name[:80] or "presentation"


def _content_disposition(filename: str) -> str:
    """RFC 5987：头必须 ASCII 安全——filename= 给 ASCII 回退，中文原名走 filename*。"""
    fallback = "".join(
        ch if 32 <= ord(ch) < 127 and ch not in '\\"' else "_" for ch in filename
    ) or "file"
    return (
        f'attachment; filename="{fallback}"; '
        f"filename*=UTF-8''{quote(filename)}"
    )


def _current_render_dir(doc_id: str, meta: Dict[str, Any]) -> Path:
    return storage.render_dir(doc_id, int(meta.get("version", 0)))


def _regenerate_render(doc_id: str, meta: Dict[str, Any]) -> Path:
    """Re-render the current version once; returns its render dir.

    Needed for render dirs produced before v0.3 (they kept no PDF), emptied by
    hand, or degraded to the previous version's pages by a render timeout (E21).
    The render goes to a scratch dir beside the real one: the committed dir
    keeps whatever previews the UI is showing until a complete render exists to
    take their place, so a second timeout can't leave the document without a
    single page. A completed render also clears the degraded state."""
    version = int(meta.get("version", 0))
    render_dir = storage.render_dir(doc_id, version)
    scratch = render_dir.with_name(f"{render_dir.name}.regen-{uuid.uuid4().hex}")
    try:
        render.render_pptx(storage.current_path(doc_id), scratch, scale=RENDER_SCALE)
        with storage.lock_for(doc_id):
            if int(storage.load_meta(doc_id).get("version", 0)) != version:
                raise edits.PatchConflict("文档在渲染期间已更新，请重试")
            render_dir.mkdir(parents=True, exist_ok=True)
            for item in scratch.iterdir():
                if item.is_file():
                    os.replace(item, render_dir / item.name)
            render.partial_marker(render_dir).unlink(missing_ok=True)
            objects_payload = storage.load_objects(doc_id)
            if objects_payload.pop("render_partial", None):
                storage.save_objects(doc_id, objects_payload)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    return render_dir


def _is_degraded(render_dir: Path) -> bool:
    """E21: the pages in this render dir are the previous version's."""
    return render.partial_marker(render_dir).exists()


def _hidden_slide_indexes(doc_id: str, meta: Dict[str, Any]) -> set:
    """当前文稿的隐藏页 0-based 索引集合（presentation.xml 的 show="0"）。

    v0.5 #5/#10：导出默认跳过隐藏页的依据；读不出来时按无隐藏处理，导出
    行为退回 v0.4（全量），不让隐藏特性把导出搞挂。"""
    from pptx import Presentation as _Presentation

    try:
        prs = _Presentation(str(storage.current_path(doc_id)))
        flags = [child.get("show") == "0" for child in prs.slides._sldIdLst]
    except Exception:
        return set()
    return {index for index, flag in enumerate(flags) if flag}


@app.get("/api/docs/{doc_id}/export/pdf")
def export_pdf(
    doc_id: str,
    include_hidden: bool = False,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    render_dir = _current_render_dir(doc_id, meta)
    pdf_path = render_dir / "export.pdf"
    if not pdf_path.exists() or _is_degraded(render_dir):
        render_dir = _regenerate_render(doc_id, meta)
        pdf_path = render_dir / "export.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError("PDF 尚未生成（渲染未产出 PDF），请稍后重试或先重新渲染")
    if not include_hidden:
        hidden = _hidden_slide_indexes(doc_id, meta)
        if hidden:
            import pymupdf

            doc = pymupdf.open(pdf_path)
            try:
                # 页数对齐才按索引剔除；渲染器若已跳过隐藏页则原样导出
                if doc.page_count == int(meta.get("slide_count", 0)):
                    doc.select([i for i in range(doc.page_count) if i not in hidden])
                    tmp_dir = storage.DATA_DIR / "tmp"
                    tmp_dir.mkdir(parents=True, exist_ok=True)
                    tmp_path = tmp_dir / f"export-{uuid.uuid4().hex}.pdf"
                    doc.save(str(tmp_path), garbage=3, deflate=True)
                    return FileResponse(
                        tmp_path,
                        media_type="application/pdf",
                        filename=f"{_export_base_name(meta)}.pdf",
                        headers={"Cache-Control": "no-store"},
                        background=BackgroundTask(lambda: tmp_path.unlink(missing_ok=True)),
                    )
            finally:
                doc.close()
    return FileResponse(
        pdf_path,
        media_type="application/pdf",
        filename=f"{_export_base_name(meta)}.pdf",
        headers={"Cache-Control": "no-store"},
    )


@app.get("/api/docs/{doc_id}/export/images.zip")
def export_images_zip(
    doc_id: str,
    include_hidden: bool = False,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    render_dir = _current_render_dir(doc_id, meta)
    pages = sorted(render_dir.glob("slide-*.png"))
    if not pages or _is_degraded(render_dir):
        # pages parked there by a render timeout belong to the previous version
        # (E21): re-render rather than ship them as this version's images
        render_dir = _regenerate_render(doc_id, meta)
        pages = sorted(render_dir.glob("slide-*.png"))
    if not pages:
        raise FileNotFoundError("暂无可打包的页面图片，请先触发一次渲染")
    if not include_hidden and len(pages) == int(meta.get("slide_count", 0)):
        # 页数对齐才按索引剔除（渲染 PNG 与页码 1:1）；渲染器已跳过隐藏页时全量
        hidden = _hidden_slide_indexes(doc_id, meta)
        if hidden:
            pages = [page for index, page in enumerate(pages) if index not in hidden]
    tmp_dir = storage.DATA_DIR / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = tmp_dir / f"export-{uuid.uuid4().hex}.zip"
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_STORED) as zf:
            # PNGs are already deflate-compressed; storing keeps export fast
            for page in pages:
                zf.write(page, arcname=page.name)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise
    return FileResponse(
        tmp_path,
        media_type="application/zip",
        filename=f"{_export_base_name(meta)}-images.zip",
        headers={"Cache-Control": "no-store"},
        background=BackgroundTask(lambda: tmp_path.unlink(missing_ok=True)),
    )


# ---------- v0.4 B：备注版 PDF / TXT 大纲 / 图片导入 / 版本对比 ----------

NOTES_PDF_CJK_FONT = "china-s"  # PyMuPDF 内置简体字体，备注文字可渲染


def _compose_notes_pdf(pptx_path: Path, slides_pdf_path: Path, out_path: Path, hidden: Optional[set] = None) -> Path:
    """幻灯片 + 演讲者备注合成讲义 PDF：A4 竖版，上页图下备注。

    hidden：要跳过的 0-based 页索引（v0.5 #10 隐藏页默认不进讲义）。"""
    import pymupdf
    from pptx import Presentation as _Presentation

    hidden = hidden or set()
    prs = _Presentation(str(pptx_path))
    notes = []
    for slide in prs.slides:
        text = ""
        if getattr(slide, "has_notes_slide", False):
            try:
                text = str(slide.notes_slide.notes_text_frame.text or "").strip()
            except Exception:
                text = ""
        notes.append(text)
    src = pymupdf.open(str(slides_pdf_path))
    out = pymupdf.open()
    page_w, page_h, margin = 595, 842, 40
    for index in range(max(src.page_count, len(notes))):
        if index in hidden:
            continue
        page = out.new_page(width=page_w, height=page_h)
        top = margin
        if index < src.page_count:
            rect = src[index].rect
            scale = min((page_w - margin * 2) / rect.width, (page_h * 0.62) / rect.height)
            w, h = rect.width * scale, rect.height * scale
            x0 = (page_w - w) / 2
            page.show_pdf_page(pymupdf.Rect(x0, top, x0 + w, top + h), src, index)
            top += h + 16
        note = notes[index].strip() if index < len(notes) else ""
        label = f"— 第 {index + 1} 页 · 演讲者备注 —"
        body = f"{label}\n{note or '（无备注）'}"
        box = pymupdf.Rect(margin, top + 4, page_w - margin, page_h - margin)
        page.insert_textbox(box, body, fontname=NOTES_PDF_CJK_FONT, fontsize=10)
    out.save(str(out_path), garbage=3, deflate=True)
    return out_path


@app.get("/api/docs/{doc_id}/export/pdf-notes")
def export_pdf_notes(
    doc_id: str,
    include_hidden: bool = False,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    render_dir = _current_render_dir(doc_id, meta)
    pdf_path = render_dir / "export.pdf"
    if not pdf_path.exists() or _is_degraded(render_dir):
        render_dir = _regenerate_render(doc_id, meta)
        pdf_path = render_dir / "export.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError("PDF 尚未生成（渲染未产出 PDF），请稍后重试或先重新渲染")
    tmp_dir = storage.DATA_DIR / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = tmp_dir / f"pdf-notes-{uuid.uuid4().hex}.pdf"
    hidden = set() if include_hidden else _hidden_slide_indexes(doc_id, meta)
    try:
        _compose_notes_pdf(storage.current_path(doc_id), pdf_path, tmp_path, hidden=hidden)
    except BaseException:
        tmp_path.unlink(missing_ok=True)
        raise
    return FileResponse(
        tmp_path,
        media_type="application/pdf",
        filename=f"{_export_base_name(meta)}-notes.pdf",
        headers={"Cache-Control": "no-store"},
        background=BackgroundTask(lambda: tmp_path.unlink(missing_ok=True)),
    )


def _outline_lines(meta: Dict[str, Any], pptx_path: Path) -> List[str]:
    """全稿文字大纲：页标题/正文/表格行/备注，贴飞书或文档用。"""
    from pptx import Presentation as _Presentation

    prs = _Presentation(str(pptx_path))
    title = str(
        meta.get("title")
        or Path(str(meta.get("file_name") or "presentation.pptx")).stem
    )
    lines = [f"# {title}", ""]
    hidden_flags = [child.get("show") == "0" for child in prs.slides._sldIdLst]

    def walk(shapes):
        for shape in shapes:
            if getattr(shape, "shape_type", None) == 6:
                try:
                    walk(shape.shapes)
                except Exception:
                    pass
                continue
            if getattr(shape, "has_table", False):
                try:
                    row_texts = []
                    for row in shape.table.rows:
                        cells = [str(cell.text or "").strip() for cell in row.cells]
                        if any(cells):
                            row_texts.append(" | ".join(cells))
                    if row_texts:
                        lines.append("")
                        lines.extend(f"  {row}" for row in row_texts)
                        lines.append("")
                except Exception:
                    pass
                continue
            if not getattr(shape, "has_text_frame", False):
                continue
            try:
                text = str(shape.text or "").strip()
            except Exception:
                continue
            if not text:
                continue
            for line in text.split("\n"):
                if line.strip():
                    lines.append(f"- {line.strip()}")

    for index, slide in enumerate(prs.slides, start=1):
        marker = "（已隐藏）" if hidden_flags[index - 1] else ""
        lines.append(f"## 第 {index} 页{marker}")
        walk(slide.shapes)
        if getattr(slide, "has_notes_slide", False):
            try:
                note = str(slide.notes_slide.notes_text_frame.text or "").strip()
            except Exception:
                note = ""
            if note:
                lines.append(f"  > 备注：{note}")
        lines.append("")
    return lines


@app.get("/api/docs/{doc_id}/export/outline")
def export_outline(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    content = "\n".join(_outline_lines(meta, storage.current_path(doc_id)))
    return Response(
        content,
        media_type="text/plain; charset=utf-8",
        headers={
            "Content-Disposition": _content_disposition(
                f"{_export_base_name(meta)}-outline.txt"
            ),
            "Cache-Control": "no-store",
        },
    )


_IMPORT_IMAGE_MAGIC = (b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff")
IMPORT_IMAGES_MAX = 60
IMPORT_IMAGE_MAX_BYTES = 30 * 1024 * 1024


@app.post("/api/docs/import-images")
def import_images(files: Optional[List[UploadFile]] = File(None), ocr: str = Form("1")):
    """多张图片（扫描件/截图）每张一页生成 PPTX，图片满宽居中。

    v0.6 #4：默认开启本机 OCR（macOS Vision via pyobjc）——识别出的文字行
    生成原生文本框叠在图片上层（可直接改字），图片保留作背景层；`ocr=0` 关闭；
    OCR 不可用/失败/无文字时行为与 v0.4 一致（纯图片页）。"""
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传 1 张图片（png/jpg）")
    if len(files) > IMPORT_IMAGES_MAX:
        raise HTTPException(
            status_code=400, detail=f"单次最多 {IMPORT_IMAGES_MAX} 张图片，请分批导入"
        )
    images: List[bytes] = []
    first_stem = ""
    total = 0
    for upload in files:
        original = Path(upload.filename or "image.png").name
        stem = re.sub(r"[\x00-\x1f]", "", original).strip()
        if not first_stem:
            first_stem = Path(stem or "image").stem
        data = upload.file.read(IMPORT_IMAGE_MAX_BYTES + 1)
        if len(data) > IMPORT_IMAGE_MAX_BYTES:
            raise HTTPException(status_code=413, detail=f"{original} 超过 30 MB 限制")
        if not data or not any(data.startswith(magic) for magic in _IMPORT_IMAGE_MAGIC):
            raise HTTPException(
                status_code=400, detail=f"{original} 不是有效的 png/jpg 图片"
            )
        total += len(data)
        if total > MAX_UPLOAD_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"图片总量超过 {MAX_UPLOAD_BYTES // (1024 * 1024)} MB 限制",
            )
        images.append(data)
    import io

    from pptx import Presentation as _Presentation
    from pptx.util import Emu as _Emu

    from services import ocr as ocr_mod

    ocr_requested = str(ocr or "").strip().lower() not in ("0", "false", "no", "off")
    ocr_ok = True
    ocr_active = ocr_requested and ocr_mod.ocr_available()
    if ocr_requested and not ocr_active:
        log.info("import-images: OCR requested but unavailable — falling back to image-only pages")
    ocr_lines_total = 0
    ocr_pages_with_text = 0

    prs = _Presentation()
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    slide_w, slide_h = int(prs.slide_width), int(prs.slide_height)
    for data in images:
        slide = prs.slides.add_slide(blank)
        picture = slide.shapes.add_picture(io.BytesIO(data), _Emu(0), _Emu(0), width=_Emu(slide_w))
        picture.top = _Emu(max(0, (slide_h - int(picture.height)) // 2))
        if ocr_active:
            try:
                lines = ocr_mod.ocr_image_bytes(data)
            except Exception as exc:
                log.warning("import-images: OCR failed for one page: %s", exc)
                lines = []
                ocr_ok = False
            if lines:
                ocr_pages_with_text += 1
                ocr_lines_total += ocr_mod.overlay_textboxes(
                    slide,
                    lines,
                    int(picture.left),
                    int(picture.top),
                    int(picture.width),
                    int(picture.height),
                )
    file_name = f"{(first_stem or '图片集')[:60]}-imported.pptx"
    storage.ensure_dirs()
    tmp_dir = storage.DATA_DIR / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_path = tmp_dir / f"{uuid.uuid4().hex}.pptx"
    prs.save(str(tmp_path))
    doc_id: Optional[str] = None
    try:
        try:
            _, parsed = parse.parse_pptx(tmp_path)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"生成的 PPTX 无法解析：{exc}") from exc
        meta = storage.create_doc(
            file_name=file_name,
            content=tmp_path.read_bytes(),
            slide_count=int(parsed["slide_count"]),
            slide_size=parsed["slide_size_emu"],
        )
        doc_id = meta["doc_id"]
        storage.save_objects(doc_id, parsed)
        _render_with_cache(doc_id, storage.current_path(doc_id), storage.render_dir(doc_id, 0))
        _prewarm_doc_async(doc_id)
        result = _serialize(meta, parsed)
        result["ocr"] = {
            "enabled": bool(ocr_active and ocr_ok),
            "pages_with_text": ocr_pages_with_text,
            "lines": ocr_lines_total,
        }
        return result
    except Exception:
        if doc_id:
            shutil.rmtree(storage.doc_dir(doc_id), ignore_errors=True)
            storage.release_lock(doc_id)
        raise
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass


def _version_number(raw: str) -> Optional[int]:
    text = str(raw or "").strip().lower()
    if re.fullmatch(r"v\d+", text):
        return int(text[1:])
    if re.fullmatch(r"\d+", text):
        return int(text)
    return None


def _page_text_lists(payload: Dict[str, Any]) -> List[List[str]]:
    count = int(payload.get("slide_count") or 0)
    pages: List[List[str]] = [[] for _ in range(count)]
    for obj in payload.get("objects", []):
        try:
            index = int(obj.get("slide_index"))
        except (TypeError, ValueError):
            continue
        if 0 <= index < count:
            text = str(obj.get("text") or "").strip()
            if text:
                pages[index].append(text)
    return pages


@app.get("/api/docs/{doc_id}/versions/{version_a}/compare/{version_b}")
def compare_versions(doc_id: str, version_a: str, version_b: str):
    """两版本文字差异摘要：新增/删除/改动的页码与内容要点（按页位对齐）。"""
    meta = storage.load_meta(doc_id)
    a_num, b_num = _version_number(version_a), _version_number(version_b)
    if a_num is None or b_num is None:
        raise HTTPException(status_code=400, detail="版本号格式无效（应为 v0 / 0）")
    a_id, b_id = f"v{a_num}", f"v{b_num}"
    path_a = storage.version_file(doc_id, a_id)  # 缺失 → 404
    path_b = storage.version_file(doc_id, b_id)
    _objs_a, payload_a = parse.parse_pptx(path_a)
    _objs_b, payload_b = parse.parse_pptx(path_b)
    pages_a = _page_text_lists(payload_a)
    pages_b = _page_text_lists(payload_b)

    def _titles(payload):
        by_slide: Dict[int, list] = {}
        for obj in payload.get("objects", []):
            try:
                by_slide.setdefault(int(obj.get("slide_index")), []).append(obj)
            except (TypeError, ValueError):
                continue
        return [
            agent._slide_title(by_slide.get(index, []))
            for index in range(int(payload.get("slide_count") or 0))
        ]

    titles_a, titles_b = _titles(payload_a), _titles(payload_b)
    shared = min(len(pages_a), len(pages_b))
    changed, added, removed = [], [], []
    unchanged = 0
    for index in range(shared):
        set_a = {edits.normalize_text(t) for t in pages_a[index]}
        set_b = {edits.normalize_text(t) for t in pages_b[index]}
        if set_a == set_b:
            unchanged += 1
            continue
        additions = [t[:80] for t in pages_b[index] if edits.normalize_text(t) not in set_a][:6]
        deletions = [t[:80] for t in pages_a[index] if edits.normalize_text(t) not in set_b][:6]
        changed.append(
            {
                "page": index + 1,
                "title": titles_b[index] if index < len(titles_b) else "",
                "additions": additions,
                "deletions": deletions,
            }
        )
    for index in range(shared, len(pages_b)):
        added.append(
            {
                "page": index + 1,
                "title": titles_b[index] if index < len(titles_b) else "",
                "points": [t[:80] for t in pages_b[index][:6]],
            }
        )
    for index in range(shared, len(pages_a)):
        removed.append(
            {
                "page": index + 1,
                "title": titles_a[index] if index < len(titles_a) else "",
                "points": [t[:80] for t in pages_a[index][:6]],
            }
        )
    diff_pages = len(changed) + len(added) + len(removed)
    struct = diff.structural_changelog(payload_a, payload_b)
    # changelog 自带版本标识，AI 单看 struct 也能知道比的是哪两版
    struct["a"] = a_id
    struct["b"] = b_id
    return {
        "a": a_id,
        "b": b_id,
        "pages_a": len(pages_a),
        "pages_b": len(pages_b),
        "changed": changed,
        "added": added,
        "removed": removed,
        "unchanged_pages": unchanged,
        # v0.5 #8：对象级结构化 changelog（新增/删除/移动/几何/样式），JSON 给 AI 消费
        "struct": struct,
        "summary": (
            f"共 {diff_pages} 页存在差异：改动 {len(changed)} 页、"
            f"新增 {len(added)} 页、删除 {len(removed)} 页；{struct['summary']['text']}"
        ),
    }


# ---------- v0.4 C：密码保护 / 只读分享 / 宏只读 ----------

def _password_digest(salt: str, password: str) -> str:
    return hashlib.sha256(f"{salt}:{password}".encode("utf-8")).hexdigest()


def _assert_doc_open(meta: Dict[str, Any], x_doc_password: Optional[str]) -> None:
    """密码保护：未提供正确密码时拒绝访问（X-Doc-Password 头）。"""
    if not meta.get("password_hash"):
        return
    supplied = str(x_doc_password or "")
    salt = str(meta.get("password_salt") or "")
    expected = _password_digest(salt, supplied)
    if not hmac.compare_digest(expected, str(meta.get("password_hash"))):
        raise HTTPException(
            status_code=401,
            detail="文档已设置密码保护：请先用正确的密码解锁（X-Doc-Password 请求头）",
        )


def _assert_editable(meta: Dict[str, Any]) -> None:
    """宏文档只读：.pptx 内嵌 vbaProject 的文件拒绝一切编辑入口。"""
    if meta.get("has_macro"):
        raise HTTPException(
            status_code=400,
            detail="含宏的文件已转为只读预览（检测到 vbaProject，防宏病毒），不能编辑",
        )


def _serialize_share(meta: Dict[str, Any], entry: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "token": entry.get("token"),
        "created_at": entry.get("created_at"),
        "revoked": bool(entry.get("revoked")),
        "share_url": f"/share/{entry.get('token')}",
    }


def _find_share(token: str):
    """token → (doc_id, meta, entry)；已撤销/不存在返回 None。"""
    token = str(token or "").strip()
    if not token or not re.fullmatch(r"[A-Za-z0-9_-]{8,64}", token):
        return None
    storage.ensure_dirs()
    for path in storage.DOCS_DIR.iterdir():
        if not path.is_dir() or not re.fullmatch(r"[a-f0-9]{32}", path.name):
            continue
        try:
            meta = json.loads((path / "meta.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        for entry in meta.get("shares") or []:
            if isinstance(entry, dict) and entry.get("token") == token and not entry.get("revoked"):
                return path.name, meta, entry
    return None


@app.post("/api/docs/{doc_id}/password")
def set_document_password(doc_id: str, body: PasswordIn, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        _assert_doc_open(meta, x_doc_password)
        if body.password:
            salt = uuid.uuid4().hex
            meta["password_salt"] = salt
            meta["password_hash"] = _password_digest(salt, body.password)
            meta["password_set_at"] = storage.now_iso()
        else:
            meta.pop("password_hash", None)
            meta.pop("password_salt", None)
            meta.pop("password_set_at", None)
        meta["updated_at"] = storage.now_iso()
        storage.save_meta(doc_id, meta)
    return {"ok": True, "locked": bool(body.password)}


@app.post("/api/docs/{doc_id}/unlock")
def unlock_document(doc_id: str, body: UnlockIn):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, body.password)
    return {"ok": True}


@app.post("/api/docs/{doc_id}/share")
def create_share(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        _assert_doc_open(meta, x_doc_password)
        token = uuid.uuid4().hex
        entry = {"token": token, "created_at": storage.now_iso(), "revoked": False}
        meta.setdefault("shares", []).append(entry)
        meta["updated_at"] = storage.now_iso()
        storage.save_meta(doc_id, meta)
    return _serialize_share(meta, entry)


@app.get("/api/docs/{doc_id}/shares")
def list_shares(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    return {"shares": [_serialize_share(meta, e) for e in meta.get("shares") or []]}


@app.delete("/api/docs/{doc_id}/share/{token}")
def revoke_share(doc_id: str, token: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        _assert_doc_open(meta, x_doc_password)
        found = False
        for entry in meta.get("shares") or []:
            if entry.get("token") == token:
                entry["revoked"] = True
                found = True
        if not found:
            raise FileNotFoundError("share token not found")
        storage.save_meta(doc_id, meta)
    return {"ok": True, "revoked": token}


_SHARE_PAGE = """<!doctype html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} · 只读分享</title>
<style>
body{{margin:0;background:#f1f5f9;font-family:-apple-system,"PingFang SC",sans-serif}}
header{{padding:14px 20px;background:#0f172a;color:#fff;display:flex;justify-content:space-between;align-items:center}}
header small{{opacity:.7}}
main{{max-width:980px;margin:0 auto;padding:20px}}
img{{width:100%;display:block;margin:0 0 14px;border-radius:8px;box-shadow:0 4px 14px rgba(0,0,0,.12);background:#fff}}
a.dl{{color:#38bdf8}}
</style></head><body>
<header><strong>{title}</strong><small>只读分享 · <a class="dl" href="/share/{token}/download">下载 PPTX</a></small></header>
<main id="pages"><p style="color:#64748b">加载中…</p></main>
<script>
fetch('/share/{token}/data').then(function(r){{return r.json()}}).then(function(d){{
  document.title = d.title + ' · 只读分享';
  var main = document.getElementById('pages');
  main.innerHTML = '';
  d.pages.forEach(function(p){{
    var img = new Image();
    img.src = p.thumb;
    img.alt = '第 ' + p.page + ' 页';
    img.loading = 'lazy';
    main.appendChild(img);
  }});
}}).catch(function(){{ document.getElementById('pages').innerHTML = '<p>分享内容不可用或已撤销</p>'; }});
</script></body></html>"""


@app.get("/share/{token}")
def share_page(token: str):
    found = _find_share(token)
    if found is None:
        raise FileNotFoundError("share not found")
    _doc_id, meta, _entry = found
    title = html.escape(str(meta.get("title") or "演示文稿"))
    safe_token = html.escape(token)
    return Response(
        _SHARE_PAGE.format(title=title, token=safe_token),
        media_type="text/html; charset=utf-8",
        headers={"Cache-Control": "no-store"},
    )


@app.get("/share/{token}/data")
def share_data(token: str):
    found = _find_share(token)
    if found is None:
        raise FileNotFoundError("share not found")
    doc_id, meta, _entry = found
    slide_count = int(meta.get("slide_count") or 0)
    version = int(meta.get("version") or 0)
    pages = [
        {
            "page": index + 1,
            "image": f"/share/{token}/slides/{index}/image",
            "thumb": f"/share/{token}/slides/{index}/image?thumb=1",
        }
        for index in range(slide_count)
    ]
    return {
        "title": meta.get("title") or "演示文稿",
        "slide_count": slide_count,
        "version_id": f"v{version}",
        "pages": pages,
        "download_url": f"/share/{token}/download",
    }


@app.get("/share/{token}/slides/{slide_index}/image")
def share_slide_image(token: str, slide_index: int, thumb: bool = False):
    found = _find_share(token)
    if found is None:
        raise FileNotFoundError("share not found")
    doc_id, meta, _entry = found
    if slide_index < 0 or slide_index >= int(meta.get("slide_count", 0)):
        raise FileNotFoundError("slide not found")
    version = int(meta.get("version", 0))
    path = (
        storage.thumb_path(doc_id, version, slide_index)
        if thumb
        else storage.image_path(doc_id, version, slide_index)
    )
    if not path.exists():
        raise FileNotFoundError("image not rendered")
    return FileResponse(
        path,
        media_type="image/png",
        headers={"Cache-Control": "no-store"},
    )


@app.get("/share/{token}/download")
def share_download(token: str):
    found = _find_share(token)
    if found is None:
        raise FileNotFoundError("share not found")
    doc_id, meta, _entry = found
    path = storage.current_path(doc_id)
    if not path.exists():
        raise FileNotFoundError("document file not found")
    name = meta.get("file_name") or "presentation.pptx"
    if not name.lower().endswith(".pptx"):
        name = f"{name}.pptx"
    return FileResponse(
        path, media_type=PPTX_MEDIA_TYPE, filename=name,
        headers={"Cache-Control": "no-store"},
    )


@app.get("/api/audit")
def get_audit(limit: int = 50, doc_id: Optional[str] = None):
    entries = storage.recent_audit(limit=limit, doc_id=doc_id)
    return {"entries": entries, "count": len(entries)}


@app.post("/api/docs/{doc_id}/duplicate")
def duplicate_document(doc_id: str):
    """另存副本：把当前版本复制成一份全新的会话（不带聊天记录/版本历史）。"""
    meta, objects_payload = _load_doc(doc_id)
    src = storage.current_path(doc_id)
    if not src.exists():
        raise FileNotFoundError("document file not found")
    # the copy is rendered from scratch, so the source's "preview is the
    # previous version" flag (E21) must not travel with its objects
    objects_payload = {k: v for k, v in objects_payload.items() if k != "render_partial"}
    new_meta: Optional[Dict[str, Any]] = None
    try:
        new_meta = storage.create_doc(
            file_name=meta.get("file_name") or "presentation.pptx",
            content=src.read_bytes(),
            slide_count=int(
                meta.get("slide_count") or objects_payload.get("slide_count") or 0
            ),
            slide_size=objects_payload.get("slide_size_emu")
            or meta.get("slide_size_emu")
            or {},
        )
        storage.save_objects(new_meta["doc_id"], objects_payload)
        title = str(meta.get("title") or Path(str(meta.get("file_name"))).stem)
        storage.touch_doc(new_meta["doc_id"], title=f"{title} 副本")
        _render_with_cache(
            new_meta["doc_id"],
            storage.current_path(new_meta["doc_id"]),
            storage.render_dir(new_meta["doc_id"], 0),
        )
        # 起一条系统备注说明副本来源，聊天上下文可追溯
        storage.append_messages(
            new_meta["doc_id"],
            [
                {
                    "id": uuid.uuid4().hex[:12],
                    "role": "system",
                    "content": f"已从「{title}」另存为副本，当前是独立会话（从 v0 起步）",
                    "at": storage.now_iso(),
                    "kind": "duplicate",
                }
            ],
        )
        return _serialize(storage.load_meta(new_meta["doc_id"]), objects_payload)
    except Exception as exc:
        if new_meta:
            shutil.rmtree(storage.doc_dir(new_meta["doc_id"]), ignore_errors=True)
            storage.release_lock(new_meta["doc_id"])
        if isinstance(exc, (render.RenderError, render.RenderBusy)):
            raise
        raise HTTPException(status_code=500, detail=_sanitize_detail(exc) or "复制失败") from exc


# ---------- v0.5 #3：merge 从另一 deck 拉页 ----------

@app.post("/api/docs/{doc_id}/merge")
def merge_document(
    doc_id: str,
    file: UploadFile = File(...),
    pages: str = Form(""),
    position: Optional[int] = Form(None),
    base_version: Optional[int] = Form(None),
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    """上传另一 PPTX，把指定页拉进当前文档（rels 全量重新家园化），生成新版本。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    _assert_editable(meta)
    original = Path(file.filename or "merge-source.pptx").name
    if not original.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="合并源只支持 .pptx 文件")
    data = file.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail=f"合并源超过 {MAX_UPLOAD_BYTES // (1024 * 1024)} MB 限制")
    if len(data) < 100:
        raise HTTPException(status_code=400, detail="合并源文件过小或已损坏")
    tmp_dir = storage.DATA_DIR / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    src_tmp = tmp_dir / f"merge-src-{uuid.uuid4().hex}.pptx"
    work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
    render_dir: Optional[Path] = None
    try:
        src_tmp.write_bytes(data)
        # ---- phase 1 (short lock): 校验 + 分配版本号 ----
        with storage.lock_for(doc_id):
            meta = storage.load_meta(doc_id)
            if base_version is not None and int(base_version) != int(meta.get("version", 0)):
                raise edits.PatchConflict("文档已经更新，请刷新后再合并")
            entry_version = int(meta.get("version", 0))
            new_version, _ = storage.next_version_id(meta)
        # ---- phase 2 (no lock): 跨包克隆 + 渲染 + 解析 ----
        merged, touched = edits.merge_slides_from_file(
            storage.current_path(doc_id), src_tmp, work_path, pages, position
        )
        render_dir, objects_payload, page_count = _prepare_commit(doc_id, work_path, new_version)
        # ---- phase 3 (short lock): 提交 ----
        try:
            with storage.lock_for(doc_id):
                committed, objects_payload = _commit_prepared_locked(
                    doc_id, new_version, f"合并 {merged} 页", work_path, render_dir,
                    objects_payload, page_count, expected_version=entry_version,
                )
                result = _commit_result(committed, objects_payload, sorted(touched))
                result["merged_slides"] = merged
                message = {
                    "id": uuid.uuid4().hex[:12],
                    "role": "system",
                    "content": f"从「{original}」合并 {merged} 页 · 保存为 {result['version_id']}",
                    "at": storage.now_iso(),
                    "kind": "merge",
                    "version": result["version_id"],
                }
                result["messages"] = storage.append_messages(doc_id, [message])
                result["message_count"] = len(result["messages"])
        except edits.PatchConflict:
            _discard_prepared(work_path, render_dir)
            raise
    except edits.PatchError as exc:
        _discard_prepared(work_path, render_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception:
        _discard_prepared(work_path, render_dir)
        raise
    finally:
        try:
            src_tmp.unlink()
        except OSError:
            pass
    storage.append_audit(
        action="merge", doc_id=doc_id, summary=f"合并 {merged} 页（{original}）",
        op_count=merged, ok=True, version=result.get("version_id"),
    )
    log.info("merged %d slides into %s -> v%d", merged, doc_id, new_version)
    return result


# ---------- v0.5 #9：watch 实时预览 ----------

# 空转轮询间隔：订阅队列短暂等待，空了就发一条 SSE 注释行保活
WATCH_EVENT_TIMEOUT = 1.0


@app.get("/api/docs/{doc_id}/events")
def doc_events(
    doc_id: str,
    max_events: Optional[int] = None,
    x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password"),
):
    """文档事件流（SSE）：连接即 hello（当前版本），此后每次版本变更推 changed。

    max_events：发出指定条数（含 hello）后发 bye 并收尾 —— 无限流无法在
    TestClient 下断开（不发 disconnect），测试用它让流自然终止；浏览器不传。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    version = int(meta.get("version", 0))
    if max_events is not None:
        max_events = max(1, min(int(max_events), 1000))
    q = broker.subscribe(doc_id)

    def stream():
        sent = 0
        try:
            yield _sse(
                "hello",
                {
                    "doc_id": doc_id,
                    "version": version,
                    "version_id": f"v{version}",
                    "slide_count": int(meta.get("slide_count", 0)),
                },
            )
            sent += 1
            while max_events is None or sent < max_events:
                try:
                    event = q.get(timeout=WATCH_EVENT_TIMEOUT)
                except queue.Empty:
                    yield ": keepalive\n\n"
                    continue
                yield _sse(str(event.get("type") or "changed"), event)
                sent += 1
            if max_events is not None:
                yield _sse("bye", {"doc_id": doc_id, "sent": sent})
        finally:
            broker.unsubscribe(doc_id, q)

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


_WATCH_PAGE = """<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>watch · 实时预览</title>
<style>
  :root { color-scheme: dark; }
  * { box-sizing: border-box; }
  body { margin: 0; font: 14px/1.5 -apple-system, "PingFang SC", sans-serif;
         background: #14161a; color: #e6e9ef; display: flex; flex-direction: column; height: 100vh; }
  header { display: flex; align-items: center; gap: 12px; padding: 10px 16px;
           background: #1b1f26; border-bottom: 1px solid #2a2f3a; }
  header h1 { font-size: 15px; margin: 0; font-weight: 600; }
  .badge { font-size: 12px; color: #8b94a7; }
  .dot { width: 9px; height: 9px; border-radius: 50%; background: #f0b429; }
  .dot.on { background: #35c46f; }
  aside { width: 190px; overflow-y: auto; background: #181b21; border-right: 1px solid #2a2f3a; }
  aside .thumb { padding: 8px; cursor: pointer; }
  aside .thumb img { width: 100%; border-radius: 4px; display: block;
                     border: 2px solid transparent; }
  aside .thumb.cur img { border-color: #4d8ef7; }
  aside .thumb span { font-size: 11px; color: #8b94a7; }
  main { flex: 1; overflow: auto; display: flex; align-items: flex-start; justify-content: center; }
  main img.page { max-width: 100%; max-height: 100%; margin: 16px; border-radius: 6px;
                  box-shadow: 0 8px 40px rgba(0,0,0,.5); background: #fff; }
  footer { padding: 6px 16px; font-size: 12px; color: #8b94a7; background: #1b1f26;
           border-top: 1px solid #2a2f3a; }
</style>
</head>
<body>
<header>
  <span class="dot" id="dot"></span>
  <h1 id="title">实时预览</h1>
  <span class="badge" id="version-badge">v-</span>
  <span class="badge" id="slide-badge"></span>
</header>
<div style="flex:1; display:flex; min-height:0;">
  <aside id="thumbs"></aside>
  <main><img class="page" id="page" alt="当前页渲染图"></main>
</div>
<footer id="log">等待连接…</footer>
<script>
const DOC_ID = "__DOC_ID__";
const base = "/api/docs/" + DOC_ID;
let current = 0, version = null, slideCount = 0;

const es = new EventSource(base + "/events");
es.addEventListener("open", () => { dot.classList.add("on"); log.textContent = "已连接"; });
es.addEventListener("error", () => { dot.classList.remove("on"); log.textContent = "连接断开，自动重连中…"; });
es.addEventListener("hello", (e) => {
  const d = JSON.parse(e.data);
  version = d.version_id;
  slideCount = d.slide_count || 0;
  log.textContent = "已连接 · 当前 " + version;
  refresh(false);
});
es.addEventListener("changed", (e) => {
  const d = JSON.parse(e.data);
  version = d.version_id || version;
  if (d.slide_count != null) slideCount = d.slide_count;
  if (Array.isArray(d.touched_slides) && d.touched_slides.length && !d.touched_slides.includes(current)) {
    current = d.touched_slides[0];
  }
  log.textContent = "变更 → " + version + (d.touched_slides ? " · 页 " + d.touched_slides.map(n => n + 1).join(",") : "");
  refresh(true);
});

function refresh(scroll) {
  versionBadge.textContent = version || "-";
  slideBadge.textContent = slideCount ? slideCount + " 页" : "";
  if (current >= slideCount) current = 0;
  page.src = base + "/slides/" + current + "/image?v=" + (version || "0");
  thumbs.innerHTML = "";
  for (let i = 0; i < slideCount; i++) {
    const div = document.createElement("div");
    div.className = "thumb" + (i === current ? " cur" : "");
    const img = document.createElement("img");
    img.src = base + "/slides/" + i + "/image?thumb=1&v=" + (version || "0");
    img.loading = "lazy";
    const span = document.createElement("span");
    span.textContent = "第 " + (i + 1) + " 页";
    div.appendChild(img); div.appendChild(span);
    div.onclick = () => { current = i; refresh(true); };
    thumbs.appendChild(div);
  }
  if (scroll && slideCount) {
    const cur = thumbs.children[current];
    if (cur) cur.scrollIntoView({ block: "nearest" });
  }
}

fetch(base).then(r => r.json()).then(d => {
  document.title = (d.title || "watch") + " · 实时预览";
  title.textContent = d.title || "实时预览";
  if (!version) { version = d.version_id; slideCount = d.slide_count || 0; refresh(false); }
}).catch(() => {});
</script>
</body>
</html>
"""


@app.get("/api/docs/{doc_id}/watch")
def watch_page(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    """浏览器分屏实时预览页：EventSource 订阅 /events，changed 后自动刷新渲染图。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    # 标题由页面 JS 拉取 /api/docs/{id} 后自行填充；模板只需注入 doc_id
    page = _WATCH_PAGE.replace("__DOC_ID__", doc_id)
    return HTMLResponse(page, headers={"Cache-Control": "no-store"})


# ---------- v0.6 #5：演讲者模式 + 排练计时 ----------

_PRESENT_PAGE = """<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>present · 演讲者模式</title>
<style>
  :root { color-scheme: dark; }
  * { box-sizing: border-box; }
  body { margin: 0; font: 14px/1.5 -apple-system, "PingFang SC", sans-serif;
         background: #0b0d10; color: #e6e9ef; height: 100vh; display: flex; flex-direction: column; }
  header { display: flex; align-items: center; gap: 12px; padding: 8px 16px;
           background: #14161a; border-bottom: 1px solid #23272f; }
  header h1 { font-size: 15px; margin: 0; font-weight: 600; }
  header .spacer { flex: 1; }
  .badge { font-size: 12px; color: #8b94a7; }
  .timer { font-variant-numeric: tabular-nums; font-size: 13px; color: #38bdf8; }
  .timer b { color: #e6e9ef; font-weight: 600; }
  button { background: #1b1f26; color: #cfd6e4; border: 1px solid #2a2f3a; border-radius: 6px;
           padding: 4px 10px; font-size: 12px; cursor: pointer; }
  button:hover { background: #232833; }
  button.primary { background: #1d4ed8; border-color: #1d4ed8; color: #fff; }
  main { flex: 1; min-height: 0; display: flex; }
  .stage { flex: 1; display: flex; align-items: center; justify-content: center; position: relative;
           background: #0b0d10; }
  .stage img.current { max-width: 96%; max-height: 96%; border-radius: 6px;
                       box-shadow: 0 10px 50px rgba(0,0,0,.6); background: #fff; }
  .pageno { position: absolute; right: 18px; bottom: 14px; font-size: 12px; color: #8b94a7;
            background: rgba(20,22,26,.8); padding: 2px 10px; border-radius: 999px; }
  aside { width: 330px; background: #101216; border-left: 1px solid #23272f; display: none;
          flex-direction: column; }
  body.speaker aside { display: flex; }
  body.speaker .stage { flex: 1; }
  aside .next { padding: 12px; border-bottom: 1px solid #23272f; }
  aside .next img { width: 100%; border-radius: 4px; background: #fff; display: block; }
  aside .next .label { font-size: 11px; color: #8b94a7; margin-bottom: 6px; }
  aside .notes { flex: 1; overflow-y: auto; padding: 12px; font-size: 13px; line-height: 1.7;
                 white-space: pre-wrap; color: #d7dce6; }
  aside .notes .label { font-size: 11px; color: #8b94a7; margin-bottom: 6px; }
  aside .hints { padding: 8px 12px; font-size: 11px; color: #6b7385; border-top: 1px solid #23272f; }
</style>
</head>
<body>
<header>
  <h1 id="title">present</h1>
  <span class="badge" id="version-badge"></span>
  <span class="spacer"></span>
  <span class="timer">总 <b id="timer-total">00:00</b> · 本页 <b id="timer-page">00:00</b></span>
  <button id="btn-pause" onclick="togglePause()">⏸ 暂停</button>
  <button onclick="resetTimers()">↺ 重置计时</button>
  <button id="btn-speaker" onclick="toggleSpeaker()">🖥 演讲者视图</button>
  <button class="primary" onclick="goFullscreen()">⛶ 全屏</button>
</header>
<main>
  <div class="stage" id="stage" onclick="next()">
    <img class="current" id="page" alt="当前页渲染图">
    <span class="pageno" id="pageno"></span>
  </div>
  <aside>
    <div class="next">
      <div class="label">下一页预览</div>
      <img id="next-img" alt="下一页预览">
    </div>
    <div class="notes"><div class="label">演讲者备注</div><div id="note">（无备注）</div></div>
    <div class="hints">←/→ 或空格翻页 · S 演讲者视图 · P 暂停 · R 重置 · F 全屏 · Esc 退出</div>
  </aside>
</main>
<script>
const DOC_ID = "__DOC_ID__";
const base = "/api/docs/" + DOC_ID;
let total = 0, current = 0, notes = {}, title = "present", version = "";
let elapsedTotal = 0, elapsedPage = 0, paused = false;

const tick = setInterval(() => {
  if (paused) return;
  elapsedTotal += 1; elapsedPage += 1; renderTimer();
}, 1000);

function fmt(s) {
  const m = Math.floor(s / 60), r = s % 60;
  return String(m).padStart(2, "0") + ":" + String(r).padStart(2, "0");
}
function renderTimer() {
  document.getElementById("timer-total").textContent = fmt(elapsedTotal);
  document.getElementById("timer-page").textContent = fmt(elapsedPage);
}
function togglePause() {
  paused = !paused;
  document.getElementById("btn-pause").textContent = paused ? "▶ 继续" : "⏸ 暂停";
}
function resetTimers() { elapsedTotal = 0; elapsedPage = 0; renderTimer(); }
function toggleSpeaker() { document.body.classList.toggle("speaker"); refresh(); }
function goFullscreen() {
  const el = document.documentElement;
  if (document.fullscreenElement) document.exitFullscreen();
  else (el.requestFullscreen && el.requestFullscreen());
}
function imgUrl(i) { return base + "/slides/" + i + "/image?v=" + (version || "0"); }
function refresh() {
  document.getElementById("page").src = imgUrl(current);
  const nextImg = document.getElementById("next-img");
  nextImg.src = (current + 1 < total) ? imgUrl(current + 1) : imgUrl(current);
  nextImg.style.opacity = (current + 1 < total) ? "1" : ".35";
  document.getElementById("note").textContent = notes[current] || "（无备注）";
  document.getElementById("pageno").textContent = (current + 1) + " / " + total;
}
function next() { if (current < total - 1) { current += 1; elapsedPage = 0; refresh(); } }
function prev() { if (current > 0) { current -= 1; elapsedPage = 0; refresh(); } }

document.addEventListener("keydown", (e) => {
  if (e.key === "ArrowRight" || e.key === " " || e.key === "PageDown") { e.preventDefault(); next(); }
  else if (e.key === "ArrowLeft" || e.key === "PageUp") { e.preventDefault(); prev(); }
  else if (e.key === "Home") { current = 0; elapsedPage = 0; refresh(); }
  else if (e.key === "End") { current = total - 1; elapsedPage = 0; refresh(); }
  else if (e.key === "s" || e.key === "S") toggleSpeaker();
  else if (e.key === "p" || e.key === "P") togglePause();
  else if (e.key === "r" || e.key === "R") resetTimers();
  else if (e.key === "f" || e.key === "F") goFullscreen();
  else if (e.key === "Escape" && document.fullscreenElement) document.exitFullscreen();
});

fetch(base).then(r => r.json()).then(d => {
  title = d.title || "present";
  total = d.slide_count || 0;
  version = d.version_id || "v0";
  document.title = title + " · 演讲者模式";
  document.getElementById("title").textContent = title;
  document.getElementById("version-badge").textContent = version;
  (d.notes || []).forEach(n => { notes[n.slide_index] = n.text; });
  // 预热全部页图，翻页零等待
  for (let i = 0; i < total; i++) { const im = new Image(); im.src = imgUrl(i); }
  refresh();
  renderTimer();
}).catch(() => {});
</script>
</body>
</html>
"""


@app.get("/api/docs/{doc_id}/present")
def present_page(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    """v0.6 #5：全屏演示 + 演讲者视图（当前页/下一页预览/备注/排练计时）。

    渲染 PNG 直接放（LibreOffice 出图，保真 100%），浏览器只做翻页不重排。"""
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    page = _PRESENT_PAGE.replace("__DOC_ID__", doc_id)
    return HTMLResponse(page, headers={"Cache-Control": "no-store"})


@app.get("/api/docs/{doc_id}/lint")
def lint_document(doc_id: str, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    """v0.5 #7：按需 lint 当前版本（apply 响应也自动带 lint_issues）。"""
    from services import lint as lint_mod

    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    issues = lint_mod.lint_pptx(storage.current_path(doc_id))
    return {
        "doc_id": doc_id,
        "version_id": f"v{int(meta.get('version', 0))}",
        "issues": issues,
        "count": len(issues),
    }


@app.get("/api/docs/{doc_id}/slides/{slide_index}/image")
def slide_image(doc_id: str, slide_index: int, v: Optional[str] = None, thumb: bool = False, x_doc_password: Optional[str] = Header(default=None, alias="X-Doc-Password")):
    meta = storage.load_meta(doc_id)
    _assert_doc_open(meta, x_doc_password)
    if slide_index < 0 or slide_index >= int(meta.get("slide_count", 0)):
        raise FileNotFoundError("slide not found")
    version = int(meta.get("version", 0))
    if v:
        if not re.fullmatch(r"v\d+", v):
            raise HTTPException(status_code=400, detail="invalid version id")
        version = int(v[1:])
        known = {str(item.get("id")) for item in meta.get("versions", [])}
        if v not in known:
            raise FileNotFoundError("version not found")
    path = (
        storage.thumb_path(doc_id, version, slide_index)
        if thumb
        else storage.image_path(doc_id, version, slide_index)
    )
    if not path.exists():
        raise FileNotFoundError("image not rendered")
    return FileResponse(
        path,
        media_type="image/png",
        headers={"Cache-Control": "public, max-age=31536000, immutable"},
    )


if WEB_DIR.exists():
    app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
else:

    @app.get("/")
    def index_missing():
        return {"service": "ppt-light", "message": "web/ directory not found"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host=os.getenv("HOST", "0.0.0.0"),
        port=int(os.getenv("PORT", "8090")),
        log_level="info",
    )
