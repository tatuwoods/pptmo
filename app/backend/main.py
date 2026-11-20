from __future__ import annotations

import json
import logging
import os
import queue
import re
import shutil
import tempfile
import threading
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.datastructures import Headers as StarletteHeaders

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

import storage  # noqa: E402
from services import agent, ai, edits, parse, render  # noqa: E402

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
AUTO_APPLY_DESTRUCTIVE = os.getenv("AUTO_APPLY_DESTRUCTIVE", "0").strip().lower() in ("1", "true", "yes")
DESTRUCTIVE_ACTIONS = {"delete_textbox", "clear"}
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


app = FastAPI(title="PPT Light", version="0.1.2", docs_url="/api/docs-ui", openapi_url="/api/openapi.json")
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
        "ai_configured": bool(os.getenv("DEEPSEEK_API_KEY", "").strip()),
        "ai_model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
        "render_scale": RENDER_SCALE,
    }


def _load_doc(doc_id: str):
    meta = storage.load_meta(doc_id)
    objects_payload = storage.load_objects(doc_id)
    return meta, objects_payload


def _sse(event: str, data: Dict[str, Any]) -> str:
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _prepare_commit(doc_id: str, work_path: Path, new_version: int):
    """Render + parse the work file WITHOUT holding the doc lock.

    LibreOffice can take up to 180s per document; doing this inside the lock
    blocked delete/rename/apply/chat on the doc (and the global conversion lock
    blocked every other doc). Returns the prepared payloads for the commit.
    """
    render_dir = storage.render_dir(doc_id, new_version)
    try:
        page_count = render.render_pptx(work_path, render_dir, scale=RENDER_SCALE)
        _objs, objects_payload = parse.parse_pptx(work_path)
        del _objs
        if page_count != objects_payload.get("slide_count"):
            page_count = objects_payload.get("slide_count", page_count)
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


def _commit_prepared(
    doc_id: str,
    new_version: int,
    label: str,
    work_path: Path,
    render_dir: Path,
    objects_payload: Dict[str, Any],
    page_count: int,
    expected_version: Optional[int] = None,
    touched: Optional[List[int]] = None,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Take the doc lock, re-check the optimistic version, commit, and serialize.

    Raises PatchConflict (cleaning up the prepared artifacts) when the doc moved
    while rendering. Returns (committed_meta, objects_payload).
    """
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        if expected_version is not None and int(meta.get("version", 0)) != int(expected_version):
            _discard_prepared(work_path, render_dir)
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


def _commit_result(committed: Dict[str, Any], objects_payload: Dict[str, Any], touched: Optional[List[int]] = None) -> Dict[str, Any]:
    result = _serialize(committed, objects_payload)
    if touched is not None:
        result["touched_slides"] = sorted(set(touched))
    return result


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
        "ai_configured": bool(os.getenv("DEEPSEEK_API_KEY", "").strip()),
        "ai_model": os.getenv("DEEPSEEK_MODEL", "deepseek-chat"),
        "libreoffice": bool(shutil.which("soffice") or shutil.which("libreoffice")),
        "docs": len(storage.list_docs()),
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
    if not safe_name.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="只支持 .pptx 文件")
    data = file.file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"文件超过 {MAX_UPLOAD_BYTES // (1024 * 1024)} MB 限制",
        )
    if len(data) < 100:
        raise HTTPException(status_code=400, detail="文件为空或损坏")

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
        meta = storage.create_doc(
            file_name=safe_name,
            content=data,
            slide_count=int(parsed["slide_count"]),
            slide_size=parsed["slide_size_emu"],
        )
        doc_id = meta["doc_id"]
        storage.save_objects(doc_id, parsed)
        render.render_pptx(
            storage.current_path(doc_id),
            storage.render_dir(doc_id, 0),
            scale=RENDER_SCALE,
        )
        return _serialize(meta, parsed)
    except Exception:
        if doc_id:
            shutil.rmtree(storage.doc_dir(doc_id), ignore_errors=True)
        raise
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass


@app.get("/api/docs/{doc_id}")
def get_document(doc_id: str):
    meta, objects_payload = _load_doc(doc_id)
    return _serialize(meta, objects_payload)


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
def chat(doc_id: str, body: ChatIn):
    message = body.message.strip()
    if not message:
        raise HTTPException(status_code=400, detail="请输入修改要求")
    meta, objects_payload = _load_doc(doc_id)
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
    if requested_state == "ok":
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
def apply(doc_id: str, body: ApplyIn):
    label = (body.label or "编辑").strip()[:80] or "编辑"
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
def restore_version(doc_id: str, version_id: str):
    if not re.fullmatch(r"v\d+", version_id or ""):
        raise HTTPException(status_code=400, detail="invalid version id")
    # ---- phase 1 (short lock): allocate the version id ----
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
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
    return result


@app.get("/api/docs/{doc_id}/download")
def download(doc_id: str):
    meta = storage.load_meta(doc_id)
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


@app.get("/api/docs/{doc_id}/slides/{slide_index}/image")
def slide_image(doc_id: str, slide_index: int, v: Optional[str] = None, thumb: bool = False):
    meta = storage.load_meta(doc_id)
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
