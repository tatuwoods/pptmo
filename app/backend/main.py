from __future__ import annotations

import json
import logging
import os
import queue
import re
import shutil
import threading
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

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

app = FastAPI(title="PPT Light", version="0.1.0", docs_url="/api/docs-ui", openapi_url="/api/openapi.json")
app.add_middleware(GZipMiddleware, minimum_size=1500)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


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
    from fastapi.responses import JSONResponse

    return JSONResponse(status_code=409, content={"detail": str(exc)})


@app.exception_handler(edits.PatchError)
async def _patch_error_handler(_request, exc: edits.PatchError):
    from fastapi.responses import JSONResponse

    return JSONResponse(status_code=400, content={"detail": str(exc)})


@app.exception_handler(ai.AIServiceError)
async def _ai_error_handler(_request, exc: ai.AIServiceError):
    from fastapi.responses import JSONResponse

    return JSONResponse(status_code=502, content={"detail": str(exc)})


@app.exception_handler(render.RenderError)
async def _render_error_handler(_request, exc: render.RenderError):
    from fastapi.responses import JSONResponse

    log.exception("render failed")
    return JSONResponse(status_code=500, content={"detail": f"渲染失败：{exc}"})


@app.exception_handler(FileNotFoundError)
async def _not_found_handler(_request, exc: FileNotFoundError):
    from fastapi.responses import JSONResponse

    return JSONResponse(status_code=404, content={"detail": str(exc) or "not found"})


@app.exception_handler(ValueError)
async def _value_error_handler(_request, exc: ValueError):
    from fastapi.responses import JSONResponse

    return JSONResponse(status_code=400, content={"detail": str(exc)})


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


def _commit_work(
    doc_id: str,
    meta: Dict[str, Any],
    work_path: Path,
    new_version: int,
    label: str,
    touched: Optional[List[int]] = None,
) -> Dict[str, Any]:
    render_dir = storage.render_dir(doc_id, new_version)
    try:
        page_count = render.render_pptx(work_path, render_dir, scale=RENDER_SCALE)
        objects_payload_objs, objects_payload = parse.parse_pptx(work_path)
        del objects_payload_objs
        if page_count != objects_payload.get("slide_count"):
            page_count = objects_payload.get("slide_count", page_count)
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
    except Exception:
        try:
            work_path.unlink()
        except OSError:
            pass
        shutil.rmtree(render_dir, ignore_errors=True)
        raise
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

        if auto_apply and operations:
            if int(meta.get("version", 0)) != int(start_version):
                apply_error = "文档在对话期间已更新，操作已暂存，请刷新后重试"
            else:
                new_version, _ = storage.next_version_id(meta)
                work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
                applied_doc: Optional[Dict[str, Any]] = None
                try:
                    touched = edits.apply_patches_to_file(
                        storage.current_path(doc_id), work_path, operations
                    )
                    applied_doc = _commit_work(
                        doc_id,
                        meta,
                        work_path,
                        new_version,
                        f"AI 对话改稿 · {len(operations)} 项",
                        touched=sorted(touched),
                    )
                except Exception as exc:
                    apply_error = str(exc) or "应用操作失败"
                    try:
                        work_path.unlink()
                    except OSError:
                        pass
                if applied_doc is not None:
                    applied = True
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

    def worker() -> None:
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

    threading.Thread(target=worker, daemon=True, name=f"agent-{doc_id[:8]}").start()

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
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        if body.base_version is not None and int(body.base_version) != int(meta.get("version", 0)):
            raise edits.PatchConflict("文档已经更新，请刷新后再试")
        new_version, _ = storage.next_version_id(meta)
        work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
        try:
            touched = edits.apply_patches_to_file(
                storage.current_path(doc_id), work_path, body.patches
            )
            result = _commit_work(
                doc_id, meta, work_path, new_version, label, touched=sorted(touched)
            )
        except Exception:
            # a rejected patch used to leak the work file; _commit_work cleans
            # up only for failures that happen inside it
            try:
                work_path.unlink()
            except OSError:
                pass
            raise
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
    log.info("applied %d patches to %s -> v%d", len(body.patches), doc_id, new_version)
    return result


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
    with storage.lock_for(doc_id):
        meta = storage.load_meta(doc_id)
        source = storage.version_file(doc_id, version_id)
        new_version, _ = storage.next_version_id(meta)
        work_path = storage.doc_dir(doc_id) / f"work-{uuid.uuid4().hex}.pptx"
        try:
            shutil.copy2(source, work_path)
            result = _commit_work(doc_id, meta, work_path, new_version, f"恢复 {version_id}")
        except Exception:
            try:
                work_path.unlink()
            except OSError:
                pass
            raise
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
