from __future__ import annotations

import json
import os
import re
import shutil
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DOCS_DIR = DATA_DIR / "docs"

_DOC_ID_RE = re.compile(r"^[a-f0-9]{32}$")
_locks: Dict[str, threading.Lock] = {}
_locks_guard = threading.Lock()


def ensure_dirs() -> None:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def doc_dir(doc_id: str) -> Path:
    if not _DOC_ID_RE.fullmatch(doc_id or ""):
        raise ValueError("invalid doc_id")
    return DOCS_DIR / doc_id


def lock_for(doc_id: str) -> threading.Lock:
    with _locks_guard:
        return _locks.setdefault(doc_id, threading.Lock())


def meta_path(doc_id: str) -> Path:
    return doc_dir(doc_id) / "meta.json"


def current_path(doc_id: str) -> Path:
    return doc_dir(doc_id) / "current.pptx"


def objects_path(doc_id: str) -> Path:
    return doc_dir(doc_id) / "objects.json"


def messages_path(doc_id: str) -> Path:
    return doc_dir(doc_id) / "messages.json"


def versions_dir(doc_id: str) -> Path:
    return doc_dir(doc_id) / "versions"


def render_dir(doc_id: str, version: int) -> Path:
    return doc_dir(doc_id) / "render" / f"v{version}"


def image_path(doc_id: str, version: int, slide_index: int) -> Path:
    return render_dir(doc_id, version) / f"slide-{slide_index + 1:03d}.png"


def thumb_path(doc_id: str, version: int, slide_index: int) -> Path:
    return render_dir(doc_id, version) / f"thumb-{slide_index + 1:03d}.png"


def load_messages(doc_id: str) -> List[Dict[str, Any]]:
    path = messages_path(doc_id)
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return data if isinstance(data, list) else []


def save_messages(doc_id: str, messages: List[Dict[str, Any]]) -> None:
    path = messages_path(doc_id)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(messages[-120:], ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def append_messages(doc_id: str, messages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    current = load_messages(doc_id)
    current.extend(messages)
    save_messages(doc_id, current)
    return current[-120:]


def list_docs() -> List[Dict[str, Any]]:
    ensure_dirs()
    docs: List[Dict[str, Any]] = []
    for path in DOCS_DIR.iterdir():
        if not path.is_dir() or not _DOC_ID_RE.fullmatch(path.name):
            continue
        try:
            meta = json.loads((path / "meta.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        messages = load_messages(path.name)
        last = messages[-1] if messages else None
        docs.append(
            {
                "doc_id": meta.get("doc_id", path.name),
                "title": meta.get("title") or Path(str(meta.get("file_name") or "presentation.pptx")).stem,
                "file_name": meta.get("file_name", "presentation.pptx"),
                "slide_count": meta.get("slide_count", 0),
                "version": meta.get("version", 0),
                "version_id": f"v{meta.get('version', 0)}",
                "created_at": meta.get("created_at"),
                "updated_at": meta.get("updated_at"),
                "message_count": len(messages),
                "last_message": (
                    {
                        "role": last.get("role"),
                        "content": str(last.get("content") or "")[:180],
                        "at": last.get("at"),
                        "kind": last.get("kind"),
                    }
                    if last
                    else None
                ),
            }
        )
    docs.sort(key=lambda item: item.get("updated_at") or "", reverse=True)
    return docs


def delete_doc(doc_id: str) -> None:
    path = doc_dir(doc_id)
    if not path.exists():
        raise FileNotFoundError("document not found")
    shutil.rmtree(path)
    with _locks_guard:
        _locks.pop(doc_id, None)


def load_meta(doc_id: str) -> Dict[str, Any]:
    path = meta_path(doc_id)
    if not path.exists():
        raise FileNotFoundError("document not found")
    return json.loads(path.read_text(encoding="utf-8"))


def save_meta(doc_id: str, meta: Dict[str, Any]) -> None:
    meta_path(doc_id).write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def touch_doc(doc_id: str, title: str | None = None) -> Dict[str, Any]:
    meta = load_meta(doc_id)
    if title is not None:
        meta["title"] = title.strip()[:100] or meta.get("title")
    meta["updated_at"] = now_iso()
    save_meta(doc_id, meta)
    return meta


def load_objects(doc_id: str) -> Dict[str, Any]:
    path = objects_path(doc_id)
    if not path.exists():
        raise FileNotFoundError("objects not found")
    return json.loads(path.read_text(encoding="utf-8"))


def save_objects(doc_id: str, payload: Dict[str, Any]) -> None:
    objects_path(doc_id).write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def create_doc(file_name: str, content: bytes, slide_count: int, slide_size: Dict[str, int]) -> Dict[str, Any]:
    doc_id = uuid.uuid4().hex
    d = doc_dir(doc_id)
    d.mkdir(parents=True, exist_ok=False)
    (d / "render").mkdir()
    versions_dir(doc_id).mkdir()
    current_path(doc_id).write_bytes(content)
    shutil.copy2(current_path(doc_id), versions_dir(doc_id) / "v0.pptx")
    ts = now_iso()
    meta = {
        "doc_id": doc_id,
        "title": Path(file_name).stem or file_name,
        "file_name": file_name,
        "slide_count": slide_count,
        "slide_size_emu": slide_size,
        "version": 0,
        "created_at": ts,
        "updated_at": ts,
        "versions": [
            {"id": "v0", "label": "upload", "created_at": ts, "size": len(content)}
        ],
    }
    save_meta(doc_id, meta)
    return meta


def next_version_id(meta: Dict[str, Any]) -> tuple[int, str]:
    version = int(meta.get("version", 0)) + 1
    return version, f"v{version}"


def _max_versions() -> int:
    try:
        return max(2, int(os.getenv("MAX_VERSIONS", "50")))
    except (TypeError, ValueError):
        return 50


def _prune_versions(doc_id: str, meta: Dict[str, Any]) -> None:
    """Keep only the newest MAX_VERSIONS entries; drop their pptx + render dirs.

    Without pruning, every apply stores a full pptx copy plus 2 PNGs per slide
    forever — an easy path to disk exhaustion on long-lived documents."""
    versions = meta.get("versions") or []
    cap = _max_versions()
    if len(versions) <= cap:
        return
    for old in versions[:-cap]:
        old_id = str(old.get("id") or "")
        if not re.fullmatch(r"v\d+", old_id):
            continue
        try:
            (versions_dir(doc_id) / f"{old_id}.pptx").unlink(missing_ok=True)
        except OSError:
            continue
        shutil.rmtree(render_dir(doc_id, int(old_id[1:])), ignore_errors=True)
    meta["versions"] = versions[-cap:]


def commit_work(
    doc_id: str,
    meta: Dict[str, Any],
    work_path: Path,
    new_version: int,
    label: str,
    slide_count: int,
    slide_size: Dict[str, int],
    objects_payload: Dict[str, Any],
) -> Dict[str, Any]:
    """把已经渲染成功的 work 文件提交为当前版本。"""
    version_id = f"v{new_version}"
    # atomic promotion: a concurrent download of current.pptx used to be able
    # to stream a torn copy while copy2 truncated+rewrote it in place
    tmp_version = versions_dir(doc_id) / f"{version_id}.pptx.tmp"
    tmp_current = current_path(doc_id).with_name("current.pptx.tmp")
    shutil.copy2(work_path, tmp_version)
    shutil.copy2(work_path, tmp_current)
    os.replace(tmp_version, versions_dir(doc_id) / f"{version_id}.pptx")
    os.replace(tmp_current, current_path(doc_id))
    try:
        work_path.unlink()
    except OSError:
        pass

    ts = now_iso()
    meta["version"] = new_version
    meta["updated_at"] = ts
    meta["slide_count"] = slide_count
    meta["slide_size_emu"] = slide_size
    meta.setdefault("versions", []).append(
        {
            "id": version_id,
            "label": label,
            "created_at": ts,
            "size": current_path(doc_id).stat().st_size,
        }
    )
    _prune_versions(doc_id, meta)
    save_objects(doc_id, objects_payload)
    save_meta(doc_id, meta)
    return meta


def reconcile_meta(doc_id: str) -> Dict[str, Any]:
    """Repair version drift: if versions/vN.pptx exists above meta.version
    (e.g. a crash between file write and meta save), adopt it.

    meta.json is written last during commit, so a crash can leave version
    files the metadata never references; the next commit would then reuse vN
    and clobber it. Adopting the orphan keeps history monotonic."""
    meta = load_meta(doc_id)
    current = int(meta.get("version", 0))
    max_found = current
    pattern = re.compile(r"v(\d+)\.pptx$")
    for path in versions_dir(doc_id).glob("v*.pptx"):
        match = pattern.fullmatch(path.name)
        if match:
            max_found = max(max_found, int(match.group(1)))
    if max_found > current:
        versions = meta.setdefault("versions", [])
        known = {str(v.get("id")) for v in versions}
        for number in range(current + 1, max_found + 1):
            candidate = f"v{number}"
            if candidate in known:
                continue
            candidate_path = versions_dir(doc_id) / f"{candidate}.pptx"
            if not candidate_path.exists():
                continue  # gapped orphan: adopt what exists, skip holes
            versions.append(
                {
                    "id": candidate,
                    "label": "recovered",
                    "created_at": now_iso(),
                    "size": candidate_path.stat().st_size,
                }
            )
        versions.sort(key=lambda item: int(str(item.get("id"))[1:]))
        meta["version"] = max_found
        meta["updated_at"] = now_iso()
        save_meta(doc_id, meta)
    return meta


def reconcile_all() -> int:
    """Run reconcile_meta over every doc; returns the number repaired."""
    ensure_dirs()
    repaired = 0
    for path in DOCS_DIR.iterdir():
        if not path.is_dir() or not _DOC_ID_RE.fullmatch(path.name):
            continue
        try:
            meta = json.loads((path / "meta.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        try:
            before = int(meta.get("version", 0))
            if int(reconcile_meta(path.name).get("version", 0)) != before:
                repaired += 1
        except (OSError, ValueError):
            continue
    return repaired


def release_lock(doc_id: str) -> None:
    """Drop a doc's lock entry (used when a create fails and rolls back)."""
    with _locks_guard:
        _locks.pop(doc_id, None)


def version_file(doc_id: str, version_id: str) -> Path:
    if not re.fullmatch(r"v\d+", version_id or ""):
        raise ValueError("invalid version_id")
    path = versions_dir(doc_id) / f"{version_id}.pptx"
    if not path.exists():
        raise FileNotFoundError("version not found")
    return path
