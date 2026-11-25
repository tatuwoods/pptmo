---
name: pptmo
description: Deploy and drive PPT Light (pptmo), a local conversational AI PPTX editor with high-fidelity preview, versioning, and precise agent-driven edits written back to a real editable .pptx. Use this skill whenever the user wants to edit/revise/polish PowerPoint files through conversation, asks to set up or start the ppt-light service, wants AI to read a deck and make precise text/style/geometry changes, needs the PPT editing API (upload/chat-SSE/apply/download) called programmatically, or wants to expose the local service remotely via Tailscale. Also use it when the user mentions "PPT 改稿", "对话改 PPT", uploading a .pptx for AI editing, or recovering/rolling back PPT versions.
---

# pptmo — Conversational AI PPTX Editor (PPT Light)

A self-hosted web app: upload a `.pptx`, get a high-fidelity render, then edit it either manually (double-click text, drag boxes) or by chatting with an AI agent (ReAct loop) that reads the deck and submits **validated, reviewable operations**. Applying them writes a real editable PPTX via `python-pptx` and creates a restorable version.

Stack: FastAPI + python-pptx + LibreOffice(`soffice`) → PDF → PyMuPDF PNG; vanilla JS frontend; any OpenAI-compatible chat API for the AI agent.

## Layout of this skill

```
app/        # full application source (backend/ web/ samples/demo.pptx)
scripts/    # setup.sh — one-shot environment bootstrap
```

## When the user just wants it running (fast path)

```bash
bash <skill-dir>/scripts/setup.sh <target-dir>          # copies app, creates venv, installs deps
cd <target-dir>/backend
.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8090
curl -s http://127.0.0.1:8090/api/health                # {"ok": true, ...}
```

Open http://127.0.0.1:8090 . Without an AI key everything works except the chat pane; with a key configured (below) AI editing unlocks.

Prerequisites: Python 3.10+ (tested on 3.12 and 3.14), `soffice` on PATH (LibreOffice; on macOS `brew install --cask libreoffice` puts a CLI at `$(brew --prefix)/bin/soffice`). Data lives in `backend/data/docs/`; deleting a session in the UI deletes its files.

## AI configuration

The agent calls any **OpenAI-compatible** `/chat/completions` endpoint. Configure via `backend/.env` (chmod 600, never commit):

```dotenv
DEEPSEEK_API_KEY=<key>
DEEPSEEK_API_BASE=<base-url-ending-in-/v1>
DEEPSEEK_MODEL=<model>
RENDER_SCALE=2
# optional: AI_MAX_STEPS=10 (2-16), AI_MAX_OPS=24 (4-60), MAX_UPLOAD_MB=60
```

The env names say "DEEPSEEK" but the provider is arbitrary. Verified working:

| Provider | BASE | MODEL | Notes |
| --- | --- | --- | --- |
| Zhipu coding plan | `https://open.bigmodel.cn/api/coding/paas/v4` | `glm-5.3-flash` | supports `response_format: json_object` + SSE; reasoning model, a few seconds slower per step |
| Zhipu standard | `https://open.bigmodel.cn/api/paas/v4` | `glm-4.6` etc. | |
| DeepSeek official | `https://api.deepseek.com/v1` | `deepseek-chat` | |

The client already retries without `response_format` once if an endpoint rejects it, and it tolerates `delta.reasoning_content`-style models (it only concatenates `delta.content`). Key must exist or `ai_configured` stays false — a placeholder value turns the badge green but every chat will 502, so leave it empty rather than fake.

## How the AI edit flow works (what to tell the user)

1. AI **cannot** write the file directly. It calls read tools (`outline`, `read_slides`, `search_text`, `read_table`, `inspect_shapes`), then submits ops via `propose_operations` (≤6 per call, ≤24 per turn).
2. Default is **review mode**: operations appear as cards, applied only after the user clicks apply. The "自动执行" toggle auto-applies after the turn.
3. Guards you can rely on: ops must target real `shape_id`s from tool output; text edits carry a `before` checked against current text (page-number/footer/logo objects are protected unless the instruction names them); grouped shapes allow text/style edits but reject geometry; every apply creates a version that can be restored (restore itself creates a new version, history is never lost).
4. Optimistic-locking: requests carry `base_version`; if the doc changed meanwhile the server answers 409 and ops stay staged.
5. Slide-level ops (v0.2): `duplicate_slide` (full-fidelity page copy: text/styles/pictures/tables/groups/background/notes), `add_slide` (blank page), `delete_slide`, `move_slide` — page numbers are 0-based like everywhere else. In a mixed batch, shape edits resolve against the pre-batch numbering and slide ops then run in submission order; after apply every `s{n}_…` id renumbers, so "copy a page then edit the copy" is a two-round flow (duplicate → apply → re-read → edit).

## Driving it programmatically (no UI needed)

The whole feature set is HTTP. Typical agent loop:

```bash
# 1. upload
DOC=$(curl -s -X POST http://127.0.0.1:8090/api/docs/upload \
  -F "file=@deck.pptx" | python3 -c 'import json,sys;print(json.load(sys.stdin)["doc_id"])')

# 2. chat (SSE stream of events: start/step/thought/tool_call/tool_result/operations/summary/done)
curl -sN -X POST http://127.0.0.1:8090/api/docs/$DOC/chat -H "Content-Type: application/json" \
  -d '{"message":"把第4页的项目总投资改成190万，保留格式","scope":"current","slide_index":3,"base_version":0}'
# done event contains operations[] — each has id/action/shape_id/before/after/description

# 3. apply accepted operations (auto_apply=false flow)
curl -s -X POST http://127.0.0.1:8090/api/docs/$DOC/apply -H "Content-Type: application/json" \
  -d '{"patches":[<op objects>],"label":"AI 改稿","base_version":0}'

# 4. download the real pptx
curl -sO http://127.0.0.1:8090/api/docs/$DOC/download
```

Slide-level ops (v0.2) go through the same channels — chat tool-call or direct
`/apply` patches (all page numbers are 0-based `slide_index`, same as shape_id
prefixes):

```json
{"action":"duplicate_slide","source_slide":0,"position":1,"description":"copy cover after itself"}
{"action":"add_slide","after_slide":3,"layout_index":6}
{"action":"delete_slide","slide":2}
{"action":"move_slide","slide":4,"to":0}
```

`duplicate_slide` copies everything (text, styles, pictures, tables, groups,
background, notes) by re-binding relationships to the same image/chart parts.
Shape-level ops in a mixed batch resolve against the pre-batch numbering;
after any slide-level op commits, all `s{n}_id{m}` IDs are renumbered on the
next parse — re-read before editing the new slides. `delete_slide` is
destructive-guarded for auto-apply like `clear`/`delete_textbox`.

Say “第 N 页” in the message and the server auto-scopes to that slide (Chinese numerals, “首页”, "page N" all work). `scope: "deck"` widens to the whole file. SSE is plain `event:`/`data:` lines — parse incrementally; the `done` payload holds `message_id`, final `operations`, `trace`, and (if auto-apply) `applied_version_id`.

Useful endpoints beyond these: `GET /api/sessions`, `GET /api/docs/{id}` (objects+messages in one shot), `GET /api/docs/{id}/versions`, `POST .../versions/{v}/restore`, `GET /api/docs/{id}/slides/{n}/image`. Interactive docs at `/api/docs-ui`.

## v0.5 operations (rebrand / merge / transitions / hide / lint / diff / watch)

Deck-wide ops go through the same `/apply` channel as patches:

```json
{"action":"theme_rebrand","payload":{"palette":{"accent1":"1A1A2E","accent2":"16213E"},"fonts":{"major":"Source Han Sans","minor":"Source Han Sans"}}}
{"action":"recolor","payload":{"mapping":{"FF0000":"0066CC","000000":"222222"}}}
{"action":"set_transition","payload":{"pages":"all","effect":"push","direction":"l","speed":"slow","advance_after":3}}
{"action":"hide_slide","payload":{"pages":"2-3"}}        // show_slide unhides; parse marks slides[].hidden
{"action":"edit_props","payload":{"title":"Q3 复盘","author":"张三","keywords":["销售","复盘"]}}
{"action":"set_alt_text","shape_id":"s0_id4","alt_text":"柱状图：Q3 销售额"}   // pictures now appear in objects (kind=picture)
{"action":"fix_issues","payload":{"kinds":["off_canvas","text_overflow"]}}    // deterministic lint self-fix
```

- **theme_rebrand vs recolor**: the first rewrites theme scheme colors + major/minor fonts (shapes referencing `accent1` etc. follow); the second remaps hard-coded fills (incl. gradient stops) via an exact hex mapping. Use both to re-brand a template deck.
- **merge**: `POST /api/docs/{id}/merge` with multipart `file` (source pptx), `pages` (1-based selector, empty = all), `position` (0-based insert index). Layouts/masters/images are re-homed across the OPC package; notes travel with the page.
- **lint**: every apply response carries `lint.issues` (text overflow / font size out of range / off-canvas / overlap warnings); `GET /api/docs/{id}/lint` checks on demand. `fix_issues` repairs deterministically (shrink font / pull back into canvas); overlaps are warn-only. Let the AI read `issues` and decide whether to fix.
- **structural diff**: `GET /api/docs/{id}/versions/{a}/compare/{b}` now also returns `struct` — object-level added/removed, geometry changes (which axis, moved vs resized), style changes (font/size/bold/color), matched by `shape_id` per page. JSON for AI consumption.
- **watch**: open `GET /api/docs/{id}/watch` in a browser — split-pane live preview. Changes publish SSE events on `GET /api/docs/{id}/events` (`hello` on connect, `changed` with version/touched_slides on every commit: apply/chat-auto-apply/undo/restore/merge).
- **hidden pages & export**: export PDF / images.zip / pdf-notes skip hidden slides by default; add `?include_hidden=1` to include them. Outline marks hidden pages「（已隐藏）」.

## Remote access via Tailscale

Best practice: keep uvicorn bound to `127.0.0.1` and publish through `tailscale serve` (auto-HTTPS, tailnet-only, streams SSE fine). **Since v0.1.2 the app runs a request guard** (anti drive-by / DNS-rebinding), so proxied browser traffic needs the public host whitelisted in `backend/.env` or every cross-origin POST gets 403:

```dotenv
TRUSTED_HOSTS=macbookpro.example.ts.net
TRUSTED_ORIGINS=https://macbookpro.example.ts.net:9443
```

(Hostnames listed in TRUSTED_ORIGINS are trusted automatically; loopback is always trusted, so plain `curl` keeps working.)

```bash
tailscale serve --bg --https=9443 http://127.0.0.1:8090
tailscale serve status        # prints https://<host>.<tailnet>.ts.net:9443/
```

On macOS with a **homebrew userspace** tailscaled (common on agent boxes), the CLI must point at its socket or every command fails with "failed to connect to local tailscaled":

```bash
TS="tailscale --socket=$HOME/.local/state/tailscale/tailscaled.sock"
$TS serve --bg --https=9443 http://127.0.0.1:8090
```

Find the socket via `ps aux | grep tailscaled` (`--socket=` flag). Userspace mode has **no tun device**, so bare connections to the machine's 100.x IP will NOT work — `serve` is the way. To test from the serving machine itself in userspace mode, route through its SOCKS5 proxy: `curl --proxy socks5h://127.0.0.1:1055 https://<host>.<tailnet>.ts.net:9443/api/health` (note: curl sends no Origin header, so the guard passes it — this does NOT prove browser access works; check the TRUSTED_* config). Undo with `tailscale serve --https=9443 off`.

**Coexisting with other serve users (important).** All serve entries share one config table on the node; different ports coexist by design, but the same port is **last-writer-wins** — issuing `serve --https=443 ...` silently replaces whatever 443 pointed at. Port 443 is the "default" everyone (humans and agent bots) tends to grab. So: always publish this app on a dedicated high port (9443 here) and leave 443 alone; check `tailscale serve status` first to see which ports are taken. If full isolation is ever required (multiple tenants fighting over serve), the escalation path is a second tailscaled instance with its own `--state`/`--socket` (a separate tailnet node, needs its own login), but a distinct port on the shared daemon is almost always enough.

Do NOT rebind uvicorn to `0.0.0.0` instead: the app has no auth and open CORS; the tailnet HTTPS proxy is the only exposure that keeps it private.

## Environment knobs (v0.1.5)

| Env | Default | Meaning |
| --- | --- | --- |
| `TRUSTED_HOSTS` | loopback | extra Host names accepted (reverse proxy) |
| `TRUSTED_ORIGINS` | loopback | extra browser Origins accepted; hostnames auto-trusted |
| `MAX_JSON_BODY_MB` | 2 | cap for JSON request bodies (apply/chat) |
| `AUTO_APPLY_DESTRUCTIVE` | 0 | allow auto-apply of delete_textbox/clear/delete_slide |
| `AI_CONCURRENCY` | 2 | max simultaneous agent chats (429 beyond) |
| `RENDER_QUEUE_MAX` | 4 | queued renders before 429 |
| `RENDER_MAX_PAGES` | 500 | LibreOffice render page cap |
| `PARSE_MAX_SLIDES` | 500 | upload rejected (413) above this |
| `PARSE_MAX_SHAPES` | 5000 | shapes cap (recursive groups incl.) |
| `PARSE_MAX_GROUP_DEPTH` | 20 | nested group cap |
| `MAX_VERSIONS` | 50 | versions kept per doc (old pruned with renders) |
| `AI_MAX_STEPS` / `AI_MAX_OPS` | 10 / 24 | agent loop limits (2–16 / 4–60) |
| `MAX_UPLOAD_MB` | 60 | upload cap; served to the frontend |

Behavior notes (v0.1.1–v0.1.5): chat finalize persists even on client disconnect; `/apply` marks exactly the sourced message (send `assistant_message_id`), partial applies recorded as `partial_applies`; out-of-range "第 N 页" → 400; `pending_operations` field keeps staged ops alive past message trimming; renders happen outside the doc lock (two-phase commit with version re-check).

## Troubleshooting

- **`ai_configured: false`** in `/api/health` → missing/empty `DEEPSEEK_API_KEY`; env is read at process start, so restart after editing `.env`.
- **Chat 502 / "AI 服务异常"** → wrong base URL (must include `/v1`-style suffix the provider expects), invalid/expired key, or proxy interference; the backend honors `HTTPS_PROXY`.
- **Upload 400 “无法解析 PPTX”** → corrupt or non-pptx; only `.pptx` (no `.ppt`).
- **First open of a big deck is slow** → LibreOffice renders synchronously (all pages, 2x PNG); tens of seconds for 50+ pages is normal, later versions only rerender once per apply.
- **`render/` grows unboundedly** → every version keeps full-page PNGs by design; clean old `backend/data/docs/<id>/render/vN/` dirs if disk matters (thumbnails of old versions will 404 harmlessly).
- **Port 8090 busy** → launch with `--port 8091`, or run `python main.py` which reads `PORT`/`HOST` env vars.

## Known limits (set expectations honestly)

Text layer + textbox geometry only — pictures/charts/SmartArt render but can't be edited (whole slides containing them can still be duplicated/moved since v0.2); table cells accept text/style but not move/resize; grouped text can be reworded but not repositioned; fonts may render slightly differently vs PowerPoint (LibreOffice rasterizes the preview; the downloaded file opens fine in PowerPoint/WPS). 60 MB upload cap. No multi-user auth — keep it behind localhost or a private tailnet.
