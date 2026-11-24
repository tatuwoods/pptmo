# Changelog

## v0.4.0 — second-round optimization, 22 items (2026-11-24)

### A. Editing & AI (edits.py + agent.py, 8 items)
1. `batch_slide_ops` — one instruction applies the same op to many slides
   (duplicate/delete; page-range syntax `2-5`, lists `1,3,5`)
2. `apply_layout` — preset layouts (title / two_column / image_text / data)
   with `layout_catalog` introspection
3. `generate_summary` — AI reads the whole deck and inserts a
   directory/summary slide at a given position
4. `edit_chart` — rewrite chart data tables and series names
   (python-pptx chart part)
5. SmartArt read/edit — SmartArt node text is readable (`smartart_node`
   objects) and editable via normal text edits; rebuilding geometry is
   rejected
6. `edit_notes` — speaker notes read/write per slide, incl. batch note
   generation hint
7. `set_hyperlink` / `remove_hyperlink` — add/clear hyperlinks on runs
8. `highlight_text` — yellow (or specified color) highlight on matching runs

### B. Import/export (main.py, 4 items)
9. `GET /api/docs/{id}/export/pdf-notes` — handout PDF (slide + notes pages)
10. `GET /api/docs/{id}/export/outline` — plain-text outline export
11. `POST /api/docs/import-images` — multi-image upload → one PPTX page per image
12. `GET /api/docs/{id}/versions/{a}/compare/{b}` — text diff summary between
    two versions (added/removed/changed slides)

### C. Collaboration & security (main.py + storage.py, 4 items)
13. `POST /api/docs/{id}/share` — tokenized read-only share URL
    (`/share/{token}`: renders + download only, revocable)
14. Optional document password (hashed in meta; session open requires
    `X-Doc-Password`)
15. Audit log — every apply records time/instruction digest/op count/result
    to `audit.jsonl`; API returns the recent 50 entries
16. Macro detection — pptx containing `vbaProject.bin` becomes read-only
    preview with a warning

### D. Performance & stability (render.py + main.py + web, 4 items)
17. Render cache — same doc+version reuses rendered PNGs (no re-render on
    restore of the same version)
18. Background prewarm — upload finishes with thumbnail warm-up and missing
    page fill
19. Lazy thumbnails — >50-page decks load thumbs via IntersectionObserver
20. Auto cleanup — uploaded tmp files and >30-day-old version files purged on
    daily first startup (logged)

### E. AI prompt upgrades (agent.py, 2 items)
21. Style imitation — system prompt directs the agent to read reference
    slides, extract font/color/layout patterns and reuse them
22. Bilingual reporting — `detect_language` switches reply language to English
    when the user asks in English (default stays Chinese)

- `/api/health` reports version 0.4.0
- Test suite grew from 311 to 403 cases, all green
- `garble_fix.py` font mapping logic untouched

## v0.3.0 — full capability upgrade, 22 items (2026-11-22)

### A. Editing (edits.py + agent.py)
1. `font_unify` — unify fonts across slide/deck (latin + east-asian + cs
   typefaces; size/bold differences preserved; symbol fonts like Wingdings
   skipped; paragraph defaults `defRPr`/`endParaRPr`/`lstStyle` covered;
   deck scope also covers masters/layouts)
2. `replace_text {find, replace, scope}` — deck-wide replace; per-run replace
   keeps other runs' styles, run-boundary-spanning needles fall back to a
   whole-text rewrite; zero matches is an explicit error
3. `align_objects {shape_ids, align}` — left/right/top/bottom/hcenter/vcenter
   + distribute_h/distribute_v (≥3 shapes, same slide, non-grouped)
4. `duplicate_object {shape_id, offset_emu}` — in-page copy of textboxes,
   pictures, groups (new cNvPr ids; same-part image rIds stay valid)
5. `z_order {shape_id, mode}` — front/back/forward/backward
6. `table_mod` — add/delete row/col, merge_cells (non-origin text cleared),
   split_cells; merge-overlap and merge-intersection guards; frame ext updated
7. `replace_image` — swap picture content by shape_id (base64 or local path;
   png/jpg magic validated; position/size preserved)
8. `slide_background` — solid color or two-stop gradient per slide (raw
   `p:bg` XML, replaces existing background)
9. `add_page_numbers {skip_first, position, start}` — batch page numbers,
   idempotency guard, bottom-left/center/right
10. `font_scale {delta_pt}` — slide/deck font size ±N pt (inherited sizes
    resolved from paragraph defaults, then made explicit; 6–400 pt clamped)
- All ten wired through the agent op store with scope validation
  (deck-wide ops need deck scope; “全局/整份/所有页” instructions auto-expand
  the scope), dedupe keys and op summaries; system prompt documents the schemas
- Fixed a latent v0.2 bug: `int(x or -1)` mapped slide_index 0 to -1, so the
  agent could never target objects on page 1

### B. Import/export (main.py)
11. `GET /api/docs/{id}/export/pdf` — render chain now keeps the produced PDF
    (`export.pdf` in the version's render dir); missing PDFs are regenerated once
12. `GET /api/docs/{id}/export/images.zip` — all rendered slide PNGs zipped
13. `POST /api/docs/{id}/duplicate` — copy the current document into a fresh
    session (same content/objects, v0, “副本” title, system note; rollback on
    render failure)

### C. Chat experience (agent.py + main.py)
14. `plan` tool — complex instructions get a step list first (“我将：1… 2… 3…”),
    streamed as a `plan` event and reported step-by-step
15. `undo_last` tool + `POST /api/docs/{id}/undo` — roll back the last applied
    round via the version snapshot (creates a new version, so nothing is lost);
    ⌘Z in the UI calls the endpoint
16. Self-healing — transient AI errors retry once without consuming a step
    (`retry` event); failed tool calls return an Observation with retry
    guidance instead of surfacing raw errors

### D. UI (web/)
17. Shortcuts: ⌘Z undo, ⌘Enter send, Esc stops generation first
18. Drag-drop upload anywhere on the window with a visible overlay; legacy
    .ppt/.pptm drops get the conversion hint
19. Thumbnail rail click switches canvas + chat scope together
20. Dark mode follows `prefers-color-scheme` (variable groups for all core
    surfaces; `color-scheme: light dark`)

### E. Stability
21. Render timeout (180s, `RENDER_TIMEOUT_S`) degrades instead of
    failing the whole commit: pages already rendered / stale previews from the
    previous version are kept, response carries `render_partial`; the web UI
    surfaces it twice — a toast when the degraded state first appears and a
    persistent status-bar note (“预览暂为上一版画面”) until the next
    successful render, so stale previews are never mistaken for the latest
22. Upload validation: .ppt/.pptm/.pps → “另存为 .pptx” hint; empty files,
    tiny/corrupt files and oversize files get friendly, specific messages

- `/api/health` reports version 0.3.0
- python-pptx stays pinned at 0.6.23; garble_fix Wingdings remap untouched

### Fixes landed on main after the v0.3.0 tag (2026-11-22)
- `replace_text` replaced twice whenever the replacement contained the find
  term (10 → 100 came out as 1000, `v1` → `v1.1` as `v1.1.1`), and a no-op
  replace collapsed mixed run formatting. Decided per paragraph now: an
  in-run hit rewrites that run only; a needle spanning runs rebuilds just that
  paragraph from its original text
- Placeholder inheritance (layout placeholders usually carry no `a:xfrm` and
  their runs no `sz`): `align_objects` and `geometry` no longer pin the axis
  they do not edit to 0; `duplicate_object` gives the copy an explicit box
  instead of stacking it under the original; `font_scale` resolves the
  inherited size (run → paragraph → shape list style → layout/master
  placeholder → master text styles) instead of assuming 18pt
- `add_page_numbers` repeat guard keys on the name it stamps on its own
  boxes — a small text box that merely reads like a page number no longer
  blocks numbering
- `z_order front` stays before a trailing `p:extLst`
- `table_mod` add/delete row/col keep the frame size in step with the grid
- `table_mod merge_cells` keeps every cell's text (python-pptx already
  concatenates them into the origin like PowerPoint; the pre-clear dropped
  them) and accepts the two corners in either order — the origin is always
  the top-left cell
- `add_page_numbers` writes a live `slidenum` field when numbering from 1, so
  numbers follow pages that are duplicated/deleted/moved (start ≠ 1 still
  writes plain digits: LibreOffice ignores the deck's start number in fields);
  duplicated slides/objects get fresh field ids; without `start` the deck's
  own first-slide number is honored
- Exports under a render-timeout degradation (E21): re-renders for export go
  to a scratch dir and only replace the committed one on success, so a second
  timeout can't wipe the previews the UI is showing; `export/images.zip` no
  longer ships the previous version's stale pages; a successful re-render
  clears `_partial` + `render_partial`; a duplicate never inherits the
  degraded flag (ported from the `fix/v0.3-audit` branch)
- `font_scale` moves `a:fld` fields (the live page-number) with the text
  around them — the field's rendered size is resolved through the same
  inheritance chain as runs and pinned explicitly, so a footer line can't
  end up with two font sizes after a shift (acceptance pass on
  samples/demo.pptx, `test_accept_a10_font_scale.py`)

## v0.2.0 — slide-level operations (2026-11-21)

- Four slide-level actions join the op pipeline (`propose_operations` → review
  cards → `/apply`), alongside the existing shape-level ones:
  - `duplicate_slide {source_slide, position}` — full-fidelity copy (text,
    styles, pictures, tables incl. merged cells, groups, background,
    transitions, speaker notes); image/chart/hyperlink relationships are
    re-bound to the same parts in the new slide's rels
  - `add_slide {after_slide?, layout_index?}` — blank page, appended by default
  - `delete_slide {slide}` — rejects deleting the only slide
  - `move_slide {slide, to}` — reorder via `_sldIdLst`
- python-pptx 0.6.23 stays pinned (it has no slide-copy API; the implementation
  works on the OPC relationship graph + `sldIdLst` directly)
- Mixed batches are deterministic: shape-level ops resolve against the
  pre-batch numbering (matching objects.json), slide-level ops then run in
  submission order
- `objects.json` is fully re-parsed on commit, so every `s{n}_id{m}` prefix is
  renumbered after any slide-level op; later agent rounds see the new IDs
- `delete_slide` joins `delete_textbox`/`clear` in the destructive auto-apply
  guard (manual confirmation unless `AUTO_APPLY_DESTRUCTIVE=1`)
- `base_version` conflict protection and partial-apply bookkeeping apply to
  slide-level ops unchanged
- Agent system prompt documents the new actions, the renumbering rule and the
  recommended two-round flow for "duplicate then edit the copy"
- 39 new tests (test_slide_ops.py): picture/table/group/notes fidelity,
  object-ID shift after delete, reorder, out-of-range params, API conflict
  guard, agent op-store validation, partname uniqueness after delete —
  126 total
- Frontend renders dedicated cards for the four actions; locate jumps to the
  affected slide

## v0.1.5 — UX polish & docs (2026-11-20)

- Inline editor confirms before discarding typed changes on navigation-away
  (Escape still cancels silently — explicit intent)
- Upload size cap served by the backend (`max_upload_mb` in `/api/health` and
  doc payloads); frontend no longer hardcodes 60 MB
- `/api/health` reports `version`
- CHANGELOG, README security section, SKILL.md env-knob reference

## v0.1.4 — API robustness (2026-11-20)

- Atomic `current.pptx` promotion (`os.replace`): no torn downloads mid-commit
- Startup reconciliation adopts orphaned `versions/vN.pptx` after a crash
- `pending_operations` served on doc payloads — staged AI ops survive message
  trimming (they used to vanish after 60 messages); frontend prefers it
- `restore` accepts optional `base_version` → 409 on drift; UI refreshes
- SSE frame delimiter CRLF-tolerant; chat aborts after 90s of stream silence
- `[DONE]` sentinel matched in the streaming client
- bbox explicit `w/h ≤ 0` rejected (was silently replaced by slide fractions)
- patch `slide_index` must agree with the `shape_id` slide prefix
- `_trim_messages` clips oversized JSON observations per-field
- per-doc lock released on failed doc creation

## v0.1.3 — performance & resource management (2026-11-20)

- Two-phase commit (apply / restore / chat auto-apply): LibreOffice render +
  parse run **outside** the doc lock; commit re-checks the optimistic version
  under a short lock → a 180s render no longer blocks delete/rename/apply/chat,
  and concurrent applies resolve with a clean 409
- Render queue bounded (`RENDER_QUEUE_MAX=4`) → 429 instead of head-of-line
  blocking; page cap `RENDER_MAX_PAGES=500`
- Chat concurrency: `AI_CONCURRENCY=2` semaphore + one in-flight chat per doc
- Structural parse caps: `PARSE_MAX_SLIDES=500`, `PARSE_MAX_SHAPES=5000`,
  `PARSE_MAX_GROUP_DEPTH=20`; uploads above the caps rejected (413)
- Version pruning: `MAX_VERSIONS=50` (drops old pptx + render dirs)
- Drag commits only above ~0.05 cm real movement; thumbnail DOM reused on save

## v0.1.2 — security hardening (2026-11-20)

- `RequestGuardMiddleware`: DNS-rebinding Host check (421), cross-site
  mutation guard via `Sec-Fetch-Site`/`Origin` (403), request-body caps
  (JSON 2 MB default; multipart = upload cap + slack) → 413 before buffering
- `/apply` patch strings capped at 20k chars; fuzzy before-match skipped for
  huge values (quadratic CPU)
- Error details sanitized (server paths masked, logged server-side)
- Auto-apply refuses `delete_textbox`/`clear` unless `AUTO_APPLY_DESTRUCTIVE=1`
- `TRUSTED_HOSTS` / `TRUSTED_ORIGINS` env for reverse-proxy (tailscale serve)
  deployments

## v0.1.1 — critical correctness (2026-11-20)

- Chat finalize moved into the worker thread: client disconnect (Stop, tab
  close) can no longer lose the assistant message or validated auto-apply ops
- `/apply` marks exactly the message it came from (`assistant_message_id`);
  partial applies recorded as `partial_applies` instead of wrongly marking a
  whole message applied (which silently un-restored pending ops)
- Out-of-range page mentions answered with 400 instead of silently editing the
  open slide
- `work-*.pptx` leaked on rejected patches — cleaned up
- Frontend: stale-ops guard keyed on the ops' `base_version`; chat 409 auto
  recovery with input restored; doc switch aborts in-flight streams; finalize
  targets the originating doc; dead code removed

## v0.1.0 — initial public release

PPT Light as packaged in this repo (skill + app), plus the pytest suite
(`tests/`) and live smoke script (`scripts/smoke_test.sh`).
