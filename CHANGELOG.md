# Changelog

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
- 34 new tests (test_slide_ops.py): picture/table/group/notes fidelity,
  object-ID shift after delete, reorder, out-of-range params, API conflict
  guard, agent op-store validation — 121 total
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
