# Changelog

## v0.6.0 — 对标 2026-11 最新 Top5，渲染验证闭环 + 保真 + 可编辑性 (2026-11-26)

对标 guizang-ppt-skill / ppt-master / image-to-editable-ppt-skill / KitionAI
kition-slide / CasualOffice slides（awesome-codex-ppt-skills 实测榜），
核心洞察「pick by whether you can edit the result」：

1. 渲染验证闭环（vision 自检）— `GET /api/docs/{id}/vision-check`：把当前
   版本渲染 PNG 转 base64 视觉消息喂多模态模型（glm-5.3-flash 等，溢出/
   乱码豆腐/塌版/对比度四类检查）；agent 新增 vision_check 工具（apply 响应
   带 vision_check_url，前端「👁 自检」按钮），发现问题自提案修复——对标
   OfficeCLI render→look→fix
2. OOXML passthrough 保真 — apply/merge 保存后未触碰的 slide part（含
   .rels/notes）按 part 名原字节回写，改一处不再丢别处（动画 timing/
   自定义扩展元素保真）；`PPTMO_PASSTHROUGH=0` 可关——对标 CasualOffice
   未映射 XML 透传
3. lint 升级设计建议引擎 — 新增字号层级混乱（同页标题<正文）/对比度过低
   （WCAG 3:1，含主题色解析）/图片拉伸变形（裁剪图与 <8px 占位图豁免）三类
   warn-only 检查；全部 8 类 issue 携带「设计师会怎么改」自然语言
   suggestion——对标 Kition designer 建议
4. 图片导入 OCR 重建可编辑 — `/api/docs/import-images` 默认本机 OCR
   （macOS Vision via pyobjc，zh-Hans 优先），文字行按行框生成原生文本框
   叠在图片上层（可改字），`ocr=0` 可关、不可用自动退回纯图片页——对标
   ningzimu 图片→可编辑
5. 演讲者模式 + 排练计时 — `GET /api/docs/{id}/present`：渲染 PNG 直放
   （保真 100%），←→/空格翻页、演讲者视图（当前页+下一页预览+备注+总/本页
   双计时器、暂停/重置）、全屏；前端「▶ 演示」按钮——对标 guizang 第 1 名
   核心体验
6. 原生图表创建 — `add_chart` 操作：python-pptx 原生 chart part
   （column/bar/bar_h/line/pie/area/doughnut + 中文别名，标题/图例/bbox，
   值数=类目数校验、饼图单系列），PowerPoint 可双击改数据；parse 输出
   kind=chart 对象供 edit_chart 续改——对标 ppt-master 数据图表
7. 母版/版式强化 — `GET /api/docs/{id}/layouts` 列出多母版版式清单（占位符
   类型+在用页数）；add_slide 缺省改用全稿最常用版式（品牌一致），支持
   layout_name 中英等价匹配——对标 ppt-master slide masters
8. 重主题预览安全网 — `POST /api/docs/{id}/theme-preview`：theme_rebrand/
   recolor 先在副本上应用并渲染 before/after 对比图（真文档一字节不动），
   确认后照常 /apply——对标 CasualOffice round-trip 大改动先看效果
9. 逐页并发渲染 — PDF 光栅化按连续页段分给多进程（默认 min(4,CPU)，
   `PPTMO_RASTER_WORKERS` 可调，失败自动回退串行），大文档预览提速——对标
   ningzimu page worker 并发
10. 交付健康报告 — `GET /api/docs/{id}/health`：文档级评分卡（lint 错误/
    警告、乱码符号残留、字体统一度 + 页数/对象统计，0-100 分 + A/B/C/D 档
    + 可解释扣分明细），apply 响应附带 health 摘要——对标 awesome 榜
    rework 量度量

## v0.5.0 — 对标 GitHub 顶级同类项目，10 项高价值功能 (2026-11-25)

对标 hands-on-deck（结构化编辑）/ OfficeCLI（render→look→fix）/ dashi-ppt-skill /
Kami，围绕「对话式精准改稿」定位取长补短：

1. `theme_rebrand` — 整套重主题：改 theme part 的 scheme colors + major/minor
   字体，一次操作整套 PPT 换品牌色（模板 deck 重品牌化的核心原语）
2. `recolor` — 调色板映射：`{mapping: {"FF0000":"0066CC"}}` 全稿扫描
   solidFill/渐变 stop 按映射替换（管硬编码颜色，与 theme_rebrand 互补）
3. `POST /api/docs/{id}/merge` — 从另一 deck 拉页：上传源 PPTX + 页码选择器 +
   position，跨 OPC 包克隆（slide part + 版式/母版/图片 rels 全量重新家园化）
4. `set_transition` — 页面切换动画：fade/push/wipe 等 21 种效果 + 方向 +
   速度 + 自动换片（手工组 `<p:transition>` XML）
5. `hide_slide` / `show_slide` — 隐藏页：`show="0"` 挂 sldId，播放/导出自动
   跳过，页码/对象结构不动（不同于删页）；parse 输出 hidden 标记
6. `edit_props` + `set_alt_text` — 文档元数据（标题/作者/主题/关键词）+
   对象替代文本（无障碍 + AI 理解图片内容）；轻量 `GET /objects` 清单
7. lint + `fix_issues` — 后置 lint：apply 后自动检查（文字溢出框外/字号超界/
   对象越画布/重叠告警）返回 issues；fix 确定性修复（缩字号/收回画布），
   重叠只告警；`GET /api/docs/{id}/lint` 按需检查
8. 结构 diff 增强 — 版本对比输出对象级结构化 changelog（新增/删除/移动了
   哪些对象、几何变化、样式变化），JSON 给 AI 消费
9. watch 实时预览 — `GET /api/docs/{id}/watch` 浏览器分屏实时看渲染图变化
   （SSE 推送版本变更事件，apply/undo/restore/merge 全覆盖）
10. 隐藏页 + 导出联动 — 导出 PDF/PNG zip/pdf-notes 默认跳过隐藏页
    （`?include_hidden=1` 可含）；outline 大纲标注「（已隐藏）」

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
