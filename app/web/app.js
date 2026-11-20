const $id = (id) => document.getElementById(id);

const dom = {
  landing: $id("landing"),
  workspace: $id("workspace"),
  docMeta: $id("docMeta"),
  docName: $id("docName"),
  docVersion: $id("docVersion"),
  backBtn: $id("backBtn"),
  uploadBtn: $id("uploadBtn"),
  landingUploadBtn: $id("landingUploadBtn"),
  historyBtn: $id("historyBtn"),
  downloadBtn: $id("downloadBtn"),
  fileInput: $id("fileInput"),
  dropzone: $id("dropzone"),
  recentSection: $id("recentSection"),
  recentList: $id("recentList"),
  refreshDocsBtn: $id("refreshDocsBtn"),
  slideCount: $id("slideCount"),
  thumbs: $id("thumbs"),
  prevBtn: $id("prevBtn"),
  nextBtn: $id("nextBtn"),
  slideLabel: $id("slideLabel"),
  addTextBtn: $id("addTextBtn"),
  stageScroll: $id("stageScroll"),
  stage: $id("stage"),
  slideImage: $id("slideImage"),
  overlay: $id("overlay"),
  stageLoader: $id("stageLoader"),
  statusText: $id("statusText"),
  statusRight: $id("statusRight"),
  aiBadge: $id("aiBadge"),
  sessionLogBtn: $id("sessionLogBtn"),
  chatList: $id("chatList"),
  chatInput: $id("chatInput"),
  chatHint: $id("chatHint"),
  sendBtn: $id("sendBtn"),
  stopBtn: $id("stopBtn"),
  autoApplyToggle: $id("autoApplyToggle"),
  chatPending: $id("chatPending"),
  sessionDialog: $id("sessionDialog"),
  sessionDialogMeta: $id("sessionDialogMeta"),
  messageList: $id("messageList"),
  closeSessionLogBtn: $id("closeSessionLogBtn"),
  tabAi: $id("tabAi"),
  tabText: $id("tabText"),
  aiPane: $id("aiPane"),
  textPane: $id("textPane"),
  scopeSegment: $id("scopeSegment"),
  promptChips: $id("promptChips"),
  aiNotice: $id("aiNotice"),
  patchCount: $id("patchCount"),
  acceptAllBtn: $id("acceptAllBtn"),
  clearPatchesBtn: $id("clearPatchesBtn"),
  patchList: $id("patchList"),
  applyBar: $id("applyBar"),
  acceptedCount: $id("acceptedCount"),
  applyBtn: $id("applyBtn"),
  noSelection: $id("noSelection"),
  inspector: $id("inspector"),
  inspectPage: $id("inspectPage"),
  inspectFont: $id("inspectFont"),
  inspectId: $id("inspectId"),
  textEditor: $id("textEditor"),
  saveTextBtn: $id("saveTextBtn"),
  deleteShapeBtn: $id("deleteShapeBtn"),
  fontSizeInput: $id("fontSizeInput"),
  colorInput: $id("colorInput"),
  alignSelect: $id("alignSelect"),
  boldToggle: $id("boldToggle"),
  italicToggle: $id("italicToggle"),
  applyStyleBtn: $id("applyStyleBtn"),
  inspectGeo: $id("inspectGeo"),
  busy: $id("busy"),
  busyTitle: $id("busyTitle"),
  busyText: $id("busyText"),
  toastStack: $id("toastStack"),
  historyDialog: $id("historyDialog"),
  closeHistoryBtn: $id("closeHistoryBtn"),
  versionList: $id("versionList"),
};

const state = {
  doc: null,
  slideIndex: 0,
  selectedId: null,
  patches: [],
  patchesStale: false,
  scope: "current",
  busy: false,
  streaming: false,
  streamController: null,
  live: null,
  inline: null,
  stageFitted: false,
  lastTrace: [],
};

const MIN_TEXT_EMU = 91440;

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[char]));
}

function formatDate(iso) {
  if (!iso) return "";
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("zh-CN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatSize(bytes) {
  const num = Number(bytes || 0);
  if (num < 1024) return `${num} B`;
  if (num < 1024 * 1024) return `${(num / 1024).toFixed(1)} KB`;
  return `${(num / 1024 / 1024).toFixed(1)} MB`;
}

function emuToCm(value) {
  return (Number(value || 0) / 360000).toFixed(2);
}

async function api(url, options = {}, timeout = 180000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    const raw = await response.text();
    let data = null;
    if (raw) {
      try {
        data = JSON.parse(raw);
      } catch {
        data = null;
      }
    }
    if (!response.ok) {
      const message = (data && data.detail) || raw || response.statusText || "请求失败";
      const error = new Error(typeof message === "string" ? message : JSON.stringify(message));
      error.status = response.status;
      throw error;
    }
    return data;
  } catch (error) {
    if (error.name === "AbortError") {
      const timeoutError = new Error("请求超时，请重试");
      timeoutError.status = 0;
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function setBusy(show, title = "处理中", text = "请稍候…") {
  state.busy = show;
  if (show) {
    dom.busyTitle.textContent = title;
    dom.busyText.textContent = text;
    dom.busy.classList.remove("is-hidden");
  } else {
    dom.busy.classList.add("is-hidden");
    renderApplyBar();
  }
}

function toast(message, type = "") {
  const node = document.createElement("div");
  node.className = `toast ${type}`.trim();
  node.textContent = message;
  dom.toastStack.appendChild(node);
  setTimeout(() => {
    node.style.opacity = "0";
    node.style.transform = "translateY(4px)";
    node.style.transition = "opacity .2s, transform .2s";
    setTimeout(() => node.remove(), 220);
  }, type === "error" ? 5600 : 3800);
}

function setStatus(text, right) {
  if (text !== undefined) dom.statusText.textContent = text;
  if (right !== undefined) dom.statusRight.textContent = right;
}

function documentObject(shapeId) {
  if (!state.doc || !shapeId) return null;
  return state.doc.objects.find((obj) => obj.shape_id === shapeId) || null;
}

function currentSlideObjects() {
  if (!state.doc) return [];
  return state.doc.objects.filter((obj) => obj.slide_index === state.slideIndex);
}

function imageUrl(slideIndex, { thumb = false, versionId } = {}) {
  if (!state.doc) return "";
  const version = versionId || state.doc.version_id;
  return `${state.doc.image_url_base}/${slideIndex}/image?${thumb ? "thumb=1&" : ""}v=${encodeURIComponent(version)}`;
}

function showLanding() {
  dom.landing.classList.remove("is-hidden");
  dom.workspace.classList.add("is-hidden");
  dom.docMeta.classList.add("is-hidden");
  dom.backBtn.classList.add("is-hidden");
  dom.uploadBtn.classList.add("is-hidden");
  dom.historyBtn.classList.add("is-hidden");
  dom.downloadBtn.classList.add("is-hidden");
  dom.historyBtn.disabled = true;
  dom.downloadBtn.disabled = true;
}

function showWorkspace() {
  dom.landing.classList.add("is-hidden");
  dom.workspace.classList.remove("is-hidden");
  dom.docMeta.classList.remove("is-hidden");
  dom.backBtn.classList.remove("is-hidden");
  dom.uploadBtn.classList.remove("is-hidden");
  dom.historyBtn.classList.remove("is-hidden");
  dom.downloadBtn.classList.remove("is-hidden");
  dom.historyBtn.disabled = false;
  dom.downloadBtn.disabled = false;
}

/* landing + recent documents */
async function loadRecent() {
  try {
    const data = await api("/api/sessions", {}, 20000);
    renderSessions(data.sessions || []);
  } catch (error) {
    dom.recentList.innerHTML = `<div class="session-empty">无法读取会话：${escapeHtml(error.message)}</div>`;
  }
}

function messagePreview(session) {
  const last = session.last_message;
  if (!last) return "尚未开始对话，打开后可以直接改稿。";
  const prefix = last.role === "user" ? "你：" : last.role === "assistant" ? "AI：" : "";
  return `${prefix}${last.content || ""}`;
}

function renderSessions(sessions) {
  if (!sessions.length) {
    dom.recentList.innerHTML = `
      <div class="session-empty">
        <strong>还没有会话</strong>
        <p>上传一份 PPTX，开始你的第一次修改。</p>
      </div>`;
    return;
  }
  dom.recentList.innerHTML = sessions
    .map((session) => {
      const thumb = `/api/docs/${session.doc_id}/slides/0/image?thumb=1&v=${encodeURIComponent(session.version_id || "v0")}`;
      const title = session.title || session.file_name || "未命名会话";
      return `
        <div class="session-row">
          <button class="session-open" type="button" data-open="${session.doc_id}" aria-label="继续编辑 ${escapeHtml(title)}">
            <img class="session-thumb" loading="lazy" src="${thumb}" alt="">
            <span class="session-content">
              <span class="session-top">
                <span class="session-title">${escapeHtml(title)}</span>
                <span class="session-version">${escapeHtml(session.version_id || "v0")}</span>
              </span>
              <span class="session-last">${escapeHtml(messagePreview(session))}</span>
              <span class="session-meta">${session.slide_count} 页 · ${session.message_count || 0} 条记录 · ${escapeHtml(formatDate(session.updated_at))}</span>
            </span>
            <span class="session-arrow"><svg viewBox="0 0 24 24"><path d="m9 5 7 7-7 7"/></svg></span>
          </button>
          <button class="session-del" type="button" data-delete="${session.doc_id}" data-title="${escapeHtml(title)}" title="删除会话" aria-label="删除会话">
            <svg viewBox="0 0 24 24"><path d="M4 7h16M9 7V4h6v3m-8 0 1 13h8l1-13"/></svg>
          </button>
        </div>`;
    })
    .join("") + '<div class="session-tip">每个会话都会保留对话、建议与版本，可随时回来继续。</div>';
}

async function deleteDocument(docId, title) {
  if (!window.confirm(`确定删除「${title}」？本地文件、会话记录和所有版本都会被移除。`)) return;
  try {
    await api(`/api/docs/${docId}`, { method: "DELETE" }, 30000);
    toast("会话已删除", "success");
    await loadRecent();
  } catch (error) {
    toast(error.message, "error");
  }
}

/* document lifecycle */
async function openDocument(docId) {
  if (state.streaming && state.streamController) state.streamController.abort();
  setBusy(true, "正在载入文档", "读取版本与渲染图…");
  try {
    const doc = await api(`/api/docs/${docId}`, {}, 60000);
    applyDocument(doc, { resetSlide: true, preserveSelection: false, clearPatches: true });
    showWorkspace();
    history.replaceState(null, "", `#doc=${docId}`);
    requestAnimationFrame(fitStage);
  } catch (error) {
    toast(error.message, "error");
    goHome();
  } finally {
    setBusy(false);
  }
}

function applyDocument(doc, options = {}) {
  const { resetSlide = false, preserveSelection = true, clearPatches = true } = options;
  const sameDoc = state.doc && state.doc.doc_id === doc.doc_id;
  const previousSlide = state.slideIndex;
  const previousSelected = state.selectedId;

  state.doc = doc;
  state.slideIndex = resetSlide || !sameDoc
    ? 0
    : clamp(previousSlide, 0, Math.max(0, doc.slide_count - 1));
  const selectedStillExists = preserveSelection && previousSelected
    && doc.objects.some((obj) => obj.shape_id === previousSelected && obj.slide_index === state.slideIndex);
  state.selectedId = selectedStillExists ? previousSelected : null;
  state.stageFitted = false;
  if (clearPatches) {
    state.patches = [];
    state.patchesStale = false;
    setNotice("");
  }
  const lastTrace = [...(doc.messages || [])]
    .reverse()
    .find((message) => Array.isArray(message.trace))?.trace;
  state.lastTrace = lastTrace || [];
  renderChrome();
  renderThumbs();
  renderStage();
  renderPatchList();
  renderInspector();
  renderChat();
  renderMessages();
  if (clearPatches) restorePendingOperations();
  updateThumbsActive();
  updateStatus();
  requestAnimationFrame(fitStage);
}

function restorePendingOperations() {
  if (state.streaming || state.patches.length || !state.doc) return;
  // server-computed pending ops survive message trimming — prefer them
  const pending = state.doc.pending_operations;
  if (pending && Array.isArray(pending.operations) && pending.operations.length) {
    state.patches = pending.operations.map((operation) => ({
      ...operation,
      uiState: "accepted",
      message_id: pending.message_id || null,
    }));
    const baseVersion = Number.isInteger(pending.base_version)
      ? pending.base_version
      : state.doc.version;
    state.patchesStale = baseVersion !== state.doc.version;
    renderPatchList();
    if (state.patchesStale) setNotice("待确认操作基于旧版本，请重新生成后再应用。", "error");
    return;
  }
  if (pending) return; // server says nothing pending — trust it
  const messages = state.doc.messages || [];
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message.role !== "assistant") continue;
    const operations = Array.isArray(message.operations) ? message.operations : [];
    if (!operations.length) continue;
    if (message.applied_version) return;
    if (messages.slice(index + 1).some((later) => later.kind === "apply")) return;
    if (messages.length - 1 - index > 4) return;
    // ops already applied through a partial apply must not be offered again
    const appliedIds = new Set(
      (Array.isArray(message.partial_applies) ? message.partial_applies : [])
        .flatMap((apply) => (Array.isArray(apply.applied_shape_ids) ? apply.applied_shape_ids : []))
    );
    const remaining = operations.filter(
      (operation) => !operation.shape_id || !appliedIds.has(String(operation.shape_id))
    );
    if (!remaining.length) return;
    state.patches = remaining.map((operation) => ({
      ...operation,
      uiState: "accepted",
      message_id: message.id || null,
    }));
    const baseVersion = Number.isInteger(message.base_version)
      ? message.base_version
      : Number.isInteger(message.version)
        ? message.version
        : state.doc.version;
    state.patchesStale = baseVersion !== state.doc.version;
    renderPatchList();
    if (state.patchesStale) setNotice("待确认操作基于旧版本，请重新生成后再应用。", "error");
    return;
  }
}

function goHome() {
  if (state.streaming && state.streamController) state.streamController.abort();
  if (state.inline) closeInline(false);
  state.doc = null;
  state.selectedId = null;
  state.patches = [];
  state.patchesStale = false;
  state.lastTrace = [];
  if (dom.sessionDialog.open) dom.sessionDialog.close();
  history.replaceState(null, "", location.pathname + location.search);
  dom.slideImage.removeAttribute("src");
  dom.slideImage.dataset.url = "";
  dom.overlay.innerHTML = "";
  dom.recentList.innerHTML = '<div class="session-empty">正在读取会话…</div>';
  showLanding();
  loadRecent();
}

function renderChrome() {
  if (!state.doc) return;
  dom.docName.textContent = state.doc.title || state.doc.file_name;
  dom.docName.title = state.doc.file_name || "";
  dom.docVersion.textContent = state.doc.version_id;
  dom.slideCount.textContent = String(state.doc.slide_count);
  dom.aiBadge.textContent = state.doc.ai_configured ? (state.doc.ai_model || "AI") : "AI 未配置";
  dom.aiBadge.classList.toggle("warn", !state.doc.ai_configured);
  dom.slideLabel.textContent = `${state.slideIndex + 1} / ${state.doc.slide_count}`;
  dom.prevBtn.disabled = state.slideIndex <= 0;
  dom.nextBtn.disabled = state.slideIndex >= state.doc.slide_count - 1;
  updateScanButtonLabel();
}

function updateStatus() {
  if (!state.doc) return;
  const count = state.doc.objects.filter((obj) => obj.slide_index === state.slideIndex).length;
  setStatus(
    `已保存 ${state.doc.version_id} · 所有修改直接写入原 PPTX`,
    `第 ${state.slideIndex + 1} 页 · ${count} 个文本 · 共 ${state.doc.objects.length} 个`
  );
}

function renderThumbs() {
  if (!state.doc) {
    dom.thumbs.innerHTML = "";
    return;
  }
  const existing = dom.thumbs.querySelectorAll(".thumb");
  if (existing.length === state.doc.slide_count) {
    // same deck shape: patch image sources only — rebuilding the DOM would
    // re-download every thumbnail and re-attach listeners on each save
    existing.forEach((button) => {
      const index = Number(button.dataset.slide);
      const img = button.querySelector("img");
      if (!img) return;
      const url = imageUrl(index, { thumb: true });
      if (img.dataset.url !== url) {
        img.src = url;
        img.dataset.url = url;
      }
    });
    return;
  }
  const parts = [];
  for (let index = 0; index < state.doc.slide_count; index += 1) {
    const url = imageUrl(index, { thumb: true });
    parts.push(`
      <button class="thumb" type="button" data-slide="${index}" aria-label="第 ${index + 1} 页">
        <span class="thumb-frame"><img loading="lazy" src="${url}" data-url="${url}" alt=""></span>
        <span class="thumb-num">${index + 1}</span>
      </button>`);
  }
  dom.thumbs.innerHTML = parts.join("");
  dom.thumbs.querySelectorAll(".thumb").forEach((button) => {
    button.addEventListener("click", () => setSlide(Number(button.dataset.slide)));
  });
}

function updateThumbsActive() {
  dom.thumbs.querySelectorAll(".thumb").forEach((button) => {
    const active = Number(button.dataset.slide) === state.slideIndex;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-current", active ? "true" : "false");
  });
  const active = dom.thumbs.querySelector(".thumb.is-active");
  if (active) active.scrollIntoView({ block: "nearest" });
}

function setSlide(index, options = {}) {
  if (!state.doc) return;
  const next = clamp(Number(index), 0, state.doc.slide_count - 1);
  if (next === state.slideIndex && !options.force) {
    updateThumbsActive();
    return;
  }
  if (state.inline) closeInline(false);
  state.slideIndex = next;
  if (!options.keepSelection) state.selectedId = null;
  dom.slideLabel.textContent = `${state.slideIndex + 1} / ${state.doc.slide_count}`;
  dom.prevBtn.disabled = state.slideIndex <= 0;
  dom.nextBtn.disabled = state.slideIndex >= state.doc.slide_count - 1;
  updateThumbsActive();
  renderStage();
  renderInspector();
  updateStatus();
  updateScanButtonLabel();
}

function renderStage() {
  if (!state.doc) return;
  const url = imageUrl(state.slideIndex);
  if (dom.slideImage.dataset.url !== url) {
    dom.slideImage.dataset.url = url;
    dom.stageLoader.classList.remove("is-hidden");
    dom.slideImage.src = url;
  }
  renderOverlay();
  fitStage();
}

function renderOverlay() {
  dom.overlay.innerHTML = "";
  if (!state.doc) return;
  const size = state.doc.slide_size_emu || { width: 1, height: 1 };
  const width = Math.max(1, Number(size.width));
  const height = Math.max(1, Number(size.height));
  for (const obj of currentSlideObjects()) {
    const bbox = obj.bbox_emu || {};
    const box = document.createElement("div");
    box.className = "hit";
    if (obj.kind) box.classList.add(`hit-${obj.kind}`);
    if (obj.shape_id === state.selectedId) box.classList.add("is-selected");
    box.dataset.shapeId = obj.shape_id;
    box.style.left = `${(Number(bbox.x || 0) / width) * 100}%`;
    box.style.top = `${(Number(bbox.y || 0) / height) * 100}%`;
    box.style.width = `${(Number(bbox.w || 1) / width) * 100}%`;
    box.style.height = `${(Number(bbox.h || 1) / height) * 100}%`;
    box.title = obj.kind === "table_cell"
      ? `${obj.text || ""} · R${Number(obj.row) + 1}C${Number(obj.col) + 1}`
      : (obj.text || "");
    const tag = document.createElement("span");
    tag.className = "hit-tag";
    tag.textContent = obj.font_size_pt
      ? `${Math.round(obj.font_size_pt)}pt`
      : `≈${Math.round(estimatedFontSize(obj))}pt`;
    box.append(tag);
    if (obj.resizable !== false) {
      const handle = document.createElement("span");
      handle.className = "hit-handle";
      handle.title = "拖动缩放";
      box.append(handle);
      handle.addEventListener("pointerdown", (event) => {
        event.stopPropagation();
        startBoxGesture(event, obj, box, "resize");
      });
    }
    if (obj.movable !== false) {
      box.addEventListener("pointerdown", (event) => startBoxGesture(event, obj, box, "move"));
    } else {
      box.addEventListener("pointerdown", (event) => {
        event.stopPropagation();
        setSelected(obj.shape_id, { switchToText: true });
      });
    }
    box.addEventListener("dblclick", (event) => {
      event.stopPropagation();
      event.preventDefault();
      openInlineEditor(obj, { draft: false });
    });
    dom.overlay.appendChild(box);
  }
}

function fitStage() {
  if (!state.doc || dom.workspace.classList.contains("is-hidden")) return;
  const size = state.doc.slide_size_emu || { width: 16, height: 9 };
  const ratio = Math.max(0.2, Number(size.width) / Math.max(1, Number(size.height)));
  const scrollRect = dom.stageScroll.getBoundingClientRect();
  const availableWidth = Math.max(280, scrollRect.width - 44);
  const availableHeight = Math.max(220, scrollRect.height - 44);
  const width = Math.min(availableWidth, availableHeight * ratio, 1440);
  dom.stage.style.width = `${Math.round(width)}px`;
}

function fitStageSoon() {
  requestAnimationFrame(fitStage);
}

function setSelected(shapeId, options = {}) {
  const { switchToText = true, flash = false } = options;
  state.selectedId = shapeId;
  dom.overlay.querySelectorAll(".hit").forEach((box) => {
    box.classList.toggle("is-selected", box.dataset.shapeId === shapeId);
  });
  if (switchToText && shapeId) switchTab("text");
  renderInspector();
  if (flash && shapeId) {
    const box = dom.overlay.querySelector(`.hit[data-shape-id="${shapeId}"]`);
    if (box) {
      box.classList.remove("is-flash");
      void box.offsetWidth;
      box.classList.add("is-flash");
      setTimeout(() => box.classList.remove("is-flash"), 1300);
    }
  }
}

function gotoObject(shapeId, options = {}) {
  const obj = documentObject(shapeId);
  if (!obj) {
    toast("这条建议对应的文本框已不存在，请重新生成建议", "error");
    return;
  }
  if (obj.slide_index !== state.slideIndex) {
    setSlide(obj.slide_index, { keepSelection: true });
  }
  setSelected(shapeId, { switchToText: false, flash: options.flash !== false });
}

function switchTab(tab) {
  const ai = tab === "ai";
  dom.tabAi.classList.toggle("is-active", ai);
  dom.tabText.classList.toggle("is-active", !ai);
  dom.aiPane.classList.toggle("is-hidden", !ai);
  dom.textPane.classList.toggle("is-hidden", ai);
  if (!ai) renderInspector();
}

function setNotice(text, type = "") {
  dom.aiNotice.className = `notice ${type}`.trim();
  dom.aiNotice.textContent = text || "";
  dom.aiNotice.classList.toggle("is-hidden", !text);
}

function updateScanButtonLabel() {
  if (!dom.chatHint) return;
  const suffix = state.scope === "current" ? `第 ${state.slideIndex + 1} 页` : "整份文稿";
  dom.chatHint.textContent = `范围：${suffix} · Enter 发送 · Shift+Enter 换行`;
  if (dom.chatInput) {
    dom.chatInput.placeholder = state.scope === "current"
      ? `描述第 ${state.slideIndex + 1} 页要怎么改，例如：把这段压成 3 个要点，其他页不动。`
      : "描述整份文稿要怎么改，例如：先检查全篇，统一术语和数字格式，再给出高价值修改。";
  }
}

function renderInspector() {
  const obj = documentObject(state.selectedId);
  if (!obj) {
    dom.noSelection.classList.remove("is-hidden");
    dom.inspector.classList.add("is-hidden");
    return;
  }
  dom.noSelection.classList.add("is-hidden");
  dom.inspector.classList.remove("is-hidden");
  const tableMeta = obj.kind === "table_cell" ? ` · R${Number(obj.row) + 1}C${Number(obj.col) + 1}` : "";
  dom.inspectPage.textContent = `第 ${obj.slide_index + 1} / ${state.doc.slide_count} 页${tableMeta}`;
  dom.inspectFont.textContent = `${obj.kind === "table_cell" ? "表格单元格" : "文本框"} · ${obj.font_name || "继承字体"} · ${obj.font_size_pt ? `${Math.round(obj.font_size_pt)}pt` : `≈${Math.round(estimatedFontSize(obj))}pt（继承）`}`;
  dom.inspectId.textContent = obj.shape_id;
  if (document.activeElement !== dom.textEditor || dom.textEditor.dataset.shapeId !== obj.shape_id) {
    dom.textEditor.value = obj.text || "";
  }
  dom.textEditor.dataset.shapeId = obj.shape_id;
  const bbox = obj.bbox_emu || {};
  dom.inspectGeo.innerHTML = `
    <div><span>X</span><strong>${emuToCm(bbox.x)} cm</strong></div>
    <div><span>Y</span><strong>${emuToCm(bbox.y)} cm</strong></div>
    <div><span>宽</span><strong>${emuToCm(bbox.w)} cm</strong></div>
    <div><span>高</span><strong>${emuToCm(bbox.h)} cm</strong></div>`;
  dom.fontSizeInput.value = Number(effectiveFontSize(obj)).toFixed(1);
  dom.boldToggle.checked = Boolean(obj.bold);
  dom.italicToggle.checked = Boolean(obj.italic);
  dom.alignSelect.value = String(effectiveAlignment(obj) || "LEFT").toUpperCase().replace("PP_ALIGN.", "");
  const color = String(obj.color || "").replace("#", "");
  dom.colorInput.value = /^[0-9a-fA-F]{6}$/.test(color) ? `#${color}` : "#1f2937";
  dom.deleteShapeBtn.textContent = obj.kind === "table_cell" ? "清空单元格" : "删除文本框";
  dom.deleteShapeBtn.disabled = obj.deletable === false;
}

async function saveSelectedText() {
  const obj = documentObject(state.selectedId);
  if (!obj) return;
  const value = dom.textEditor.value;
  if (!value.trim()) {
    if (obj.kind === "table_cell") {
      await runApply(
        [{
          action: "clear",
          slide_index: obj.slide_index,
          shape_id: obj.shape_id,
          before: obj.text,
          description: "清空表格单元格",
        }],
        "清空单元格",
        { successMessage: "单元格已清空" }
      );
      return;
    }
    toast("文本不能为空。如需移除，请使用「删除文本框」。", "error");
    dom.textEditor.focus();
    return;
  }
  if (value === obj.text) {
    toast("文本没有变化");
    return;
  }
  await runApply(
    [{
      action: "edit",
      slide_index: obj.slide_index,
      shape_id: obj.shape_id,
      before: obj.text,
      after: value,
      description: obj.kind === "table_cell" ? "手动修改表格单元格" : "手动修改文本",
    }],
    obj.kind === "table_cell" ? "修改单元格" : "手动改字",
    { successMessage: "文本已写入原 PPTX" }
  );
}

async function deleteSelectedShape() {
  const obj = documentObject(state.selectedId);
  if (!obj) return;
  if (obj.kind === "table_cell") {
    if (!window.confirm("清空这个单元格的文字？该操作会生成新版本，可在历史中恢复。")) return;
    await runApply(
      [{
        action: "clear",
        slide_index: obj.slide_index,
        shape_id: obj.shape_id,
        before: obj.text,
        description: "清空表格单元格",
      }],
      "清空单元格",
      { successMessage: "单元格已清空" }
    );
    return;
  }
  if (obj.deletable === false) {
    toast("分组内的文字不能删除，只能修改", "error");
    return;
  }
  if (!window.confirm("确定删除这个文本框？该操作会生成新版本，可在历史中恢复。")) return;
  await runApply(
    [{
      action: "delete_textbox",
      slide_index: obj.slide_index,
      shape_id: obj.shape_id,
      before: obj.text,
      after: "",
      description: "删除文本框",
    }],
    "删除文本框",
    { successMessage: "文本框已删除" }
  );
}

async function applySelectedStyle() {
  const obj = documentObject(state.selectedId);
  if (!obj) return;
  const fontValue = Number(dom.fontSizeInput.value);
  const style = {
    bold: dom.boldToggle.checked,
    italic: dom.italicToggle.checked,
    alignment: dom.alignSelect.value,
    color: dom.colorInput.value.replace("#", "").toUpperCase(),
  };
  if (Number.isFinite(fontValue) && fontValue >= 6) style.font_size_pt = fontValue;
  await runApply(
    [{
      action: "style",
      slide_index: obj.slide_index,
      shape_id: obj.shape_id,
      style,
      description: "精确调整文字样式",
    }],
    "调整样式",
    { successMessage: "样式已写入原 PPTX" }
  );
}

/* inline editor */
function draftBbox() {
  const size = state.doc.slide_size_emu || { width: 9144000, height: 5143500 };
  const width = Math.round(size.width * 0.8);
  const height = Math.round(size.height * 0.11);
  return {
    x: Math.round(size.width * 0.1),
    y: Math.round(size.height * 0.62),
    w: width,
    h: height,
  };
}

function estimatedFontSize(obj) {
  const size = state.doc?.slide_size_emu || { width: 1, height: 1 };
  const bbox = obj.bbox_emu || {};
  const heightPt = Number(bbox.h || 0) / 12700;
  const lines = Math.max(1, String(obj.text || "").split("\n").length);
  const topRatio = Number(bbox.y || 0) / Math.max(1, Number(size.height));
  if (topRatio < 0.2 && heightPt > 80) return clamp(heightPt * 0.3, 24, 54);
  return clamp((heightPt / lines) * 0.55, 10, 40);
}

function effectiveFontSize(obj) {
  return Number(obj.font_size_pt) || estimatedFontSize(obj);
}

function effectiveAlignment(obj) {
  if (obj.alignment) return obj.alignment;
  const size = state.doc?.slide_size_emu || { width: 1, height: 1 };
  const bbox = obj.bbox_emu || {};
  const left = Number(bbox.x || 0);
  const right = Number(size.width) - left - Number(bbox.w || 0);
  if (Math.abs(left - right) < Number(size.width) * 0.08) return "CENTER";
  return "LEFT";
}

function draftFontSize() {
  const body = currentSlideObjects()
    .filter((obj) => (obj.text || "").length > 8)
    .sort((a, b) => effectiveFontSize(b) - effectiveFontSize(a))[0];
  return body ? effectiveFontSize(body) : 20;
}

function closeInline(commit = false) {
  if (!state.inline) return;
  const session = state.inline;
  // Escape calls finish(false) directly (explicit cancel); this path is
  // navigation-away, where silently dropping typed text is a data loss
  if (!commit && !session.finished) {
    const changed = session.draft
      ? Boolean(session.textarea.value.trim())
      : session.textarea.value !== (session.obj?.text || "");
    if (changed && !window.confirm("有未保存的修改，确定放弃吗？")) return;
  }
  if (commit) session.finish(true);
  else session.finish(false);
}

function openInlineEditor(obj, options = {}) {
  const draft = Boolean(options.draft);
  if (!state.doc || state.busy) return;
  if (state.inline) closeInline(false);

  const stageWidth = dom.slideImage.clientWidth || dom.stage.clientWidth;
  const stageHeight = dom.slideImage.clientHeight || dom.stage.clientHeight;
  if (!stageWidth || !stageHeight) {
    toast("预览图还未加载完成，请稍后再试", "error");
    return;
  }
  const size = state.doc.slide_size_emu || { width: 9144000, height: 5143500 };
  const bbox = draft ? draftBbox() : obj.bbox_emu;
  const pxPerEmuX = stageWidth / Math.max(1, size.width);
  const pxPerEmuY = stageHeight / Math.max(1, size.height);
  const fontPt = draft ? draftFontSize() : effectiveFontSize(obj);
  const fontPx = clamp(fontPt * (stageHeight / (Number(size.height) / 12700)), 8, 120);
  const left = clamp(Number(bbox.x) * pxPerEmuX, 0, stageWidth - 60);
  const top = clamp(Number(bbox.y) * pxPerEmuY, 0, stageHeight - 28);
  const width = clamp(Number(bbox.w) * pxPerEmuX, 80, stageWidth - left);
  const height = clamp(Math.max(Number(bbox.h) * pxPerEmuY, fontPx * 1.45), 28, stageHeight - top);

  const textarea = document.createElement("textarea");
  textarea.className = "inline-editor";
  textarea.value = draft ? "" : (obj.text || "");
  textarea.placeholder = draft ? "输入新文本框内容，Ctrl+Enter 保存" : "";
  textarea.spellcheck = false;
  textarea.style.left = `${Math.round(left)}px`;
  textarea.style.top = `${Math.round(top)}px`;
  textarea.style.width = `${Math.round(width)}px`;
  textarea.style.height = `${Math.round(height)}px`;
  textarea.style.fontSize = `${fontPx.toFixed(1)}px`;
  textarea.style.textAlign = alignmentToCss(effectiveAlignment(obj));
  textarea.style.fontWeight = obj.bold ? "700" : "400";
  textarea.style.fontStyle = obj.italic ? "italic" : "normal";

  const session = { obj, textarea, draft, finished: false };
  session.finish = async (commit) => {
    if (session.finished) return;
    session.finished = true;
    const value = textarea.value;
    textarea.remove();
    if (state.inline === session) state.inline = null;
    if (!commit) return;
    if (draft && !value.trim()) return;
    if (!draft && value === (obj.text || "")) return;
    if (!draft && !value.trim() && obj.kind !== "table_cell") {
      toast("文本不能为空。如需移除，请使用「删除文本框」。", "error");
      return;
    }

    let patch;
    if (draft) {
      patch = {
        action: "add_textbox",
        slide_index: state.slideIndex,
        after: value,
        description: "新增文本框",
        payload: {
          bbox_emu: bbox,
          text: value,
          font_size_pt: fontPt,
          alignment: "LEFT",
        },
      };
    } else if (obj.kind === "table_cell" && !value.trim()) {
      patch = {
        action: "clear",
        slide_index: obj.slide_index,
        shape_id: obj.shape_id,
        before: obj.text,
        description: "清空表格单元格",
      };
    } else {
      patch = {
        action: "edit",
        slide_index: obj.slide_index,
        shape_id: obj.shape_id,
        before: obj.text,
        after: value,
        description: obj.kind === "table_cell" ? "原位修改表格单元格" : "原位修改文本",
      };
    }
    const slideIndex = state.slideIndex;
    const previousIds = new Set(state.doc.objects.map((item) => item.shape_id));
    const doc = await runApply(
      [patch],
      draft ? "新增文本框" : (obj.kind === "table_cell" ? "修改单元格" : "原位改字"),
      { successMessage: draft ? "文本框已添加" : "文本已写入原 PPTX" }
    );
    if (doc && draft) {
      const added = doc.objects
        .filter((item) => item.slide_index === slideIndex && !previousIds.has(item.shape_id))
        .sort((a, b) => Number(b.shape_index || 0) - Number(a.shape_index || 0))[0];
      if (added) {
        state.selectedId = added.shape_id;
        setSelected(added.shape_id, { switchToText: true, flash: true });
        renderInspector();
      }
    }
  };

  textarea.addEventListener("blur", () => { void session.finish(true); });
  textarea.addEventListener("pointerdown", (event) => event.stopPropagation());
  textarea.addEventListener("keydown", (event) => {
    event.stopPropagation();
    if (event.key === "Escape") {
      event.preventDefault();
      void session.finish(false);
    } else if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
      event.preventDefault();
      void session.finish(true);
    }
  });

  state.inline = session;
  dom.overlay.appendChild(textarea);
  textarea.focus();
  if (!draft) textarea.select();
}

function alignmentToCss(alignment) {
  const value = String(alignment || "").toUpperCase();
  if (value.includes("CENTER")) return "center";
  if (value.includes("RIGHT")) return "right";
  if (value.includes("JUSTIFY")) return "justify";
  return "left";
}

function startDraftTextbox() {
  if (!state.doc || state.busy) return;
  if (state.inline) closeInline(false);
  switchTab("text");
  openInlineEditor(
    {
      shape_id: "__draft__",
      slide_index: state.slideIndex,
      text: "",
      bbox_emu: draftBbox(),
      font_size_pt: draftFontSize(),
      alignment: "LEFT",
    },
    { draft: true }
  );
}

/* drag / resize */
function startBoxGesture(event, obj, box, mode) {
  if (event.button !== undefined && event.button !== 0) return;
  if (state.busy) return;
  event.preventDefault();
  event.stopPropagation();
  if (state.inline) closeInline(false);
  setSelected(obj.shape_id, { switchToText: true });

  const stageWidth = dom.slideImage.clientWidth || dom.stage.clientWidth;
  const stageHeight = dom.slideImage.clientHeight || dom.stage.clientHeight;
  if (!stageWidth || !stageHeight) return;
  const size = state.doc.slide_size_emu || { width: 9144000, height: 5143500 };
  const bbox = obj.bbox_emu || {};
  const orig = {
    left: Number(bbox.x) / size.width * stageWidth,
    top: Number(bbox.y) / size.height * stageHeight,
    width: Number(bbox.w) / size.width * stageWidth,
    height: Number(bbox.h) / size.height * stageHeight,
  };
  const startX = event.clientX;
  const startY = event.clientY;
  const current = { ...orig };
  let moved = false;
  const previousUserSelect = document.body.style.userSelect;
  document.body.style.userSelect = "none";
  box.style.cursor = mode === "resize" ? "nwse-resize" : "grabbing";

  const onMove = (moveEvent) => {
    const dx = moveEvent.clientX - startX;
    const dy = moveEvent.clientY - startY;
    if (Math.abs(dx) + Math.abs(dy) > 6) moved = true;
    if (mode === "move") {
      current.left = clamp(orig.left + dx, 0, stageWidth - orig.width);
      current.top = clamp(orig.top + dy, 0, stageHeight - orig.height);
      box.style.left = `${current.left}px`;
      box.style.top = `${current.top}px`;
    } else {
      // clamp the on-screen box to the same min size the commit enforces, so
      // the re-rendered overlay doesn't jump bigger than what was positioned
      const minPxW = Math.max(6, MIN_TEXT_EMU * stageWidth / Math.max(1, size.width));
      const minPxH = Math.max(6, MIN_TEXT_EMU * stageHeight / Math.max(1, size.height));
      current.width = clamp(orig.width + dx, minPxW, stageWidth - orig.left);
      current.height = clamp(orig.height + dy, minPxH, stageHeight - orig.top);
      box.style.width = `${current.width}px`;
      box.style.height = `${current.height}px`;
    }
  };

  const onUp = () => {
    window.removeEventListener("pointermove", onMove);
    window.removeEventListener("pointerup", onUp);
    window.removeEventListener("pointercancel", onUp);
    document.body.style.userSelect = previousUserSelect;
    box.style.cursor = "";
    if (!moved) return;
    const pxPerEmuX = Math.max(1, size.width) / stageWidth;
    const pxPerEmuY = Math.max(1, size.height) / stageHeight;
    const nextBbox = {
      x: Math.round(current.left * pxPerEmuX),
      y: Math.round(current.top * pxPerEmuY),
      w: Math.max(MIN_TEXT_EMU, Math.round(current.width * pxPerEmuX)),
      h: Math.max(MIN_TEXT_EMU, Math.round(current.height * pxPerEmuY)),
    };
    // ~20000 EMU (≈0.055 cm): stray 1px jitter must not trigger a full
    // LibreOffice re-render of every page
    const changed = ["x", "y", "w", "h"].some(
      (key) => Math.abs(Number(nextBbox[key]) - Number(bbox[key] || 0)) > 20000
    );
    if (!changed) return;
    const patch = {
      action: "edit",
      slide_index: obj.slide_index,
      shape_id: obj.shape_id,
      before: obj.text,
      description: mode === "move" ? "调整文本框位置" : "调整文本框大小",
      payload: { bbox_emu: nextBbox },
    };
    void runApply([patch], mode === "move" ? "移动文本框" : "缩放文本框", {
      successMessage: mode === "move" ? "位置已保存" : "尺寸已保存",
    }).then((doc) => {
      if (!doc) {
        renderOverlay();
        renderInspector();
      }
    });
  };

  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
  window.addEventListener("pointercancel", onUp);
}

/* server write */
async function runApply(patches, label, options = {}) {
  if (!state.doc || state.busy) return null;
  const { successMessage, preserveSelection = true, assistantMessageId } = options;
  setBusy(
    true,
    "正在回写 PPTX",
    `应用 ${patches.length} 条修改，LibreOffice 正在重新渲染每一页…`
  );
  setStatus("正在保存并重新渲染…");
  try {
    const doc = await api(
      `/api/docs/${state.doc.doc_id}/apply`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          patches,
          label,
          base_version: state.doc.version,
          ...(assistantMessageId ? { assistant_message_id: assistantMessageId } : {}),
        }),
      },
      300000
    );
    applyDocument(doc, { preserveSelection, clearPatches: true });
    toast(successMessage || `已保存为 ${doc.version_id}`, "success");
    setStatus(`已保存 ${doc.version_id} · 所有修改直接写入原 PPTX`);
    return doc;
  } catch (error) {
    let recovered = false;
    if (error.status === 409) {
      try {
        const fresh = await api(`/api/docs/${state.doc.doc_id}`, {}, 60000);
        applyDocument(fresh, { preserveSelection: true, clearPatches: true });
        toast("文档已在别处更新，已加载最新版本", "error");
        recovered = true;
      } catch {
        /* keep original error */
      }
    }
    if (!recovered) toast(error.message, "error");
    return null;
  } finally {
    setBusy(false);
  }
}

/* AI suggestions */
function diffPair(before, after) {
  const a = String(before ?? "");
  const b = String(after ?? "");
  if (a === b) return { before: escapeHtml(a), after: escapeHtml(b) };
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start]) start += 1;
  let endA = a.length;
  let endB = b.length;
  while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) {
    endA -= 1;
    endB -= 1;
  }
  const context = 84;
  const changeLimit = 280;
  const trim = (text) => text.length > changeLimit ? `${text.slice(0, changeLimit)}…` : text;
  const prefix = start > context ? `…${a.slice(start - context, start)}` : a.slice(0, start);
  const suffix = a.length - endA > context ? `${a.slice(endA, endA + context)}…` : a.slice(endA);
  const removed = trim(a.slice(start, endA));
  const inserted = trim(b.slice(start, endB));
  return {
    before: `${escapeHtml(prefix)}${removed ? `<span class="del">${escapeHtml(removed)}</span>` : ""}${escapeHtml(suffix)}`,
    after: `${escapeHtml(prefix)}${inserted ? `<span class="ins">${escapeHtml(inserted)}</span>` : ""}${escapeHtml(suffix)}`,
  };
}

function patchStateLabel(patch) {
  if (patch.uiState === "accepted") return "已接受";
  if (patch.uiState === "ignored") return "已忽略";
  return "待审阅";
}

function styleSummary(style) {
  if (!style || typeof style !== "object") return "";
  const parts = [];
  if (style.font_size_pt) parts.push(`${Number(style.font_size_pt)}pt`);
  if (style.bold) parts.push("加粗");
  if (style.italic) parts.push("斜体");
  if (style.color) parts.push(`#${style.color}`);
  if (style.alignment) parts.push(String(style.alignment).toLowerCase());
  return parts.join(" · ");
}

function renderPatchList() {
  dom.patchCount.textContent = String(state.patches.length);
  if (!state.patches.length) {
    dom.patchList.innerHTML = `
      <div class="empty-soft">
        <span class="empty-mark">✦</span>
        <p>暂无可应用的操作。</p>
      </div>`;
  } else {
    dom.patchList.innerHTML = state.patches.map((patch) => {
      const accepted = patch.uiState === "accepted";
      const ignored = patch.uiState === "ignored";
      const classes = ["patch-card"];
      if (accepted) classes.push("is-accepted");
      if (ignored) classes.push("is-ignored");
      if (state.patchesStale) classes.push("is-stale");
      const action = String(patch.action || "edit");
      let body = "";
      if (action === "style") {
        body = `<div class="patch-text after">样式：${escapeHtml(styleSummary(patch.style) || "未指定")}</div>
          <div class="patch-text before">${escapeHtml(patch.before || "（无原文）")}</div>`;
      } else if (action === "geometry") {
        const bbox = patch.bbox_emu || {};
        const loc = [`x ${emuToCm(bbox.x)} cm`, `y ${emuToCm(bbox.y)} cm`];
        if (bbox.w) loc.push(`宽 ${emuToCm(bbox.w)} cm`);
        if (bbox.h) loc.push(`高 ${emuToCm(bbox.h)} cm`);
        body = `<div class="patch-text after">精确调整：${escapeHtml(loc.join(" · "))}</div>`;
      } else if (action === "clear") {
        body = `<div class="patch-text before">${escapeHtml(patch.before || "（无原文）")}</div>
          <div class="patch-text after">（清空此文字）</div>`;
      } else if (action === "add_textbox") {
        const bbox = patch.bbox_emu || {};
        body = `<div class="patch-text after">新增文本框：${escapeHtml(patch.after || "")}</div>
          <div class="patch-text before">位置：x ${emuToCm(bbox.x)} cm · y ${emuToCm(bbox.y)} cm · ${emuToCm(bbox.w)} × ${emuToCm(bbox.h)} cm</div>`;
      } else {
        const diff = diffPair(patch.before, patch.after);
        const afterHtml = action === "delete_textbox" ? "（将删除此文本框）" : diff.after;
        body = `<div class="patch-text before">${diff.before || "（原文为空）"}</div>
          <div class="patch-text after">${afterHtml}</div>`;
      }
      return `
        <article class="${classes.join(" ")}" data-id="${escapeHtml(patch.id)}">
          <div class="patch-top">
            <button class="patch-loc" type="button" data-action="locate" title="定位到对象">第${Number(patch.slide_index) + 1}页</button>
            <div class="patch-desc" title="${escapeHtml(patch.description || "")}">${escapeHtml(patch.description || "AI 操作")}</div>
            <span class="patch-state">${patchStateLabel(patch)}</span>
          </div>
          ${body}
          <div class="patch-actions">
            <button class="btn tiny ghost" type="button" data-action="locate">定位</button>
            <button class="btn tiny ghost" type="button" data-action="ignore">${ignored ? "恢复" : "忽略"}</button>
            <button class="btn tiny ${accepted ? "ghost" : "primary"}" type="button" data-action="accept">${accepted ? "取消接受" : "接受"}</button>
          </div>
        </article>`;
    }).join("");
  }
  const allAccepted = state.patches.length > 0
    && state.patches.every((patch) => patch.uiState === "accepted");
  dom.acceptAllBtn.textContent = allAccepted ? "取消全选" : "全选";
  if (dom.chatPending) dom.chatPending.classList.toggle("is-hidden", state.patches.length === 0);
  renderApplyBar();
}

function chatTraceHtml(trace) {
  if (!Array.isArray(trace) || !trace.length) return "";
  return `<details class="chat-trace"><summary>执行轨迹 ${trace.length} 步</summary><div class="chat-trace-list">${trace.map((step) => `
    <div class="chat-trace-step"><b>${Number(step.step) || ""}</b><span>${escapeHtml(step.detail || step.action || "")}</span></div>`).join("")}</div></details>`;
}

function chatMessageHtml(message) {
  const role = message.role === "user" ? "user" : (message.role === "system" ? "system" : "assistant");
  const content = escapeHtml(message.content || "");
  if (role === "system") {
    return `<div class="chat-msg system"><div class="chat-bubble"><div class="chat-text">${content}</div><div class="chat-meta">${escapeHtml(formatDate(message.at))}</div></div></div>`;
  }
  const ops = Array.isArray(message.operations) ? message.operations : [];
  const tags = [];
  if (ops.length) tags.push(`<span class="chat-ops-badge">${message.applied_version ? "已执行" : "待确认"} ${ops.length} 项</span>`);
  if (message.applied_version) tags.push(`<span class="chat-tag">${escapeHtml(message.applied_version)}</span>`);
  if (message.scope) tags.push(`<span>${escapeHtml(message.scope)}</span>`);
  if (message.model) tags.push(`<span>${escapeHtml(message.model)}</span>`);
  const trace = chatTraceHtml(message.trace);
  return `<div class="chat-msg ${role}"><div class="chat-avatar">${role === "user" ? "你" : "AI"}</div><div class="chat-bubble"><div class="chat-text">${content}</div><div class="chat-meta">${tags.join("")}<span>${escapeHtml(formatDate(message.at))}</span></div>${trace}</div></div>`;
}

function renderChat() {
  if (!dom.chatList) return;
  const messages = state.doc?.messages || [];
  if (!messages.length) {
    dom.chatList.innerHTML = `
      <div class="chat-empty">
        <span class="empty-mark">✦</span>
        <strong>先看清，再动 PPT</strong>
        <p>告诉我页码和要改的内容。AI 会自己读页、定位表格单元格、提出精确操作；开启「自动执行」后会直接写入并生成版本。</p>
      </div>`;
    return;
  }
  dom.chatList.innerHTML = messages.map(chatMessageHtml).join("");
  dom.chatList.scrollTop = dom.chatList.scrollHeight;
}

function renderMessages() {
  if (!state.doc) {
    dom.messageList.innerHTML = '<div class="message-empty">还没有会话记录</div>';
    return;
  }
  const messages = (state.doc.messages || []).slice(-60);
  dom.sessionDialogMeta.textContent = `${state.doc.title || state.doc.file_name} · ${messages.length} 条记录`;
  if (!messages.length) {
    dom.messageList.innerHTML = '<div class="message-empty">还没有对话。向 AI 描述一次修改，或直接点选文字编辑。</div>';
    return;
  }
  const kindLabels = {
    request: "需求",
    suggest: "AI 建议",
    apply: "已保存",
    restore: "版本恢复",
  };
  dom.messageList.innerHTML = messages.map((message) => {
    const role = message.role === "user" ? "user" : message.role === "assistant" ? "assistant" : "system";
    const tag = kindLabels[message.kind];
    const meta = [];
    if (tag) meta.push(`<span class="message-tag">${escapeHtml(tag)}</span>`);
    if (message.patch_count) meta.push(`<span>${Number(message.patch_count)} 条建议</span>`);
    if (message.version) meta.push(`<span>${escapeHtml(message.version)}</span>`);
    if (message.mode === "react" || message.mode === "react+fallback") {
      meta.push(`<span>${message.mode === "react" ? "ReAct" : "ReAct + 回退"}</span>`);
    }
    if (Array.isArray(message.trace) && message.trace.length) {
      meta.push(`<span>${message.trace.length} 步</span>`);
    }
    return `
      <div class="message ${role}">
        <div class="message-text">${escapeHtml(message.content || "")}</div>
        <div class="message-meta">
          ${meta.join("")}
          <span>${escapeHtml(formatDate(message.at))}</span>
        </div>
      </div>`;
  }).join("");
  dom.messageList.scrollTop = dom.messageList.scrollHeight;
}

function openSessionLog() {
  if (!state.doc) return;
  renderMessages();
  if (!dom.sessionDialog.open) dom.sessionDialog.showModal();
}

function renderApplyBar() {
  const accepted = state.patches.filter((patch) => patch.uiState === "accepted");
  const visible = accepted.length > 0;
  dom.applyBar.classList.toggle("is-hidden", !visible);
  dom.acceptedCount.textContent = state.patchesStale
    ? "建议已过期，请重新生成"
    : `已接受 ${accepted.length} 条 · 将保存为新版本`;
  dom.applyBtn.disabled = !visible || state.patchesStale || state.busy;
}

function setScope(scope) {
  state.scope = scope === "deck" ? "deck" : "current";
  dom.scopeSegment.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.scope === state.scope);
  });
  updateScanButtonLabel();
}

function scrollChatBottom() {
  if (dom.chatList) dom.chatList.scrollTop = dom.chatList.scrollHeight;
}

function appendChatElement(html) {
  dom.chatList.querySelector(".chat-empty")?.remove();
  const host = document.createElement("div");
  host.innerHTML = String(html).trim();
  const node = host.firstElementChild;
  if (node) {
    dom.chatList.appendChild(node);
    scrollChatBottom();
  }
  return node;
}

function partialThought(raw) {
  const match = String(raw || "").match(/"thought"\s*:\s*"((?:\\.|[^"\\])*)/);
  if (!match) return "";
  return match[1]
    .replace(/\\n/g, "\n")
    .replace(/\\"/g, '"')
    .replace(/\\\\/g, "\\")
    .slice(0, 800);
}

function toolLabel(action, args = {}) {
  if (action === "outline") return "读取整份大纲";
  if (action === "read_slides" || action === "read_slide" || action === "get_slides") {
    const list = Array.isArray(args.slide_indices) ? args.slide_indices : [args.slide_index];
    return `读取第 ${list.filter((v) => v !== undefined && v !== null).map((v) => Number(v) + 1).join("、") || "?"} 页`;
  }
  if (action === "search_text") return `搜索「${args.query || ""}」`;
  if (action === "read_table") return `读取表格 ${args.table_shape_id || ""}`;
  if (action === "inspect_shapes") return "查看对象细节";
  if (action === "propose_operations") return `提交 ${Number(args.count || 0)} 项操作`;
  if (action === "finish") return "完成";
  return action || "执行工具";
}

function createLiveBubble() {
  const root = appendChatElement(`
    <div class="chat-msg assistant"><div class="chat-avatar">AI</div>
      <div class="chat-bubble chat-live">
        <div class="live-head"><span class="live-dot"></span><span class="live-head-text">AI 正在连接…</span></div>
        <div class="live-thought"></div>
        <div class="live-steps"></div>
        <div class="live-output"></div>
      </div>
    </div>`);
  if (!root) return null;
  return {
    root,
    head: root.querySelector(".live-head-text"),
    thought: root.querySelector(".live-thought"),
    steps: root.querySelector(".live-steps"),
    output: root.querySelector(".live-output"),
    raw: "",
    step: 0,
    stepNodes: new Map(),
    ops: [],
  };
}

function addLiveStep(live, step, action, args) {
  if (!live) return;
  const row = document.createElement("div");
  row.className = "live-step";
  row.innerHTML = `<b>${Number(step) || ""}</b><span>${escapeHtml(toolLabel(action, args))}</span>`;
  live.steps.appendChild(row);
  live.stepNodes.set(Number(step), row.querySelector("span"));
  scrollChatBottom();
}

function handleChatEvent(event, data, live, doneRef) {
  if (!live) return;
  if (event === "start") {
    live.head.textContent = `AI 正在执行${data.model ? ` · ${data.model}` : ""}`;
  } else if (event === "step") {
    live.step = Number(data.step) || 0;
    live.raw = "";
    live.head.textContent = `第 ${live.step} 步 · 思考中…`;
  } else if (event === "token") {
    live.raw += String(data.text || "");
    const thought = partialThought(live.raw);
    if (thought) live.thought.textContent = thought;
  } else if (event === "thought") {
    live.thought.textContent = data.text || "";
  } else if (event === "tool_call") {
    addLiveStep(live, data.step, data.action, data.args || {});
    live.head.textContent = `第 ${Number(data.step) || ""} 步 · ${toolLabel(data.action, data.args || {})}`;
  } else if (event === "tool_result") {
    const node = live.stepNodes.get(Number(data.step));
    if (node && data.detail) node.textContent = data.detail;
  } else if (event === "operations") {
    const ops = Array.isArray(data.operations) ? data.operations : [];
    live.ops.push(...ops);
    live.output.textContent = `已提交 ${ops.length} 项操作，等待收尾…`;
  } else if (event === "summary") {
    live.output.textContent = data.text || data.summary || "";
  } else if (event === "error") {
    live.output.textContent = `执行出错：${data.message || "未知错误"}`;
    live.output.style.color = "var(--red)";
  } else if (event === "applying") {
    live.head.textContent = "正在写回 PPTX…";
    live.output.textContent = `正在应用 ${data.count || 0} 项操作，LibreOffice 重新渲染中…`;
  } else if (event === "applied") {
    live.head.textContent = `已执行 · ${data.version_id || ""}`;
    live.output.textContent = `已写入 ${data.operation_count || 0} 项操作，保存为 ${data.version_id || "新版本"}。`;
  } else if (event === "apply_error") {
    live.output.textContent = `自动执行失败：${data.message || "未知错误"}。操作已保留，可手动确认。`;
    live.output.style.color = "var(--red)";
  } else if (event === "done") {
    doneRef.value = data;
    if (!live.output.textContent) live.output.textContent = data.reply || data.summary || "";
    live.head.textContent = data.applied ? "已完成并保存" : "已完成，等待确认";
  }
}

function parseSseChunk(chunk) {
  let event = "message";
  const dataLines = [];
  for (const rawLine of String(chunk).split("\n")) {
    const line = rawLine.replace(/\r$/, "");
    if (line.startsWith("event:")) event = line.slice(6).trim();
    else if (line.startsWith("data:")) dataLines.push(line.slice(5).trimStart());
  }
  if (!dataLines.length) return null;
  try {
    return { event, data: JSON.parse(dataLines.join("\n")) };
  } catch {
    return null;
  }
}

async function refreshChatMessages(docId) {
  const targetDocId = docId || state.doc?.doc_id;
  if (!targetDocId) return;
  const data = await api(`/api/docs/${targetDocId}/messages`, {}, 30000);
  if (state.doc?.doc_id !== targetDocId) return; // user switched docs mid-flight
  state.doc.messages = data.messages || [];
  state.doc.message_count = state.doc.messages.length;
  renderChat();
  renderMessages();
  if (!state.streaming) restorePendingOperations();
}

function showPendingOperations(operations, baseVersion, messageId) {
  const ops = Array.isArray(operations) ? operations : [];
  state.patches = ops.map((op) => ({
    ...op,
    uiState: "accepted",
    ...(messageId ? { message_id: messageId } : {}),
  }));
  // staleness must come from the version the ops were computed against, not
  // from whatever version the doc happens to be on when the stream ends
  state.patchesStale = Number.isInteger(baseVersion) && baseVersion !== state.doc.version;
  renderPatchList();
  dom.chatPending.classList.toggle("is-hidden", state.patches.length === 0);
  if (state.patchesStale) {
    setNotice("这组操作基于旧版本生成，请重新生成后再应用。", "error");
  }
  if (state.patches.length) {
    dom.chatPending.scrollTop = 0;
  }
}

async function finalizeChat(done, slideIndex, docId) {
  const targetDocId = docId || state.doc?.doc_id;
  if (!targetDocId) return;
  if (done.applied) {
    const doc = await api(`/api/docs/${targetDocId}`, {}, 90000);
    if (state.doc?.doc_id !== targetDocId) {
      toast(`AI 已执行并保存为 ${done.applied_version_id || doc.version_id}`, "success");
      return;
    }
    const selected = state.selectedId;
    applyDocument(doc, { preserveSelection: Boolean(selected), clearPatches: true });
    const touched = (done.operations || [])
      .map((operation) => Number(operation.slide_index))
      .filter((index) => Number.isInteger(index) && index >= 0);
    const target = touched.length ? touched[0] : slideIndex;
    if (state.slideIndex !== target) setSlide(target, { keepSelection: true });
    state.patches = [];
    dom.chatPending.classList.add("is-hidden");
    toast(`AI 已执行并保存为 ${done.applied_version_id || doc.version_id}`, "success");
    setNotice(`已自动执行 ${(done.operations || []).length} 项操作 · ${done.reply || done.summary || ""}`, "success");
    return;
  }
  try {
    await refreshChatMessages(targetDocId);
  } catch (error) {
    if (done.assistant && state.doc?.doc_id === targetDocId) {
      state.doc.messages = [...(state.doc.messages || []), done.assistant];
      renderChat();
    }
  }
  if (state.doc?.doc_id !== targetDocId) return;
  const ops = done.operations || [];
  if (ops.length) {
    const target = Number(ops[0]?.slide_index);
    if (Number.isInteger(target) && target !== state.slideIndex) {
      setSlide(target, { keepSelection: true });
    }
    showPendingOperations(ops, done.base_version, done.message_id);
    setNotice(`AI 提出 ${ops.length} 项操作，已默认选中，确认后写入。`, "");
  } else {
    state.patches = [];
    renderPatchList();
    dom.chatPending.classList.add("is-hidden");
    setNotice(done.apply_error ? `自动执行失败：${done.apply_error}` : "", done.apply_error ? "error" : "");
  }
  if (done.apply_error) {
    toast(`自动执行失败：${done.apply_error}`, "error");
    if (ops.length) showPendingOperations(ops, done.base_version, done.message_id);
  }
}

async function sendChat() {
  if (!state.doc || state.streaming) return;
  const instruction = dom.chatInput.value.trim();
  if (!instruction) {
    toast("先说清楚要改什么", "error");
    dom.chatInput.focus();
    return;
  }
  if (!state.doc.ai_configured) {
    switchTab("ai");
    setNotice("后端还没有配置 AI key，无法使用对话改稿。", "error");
    return;
  }
  switchTab("ai");
  setNotice("");
  state.streaming = true;
  state.busy = true;
  state.patches = [];
  state.patchesStale = false;
  renderPatchList();
  dom.chatPending.classList.add("is-hidden");
  dom.sendBtn.disabled = true;
  dom.stopBtn.classList.remove("is-hidden");
  dom.chatInput.disabled = true;
  const controller = new AbortController();
  state.streamController = controller;
  const docId = state.doc.doc_id;
  const docVersion = state.doc.version;
  const slideIndex = state.slideIndex;
  const scope = state.scope;
  const selected = state.selectedId;
  const auto = Boolean(dom.autoApplyToggle.checked);
  const scopeLabel = scope === "deck" ? "整份文稿" : `第 ${slideIndex + 1} 页`;
  appendChatElement(chatMessageHtml({
    role: "user",
    content: instruction,
    at: new Date().toISOString(),
    scope: scopeLabel,
    model: "",
  }));
  const live = createLiveBubble();
  state.live = live;
  const doneRef = { value: null };
  let failure = "";
  let idleAborted = false;
  try {
    const response = await fetch(`/api/docs/${docId}/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        message: instruction,
        scope,
        slide_index: slideIndex,
        selected_shape_id: selected || null,
        auto_apply: auto,
        base_version: docVersion,
      }),
      signal: controller.signal,
    });
    if (!response.ok || !response.body) {
      let detail = `HTTP ${response.status}`;
      try {
        const data = await response.json();
        detail = data.detail || detail;
      } catch { /* ignore */ }
      const error = new Error(detail);
      error.status = response.status;
      throw error;
    }
    // request accepted — only now clear the composer, so a rejected send
    // (409 conflict, validation) keeps what the user typed
    dom.chatInput.value = "";
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    let idleTimer = setTimeout(() => {
      idleAborted = true;
      try {
        controller.abort();
      } catch { /* already aborted */ }
    }, 90000); // dead-connection guard: no bytes for 90s → abort
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        clearTimeout(idleTimer);
        idleTimer = setTimeout(() => {
          idleAborted = true;
          try {
            controller.abort();
          } catch { /* already aborted */ }
        }, 90000);
        buffer += decoder.decode(value, { stream: true });
        // CRLF-tolerant frame delimiter: a proxy normalizing \n to \r\n must
        // not fuse the whole stream into one unparseable buffer
        let boundary = buffer.match(/\r?\n\r?\n/);
        while (boundary) {
          const chunk = buffer.slice(0, boundary.index);
          buffer = buffer.slice(boundary.index + boundary[0].length);
          const parsed = parseSseChunk(chunk);
          if (parsed) {
            handleChatEvent(parsed.event, parsed.data, live, doneRef);
            if (parsed.event === "done") doneRef.value = parsed.data;
          }
          boundary = buffer.match(/\r?\n\r?\n/);
        }
      }
    } finally {
      clearTimeout(idleTimer);
    }
  } catch (error) {
    if (error.name === "AbortError") {
      if (idleAborted) {
        setNotice("连接超过 90 秒没有数据，已中断；输入已保留，可重试。", "error");
        if (!dom.chatInput.value) dom.chatInput.value = instruction;
        if (live?.output) live.output.textContent = "执行中断：连接超时";
        toast("对话连接超时", "error");
        failure = "";
      } else {
        failure = "已停止生成";
        setNotice("已停止生成；已写入的用户消息会保留在会话里。", "");
        if (!dom.chatInput.value) dom.chatInput.value = instruction;
      }
    } else {
      failure = error.message || "请求失败";
      if (error.status === 409 || /已经更新/.test(failure)) {
        // stale doc: refresh to the latest version and give the input back
        try {
          const fresh = await api(`/api/docs/${docId}`, {}, 60000);
          if (state.doc?.doc_id === docId) {
            applyDocument(fresh, { preserveSelection: true, clearPatches: true });
          }
          dom.chatInput.value = instruction;
          setNotice("文档已在别处更新，已加载最新版本；输入已保留，请重新发送。", "error");
          failure = "";
        } catch { /* fall through to generic error handling */ }
      }
      if (failure) {
        if (!dom.chatInput.value) dom.chatInput.value = instruction;
        setNotice(`对话失败：${failure}`, "error");
        if (live?.output) live.output.textContent = `执行中断：${failure}`;
        toast(`对话失败：${failure}`, "error");
      }
    }
  } finally {
    state.streaming = false;
    state.streamController = null;
    dom.sendBtn.disabled = false;
    dom.stopBtn.classList.add("is-hidden");
    dom.chatInput.disabled = false;
    if (live?.root) live.root.remove();
    state.live = null;
    try {
      if (doneRef.value) await finalizeChat(doneRef.value, slideIndex, docId);
      else await refreshChatMessages(docId);
    } catch (error) {
      toast(`收尾失败：${error.message || error}`, "error");
    }
    state.busy = false;
    updateScanButtonLabel();
    dom.chatInput.focus();
    scrollChatBottom();
  }
  if (failure && !doneRef.value) setNotice(failure, "error");
}

function togglePatchAccept(patchId) {
  const patch = state.patches.find((item) => item.id === patchId);
  if (!patch) return;
  patch.uiState = patch.uiState === "accepted" ? "pending" : "accepted";
  renderPatchList();
}

function togglePatchIgnore(patchId) {
  const patch = state.patches.find((item) => item.id === patchId);
  if (!patch) return;
  patch.uiState = patch.uiState === "ignored" ? "pending" : "ignored";
  renderPatchList();
}

function previewPatch(patch) {
  if (!patch) return;
  if (patch.action === "add_textbox") {
    setSlide(Number(patch.slide_index) || 0);
    toast(`新增文本框位置：x ${emuToCm((patch.bbox_emu || {}).x)} cm · y ${emuToCm((patch.bbox_emu || {}).y)} cm`);
    return;
  }
  if (!patch.shape_id) return;
  gotoObject(patch.shape_id, { flash: true });
}

function clearPatches() {
  state.patches = [];
  state.patchesStale = false;
  setNotice("");
  renderPatchList();
}

function toggleAcceptAll() {
  if (!state.patches.length) return;
  const allAccepted = state.patches.every((patch) => patch.uiState === "accepted");
  state.patches.forEach((patch) => {
    patch.uiState = allAccepted ? "pending" : "accepted";
  });
  renderPatchList();
}

async function applyAcceptedPatches() {
  if (!state.doc || state.busy) return;
  const accepted = state.patches.filter((patch) => patch.uiState === "accepted");
  if (!accepted.length) {
    toast("先接受至少一条建议", "error");
    return;
  }
  if (state.patchesStale) {
    toast("这组建议已过期，请重新生成", "error");
    return;
  }
  const patches = accepted.map((patch) => {
    const { uiState, message_id, ...operation } = patch;
    return operation;
  });
  const messageId = state.patches.find((patch) => patch.message_id)?.message_id || null;
  const doc = await runApply(patches, `AI 改稿 · ${patches.length} 项`, {
    successMessage: `${patches.length} 项 AI 操作已写入原 PPTX`,
    assistantMessageId: messageId,
  });
  const target = Number(accepted[0]?.slide_index);
  if (doc && Number.isInteger(target) && target !== state.slideIndex) {
    setSlide(target, { keepSelection: true });
  }
}

/* version history */
async function openHistory() {
  if (!state.doc) return;
  dom.historyDialog.showModal();
  dom.versionList.innerHTML = '<div class="version-empty">正在读取版本…</div>';
  try {
    const data = await api(`/api/docs/${state.doc.doc_id}/versions`, {}, 30000);
    renderVersionList(data);
  } catch (error) {
    dom.versionList.innerHTML = `<div class="version-empty">${escapeHtml(error.message)}</div>`;
  }
}

function renderVersionList(data) {
  const versions = [...(data.versions || [])].reverse();
  if (!versions.length) {
    dom.versionList.innerHTML = '<div class="version-empty">还没有版本记录</div>';
    return;
  }
  dom.versionList.innerHTML = versions.map((version) => {
    const current = version.is_current || version.id === state.doc.version_id;
    const thumb = `/api/docs/${state.doc.doc_id}/slides/0/image?thumb=1&v=${encodeURIComponent(version.id)}`;
    return `
      <div class="version-row ${current ? "is-current" : ""}">
        <img class="version-thumb" loading="lazy" src="${thumb}" alt="">
        <div class="version-info">
          <strong>${escapeHtml(version.id)} · ${escapeHtml(version.label || "编辑")}${current ? '<span class="current-tag">当前</span>' : ""}</strong>
          <p>${escapeHtml(formatDate(version.created_at))} · ${escapeHtml(formatSize(version.size))}</p>
        </div>
        <div class="version-actions">
          ${current
            ? '<button class="btn tiny ghost" type="button" disabled>当前版本</button>'
            : `<button class="btn tiny ghost" type="button" data-restore="${escapeHtml(version.id)}">恢复</button>`}
        </div>
      </div>`;
  }).join("");
}

async function restoreVersion(versionId) {
  if (!state.doc || state.busy || state.streaming) return;
  if (!window.confirm(`恢复到 ${versionId}？当前内容会保留在历史中，并生成一个新版本。`)) return;
  dom.historyDialog.close();
  setBusy(true, "正在恢复版本", `读取 ${versionId} 并重新渲染…`);
  try {
    const doc = await api(
      `/api/docs/${state.doc.doc_id}/versions/${encodeURIComponent(versionId)}/restore`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ base_version: state.doc.version }),
      },
      300000
    );
    applyDocument(doc, { preserveSelection: false, clearPatches: true });
    toast(`已恢复 ${versionId}，并保存为 ${doc.version_id}`, "success");
  } catch (error) {
    if (error.status === 409) {
      try {
        const fresh = await api(`/api/docs/${state.doc.doc_id}`, {}, 60000);
        applyDocument(fresh, { preserveSelection: true, clearPatches: true });
        toast("文档已在别处更新，已加载最新版本；请重新选择要恢复的版本", "error");
      } catch {
        toast(error.message, "error");
      }
    } else {
      toast(error.message, "error");
    }
  } finally {
    setBusy(false);
  }
}

/* upload */
async function handleFile(file) {
  if (!file || state.busy) return;
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    toast("只支持 .pptx 文件", "error");
    return;
  }
  const capMb = Number(state.doc?.max_upload_mb) || 60;
  if (file.size > capMb * 1024 * 1024) {
    toast(`文件超过 ${capMb} MB 限制`, "error");
    return;
  }
  const form = new FormData();
  form.append("file", file, file.name);
  setBusy(true, "正在上传并高保真渲染", "LibreOffice 正在生成每一页预览，通常需要几秒…");
  try {
    const doc = await api(
      "/api/docs/upload",
      { method: "POST", body: form },
      360000
    );
    applyDocument(doc, { resetSlide: true, preserveSelection: false, clearPatches: true });
    showWorkspace();
    history.replaceState(null, "", `#doc=${doc.doc_id}`);
    toast(`已上传「${doc.file_name}」，v0 渲染完成`, "success");
    requestAnimationFrame(fitStage);
  } catch (error) {
    toast(error.message, "error");
  } finally {
    setBusy(false);
    dom.fileInput.value = "";
  }
}

function downloadCurrent() {
  if (!state.doc) return;
  const anchor = document.createElement("a");
  anchor.href = state.doc.download_url;
  anchor.download = state.doc.file_name || "presentation.pptx";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

/* events */
function bindEvents() {
  const chooseFile = () => dom.fileInput.click();
  dom.uploadBtn.addEventListener("click", chooseFile);
  dom.landingUploadBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    chooseFile();
  });
  dom.fileInput.addEventListener("change", () => {
    const file = dom.fileInput.files && dom.fileInput.files[0];
    if (file) void handleFile(file);
  });

  dom.dropzone.addEventListener("click", (event) => {
    if (event.target.closest("button")) return;
    chooseFile();
  });
  dom.dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      chooseFile();
    }
  });
  dom.dropzone.addEventListener("dragenter", (event) => {
    event.preventDefault();
    dom.dropzone.classList.add("is-dragover");
  });
  dom.dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
  });
  dom.dropzone.addEventListener("dragleave", (event) => {
    if (!dom.dropzone.contains(event.relatedTarget)) dom.dropzone.classList.remove("is-dragover");
  });
  dom.dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    event.stopPropagation();
    dom.dropzone.classList.remove("is-dragover");
    const file = event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) void handleFile(file);
  });
  document.addEventListener("dragover", (event) => event.preventDefault());
  document.addEventListener("drop", (event) => {
    event.preventDefault();
    if (dom.dropzone.contains(event.target)) return;
    const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) void handleFile(file);
  });

  dom.recentList.addEventListener("click", (event) => {
    const deleteButton = event.target.closest("[data-delete]");
    if (deleteButton) {
      event.stopPropagation();
      void deleteDocument(deleteButton.dataset.delete, deleteButton.dataset.title || "会话");
      return;
    }
    const openButton = event.target.closest("[data-open]");
    if (openButton) void openDocument(openButton.dataset.open);
  });
  dom.refreshDocsBtn.addEventListener("click", () => void loadRecent());
  dom.backBtn.addEventListener("click", goHome);
  $id("brandLink").addEventListener("click", (event) => {
    event.preventDefault();
    goHome();
  });

  dom.downloadBtn.addEventListener("click", downloadCurrent);
  dom.historyBtn.addEventListener("click", () => void openHistory());
  dom.sessionLogBtn.addEventListener("click", openSessionLog);
  dom.closeSessionLogBtn.addEventListener("click", () => dom.sessionDialog.close());
  dom.sessionDialog.addEventListener("click", (event) => {
    if (event.target === dom.sessionDialog) dom.sessionDialog.close();
  });
  dom.closeHistoryBtn.addEventListener("click", () => dom.historyDialog.close());
  dom.historyDialog.addEventListener("click", (event) => {
    if (event.target === dom.historyDialog) dom.historyDialog.close();
  });
  dom.versionList.addEventListener("click", (event) => {
    const button = event.target.closest("[data-restore]");
    if (button) void restoreVersion(button.dataset.restore);
  });

  dom.prevBtn.addEventListener("click", () => setSlide(state.slideIndex - 1));
  dom.nextBtn.addEventListener("click", () => setSlide(state.slideIndex + 1));
  dom.addTextBtn.addEventListener("click", startDraftTextbox);

  dom.tabAi.addEventListener("click", () => switchTab("ai"));
  dom.tabText.addEventListener("click", () => switchTab("text"));

  dom.scopeSegment.addEventListener("click", (event) => {
    const button = event.target.closest("[data-scope]");
    if (button) setScope(button.dataset.scope);
  });
  dom.promptChips.addEventListener("click", (event) => {
    const button = event.target.closest("[data-prompt]");
    if (!button) return;
    const prompt = button.dataset.prompt;
    const current = dom.chatInput.value.trim();
    dom.chatInput.value = current ? `${current}；${prompt}` : prompt;
    dom.chatInput.focus();
  });
  dom.sendBtn.addEventListener("click", () => void sendChat());
  dom.stopBtn.addEventListener("click", () => {
    if (state.streamController) state.streamController.abort();
  });
  dom.chatInput.addEventListener("keydown", (event) => {
    event.stopPropagation();
    if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
      event.preventDefault();
      void sendChat();
    }
  });
  try {
    dom.autoApplyToggle.checked = localStorage.getItem("ppt-light-auto-apply") === "1";
  } catch { /* ignore */ }
  dom.autoApplyToggle.addEventListener("change", () => {
    try { localStorage.setItem("ppt-light-auto-apply", dom.autoApplyToggle.checked ? "1" : "0"); } catch { /* ignore */ }
  });
  dom.acceptAllBtn.addEventListener("click", toggleAcceptAll);
  dom.clearPatchesBtn.addEventListener("click", clearPatches);
  dom.applyBtn.addEventListener("click", () => void applyAcceptedPatches());

  dom.patchList.addEventListener("click", (event) => {
    const card = event.target.closest(".patch-card");
    if (!card) return;
    const patch = state.patches.find((item) => item.id === card.dataset.id);
    if (!patch) return;
    const actionButton = event.target.closest("[data-action]");
    if (actionButton) {
      event.stopPropagation();
      const action = actionButton.dataset.action;
      if (action === "accept") togglePatchAccept(patch.id);
      else if (action === "ignore") togglePatchIgnore(patch.id);
      else previewPatch(patch);
      return;
    }
    previewPatch(patch);
  });

  dom.saveTextBtn.addEventListener("click", () => void saveSelectedText());
  dom.deleteShapeBtn.addEventListener("click", () => void deleteSelectedShape());
  dom.applyStyleBtn.addEventListener("click", () => void applySelectedStyle());
  dom.textEditor.addEventListener("keydown", (event) => {
    if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
      event.preventDefault();
      void saveSelectedText();
    }
  });

  dom.slideImage.addEventListener("pointerdown", () => {
    if (!state.inline) setSelected(null);
  });
  dom.stage.addEventListener("pointerdown", (event) => {
    if (event.target === dom.stage) setSelected(null);
  });
  dom.slideImage.addEventListener("load", () => {
    dom.stageLoader.classList.add("is-hidden");
    fitStage();
  });
  dom.slideImage.addEventListener("error", () => {
    dom.stageLoader.classList.add("is-hidden");
    toast("预览图加载失败，请检查渲染版本", "error");
  });

  let resizeTimer = null;
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(fitStage, 100);
  });

  document.addEventListener("keydown", (event) => {
    if (state.busy || state.inline || dom.historyDialog.open) return;
    const target = event.target;
    const typing = target instanceof HTMLElement
      && (["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName) || target.isContentEditable);
    if (typing) return;
    if (event.key === "ArrowLeft") {
      event.preventDefault();
      setSlide(state.slideIndex - 1);
    } else if (event.key === "ArrowRight") {
      event.preventDefault();
      setSlide(state.slideIndex + 1);
    } else if (event.key === "Escape") {
      setSelected(null);
    }
  });

  window.addEventListener("hashchange", () => {
    const match = location.hash.match(/^#doc=([a-f0-9]{32})$/);
    if (match) {
      if (!state.doc || state.doc.doc_id !== match[1]) void openDocument(match[1]);
    } else if (state.doc) {
      goHome();
    }
  });
}

function init() {
  bindEvents();
  const match = location.hash.match(/^#doc=([a-f0-9]{32})$/);
  if (match) {
    void openDocument(match[1]);
  } else {
    showLanding();
    void loadRecent();
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init, { once: true });
} else {
  init();
}


