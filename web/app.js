/** @typedef {{ name: string, type: 'text' }} TextField */
/** @typedef {{ name: string, type: 'image' }} ImageField */
/** @typedef {{ name: string, type: 'bool', trueText: string, falseText: string, label: string }} BoolField */
/** @typedef {{ name: string, type: 'auto', label: string }} AutoField */
/** @typedef {TextField | BoolField | AutoField} ListItemField */
/** @typedef {{ name: string, type: 'list', itemFields: ListItemField[] }} ListField */
/** @typedef {TextField | ImageField | ListField | BoolField | AutoField} FormField */

const app = document.getElementById("app");
const banner = document.getElementById("banner");
const statusDot = document.getElementById("statusDot");
const statusText = document.getElementById("statusText");

/**
 * @type {{
 *   file: File | null,
 *   filename: string,
 *   fields: FormField[],
 *   values: Record<string, any>,
 *   images: Record<string, File | null>,
 *   skip: Record<string, boolean>,
 *   rowSkip: Record<string, Record<string, boolean>[]>,
 *   showFullContent: boolean,
 *   fileContent: string,
 *   fileLayout: any[],
 *   objectUrls: string[],
 * }}
 */
const state = {
  file: null,
  filename: "",
  fields: [],
  values: {},
  images: {},
  skip: {},
  rowSkip: {},
  showFullContent: true,
  fileContent: "",
  fileLayout: [],
  objectUrls: [],
};

/** View-only column widths (px). Never written to the .docx. */
const viewTableColWidths = new Map();

const DRAFT_DB = "docx-form";
const DRAFT_STORE = "draft";
const DRAFT_KEY = "current";
let draftSaveTimer = 0;
let draftRestoring = false;

function openDraftDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DRAFT_DB, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(DRAFT_STORE)) db.createObjectStore(DRAFT_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("indexedDB open failed"));
  });
}

function draftStore(mode) {
  return openDraftDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const tx = db.transaction(DRAFT_STORE, mode);
        const store = tx.objectStore(DRAFT_STORE);
        resolve({ db, tx, store });
        tx.onerror = () => reject(tx.error || new Error("indexedDB tx failed"));
      }),
  );
}

async function draftGet() {
  const { store } = await draftStore("readonly");
  return new Promise((resolve, reject) => {
    const req = store.get(DRAFT_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function draftSet(value) {
  const { store } = await draftStore("readwrite");
  return new Promise((resolve, reject) => {
    const req = store.put(value, DRAFT_KEY);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

async function draftClear() {
  try {
    const { store } = await draftStore("readwrite");
    await new Promise((resolve, reject) => {
      const req = store.delete(DRAFT_KEY);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  } catch {
    /* ignore */
  }
}

async function fileToDraftBlob(file) {
  if (!file) return null;
  const buffer = await file.arrayBuffer();
  return { name: file.name, type: file.type || "", buffer };
}

function draftBlobToFile(blob) {
  if (!blob?.buffer) return null;
  return new File([blob.buffer], blob.name || "file", {
    type: blob.type || "application/octet-stream",
  });
}

async function saveDraftNow() {
  if (draftRestoring) return;
  if (!state.file) {
    await draftClear();
    return;
  }
  try {
    /** @type {Record<string, any>} */
    const images = {};
    for (const [name, file] of Object.entries(state.images)) {
      images[name] = file ? await fileToDraftBlob(file) : null;
    }
    await draftSet({
      v: 1,
      savedAt: new Date().toISOString(),
      filename: state.filename,
      fields: state.fields,
      values: state.values,
      skip: state.skip,
      rowSkip: state.rowSkip,
      showFullContent: state.showFullContent,
      fileContent: state.fileContent,
      fileLayout: state.fileLayout,
      viewColWidths: Object.fromEntries(viewTableColWidths),
      file: await fileToDraftBlob(state.file),
      images,
    });
  } catch (err) {
    console.warn("save draft failed", err);
  }
}

function scheduleSaveDraft() {
  if (draftRestoring) return;
  window.clearTimeout(draftSaveTimer);
  draftSaveTimer = window.setTimeout(() => {
    void saveDraftNow();
  }, 250);
}

async function restoreDraft() {
  try {
    const draft = await draftGet();
    if (!draft?.file?.buffer || !Array.isArray(draft.fields) || !draft.fields.length) {
      return false;
    }
    draftRestoring = true;
    revokeObjectUrls();
    state.file = draftBlobToFile(draft.file);
    state.filename = draft.filename || state.file?.name || "";
    state.fields = draft.fields;
    state.values = draft.values && typeof draft.values === "object" ? draft.values : {};
    state.skip = draft.skip && typeof draft.skip === "object" ? draft.skip : {};
    state.rowSkip = draft.rowSkip && typeof draft.rowSkip === "object" ? draft.rowSkip : {};
    state.showFullContent = draft.showFullContent !== false;
    state.fileContent = typeof draft.fileContent === "string" ? draft.fileContent : "";
    state.fileLayout = Array.isArray(draft.fileLayout) ? draft.fileLayout : [];
    state.images = {};
    if (draft.images && typeof draft.images === "object") {
      for (const [name, blob] of Object.entries(draft.images)) {
        state.images[name] = draftBlobToFile(blob);
      }
    }
    viewTableColWidths.clear();
    if (draft.viewColWidths && typeof draft.viewColWidths === "object") {
      for (const [k, widths] of Object.entries(draft.viewColWidths)) {
        if (Array.isArray(widths)) viewTableColWidths.set(k, widths.map(Number));
      }
    }
    return !!state.file;
  } catch (err) {
    console.warn("restore draft failed", err);
    return false;
  } finally {
    draftRestoring = false;
  }
}

function showBanner(message, kind = "error") {
  banner.hidden = !message;
  banner.textContent = message || "";
  banner.classList.toggle("success", kind === "success");
  if (message && kind === "error") {
    banner.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

/**
 * In-page dialog (native confirm is often blocked in Simple Browser).
 * @param {{
 *   title: string,
 *   body: string,
 *   actions: { id: string, label: string, kind?: 'primary' | 'secondary' | 'ghost' }[],
 * }} opts
 * @returns {Promise<string>} action id, or "abort" on Escape / overlay click
 */
function askDialog(opts) {
  return new Promise((resolve) => {
    const existing = document.getElementById("appDialog");
    if (existing) existing.remove();

    const actionsHtml = opts.actions
      .map((a) => {
        const kind = a.kind || "secondary";
        const cls =
          kind === "primary" ? "btn-primary" : kind === "ghost" ? "btn-ghost" : "btn-secondary";
        return `<button type="button" class="${cls}" data-dialog="${escapeAttr(a.id)}">${escapeHtml(a.label)}</button>`;
      })
      .join("");

    const overlay = document.createElement("div");
    overlay.id = "appDialog";
    overlay.className = "dialog-overlay";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = `
      <div class="dialog-card">
        <h2 class="dialog-title">${escapeHtml(opts.title)}</h2>
        <p class="dialog-body">${escapeHtml(opts.body)}</p>
        <div class="dialog-actions dialog-actions-stack">${actionsHtml}</div>
      </div>
    `;

    const finish = (value) => {
      overlay.remove();
      document.removeEventListener("keydown", onKey);
      resolve(value);
    };
    const onKey = (e) => {
      if (e.key === "Escape") finish("abort");
    };

    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) finish("abort");
    });
    for (const a of opts.actions) {
      overlay
        .querySelector(`[data-dialog="${cssAttrSelector(a.id)}"]`)
        ?.addEventListener("click", () => finish(a.id));
    }
    document.addEventListener("keydown", onKey);
    document.body.appendChild(overlay);
    const focusId = opts.actions.find((a) => a.kind === "primary")?.id || opts.actions[0]?.id;
    if (focusId) {
      /** @type {HTMLButtonElement | null} */ (
        overlay.querySelector(`[data-dialog="${cssAttrSelector(focusId)}"]`)
      )?.focus();
    }
  });
}

/** @returns {Promise<'clear' | 'keep' | 'abort'>} */
function askClearAfterCreate() {
  return askDialog({
    title: "ล้างช่องหลังสร้างเอกสารไหม?",
    body: "จะสร้างและดาวน์โหลดไฟล์ก่อน จากนั้นเลือกได้ว่าจะล้างช่องที่กรอกไว้หรือเก็บค่าไว้กรอกต่อ (เทมเพลตยังอยู่)",
    actions: [
      { id: "clear", label: "สร้างและล้างช่อง", kind: "primary" },
      { id: "keep", label: "สร้างแต่เก็บค่าไว้", kind: "secondary" },
      { id: "abort", label: "ยกเลิก", kind: "ghost" },
    ],
  });
}

/** @returns {Promise<boolean>} */
async function askChangeFile() {
  const choice = await askDialog({
    title: "เปลี่ยนไฟล์เทมเพลต?",
    body: "การเปลี่ยนไฟล์จะล้างค่าที่กรอกและรูปที่เลือกไว้ทั้งหมด",
    actions: [
      { id: "ok", label: "เปลี่ยนไฟล์", kind: "primary" },
      { id: "abort", label: "ยกเลิก", kind: "ghost" },
    ],
  });
  return choice === "ok";
}

function clearFieldErrors() {
  app.querySelectorAll(".field-error, .has-error").forEach((el) => {
    if (el.classList.contains("field-error")) el.remove();
    else el.classList.remove("has-error");
  });
}

/** Highlight empty required fields and return their display names. */
function validateBeforeSubmit() {
  clearFieldErrors();
  /** @type {string[]} */
  const missing = [];

  for (const f of state.fields) {
    if (state.skip[f.name]) continue;

    if (f.type === "text") {
      if (!String(state.values[f.name] ?? "").trim()) {
        missing.push(f.name);
        markFieldError(`#f-${cssId(f.name)}`, f.name);
      }
    } else if (f.type === "bool") {
      // always true/false — no required check
    } else if (f.type === "image") {
      if (!state.images[f.name]) {
        missing.push(f.name);
        const wrap = app.querySelector(`[data-image-field="${cssAttrSelector(f.name)}"]`);
        if (wrap) {
          wrap.classList.add("has-error");
          wrap.insertAdjacentHTML(
            "beforeend",
            `<p class="field-error">กรุณาเลือกไฟล์ หรือติ๊ก “ไม่ประสงค์กรอก”</p>`,
          );
        }
      }
    } else if (f.type === "list") {
      const rows = state.values[f.name] || [];
      if (rows.length === 0) {
        missing.push(`${f.name} (ต้องมีอย่างน้อย 1 แถว)`);
        const rep = app.querySelector(`[data-list-field="${cssAttrSelector(f.name)}"]`);
        if (rep) {
          rep.classList.add("has-error");
          rep.insertAdjacentHTML(
            "beforeend",
            `<p class="field-error">เพิ่มแถว หรือติ๊ก “ไม่ประสงค์กรอก”</p>`,
          );
        }
        continue;
      }
      rows.forEach((row, idx) => {
        for (const item of f.itemFields) {
          if (item.type === "bool" || item.type === "auto") continue;
          if (state.rowSkip[f.name]?.[idx]?.[item.name]) continue;
          if (!String(row[item.name] ?? "").trim()) {
            const label = `${f.name} แถว ${idx + 1}.${item.name}`;
            missing.push(label);
            markFieldError(
              `#l-${cssId(f.name)}-${idx}-${cssId(item.name)}`,
              item.name,
            );
          }
        }
      });
    }
  }

  return missing;
}

function markFieldError(inputSelector, name) {
  let input = app.querySelector(inputSelector);
  if (!input && name) {
    input = app.querySelector(`[data-doc-field="${cssAttrSelector(name)}"]`);
  }
  if (!input) return;
  const field =
    input.closest(".field") ||
    input.closest(".doc-blank-wrap") ||
    input.parentElement;
  if (field) field.classList.add("has-error");
  if (!field?.querySelector(".field-error")) {
    input.insertAdjacentHTML(
      "afterend",
      `<p class="field-error">ยังไม่ได้กรอก — กรอกค่า หรือติ๊ก “ไม่ประสงค์กรอก”</p>`,
    );
  }
}

function cssAttrSelector(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function revokeObjectUrls() {
  for (const u of state.objectUrls) URL.revokeObjectURL(u);
  state.objectUrls = [];
}

function resetFormState() {
  revokeObjectUrls();
  state.file = null;
  state.filename = "";
  state.fields = [];
  state.values = {};
  state.images = {};
  state.skip = {};
  state.rowSkip = {};
  // keep showFullContent preference across re-uploads in the same session
  state.fileContent = "";
  state.fileLayout = [];
  viewTableColWidths.clear();
}

/** Clear entered values/images/skip flags but keep the loaded template. */
function clearFilledFields() {
  revokeObjectUrls();
  state.values = {};
  state.images = {};
  state.skip = {};
  state.rowSkip = {};
  for (const f of state.fields) {
    state.skip[f.name] = false;
    if (f.type === "text") state.values[f.name] = "";
    else if (f.type === "bool") state.values[f.name] = false;
    else if (f.type === "image") state.images[f.name] = null;
    else if (f.type === "list") ensureListRows(f.name, f.itemFields);
  }
  scheduleSaveDraft();
}

function skipCheckbox(fieldName, checked) {
  const id = `skip-${cssId(fieldName)}`;
  return `
    <label class="skip-opt" for="${id}">
      <input type="checkbox" id="${id}" data-skip-field="${escapeAttr(fieldName)}" ${checked ? "checked" : ""} />
      ไม่ประสงค์กรอก
    </label>`;
}

function rowSkipCheckbox(listName, rowIdx, itemName, checked) {
  const id = `skip-${cssId(listName)}-${rowIdx}-${cssId(itemName)}`;
  return `
    <label class="skip-opt skip-opt-sm" for="${id}">
      <input type="checkbox" id="${id}"
        data-skip-list="${escapeAttr(listName)}"
        data-skip-row="${rowIdx}"
        data-skip-item="${escapeAttr(itemName)}"
        ${checked ? "checked" : ""} />
      ไม่ประสงค์กรอก
    </label>`;
}

function ensureRowSkip(listName, rowCount, itemFields) {
  if (!Array.isArray(state.rowSkip[listName])) state.rowSkip[listName] = [];
  while (state.rowSkip[listName].length < rowCount) {
    const map = {};
    for (const f of itemFields) map[f.name] = false;
    state.rowSkip[listName].push(map);
  }
  if (state.rowSkip[listName].length > rowCount) {
    state.rowSkip[listName].length = rowCount;
  }
  for (const row of state.rowSkip[listName]) {
    for (const f of itemFields) {
      if (row[f.name] === undefined) row[f.name] = false;
    }
  }
}

function setCreditsVersion(version) {
  const label = version ? `v${String(version).replace(/^v/i, "")}` : "v…";
  const appVer = document.getElementById("appVersion");
  const creditsVer = document.getElementById("creditsVersion");
  if (appVer) appVer.textContent = label;
  if (creditsVer) creditsVer.textContent = label;
}

async function checkHealth() {
  try {
    const res = await fetch("/api/health");
    if (!res.ok) throw new Error("bad");
    const data = await res.json().catch(() => ({}));
    setCreditsVersion(data.version || "0.0.0");
    statusDot.classList.add("ok");
    statusDot.classList.remove("bad");
    statusText.textContent = "localhost · พร้อมใช้";
  } catch {
    statusDot.classList.remove("ok");
    statusDot.classList.add("bad");
    statusText.textContent = "เชื่อมต่อไม่ได้";
    showBanner("เซิร์ฟเวอร์ไม่ตอบสนอง — ลองรีเฟรชหน้า หรือรัน docx-form ใหม่");
  }
}

function summaryText(fields) {
  let text = 0;
  let image = 0;
  let list = 0;
  let bool = 0;
  for (const f of fields) {
    if (f.type === "text") text += 1;
    else if (f.type === "image") image += 1;
    else if (f.type === "bool") bool += 1;
    else list += 1;
  }
  const parts = [];
  if (text) parts.push(`${text} ข้อความ`);
  if (bool) parts.push(`${bool} ติ๊ก`);
  if (image) parts.push(`${image} รูป`);
  if (list) parts.push(`${list} รายการ`);
  return parts.join(" · ") || "ไม่มีฟิลด์";
}

function ensureListRows(name, itemFields) {
  if (!Array.isArray(state.values[name])) {
    state.values[name] = [{}];
  }
  for (const row of state.values[name]) {
    for (const f of itemFields) {
      if (row[f.name] === undefined) {
        if (f.type === "bool") row[f.name] = false;
        else if (f.type === "auto") row[f.name] = ""; // filled on submit by row index
        else row[f.name] = "";
      }
    }
  }
  ensureRowSkip(name, state.values[name].length, itemFields);
}

function renderEmpty() {
  app.innerHTML = `
    <h1 class="headline">เติมตัวแปรใน Word</h1>
    <p class="subhead">อัปโหลดเทมเพลต แล้วกรอกค่าในฟอร์ม</p>
    <div class="upload-layout">
      <button type="button" class="dropzone" id="dropzone" aria-label="อัปโหลดไฟล์ Word">
        <div class="dropzone-title">ลากไฟล์ .docx มาวางที่นี่</div>
        <div class="dropzone-sub">หรือ <span class="linkish">เลือกไฟล์</span></div>
      </button>
      <aside class="upload-side">
        <label class="upload-opt" for="showFullContent">
          <input type="checkbox" id="showFullContent" ${state.showFullContent ? "checked" : ""} />
          <span>
            <strong>แสดงรายละเอียดเนื้อหาในไฟล์ทั้งหมด</strong>
            <small>ข้อความเป็นตัวอักษร ส่วนตัวแปรเป็นช่องว่างให้กรอก เหมือนใน Word</small>
          </span>
        </label>
        <div class="upload-tips">
          <p class="upload-tips-title">วิธีใส่ตัวแปรในเทมเพลต</p>
          <ul>
            <li><span>ข้อความ</span> <code>{{ชื่อ}}</code></li>
            <li>
              <span>รูป</span>
              <code>{{logo}}</code>
              <small class="tip-note">วางรูปใน Word → คลิกขวาที่รูป → View Alt Text / Edit Alt Text → ใส่แท็ก เช่น {{logo}}</small>
            </li>
            <li><span>ตาราง / ลิสต์</span> <code>{{#รายการ}}…{{/}}</code></li>
          </ul>
        </div>
      </aside>
    </div>
    <input class="sr-only" type="file" id="fileInput" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" />
    <p class="hint-error" id="uploadError" hidden></p>
  `;

  const dropzone = document.getElementById("dropzone");
  const fileInput = /** @type {HTMLInputElement} */ (document.getElementById("fileInput"));
  const uploadError = document.getElementById("uploadError");
  const showFull = /** @type {HTMLInputElement} */ (document.getElementById("showFullContent"));

  showFull.addEventListener("change", () => {
    state.showFullContent = showFull.checked;
  });

  const onFiles = (files) => {
    const file = files?.[0];
    if (!file) return;
    state.showFullContent = showFull.checked;
    void parseFile(file, uploadError);
  };

  dropzone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => onFiles(fileInput.files));

  ["dragenter", "dragover"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });
  dropzone.addEventListener("drop", (e) => {
    onFiles(e.dataTransfer?.files);
  });
}

/**
 * @param {File} file
 * @param {HTMLElement | null} errorEl
 */
async function parseFile(file, errorEl) {
  showBanner("");
  if (errorEl) {
    errorEl.hidden = true;
    errorEl.textContent = "";
  }

  const lower = file.name.toLowerCase();
  if (lower.endsWith(".doc") && !lower.endsWith(".docx")) {
    const msg = "รองรับเฉพาะ .docx — ให้ Save As เป็น .docx ใน Word ก่อน";
    if (errorEl) {
      errorEl.hidden = false;
      errorEl.textContent = msg;
    } else showBanner(msg);
    return;
  }
  if (!lower.endsWith(".docx")) {
    const msg = "ไฟล์ต้องเป็น .docx เท่านั้น";
    if (errorEl) {
      errorEl.hidden = false;
      errorEl.textContent = msg;
    } else showBanner(msg);
    return;
  }

  const fd = new FormData();
  fd.append("file", file);
  if (state.showFullContent) fd.append("includeContent", "1");
  app.setAttribute("aria-busy", "true");
  try {
    const res = await fetch("/api/parse", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) {
      const msg = data.error || "อ่านเทมเพลตไม่สำเร็จ";
      if (errorEl) {
        errorEl.hidden = false;
        errorEl.textContent = msg;
      } else showBanner(msg);
      return;
    }

    const preferContent = state.showFullContent;
    resetFormState();
    state.showFullContent = preferContent;
    state.file = file;
    state.filename = data.filename || file.name;
    state.fields = data.fields || [];
    state.fileContent =
      preferContent && typeof data.content === "string" ? data.content : "";
    state.fileLayout =
      preferContent && Array.isArray(data.layout) ? data.layout : [];
    for (const f of state.fields) {
      state.skip[f.name] = false;
      if (f.type === "text") state.values[f.name] = "";
      else if (f.type === "bool") state.values[f.name] = false;
      else if (f.type === "image") state.images[f.name] = null;
      else if (f.type === "list") ensureListRows(f.name, f.itemFields);
    }
    renderForm();
    scheduleSaveDraft();
  } catch {
    const msg = "เชื่อมต่อเซิร์ฟเวอร์ไม่ได้";
    if (errorEl) {
      errorEl.hidden = false;
      errorEl.textContent = msg;
    } else showBanner(msg);
  } finally {
    app.removeAttribute("aria-busy");
  }
}

/** Fields covered by inline document preview (text + lists). Images stay in sidebar form. */
function fieldsOutsidePreview() {
  if (!state.showFullContent) return state.fields;
  if (!state.fileLayout.length && !state.fileContent) return state.fields;
  return state.fields.filter((f) => f.type === "image");
}

function renderForm() {
  const useDocLayout =
    state.showFullContent && (state.fileLayout.length > 0 || !!state.fileContent);
  const fieldsHtml = fieldsOutsidePreview()
    .map((f) => renderField(f))
    .join("");
  const bodyHtml = state.fileLayout.length
    ? renderDocumentFromLayout(state.fileLayout)
    : renderDocumentLayout(state.fileContent);
  const contentPanel = useDocLayout
    ? `
    <div class="doc-sheet" id="docPreview">
      <div class="doc-sheet-label">ตัวอย่างเนื้อหาในไฟล์ (กรอกในช่องว่างได้)</div>
      <div class="doc-body">${bodyHtml}</div>
    </div>`
    : state.showFullContent
      ? `<p class="dropzone-hint">เลือกแสดงเนื้อหาแล้ว แต่ไม่พบข้อความในไฟล์</p>`
      : "";

  const extraFields =
    fieldsHtml &&
    `<div class="extra-fields">
      ${useDocLayout ? `<h2 class="extra-fields-title">ฟิลด์เพิ่มเติม (รูปภาพ)</h2>` : ""}
      ${fieldsHtml}
    </div>`;

  app.innerHTML = `
    <div class="file-chip">
      <strong>${escapeHtml(state.filename)}</strong>
      <span class="file-meta">${escapeHtml(summaryText(state.fields))}</span>
      <div class="file-chip-actions">
        <button type="button" class="btn-secondary btn-compact" id="exportJson" title="ส่งออกค่าที่กรอกเป็น JSON">Export JSON</button>
        <button type="button" class="btn-secondary btn-compact" id="importJson" title="นำเข้าค่าจากไฟล์ JSON">Import JSON</button>
        <input class="sr-only" type="file" id="importJsonInput" accept="application/json,.json" />
        <button type="button" class="linkish" id="changeFile">เปลี่ยนไฟล์</button>
      </div>
    </div>
    <form id="fillForm">
      ${contentPanel}
      ${useDocLayout ? extraFields || "" : fieldsHtml}
      <div class="sticky-bar">
        <button type="submit" class="btn-primary" id="submitBtn">สร้างเอกสาร</button>
        <p class="dropzone-hint" style="margin:0.5rem 0 0;text-align:center">
          จะดาวน์โหลดไฟล์ใหม่ โดยไม่แก้ไฟล์ต้นฉบับ · ช่องว่างต้องติ๊ก “ไม่ประสงค์กรอก” · Export/Import JSON ได้ด้านบน
        </p>
      </div>
    </form>
  `;

  document.getElementById("changeFile")?.addEventListener("click", async () => {
    if (!(await askChangeFile())) return;
    resetFormState();
    void draftClear();
    showBanner("");
    renderEmpty();
  });

  document.getElementById("exportJson")?.addEventListener("click", () => {
    void exportValuesJson();
  });
  const importInput = /** @type {HTMLInputElement | null} */ (document.getElementById("importJsonInput"));
  document.getElementById("importJson")?.addEventListener("click", () => importInput?.click());
  importInput?.addEventListener("change", () => {
    const file = importInput.files?.[0];
    importInput.value = "";
    if (file) void importValuesJson(file);
  });

  wireFieldEvents();
  wireDocumentPreviewEvents();
  document.getElementById("fillForm")?.addEventListener("submit", onSubmit);
}

function bufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function base64ToUint8Array(b64) {
  const binary = atob(String(b64 || ""));
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

function valuesJsonBasename() {
  const base = String(state.filename || "template").replace(/\.docx$/i, "");
  return `${base}-values.json`;
}

async function buildValuesExportPayload() {
  /** @type {Record<string, any>} */
  const images = {};
  for (const [name, file] of Object.entries(state.images)) {
    if (!file) {
      images[name] = null;
      continue;
    }
    const buffer = await file.arrayBuffer();
    images[name] = {
      name: file.name,
      type: file.type || "application/octet-stream",
      base64: bufferToBase64(buffer),
    };
  }
  return {
    format: "docx-form-values",
    version: 1,
    exportedAt: new Date().toISOString(),
    template: state.filename || "",
    fields: state.fields.map((f) => ({ name: f.name, type: f.type })),
    values: state.values,
    skip: state.skip,
    rowSkip: state.rowSkip,
    images,
  };
}

async function exportValuesJson() {
  if (!state.file) {
    showBanner("ยังไม่มีเทมเพลต — อัปโหลดไฟล์ก่อน");
    return;
  }
  try {
    const payload = await buildValuesExportPayload();
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = valuesJsonBasename();
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showBanner(`ส่งออก ${valuesJsonBasename()} แล้ว`, "success");
    setTimeout(() => showBanner(""), 2200);
  } catch (err) {
    console.warn(err);
    showBanner("ส่งออก JSON ไม่สำเร็จ");
  }
}

/**
 * @param {File} file
 */
async function importValuesJson(file) {
  if (!state.file) {
    showBanner("อัปโหลดเทมเพลต .docx ก่อน แล้วค่อย Import JSON");
    return;
  }
  try {
    const text = await file.text();
    const data = JSON.parse(text);
    if (!data || typeof data !== "object") throw new Error("invalid");
    if (data.format && data.format !== "docx-form-values") {
      showBanner("ไฟล์ JSON นี้ไม่ใช่รูปแบบ docx-form-values");
      return;
    }

    if (data.template && state.filename && data.template !== state.filename) {
      const go = await askDialog({
        title: "เทมเพลตไม่ตรงกัน",
        body: `JSON จากไฟล์ «${data.template}» แต่ตอนนี้เปิด «${state.filename}» — นำเข้าเฉพาะฟิลด์ที่ชื่อตรงกันหรือไม่?`,
        actions: [
          { id: "abort", label: "ยกเลิก", kind: "ghost" },
          { id: "ok", label: "นำเข้า", kind: "primary" },
        ],
      });
      if (go !== "ok") return;
    }

    const known = new Set(state.fields.map((f) => f.name));
    const importedValues =
      data.values && typeof data.values === "object" ? data.values : {};
    const importedImages =
      data.images && typeof data.images === "object" ? data.images : {};
    const importedSkip =
      data.skip && typeof data.skip === "object" ? data.skip : {};
    const importedRowSkip =
      data.rowSkip && typeof data.rowSkip === "object" ? data.rowSkip : {};

    for (const [name, val] of Object.entries(importedValues)) {
      if (!known.has(name)) continue;
      state.values[name] = val;
    }

    for (const [name, img] of Object.entries(importedImages)) {
      const field = state.fields.find((f) => f.name === name && f.type === "image");
      if (!field) continue;
      if (!img || !img.base64) {
        state.images[name] = null;
        continue;
      }
      const bytes = base64ToUint8Array(img.base64);
      state.images[name] = new File([bytes], img.name || `${name}.png`, {
        type: img.type || "application/octet-stream",
      });
    }

    for (const f of state.fields) {
      if (f.type === "list") ensureListRows(f.name, f.itemFields);
    }

    const hasOwn = (obj, key) => Object.prototype.hasOwnProperty.call(obj, key);
    /** Empty / missing fill → should tick ข้าม (bool false is a real value). */
    const isBlankFill = (val, type) => {
      if (type === "bool" || type === "auto") return false;
      if (val == null) return true;
      if (typeof val === "string" && !val.trim()) return true;
      return false;
    };

    // Start from exported skip maps (if any)
    for (const [name, val] of Object.entries(importedSkip)) {
      if (!known.has(name)) continue;
      state.skip[name] = !!val;
    }
    for (const [name, rows] of Object.entries(importedRowSkip)) {
      if (!known.has(name) || !Array.isArray(rows)) continue;
      const listField = state.fields.find((f) => f.type === "list" && f.name === name);
      ensureRowSkip(
        name,
        Math.max(rows.length, state.values[name]?.length || 0),
        listField?.itemFields || [],
      );
      rows.forEach((row, idx) => {
        if (!row || typeof row !== "object") return;
        if (!state.rowSkip[name][idx]) state.rowSkip[name][idx] = {};
        for (const [item, skipped] of Object.entries(row)) {
          state.rowSkip[name][idx][item] = !!skipped;
        }
      });
    }

    // Missing keys OR blank strings → force ข้าม (wins over skip:false in JSON)
    for (const f of state.fields) {
      if (f.type === "auto") continue;

      if (f.type === "image") {
        const provided = hasOwn(importedImages, f.name) && !!importedImages[f.name]?.base64;
        if (!provided) {
          state.skip[f.name] = true;
          state.images[f.name] = null;
        }
        continue;
      }

      if (f.type === "list") {
        const rows = hasOwn(importedValues, f.name) ? importedValues[f.name] : null;
        if (!Array.isArray(rows) || rows.length === 0) {
          state.skip[f.name] = true;
          continue;
        }
        ensureRowSkip(f.name, state.values[f.name].length, f.itemFields);
        state.values[f.name].forEach((row, idx) => {
          const srcRow =
            Array.isArray(rows) && rows[idx] && typeof rows[idx] === "object" ? rows[idx] : null;
          for (const item of f.itemFields) {
            if (item.type === "auto") continue;
            const missing = !srcRow || !hasOwn(srcRow, item.name);
            const blank = !missing && isBlankFill(srcRow[item.name], item.type);
            if (missing || blank) {
              state.rowSkip[f.name][idx][item.name] = true;
              if (item.type !== "bool") state.values[f.name][idx][item.name] = "";
            }
          }
        });
        continue;
      }

      // text / bool (top-level)
      const missing = !hasOwn(importedValues, f.name);
      const blank = !missing && isBlankFill(importedValues[f.name], f.type);
      if (missing || blank) {
        state.skip[f.name] = true;
        state.values[f.name] = f.type === "bool" ? false : "";
      }
    }

    renderForm();
    scheduleSaveDraft();
    showBanner(`นำเข้าจาก ${file.name} แล้ว`, "success");
    setTimeout(() => showBanner(""), 2400);
  } catch (err) {
    console.warn(err);
    showBanner("อ่าน JSON ไม่ได้ — ตรวจว่าเป็นไฟล์ที่ Export จาก docx-form");
  }
}

/** Render structured layout from Word (real tables + paragraph text). */
function renderDocumentFromLayout(blocks) {
  if (!blocks?.length) return `<p class="dropzone-hint">ไม่มีข้อความในไฟล์</p>`;
  return blocks
    .map((b) => {
      if (b.type === "text") return renderDocumentLayout(b.text || "");
      if (b.type === "table") return renderWordTableBlock(b);
      return "";
    })
    .join("");
}

function cellTextOf(cell) {
  return typeof cell === "string" ? cell : cell?.text || "";
}

function tableViewKey(block) {
  if (block.listName) return `list:${block.listName}`;
  const sig = (block.colWidths || []).join("x");
  if (sig) return `tbl:${sig}`;
  const n =
    block.templateCells?.length ||
    block.staticRows?.[0]?.length ||
    block.headerRows?.[0]?.reduce((s, c) => s + (c.colspan || 1), 0) ||
    0;
  return `tbl:cols${n}`;
}

function countTableDataCols(block) {
  if (block.colWidths?.length) return block.colWidths.length;
  if (block.templateCells?.length) return block.templateCells.length;
  if (block.staticRows?.[0]?.length) {
    return block.staticRows[0].reduce((s, c) => s + (typeof c === "object" && c.colspan ? c.colspan : 1), 0);
  }
  if (block.headerRows?.[0]?.length) {
    return block.headerRows[0].reduce((s, c) => s + (c.colspan || 1), 0);
  }
  return 0;
}

function layoutCellAttrs(cell, opts = {}) {
  if (typeof cell === "string") return { text: cell, attrs: "" };
  const parts = [];
  const classes = [];
  if (opts.extraClass) classes.push(opts.extraClass);
  if (cell.align) classes.push(`align-${cell.align}`);
  if (classes.length) parts.push(`class="${classes.join(" ")}"`);
  if (cell.colspan && cell.colspan > 1) parts.push(`colspan="${cell.colspan}"`);
  if (cell.rowspan && cell.rowspan > 1) parts.push(`rowspan="${cell.rowspan}"`);
  const styles = [];
  if (cell.widthDxa && cell.widthDxa > 0) {
    // dxa (twips) → px ≈ /15; allow grow past Word width for long input
    const px = Math.max(48, Math.round(cell.widthDxa / 15));
    styles.push(`min-width:${px}px`);
  }
  if (styles.length) parts.push(`style="${styles.join(";")}"`);
  return { text: cell.text || "", attrs: parts.length ? ` ${parts.join(" ")}` : "" };
}

function renderLayoutCellTag(tag, cell, extraClass) {
  // Word headers without explicit w:jc often read centered in narrow columns
  const normalized =
    tag === "th" && typeof cell === "object" && cell && !cell.align
      ? { ...cell, align: "center" }
      : cell;
  const { text, attrs } = layoutCellAttrs(normalized, { extraClass });
  return `<${tag}${attrs}>${formatPlainText(text)}</${tag}>`;
}

function renderColgroup(block, withRowNum) {
  const count = countTableDataCols(block);
  if (!count) return "";
  const key = tableViewKey(block);
  const saved = viewTableColWidths.get(key);
  const wordWidths = block.colWidths || [];
  const cols = [];
  if (withRowNum) {
    const w = saved?.[0];
    cols.push(
      w
        ? `<col class="doc-table-rownum-col" style="width:${w}px;min-width:${w}px" />`
        : `<col class="doc-table-rownum-col" />`,
    );
  }
  for (let i = 0; i < count; i++) {
    const savedIdx = withRowNum ? i + 1 : i;
    const savedW = saved?.[savedIdx];
    const fromWord = wordWidths[i] ? Math.max(48, Math.round(Number(wordWidths[i]) / 15)) : 96;
    const px = savedW || fromWord;
    cols.push(`<col style="width:${px}px;min-width:${px}px" />`);
  }
  return `<colgroup data-view-key="${escapeAttr(key)}">${cols.join("")}</colgroup>`;
}

function renderWordTableBlock(block) {
  if (block.staticRows?.length) {
    const rows = block.staticRows
      .map((r) => `<tr>${r.map((c) => renderLayoutCellTag("td", c)).join("")}</tr>`)
      .join("");
    return `<div class="doc-table-scroll doc-table-real"><table class="doc-table doc-table-stretch">${renderColgroup(block, false)}${rows}</table></div>`;
  }

  const listName = block.listName || "";
  const field = state.fields.find((f) => f.type === "list" && f.name === listName);
  const renderHeaderRows = (withRowNum) =>
    (block.headerRows || [])
      .map((r, ri) => {
        // Row number column only on first header row; rowspan covers the rest
        const headerRows = block.headerRows || [];
        let num = "";
        if (withRowNum && ri === 0) {
          const span = Math.max(headerRows.length, 1);
          num = `<th class="doc-table-rownum"${span > 1 ? ` rowspan="${span}"` : ""}>#</th>`;
        }
        return `<tr>${num}${r.map((c) => renderLayoutCellTag("th", c)).join("")}</tr>`;
      })
      .join("");

  if (!field || field.type !== "list") {
    const tmpl = (block.templateCells || [])
      .map((c) => {
        const { attrs } = layoutCellAttrs(typeof c === "string" ? { text: c } : c);
        return `<td${attrs}>${formatPlainText(cellTextOf(c))}</td>`;
      })
      .join("");
    return `<div class="doc-list doc-list-table" data-list-field="${escapeAttr(listName)}">
      <div class="doc-table-scroll doc-table-real">
        <table class="doc-table doc-table-stretch">
          ${renderColgroup(block, false)}
          <thead>${renderHeaderRows(false)}</thead>
          <tbody><tr>${tmpl}</tr></tbody>
        </table>
      </div>
    </div>`;
  }

  const skipped = !!state.skip[listName];
  ensureListRows(listName, field.itemFields);
  if (skipped) {
    return `<div class="doc-list is-skipped" data-list-field="${escapeAttr(listName)}">
      <div class="doc-list-toolbar">
        <span class="doc-list-name">${escapeHtml(listName)} <span class="badge">ตาราง</span></span>
        <label class="doc-skip">
          <input type="checkbox" data-skip-field="${escapeAttr(listName)}" checked />
          ไม่ประสงค์กรอก
        </label>
      </div>
      <p class="row-empty">ไม่แสดงแถวในเอกสาร</p>
    </div>`;
  }

  const rows = /** @type {Record<string, any>[]} */ (state.values[listName] || []);
  const bodyRows = rows
    .map((row, idx) => {
      const cells = (block.templateCells || [])
        .map((cell) => {
          const cellObj = typeof cell === "string" ? { text: cell } : cell;
          const { attrs } = layoutCellAttrs(cellObj);
          return `<td${attrs}>${renderTableTemplateCell(listName, field, row, idx, cellTextOf(cellObj))}</td>`;
        })
        .join("");
      return `<tr>
        <td class="doc-table-rownum">
          <span>${idx + 1}</span>
          <button type="button" class="linkish danger remove-row" data-list="${escapeAttr(listName)}" data-row="${idx}" title="ลบแถว">ลบ</button>
        </td>
        ${cells}
      </tr>`;
    })
    .join("");

  const colCount = (block.templateCells || []).length + 1;
  return `<div class="doc-list doc-list-table" data-list-field="${escapeAttr(listName)}">
    <div class="doc-list-toolbar">
      <span class="doc-list-name">${escapeHtml(listName)} <span class="badge">ตารางจากไฟล์</span></span>
      <label class="doc-skip">
        <input type="checkbox" data-skip-field="${escapeAttr(listName)}" />
        ไม่ประสงค์กรอก
      </label>
      <button type="button" class="btn-secondary add-row" data-list="${escapeAttr(listName)}">+ เพิ่มแถว</button>
    </div>
    <div class="doc-table-scroll doc-table-real">
      <table class="doc-table doc-table-stretch">
        ${renderColgroup(block, true)}
        <thead>
          ${renderHeaderRows(true) || `<tr><th class="doc-table-rownum">#</th>${(block.templateCells || []).map(() => "<th></th>").join("")}</tr>`}
        </thead>
        <tbody>
          ${bodyRows || `<tr><td colspan="${colCount}" class="row-empty">ยังไม่มีแถว</td></tr>`}
        </tbody>
      </table>
    </div>
  </div>`;
}

function renderTableTemplateCell(listName, field, row, idx, cellText) {
  const tokens = tokenizeTemplate(cellText || "");
  if (!tokens.length) return "";
  return tokens
    .map((t) => {
      if (t.type === "text") {
        const v = String(t.value || "").trim();
        return v ? formatPlainText(v) : "";
      }
      if (t.type === "self") {
        const item = t.name || "";
        const itemField = field.itemFields.find((x) => x.name === item);
        return renderListCellControl(listName, row, idx, item, itemField);
      }
      return "";
    })
    .join("");
}

/**
 * Turn template text (with {{tags}}) into Word-like HTML:
 * plain text stays text; variables become blank inputs.
 */
function renderDocumentLayout(content) {
  const tokens = tokenizeTemplate(content);
  let html = "";
  let i = 0;
  while (i < tokens.length) {
    const tok = tokens[i];
    if (tok.type === "text") {
      html += formatPlainText(tok.value);
      i += 1;
      continue;
    }
    if (tok.type === "open") {
      const listName = tok.name;
      let depth = 1;
      let j = i + 1;
      const inner = [];
      while (j < tokens.length && depth > 0) {
        if (tokens[j].type === "open") depth += 1;
        else if (tokens[j].type === "close") depth -= 1;
        if (depth > 0) inner.push(tokens[j]);
        j += 1;
      }
      html += renderListBlock(listName, inner);
      i = j;
      continue;
    }
    if (tok.type === "close") {
      // stray close — skip
      i += 1;
      continue;
    }
    // self tag
    html += renderInlineBlank(tok.name);
    i += 1;
  }
  return html || `<p class="dropzone-hint">ไม่มีข้อความในไฟล์</p>`;
}

function tokenizeTemplate(content) {
  const re = /\{\{(\/?#?[^}]*)\}\}/g;
  /** @type {{ type: string, value?: string, name?: string }[]} */
  const tokens = [];
  let last = 0;
  let m;
  while ((m = re.exec(content)) !== null) {
    if (m.index > last) {
      tokens.push({ type: "text", value: content.slice(last, m.index) });
    }
    const raw = m[1].trim();
    if (raw.startsWith("#")) {
      tokens.push({ type: "open", name: raw.slice(1).trim() });
    } else if (raw.startsWith("/")) {
      tokens.push({ type: "close", name: raw.slice(1).trim() });
    } else {
      tokens.push({ type: "self", name: raw });
    }
    last = m.index + m[0].length;
  }
  if (last < content.length) {
    tokens.push({ type: "text", value: content.slice(last) });
  }
  return tokens;
}

function formatPlainText(text) {
  return escapeHtml(text)
    .replace(/\n/g, "<br>")
    .replace(/ {2}/g, " &nbsp;");
}

/** Compact bool control: pair √|x when both set; single chip when false side is empty. */
function renderDocBoolControl({ checked, skipped, trueText, falseText, label, inputAttrs }) {
  const tip = `${label || ""} — ติ๊ก → "${trueText}" · ไม่ติ๊ก → "${falseText || "(ว่าง)"}"`;
  const nameText = String(label || "").trim() || "ตัวเลือก";
  const t = String(trueText ?? "");
  const f = String(falseText ?? "");
  const showPair = t !== "" && f !== "";

  let chipsHtml;
  if (showPair) {
    chipsHtml = `<span class="bool-pair" aria-hidden="true" role="presentation">
        <span class="bool-chip bool-chip-true ${checked ? "is-active" : ""}">${escapeHtml(t)}</span>
        <span class="bool-chip bool-chip-false ${checked ? "" : "is-active"}">${escapeHtml(f)}</span>
      </span>`;
  } else {
    const only = t !== "" ? t : f !== "" ? f : "·";
    const active = t !== "" ? checked : !checked;
    chipsHtml = `<span class="bool-chip bool-chip-single ${active ? "is-active" : ""} ${only === "·" ? "is-empty" : ""}" aria-hidden="true">${escapeHtml(only)}</span>`;
  }

  return `<span class="doc-blank-wrap doc-bool-wrap ${skipped ? "is-skipped" : ""} ${checked ? "is-on" : "is-off"}">
    <label class="bool-opt doc-bool-opt" title="${escapeAttr(tip)}">
      <input type="checkbox" ${inputAttrs}
        ${checked ? "checked" : ""} ${skipped ? "disabled" : ""}
        aria-label="${escapeAttr(tip)}" />
      <span class="bool-name">${escapeHtml(nameText)}</span>
      ${chipsHtml}
    </label>
  </span>`;
}

function renderInlineBlank(name) {
  const field = state.fields.find((f) => f.name === name);
  const skipped = !!state.skip[name];

  if (field?.type === "image") {
    return `<span class="doc-blank doc-blank-image" title="${escapeAttr(name)}">[รูป: ${escapeHtml(name)}]</span>`;
  }

  if (field?.type === "bool") {
    const checked = !!state.values[name];
    return renderDocBoolControl({
      checked,
      skipped,
      trueText: field.trueText,
      falseText: field.falseText,
      label: field.label,
      inputAttrs: `data-kind="bool" data-bool-field="${escapeAttr(name)}"`,
    });
  }

  if (field?.type === "auto") {
    return renderAutoRunDisplay(field.label, 0);
  }

  const val = state.values[name] ?? "";
  return `<span class="doc-blank-wrap ${skipped ? "is-skipped" : ""}">
    <input type="text" class="doc-blank" data-kind="text" data-doc-field="${escapeAttr(name)}"
      name="${escapeAttr(name)}" value="${escapeAttr(String(val))}"
      placeholder="${escapeAttr(name)}" ${skipped ? "disabled" : ""}
      aria-label="${escapeAttr(name)}" />
    <label class="doc-skip">
      <input type="checkbox" data-skip-field="${escapeAttr(name)}" ${skipped ? "checked" : ""} />
      ไม่ประสงค์กรอก
    </label>
  </span>`;
}

/** Table-like loop: many adjacent {{tags}} with little prose between them. */
function isTabularList(innerTokens, itemFields) {
  if (!itemFields || itemFields.length < 2) return false;
  const selfCount = innerTokens.filter((t) => t.type === "self").length;
  if (selfCount < 2) return false;
  const prose = innerTokens
    .filter((t) => t.type === "text")
    .map((t) => String(t.value || "").replace(/\s+/g, ""))
    .join("");
  return prose.length <= Math.max(8, selfCount * 2);
}

function itemFieldLabel(itemField) {
  if (!itemField) return "";
  if (itemField.type === "bool" || itemField.type === "auto") {
    return itemField.label || itemField.name;
  }
  return itemField.name;
}

function renderAutoRunDisplay(label, rowIdx) {
  const n = rowIdx + 1;
  return `<span class="auto-run" title="${escapeAttr(`${label || "ลำดับ"} — รันอัตโนมัติ`)}">
    <span class="auto-run-num">${n}</span>
    <span class="badge">รันเอง</span>
  </span>`;
}

function renderListCellControl(listName, row, idx, item, itemField) {
  const itemSkipped = !!state.rowSkip[listName]?.[idx]?.[item];
  if (itemField?.type === "auto") {
    return renderAutoRunDisplay(itemField.label, idx);
  }
  if (itemField?.type === "bool") {
    const checked = !!row[item];
    return renderDocBoolControl({
      checked,
      skipped: itemSkipped,
      trueText: itemField.trueText,
      falseText: itemField.falseText,
      label: itemField.label,
      inputAttrs: `data-kind="bool" data-list="${escapeAttr(listName)}" data-row="${idx}" data-item="${escapeAttr(item)}"`,
    });
  }
  return `<span class="doc-blank-wrap doc-table-cell-wrap ${itemSkipped ? "is-skipped" : ""}">
    <input type="text" class="doc-blank doc-table-input" data-list="${escapeAttr(listName)}" data-row="${idx}" data-item="${escapeAttr(item)}"
      value="${escapeAttr(String(row[item] ?? ""))}" placeholder="${escapeAttr(item)}"
      ${itemSkipped ? "disabled" : ""} aria-label="${escapeAttr(item)}" />
    <label class="doc-skip doc-skip-compact" title="ไม่ประสงค์กรอก">
      <input type="checkbox" data-skip-list="${escapeAttr(listName)}" data-skip-row="${idx}" data-skip-item="${escapeAttr(item)}" ${itemSkipped ? "checked" : ""} />
      <span>ข้าม</span>
    </label>
  </span>`;
}

function renderListAsTable(listName, field, rows) {
  const head = field.itemFields
    .map((item) => `<th scope="col">${escapeHtml(itemFieldLabel(item))}</th>`)
    .join("");
  const body = rows
    .map((row, idx) => {
      const cells = field.itemFields
        .map((itemField) => {
          const item = itemField.name;
          return `<td>${renderListCellControl(listName, row, idx, item, itemField)}</td>`;
        })
        .join("");
      return `<tr>
        <th scope="row" class="doc-table-rownum">
          <span>${idx + 1}</span>
          <button type="button" class="linkish danger remove-row" data-list="${escapeAttr(listName)}" data-row="${idx}" title="ลบแถว">ลบ</button>
        </th>
        ${cells}
      </tr>`;
    })
    .join("");

  return `<div class="doc-table-scroll">
    <table class="doc-table">
      <thead>
        <tr>
          <th scope="col" class="doc-table-rownum">#</th>
          ${head}
        </tr>
      </thead>
      <tbody>
        ${body || `<tr><td colspan="${field.itemFields.length + 1}" class="row-empty">ยังไม่มีแถว</td></tr>`}
      </tbody>
    </table>
  </div>`;
}

function renderListAsCards(listName, field, innerTokens, rows) {
  return rows
    .map((row, idx) => {
      const line = innerTokens
        .map((t) => {
          if (t.type === "text") return formatPlainText(t.value || "");
          if (t.type === "self") {
            const item = t.name || "";
            const itemField = field.itemFields.find((x) => x.name === item);
            return renderListCellControl(listName, row, idx, item, itemField);
          }
          return "";
        })
        .join("");
      return `<div class="doc-list-row">
        <div class="doc-list-row-head">
          <span>แถว ${idx + 1}</span>
          <button type="button" class="linkish danger remove-row" data-list="${escapeAttr(listName)}" data-row="${idx}">ลบแถว</button>
        </div>
        <div class="doc-list-row-body">${line}</div>
      </div>`;
    })
    .join("");
}

function renderListBlock(listName, innerTokens) {
  const field = state.fields.find((f) => f.type === "list" && f.name === listName);
  const skipped = !!state.skip[listName];
  if (!field || field.type !== "list") {
    return innerTokens
      .map((t) => {
        if (t.type === "text") return formatPlainText(t.value || "");
        if (t.type === "self") return renderInlineBlank(t.name || "");
        return "";
      })
      .join("");
  }

  ensureListRows(listName, field.itemFields);
  if (skipped) {
    return `<div class="doc-list is-skipped">
      <div class="doc-list-toolbar">
        <span class="doc-list-name">${escapeHtml(listName)}</span>
        <label class="doc-skip">
          <input type="checkbox" data-skip-field="${escapeAttr(listName)}" checked />
          ไม่ประสงค์กรอก
        </label>
      </div>
      <p class="row-empty">ไม่แสดงแถวในเอกสาร</p>
    </div>`;
  }

  const rows = /** @type {Record<string, string>[]} */ (state.values[listName] || []);
  const asTable = isTabularList(innerTokens, field.itemFields);
  const bodyHtml = asTable
    ? renderListAsTable(listName, field, rows)
    : renderListAsCards(listName, field, innerTokens, rows) ||
      `<p class="row-empty">ยังไม่มีแถว</p>`;

  return `<div class="doc-list ${asTable ? "doc-list-table" : ""}" data-list-field="${escapeAttr(listName)}">
    <div class="doc-list-toolbar">
      <span class="doc-list-name">${escapeHtml(listName)}${asTable ? ` <span class="badge">ตาราง</span>` : ""}</span>
      <label class="doc-skip">
        <input type="checkbox" data-skip-field="${escapeAttr(listName)}" />
        ไม่ประสงค์กรอก
      </label>
      <button type="button" class="btn-secondary add-row" data-list="${escapeAttr(listName)}">+ เพิ่มแถว</button>
    </div>
    ${bodyHtml}
  </div>`;
}

function wireDocumentPreviewEvents() {
  const root = document.getElementById("docPreview");
  if (!root) return;
  // text / list / skip handlers already covered by wireFieldEvents via same data-* attrs
  root.querySelectorAll("table.doc-table-stretch").forEach((table) => wireTableColResize(table));
}

/**
 * Drag column edges to resize in the preview only (not saved to the document).
 * Uses an overlay so every grid line is hittable (cell-edge handles get covered by neighbors).
 * @param {HTMLTableElement} table
 */
function wireTableColResize(table) {
  if (table.dataset.colResizeWired === "1") return;
  const colgroup = table.querySelector(":scope > colgroup");
  const wrap = table.closest(".doc-table-scroll");
  if (!colgroup || !wrap) return;
  table.dataset.colResizeWired = "1";
  table.classList.add("doc-table-resizable");
  table.style.tableLayout = "fixed";
  wrap.classList.add("is-col-resizable");

  const cols = /** @type {HTMLTableColElement[]} */ ([...colgroup.querySelectorAll("col")]);
  if (cols.length < 1) return;
  const key = colgroup.dataset.viewKey || "";

  /** @param {number} [forceIndex] @param {number} [forcePx] */
  const syncTableWidth = (forceIndex, forcePx) => {
    let total = 0;
    cols.forEach((col, idx) => {
      const w =
        forceIndex === idx && forcePx != null
          ? forcePx
          : Math.round(parseFloat(col.style.width) || col.getBoundingClientRect().width || 80);
      total += w;
    });
    table.style.width = `${total}px`;
    table.style.minWidth = `${total}px`;
    table.style.maxWidth = "none";
  };

  /** @param {number} index @param {number} px */
  const applyColWidth = (index, px) => {
    const col = cols[index];
    const w = Math.max(48, Math.round(px));
    col.style.width = `${w}px`;
    col.style.minWidth = `${w}px`;
    col.style.maxWidth = `${w}px`;
    syncTableWidth(index, w);
    return w;
  };

  const lockWidths = () => {
    cols.forEach((col, idx) => {
      const measured = Math.round(col.getBoundingClientRect().width);
      const fallback = Math.round(parseFloat(col.style.width) || parseFloat(col.style.minWidth) || 80);
      applyColWidth(idx, measured > 0 ? measured : fallback);
    });
    syncTableWidth();
  };
  lockWidths();

  const persistViewWidths = () => {
    if (!key) return;
    viewTableColWidths.set(
      key,
      cols.map((col) => Math.round(parseFloat(col.style.width) || col.getBoundingClientRect().width || 80)),
    );
    scheduleSaveDraft();
  };

  /** Prefer a 1-cell-per-column row (body), else expand header leaf row. */
  const measureColRights = () => {
    const wrapRect = wrap.getBoundingClientRect();
    const bodyRow = table.tBodies[0]?.rows[0];
    /** @type {number[]} */
    const rights = [];
    if (bodyRow && bodyRow.cells.length === cols.length) {
      [...bodyRow.cells].forEach((cell) => {
        rights.push(cell.getBoundingClientRect().right - wrapRect.left + wrap.scrollLeft);
      });
      return rights;
    }
    const theadRows = [...(table.tHead?.rows || [])];
    const leafRow = theadRows[theadRows.length - 1] || table.rows[0];
    if (!leafRow) return rights;
    const hasRownum = cols[0]?.classList.contains("doc-table-rownum-col");
    if (hasRownum && theadRows[0]?.cells[0]?.classList.contains("doc-table-rownum")) {
      rights.push(theadRows[0].cells[0].getBoundingClientRect().right - wrapRect.left + wrap.scrollLeft);
    }
    [...leafRow.cells].forEach((cell) => {
      if (cell.classList.contains("doc-table-rownum")) return;
      const span = cell.colSpan || 1;
      const rect = cell.getBoundingClientRect();
      const left = rect.left - wrapRect.left + wrap.scrollLeft;
      const width = rect.width;
      for (let s = 0; s < span; s++) {
        // Approximate equal split when only a colspan cell is available
        rights.push(left + (width * (s + 1)) / span);
      }
    });
    return rights.slice(0, cols.length);
  };

  let overlay = wrap.querySelector(":scope > .col-resize-overlay");
  if (overlay && overlay.dataset.forTable !== key) {
    overlay.remove();
    overlay = null;
  }
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.className = "col-resize-overlay";
    overlay.dataset.forTable = key;
    overlay.setAttribute("aria-hidden", "true");
    wrap.appendChild(overlay);
  }

  const layoutHandles = () => {
    if (wrap.classList.contains("is-resizing-col")) return;
    const rights = measureColRights();
    if (rights.length !== cols.length) {
      // Fall back: cumulative col widths from table left
      const tableLeft =
        table.getBoundingClientRect().left - wrap.getBoundingClientRect().left + wrap.scrollLeft;
      let x = tableLeft;
      rights.length = 0;
      cols.forEach((col) => {
        x += parseFloat(col.style.width) || col.getBoundingClientRect().width || 80;
        rights.push(x);
      });
    }

    // Cover full table height so every vertical grid line is draggable
    overlay.style.height = `${Math.round(table.offsetHeight)}px`;
    overlay.innerHTML = "";

    rights.forEach((right, i) => {
      const handle = document.createElement("div");
      handle.className = "col-resizer";
      handle.dataset.colIndex = String(i);
      handle.title = "ลากปรับความกว้าง (เฉพาะหน้าจอ ไม่บันทึกลงไฟล์)";
      handle.style.left = `${Math.round(right)}px`;
      handle.addEventListener("pointerdown", (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        const startX = e.clientX;
        const startW = parseFloat(cols[i].style.width) || cols[i].getBoundingClientRect().width;
        const startLeft = parseFloat(handle.style.left) || Math.round(right);
        handle.classList.add("is-active");
        table.classList.add("is-resizing-col");
        wrap.classList.add("is-resizing-col");
        handle.setPointerCapture(e.pointerId);

        const onMove = (ev) => {
          const dx = ev.clientX - startX;
          applyColWidth(i, startW + dx);
          // Handle tracks the pointer delta directly (not post-layout measure)
          handle.style.left = `${startLeft + dx}px`;
        };
        const onUp = (ev) => {
          const dx = ev.clientX - startX;
          applyColWidth(i, startW + dx);
          handle.classList.remove("is-active");
          table.classList.remove("is-resizing-col");
          wrap.classList.remove("is-resizing-col");
          window.removeEventListener("pointermove", onMove);
          window.removeEventListener("pointerup", onUp);
          window.removeEventListener("pointercancel", onUp);
          persistViewWidths();
          layoutHandles();
        };
        window.addEventListener("pointermove", onMove);
        window.addEventListener("pointerup", onUp);
        window.addEventListener("pointercancel", onUp);
      });
      overlay.appendChild(handle);
    });
  };

  requestAnimationFrame(() => {
    lockWidths();
    layoutHandles();
  });
  wrap.addEventListener(
    "scroll",
    () => {
      if (!wrap.classList.contains("is-resizing-col")) layoutHandles();
    },
    { passive: true },
  );
  window.addEventListener("resize", () => {
    if (!wrap.classList.contains("is-resizing-col")) layoutHandles();
  });
}

/** @param {FormField} field */
function renderField(field) {
  const skipped = !!state.skip[field.name];

  if (field.type === "text") {
    const val = state.values[field.name] ?? "";
    return `
      <div class="field ${skipped ? "is-skipped" : ""}">
        <div class="field-head">
          <label for="f-${cssId(field.name)}">${escapeHtml(field.name)}</label>
          ${skipCheckbox(field.name, skipped)}
        </div>
        <input type="text" id="f-${cssId(field.name)}" name="${escapeAttr(field.name)}"
          value="${escapeAttr(String(val))}" placeholder="กรอก ${escapeAttr(field.name)}"
          data-kind="text" ${skipped ? "disabled" : ""} />
      </div>`;
  }

  if (field.type === "bool") {
    const checked = !!state.values[field.name];
    return `
      <div class="field ${skipped ? "is-skipped" : ""}">
        <div class="field-head">
          <div class="field-label">${escapeHtml(field.label)} <span class="badge">ติ๊ก</span></div>
          ${skipCheckbox(field.name, skipped)}
        </div>
        <label class="bool-opt bool-opt-block" for="f-${cssId(field.name)}">
          <input type="checkbox" id="f-${cssId(field.name)}" data-kind="bool" data-bool-field="${escapeAttr(field.name)}"
            ${checked ? "checked" : ""} ${skipped ? "disabled" : ""} />
          <span class="bool-chip">${escapeHtml(field.trueText || "·")}</span>
          <span class="bool-hint">ติ๊ก → <code>${escapeHtml(field.trueText || "(ว่าง)")}</code> · ไม่ติ๊ก → <code>${escapeHtml(field.falseText || "(ว่าง)")}</code></span>
        </label>
      </div>`;
  }

  if (field.type === "auto") {
    return `
      <div class="field">
        <div class="field-head">
          <div class="field-label">${escapeHtml(field.label)} <span class="badge">รันเอง</span></div>
        </div>
        <p class="dropzone-hint" style="margin:0">ใช้ในลูปตาราง/รายการ — จะใส่เลขแถวอัตโนมัติ</p>
        ${renderAutoRunDisplay(field.label, 0)}
      </div>`;
  }

  if (field.type === "image") {
    const file = state.images[field.name];
    let preview = "";
    if (file && !skipped) {
      const url = URL.createObjectURL(file);
      state.objectUrls.push(url);
      preview = `<img class="thumb" src="${url}" alt="พรีวิว ${escapeAttr(field.name)}" />`;
    } else {
      preview = `<div class="thumb" aria-hidden="true"></div>`;
    }
    return `
      <div class="field ${skipped ? "is-skipped" : ""}" data-image-field="${escapeAttr(field.name)}">
        <div class="field-head">
          <div class="field-label">${escapeHtml(field.name)} <span class="badge">รูปภาพ</span></div>
          ${skipCheckbox(field.name, skipped)}
        </div>
        <div class="image-box" tabindex="0" role="button" aria-label="อัปโหลดรูป ${escapeAttr(field.name)}" ${skipped ? 'aria-disabled="true"' : ""}>
          ${preview}
          <div class="image-actions">
            <button type="button" class="btn-secondary pick-image" ${skipped ? "disabled" : ""}>เลือกหรือลากรูป PNG/JPEG</button>
            ${file && !skipped ? `<button type="button" class="linkish danger clear-image">ลบรูป</button>` : ""}
          </div>
          <input class="sr-only image-input" type="file" accept="image/png,image/jpeg,image/gif,image/webp,.png,.jpg,.jpeg,.gif,.webp" ${skipped ? "disabled" : ""} />
        </div>
      </div>`;
  }

  // list
  ensureListRows(field.name, field.itemFields);
  const rows = /** @type {Record<string, string>[]} */ (state.values[field.name]);
  const rowsHtml = skipped
    ? `<p class="row-empty">ไม่ประสงค์กรอกรายการนี้ — จะไม่แสดงแถวในเอกสาร</p>`
    : rows.length === 0
      ? `<p class="row-empty">ยังไม่มีแถว — กดเพิ่มแถว หรือติ๊กไม่ประสงค์กรอก</p>`
      : rows
          .map((row, idx) => {
            const inputs = field.itemFields
              .map((item) => {
                const itemSkipped = !!state.rowSkip[field.name]?.[idx]?.[item.name];
                if (item.type === "auto") {
                  return `
              <div class="field" style="margin:0">
                <div class="field-head">
                  <label>${escapeHtml(item.label)} <span class="badge">รันเอง</span></label>
                </div>
                ${renderAutoRunDisplay(item.label, idx)}
              </div>`;
                }
                if (item.type === "bool") {
                  const checked = !!row[item.name];
                  return `
              <div class="field ${itemSkipped ? "is-skipped" : ""}" style="margin:0">
                <div class="field-head">
                  <label for="l-${cssId(field.name)}-${idx}-${cssId(item.name)}">${escapeHtml(item.label)} <span class="badge">ติ๊ก</span></label>
                  ${rowSkipCheckbox(field.name, idx, item.name, itemSkipped)}
                </div>
                <label class="bool-opt" for="l-${cssId(field.name)}-${idx}-${cssId(item.name)}">
                  <input type="checkbox" id="l-${cssId(field.name)}-${idx}-${cssId(item.name)}"
                    data-kind="bool" data-list="${escapeAttr(field.name)}" data-row="${idx}" data-item="${escapeAttr(item.name)}"
                    ${checked ? "checked" : ""} ${itemSkipped ? "disabled" : ""} />
                  <span>จริง → <code>${escapeHtml(item.trueText || "(ว่าง)")}</code> · ไม่จริง → <code>${escapeHtml(item.falseText || "(ว่าง)")}</code></span>
                </label>
              </div>`;
                }
                return `
              <div class="field ${itemSkipped ? "is-skipped" : ""}" style="margin:0">
                <div class="field-head">
                  <label for="l-${cssId(field.name)}-${idx}-${cssId(item.name)}">${escapeHtml(item.name)}</label>
                  ${rowSkipCheckbox(field.name, idx, item.name, itemSkipped)}
                </div>
                <input type="text" id="l-${cssId(field.name)}-${idx}-${cssId(item.name)}"
                  value="${escapeAttr(String(row[item.name] ?? ""))}"
                  data-list="${escapeAttr(field.name)}" data-row="${idx}" data-item="${escapeAttr(item.name)}"
                  ${itemSkipped ? "disabled" : ""} />
              </div>`;
              })
              .join("");
            return `
              <div class="row-card" aria-label="แถว ${idx + 1} ของ ${escapeAttr(field.name)}">
                <div class="row-card-head">
                  <span>แถว ${idx + 1}</span>
                  <button type="button" class="linkish danger remove-row" data-list="${escapeAttr(field.name)}" data-row="${idx}">ลบแถว</button>
                </div>
                <div class="row-grid">${inputs}</div>
              </div>`;
          })
          .join("");

  return `
    <div class="repeater ${skipped ? "is-skipped" : ""}" data-list-field="${escapeAttr(field.name)}">
      <div class="repeater-head">
        <h3>${escapeHtml(field.name)} <span class="badge">รายการ</span></h3>
        <div class="repeater-actions">
          ${skipCheckbox(field.name, skipped)}
          <button type="button" class="btn-secondary add-row" data-list="${escapeAttr(field.name)}" ${skipped ? "disabled" : ""}>+ เพิ่มแถว</button>
        </div>
      </div>
      ${rowsHtml}
    </div>`;
}

function wireFieldEvents() {
  app.querySelectorAll("input[data-skip-field]").forEach((el) => {
    el.addEventListener("change", (e) => {
      const t = /** @type {HTMLInputElement} */ (e.target);
      const name = t.dataset.skipField;
      if (!name) return;
      state.skip[name] = t.checked;
      if (t.checked) {
        const field = state.fields.find((f) => f.name === name);
        if (field?.type === "text") state.values[name] = "";
        if (field?.type === "bool") state.values[name] = false;
        if (field?.type === "image") state.images[name] = null;
      }
      renderForm();
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll("input[data-skip-list]").forEach((el) => {
    el.addEventListener("change", (e) => {
      const t = /** @type {HTMLInputElement} */ (e.target);
      const list = t.dataset.skipList;
      const row = Number(t.dataset.skipRow);
      const item = t.dataset.skipItem;
      if (!list || item == null || Number.isNaN(row)) return;
      if (!state.rowSkip[list]) state.rowSkip[list] = [];
      if (!state.rowSkip[list][row]) state.rowSkip[list][row] = {};
      state.rowSkip[list][row][item] = t.checked;
      if (t.checked && state.values[list]?.[row]) {
        const listField = state.fields.find((f) => f.type === "list" && f.name === list);
        const itemField = listField?.type === "list"
          ? listField.itemFields.find((x) => x.name === item)
          : null;
        state.values[list][row][item] = itemField?.type === "bool" ? false : "";
      }
      renderForm();
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll('input[data-kind="bool"]').forEach((el) => {
    el.addEventListener("change", (e) => {
      const t = /** @type {HTMLInputElement} */ (e.target);
      const wrap = t.closest(".doc-bool-wrap");
      if (wrap) {
        wrap.classList.toggle("is-on", t.checked);
        wrap.classList.toggle("is-off", !t.checked);
        const trueChip = wrap.querySelector(".bool-chip-true");
        const falseChip = wrap.querySelector(".bool-chip-false");
        const singleChip = wrap.querySelector(".bool-chip-single");
        trueChip?.classList.toggle("is-active", t.checked);
        falseChip?.classList.toggle("is-active", !t.checked);
        // single-chip mode (e.g. {{พบ?√,}}) — highlight only when checked
        singleChip?.classList.toggle("is-active", t.checked);
      }
      const list = t.dataset.list;
      if (list != null) {
        const row = Number(t.dataset.row);
        const item = t.dataset.item;
        if (item == null || Number.isNaN(row)) return;
        if (!state.values[list]) state.values[list] = [];
        if (!state.values[list][row]) state.values[list][row] = {};
        state.values[list][row][item] = t.checked;
        scheduleSaveDraft();
        return;
      }
      const name = t.dataset.boolField;
      if (!name) return;
      state.values[name] = t.checked;
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll('input[data-kind="text"]').forEach((el) => {
    el.addEventListener("input", (e) => {
      const t = /** @type {HTMLInputElement} */ (e.target);
      const name = t.name || t.dataset.docField;
      if (!name) return;
      state.values[name] = t.value;
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll("input[data-list]:not([data-kind='bool'])").forEach((el) => {
    el.addEventListener("input", (e) => {
      const t = /** @type {HTMLInputElement} */ (e.target);
      const list = t.dataset.list;
      const row = Number(t.dataset.row);
      const item = t.dataset.item;
      if (!list || item == null || Number.isNaN(row)) return;
      if (!state.values[list][row]) state.values[list][row] = {};
      state.values[list][row][item] = t.value;
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll(".add-row").forEach((btn) => {
    btn.addEventListener("click", () => {
      const list = /** @type {HTMLElement} */ (btn).dataset.list;
      const field = state.fields.find((f) => f.type === "list" && f.name === list);
      if (!field || field.type !== "list" || state.skip[list]) return;
      const blank = {};
      for (const item of field.itemFields) {
        if (item.type === "bool") blank[item.name] = false;
        else blank[item.name] = "";
      }
      state.values[list].push(blank);
      ensureRowSkip(list, state.values[list].length, field.itemFields);
      renderForm();
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll(".remove-row").forEach((btn) => {
    btn.addEventListener("click", () => {
      const el = /** @type {HTMLElement} */ (btn);
      const list = el.dataset.list;
      const row = Number(el.dataset.row);
      if (!list || Number.isNaN(row)) return;
      state.values[list].splice(row, 1);
      state.rowSkip[list]?.splice(row, 1);
      renderForm();
      scheduleSaveDraft();
    });
  });

  app.querySelectorAll("[data-image-field]").forEach((wrap) => {
    const name = /** @type {HTMLElement} */ (wrap).dataset.imageField;
    const box = wrap.querySelector(".image-box");
    const input = /** @type {HTMLInputElement | null} */ (wrap.querySelector(".image-input"));
    const pick = wrap.querySelector(".pick-image");
    const clear = wrap.querySelector(".clear-image");
    if (!name || !box || !input || state.skip[name]) return;

    const assign = (file) => {
      if (!file) return;
      if (file.size > 8 * 1024 * 1024) {
        showBanner(`รูป ${name} ใหญ่เกิน 8 MB`);
        return;
      }
      state.images[name] = file;
      showBanner("");
      renderForm();
      scheduleSaveDraft();
    };

    pick?.addEventListener("click", (e) => {
      e.stopPropagation();
      input.click();
    });
    box.addEventListener("click", (e) => {
      if (e.target === clear) return;
      input.click();
    });
    input.addEventListener("change", () => assign(input.files?.[0] || null));
    clear?.addEventListener("click", (e) => {
      e.stopPropagation();
      state.images[name] = null;
      renderForm();
      scheduleSaveDraft();
    });

    ["dragenter", "dragover"].forEach((ev) => {
      box.addEventListener(ev, (e) => {
        e.preventDefault();
        box.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach((ev) => {
      box.addEventListener(ev, (e) => {
        e.preventDefault();
        box.classList.remove("dragover");
      });
    });
    box.addEventListener("drop", (e) => {
      assign(e.dataTransfer?.files?.[0] || null);
    });
  });
}

/** Build values object for the API (skipped fields become empty / []). */
function buildValuesPayload() {
  /** @type {Record<string, unknown>} */
  const values = {};

  for (const f of state.fields) {
    if (f.type === "text") {
      values[f.name] = state.skip[f.name]
        ? ""
        : String(state.values[f.name] ?? "");
    } else if (f.type === "bool") {
      const on = !state.skip[f.name] && !!state.values[f.name];
      values[f.name] = on ? f.trueText : f.falseText;
    } else if (f.type === "auto") {
      values[f.name] = "1";
    } else if (f.type === "image") {
      // Image binary sent separately; empty string clears the tag when skipped / no file
      if (state.skip[f.name] || !state.images[f.name]) {
        values[f.name] = "";
      }
    } else if (f.type === "list") {
      if (state.skip[f.name]) {
        values[f.name] = [];
      } else {
        const rows = (state.values[f.name] || []).map((row, idx) => {
          /** @type {Record<string, string>} */
          const out = {};
          for (const item of f.itemFields) {
            if (item.type === "auto") {
              out[item.name] = String(idx + 1);
            } else if (item.type === "bool") {
              const on =
                !state.rowSkip[f.name]?.[idx]?.[item.name] && !!row[item.name];
              out[item.name] = on ? item.trueText : item.falseText;
            } else {
              out[item.name] = state.rowSkip[f.name]?.[idx]?.[item.name]
                ? ""
                : String(row[item.name] ?? "");
            }
          }
          return out;
        });
        values[f.name] = rows;
      }
    }
  }

  return values;
}

async function onSubmit(e) {
  e.preventDefault();
  if (!state.file) return;

  const missing = validateBeforeSubmit();
  if (missing.length) {
    const listed =
      missing.length <= 10
        ? missing.join(", ")
        : `${missing.slice(0, 10).join(", ")} และอีก ${missing.length - 10} ช่อง`;
    showBanner(
      `ยังมีช่องที่ไม่ได้กรอก (${missing.length}): ${listed} — กรอกให้ครบ หรือติ๊ก “ไม่ประสงค์กรอก” ก่อนสร้างเอกสาร`,
    );
    return;
  }

  const choice = await askClearAfterCreate();
  if (choice === "abort") return;

  const btn = /** @type {HTMLButtonElement | null} */ (document.getElementById("submitBtn"));
  if (btn) {
    btn.disabled = true;
    btn.textContent = "กำลังสร้าง…";
  }
  showBanner("");

  const values = buildValuesPayload();
  const fd = new FormData();
  fd.append("file", state.file);
  fd.append("values", JSON.stringify(values));
  for (const [name, file] of Object.entries(state.images)) {
    if (file && !state.skip[name]) {
      fd.append(`image:${name}`, file, file.name);
    }
  }

  try {
    const res = await fetch("/api/fill", { method: "POST", body: fd });
    if (!res.ok) {
      let msg = "สร้างเอกสารไม่สำเร็จ";
      try {
        const data = await res.json();
        if (data.error) msg = data.error;
      } catch {
        /* ignore */
      }
      showBanner(msg);
      return;
    }
    const blob = await res.blob();
    const cd = res.headers.get("Content-Disposition") || "";
    const star = /filename\*\s*=\s*UTF-8''([^;]+)/i.exec(cd);
    const plain = /filename="([^"]+)"/.exec(cd);
    const filename = star
      ? decodeURIComponent(star[1].trim())
      : plain?.[1] || "filled.docx";
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    if (choice === "clear") {
      clearFilledFields();
      renderForm();
      showBanner("ดาวน์โหลดแล้ว และล้างช่องเรียบร้อย", "success");
    } else {
      scheduleSaveDraft();
      showBanner("ดาวน์โหลดเอกสารแล้ว — เก็บค่าที่กรอกไว้", "success");
    }
    setTimeout(() => showBanner(""), 2800);
  } catch {
    showBanner("เชื่อมต่อเซิร์ฟเวอร์ไม่ได้");
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "สร้างเอกสาร";
    }
  }
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeAttr(s) {
  return escapeHtml(s).replace(/'/g, "&#39;");
}

function cssId(s) {
  return String(s).replace(/[^a-zA-Z0-9_-]/g, "_");
}

async function boot() {
  await checkHealth();
  const restored = await restoreDraft();
  if (restored) {
    renderForm();
    showBanner("กู้คืนข้อมูลหลังรีเฟรชแล้ว", "success");
    setTimeout(() => showBanner(""), 2800);
  } else {
    renderEmpty();
  }
}

window.addEventListener("pagehide", () => {
  window.clearTimeout(draftSaveTimer);
  void saveDraftNow();
});
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") {
    window.clearTimeout(draftSaveTimer);
    void saveDraftNow();
  }
});

void boot();
setInterval(checkHealth, 30000);
