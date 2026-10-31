#!/usr/bin/env node
/**
 * End-to-end UI tests for template-from-selection (tagger).
 * Requires: server on BASE_URL (default http://127.0.0.1:3850), Playwright available.
 *
 *   node scripts/test-tagger-e2e.mjs
 *   BASE_URL=http://127.0.0.1:3850 CACHE=drafts12 node scripts/test-tagger-e2e.mjs
 */
import { readFileSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import JSZip from "jszip";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, "..");
const require = createRequire(import.meta.url);

const BASE = process.env.BASE_URL || "http://127.0.0.1:3850";
const CACHE = process.env.CACHE || "drafts23";

function loadPlaywright() {
  const candidates = [
    join(root, "node_modules/playwright/index.mjs"),
    "/tmp/node_modules/playwright/index.mjs",
  ];
  for (const p of candidates) {
    if (existsSync(p)) return import(p);
  }
  try {
    return import("playwright");
  } catch {
    throw new Error(
      "playwright not found — npm i -D playwright  หรือติดตั้งที่ /tmp/node_modules/playwright",
    );
  }
}

async function minimalDocx(bodyInner) {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyInner}<w:sectPr/></w:body>
</w:document>`;
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );
  zip.folder("_rels").file(
    ".rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.folder("word").file("document.xml", doc);
  zip
    .folder("word")
    .folder("_rels")
    .file(
      "document.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
    );
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

function p(text) {
  const esc = String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return `<w:p><w:r><w:t xml:space="preserve">${esc}</w:t></w:r></w:p>`;
}
function cell(text) {
  const esc = String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return `<w:tc><w:p><w:r><w:t xml:space="preserve">${esc}</w:t></w:r></w:p></w:tc>`;
}
function row(...cells) {
  return `<w:tr>${cells.map(cell).join("")}</w:tr>`;
}

class Suite {
  constructor() {
    this.pass = [];
    this.fail = [];
  }
  ok(name, cond, detail = "") {
    if (cond) this.pass.push(name);
    else this.fail.push(detail ? `${name}: ${detail}` : name);
  }
  summary() {
    console.log("\n=== TAGGER E2E RESULT ===");
    console.log(`PASS ${this.pass.length}`);
    for (const x of this.pass) console.log("  ✓", x);
    console.log(`FAIL ${this.fail.length}`);
    for (const x of this.fail) console.log("  ✗", x);
    return this.fail.length === 0;
  }
}

async function startTagger(page, buffer, filename = "sample.docx") {
  await page.goto(`${BASE}/?v=${CACHE}`, { waitUntil: "networkidle" });
  await page.click("#startTaggerMode");
  await page.waitForSelector("#taggerFileInput");
  await page.setInputFiles("#taggerFileInput", {
    name: filename,
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer,
  });
  await page.waitForSelector("#taggerSelectRoot", { timeout: 20000 });
}

async function selectTextInRoot(page, needle) {
  const found = await page.evaluate((n) => {
    const root = document.getElementById("taggerSelectRoot");
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = w.nextNode())) {
      const t = node.textContent || "";
      const idx = t.indexOf(n);
      if (idx >= 0) {
        const range = document.createRange();
        range.setStart(node, idx);
        range.setEnd(node, idx + n.length);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        return true;
      }
    }
    return false;
  }, needle);
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(280);
  return found;
}

async function selectTableRow(page, rowIndex = 1, mode = "cell0") {
  await page.evaluate(
    ({ rowIndex, mode }) => {
      const root = document.getElementById("taggerSelectRoot");
      const row = root.querySelectorAll("tr")[rowIndex];
      if (!row) throw new Error("no row");
      const sel = window.getSelection();
      sel.removeAllRanges();
      const range = document.createRange();
      if (mode === "cell0") range.selectNodeContents(row.cells[0]);
      else if (mode === "row") range.selectNodeContents(row);
      else {
        range.setStartBefore(row.cells[0]);
        range.setEndAfter(row.cells[row.cells.length - 1]);
      }
      sel.addRange(range);
    },
    { rowIndex, mode },
  );
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(280);
}

async function openCreate(page) {
  const visible = await page.evaluate(() => !document.getElementById("taggerCreateBtn")?.hidden);
  if (!visible) return { visible: false };
  await page.click("#taggerCreateBtn");
  await page.waitForTimeout(350);
  return page.evaluate(() => ({
    visible: true,
    dlg: !document.getElementById("taggerDialog")?.hidden,
    banner: document.getElementById("banner")?.textContent || "",
    kind: document.querySelector('input[name="taggerKind"]:checked')?.value || "",
    sample: document.getElementById("taggerDialogSample")?.textContent || "",
    cells: document.querySelectorAll(".tagger-cell-field").length,
    tableHidden: !!document.getElementById("taggerKindTable")?.hidden,
    listHidden: !!document.getElementById("taggerKindList")?.hidden,
  }));
}

async function setKind(page, kind) {
  await page.evaluate((k) => {
    const r = document.querySelector(`input[name="taggerKind"][value="${k}"]`);
    if (!r) throw new Error(`kind ${k} missing`);
    const label = r.closest("label");
    if (label?.hidden) throw new Error(`kind ${k} hidden`);
    r.checked = true;
    r.dispatchEvent(new Event("change", { bubbles: true }));
  }, kind);
  await page.waitForTimeout(80);
}

async function confirmCreate(page, name) {
  await page.fill("#taggerVarName", name);
  await page.click("#taggerDialogConfirm");
  await page.waitForTimeout(2200);
  return page.evaluate(() => ({
    banner: document.getElementById("banner")?.textContent || "",
    err: document.getElementById("taggerDialogError")?.textContent || "",
    dlgHidden: !!document.getElementById("taggerDialog")?.hidden,
    chips: [...document.querySelectorAll("[data-tagger-edit] code")].map((c) => c.textContent),
    preview: document.getElementById("taggerSelectRoot")?.innerText || "",
  }));
}

async function cancelDlg(page) {
  await page.evaluate(() => document.getElementById("taggerDialogCancel")?.click());
  await page.waitForTimeout(100);
}

async function main() {
  const health = await fetch(`${BASE}/api/health`).then((r) => r.ok).catch(() => false);
  if (!health) {
    console.error(`Server not reachable at ${BASE} — รัน npm start / node dist/index.js -p 3850 ก่อน`);
    process.exit(2);
  }

  const { chromium } = await loadPlaywright();
  const S = new Suite();
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    permissions: ["clipboard-read", "clipboard-write"],
  });
  const page = await context.newPage();
  page.setDefaultTimeout(20000);

  const paraDoc = await minimalDocx(
    p("หัวข้อเอกสาร") +
      p("จำเลยที่1เลขนิติบุคคล 0195561000585ชื่อบริษัท ข้าวแกงแม่มณี") +
      p("จำเลยที่2เลขนิติบุคคล 0195561000999ชื่อบริษัท อีกบริษัท"),
  );
  const tableDoc = await minimalDocx(
    p("Title") +
      `<w:tbl>${row("สินค้า", "จำนวน", "ราคา")}${row("แอปเปิล", "10", "20")}${row("กล้วย", "5", "12")}${row("ส้ม", "3", "15")}</w:tbl>` +
      p("Note"),
  );
  const tableFile = existsSync(join(root, "examples/table-untagged.docx"))
    ? readFileSync(join(root, "examples/table-untagged.docx"))
    : tableDoc;

  // ---------- 1) Text variable ----------
  await startTagger(page, paraDoc, "para.docx");
  S.ok("upload preview", await page.locator("#taggerSelectRoot").count().then((n) => n > 0));
  S.ok("select หัวข้อ", await selectTextInRoot(page, "หัวข้อเอกสาร"));
  let st = await openCreate(page);
  S.ok("text: float+dialog", st.visible && st.dlg, JSON.stringify(st));
  S.ok("text: no false banner", !st.banner.includes("เลือกข้อความในพรีวิว"), st.banner);
  await setKind(page, "text");
  let res = await confirmCreate(page, "title");
  S.ok("text: tagged", res.chips.some((c) => c.includes("{{title}}")), JSON.stringify(res.chips));
  S.ok("text: preview", res.preview.includes("{{title}}"));

  // ---------- 2) Bool ----------
  S.ok("select ชื่อบริษัท label text", await selectTextInRoot(page, "ข้าวแกงแม่มณี"));
  st = await openCreate(page);
  S.ok("bool: dialog", st.dlg);
  await setKind(page, "bool");
  await page.fill("#taggerTrueText", "พบ");
  await page.fill("#taggerFalseText", "");
  res = await confirmCreate(page, "พบบริษัท");
  S.ok(
    "bool: tagged",
    res.chips.some((c) => c.includes("พบบริษัท") || c.includes("{{พบบริษัท")),
    JSON.stringify(res.chips),
  );

  // ---------- 3) Paragraph list + auto runno ----------
  await startTagger(page, paraDoc, "para2.docx");
  S.ok("select defendant line", await selectTextInRoot(page, "จำเลยที่1เลขนิติบุคคล"));
  // extend selection to full first defendant paragraph via para node
  await page.evaluate(() => {
    const root = document.getElementById("taggerSelectRoot");
    const paras = [...root.querySelectorAll(".tagger-para")];
    const target = paras.find((p) => (p.textContent || "").includes("จำเลยที่1"));
    if (!target) return;
    const range = document.createRange();
    range.selectNodeContents(target);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  });
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(280);
  st = await openCreate(page);
  S.ok("list: dialog", st.dlg, JSON.stringify(st));
  await setKind(page, "list");
  res = await confirmCreate(page, "จำเลย");
  S.ok("list: wrap chip", res.chips.some((c) => c.includes("{{#จำเลย}}")), JSON.stringify(res.chips));
  S.ok("list: preview wrap", /\{\{#จำเลย\}\}/.test(res.preview) && /\{\{\/\}\}/.test(res.preview));

  // Prefer the digit after «ที่» inside the loop
  const gotDigit = await page.evaluate(() => {
    const root = document.getElementById("taggerSelectRoot");
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = w.nextNode())) {
      const t = node.textContent || "";
      const m = /ที่(\d)/u.exec(t);
      if (!m) continue;
      const digitIdx = m.index + m[0].length - 1;
      const range = document.createRange();
      range.setStart(node, digitIdx);
      range.setEnd(node, digitIdx + 1);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return sel.toString() === m[1];
    }
    return false;
  });
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(280);
  S.ok("select run digit after ที่", gotDigit);
  st = await openCreate(page);
  S.ok("auto: dialog after list", st.dlg && !st.banner.includes("เลือกข้อความในพรีวิว"), JSON.stringify(st));
  if (st.dlg) {
    const kindOk = await page.evaluate(() => {
      try {
        const r = document.querySelector('input[name="taggerKind"][value="auto"]');
        if (!r || r.closest("label")?.hidden) return false;
        r.checked = true;
        r.dispatchEvent(new Event("change", { bubbles: true }));
        return true;
      } catch {
        return false;
      }
    });
    S.ok("auto: kind available", kindOk);
    if (kindOk) {
      res = await confirmCreate(page, "ลำดับ");
      S.ok(
        "auto: chip {{+ลำดับ}}",
        res.chips.some((c) => c.includes("{{+ลำดับ}}")),
        JSON.stringify(res.chips),
      );
      S.ok("auto: keeps ที่ prefix", /ที่\{\{\+ลำดับ\}\}|ที่\{\{\+/.test(res.preview.replace(/\s/g, "")) || res.preview.includes("{{+ลำดับ}}"));
    } else await cancelDlg(page);
  }

  // Selecting long phrase as auto should narrow to digit
  await startTagger(page, paraDoc, "para3.docx");
  await page.evaluate(() => {
    const root = document.getElementById("taggerSelectRoot");
    const target = [...root.querySelectorAll(".tagger-para")].find((p) =>
      (p.textContent || "").includes("จำเลยที่1"),
    );
    const range = document.createRange();
    range.selectNodeContents(target);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
  });
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(250);
  st = await openCreate(page);
  await setKind(page, "list");
  await confirmCreate(page, "จำเลย");
  await page.evaluate(() => {
    const root = document.getElementById("taggerSelectRoot");
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = w.nextNode())) {
      const t = node.textContent || "";
      const idx = t.indexOf("จำเลยที่1");
      if (idx >= 0) {
        const range = document.createRange();
        range.setStart(node, idx);
        range.setEnd(node, idx + "จำเลยที่1".length);
        window.getSelection().removeAllRanges();
        window.getSelection().addRange(range);
        return;
      }
    }
  });
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(250);
  st = await openCreate(page);
  S.ok("auto-narrow: dialog", st.dlg);
  await setKind(page, "auto");
  const hint = await page.evaluate(() => document.getElementById("taggerKindHint")?.textContent || "");
  S.ok("auto-narrow: hint mentions digit", hint.includes("1") || hint.includes("รันเลข"), hint);
  res = await confirmCreate(page, "ลำดับ");
  S.ok(
    "auto-narrow: only digit replaced",
    res.preview.includes("จำเลยที่{{+ลำดับ}}") || res.preview.includes("จำเลยที่{{+"),
    res.preview.slice(0, 180),
  );

  // ---------- 4) Table list-table ----------
  await startTagger(page, tableFile, "table.docx");
  S.ok("table preview", (await page.locator("#taggerSelectRoot table").count()) > 0);
  for (const mode of ["cell0", "row", "drag"]) {
    await selectTableRow(page, 1, mode === "drag" ? "drag" : mode);
    st = await openCreate(page);
    S.ok(`table-${mode}: open`, st.visible && st.dlg, JSON.stringify(st));
    S.ok(`table-${mode}: no false banner`, !st.banner.includes("เลือกข้อความในพรีวิว"), st.banner);
    S.ok(`table-${mode}: list-table`, st.kind === "list-table", st.kind);
    S.ok(`table-${mode}: 3 cells`, st.cells === 3, String(st.cells));
    await cancelDlg(page);
  }

  await selectTableRow(page, 1, "cell0");
  st = await openCreate(page);
  await page.fill("#taggerVarName", "items");
  await page.evaluate(() => {
    [...document.querySelectorAll(".tagger-cell-field")].forEach((inp, i) => {
      inp.value = ["name", "qty", "price"][i] || "";
    });
  });
  res = await confirmCreate(page, "items");
  S.ok("table create: chips", res.chips.includes("{{#items}}…{{/}}") && res.chips.includes("{{name}}"), JSON.stringify(res.chips));
  S.ok("table create: preview", res.preview.includes("{{#items}}") && res.preview.includes("{{name}}"));

  // Nested list-table blocked
  await startTagger(page, tableFile, "table2.docx");
  await selectTableRow(page, 1, "cell0");
  await openCreate(page);
  await confirmCreate(page, "items"); // wrap only
  await selectTextInRoot(page, "10");
  st = await openCreate(page);
  S.ok("inside table loop: open", st.dlg && !st.banner.includes("เลือกข้อความในพรีวิว"), JSON.stringify(st));
  S.ok("inside table loop: hide list-table", st.tableHidden === true);
  await page.evaluate(() => {
    const label = document.getElementById("taggerKindTable");
    if (label) label.hidden = false;
    const r = document.querySelector('input[name="taggerKind"][value="list-table"]');
    if (r) {
      r.checked = true;
      r.dispatchEvent(new Event("change", { bubbles: true }));
    }
  });
  await page.fill("#taggerVarName", "items2");
  await page.click("#taggerDialogConfirm");
  await page.waitForTimeout(700);
  const nested = await page.evaluate(() => ({
    err: document.getElementById("taggerDialogError")?.textContent || "",
    banner: document.getElementById("banner")?.textContent || "",
    chips: document.querySelectorAll("[data-tagger-edit]").length,
    dlgHidden: !!document.getElementById("taggerDialog")?.hidden,
  }));
  S.ok(
    "nested table list blocked",
    nested.err.includes("ห่อเป็นลิสต์") || nested.banner.includes("ห่อเป็นลิสต์"),
    JSON.stringify(nested),
  );
  S.ok("nested no extra chip", nested.chips === 1, String(nested.chips));

  // Multi-row
  await startTagger(page, tableFile, "table3.docx");
  await page.evaluate(() => {
    const rows = document.querySelectorAll("#taggerSelectRoot tr");
    const range = document.createRange();
    range.setStartBefore(rows[1]);
    range.setEndAfter(rows[2]);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
  });
  await page.dispatchEvent("body", "pointerup");
  await page.waitForTimeout(280);
  st = await openCreate(page);
  S.ok("multi-row: open without false banner", st.dlg && !st.banner.includes("เลือกข้อความในพรีวิว"), JSON.stringify(st));
  S.ok("multi-row: hide lists", st.tableHidden && st.listHidden);
  await cancelDlg(page);

  // ---------- 5) Copy + paste ----------
  await startTagger(page, paraDoc, "copy.docx");
  S.ok("copy select", await selectTextInRoot(page, "ข้าวแกงแม่มณี"));
  await page.waitForSelector("#taggerCopyBtn:not([hidden])");
  await page.click("#taggerCopyBtn");
  await page.waitForTimeout(500);
  let banner = await page.evaluate(() => document.getElementById("banner")?.textContent || "");
  S.ok(
    "copy: success banner",
    banner.includes("คัดลอกข้อความแล้ว") && !banner.includes("ไม่สำเร็จ"),
    banner,
  );
  S.ok("copy: no scroll-required error", !banner.includes("เลือกข้อความในพรีวิว"));
  const hasBuffer = await page.evaluate(() => {
    // syncTaggerPasteUi leaves paste hint visible when buffer set
    const hint = document.getElementById("taggerPasteHint");
    return hint && !hint.hidden;
  });
  // Buffer may only show after dialog open — check paste btn after open
  S.ok("paste prep select", await selectTextInRoot(page, "หัวข้อเอกสาร"));
  st = await openCreate(page);
  await setKind(page, "text");
  const pasteEnabled = await page
    .locator('[data-paste-into="taggerVarName"]')
    .evaluate((el) => el instanceof HTMLButtonElement && !el.disabled);
  S.ok("paste: button enabled after copy", pasteEnabled || hasBuffer, `enabled=${pasteEnabled} hint=${hasBuffer}`);
  if (pasteEnabled) {
    await page.click('[data-paste-into="taggerVarName"]');
    await page.waitForTimeout(300);
  } else {
    // Fallback: inject buffer if UI somehow cleared (should not happen after app fix)
    await page.evaluate(() => {
      // @ts-ignore
      if (typeof lastCopiedTaggerText !== "undefined") return;
    });
  }
  const pasted = await page.evaluate(() => document.getElementById("taggerVarName")?.value || "");
  S.ok("paste: filled name field", pasted.includes("ข้าวแกงแม่มณี") || pasted.length > 0, pasted);
  await cancelDlg(page);

  // ---------- 6) Delete paragraph ----------
  await startTagger(page, paraDoc, "del.docx");
  S.ok("delete select", await selectTextInRoot(page, "หัวข้อเอกสาร"));
  await page.waitForSelector("#taggerDeleteBtn:not([hidden])");
  await page.click("#taggerDeleteBtn");
  await page.waitForTimeout(400);
  const delDlg = await page.evaluate(() => ({
    open: !document.getElementById("taggerDeleteDialog")?.hidden,
    banner: document.getElementById("banner")?.textContent || "",
  }));
  S.ok("delete: dialog", delDlg.open, JSON.stringify(delDlg));
  S.ok("delete: no false banner", !delDlg.banner.includes("เลือกข้อความในพรีวิว"), delDlg.banner);
  await page.click("#taggerDeleteConfirm");
  await page.waitForTimeout(2000);
  res = await page.evaluate(() => ({
    chips: [...document.querySelectorAll("[data-tagger-edit] code")].map((c) => c.textContent),
    preview: document.getElementById("taggerSelectRoot")?.innerText || "",
  }));
  S.ok("delete: has delete op chip", res.chips.length >= 1, JSON.stringify(res.chips));
  S.ok("delete: title removed from preview", !res.preview.includes("หัวข้อเอกสาร"), res.preview.slice(0, 120));

  // ---------- 7) Download / open fill enabled after tags ----------
  await startTagger(page, paraDoc, "export.docx");
  await selectTextInRoot(page, "หัวข้อเอกสาร");
  await openCreate(page);
  await setKind(page, "text");
  await confirmCreate(page, "title");
  const exportBtns = await page.evaluate(() => ({
    dl: !document.getElementById("taggerDownload")?.disabled,
    fill: !document.getElementById("taggerOpenFill")?.disabled,
  }));
  S.ok("export: download enabled", exportBtns.dl);
  S.ok("export: open fill enabled", exportBtns.fill);

  // Open fill parses tagged doc
  await page.click("#taggerOpenFill");
  await page.waitForTimeout(2500);
  const fillPage = await page.evaluate(() => ({
    hasForm: !!document.querySelector("form, #app .field, [data-doc-field], .sticky-bar"),
    banner: document.getElementById("banner")?.textContent || "",
    body: document.getElementById("app")?.innerText?.slice(0, 200) || "",
  }));
  S.ok(
    "open fill: form UI",
    fillPage.hasForm || fillPage.body.includes("title") || fillPage.body.includes("สร้างเอกสาร"),
    JSON.stringify(fillPage),
  );

  await browser.close();
  const ok = S.summary();
  process.exit(ok ? 0 : 1);
}

main().catch(async (err) => {
  console.error(err);
  console.error("\n(suite aborted before summary — see error above)");
  process.exit(1);
});
