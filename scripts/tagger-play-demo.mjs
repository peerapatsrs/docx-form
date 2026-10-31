#!/usr/bin/env node
/**
 * Visual "play" demo — Chromium เปิดบนจอจริง คลิก flow สร้างเทมเพลตให้ดูทีละขั้น
 *
 *   npm run test:tagger:play
 *   BASE_URL=http://127.0.0.1:3850 node scripts/tagger-play-demo.mjs
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
const SLOW = Number(process.env.SLOW_MS || 450);

function loadPlaywright() {
  const candidates = [
    join(root, "node_modules/playwright/index.mjs"),
    "/tmp/node_modules/playwright/index.mjs",
  ];
  for (const p of candidates) {
    if (existsSync(p)) return import(p);
  }
  return import("playwright");
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

async function pause(page, ms = SLOW) {
  await page.waitForTimeout(ms);
}

function say(msg) {
  console.log(`▶ ${msg}`);
}

async function selectTextInRoot(page, needle) {
  return page.evaluate((n) => {
    const root = document.getElementById("taggerSelectRoot");
    if (!root) return false;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = w.nextNode())) {
      const t = node.textContent || "";
      const i = t.indexOf(n);
      if (i < 0) continue;
      const range = document.createRange();
      range.setStart(node, i);
      range.setEnd(node, i + n.length);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return sel.toString().includes(n);
    }
    return false;
  }, needle);
}

async function openCreate(page) {
  await page.dispatchEvent("body", "pointerup");
  await pause(page, 350);
  const float = page.locator("#taggerCreateBtn:not([hidden])");
  if (await float.count()) {
    await float.click();
  } else {
    await page.evaluate(() => document.getElementById("taggerCreateBtn")?.click());
  }
  await page.waitForFunction(() => {
    const dlg = document.getElementById("taggerDialog");
    return dlg && !dlg.hidden;
  });
  await pause(page, 500);
}

async function setKind(page, kind) {
  await page.evaluate((k) => {
    const r = document.querySelector(`input[name="taggerKind"][value="${k}"]`);
    if (r) {
      r.checked = true;
      r.dispatchEvent(new Event("change", { bubbles: true }));
    }
  }, kind);
  await pause(page, 250);
}

async function confirm(page, name) {
  await page.fill("#taggerVarName", name);
  await pause(page, 300);
  await page.click("#taggerDialogConfirm");
  await pause(page, 1800);
}

async function startTagger(page, buffer, filename) {
  await page.goto(`${BASE}/?v=${CACHE}`, { waitUntil: "networkidle" });
  await pause(page, 600);
  await page.click("#startTaggerMode");
  await page.waitForSelector("#taggerFileInput");
  await pause(page, 400);
  await page.setInputFiles("#taggerFileInput", {
    name: filename,
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer,
  });
  await page.waitForSelector("#taggerSelectRoot", { timeout: 20000 });
  await pause(page, 800);
}

async function main() {
  const health = await fetch(`${BASE}/api/health`).then((r) => r.ok).catch(() => false);
  if (!health) {
    console.error(`Server not reachable at ${BASE}`);
    process.exit(2);
  }

  const { chromium } = await loadPlaywright();
  say("เปิด Chromium บนจอ (headed) — ดูหน้าต่างที่โผล่ขึ้นมา");
  const browser = await chromium.launch({
    headless: false,
    slowMo: Math.max(80, Math.floor(SLOW / 3)),
    args: ["--start-maximized"],
  });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 860 },
    permissions: ["clipboard-read", "clipboard-write"],
  });
  const page = await context.newPage();
  page.setDefaultTimeout(30000);

  const paraDoc = await minimalDocx(
    p("หัวข้อเอกสาร") +
      p("จำเลยที่1เลขนิติบุคคล 0195561000585ชื่อบริษัท ข้าวแกงแม่มณี") +
      p("จำเลยที่2เลขนิติบุคคล 0195561000999ชื่อบริษัท อีกบริษัท"),
  );
  const tableDoc = await minimalDocx(
    p("Title") +
      `<w:tbl>${row("สินค้า", "จำนวน", "ราคา")}${row("แอปเปิล", "10", "20")}${row("กล้วย", "5", "12")}</w:tbl>` +
      p("Note"),
  );
  const tableFile = existsSync(join(root, "examples/table-untagged.docx"))
    ? readFileSync(join(root, "examples/table-untagged.docx"))
    : tableDoc;

  // 1) Text
  say("1/6 อัปโหลด + แท็กข้อความ (text)");
  await startTagger(page, paraDoc, "play-para.docx");
  await selectTextInRoot(page, "หัวข้อเอกสาร");
  await openCreate(page);
  await setKind(page, "text");
  await confirm(page, "title");

  // 2) Bool
  say("2/6 แท็ก bool");
  await selectTextInRoot(page, "ข้าวแกงแม่มณี");
  await openCreate(page);
  await setKind(page, "bool");
  await page.fill("#taggerTrueText", "พบ");
  await page.fill("#taggerFalseText", "");
  await confirm(page, "พบบริษัท");

  // 3) List + auto
  say("3/6 ลิสต์ย่อหน้า + auto ลำดับ");
  await startTagger(page, paraDoc, "play-list.docx");
  await page.evaluate(() => {
    const root = document.getElementById("taggerSelectRoot");
    const paras = [...root.querySelectorAll(".tagger-para")];
    const target = paras.find((el) => (el.textContent || "").includes("จำเลยที่1"));
    if (!target) return;
    const range = document.createRange();
    range.setStartBefore(target);
    range.setEndAfter(target);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  });
  await openCreate(page);
  await setKind(page, "list");
  await confirm(page, "จำเลย");
  await page.evaluate(() => {
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
      return;
    }
  });
  await openCreate(page);
  await setKind(page, "auto");
  await confirm(page, "ลำดับ");

  // 4) Table list-table
  say("4/6 ตาราง → list-table");
  await startTagger(page, tableFile, "play-table.docx");
  await page.evaluate(() => {
    const tr = document.querySelectorAll("#taggerSelectRoot tr")[1];
    if (!tr) return;
    const range = document.createRange();
    range.selectNodeContents(tr);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  });
  await openCreate(page);
  await page.fill("#taggerVarName", "items");
  await page.evaluate(() => {
    [...document.querySelectorAll(".tagger-cell-field")].forEach((inp, i) => {
      inp.value = ["name", "qty", "price"][i] || "";
    });
  });
  await pause(page, 400);
  await page.click("#taggerDialogConfirm");
  await pause(page, 2000);

  // 5) Copy + paste
  say("5/6 คัดลอกข้อความ → วางชื่อตัวแปร");
  await startTagger(page, paraDoc, "play-copy.docx");
  await selectTextInRoot(page, "ข้าวแกงแม่มณี");
  await page.dispatchEvent("body", "pointerup");
  await pause(page, 400);
  await page.click("#taggerCopyBtn");
  await pause(page, 900);
  await selectTextInRoot(page, "หัวข้อเอกสาร");
  await openCreate(page);
  await setKind(page, "text");
  await page.click('[data-paste-into="taggerVarName"]');
  await pause(page, 900);
  await page.click("#taggerDialogCancel");
  await pause(page, 400);

  // 6) Delete + export
  say("6/6 ลบย่อหน้า + เปิดฟอร์มกรอก");
  await selectTextInRoot(page, "หัวข้อเอกสาร");
  await page.dispatchEvent("body", "pointerup");
  await pause(page, 350);
  await page.click("#taggerDeleteBtn");
  await pause(page, 500);
  await page.click("#taggerDeleteConfirm");
  await pause(page, 1600);
  await selectTextInRoot(page, "ข้าวแกงแม่มณี");
  await openCreate(page);
  await setKind(page, "text");
  await confirm(page, "company");
  say("เปิดหน้ากรอกเอกสาร…");
  await page.click("#taggerOpenFill");
  await pause(page, 2500);

  say("จบเดโม — หน้าต่างจะปิดใน 8 วินาที (Ctrl+C เพื่อค้างไว้)");
  await pause(page, 8000);
  await browser.close();
  console.log("PLAY DEMO DONE");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
