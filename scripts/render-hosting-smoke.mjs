#!/usr/bin/env node
/**
 * Smoke test for Render-style hosting:
 * PORT=10000 HOST=0.0.0.0 → GET / → tag → parse → fill → verify docx.
 */
import { spawn } from "node:child_process";
import { mkdirSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import JSZip from "jszip";

const PORT = Number(process.env.SMOKE_PORT || 10000);
const HOST = process.env.SMOKE_HOST || "0.0.0.0";
const BASE = `http://127.0.0.1:${PORT}`;
const OUT_DIR = join(tmpdir(), "docx-form-smoke");
mkdirSync(OUT_DIR, { recursive: true });

async function minimalDocx(bodyInner) {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyInner}
    <w:sectPr/>
  </w:body>
</w:document>`;
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", ct);
  zip.folder("_rels").file(".rels", rels);
  zip.folder("word").file("document.xml", doc);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

function p(text) {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function waitForHealth(timeoutMs = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(`${BASE}/api/health`);
      if (res.ok) return await res.json();
    } catch {
      // retry
    }
    await sleep(200);
  }
  throw new Error("Server did not become healthy in time");
}

async function main() {
  console.log(`Starting server PORT=${PORT} HOST=${HOST} …`);
  const child = spawn("node", ["dist/index.js", "--no-open"], {
    env: {
      ...process.env,
      PORT: String(PORT),
      HOST,
      NODE_ENV: "production",
      DOCX_FORM_DATA_DIR: join(OUT_DIR, "data"),
    },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let serverLog = "";
  child.stdout.on("data", (d) => {
    serverLog += d.toString();
    process.stdout.write(`[server] ${d}`);
  });
  child.stderr.on("data", (d) => {
    serverLog += d.toString();
    process.stderr.write(`[server:err] ${d}`);
  });

  const shutdown = () => {
    if (!child.killed) child.kill("SIGTERM");
  };
  process.on("exit", shutdown);
  process.on("SIGINT", () => {
    shutdown();
    process.exit(130);
  });

  try {
    await waitForHealth();

    // 1) GET /
    const home = await fetch(`${BASE}/`);
    console.log(`GET / → ${home.status}`);
    if (home.status !== 200) throw new Error(`GET / expected 200, got ${home.status}`);
    const homeHtml = await home.text();
    if (!homeHtml.includes("<!DOCTYPE html") && !homeHtml.includes("<!doctype html") && !homeHtml.includes("<html")) {
      throw new Error("GET / did not return HTML");
    }

    // 2) Create untagged sample .docx with plain text to tag
    const sampleBuf = await minimalDocx(p("Customer: PLACEHOLDER_NAME") + p("Thank you."));
    const samplePath = join(OUT_DIR, "sample-untagged.docx");
    writeFileSync(samplePath, sampleBuf);
    console.log(`Wrote sample: ${samplePath}`);

    // 3) Tag a field via /api/template/tag
    const tagFd = new FormData();
    tagFd.append(
      "file",
      new Blob([sampleBuf], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
      "sample-untagged.docx",
    );
    tagFd.append(
      "replacements",
      JSON.stringify([
        {
          id: "f1",
          plainText: "PLACEHOLDER_NAME",
          occurrence: 0,
          tag: "{{customer_name}}",
        },
      ]),
    );
    const tagRes = await fetch(`${BASE}/api/template/tag`, { method: "POST", body: tagFd });
    console.log(`POST /api/template/tag → ${tagRes.status}`);
    const tagJson = await tagRes.json();
    if (!tagRes.ok) throw new Error(`tag failed: ${JSON.stringify(tagJson)}`);
    if (!tagJson.docxBase64) throw new Error("tag response missing docxBase64");
    const taggedBuf = Buffer.from(tagJson.docxBase64, "base64");
    writeFileSync(join(OUT_DIR, "sample-tagged.docx"), taggedBuf);
    console.log(`Tagged filename=${tagJson.filename} replacementCount=${tagJson.replacementCount}`);

    // 4) Parse tagged template (as the web UI does after tagging / on upload)
    const parseFd = new FormData();
    parseFd.append(
      "file",
      new Blob([taggedBuf], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
      "sample-tagged.docx",
    );
    parseFd.append("includeContent", "1");
    const parseRes = await fetch(`${BASE}/api/parse`, { method: "POST", body: parseFd });
    console.log(`POST /api/parse → ${parseRes.status}`);
    const parseJson = await parseRes.json();
    if (!parseRes.ok) throw new Error(`parse failed: ${JSON.stringify(parseJson)}`);
    const fieldNames = (parseJson.fields || []).map((f) => f.name);
    console.log(`Parsed fields: ${JSON.stringify(fieldNames)}`);
    if (!fieldNames.includes("customer_name")) {
      throw new Error("expected customer_name field after tagging");
    }

    // 5) Fill form / generate via /api/fill
    const fillValue = "Render Smoke Co., Ltd.";
    const fillFd = new FormData();
    fillFd.append(
      "file",
      new Blob([taggedBuf], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
      "sample-tagged.docx",
    );
    fillFd.append("values", JSON.stringify({ customer_name: fillValue }));
    const fillRes = await fetch(`${BASE}/api/fill`, { method: "POST", body: fillFd });
    console.log(`POST /api/fill → ${fillRes.status}`);
    if (!fillRes.ok) {
      const errBody = await fillRes.text();
      throw new Error(`fill failed: ${errBody}`);
    }
    const filledBuf = Buffer.from(await fillRes.arrayBuffer());
    const filledPath = join(OUT_DIR, "sample-filled.docx");
    writeFileSync(filledPath, filledBuf);
    console.log(`Saved filled docx: ${filledPath} (${filledBuf.length} bytes)`);

    // 6) Confirm filled value is in the downloaded docx
    const zip = await JSZip.loadAsync(filledBuf);
    const xml = await zip.file("word/document.xml").async("string");
    if (!xml.includes(fillValue)) {
      throw new Error(`filled docx missing value "${fillValue}"`);
    }
    if (xml.includes("{{customer_name}}") || xml.includes("PLACEHOLDER_NAME")) {
      throw new Error("filled docx still contains placeholder/tag");
    }
    console.log(`OK: filled docx contains "${fillValue}"`);
    console.log("SMOKE PASS");
  } finally {
    shutdown();
    await sleep(300);
  }
}

main().catch((err) => {
  console.error("SMOKE FAIL:", err);
  process.exit(1);
});
