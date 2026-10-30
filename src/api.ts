import { Hono } from "hono";
import {
  extractDocumentLayout,
  extractDocumentText,
  fillTemplate,
  formatTemplateError,
  parseTemplate,
  type FormField,
} from "./docx-engine.js";

const MAX_DOCX_BYTES = 25 * 1024 * 1024;
const MAX_IMAGE_BYTES = 8 * 1024 * 1024;

function isDocx(file: File): boolean {
  const name = file.name.toLowerCase();
  if (name.endsWith(".docx")) return true;
  const t = (file.type || "").toLowerCase();
  return (
    t === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    t === "application/octet-stream"
  );
}

function summarizeFields(fields: FormField[]) {
  let text = 0;
  let image = 0;
  let list = 0;
  let bool = 0;
  let auto = 0;
  for (const f of fields) {
    if (f.type === "text") text += 1;
    else if (f.type === "image") image += 1;
    else if (f.type === "bool") bool += 1;
    else if (f.type === "auto") auto += 1;
    else list += 1;
  }
  return { text, image, list, bool, auto };
}

export function createApi(options: { startedAt: Date; version?: string }) {
  const api = new Hono();

  api.get("/health", (c) =>
    c.json({
      ok: true,
      name: "@peerapat.srs/docx-form",
      version: options.version || "0.0.0",
      author: "Peerapat Surasing",
      email: "peerapat.srs@gmail.com",
      startedAt: options.startedAt.toISOString(),
    }),
  );

  api.post("/parse", async (c) => {
    try {
      const form = await c.req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        return c.json({ error: "กรุณาอัปโหลดไฟล์ .docx" }, 400);
      }
      if (!isDocx(file)) {
        return c.json(
          {
            error:
              file.name.toLowerCase().endsWith(".doc")
                ? "รองรับเฉพาะ .docx — ให้ Save As เป็น .docx ใน Word ก่อน"
                : "ไฟล์ต้องเป็น .docx เท่านั้น",
          },
          400,
        );
      }
      if (file.size > MAX_DOCX_BYTES) {
        return c.json({ error: "ไฟล์ใหญ่เกินไป (สูงสุด 25 MB)" }, 400);
      }

      const buffer = Buffer.from(await file.arrayBuffer());
      const includeContent =
        String(form.get("includeContent") || "").toLowerCase() === "1" ||
        String(form.get("includeContent") || "").toLowerCase() === "true";

      const { fields } = await parseTemplate(buffer);

      if (fields.length === 0) {
        return c.json(
          {
            error:
              "ไม่พบตัวแปรในเทมเพลต — ใส่ {{ชื่อ}} ในข้อความ, แท็กใน Alt Text ของรูป, หรือ {{#รายการ}}…{{/}}",
            fields: [],
            filename: file.name,
          },
          400,
        );
      }

      const payload: Record<string, unknown> = {
        filename: file.name,
        fields,
        summary: summarizeFields(fields),
      };

      if (includeContent) {
        const layout = await extractDocumentLayout(buffer);
        payload.layout = layout.blocks;
        payload.content = await extractDocumentText(buffer);
        payload.contentIncluded = true;
      } else {
        payload.contentIncluded = false;
      }

      return c.json(payload);
    } catch (err) {
      const message = formatTemplateError(err);
      return c.json({ error: `อ่านเทมเพลตไม่สำเร็จ: ${message}` }, 400);
    }
  });

  api.post("/fill", async (c) => {
    try {
      const form = await c.req.formData();
      const file = form.get("file");
      if (!(file instanceof File)) {
        return c.json({ error: "กรุณาส่งไฟล์เทมเพลต .docx" }, 400);
      }
      if (!isDocx(file) || file.size > MAX_DOCX_BYTES) {
        return c.json({ error: "ไฟล์เทมเพลตไม่ถูกต้องหรือใหญ่เกินไป" }, 400);
      }

      const valuesRaw = form.get("values");
      if (typeof valuesRaw !== "string") {
        return c.json({ error: "ต้องส่งค่า values เป็น JSON" }, 400);
      }

      let values: Record<string, unknown>;
      try {
        values = JSON.parse(valuesRaw) as Record<string, unknown>;
      } catch {
        return c.json({ error: "values ต้องเป็น JSON ที่ถูกต้อง" }, 400);
      }

      const images: {
        name: string;
        buffer: Buffer;
        filename: string;
        mimeType?: string;
      }[] = [];

      for (const [key, val] of form.entries()) {
        if (!key.startsWith("image:") || !(val instanceof File)) continue;
        const name = key.slice("image:".length);
        if (!name) continue;
        if (val.size === 0) continue;
        if (val.size > MAX_IMAGE_BYTES) {
          return c.json({ error: `รูป ${name} ใหญ่เกิน 8 MB` }, 400);
        }
        const type = (val.type || "").toLowerCase();
        if (
          type &&
          !type.startsWith("image/") &&
          type !== "application/octet-stream"
        ) {
          return c.json({ error: `ไฟล์สำหรับ ${name} ไม่ใช่รูปภาพ` }, 400);
        }
        images.push({
          name,
          buffer: Buffer.from(await val.arrayBuffer()),
          filename: val.name || `${name}.png`,
          mimeType: val.type,
        });
      }

      const template = Buffer.from(await file.arrayBuffer());
      const out = await fillTemplate({ template, values, images });

      const base = file.name.replace(/\.docx$/i, "") || "document";
      const utf8Name = `${base}-filled.docx`;
      // HTTP headers must be ASCII; keep a safe fallback + RFC 5987 for Thai names
      const asciiName =
        utf8Name.replace(/[^\x20-\x7E]/g, "_").replace(/["\\]/g, "_") ||
        "document-filled.docx";
      const disposition = `attachment; filename="${asciiName}"; filename*=UTF-8''${encodeURIComponent(utf8Name)}`;

      return new Response(new Uint8Array(out), {
        status: 200,
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "Content-Disposition": disposition,
        },
      });
    } catch (err) {
      const message = formatTemplateError(err);
      return c.json({ error: `สร้างเอกสารไม่สำเร็จ: ${message}` }, 400);
    }
  });

  return api;
}
