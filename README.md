# @peerapat.srs/docx-form

Local CLI with a Web UI to fill Word (`.docx`) templates.

อ่านตัวแปรจากเทมเพลต Word แล้วเปิดฟอร์มให้กรอก แทนค่าด้วย [`easy-template-x`](https://www.npmjs.com/package/easy-template-x) — ข้อความ / ติ๊ก / รูป / ตาราง / เลขรันอัตโนมัติ

**npm:** https://www.npmjs.com/package/@peerapat.srs/docx-form  
**GitHub:** https://github.com/peerapatsrs/docx-form

**Author:** Peerapat Surasing · peerapat.srs@gmail.com

---

## เริ่มใช้งาน (ไทย)

### ติดตั้ง

ต้องมี **Node.js ≥ 20**

```bash
npm i -g @peerapat.srs/docx-form
docx-form
```

ทางเลือกโดยไม่ติดตั้ง global:

```bash
npx @peerapat.srs/docx-form
```

> หมายเหตุ: กล่อง **Install** บน npmjs.com แสดง `npm i @peerapat.srs/docx-form` (local) เสมอ — แก้ไม่ได้ผ่าน package.json สำหรับ CLI ให้ใช้ `-g` ตามด้านบน

### รัน

```bash
docx-form
# หรือ
docx-form start
```

เปิดเบราว์เซอร์ที่ `http://127.0.0.1:3850` อัตโนมัติ → อัปโหลด `.docx` → กรอก → สร้างเอกสาร

ไม่เปิดเบราว์เซอร์:

```bash
docx-form --no-open
```

เปลี่ยนพอร์ต:

```bash
docx-form -p 4000
```

### เขียนเทมเพลตใน Word

| ชนิด | วิธีใส่ในไฟล์ |
|------|----------------|
| ข้อความ | `{{customer_name}}` |
| ติ๊ก (bool) | `{{ชื่อ?ข้อความจริง,ข้อความเท็จ}}` เช่น `{{พบ?√,x}}` |
| รันเลข (auto) | `{{+ลำดับ}}` ในลูป — ใส่ 1, 2, 3… ตามแถว |
| รูป | แทรกรูป placeholder → ตั้ง **Alt Text** เป็น `{{logo}}` |
| ตาราง / ลิสต์ | หุ้มแถวด้วย `{{#items}}` … `{{/}}` |

ตัวอย่างติ๊ก:

```text
{{พบ?√,x}}
{{ไม่พบ?√,}}
```

- `{{พบ?√,x}}` → ติ๊กได้ `√` · ไม่ติ๊กได้ `x`
- `{{ไม่พบ?√,}}` → ติ๊กได้ `√` · ไม่ติ๊กว่าง

ตัวอย่างตาราง:

```text
| สินค้า | จำนวน | ราคา |
| {{#items}}{{name}} | {{qty}} | {{price}}{{/}} |
```

ตัวอย่างลิสต์:

```text
{{#items}}
- {{name}} ({{qty}})
{{/}}
```

### ใน Web UI

- พรีวิวเนื้อหาตามโครง Word (รวมตารางจริง / รวมเซลล์)
- กรอกในช่องว่างของเอกสาร · เพิ่ม/ลบแถวตาราง
- **Export / Import JSON** — ส่งออกหรือนำเข้าค่าที่กรอก (รวมรูปเป็น base64) เพื่อใช้ซ้ำกับเทมเพลตเดิม
- ลากปรับความกว้างคอลัมน์ได้เฉพาะหน้าจอ (ไม่บันทึกลงไฟล์)
- รีเฟรชหน้าแล้วยังกู้ค่าที่กรอกไว้ได้ (เก็บในเบราว์เซอร์)
- กดเปลี่ยนไฟล์จะล้างฉบับร่าง

ไฟล์ตัวอย่าง:

```bash
examples/sample-quote.docx
```

---

## What it is

`docx-form` starts a local HTTP server and opens a Web UI so you can:

- Upload a `.docx` template with `{{tags}}`
- Fill text, bool chips, images, and repeating list/table rows
- Preview layout close to Word (including real tables)
- Download a filled `.docx` without modifying the original

Templating engine: [`easy-template-x`](https://www.npmjs.com/package/easy-template-x) (MIT).

## Prerequisites

| Requirement | Notes |
|---|---|
| **Node.js** | `>= 20` |

## Install

Primary (global — recommended for a CLI):

```bash
npm i -g @peerapat.srs/docx-form
docx-form
```

Alternative without a global install:

```bash
npx @peerapat.srs/docx-form
```

> Note: The npmjs.com **Install** sidebar always shows local `npm i @peerapat.srs/docx-form` — that string is hardcoded by npm and cannot be changed via package.json. For this CLI, use `-g` as above.

## Run

Binary name: **`docx-form`**

Default command is `start` (so plain `docx-form` is enough):

```bash
docx-form
docx-form start
```

Default bind: `http://127.0.0.1:3850` — browser opens automatically.

### Common flags

| Flag | Default | Description |
|---|---|---|
| `-p, --port <number>` | `3850` | Listen port |
| `--host <host>` | `127.0.0.1` | Bind address |
| `--no-open` | (opens browser) | Do not open the browser |
| `-V, --version` | | Print version |
| `-h, --help` | | Help |

Examples:

```bash
docx-form --no-open
docx-form -p 4000
docx-form --host 127.0.0.1 --port 4000 --no-open
docx-form start -p 3850
```

## Template tags

| Kind | Syntax | Notes |
|---|---|---|
| Text | `{{name}}` | Plain replacement |
| Bool | `{{name?true,false}}` | Checked → true text; unchecked → false text |
| Auto | `{{+name}}` | Row index inside a loop (`1`, `2`, …) |
| Image | Alt text `{{name}}` on a placeholder picture | PNG/JPEG/… upload in UI |
| Loop | `{{#items}}` … `{{/}}` | Repeating rows / list items |

## Typical flow

1. `npm i -g @peerapat.srs/docx-form`
2. Run `docx-form`
3. Upload a `.docx` template
4. Fill fields in the preview / form
5. Click **สร้างเอกสาร** to download

Draft values (including the uploaded template) are kept in the browser so a refresh does not wipe the form. Choosing **เปลี่ยนไฟล์** clears that draft.

## Development

```bash
npm i
npm run dev
npm run build
npm start
```

## License

**ฟรี (Free)** — [MIT](./LICENSE)

ใช้งาน แก้ไข และแจกต่อได้ฟรี ไม่มีค่าลิขสิทธิ์แพ็กเกจนี้

## Credits

Peerapat Surasing · peerapat.srs@gmail.com
