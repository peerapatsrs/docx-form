# Design: @peerapat.srs/docx-form

## Design principles

1. งานเดียวจบในหน้าเดียว — อัปโหลด → กรอก → ดาวน์โหลด
2. เอกสารเป็นศูนย์กลาง — ชื่อไฟล์และจำนวนตัวแปรเห็นชัด
3. กรอกเร็ว — Tab ได้; ค่าว่างได้; รูปไม่อัป = คง placeholder
4. ชนิดฟิลด์ชัด — ข้อความ / รูป / รายการ
5. แถวซ้ำกรอกง่าย — เพิ่ม/ลบแถว
6. ฟีดแบ็กทันที — ภาษาไทย
7. เครื่องมือบนเครื่อง — ไม่เก็บขึ้นคลาวด์

## Information architecture

หน้าเดียว: Top bar + Zone อัปโหลด / ฟอร์ม / สร้างเอกสาร

## User flows

อัปโหลด `.docx` → parse tags → แสดงฟอร์ม → กรอก/แนบรูป/เพิ่มแถว → สร้าง → ดาวน์โหลด

## Screens

### Empty
Dropzone + hint `{{ชื่อ}}` / Alt รูป / `{{#รายการ}}…{{/}}`

### Form
File chip, text inputs, image dropzones, list repeater, sticky primary button

## Design system

- พื้น: `#E8EEF2` → `#F4F7F9`, sheet `#FBFCFF`
- หมึก `#1B2A33`, muted `#5A6B76`, accent `#0F6B5C`
- ฟอนต์ Sarabun

## Accessibility

label ทุกช่อง, aria-live สำหรับ error, Tab ตามลำดับฟิลด์
