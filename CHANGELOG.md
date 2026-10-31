# Changelog

All notable changes to this project are documented in this file.

## [2.0.0] — 2026-11-01

Major release: template builder + AI source mapping are first-class, with fixes for Thai legal tables and list tags.

### Added

- **สร้างเทมเพลตจากเอกสาร (template tagger)** — highlight text in a sample `.docx` and replace with `{{tags}}` while keeping Word layout (OOXML in place).
  - Field / bool / list wraps; **wrap table rows** (`{{#…}}` … `{{/}}`)
  - Fill empty table cells; delete paragraph or table row from the UI
  - Floating selection toolbar; drafts saved in the browser (IndexedDB)
- **AI source map (Google Gemini)** — after Login Google AI, attach PDF / Word / images; extract text (incl. Thai+English OCR) and propose values for every template field.
  - Structured JSON schema for scalars + list rows
  - Second pass fills paths still missing
  - Panel lists every variable (found or skipped) before apply
- **Table-aware extraction** — empty cells kept as placeholders so column order is not shifted (e.g. ภูมิลำเนา vs ปัจจุบัน)
- Auto-repair list tables that only had `{{/}}` without `{{#…}}` on the template row (and mark header rows)
- Word `<w:tab/>` treated as tabs in tagger plain text / matching
- CLI `--host` bind (default `127.0.0.1`; use `0.0.0.0` when needed)

### Changed

- Version bump to **2.0.0**
- Footer credits: identity + version · MIT License (no “ฟรี” badge)
- Upload screen: tagger entry as a vertical CTA card

### Fixed

- List fields mis-mapped when sparse Word tables dropped empty cells in the AI corpus
- Long digit IDs selectable next to Thai text in the tagger
- Safari blank screen when IndexedDB draft load hung — timeout + paint empty UI first
- Existing `{{#}}` / `{{/}}` no longer re-sent as field replaces after cancel / re-wrap

---

## [1.1.0]

Earlier Web UI fill flow: text / bool / image / list loops, local drafts, Export/Import JSON.
