import type { FormField } from "./docx-engine.js";
import type { ExtractedSource } from "./source-extract.js";

export type MapCandidate = {
  value: string | boolean;
  sourceFile: string;
  rule: string;
  confidence: number;
};

export type FieldProposal = {
  path: string;
  candidates: MapCandidate[];
  selectedIndex: number;
  /** When true, draft should skip this path */
  preferSkip?: boolean;
};

export type SourceMapResult = {
  proposals: FieldProposal[];
  values: Record<string, unknown>;
  skip: Record<string, boolean>;
  rowSkip: Record<string, Record<string, boolean>[]>;
  extracted: {
    filename: string;
    kind: string;
    charCount: number;
    warning?: string;
  }[];
};

export function extractedMeta(sources: ExtractedSource[]): SourceMapResult["extracted"] {
  return sources.map((s) => ({
    filename: s.filename,
    kind: s.kind,
    charCount: s.charCount,
    warning: s.warning,
  }));
}

/** Group Gemini hits into FieldProposal[] (one proposal per path). */
export function proposalsFromHits(
  hits: Array<MapCandidate & { path: string }>,
): FieldProposal[] {
  const byPath = new Map<string, FieldProposal>();
  for (const hit of hits) {
    const { path, ...cand } = hit;
    let prop = byPath.get(path);
    if (!prop) {
      prop = { path, candidates: [], selectedIndex: 0, preferSkip: false };
      byPath.set(path, prop);
    }
    const dup = prop.candidates.find(
      (c) => String(c.value) === String(cand.value) && c.rule === cand.rule,
    );
    if (dup) {
      dup.confidence = Math.max(dup.confidence, cand.confidence);
      continue;
    }
    prop.candidates.push(cand);
    prop.candidates.sort((a, b) => b.confidence - a.confidence);
    prop.selectedIndex = 0;
  }
  return [...byPath.values()].sort((a, b) => a.path.localeCompare(b.path, "th"));
}

const MISS_CAND: MapCandidate = {
  value: "(ไม่พบในเอกสาร)",
  sourceFile: "Google Gemini",
  rule: "gemini-miss",
  confidence: 0,
};

/**
 * Ensure every fillable template path has a proposal entry.
 * Missing paths get a skip-selected placeholder so the map panel lists them all.
 */
export function ensureCompleteProposals(
  fields: FormField[],
  proposals: FieldProposal[],
): FieldProposal[] {
  const byPath = new Map(proposals.map((p) => [p.path, p]));

  const ensure = (path: string, preferSkip = false) => {
    if (byPath.has(path)) return;
    const prop: FieldProposal = {
      path,
      candidates: [{ ...MISS_CAND }],
      selectedIndex: preferSkip ? -1 : 0,
      preferSkip,
    };
    byPath.set(path, prop);
  };

  for (const f of fields) {
    if (f.type === "auto" || f.type === "image") continue;
    if (f.type === "list") {
      const prefix = `${f.name}.`;
      const rowIdxs = new Set<number>();
      for (const p of byPath.keys()) {
        if (!p.startsWith(prefix)) continue;
        const m = /^\S+\.(\d+)\./.exec(p);
        if (m) rowIdxs.add(Number(m[1]));
      }
      // If Gemini found nothing for this list, still expose row 0 cells as skipped
      // so every item variable is visible and can be filled manually after apply.
      const rows = rowIdxs.size ? [...rowIdxs] : [0];
      for (const i of rows) {
        for (const item of f.itemFields) {
          if (item.type === "auto") continue;
          ensure(`${f.name}.${i}.${item.name}`, true);
        }
      }
      continue;
    }
    ensure(f.name, true);
  }

  return [...byPath.values()].sort((a, b) => a.path.localeCompare(b.path, "th"));
}

export function buildDraftFromProposals(
  fields: FormField[],
  proposals: FieldProposal[],
): Pick<SourceMapResult, "values" | "skip" | "rowSkip"> {
  const values: Record<string, unknown> = {};
  const skip: Record<string, boolean> = {};
  const rowSkip: Record<string, Record<string, boolean>[]> = {};
  const byPath = new Map(proposals.map((p) => [p.path, p]));

  const pick = (path: string): MapCandidate | null => {
    const p = byPath.get(path);
    if (!p || p.preferSkip || p.selectedIndex < 0 || !p.candidates.length) return null;
    return p.candidates[p.selectedIndex] || null;
  };

  for (const f of fields) {
    if (f.type === "auto" || f.type === "image") {
      skip[f.name] = f.type === "image";
      continue;
    }

    if (f.type === "list") {
      const prefix = `${f.name}.`;
      const rowIdxs = new Set<number>();
      for (const p of proposals) {
        if (!p.path.startsWith(prefix)) continue;
        const rest = p.path.slice(prefix.length);
        const m = /^(\d+)\./.exec(rest);
        if (m) rowIdxs.add(Number(m[1]));
      }
      const maxRow = rowIdxs.size ? Math.max(...rowIdxs) + 1 : 0;
      if (maxRow === 0) {
        skip[f.name] = true;
        values[f.name] = [];
        continue;
      }
      skip[f.name] = false;
      const rows: Record<string, unknown>[] = [];
      const rowSkipArr: Record<string, boolean>[] = [];
      for (let i = 0; i < maxRow; i += 1) {
        const row: Record<string, unknown> = {};
        const rs: Record<string, boolean> = {};
        for (const item of f.itemFields) {
          if (item.type === "auto") {
            row[item.name] = "";
            rs[item.name] = false;
            continue;
          }
          const cand = pick(`${f.name}.${i}.${item.name}`);
          if (!cand) {
            if (item.type === "bool") {
              rs[item.name] = false;
              row[item.name] = false;
            } else {
              rs[item.name] = true;
              row[item.name] = "";
            }
          } else {
            rs[item.name] = false;
            row[item.name] = cand.value;
          }
        }
        rows.push(row);
        rowSkipArr.push(rs);
      }
      values[f.name] = rows;
      rowSkip[f.name] = rowSkipArr;
      continue;
    }

    const cand = pick(f.name);
    if (!cand) {
      // Bool false is a real value — do not auto-skip (keeps control editable).
      if (f.type === "bool") {
        skip[f.name] = false;
        values[f.name] = false;
      } else {
        skip[f.name] = true;
        values[f.name] = "";
      }
    } else {
      skip[f.name] = false;
      values[f.name] = cand.value;
    }
  }

  return { values, skip, rowSkip };
}
