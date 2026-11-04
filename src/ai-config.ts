import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";

export type AiConfig = {
  /** Google AI Studio / Gemini API key */
  geminiApiKey?: string;
  /** Preferred model id */
  geminiModel?: string;
};

/**
 * Writable directory for uploads, outputs, or other server-side files.
 * Override with DOCX_FORM_DATA_DIR; defaults to OS temp (`<tmpdir>/docx-form`).
 */
export function getDataDir(): string {
  const fromEnv = process.env.DOCX_FORM_DATA_DIR?.trim();
  if (fromEnv) return fromEnv;
  return join(tmpdir(), "docx-form");
}

/**
 * AI config directory. Respects DOCX_FORM_DATA_DIR; on hosted/production uses
 * getDataDir() (temp); locally keeps ~/.docx-form so CLI behaviour is unchanged.
 */
function getConfigDir(): string {
  if (process.env.DOCX_FORM_DATA_DIR?.trim()) {
    return process.env.DOCX_FORM_DATA_DIR.trim();
  }
  const isHosted =
    process.env.NODE_ENV === "production" || process.env.RENDER === "true";
  if (isHosted) return getDataDir();
  return join(homedir(), ".docx-form");
}

function getConfigPath(): string {
  return join(getConfigDir(), "config.json");
}

export const DEFAULT_GEMINI_MODEL = "gemini-3.8-flash";

/** Curated text models suitable for source→field mapping. */
export const GEMINI_MODEL_OPTIONS: { id: string; label: string }[] = [
  { id: "gemini-3.8-flash", label: "Gemini 3.8 Flash (แนะนำ)" },
  { id: "gemini-3.6-flash", label: "Gemini 3.6 Flash" },
  { id: "gemini-3.5-flash-lite", label: "Gemini 3.5 Flash-Lite" },
  { id: "gemini-3.1-flash-lite", label: "Gemini 3.1 Flash-Lite" },
  { id: "gemini-3.1-pro-preview", label: "Gemini 3.1 Pro (preview)" },
  { id: "gemini-3-flash-preview", label: "Gemini 3 Flash (preview)" },
  { id: "gemini-flash-latest", label: "gemini-flash-latest" },
];

export { getConfigPath };

export function readAiConfig(): AiConfig {
  try {
    const configPath = getConfigPath();
    if (!existsSync(configPath)) return {};
    const raw = JSON.parse(readFileSync(configPath, "utf-8")) as AiConfig;
    return raw && typeof raw === "object" ? raw : {};
  } catch {
    return {};
  }
}

export function writeAiConfig(patch: Partial<AiConfig>): AiConfig {
  const next: AiConfig = { ...readAiConfig(), ...patch };
  const configDir = getConfigDir();
  const configPath = getConfigPath();
  if (!existsSync(configDir)) mkdirSync(configDir, { recursive: true });
  writeFileSync(configPath, JSON.stringify(next, null, 2), "utf-8");
  return next;
}

export function clearGeminiKey(): AiConfig {
  const cur = readAiConfig();
  delete cur.geminiApiKey;
  const configDir = getConfigDir();
  const configPath = getConfigPath();
  if (!existsSync(configDir)) mkdirSync(configDir, { recursive: true });
  writeFileSync(configPath, JSON.stringify(cur, null, 2), "utf-8");
  return cur;
}

export function getGeminiApiKey(): string | undefined {
  const fromEnv =
    process.env.DOCX_FORM_GEMINI_API_KEY ||
    process.env.GEMINI_API_KEY ||
    process.env.GOOGLE_API_KEY ||
    process.env.GOOGLE_GENERATIVE_AI_API_KEY;
  if (fromEnv && String(fromEnv).trim()) return String(fromEnv).trim();
  const cfg = readAiConfig().geminiApiKey;
  return cfg && String(cfg).trim() ? String(cfg).trim() : undefined;
}

export function normalizeGeminiModel(raw: unknown): string | null {
  const id = String(raw || "")
    .trim()
    .replace(/^models\//i, "");
  if (!id) return null;
  if (!/^[a-zA-Z0-9._-]+$/.test(id)) return null;
  if (id.length > 80) return null;
  return id;
}

export function getGeminiModel(): string {
  return normalizeGeminiModel(readAiConfig().geminiModel) || DEFAULT_GEMINI_MODEL;
}

export function geminiStatus() {
  const key = getGeminiApiKey();
  const masked = key ? `${key.slice(0, 6)}…${key.slice(-4)}` : null;
  return {
    connected: !!key,
    provider: "google-gemini" as const,
    model: getGeminiModel(),
    models: GEMINI_MODEL_OPTIONS,
    defaultModel: DEFAULT_GEMINI_MODEL,
    maskedKey: masked,
    configPath: getConfigPath(),
    dataDir: getDataDir(),
    studioUrl: "https://aistudio.google.com/apikey",
  };
}
