import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

export type AiConfig = {
  /** Google AI Studio / Gemini API key */
  geminiApiKey?: string;
  /** Preferred model id */
  geminiModel?: string;
};

const CONFIG_DIR = join(homedir(), ".docx-form");
const CONFIG_PATH = join(CONFIG_DIR, "config.json");
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

export function getConfigPath(): string {
  return CONFIG_PATH;
}

export function readAiConfig(): AiConfig {
  try {
    if (!existsSync(CONFIG_PATH)) return {};
    const raw = JSON.parse(readFileSync(CONFIG_PATH, "utf-8")) as AiConfig;
    return raw && typeof raw === "object" ? raw : {};
  } catch {
    return {};
  }
}

export function writeAiConfig(patch: Partial<AiConfig>): AiConfig {
  const next: AiConfig = { ...readAiConfig(), ...patch };
  if (!existsSync(CONFIG_DIR)) mkdirSync(CONFIG_DIR, { recursive: true });
  writeFileSync(CONFIG_PATH, JSON.stringify(next, null, 2), "utf-8");
  return next;
}

export function clearGeminiKey(): AiConfig {
  const cur = readAiConfig();
  delete cur.geminiApiKey;
  if (!existsSync(CONFIG_DIR)) mkdirSync(CONFIG_DIR, { recursive: true });
  writeFileSync(CONFIG_PATH, JSON.stringify(cur, null, 2), "utf-8");
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
    configPath: CONFIG_PATH,
    studioUrl: "https://aistudio.google.com/apikey",
  };
}
