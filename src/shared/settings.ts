/** 用户本地设置（仅存本机，不上传我们服务器） */

const KEY_API = "thesis_format_llm_api_key";
const KEY_BASE = "thesis_format_llm_base_url";
const KEY_MODEL = "thesis_format_llm_model";
const KEY_ENABLED = "thesis_format_llm_enabled";

export type LlmSettings = {
  apiKey: string;
  /** 默认 DeepSeek；开发态可用相对路径走 Vite 代理避免 CORS */
  baseUrl: string;
  model: string;
  enabled: boolean;
};

export const DEFAULT_LLM_SETTINGS: LlmSettings = {
  apiKey: "",
  baseUrl: "/llm",
  model: "deepseek-chat",
  enabled: true,
};

export function loadLlmSettings(): LlmSettings {
  try {
    return {
      apiKey: localStorage.getItem(KEY_API) ?? "",
      baseUrl: localStorage.getItem(KEY_BASE) ?? DEFAULT_LLM_SETTINGS.baseUrl,
      model: localStorage.getItem(KEY_MODEL) ?? DEFAULT_LLM_SETTINGS.model,
      enabled: localStorage.getItem(KEY_ENABLED) !== "0",
    };
  } catch {
    return { ...DEFAULT_LLM_SETTINGS };
  }
}

export function saveLlmSettings(partial: Partial<LlmSettings>) {
  const cur = loadLlmSettings();
  const next = { ...cur, ...partial };
  try {
    localStorage.setItem(KEY_API, next.apiKey);
    localStorage.setItem(KEY_BASE, next.baseUrl);
    localStorage.setItem(KEY_MODEL, next.model);
    localStorage.setItem(KEY_ENABLED, next.enabled ? "1" : "0");
  } catch {
    /* ignore */
  }
  return next;
}

export function maskApiKey(key: string): string {
  if (!key) return "未设置";
  if (key.length <= 8) return "****";
  return `${key.slice(0, 4)}…${key.slice(-4)}`;
}
