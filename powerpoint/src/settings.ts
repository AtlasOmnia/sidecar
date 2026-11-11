export interface HermesSettings {
  /** Chat endpoint base: the Sidecar bridge (set at runtime once the bridge is found; never stored). */
  baseUrl: string;
  model: string;
  maxIterations: number;
  /** Per-request bridge session headers (runtime only; never stored). No API key ever lives here. */
  requestHeaders?: Record<string, string>;
}

const STORAGE_KEY = "hermes.powerpoint.settings";
const DEFAULT_MODEL = "hermes-agent";
const LEGACY_MODELS = new Set(["hermes-4", ""]);

export const defaultSettings: HermesSettings = {
  baseUrl: "",
  model: DEFAULT_MODEL,
  maxIterations: 15
};

export function loadSettings(): HermesSettings {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return { ...defaultSettings };
  }

  try {
    const parsed = JSON.parse(raw) as Partial<HermesSettings> & Record<string, unknown>;
    const normalized = normalizeSettings(parsed);
    // Older builds stored a base URL and an API key here; neither belongs in the sidebar any more.
    if ("apiKey" in parsed || "baseUrl" in parsed || normalized.model !== parsed.model) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(persisted(normalized)));
    }
    return normalized;
  } catch {
    return { ...defaultSettings };
  }
}

export function saveSettings(settings: HermesSettings): HermesSettings {
  const normalized = normalizeSettings(settings);
  localStorage.setItem(STORAGE_KEY, JSON.stringify(persisted(normalized)));
  return { ...normalized, baseUrl: settings.baseUrl ?? "" };
}

function persisted(settings: HermesSettings): { model: string; maxIterations: number } {
  return { model: settings.model, maxIterations: settings.maxIterations };
}

function normalizeSettings(settings: Partial<HermesSettings>): HermesSettings {
  const maxIterations = Number(settings.maxIterations);
  const rawModel = (settings.model || defaultSettings.model).trim();
  return {
    baseUrl: "",
    model: LEGACY_MODELS.has(rawModel) ? DEFAULT_MODEL : rawModel,
    maxIterations: Number.isFinite(maxIterations) ? Math.min(Math.max(Math.floor(maxIterations), 1), 20) : defaultSettings.maxIterations
  };
}
