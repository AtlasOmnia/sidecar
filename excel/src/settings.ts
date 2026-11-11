export interface HermesSettings {
  baseUrl: string;
  model: string;
  /** Direct-provider compatibility only, never persisted or entered in the Desk UI. */
  apiKey: string;
  requestHeaders?: Record<string, string>;
  maxIterations: number;
  /** Enables the direct-provider WorkbookPlan turn only when explicitly true. */
  workbookPlanRolloutEnabled?: boolean;
}

const STORAGE_KEY = "hermes.excel.settings";
const DEFAULT_BASE_URL = "";
const DEFAULT_MODEL = "hermes-agent";
const DEFAULT_MAX_ITERATIONS = 8;
const LEGACY_MODELS = new Set(["hermes-4", ""]);

export const defaultSettings: HermesSettings = {
  baseUrl: DEFAULT_BASE_URL,
  model: DEFAULT_MODEL,
  apiKey: "",
  maxIterations: DEFAULT_MAX_ITERATIONS,
  workbookPlanRolloutEnabled: false,
};

export function loadSettings(): HermesSettings {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return { ...defaultSettings };
  }

  try {
    const parsed = JSON.parse(raw) as Partial<HermesSettings>;
    const normalized = normalizeSettings(parsed);
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
  return { ...normalized, baseUrl: settings.baseUrl };
}

function persisted(settings: HermesSettings) {
  return { model: settings.model, maxIterations: settings.maxIterations, workbookPlanRolloutEnabled: settings.workbookPlanRolloutEnabled };
}

function normalizeSettings(settings: Partial<HermesSettings>): HermesSettings {
  const maxIterations = Number(settings.maxIterations);
  const rawModel = (settings.model || defaultSettings.model).trim();
  return {
    baseUrl: "",
    model: LEGACY_MODELS.has(rawModel) ? DEFAULT_MODEL : rawModel,
    apiKey: "",
    maxIterations: Number.isFinite(maxIterations) ? Math.min(Math.max(Math.floor(maxIterations), 1), 10) : defaultSettings.maxIterations,
    workbookPlanRolloutEnabled: settings.workbookPlanRolloutEnabled === true,
  };
}
