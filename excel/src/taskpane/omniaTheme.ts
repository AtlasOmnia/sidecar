export type Palette = "teal" | "blue" | "gold" | "red";
export type ThemeMode = "auto" | "dark" | "light";
export interface ThemePreferences { palette: Palette; mode: ThemeMode }
export interface OfficeThemeSnapshot { isDarkTheme?: boolean; bodyBackgroundColor?: string }
const THEME_KEY = "omnia.theme";
const DEFAULT_THEME: ThemePreferences = { palette: "teal", mode: "auto" };
type ThemeStorage = Pick<Storage, "getItem" | "setItem">;

export function normalizeTheme(value: unknown): ThemePreferences {
  if (!value || typeof value !== "object" || Array.isArray(value)) return { ...DEFAULT_THEME };
  const stored = value as Partial<ThemePreferences>;
  return {
    palette: ["teal", "blue", "gold", "red"].includes(stored.palette ?? "") ? stored.palette! : "teal",
    mode: ["auto", "dark", "light"].includes(stored.mode ?? "") ? stored.mode! : "auto",
  };
}

export function loadTheme(storage?: Pick<Storage, "getItem">): ThemePreferences {
  try { return normalizeTheme(JSON.parse(storage?.getItem(THEME_KEY) ?? "null")); }
  catch { return { ...DEFAULT_THEME }; }
}

export function saveTheme(storage: Pick<Storage, "setItem"> | undefined, value: ThemePreferences): boolean {
  try {
    if (!storage) return false;
    storage.setItem(THEME_KEY, JSON.stringify(normalizeTheme(value)));
    return true;
  } catch { return false; }
}

export function resolveScheme(mode: ThemeMode, office: OfficeThemeSnapshot | undefined, prefersDark: boolean): "dark" | "light" {
  if (mode !== "auto") return mode;
  if (typeof office?.isDarkTheme === "boolean") return office.isDarkTheme ? "dark" : "light";
  const color = office?.bodyBackgroundColor;
  if (typeof color === "string" && /^#(?:[\da-f]{3}|[\da-f]{6})$/i.test(color)) {
    const hex = color.length === 4 ? [...color.slice(1)].map(c => c + c).join("") : color.slice(1);
    const [r, g, b] = [0, 2, 4].map(i => {
      const channel = parseInt(hex.slice(i, i + 2), 16) / 255;
      return channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4;
    });
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) < 0.5 ? "dark" : "light";
  }
  return prefersDark ? "dark" : "light";
}

export function applyTheme(root: HTMLElement, preferences: ThemePreferences, office: OfficeThemeSnapshot | undefined, prefersDark: boolean): void {
  root.dataset.palette = preferences.palette;
  root.dataset.scheme = resolveScheme(preferences.mode, office, prefersDark);
}

// Office context and storage can throw before host initialization or under storage restrictions.
export function browserStorage(): Storage | undefined {
  try { return window.localStorage; } catch { return undefined; }
}
export function currentOfficeTheme(): OfficeThemeSnapshot | undefined {
  try { return typeof Office === "undefined" ? undefined : Office.context?.officeTheme; }
  catch { return undefined; }
}
export function colorSchemeMedia(): MediaQueryList | undefined {
  return typeof window.matchMedia === "function" ? window.matchMedia("(prefers-color-scheme: dark)") : undefined;
}

interface ThemeControlsOptions {
  document: Document;
  storage?: ThemeStorage;
  media?: Pick<MediaQueryList, "matches" | "addEventListener" | "removeEventListener">;
  getOfficeTheme: () => OfficeThemeSnapshot | undefined;
}

/** Palette/appearance apply immediately, independently of the connection form's Save/Cancel. */
export function installThemeControls({ document: doc, storage, media, getOfficeTheme }: ThemeControlsOptions): { refresh: () => void; dispose: () => void } {
  let preferences = loadTheme(storage);
  const controls = Array.from(doc.querySelectorAll<HTMLInputElement>('[name="omnia-palette"], [name="omnia-mode"]'));
  const refresh = (): void => {
    let office: OfficeThemeSnapshot | undefined;
    try { office = getOfficeTheme(); } catch { office = undefined; }
    applyTheme(doc.documentElement, preferences, office, media?.matches ?? false);
    for (const control of controls) {
      control.checked = control.value === (control.name === "omnia-palette" ? preferences.palette : preferences.mode);
    }
    for (const swatch of Array.from(doc.querySelectorAll<HTMLElement>(".palette-swatch"))) {
      swatch.dataset.scheme = doc.documentElement.dataset.scheme;
    }
  };
  const change = (event: Event): void => {
    const input = event.currentTarget as HTMLInputElement;
    preferences = normalizeTheme({ ...preferences, [input.name === "omnia-palette" ? "palette" : "mode"]: input.value });
    saveTheme(storage, preferences);
    refresh();
  };
  controls.forEach(input => input.addEventListener("change", change));
  media?.addEventListener("change", refresh);
  doc.defaultView?.addEventListener("focus", refresh);
  doc.addEventListener("visibilitychange", refresh);
  // Excel has no portable OfficeThemeChanged event across supported hosts. Re-read
  // its lightweight theme snapshot, so Auto follows in-place host theme changes too.
  const poll = setInterval(() => { if (preferences.mode === "auto") refresh(); }, 1000);
  refresh();
  return { refresh, dispose: () => {
    clearInterval(poll);
    controls.forEach(input => input.removeEventListener("change", change));
    media?.removeEventListener("change", refresh);
    doc.defaultView?.removeEventListener("focus", refresh);
    doc.removeEventListener("visibilitychange", refresh);
  } };
}
