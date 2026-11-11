import { describe, expect, it, beforeEach } from "vitest";
import { defaultSettings, loadSettings, saveSettings } from "../settings";

const storage = new Map<string, string>();
const testLocalStorage = {
  clear: () => storage.clear(),
  getItem: (key: string) => storage.get(key) ?? null,
  setItem: (key: string, value: string) => { storage.set(key, value); },
  removeItem: (key: string) => { storage.delete(key); },
};

beforeEach(() => {
  storage.clear();
  Object.defineProperty(globalThis, "localStorage", { configurable: true, value: testLocalStorage });
});

describe("workbook plan rollout setting", () => {

  it("defaults to a strict boolean false and ignores malformed stored values", () => {
    expect(defaultSettings.workbookPlanRolloutEnabled).toBe(false);
    expect(typeof defaultSettings.workbookPlanRolloutEnabled).toBe("boolean");
    localStorage.setItem("hermes.excel.settings", JSON.stringify({ workbookPlanRolloutEnabled: "true" }));
    expect(loadSettings().workbookPlanRolloutEnabled).toBe(false);
  });

  it("persists only a strict boolean rollout value", () => {
    const saved = saveSettings({ ...defaultSettings, workbookPlanRolloutEnabled: true });
    expect(saved.workbookPlanRolloutEnabled).toBe(true);
    expect(JSON.parse(localStorage.getItem("hermes.excel.settings")!).workbookPlanRolloutEnabled).toBe(true);
  });
});
