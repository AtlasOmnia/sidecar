// @vitest-environment jsdom
import { describe, expect, it, beforeEach } from "vitest";
import {
  defaultSettings,
  loadSettings,
  saveSettings
} from "../settings";

// Helper to set up localStorage mock (vitest with jsdom provides it)
function setLocalStorage(key: string, value: string): void {
  localStorage.setItem(key, value);
}

describe("defaultSettings", () => {
  it("default apiKey is empty string, not placeholder", () => {
    expect(defaultSettings).not.toHaveProperty("apiKey");
  });
});

describe("loadSettings empty localStorage", () => {
  beforeEach(() => localStorage.clear());

  it("returns default settings when no stored value", () => {
    const loaded = loadSettings();
    expect(loaded).not.toHaveProperty("apiKey");
    expect(loaded.baseUrl).toBe("");
    expect(loaded.model).toBe("hermes-agent");
    expect(loaded.maxIterations).toBe(15);
  });
});

describe("loadSettings normalization", () => {
  beforeEach(() => localStorage.clear());

  it("normalizes legacy base URL alias and legacy model", () => {
    setLocalStorage(
      "hermes.powerpoint.settings",
      JSON.stringify({
        baseUrl: "/hermes/v1",
        model: "",
        apiKey: " sk-test \n",
        maxIterations: 15
      })
    );

    const loaded = loadSettings();
    expect(loaded.baseUrl).toBe("");
    expect(loaded.model).toBe("hermes-agent");
    expect(loaded).not.toHaveProperty("apiKey");
  });

  it("clamps maxIterations to safe range", () => {
    setLocalStorage(
      "hermes.powerpoint.settings",
      JSON.stringify({
        baseUrl: "http://custom:8080/v1",
        model: "gpt-4",
        apiKey: [""].join(""),
        maxIterations: 1234
      })
    );

    const loaded = loadSettings();
    expect(loaded.maxIterations).toBe(20);
  });
});

describe("saveSettings round-trip", () => {
  beforeEach(() => localStorage.clear());

  it("preserves a non-empty apiKey and trims spaces", () => {
    const saved = saveSettings({ ...defaultSettings, ...{ apiKey: " sk-test " } });
    expect(saved).not.toHaveProperty("apiKey");
    const loaded = loadSettings();
    expect(loaded).not.toHaveProperty("apiKey");
  });

  it("round-trips complex settings", () => {
    saveSettings({
      ...defaultSettings,
      baseUrl: "http://custom:8080/v1",
      model: "gpt-4",
      maxIterations: 20,
      ...{ apiKey: " key-abc " }
    });
    const loaded = loadSettings();
    expect(loaded.baseUrl).toBe("");
    expect(loaded.model).toBe("gpt-4");
    expect(loaded.maxIterations).toBe(20);
    expect(loaded).not.toHaveProperty("apiKey");
  });
});
