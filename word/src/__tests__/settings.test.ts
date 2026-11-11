import { describe, expect, it, beforeEach } from "vitest";

import { defaultSettings, loadSettings, saveSettings } from "../settings";

describe("settings defaults", () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it("defaults maxIterations to 8", () => {
    expect(defaultSettings.maxIterations).toBe(8);
  });

  it("normalizes non-finite maxIterations to default", () => {
    const saved = saveSettings({
      baseUrl: "http://127.0.0.1:8645/v1",
      model: "hermes-agent",
      maxIterations: Infinity as unknown as number
    });
    expect(saved.maxIterations).toBe(8);
  });

  it("clamps maxIterations to min 1 and max 20", () => {
    expect(loadSettings()).toEqual(defaultSettings);

    const below = saveSettings({
      baseUrl: "http://127.0.0.1:8645/v1",
      model: "hermes-agent",
      maxIterations: -2
    });
    expect(below.maxIterations).toBe(1);

    const above = saveSettings({
      baseUrl: "http://127.0.0.1:8645/v1",
      model: "hermes-agent",
      maxIterations: 100
    });
    expect(above.maxIterations).toBe(20);

    expect(loadSettings().maxIterations).toBe(20);
  });
});
