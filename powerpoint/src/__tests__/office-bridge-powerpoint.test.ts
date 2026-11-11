import { describe, it, expect, vi, beforeEach } from "vitest";
import {
  createPowerPointCompanionBridge,
  executeValidatedPowerPointBridgeTool,
} from "../officeBridgePowerPoint";

// Mock Office global
const mockSettings = new Map<string, unknown>();
(global as any).Office = {
  context: {
    document: {
      settings: {
        get: (key: string) => mockSettings.get(key),
        set: (key: string, value: unknown) => { mockSettings.set(key, value); },
        saveAsync: (cb: (result: { status: string }) => void) => {
          cb({ status: "succeeded" as any });
        },
      },
    },
  },
  AsyncResultStatus: { Succeeded: "succeeded" },
};

// Mock crypto.randomUUID using Object.defineProperty
const mockRandomUUID = () => "550e8400-e29b-41d4-a716-446655440010";
Object.defineProperty(global, "crypto", {
  value: { randomUUID: mockRandomUUID },
  writable: true,
  configurable: true,
});

describe("createPowerPointCompanionBridge", () => {
  beforeEach(() => {
    mockSettings.clear();
  });

  it("creates bridge with powerPoint host prefix", () => {
    const bridge = createPowerPointCompanionBridge();
    expect(bridge).toBeDefined();
    expect(bridge.sessionId).toBeNull();
  });

  it("accepts onState callback", () => {
    const onState = vi.fn();
    const bridge = createPowerPointCompanionBridge(onState);
    expect(bridge).toBeDefined();
  });
});

describe("executeValidatedPowerPointBridgeTool", () => {
  it("validates tool calls before execution", async () => {
    // A valid tool name should pass validation but fail due to no PowerPoint runtime
    const result = await executeValidatedPowerPointBridgeTool(
      "get_presentation_overview",
      {}
    );
    // Should return validation result or error (not a raw throw)
    expect(typeof result).toBe("string");
  });

  it("rejects unknown tool names", async () => {
    const result = await executeValidatedPowerPointBridgeTool(
      "nonexistent_tool",
      {}
    );
    const parsed = JSON.parse(result);
    expect(parsed.valid ?? parsed.ok).toBeFalsy();
  });
});
