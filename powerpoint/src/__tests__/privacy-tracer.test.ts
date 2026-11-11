// @vitest-environment jsdom
import { describe, expect, it, beforeEach, vi } from "vitest";
import { classifyPrivacy, PRIVACY_LABELS } from "../privacy";
import { startTrace, recordModelCall, recordToolResult, recordCheckpoint, finalizeTrace, loadTraces, clearTraces } from "../turnTracer";

beforeEach(() => { localStorage.clear(); });

describe("PowerPoint privacy classification", () => {
  it("labels loopback as local bridge with cloud-backend warning", () => {
    expect(classifyPrivacy("http://127.0.0.1:8646/v1")).toBe("local");
    expect(classifyPrivacy("http://[::1]/v1")).toBe("local");
    expect(PRIVACY_LABELS.local.tooltip).toContain("cloud");
  });
  it("labels known cloud providers as cloud and other URLs external", () => {
    expect(classifyPrivacy("https://api.openai.com/v1")).toBe("cloud");
    expect(classifyPrivacy("https://example.test/v1")).toBe("external");
    expect(classifyPrivacy("not a url")).toBe("external");
    expect(classifyPrivacy("http://[::1")).toBe("external");
  });
});

describe("metadata-only PowerPoint traces", () => {
  it("stores only bounded metadata and evicts oldest traces", () => {
    for (let i = 0; i < 52; i += 1) {
      const trace = startTrace(20);
      recordModelCall(trace, 0, 2, 200);
      recordToolResult(trace, "format_text", 3, true);
      trace.outcome = "completed";
      finalizeTrace(trace);
    }
    expect(loadTraces()).toHaveLength(50);
    const serialized = JSON.stringify(loadTraces());
    for (const secret of ["SLIDE_SENTINEL", "sk-secret", "provider-body", "font.secret", "arguments", "results"]) expect(serialized).not.toContain(secret);
  });
  it("clears traces", () => { const trace = startTrace(20); trace.outcome = "completed"; finalizeTrace(trace); clearTraces(); expect(loadTraces()).toEqual([]); });
  it("treats localStorage persistence failures as best-effort telemetry", () => {
    const setItem = vi.spyOn(Storage.prototype, "setItem").mockImplementation(() => { throw new DOMException("quota", "QuotaExceededError"); });
    try {
      expect(() => finalizeTrace(startTrace(20))).not.toThrow();
    } finally {
      setItem.mockRestore();
    }
  });
  it("persists model status/category, tool outcome/category, and checkpoint metadata", () => {
    const trace = startTrace(20);
    recordModelCall(trace, { attempt: 1, durationMs: 12, statusCode: 503, category: "provider", outcome: "error" });
    recordToolResult(trace, "format_text", 3, false, "repetition");
    recordCheckpoint(trace, "continue");
    trace.outcome = "completed";
    finalizeTrace(trace);
    const stored = loadTraces()[0];
    expect(stored.modelCalls).toBe(1);
    expect(stored.modelAttempts[0]).toMatchObject({ attempt: 1, durationMs: 12, statusCode: 503, category: "provider", outcome: "error" });
    expect(stored.tools[0]).toMatchObject({ name: "format_text", durationMs: 3, ok: false, category: "repetition" });
    expect(stored.checkpoint).toBe("continue");
    expect(stored.outcome).toBe("completed");
  });

});
