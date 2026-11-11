import { describe, expect, it } from "vitest";
import type { ChatMessage } from "../types";
import {
  buildServerOwnedRequest,
  isHermesAgentMode,
  shouldUseLocalArithmeticShortcut,
} from "../serverOwnedMode";

const BRIDGE_SESSION_ID = "660e8400-e29b-41d4-a716-446655440001";

describe("isHermesAgentMode", () => {
  it("returns true for hermes-agent model", () => {
    expect(isHermesAgentMode("hermes-agent")).toBe(true);
  });

  it("returns false for non-hermes-agent models", () => {
    expect(isHermesAgentMode("qwen3.6-27b")).toBe(false);
    expect(isHermesAgentMode("gpt-4o")).toBe(false);
    expect(isHermesAgentMode("hermes-4")).toBe(false);
  });

  it("ignores leading/trailing whitespace", () => {
    expect(isHermesAgentMode(" hermes-agent ")).toBe(true);
    expect(isHermesAgentMode(" qwen3.6 ")).toBe(false);
  });

  it("disables task-pane local arithmetic only for hermes-agent mode", () => {
    expect(shouldUseLocalArithmeticShortcut("hermes-agent")).toBe(false);
    expect(shouldUseLocalArithmeticShortcut(" hermes-agent ")).toBe(false);
    expect(shouldUseLocalArithmeticShortcut("qwen3.6-27b")).toBe(true);
  });
});

describe("buildServerOwnedRequest", () => {
  const settings = { model: "hermes-agent" as string, baseUrl: "", apiKey: "", maxIterations: 8 };

  it("keeps the bridge session out of model-visible messages and sends no tools or tool_choice", () => {
    const messages: ChatMessage[] = [
      { role: "system", content: "You are Hermes." },
      { role: "user", content: "Fresh workbook context:\nSheet1 A1=5\n\nUser request:\nsum column A" },
    ];

    const result = buildServerOwnedRequest(
      settings,
      messages,
      BRIDGE_SESSION_ID,
    );

    expect(result.model).toBe("hermes-agent");
    expect(result.messages).toHaveLength(2);
    // Verify tools and tool_choice are NOT present (server-owned mode sends none)
    const bodyKeys = Object.keys(result);
    expect(bodyKeys).not.toContain("tools");
    expect(bodyKeys).not.toContain("tool_choice");
    expect(result.stream).toBe(true);
    expect(result.temperature).toBe(0.2);
    expect(result.messages).toEqual(messages);
    expect(JSON.stringify(result.messages)).not.toContain("bridgeSessionId");
    expect(JSON.stringify(result.messages)).not.toContain(BRIDGE_SESSION_ID);
  });

  it("fails closed when no bridge session ID is provided", () => {
    const messages: ChatMessage[] = [{ role: "user", content: "do something" }];
    expect(() => buildServerOwnedRequest(settings, messages, null)).toThrow(/bridge.*session/i);
  });

  it("fails closed when the bridge session ID is not a valid UUID", () => {
    const messages: ChatMessage[] = [{ role: "user", content: "do something" }];
    expect(() => buildServerOwnedRequest(settings, messages, "not-a-uuid")).toThrow(/valid.*UUID/i);
  });

  it("does not append bridgeSessionId to the last user message context", () => {
    const messages: ChatMessage[] = [{ role: "user", content: "hello" }];
    const result = buildServerOwnedRequest(settings, messages, BRIDGE_SESSION_ID);

    const userMsg = result.messages.find((m) => m.role === "user");
    expect(userMsg?.content).toBe("hello");
    expect(userMsg?.content).not.toContain("bridgeSessionId");
    expect(userMsg?.content).not.toContain(BRIDGE_SESSION_ID);
  });
});
