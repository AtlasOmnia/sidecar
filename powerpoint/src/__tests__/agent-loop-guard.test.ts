import { describe, expect, it } from "vitest";
import { makeToolCallSignature, isBlockedRepeatedToolCall } from "../agent";
import { ToolCall } from "../types";

describe("tool call loop guard helpers", () => {
  it("builds a stable signature from tool name and serialized arguments", () => {
    const toolCall: ToolCall = {
      id: "test",
      type: "function",
      function: {
        name: "get_presentation_overview",
        arguments: "{\"name\":\"current\"}"
      }
    };

    expect(makeToolCallSignature(toolCall)).toBe("get_presentation_overview:{\"name\":\"current\"}");
  });

  it("marks calls as blocked only after the third identical appearance", () => {
    expect(isBlockedRepeatedToolCall(1)).toBe(false);
    expect(isBlockedRepeatedToolCall(2)).toBe(false);
    expect(isBlockedRepeatedToolCall(3)).toBe(true);
    expect(isBlockedRepeatedToolCall(4)).toBe(true);
  });
});
