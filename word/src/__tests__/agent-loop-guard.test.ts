import { describe, expect, it } from "vitest";
import { makeToolCallSignature, isBlockedRepeatedToolCall } from "../agent";
import { ToolCall } from "../types";

describe("tool call loop guard helpers", () => {
  it("builds a stable signature from tool name and serialized arguments", () => {
    const toolCall: ToolCall = {
      id: "test",
      type: "function",
      function: {
        name: "get_document_overview",
        arguments: "{\"name\":\"current\"}"
      }
    };

    expect(makeToolCallSignature(toolCall)).toBe("get_document_overview:{\"name\":\"current\"}");
  });

  it("normalizes JSON object key order in tool signatures", () => {
    const first: ToolCall = {
      id: "first",
      type: "function",
      function: { name: "find_replace", arguments: '{"find":"a","replace":"b","options":{"matchCase":false,"whole":true}}' }
    };
    const second: ToolCall = {
      id: "second",
      type: "function",
      function: { name: "find_replace", arguments: '{"options":{"whole":true,"matchCase":false},"replace":"b","find":"a"}' }
    };

    expect(makeToolCallSignature(first)).toBe(makeToolCallSignature(second));
  });

  it("canonicalizes compatibility aliases in tool signatures", () => {
    const canonical: ToolCall = {
      id: "canonical",
      type: "function",
      function: { name: "read_document_text", arguments: "{}" }
    };
    const alias: ToolCall = {
      id: "alias",
      type: "function",
      function: { name: "get_document_text", arguments: "{}" }
    };

    expect(makeToolCallSignature(canonical)).toBe(makeToolCallSignature(alias));
  });

  it("allows two identical reads but blocks the third", () => {
    expect(isBlockedRepeatedToolCall(1)).toBe(false);
    expect(isBlockedRepeatedToolCall(2)).toBe(false);
    expect(isBlockedRepeatedToolCall(3)).toBe(true);
    expect(isBlockedRepeatedToolCall(4)).toBe(true);
  });

  it("blocks an identical write after its first execution", () => {
    expect(isBlockedRepeatedToolCall(1, "find_replace")).toBe(false);
    expect(isBlockedRepeatedToolCall(2, "find_replace")).toBe(true);
    expect(isBlockedRepeatedToolCall(2, "apply_edit_plan")).toBe(true);
    expect(isBlockedRepeatedToolCall(2, "read_document_text")).toBe(false);
  });
});
