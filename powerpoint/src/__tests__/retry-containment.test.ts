import { describe, expect, it } from "vitest";
import { ABSOLUTE_MAX_ITERATIONS, canonicalizeJson, generateXmlCallId } from "../agent";

describe("retry and repeat containment helpers", () => {
  it("canonicalizes nested object key order but preserves array order", () => {
    expect(canonicalizeJson({ b: { z: 1, a: 2 }, a: [2, 1] })).toBe("{\"a\":[2,1],\"b\":{\"a\":2,\"z\":1}}");
    expect(canonicalizeJson({ a: [1, 2] })).not.toBe(canonicalizeJson({ a: [2, 1] }));
  });

  it("gives malformed JSON a stable signature payload", () => {
    expect(canonicalizeJson(" {bad } ")).toBe("raw:{bad }");
    expect(canonicalizeJson("{bad }")).toBe(canonicalizeJson(" {bad } "));
  });

  it("exports an absolute twenty-iteration cap", () => {
    expect(ABSOLUTE_MAX_ITERATIONS).toBe(20);
  });

  it("generates collision-safe per-run XML call ids", () => {
    const first = generateXmlCallId("run-a", 1);
    const second = generateXmlCallId("run-a", 2);
    const otherRun = generateXmlCallId("run-b", 1);
    expect(new Set([first, second, otherRun]).size).toBe(3);
    expect(first).toContain("run-a");
  });
});
