import { describe, expect, it } from "vitest";
import { resolveCreatedSlide, validateExplicitSlideIndex } from "../powerPointTools";

describe("exact PowerPoint slide targeting", () => {
  it.each([[-1, 3], [3, 3], [99, 3], [1.5, 3]])("rejects unsafe explicit index %s for count %s", (index, count) => {
    expect(() => validateExplicitSlideIndex(index, count)).toThrow(/slideIndex|index/i);
  });

  it("accepts a finite in-range integer", () => {
    expect(validateExplicitSlideIndex(2, 3)).toBe(2);
  });

  it("identifies the one newly inserted slide by id and actual index", () => {
    expect(resolveCreatedSlide(["a", "b"], ["a", "new", "b"])).toEqual({ id: "new", index: 1 });
  });

  it("rejects ambiguous creation results", () => {
    expect(() => resolveCreatedSlide(["a"], ["new", "a", "other"])).toThrow(/one new slide/i);
  });
});
