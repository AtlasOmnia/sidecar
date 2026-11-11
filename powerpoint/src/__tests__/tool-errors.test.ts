import { describe, expect, it } from "vitest";
import { asToolError, stringifyToolError, validationError, type ToolError } from "../toolErrors";

describe("typed tool errors", () => {
  it("round-trips a typed validation error through JSON", () => {
    const error = Object.assign(validationError("bad argument"), { statusCode: 422 });
    expect(error).toMatchObject({ ok: false, category: "validation", error: "bad argument", statusCode: 422 });
    expect(JSON.parse(JSON.stringify(error))).toEqual(error);
  });

  it.each([
    [new DOMException("stop", "AbortError"), "cancellation", "AbortError"],
    [new DOMException("late", "TimeoutError"), "timeout", "TimeoutError"],
    [new TypeError("programming or parser failure"), "internal", "TypeError"],
  ])("classifies %s safely", (input, category, name) => {
    const error = asToolError(input);
    expect(error.ok).toBe(false);
    expect(error.category).toBe(category);
    expect(error.name).toBe(name);
  });

  it("preserves HTTP status without provider body content", () => {
    const error = asToolError({ status: 503, body: "SECRET_PROVIDER_BODY" });
    expect(error).toMatchObject({ category: "provider", statusCode: 503 });
    expect(JSON.stringify(error)).not.toContain("SECRET_PROVIDER_BODY");
  });

  it("maps Office code objects to an office-host category", () => {
    const error = asToolError({ code: "InvalidArgument", message: "invalid slide" });
    expect(error.category).toBe("office_host");
    expect(error.error).not.toContain("invalid slide");
    expect(error.details).toMatchObject({ code: "InvalidArgument" });
  });

  it("returns a deterministic serialized failure for arbitrary values", () => {
    const parsed = JSON.parse(stringifyToolError("boom")) as ToolError;
    expect(parsed.ok).toBe(false);
    expect(parsed.category).toBe("internal");
    expect(parsed.error).toBe("An internal Hermes error occurred.");
  });
});
