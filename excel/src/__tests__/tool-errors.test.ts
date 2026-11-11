import { describe, expect, it } from "vitest";
import {
  ToolErrorCategory,
  validationError,
  timeoutError,
  cancellationError,
  repetitionError,
  officeHostError,
  networkError,
  providerError,
  protocolError,
  internalError,
  categorizeError,
  ToolError,
} from "../toolErrors";

describe("tool error factories", () => {
  it("creates a validation error with the correct category", () => {
    const err = validationError("bad args", "TypeError");
    expect(err.ok).toBe(false);
    expect(err.category).toBe(ToolErrorCategory.Validation);
    expect(err.error).toBe("bad args");
    expect(err.name).toBe("TypeError");
    expect(err.statusCode).toBeUndefined();
  });

  it("creates a timeout error", () => {
    const err = timeoutError("exceeded 20s");
    expect(err.category).toBe(ToolErrorCategory.Timeout);
  });

  it("creates a cancellation error", () => {
    const err = cancellationError("user stopped");
    expect(err.category).toBe(ToolErrorCategory.Cancellation);
  });

  it("creates a repetition error", () => {
    const err = repetitionError("blocked repeat");
    expect(err.category).toBe(ToolErrorCategory.Repetition);
  });

  it("creates an office host error", () => {
    const err = officeHostError("sheet not found", "ItemNotFound");
    expect(err.category).toBe(ToolErrorCategory.OfficeHost);
    expect(err.name).toBe("ItemNotFound");
  });

  it("creates a network error", () => {
    const err = networkError("fetch failed");
    expect(err.category).toBe(ToolErrorCategory.Network);
  });

  it("creates a provider error with status code", () => {
    const err = providerError("rate limited", 429);
    expect(err.category).toBe(ToolErrorCategory.Provider);
    expect(err.statusCode).toBe(429);
  });

  it("creates a protocol error", () => {
    const err = protocolError("malformed SSE chunk");
    expect(err.category).toBe(ToolErrorCategory.Protocol);
  });

  it("creates an internal error", () => {
    const err = internalError("unexpected null", "ReferenceError");
    expect(err.category).toBe(ToolErrorCategory.Internal);
    expect(err.name).toBe("ReferenceError");
  });

  it("every error factory produces ok:false", () => {
    const factories = [
      validationError("x"),
      timeoutError("x"),
      cancellationError("x"),
      repetitionError("x"),
      officeHostError("x"),
      networkError("x"),
      providerError("x"),
      protocolError("x"),
      internalError("x"),
    ];
    for (const err of factories) {
      expect(err.ok).toBe(false);
    }
  });
});

describe("categorizeError", () => {
  it("categorizes DOMException AbortError as Cancellation", () => {
    const err = new DOMException("aborted", "AbortError");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Cancellation);
  });

  it("categorizes DOMException TimeoutError as Timeout", () => {
    const err = new DOMException("timed out", "TimeoutError");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Timeout);
  });

  it("categorizes TypeError as Network", () => {
    const err = new TypeError("Failed to fetch");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Network);
  });

  it("categorizes plain Error as Internal", () => {
    const err = new Error("something went wrong");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Internal);
  });

  it("categorizes ItemNotFound-like object as OfficeHost", () => {
    const err = { code: "ItemNotFound", message: "ItemNotFound: sheet", name: "RichApi.Error" };
    expect(categorizeError(err)).toBe(ToolErrorCategory.OfficeHost);
  });

  it("categorizes conflict-like object as OfficeHost", () => {
    const err = { message: "name conflict: Sheet1 already exists", name: "Error" };
    expect(categorizeError(err)).toBe(ToolErrorCategory.OfficeHost);
  });

  it("categorizes object with AbortError name as Cancellation", () => {
    const err = { name: "AbortError", message: "cancelled" };
    expect(categorizeError(err)).toBe(ToolErrorCategory.Cancellation);
  });

  it("categorizes unknown object as Internal", () => {
    const err = { name: "WeirdThing", message: "???" };
    expect(categorizeError(err)).toBe(ToolErrorCategory.Internal);
  });

  it("categorizes null/undefined as Internal", () => {
    expect(categorizeError(null)).toBe(ToolErrorCategory.Internal);
    expect(categorizeError(undefined)).toBe(ToolErrorCategory.Internal);
  });

  it("categorizes string as Internal", () => {
    expect(categorizeError("raw string error")).toBe(ToolErrorCategory.Internal);
  });
});

describe("tool error JSON round-trip", () => {
  it("serializes and deserializes preserving category", () => {
    const original: ToolError = {
      ok: false,
      error: "Something failed",
      category: ToolErrorCategory.Network,
      statusCode: 502,
      name: "FetchError",
    };
    const json = JSON.stringify(original);
    const parsed = JSON.parse(json) as ToolError;
    expect(parsed.ok).toBe(false);
    expect(parsed.error).toBe("Something failed");
    expect(parsed.category).toBe("network");
    expect(parsed.statusCode).toBe(502);
    expect(parsed.name).toBe("FetchError");
  });
});
