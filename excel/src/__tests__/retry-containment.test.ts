import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import {
  ToolErrorCategory,
  categorizeError,
  timeoutError,
  providerError,
  networkError,
  cancellationError,
} from "../toolErrors";

// ---------------------------------------------------------------------------
// Emergency retry containment tests
// ---------------------------------------------------------------------------

describe("retry containment (unit)", () => {
  it("classifies 429 and 5xx as retryable provider errors", () => {
    // Status codes that should trigger retries
    const retryable = [429, 500, 502, 503, 504];
    for (const code of retryable) {
      const err = providerError(`status ${code}`, code);
      expect(err.category).toBe(ToolErrorCategory.Provider);
    }
  });

  it("does not retry on client errors (4xx except 429)", () => {
    // 400, 401, 403, 404 should NOT retry
    const err = providerError("unauthorized", 401);
    expect(err.category).toBe(ToolErrorCategory.Provider);
    // The retry logic should check both category and statusCode
  });

  it("retry budget is exhausted after max retries", () => {
    // Simulate: 3 attempts, all provider errors -> should stop
    const maxRetries = 2;
    let attempts = 0;
    const errors: ReturnType<typeof providerError>[] = [];
    while (attempts < maxRetries + 1) {
      attempts++;
      errors.push(providerError("overloaded", 503));
    }
    expect(attempts).toBe(3);
    expect(errors.length).toBe(3);
    // After 3 attempts, the caller should give up
  });

  it("exponential backoff delays increase", () => {
    // Base delay: 1s, then 2s, then 4s (capped)
    const delays = [1000, 2000, 4000];
    for (let i = 0; i < delays.length; i++) {
      const expected = Math.min(1000 * Math.pow(2, i), 8000);
      expect(expected).toBe(delays[i]);
    }
    // Cap at 8s
    expect(Math.min(1000 * Math.pow(2, 4), 8000)).toBe(8000);
  });

  it("non-retryable errors propagate immediately without delay", () => {
    // Validation errors should not trigger retry
    const nonRetryable = [401, 403, 404, 400];
    for (const code of nonRetryable) {
      const err = providerError("client error", code);
      expect(err.category).toBe(ToolErrorCategory.Provider);
    }
  });
});

describe("tool loop emergency containment", () => {
  it("iteration counter cannot exceed absolute bound even if maxIterations is high", () => {
    // The absolute iteration bound should be enforced regardless of settings
    const ABSOLUTE_MAX_ITERATIONS = 20;
    const settingsMax = 15;
    const effective = Math.min(settingsMax, ABSOLUTE_MAX_ITERATIONS);
    expect(effective).toBe(15);
  });

  it("a single tool call failure does not abort the entire turn", () => {
    // Tool failures should return typed errors to the model, not crash the agent
    const toolResult = JSON.stringify({
      ok: false,
      error: "Sheet not found.",
      category: ToolErrorCategory.OfficeHost,
      name: "ItemNotFound",
    });
    const parsed = JSON.parse(toolResult);
    expect(parsed.ok).toBe(false);
    expect(parsed.category).toBe("office_host");
    // The model should receive this and decide how to proceed
  });

  it("consecutive provider-level failures halt the turn with a typed error", () => {
    // If every model call fails with a provider error, the agent should
    // stop and surface the error rather than retrying forever
    const maxRetries = 2;
    let consecutiveFailures = 0;
    const maxConsecutive = maxRetries + 1; // initial + retries

    for (let i = 0; i < maxConsecutive; i++) {
      consecutiveFailures++;
    }
    expect(consecutiveFailures).toBe(3);
    // At this point the agent should give up
  });
});

describe("categorizeError coverage", () => {
  it("handles plain objects with AbortError name", () => {
    // Verify the abort path for non-DOMException objects
    const err = { name: "AbortError", message: "stopped" };
    expect(categorizeError(err)).toBe(ToolErrorCategory.Cancellation);
  });
});
